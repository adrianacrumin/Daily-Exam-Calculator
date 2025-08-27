import re
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

#-----------------------------
#UI
#-----------------------------
st.set_page_config(page_title="CPI – Modality Tables", page_icon="page", layout="wide")
st.title("Daily/Weekly/Monthly/Quarterly/Yearly Tables")
uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])

#-----------------------------
#Helpers
#-----------------------------
def _norm_text(x: str) -> str:
    s = str(x).replace("\xa0", " ").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def canonical_room(name: str) -> str:
    #collapse variants: uppercase, strip, remove non-alphanumerics
    #e.g. "CPI CT1", "cpi-ct1", "CPICT1 " -> "CPICT1"
    return re.sub(r"[^A-Za-z0-9]+", "", str(name).upper().strip())

US_PATTERN = re.compile(r"US", re.I)  #we'll test it on canonical names

def scheduled_days_for_room(room_name: str) -> int:
    #US rooms are 4-day weeks; others 5
    return 4 if US_PATTERN.search(canonical_room(room_name)) else 5

def _sniff_header(df: pd.DataFrame, look_for=("room", "study date")):
    max_scan = min(10, len(df))
    targets = set(_norm_text(t) for t in look_for)
    for i in range(max_scan):
        row_norm = [_norm_text(v) for v in df.iloc[i].tolist()]
        if any("room" == c or "room" in c for c in row_norm) and any(
            "study date" == c or "scheduled date" == c or ("study" in c and "date" in c)
            for c in row_norm
        ):
            return i
    return None

def _read_sheet_robust(src_bytes: BytesIO, sheet_name: str):
    #read raw, find the header row that contains Room and a Study/Scheduled Date
    src_bytes.seek(0)
    raw = pd.read_excel(src_bytes, sheet_name=sheet_name, header=None, dtype=str)
    if raw.empty:
        return None, None, None

    hdr_idx = _sniff_header(raw, ("room", "study date"))
    if hdr_idx is None:
        hdr_idx = _sniff_header(raw, ("room", "study"))
    if hdr_idx is None:
        return raw, None, None

    headers = raw.iloc[hdr_idx].apply(lambda x: str(x).replace("\xa0", " ").strip())
    df = raw.iloc[hdr_idx + 1 :].reset_index(drop=True)
    df.columns = headers

    nmap = {_norm_text(c): c for c in df.columns}
    #Room column
    room_col = nmap.get("room")
    if not room_col:
        for k, v in nmap.items():
            if "room" in k:
                room_col = v
                break
    #Date column: prefer Study Date, else Scheduled Date, else any "study" + "date"
    date_col = nmap.get("study date") or nmap.get("scheduled date")
    if not date_col:
        for k, v in nmap.items():
            if ("study" in k and "date" in k):
                date_col = v
                break

    return df, room_col, date_col

def make_wide_table(long_df: pd.DataFrame, index_col: str, period_name: str) -> pd.DataFrame:
    pivot = long_df.pivot_table(index=index_col, columns="Room", values="Volume", aggfunc="sum").fillna(0)
    pivot["Total Exams"] = pivot.sum(axis=1)
    pivot = pivot.sort_index()
    pivot.index.name = period_name
    return pivot.reset_index()

def add_pct_change_table(wide_df: pd.DataFrame, period_col: str) -> pd.DataFrame:
    df = wide_df.copy()
    num_cols = [c for c in df.columns if c != period_col]
    pct = df[num_cols].pct_change().round(4)
    pct.insert(0, period_col, df[period_col])
    pct.columns = [period_col] + [f"{c} %Δ" for c in num_cols]
    return pct

def append_overall_average_row(wide_df: pd.DataFrame, period_col: str, label: str) -> pd.DataFrame:
    df = wide_df.copy()
    num_cols = [c for c in df.columns if c != period_col]
    avg_row = {period_col: label}
    for c in num_cols:
        avg_row[c] = df[c].mean()
    return pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)

def round_numeric(df: pd.DataFrame, digits: int = 1) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if pd.api.types.is_numeric_dtype(out[c]):
            out[c] = out[c].round(digits)
    return out

def business_days_range(start_date: pd.Timestamp, end_date: pd.Timestamp) -> pd.DatetimeIndex:
    #Mon–Fri only
    return pd.date_range(start_date, end_date, freq="B")

def insert_weekly_avg_rows_in_daily_strict(daily_wide: pd.DataFrame) -> pd.DataFrame:
    """
    Insert one 'Weekly Avg YYYY-MM-DD→YYYY-MM-DD' row AFTER each complete Mon–Fri block.
    - Weekends excluded.
    - Per-room weekly avg divides by 5 (or 4 for rooms containing 'US').
    - 'Total Exams' on avg row is the sum of per-room averages.
    - Trailing partial week: show daily rows only (no avg row).
    """
    df = daily_wide.copy()
    date_col = df.columns[0]  #should be 'Date'
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.normalize()
    df = df.dropna(subset=[date_col])

    #fill missing weekdays with zeros so averages are correct by schedule
    all_bd = business_days_range(df[date_col].min(), df[date_col].max())
    df = df.set_index(date_col).reindex(all_bd).fillna(0.0).rename_axis(date_col).reset_index()

    dow = df[date_col].dt.dayofweek  #Mon=0..Sun=6
    week_start = df[date_col] - pd.to_timedelta(dow, unit="D")
    week_end = week_start + pd.to_timedelta(4, unit="D")
    df["_wk_start"] = week_start
    df["_wk_end"] = week_end

    all_cols = list(df.columns)
    room_cols = [c for c in all_cols if c not in [date_col, "Total Exams", "_wk_start", "_wk_end"]]

    blocks = []
    for (ws, we), g in df.groupby(["_wk_start", "_wk_end"], sort=True):
        g2 = g.drop(columns=["_wk_start", "_wk_end"]).copy()
        g2 = g2[g2[date_col].dt.dayofweek <= 4]  #Mon–Fri only
        if len(g2) == 5:
            avg_row = {date_col: f"Weekly Avg {ws.date()}→{we.date()}"}
            total_avg_sum = 0.0
            for col in room_cols:
                denom = scheduled_days_for_room(col)  #use canonicalized logic
                val = g2[col].sum() / denom if denom else 0.0
                avg_row[col] = val
                total_avg_sum += val
            avg_row["Total Exams"] = total_avg_sum
            blocks.append(g2)
            blocks.append(pd.DataFrame([avg_row]))
        else:
            blocks.append(g2)  #partial week: no avg row

    out = pd.concat(blocks, ignore_index=True)
    out[date_col] = out[date_col].apply(lambda x: x.date() if isinstance(x, pd.Timestamp) else x)
    return out

#-----------------------------
#Main
#-----------------------------
if uploaded_file:
    src_bytes = BytesIO(uploaded_file.read())
    src_bytes.seek(0)
    xls = pd.ExcelFile(src_bytes)
    sheet_names = xls.sheet_names

    results = {}

    for sname in sheet_names:
        df, room_col, date_col = _read_sheet_robust(src_bytes, sname)
        if df is None or df.empty:
            #likely an empty or formatting-only sheet
            continue
        if not room_col or not date_col:
            #Skip non-detail sheets (like prior pivots) gracefully
            if sname.lower() not in ("sheet1", "pivot", "summary"):
                st.warning(f"sheet '{sname}': couldn't find Room/Study Date — skipping. sample columns: {list(df.columns)[:8]}")
            continue

        #prep
        work = df.copy()
        work[date_col] = pd.to_datetime(work[date_col], errors="coerce")
        work = work.dropna(subset=[date_col])
        if work.empty:
            continue

        #canonicalize room names so variants roll up under one column
        work["Room"] = work[room_col].apply(canonical_room)
        work["Volume"] = 1
        work["StudyDate"] = work[date_col].dt.date

        #debug expanders (helps verify what was detected)
        with st.expander(f"Debug: detected columns and room counts on '{sname}'", expanded=False):
            st.write({"room_col": room_col, "date_col": date_col})
            st.dataframe(work["Room"].value_counts().head(30))

        #Daily long
        daily_long = work.groupby(["Room", "StudyDate"], as_index=False)["Volume"].sum()

        #Daily wide + weekly avg after each complete Mon–Fri
        daily_wide = make_wide_table(daily_long, "StudyDate", "Date")
        daily_wide = insert_weekly_avg_rows_in_daily_strict(daily_wide)
        daily_wide = round_numeric(daily_wide, 1)

        #Weekly totals (Mon-start)
        daily_long["Week"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("W-MON")
        weekly_long = daily_long.groupby(["Room", "Week"], as_index=False)["Volume"].sum()
        weekly_wide = make_wide_table(weekly_long, "Week", "Week")
        weekly_wide = append_overall_average_row(weekly_wide, "Week", "Average (Weekly)")
        weekly_wide = round_numeric(weekly_wide, 1)

        #Monthly
        daily_long["Month"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("M")
        monthly_long = daily_long.groupby(["Room", "Month"], as_index=False)["Volume"].sum()
        monthly_wide = make_wide_table(monthly_long, "Month", "Month")
        monthly_wide = append_overall_average_row(monthly_wide, "Month", "Average (Monthly)")
        monthly_wide = round_numeric(monthly_wide, 1)
        monthly_mom = add_pct_change_table(monthly_wide.drop(monthly_wide.index[-1]), "Month")

        #Quarterly
        daily_long["Quarter"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Q")
        quarterly_long = daily_long.groupby(["Room", "Quarter"], as_index=False)["Volume"].sum()
        quarterly_wide = make_wide_table(quarterly_long, "Quarter", "Quarter")
        quarterly_wide = append_overall_average_row(quarterly_wide, "Quarter", "Average (Quarterly)")
        quarterly_wide = round_numeric(quarterly_wide, 1)

        #Yearly
        daily_long["Year"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Y")
        yearly_long = daily_long.groupby(["Room", "Year"], as_index=False)["Volume"].sum()
        yearly_wide = make_wide_table(yearly_long, "Year", "Year")
        yearly_wide = append_overall_average_row(yearly_wide, "Year", "Average (Yearly)")
        yearly_wide = round_numeric(yearly_wide, 1)
        yearly_yoy = add_pct_change_table(yearly_wide.drop(yearly_wide.index[-1]), "Year")

        results[sname] = {
            "Daily": daily_wide,
            "Weekly": weekly_wide,
            "Monthly": monthly_wide,
            "Monthly_MoM_%": monthly_mom,
            "Quarterly": quarterly_wide,
            "Yearly": yearly_wide,
            "Yearly_YoY_%": yearly_yoy,
        }

    if not results:
        st.error("No usable sheets found. Make sure a sheet has 'Room' and 'Study Date' (or 'Scheduled Date').")
    else:
        #show the tables
        for sname, tabs in results.items():
            st.header(sname)
            st.subheader("Daily (with Weekly Avg rows; Mon–Fri complete weeks only)")
            st.dataframe(tabs["Daily"], use_container_width=True)
            st.subheader("Weekly (with bottom Average)")
            st.dataframe(tabs["Weekly"], use_container_width=True)
            st.subheader("Monthly (with bottom Average)")
            st.dataframe(tabs["Monthly"], use_container_width=True)
            st.subheader("Monthly (MoM % change)")
            st.dataframe(tabs["Monthly_MoM_%"], use_container_width=True)
            st.subheader("Quarterly (with bottom Average)")
            st.dataframe(tabs["Quarterly"], use_container_width=True)
            st.subheader("Yearly (with bottom Average)")
            st.dataframe(tabs["Yearly"], use_container_width=True)
            st.subheader("Yearly (YoY % change)")
            st.dataframe(tabs["Yearly_YoY_%"], use_container_width=True)
            st.divider()

        #append to original workbook
        src_bytes.seek(0)
        wb = load_workbook(filename=src_bytes)
        #remove prior result sheets to avoid duplicates
        for ws in list(wb.worksheets):
            if ws.title.endswith(("_Daily","_Weekly","_Monthly","_Monthly_MoM_%","_Quarterly","_Yearly","_Yearly_YoY_%")):
                wb.remove(ws)

        out_bytes = BytesIO()
        with pd.ExcelWriter(out_bytes, engine="openpyxl") as writer:
            writer.book = wb
            writer.sheets = {ws.title: ws for ws in wb.worksheets}
            for sname, tabs in results.items():
                tabs["Daily"].to_excel(writer, sheet_name=f"{sname}_Daily", index=False)
                tabs["Weekly"].to_excel(writer, sheet_name=f"{sname}_Weekly", index=False)
                tabs["Monthly"].to_excel(writer, sheet_name=f"{sname}_Monthly", index=False)
                tabs["Monthly_MoM_%"].to_excel(writer, sheet_name=f"{sname}_Monthly_MoM_%", index=False)
                tabs["Quarterly"].to_excel(writer, sheet_name=f"{sname}_Quarterly", index=False)
                tabs["Yearly"].to_excel(writer, sheet_name=f"{sname}_Yearly", index=False)
                tabs["Yearly_YoY_%"].to_excel(writer, sheet_name=f"{sname}_Yearly_YoY_%", index=False)
            writer.book.save(out_bytes)

        st.download_button(
            label="Download Results Excel",
            data=out_bytes.getvalue(),
            file_name=uploaded_file.name.replace(".xlsx", "_with_tables.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Tables updated. Room names canonicalized; US=4-day weeks, others=5; Mon–Fri weekly averages only.")
