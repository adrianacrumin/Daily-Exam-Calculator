import re  #import regex for cleaning headers
import streamlit as st  #import streamlit ui
import pandas as pd  #import pandas
from io import BytesIO  #import BytesIO for in-memory files
from openpyxl import load_workbook  #import openpyxl to edit existing workbook

#-----------------------------
#Streamlit UI
#-----------------------------
st.set_page_config(page_title="CPI – Modality Averages", page_icon="page", layout="wide")  #set page
st.title("Daily/Weekly/Monthly/Quarterly/Yearly Calculator")  #title

uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])  #file uploader

#-----------------------------
#Helpers
#-----------------------------
def _norm_text(x:str)->str:  #normalize a header cell
    s = str(x).replace("\xa0"," ").strip().lower()  #strip + fix non-breaking spaces
    s = re.sub(r"\s+"," ", s)  #collapse spaces
    return s

def _period_to_ts(s: pd.Series, freq: str) -> pd.Series:  #convert period strings to timestamps
    p = pd.PeriodIndex(s.astype(str), freq=freq)
    return p.to_timestamp()

def _sniff_header(df: pd.DataFrame, look_for=("room", "study date")):  #find header row index
    max_scan = min(10, len(df))
    targets = set([_norm_text(t) for t in look_for])
    for i in range(max_scan):
        row_norm = [_norm_text(v) for v in df.iloc[i].tolist()]
        if all(any(t == c for c in row_norm) for t in targets):
            return i
    return None

def _read_sheet_robust(src_bytes: BytesIO, sheet_name: str):  #read a sheet with unknown header row
    src_bytes.seek(0)
    raw = pd.read_excel(src_bytes, sheet_name=sheet_name, header=None, dtype=str)
    if raw.empty:
        return None, None, None
    hdr_idx = _sniff_header(raw, ("room","study date"))  #prefer exact
    if hdr_idx is None:  #fallback: accept any “study” + “room”
        hdr_idx = _sniff_header(raw, ("room","study"))
    if hdr_idx is None:
        return raw, None, None
    headers = raw.iloc[hdr_idx].apply(lambda x: str(x).replace("\xa0"," ").strip())
    df = raw.iloc[hdr_idx+1:].reset_index(drop=True)
    df.columns = headers
    #detect columns with robust matching
    cols_norm = { _norm_text(c): c for c in df.columns }
    room_col = cols_norm.get("room")
    study_col = cols_norm.get("study date")
    if not room_col:
        #fallback partial for room
        for c in df.columns:
            if "room" in _norm_text(c):
                room_col = c; break
    if not study_col:
        #fallback partial for study date/time variants
        for c in df.columns:
            n = _norm_text(c)
            if ("study" in n and "date" in n) or n in ("studydate","study datetime","study date/time"):
                study_col = c; break
    return df, room_col, study_col

#-----------------------------
#Processing
#-----------------------------
if uploaded_file:
    src_bytes = BytesIO(uploaded_file.read())  #read upload
    src_bytes.seek(0)
    xls = pd.ExcelFile(src_bytes)
    sheet_names = xls.sheet_names

    results = {}
    modality_days = {"US": 4}  #ultrasound 4 days/week
    default_days = 5  #others 5 days/week

    for sname in sheet_names:
        df, room_col, date_col = _read_sheet_robust(src_bytes, sname)
        if df is None or df.empty:
            st.warning(f"sheet '{sname}': empty or unreadable — skipping")
            continue
        if not room_col or not date_col:
            st.warning(f"sheet '{sname}': couldn't find Room/Study Date — skipping. columns seen: {list(df.columns)[:8]}...")
            continue

        work = df.copy()
        work[date_col] = pd.to_datetime(work[date_col], errors="coerce")
        work = work.dropna(subset=[date_col])
        if work.empty:
            st.warning(f"sheet '{sname}': no valid Study Date values — skipping")
            continue

        work["Volume"] = 1
        work["StudyDate"] = work[date_col].dt.date

        daily = work.groupby([room_col, "StudyDate"], as_index=False)["Volume"].sum()
        daily["Week"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("W-MON")
        weekly = daily.groupby([room_col, "Week"], as_index=False)["Volume"].sum()
        daily["Month"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("M")
        monthly = daily.groupby([room_col, "Month"], as_index=False)["Volume"].sum()
        daily["Quarter"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("Q")
        quarterly = daily.groupby([room_col, "Quarter"], as_index=False)["Volume"].sum()
        daily["Year"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("Y")
        yearly = daily.groupby([room_col, "Year"], as_index=False)["Volume"].sum()

        rows = []
        for room in daily[room_col].astype(str).unique():
            d_per_wk = modality_days.get(room, default_days)
            dsub = daily[daily[room_col].astype(str) == room]
            avg_day = dsub["Volume"].mean()
            avg_week_total = dsub.groupby("Week")["Volume"].sum().mean()
            avg_week_per_day = (avg_week_total / d_per_wk) if d_per_wk else None
            avg_month = dsub.groupby("Month")["Volume"].sum().mean()
            avg_quarter = dsub.groupby("Quarter")["Volume"].sum().mean()
            avg_year = dsub.groupby("Year")["Volume"].sum().mean()
            rows.append([room, avg_day, avg_week_per_day, avg_month, avg_quarter, avg_year])

        averages = pd.DataFrame(
            rows,
            columns=["Room","Avg/Day","Avg/Week(per scheduled day)","Avg/Month","Avg/Quarter","Avg/Year"]
        )

        results[sname] = {
            "Daily": daily.rename(columns={room_col:"Room"}),
            "Weekly": weekly.rename(columns={room_col:"Room"}),
            "Monthly": monthly.rename(columns={room_col:"Room"}),
            "Quarterly": quarterly.rename(columns={room_col:"Room"}),
            "Yearly": yearly.rename(columns={room_col:"Room"}),
            "Averages": averages
        }

    if not results:
        st.error("No usable sheets. If headers aren’t on the first row, this version scans the first 10 rows—upload a sample if it still fails.")
        st.stop()

    #previews + charts
    for sname, t in results.items():
        st.header(sname)
        st.caption("preview of daily (first 10 rows)")
        st.dataframe(t["Daily"].head(10), use_container_width=True)

        pick_opts = ["All"] + sorted(t["Daily"]["Room"].astype(str).unique().tolist())
        pick = st.selectbox(f"select room for charts ({sname})", pick_opts, key=f"pick_{sname}")

        with st.expander("daily chart", expanded=False):
            dfc = t["Daily"].copy()
            dfc["StudyDate"] = pd.to_datetime(dfc["StudyDate"])
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.pivot_table(index="StudyDate", values="Volume", aggfunc="sum").sort_index()
            else:
                pivot = dfc.pivot_table(index="StudyDate", columns="Room", values="Volume", aggfunc="sum").fillna(0).sort_index()
            st.line_chart(pivot) if not pivot.empty else st.info("no daily data")

        with st.expander("weekly chart", expanded=False):
            dfc = t["Weekly"].copy()
            dfc["WeekTS"] = _period_to_ts(dfc["Week"].astype(str), "W-MON")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("WeekTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="WeekTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            st.bar_chart(pivot) if not pivot.empty else st.info("no weekly data")

        with st.expander("monthly chart", expanded=False):
            dfc = t["Monthly"].copy()
            dfc["MonthTS"] = _period_to_ts(dfc["Month"].astype(str), "M")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("MonthTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="MonthTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            st.bar_chart(pivot) if not pivot.empty else st.info("no monthly data")

        with st.expander("quarterly chart", expanded=False):
            dfc = t["Quarterly"].copy()
            dfc["QuarterTS"] = _period_to_ts(dfc["Quarter"].astype(str), "Q")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("QuarterTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="QuarterTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            st.bar_chart(pivot) if not pivot.empty else st.info("no quarterly data")

        with st.expander("yearly chart", expanded=False):
            dfc = t["Yearly"].copy()
            dfc["YearTS"] = _period_to_ts(dfc["Year"].astype(str), "Y")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("YearTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="YearTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            st.bar_chart(pivot) if not pivot.empty else st.info("no yearly data")

        st.divider()

    #append new sheets to same workbook
    src_bytes.seek(0)
    wb = load_workbook(filename=src_bytes)
    #remove old result sheets to avoid duplicates
    for ws in list(wb.worksheets):
        if ws.title.endswith(("_Daily","_Weekly","_Monthly","_Quarterly","_Yearly","_Averages")):
            wb.remove(ws)

    out_bytes = BytesIO()
    with pd.ExcelWriter(out_bytes, engine="openpyxl") as writer:
        writer.book = wb
        writer.sheets = {ws.title: ws for ws in wb.worksheets}
        for sheet, tables in results.items():
            tables["Daily"].to_excel(writer, sheet_name=f"{sheet}_Daily", index=False)
            tables["Weekly"].to_excel(writer, sheet_name=f"{sheet}_Weekly", index=False)
            tables["Monthly"].to_excel(writer, sheet_name=f"{sheet}_Monthly", index=False)
            tables["Quarterly"].to_excel(writer, sheet_name=f"{sheet}_Quarterly", index=False)
            tables["Yearly"].to_excel(writer, sheet_name=f"{sheet}_Yearly", index=False)
            tables["Averages"].to_excel(writer, sheet_name=f"{sheet}_Averages", index=False)
        writer.book.save(out_bytes)

    st.download_button(
        label="Download Results Excel",
        data=out_bytes.getvalue(),
        file_name=uploaded_file.name.replace(".xlsx","_with_averages.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Added new result sheets to your original workbook.")
