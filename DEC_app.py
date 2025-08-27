import streamlit as st  #import streamlit ui
import pandas as pd  #import pandas
from io import BytesIO  #import BytesIO for in-memory files
from openpyxl import load_workbook  #import openpyxl to edit existing workbook

#-----------------------------
#Streamlit UI
#-----------------------------
st.set_page_config(page_title="CPI – Modality Averages", page_icon="page", layout="wide")  #set page
st.title("Modality Daily/Weekly/Monthly/Quarterly/Yearly Calculator")  #title

uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])  #file uploader

#-----------------------------
#Helpers
#-----------------------------
def _period_to_ts(s: pd.Series, freq: str) -> pd.Series:  #convert period strings to timestamps
    p = pd.PeriodIndex(s.astype(str), freq=freq)  #period index
    return p.to_timestamp()  #to timestamps

def _find_col_case_insensitive(df: pd.DataFrame, wanted: str) -> str:  #find exact name ignoring case/space
    want = wanted.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == want:
            return c
    raise KeyError(f"column '{wanted}' not found")

#-----------------------------
#Processing
#-----------------------------
if uploaded_file:  #run when file present
    src_bytes = BytesIO(uploaded_file.read())  #read upload into memory

    #read list of sheet names
    src_bytes.seek(0)  #rewind before read
    xls = pd.ExcelFile(src_bytes)  #inspect workbook
    sheet_names = xls.sheet_names  #all sheets

    results = {}  #store per-sheet tables
    modality_days = {"US": 4}  #us scheduled 4 days/week
    default_days = 5  #others scheduled 5 days/week

    for sname in sheet_names:  #loop sheets
        src_bytes.seek(0)  #rewind before each read
        df = pd.read_excel(src_bytes, sheet_name=sname)  #read one sheet
        if df.empty:
            continue  #skip empty

        try:
            room_col = _find_col_case_insensitive(df, "Room")  #use Room as modality
            date_col = _find_col_case_insensitive(df, "Study Date")  #use Study Date as date
        except KeyError as e:
            st.warning(f"sheet '{sname}': {e} — skipping")
            continue

        #clean
        work = df.copy()  #copy
        work[date_col] = pd.to_datetime(work[date_col], errors="coerce")  #to datetime
        work = work.dropna(subset=[date_col])  #drop bad dates
        work["Volume"] = 1  #each row counts as 1
        work["StudyDate"] = work[date_col].dt.date  #pure date for grouping

        #daily
        daily = (
            work.groupby([room_col, "StudyDate"], as_index=False)["Volume"].sum()
        )  #daily totals per room

        #period rollups
        daily["Week"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("W-MON")  #week
        weekly = daily.groupby([room_col, "Week"], as_index=False)["Volume"].sum()

        daily["Month"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("M")  #month
        monthly = daily.groupby([room_col, "Month"], as_index=False)["Volume"].sum()

        daily["Quarter"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("Q")  #quarter
        quarterly = daily.groupby([room_col, "Quarter"], as_index=False)["Volume"].sum()

        daily["Year"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("Y")  #year
        yearly = daily.groupby([room_col, "Year"], as_index=False)["Volume"].sum()

        #averages per room
        rows = []  #collector
        for room in daily[room_col].astype(str).unique():
            d_per_wk = modality_days.get(room, default_days)  #4 for US else 5
            dsub = daily[daily[room_col].astype(str) == room]  #subset
            avg_day = dsub["Volume"].mean()  #avg/day
            avg_week_total = dsub.groupby("Week")["Volume"].sum().mean()  #avg weekly total
            avg_week_per_day = (avg_week_total / d_per_wk) if d_per_wk else None  #per scheduled day
            avg_month = dsub.groupby("Month")["Volume"].sum().mean()  #avg/month
            avg_quarter = dsub.groupby("Quarter")["Volume"].sum().mean()  #avg/quarter
            avg_year = dsub.groupby("Year")["Volume"].sum().mean()  #avg/year
            rows.append([room, avg_day, avg_week_per_day, avg_month, avg_quarter, avg_year])

        averages = pd.DataFrame(
            rows,
            columns=["Room","Avg/Day","Avg/Week(per scheduled day)","Avg/Month","Avg/Quarter","Avg/Year"]
        )  #averages table

        results[sname] = {
            "room_col": room_col,  #save names for charts
            "Daily": daily.rename(columns={room_col: "Room"}),
            "Weekly": weekly.rename(columns={room_col: "Room"}),
            "Monthly": monthly.rename(columns={room_col: "Room"}),
            "Quarterly": quarterly.rename(columns={room_col: "Room"}),
            "Yearly": yearly.rename(columns={room_col: "Room"}),
            "Averages": averages
        }  #store

    #show previews + charts
    if not results:
        st.error("No usable sheets. Ensure columns are exactly 'Room' and 'Study Date'.")
        st.stop()

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
            if pivot.empty: st.info("no daily data"); 
            else: st.line_chart(pivot)

        with st.expander("weekly chart", expanded=False):
            dfc = t["Weekly"].copy()
            dfc["WeekTS"] = _period_to_ts(dfc["Week"].astype(str), "W-MON")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("WeekTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="WeekTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            if pivot.empty: st.info("no weekly data")
            else: st.bar_chart(pivot)

        with st.expander("monthly chart", expanded=False):
            dfc = t["Monthly"].copy()
            dfc["MonthTS"] = _period_to_ts(dfc["Month"].astype(str), "M")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("MonthTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="MonthTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            if pivot.empty: st.info("no monthly data")
            else: st.bar_chart(pivot)

        with st.expander("quarterly chart", expanded=False):
            dfc = t["Quarterly"].copy()
            dfc["QuarterTS"] = _period_to_ts(dfc["Quarter"].astype(str), "Q")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("QuarterTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="QuarterTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            if pivot.empty: st.info("no quarterly data")
            else: st.bar_chart(pivot)

        with st.expander("yearly chart", expanded=False):
            dfc = t["Yearly"].copy()
            dfc["YearTS"] = _period_to_ts(dfc["Year"].astype(str), "Y")
            if pick != "All":
                dfc = dfc[dfc["Room"].astype(str) == str(pick)]
                pivot = dfc.groupby("YearTS")["Volume"].sum().to_frame()
            else:
                pivot = dfc.pivot_table(index="YearTS", columns="Room", values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            if pivot.empty: st.info("no yearly data")
            else: st.bar_chart(pivot)

        st.divider()

    #append new sheets to the same workbook
    src_bytes.seek(0)  #rewind to load workbook
    wb = load_workbook(filename=src_bytes)  #load existing workbook

    #optionally remove old result sheets to avoid duplicates
    to_remove = [ws.title for ws in wb.worksheets if ws.title.endswith(("_Daily","_Weekly","_Monthly","_Quarterly","_Yearly","_Averages"))]
    for name in to_remove:
        ws = wb[name]
        wb.remove(ws)

    out_bytes = BytesIO()  #new output buffer
    with pd.ExcelWriter(out_bytes, engine="openpyxl") as writer:  #open writer
        writer.book = wb  #attach existing workbook
        writer.sheets = {ws.title: ws for ws in wb.worksheets}  #map sheets

        for sheet, tables in results.items():  #write result sheets
            tables["Daily"].to_excel(writer, sheet_name=f"{sheet}_Daily", index=False)
            tables["Weekly"].to_excel(writer, sheet_name=f"{sheet}_Weekly", index=False)
            tables["Monthly"].to_excel(writer, sheet_name=f"{sheet}_Monthly", index=False)
            tables["Quarterly"].to_excel(writer, sheet_name=f"{sheet}_Quarterly", index=False)
            tables["Yearly"].to_excel(writer, sheet_name=f"{sheet}_Yearly", index=False)
            tables["Averages"].to_excel(writer, sheet_name=f"{sheet}_Averages", index=False)

        writer.book.save(out_bytes)  #save workbook with new sheets

    st.download_button(  #download combined workbook
        label="Download Results Excel",
        data=out_bytes.getvalue(),
        file_name=uploaded_file.name.replace(".xlsx", "_with_averages.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Added new *_Daily/*_Weekly/*_Monthly/*_Quarterly/*_Yearly/*_Averages sheets to your original workbook.")
