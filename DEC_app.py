import re  #import regex for cleaning headers
import streamlit as st  #import streamlit ui
import pandas as pd  #import pandas
from io import BytesIO  #import BytesIO for in-memory files
from openpyxl import load_workbook  #import openpyxl to edit existing workbook

#-----------------------------
#Streamlit UI
#-----------------------------
st.set_page_config(page_title="CPI – Modality Tables", page_icon="page", layout="wide")  #set page
st.title("Daily/Weekly/Monthly/Quarterly/Yearly Tables")  #title
uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])  #file uploader

#-----------------------------
#Helpers
#-----------------------------
def _norm_text(x:str)->str:  #normalize a header cell
    s = str(x).replace("\xa0"," ").strip().lower()  #strip + fix non-breaking spaces
    s = re.sub(r"\s+"," ", s)  #collapse spaces
    return s

def _period_to_ts(s: pd.Series, freq: str) -> pd.Series:  #period string→timestamp
    p = pd.PeriodIndex(s.astype(str), freq=freq)  #period index
    return p.to_timestamp()  #start-of-period timestamp

def _sniff_header(df: pd.DataFrame, look_for=("room", "study date")):  #find header row index
    max_scan = min(10, len(df))  #scan first 10 rows
    targets = set([_norm_text(t) for t in look_for])  #normalize targets
    for i in range(max_scan):  #check each candidate row
        row_norm = [_norm_text(v) for v in df.iloc[i].tolist()]
        if all(any(t == c for c in row_norm) for t in targets):
            return i
    return None  #not found

def _read_sheet_robust(src_bytes: BytesIO, sheet_name: str):  #read sheet with unknown header row
    src_bytes.seek(0)  #rewind
    raw = pd.read_excel(src_bytes, sheet_name=sheet_name, header=None, dtype=str)  #raw read
    if raw.empty:
        return None, None, None  #empty
    hdr_idx = _sniff_header(raw, ("room","study date"))  #try exact
    if hdr_idx is None:  #fallback: accept any “study” + “room”
        hdr_idx = _sniff_header(raw, ("room","study"))
    if hdr_idx is None:
        return raw, None, None  #no header found
    headers = raw.iloc[hdr_idx].apply(lambda x: str(x).replace("\xa0"," ").strip())  #clean headers
    df = raw.iloc[hdr_idx+1:].reset_index(drop=True)  #data below header
    df.columns = headers  #assign headers

    cols_norm = { _norm_text(c): c for c in df.columns }  #normalized map
    room_col = cols_norm.get("room")  #room column
    study_col = cols_norm.get("study date")  #study date column
    if not room_col:
        for c in df.columns:
            if "room" in _norm_text(c):
                room_col = c; break
    if not study_col:
        for c in df.columns:
            n = _norm_text(c)
            if ("study" in n and "date" in n) or n in ("studydate","study datetime","study date/time"):
                study_col = c; break
    return df, room_col, study_col  #return cleaned df and column names

def make_wide_table(long_df: pd.DataFrame, index_col: str, period_name: str)->pd.DataFrame:  #long→wide
    pivot = long_df.pivot_table(index=index_col, columns="Room", values="Volume", aggfunc="sum").fillna(0)  #pivot
    pivot["Total Exams"] = pivot.sum(axis=1)  #total column
    pivot = pivot.sort_index()  #sort by index
    pivot.index.name = period_name  #label index
    out = pivot.reset_index()  #back to columns
    return out  #wide table

def add_pct_change_table(wide_df: pd.DataFrame, period_col: str)->pd.DataFrame:  #add % change columns
    df = wide_df.copy()  #copy
    pct = df[[c for c in df.columns if c != period_col]].pct_change().round(4)  #pct change
    pct.insert(0, period_col, df[period_col])  #keep period column
    #rename columns with suffix
    pct.columns = [period_col] + [f"{c} %Δ" for c in df.columns if c != period_col]  #suffix
    return pct  #table of % change

#-----------------------------
#Main processing
#-----------------------------
if uploaded_file:  #run when file present
    src_bytes = BytesIO(uploaded_file.read())  #read upload into memory
    src_bytes.seek(0)  #rewind
    xls = pd.ExcelFile(src_bytes)  #inspect workbook
    sheet_names = xls.sheet_names  #all sheet names

    results = {}  #collect per-sheet outputs

    for sname in sheet_names:  #loop sheets
        df, room_col, date_col = _read_sheet_robust(src_bytes, sname)  #robust read
        if df is None or df.empty:
            st.warning(f"sheet '{sname}': empty or unreadable — skipping")  #warn
            continue  #next
        if not room_col or not date_col:
            st.warning(f"sheet '{sname}': couldn't find Room/Study Date — skipping. sample columns: {list(df.columns)[:8]}")  #warn
            continue  #next

        work = df.copy()  #copy
        work[date_col] = pd.to_datetime(work[date_col], errors="coerce")  #to datetime
        work = work.dropna(subset=[date_col])  #drop bad dates
        if work.empty:
            st.warning(f"sheet '{sname}': no valid Study Date values — skipping")  #warn
            continue  #next

        work["Volume"] = 1  #each row counts as one exam
        work["Room"] = work[room_col].astype(str).str.strip()  #normalize room
        work["StudyDate"] = work[date_col].dt.date  #pure date

        daily_long = work.groupby(["Room","StudyDate"], as_index=False)["Volume"].sum()  #daily long
        daily_wide = make_wide_table(daily_long, "StudyDate", "Date")  #daily wide

        daily_long["Week"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("W-MON")  #week period
        weekly_long = daily_long.groupby(["Room","Week"], as_index=False)["Volume"].sum()  #weekly long
        weekly_wide = make_wide_table(weekly_long, "Week", "Week")  #weekly wide

        daily_long["Month"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("M")  #month period
        monthly_long = daily_long.groupby(["Room","Month"], as_index=False)["Volume"].sum()  #monthly long
        monthly_wide = make_wide_table(monthly_long, "Month", "Month")  #monthly wide
        monthly_mom = add_pct_change_table(monthly_wide, "Month")  #MoM % change

        daily_long["Quarter"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Q")  #quarter period
        quarterly_long = daily_long.groupby(["Room","Quarter"], as_index=False)["Volume"].sum()  #quarterly long
        quarterly_wide = make_wide_table(quarterly_long, "Quarter", "Quarter")  #quarterly wide

        daily_long["Year"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Y")  #year period
        yearly_long = daily_long.groupby(["Room","Year"], as_index=False)["Volume"].sum()  #yearly long
        yearly_wide = make_wide_table(yearly_long, "Year", "Year")  #yearly wide
        yearly_yoy = add_pct_change_table(yearly_wide, "Year")  #YoY % change

        results[sname] = {  #store tables
            "Daily": daily_wide,
            "Weekly": weekly_wide,
            "Monthly": monthly_wide,
            "Monthly_MoM_%": monthly_mom,
            "Quarterly": quarterly_wide,
            "Yearly": yearly_wide,
            "Yearly_YoY_%": yearly_yoy,
        }

    if not results:  #no usable sheets
        st.error("No usable sheets found. Ensure headers include 'Room' and 'Study Date' (any row, first 10 scanned).")
        st.stop()

    #-----------------------------
    #Show tables in the app
    #-----------------------------
    for sname, tabs in results.items():  #display per sheet
        st.header(sname)  #sheet name
        st.subheader("Daily")  #daily table
        st.dataframe(tabs["Daily"], use_container_width=True)  #show
        st.subheader("Weekly")  #weekly table
        st.dataframe(tabs["Weekly"], use_container_width=True)  #show
        st.subheader("Monthly")  #monthly table
        st.dataframe(tabs["Monthly"], use_container_width=True)  #show
        st.subheader("Monthly (MoM % change)")  #MoM table
        st.dataframe(tabs["Monthly_MoM_%"], use_container_width=True)  #show
        st.subheader("Quarterly")  #quarterly table
        st.dataframe(tabs["Quarterly"], use_container_width=True)  #show
        st.subheader("Yearly")  #yearly table
        st.dataframe(tabs["Yearly"], use_container_width=True)  #show
        st.subheader("Yearly (YoY % change)")  #YoY table
        st.dataframe(tabs["Yearly_YoY_%"], use_container_width=True)  #show
        st.divider()  #separator

    #-----------------------------
    #Append tables into the same Excel
    #-----------------------------
    src_bytes.seek(0)  #rewind to load workbook
    wb = load_workbook(filename=src_bytes)  #load workbook

    #remove prior result sheets to avoid duplicates
    for ws in list(wb.worksheets):
        if ws.title.endswith(("_Daily","_Weekly","_Monthly","_Monthly_MoM_%","_Quarterly","_Yearly","_Yearly_YoY_%")):
            wb.remove(ws)

    out_bytes = BytesIO()  #prepare new output buffer
    with pd.ExcelWriter(out_bytes, engine="openpyxl") as writer:  #open writer
        writer.book = wb  #attach existing workbook
        writer.sheets = {ws.title: ws for ws in wb.worksheets}  #sheet map

        for sname, tabs in results.items():  #write each table as its own sheet
            tabs["Daily"].to_excel(writer, sheet_name=f"{sname}_Daily", index=False)
            tabs["Weekly"].to_excel(writer, sheet_name=f"{sname}_Weekly", index=False)
            tabs["Monthly"].to_excel(writer, sheet_name=f"{sname}_Monthly", index=False)
            tabs["Monthly_MoM_%"].to_excel(writer, sheet_name=f"{sname}_Monthly_MoM_%", index=False)
            tabs["Quarterly"].to_excel(writer, sheet_name=f"{sname}_Quarterly", index=False)
            tabs["Yearly"].to_excel(writer, sheet_name=f"{sname}_Yearly", index=False)
            tabs["Yearly_YoY_%"].to_excel(writer, sheet_name=f"{sname}_Yearly_YoY_%", index=False)

        writer.book.save(out_bytes)  #save workbook

    st.download_button(  #download combined workbook
        label="Download Results Excel",
        data=out_bytes.getvalue(),
        file_name=uploaded_file.name.replace(".xlsx","_with_tables.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )  #end button

    st.success("Added Daily/Weekly/Monthly/Quarterly/Yearly tables (+MoM/YoY) to your workbook.")  #done
