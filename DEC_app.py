import re  #import regex utilities
import streamlit as st  #import streamlit UI
import pandas as pd  #import pandas
from io import BytesIO  #import in-memory bytes buffer
from openpyxl import load_workbook  #import to append sheets to existing workbook

#UI
st.set_page_config(page_title="CPI – Exam Volumes Tables", page_icon="page", layout="wide")  #set page config
st.title("Exam Volumes → Daily / Weekly / Monthly / Quarterly / Yearly Tables")  #set app title
uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])  #file uploader

#-----CONFIG:room filtering/mapping----------------------------------------------------
EXCLUDE_PATTERNS = [r"^OUTSIDEREAD"]  #drop any room whose canonical name starts with OUTSIDEREAD
#OPTIONAL:if you want to keep only your core rooms, list patterns here; leave empty to allow all
INCLUDE_PATTERNS = []  #examples: [r"^CPICT", r"^CPIMRI", r"^CPIUS", r"^CPIXRY", r"^GMA_"]

#OPTIONAL:mappings from messy labels to a single final label after canonicalization
ALIAS_MAP = {  #left side and right side must be canonicalized strings (A–Z,0–9 only)
    #examples:
    # "CPICT": "CPICT1",  #map "CPICT" to "CPICT1"
    # "CPICT01": "CPICT1",
}

#-----HELPERS------------------------------------------------------------------------
def canonical_room(name:str)->str:
    #uppercase, strip, drop all non-alphanumerics (spaces/underscores/slashes/dashes)
    return re.sub(r"[^A-Z0-9]+","",str(name).upper().strip())

def apply_room_filters_and_alias(room:str)->str|None:
    #canonicalize first
    canon = canonical_room(room)
    #apply alias mapping
    canon = ALIAS_MAP.get(canon, canon)
    #exclude by pattern
    for pat in EXCLUDE_PATTERNS:
        if re.search(pat, canon, re.I):
            return None
    #include whitelist if provided
    if INCLUDE_PATTERNS:
        keep = any(re.search(p, canon, re.I) for p in INCLUDE_PATTERNS)
        if not keep:
            return None
    return canon

US_PATTERN = re.compile(r"US", re.I)  #detects any room containing 'US' after canonicalization

def scheduled_days_for_room(room_name:str)->int:
    #US modalities are 4-day weeks, all others 5-day
    return 4 if US_PATTERN.search(canonical_room(room_name)) else 5

def _norm(x:str)->str:
    #normalize header text for detection
    s = str(x).replace("\xa0"," ").strip().lower()
    return re.sub(r"\s+"," ", s)

def read_exam_volumes_two_cols(src:BytesIO, sheet_name:str="Exam Volumes")->pd.DataFrame|None:
    #read the sheet with no header to sniff the header row
    src.seek(0)  #rewind file
    raw = pd.read_excel(src, sheet_name=sheet_name, header=None, dtype=str)  #read raw data
    if raw.empty:  #return None if empty
        return None
    hdr_row = None  #placeholder for header row index
    #scan first 15 rows to find header containing both "room" and a study/scheduled date
    for i in range(min(15, len(raw))):  #loop over candidate header rows
        vals = [_norm(v) for v in raw.iloc[i].tolist()]  #normalize each cell
        has_room = any(("room" == v) or ("room" in v) for v in vals)  #room present
        has_date = any(v in ("study date","scheduled date") or ("study" in v and "date" in v) for v in vals)  #date present
        if has_room and has_date:  #if both present, found header row
            hdr_row = i
            break
    if hdr_row is None:  #no header found
        return None
    headers = [str(x).replace("\xa0"," ").strip() for x in raw.iloc[hdr_row].tolist()]  #clean header row
    col_map = {_norm(h): idx for idx, h in enumerate(headers)}  #map normalized header->index
    room_idx = None  #init
    date_idx = None  #init
    #find room column
    for k, idx in col_map.items():  #iterate all headers
        if ("room" == k) or ("room" in k):  #header contains room
            room_idx = idx
            break
    #find date column (prefer Study Date, else Scheduled Date, else any Study+Date)
    date_idx = col_map.get("study date", None)  #prefer study date
    if date_idx is None:  #if not present
        date_idx = col_map.get("scheduled date", None)  #try scheduled date
    if date_idx is None:  #fallback generic study+date
        for k, idx in col_map.items():
            if ("study" in k) and ("date" in k):
                date_idx = idx
                break
    if room_idx is None or date_idx is None:  #if still missing any column
        return None
    data = raw.iloc[hdr_row+1:, [room_idx, date_idx]].copy()  #slice to just two columns under header
    data.columns = ["Room","Study Date"]  #assign fixed names
    return data  #return two-column dataframe

def make_wide_table(long_df:pd.DataFrame, index_col:str, period_name:str)->pd.DataFrame:
    #pivot long [Room,Period,Volume] → wide table with rooms as columns
    pivot = long_df.pivot_table(index=index_col, columns="Room", values="Volume", aggfunc="sum").fillna(0)  #pivot counts
    pivot["Total Exams"] = pivot.sum(axis=1)  #add total column
    pivot = pivot.sort_index()  #sort by period
    pivot.index.name = period_name  #set index name
    return pivot.reset_index()  #return as flat table

def add_pct_change_table(wide_df:pd.DataFrame, period_col:str)->pd.DataFrame:
    #compute period-over-period percent change for each numeric column
    df = wide_df.copy()  #copy
    num_cols = [c for c in df.columns if c != period_col]  #numeric columns
    pct = df[num_cols].pct_change().round(4)  #percent change
    pct.insert(0, period_col, df[period_col])  #insert period column back
    pct.columns = [period_col] + [f"{c} %Δ" for c in num_cols]  #rename cols with %Δ
    return pct  #return pct-change table

def append_overall_average_row(wide_df:pd.DataFrame, period_col:str, label:str)->pd.DataFrame:
    #append one row with averages across all rows for each numeric column
    df = wide_df.copy()  #copy
    num_cols = [c for c in df.columns if c != period_col]  #numeric columns
    avg_row = {period_col: label}  #init label cell
    for c in num_cols:  #loop numeric columns
        avg_row[c] = df[c].mean()  #compute mean
    return pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)  #append and return

def round_numeric(df:pd.DataFrame, digits:int=1)->pd.DataFrame:
    #round all numeric columns to given digits
    out = df.copy()  #copy
    for c in out.columns:  #iterate columns
        if pd.api.types.is_numeric_dtype(out[c]):  #if numeric
            out[c] = out[c].round(digits)  #round values
    return out  #return rounded df

def business_days_range(start_date:pd.Timestamp, end_date:pd.Timestamp)->pd.DatetimeIndex:
    #generate Mon–Fri date range inclusive
    return pd.date_range(start_date, end_date, freq="B")  #B = business day

def insert_weekly_avg_rows(daily_wide:pd.DataFrame)->pd.DataFrame:
    #insert one "Weekly Avg" row after each complete Mon–Fri week; US rooms divide by 4, others by 5
    date_col = daily_wide.columns[0]  #first column is Date
    df = daily_wide.copy()  #copy
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce").dt.normalize()  #ensure datetime days
    df = df.dropna(subset=[date_col])  #drop bad dates
    all_bd = business_days_range(df[date_col].min(), df[date_col].max())  #full business-day span
    df = df.set_index(date_col).reindex(all_bd).fillna(0.0).rename_axis(date_col).reset_index()  #fill missing weekdays with zeros
    dow = df[date_col].dt.dayofweek  #0=Mon..6=Sun
    wk_start = df[date_col] - pd.to_timedelta(dow, unit="D")  #monday of week
    wk_end = wk_start + pd.to_timedelta(4, unit="D")  #friday of week
    df["_ws"] = wk_start  #store week start
    df["_we"] = wk_end  #store week end
    room_cols = [c for c in df.columns if c not in [date_col, "Total Exams", "_ws", "_we"]]  #room columns
    blocks = []  #accumulate output blocks
    for (ws, we), g in df.groupby(["_ws","_we"], sort=True):  #loop each week
        g2 = g.drop(columns=["_ws","_we"])  #drop helper cols
        g2 = g2[g2[date_col].dt.dayofweek <= 4]  #keep Mon–Fri only
        if len(g2) == 5:  #only complete weeks
            avg_row = {date_col: f"Weekly Avg {ws.date()}→{we.date()}"}  #label
            total = 0.0  #init total
            for col in room_cols:  #each modality column
                denom = scheduled_days_for_room(col)  #4 for US, else 5
                val = g2[col].sum() / denom if denom else 0.0  #weekly average for that modality
                avg_row[col] = val  #store value
                total += val  #add to total
            avg_row["Total Exams"] = total  #sum of per-room averages
            blocks.append(g2)  #append the week's daily rows
            blocks.append(pd.DataFrame([avg_row]))  #append the avg row
        else:
            blocks.append(g2)  #partial week: no avg row
    out = pd.concat(blocks, ignore_index=True)  #combine all blocks
    out[date_col] = out[date_col].apply(lambda x: x.date() if isinstance(x, pd.Timestamp) else x)  #format dates
    return out  #return with weekly avg rows

#-----MAIN-------------------------------------------------------------------------
if uploaded_file:  #run after a file is uploaded
    src = BytesIO(uploaded_file.read())  #read into memory
    src.seek(0)  #rewind
    xl = pd.ExcelFile(src)  #open workbook
    exam_sheet = None  #placeholder for the 'Exam Volumes' sheet name
    for nm in xl.sheet_names:  #scan sheet names
        if _norm(nm) == "exam volumes" or "exam volumes" in _norm(nm):  #match case-insensitively
            exam_sheet = nm
            break
    if not exam_sheet:  #error if not found
        st.error("Could not find a sheet named 'Exam Volumes'.")
        st.stop()

    df = read_exam_volumes_two_cols(src, exam_sheet)  #read only Room + Study Date
    if df is None or df.empty:  #validate
        st.error("Could not locate usable 'Room' and 'Study Date' columns in 'Exam Volumes'.")
        st.stop()

    #CLEAN:canonicalize, map, and filter rooms
    df["Room"] = df["Room"].apply(apply_room_filters_and_alias)  #apply filters and aliases
    df = df.dropna(subset=["Room"])  #drop rows removed by filters
    #DATE:parse Study Date to date
    df["StudyDate"] = pd.to_datetime(df["Study Date"], errors="coerce").dt.date  #to date
    df = df.dropna(subset=["StudyDate"])  #drop bad dates
    #COUNT:each row is one exam
    df["Volume"] = 1  #set each row to count 1

    #LONG:daily counts per room
    daily_long = df.groupby(["Room","StudyDate"], as_index=False)["Volume"].sum()  #group to daily totals

    #DAILY:wide pivot + weekly avg rows
    daily_wide = make_wide_table(daily_long, "StudyDate", "Date")  #pivot to wide
    daily_wide = insert_weekly_avg_rows(daily_wide)  #insert weekly averages
    daily_wide = round_numeric(daily_wide, 1)  #round numbers

    #WEEKLY:totals per week (weeks start Monday)
    daily_long["Week"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("W-MON")  #period weeks
    weekly_long = daily_long.groupby(["Room","Week"], as_index=False)["Volume"].sum()  #weekly totals
    weekly_wide = make_wide_table(weekly_long, "Week", "Week")  #wide weekly
    weekly_wide = append_overall_average_row(weekly_wide, "Week", "Average (Weekly)")  #append average row
    weekly_wide = round_numeric(weekly_wide, 1)  #round

    #MONTHLY
    daily_long["Month"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("M")  #month period
    monthly_long = daily_long.groupby(["Room","Month"], as_index=False)["Volume"].sum()  #sum by month
    monthly_wide = make_wide_table(monthly_long, "Month", "Month")  #wide month
    monthly_wide = append_overall_average_row(monthly_wide, "Month", "Average (Monthly)")  #avg row
    monthly_wide = round_numeric(monthly_wide, 1)  #round
    monthly_mom = add_pct_change_table(monthly_wide.drop(monthly_wide.index[-1]), "Month")  #MoM excluding avg row

    #QUARTERLY
    daily_long["Quarter"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Q")  #quarter period
    quarterly_long = daily_long.groupby(["Room","Quarter"], as_index=False)["Volume"].sum()  #sum by quarter
    quarterly_wide = make_wide_table(quarterly_long, "Quarter", "Quarter")  #wide quarter
    quarterly_wide = append_overall_average_row(quarterly_wide, "Quarter", "Average (Quarterly)")  #avg row
    quarterly_wide = round_numeric(quarterly_wide, 1)  #round

    #YEARLY
    daily_long["Year"] = pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Y")  #year period
    yearly_long = daily_long.groupby(["Room","Year"], as_index=False)["Volume"].sum()  #sum by year
    yearly_wide = make_wide_table(yearly_long, "Year", "Year")  #wide year
    yearly_wide = append_overall_average_row(yearly_wide, "Year", "Average (Yearly)")  #avg row
    yearly_wide = round_numeric(yearly_wide, 1)  #round
    yearly_yoy = add_pct_change_table(yearly_wide.drop(yearly_wide.index[-1]), "Year")  #YoY excluding avg row

    #SHOW:tables in the app
    st.header(exam_sheet)  #sheet header
    st.subheader("Daily (with Weekly Avg rows; Mon–Fri complete weeks only)")  #label
    st.dataframe(daily_wide, use_container_width=True)  #daily table
    st.subheader("Weekly (with bottom Average)")  #label
    st.dataframe(weekly_wide, use_container_width=True)  #weekly table
    st.subheader("Monthly (with bottom Average)")  #label
    st.dataframe(monthly_wide, use_container_width=True)  #monthly table
    st.subheader("Monthly (MoM % change)")  #label
    st.dataframe(monthly_mom, use_container_width=True)  #MoM table
    st.subheader("Quarterly (with bottom Average)")  #label
    st.dataframe(quarterly_wide, use_container_width=True)  #quarterly table
    st.subheader("Yearly (with bottom Average)")  #label
    st.dataframe(yearly_wide, use_container_width=True)  #yearly table
    st.subheader("Yearly (YoY % change)")  #label
    st.dataframe(yearly_yoy, use_container_width=True)  #YoY table

    #WRITE:append new sheets to the same workbook
    src.seek(0)  #rewind again for openpyxl
    wb = load_workbook(filename=src)  #load workbook
    #remove prior result sheets to avoid duplicates
    for ws in list(wb.worksheets):  #loop sheets
        if ws.title.endswith(("_Daily","_Weekly","_Monthly","_Monthly_MoM_%","_Quarterly","_Yearly","_Yearly_YoY_%")):
            wb.remove(ws)
    out = BytesIO()  #prepare output buffer
    with pd.ExcelWriter(out, engine="openpyxl") as writer:  #open writer with existing workbook
        writer.book = wb  #attach existing book
        writer.sheets = {ws.title: ws for ws in wb.worksheets}  #existing sheets map
        daily_wide.to_excel(writer, sheet_name=f"{exam_sheet}_Daily", index=False)  #write daily
        weekly_wide.to_excel(writer, sheet_name=f"{exam_sheet}_Weekly", index=False)  #write weekly
        monthly_wide.to_excel(writer, sheet_name=f"{exam_sheet}_Monthly", index=False)  #write monthly
        monthly_mom.to_excel(writer, sheet_name=f"{exam_sheet}_Monthly_MoM_%", index=False)  #write MoM
        quarterly_wide.to_excel(writer, sheet_name=f"{exam_sheet}_Quarterly", index=False)  #write quarterly
        yearly_wide.to_excel(writer, sheet_name=f"{exam_sheet}_Yearly", index=False)  #write yearly
        yearly_yoy.to_excel(writer, sheet_name=f"{exam_sheet}_Yearly_YoY_%", index=False)  #write YoY
        writer.book.save(out)  #save to buffer
    st.download_button(  #download button
        label="Download Results Excel",  #label
        data=out.getvalue(),  #data bytes
        file_name=uploaded_file.name.replace(".xlsx","_with_tables.xlsx"),  #filename
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"  #mime type
    )  #end download button
