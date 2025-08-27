import streamlit as st  #import streamlit for the web app
import pandas as pd  #import pandas for data handling
from io import BytesIO  #import BytesIO to hold files in memory

#-----------------------------
#Streamlit UI
#-----------------------------

st.set_page_config(page_title="CPI – Modality Averages", page_icon="page", layout="wide")  #set page config
st.title("Modality Daily/Weekly/Monthly/Quarterly/Yearly Calculator")  #set page title

uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])  #upload excel file

#-----------------------------
#Helpers
#-----------------------------

def _period_to_ts(s: pd.Series, freq: str) -> pd.Series:  #convert period strings to timestamps for chart x-axes
    p = pd.PeriodIndex(s.astype(str), freq=freq)  #build period index
    return p.to_timestamp()  #convert to start-of-period timestamps

#-----------------------------
#Processing logic
#-----------------------------

if uploaded_file:  #run only after a file is uploaded
    file_bytes = BytesIO(uploaded_file.read())  #read uploaded file into memory buffer

    all_sheets = pd.read_excel(file_bytes, sheet_name=None)  #load all sheets into a dict of DataFrames
    results = {}  #store per-sheet result tables and metadata here

    modality_days = {"US": 4}  #ultrasound has 4 scheduled days per week
    default_days = 5  #all other modalities have 5 scheduled days per week

    for sheet_name, df in all_sheets.items():  #loop through each sheet
        if df.empty:  #skip empty sheets
            continue  #move to next sheet

        date_col, modality_col = None, None  #placeholders for detected column names
        for c in df.columns:  #scan column names to detect date and modality/room
            cl = str(c).lower()  #lowercase for matching
            if ("date" in cl) and (date_col is None):  #pick the first date-like column
                date_col = c  #save date column
            if ("modality" in cl or "room" in cl) and (modality_col is None):  #pick the first modality/room column
                modality_col = c  #save modality column
        if not date_col or not modality_col:  #if we could not find needed columns
            continue  #skip this sheet

        df_clean = df.copy()  #work on a copy
        df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors="coerce")  #coerce date column to datetime
        df_clean = df_clean.dropna(subset=[date_col])  #drop rows with invalid/missing dates
        df_clean["Volume"] = 1  #count each row as one exam
        df_clean["StudyDate"] = df_clean[date_col].dt.date  #make a pure date column for grouping

        daily = (  #aggregate daily counts per modality
            df_clean
            .groupby([modality_col, "StudyDate"], as_index=False)["Volume"]
            .sum()
        )  #end daily

        daily["Week"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("W-MON")  #add week period (weeks starting Monday)
        weekly = daily.groupby([modality_col, "Week"], as_index=False)["Volume"].sum()  #weekly totals per modality

        daily["Month"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("M")  #add month period
        monthly = daily.groupby([modality_col, "Month"], as_index=False)["Volume"].sum()  #monthly totals per modality

        daily["Quarter"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("Q")  #add quarter period
        quarterly = daily.groupby([modality_col, "Quarter"], as_index=False)["Volume"].sum()  #quarterly totals per modality

        daily["Year"] = pd.to_datetime(daily["StudyDate"]).dt.to_period("Y")  #add year period
        yearly = daily.groupby([modality_col, "Year"], as_index=False)["Volume"].sum()  #yearly totals per modality

        avg_per_modality = []  #collect average rows
        for modality in daily[modality_col].unique():  #iterate each modality
            days_per_week = modality_days.get(str(modality), default_days)  #4 for US else 5
            mod_data = daily[daily[modality_col] == modality]  #subset for this modality
            avg_day = mod_data["Volume"].mean()  #mean daily count
            avg_week_total = mod_data.groupby("Week")["Volume"].sum().mean()  #mean weekly total
            avg_week_per_day = (avg_week_total / days_per_week) if days_per_week else None  #weekly total divided by scheduled days
            avg_month = mod_data.groupby("Month")["Volume"].sum().mean()  #mean monthly total
            avg_quarter = mod_data.groupby("Quarter")["Volume"].sum().mean()  #mean quarterly total
            avg_year = mod_data.groupby("Year")["Volume"].sum().mean()  #mean yearly total
            avg_per_modality.append([modality, avg_day, avg_week_per_day, avg_month, avg_quarter, avg_year])  #append row

        avg_df = pd.DataFrame(  #create averages table
            avg_per_modality,
            columns=["Modality", "Avg/Day", "Avg/Week(per scheduled day)", "Avg/Month", "Avg/Quarter", "Avg/Year"]
        )  #end averages table

        results[sheet_name] = {  #save tables and metadata for this sheet
            "Daily": daily,
            "Weekly": weekly,
            "Monthly": monthly,
            "Quarterly": quarterly,
            "Yearly": yearly,
            "Averages": avg_df,
            "_meta": {"modality_col": modality_col}  #remember column name for charts
        }  #end results per sheet

    #-----------------------------
    #Charts (interactive per sheet)
    #-----------------------------

    for sname, tables in results.items():  #iterate sheets for charts
        st.header(f"{sname}")  #show sheet name
        modality_col = tables["_meta"]["modality_col"]  #get modality column name
        all_modalities = sorted(tables["Daily"][modality_col].unique().tolist())  #list of modalities
        pick = st.selectbox(  #chooser for modality or All
            f"Select modality for charts ({sname})",
            options=["All"] + all_modalities,
            key=f"modpick_{sname}"
        )  #end selectbox

        with st.expander("Daily chart", expanded=False):  #daily chart expander
            df = tables["Daily"].copy()  #copy daily table
            df["StudyDate"] = pd.to_datetime(df["StudyDate"])  #ensure datetime
            if pick != "All":  #single modality
                df = df[df[modality_col] == pick]  #filter
                pivot = df.pivot_table(index="StudyDate", values="Volume", aggfunc="sum").sort_index()  #sum per day
            else:  #all modalities stacked as series
                pivot = df.pivot_table(index="StudyDate", columns=modality_col, values="Volume", aggfunc="sum").fillna(0).sort_index()  #wide table
            st.line_chart(pivot)  #render line chart

        with st.expander("Weekly chart", expanded=False):  #weekly chart expander
            df = tables["Weekly"].copy()  #copy weekly table
            df["WeekTS"] = _period_to_ts(df["Week"].astype(str), "W-MON")  #period→timestamp
            if pick != "All":  #single modality
                df = df[df[modality_col] == pick]  #filter
                pivot = df.groupby("WeekTS")["Volume"].sum().to_frame()  #sum per week
            else:  #all modalities
                pivot = df.pivot_table(index="WeekTS", columns=modality_col, values="Volume", aggfunc="sum").fillna(0)  #wide
            pivot = pivot.sort_index()  #sort by time
            st.bar_chart(pivot)  #render bar chart

        with st.expander("Monthly chart", expanded=False):  #monthly chart expander
            df = tables["Monthly"].copy()  #copy monthly table
            df["MonthTS"] = _period_to_ts(df["Month"].astype(str), "M")  #period→timestamp
            if pick != "All":
                df = df[df[modality_col] == pick]
                pivot = df.groupby("MonthTS")["Volume"].sum().to_frame()
            else:
                pivot = df.pivot_table(index="MonthTS", columns=modality_col, values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            st.bar_chart(pivot)

        with st.expander("Quarterly chart", expanded=False):  #quarterly chart expander
            df = tables["Quarterly"].copy()
            df["QuarterTS"] = _period_to_ts(df["Quarter"].astype(str), "Q")
            if pick != "All":
                df = df[df[modality_col] == pick]
                pivot = df.groupby("QuarterTS")["Volume"].sum().to_frame()
            else:
                pivot = df.pivot_table(index="QuarterTS", columns=modality_col, values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            st.bar_chart(pivot)

        with st.expander("Yearly chart", expanded=False):  #yearly chart expander
            df = tables["Yearly"].copy()
            df["YearTS"] = _period_to_ts(df["Year"].astype(str), "Y")
            if pick != "All":
                df = df[df[modality_col] == pick]
                pivot = df.groupby("YearTS")["Volume"].sum().to_frame()
            else:
                pivot = df.pivot_table(index="YearTS", columns=modality_col, values="Volume", aggfunc="sum").fillna(0)
            pivot = pivot.sort_index()
            st.bar_chart(pivot)

        st.divider()  #visual divider between sheets

    #-----------------------------
    #Write results to Excel
    #-----------------------------

    output = BytesIO()  #prepare an in-memory output workbook
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:  #open excel writer
        for sheet, tables in results.items():  #loop sheets
            tables["Daily"].to_excel(writer, sheet_name=f"{sheet}_Daily", index=False)  #write daily sheet
            tables["Weekly"].to_excel(writer, sheet_name=f"{sheet}_Weekly", index=False)  #write weekly sheet
            tables["Monthly"].to_excel(writer, sheet_name=f"{sheet}_Monthly", index=False)  #write monthly sheet
            tables["Quarterly"].to_excel(writer, sheet_name=f"{sheet}_Quarterly", index=False)  #write quarterly sheet
            tables["Yearly"].to_excel(writer, sheet_name=f"{sheet}_Yearly", index=False)  #write yearly sheet
            tables["Averages"].to_excel(writer, sheet_name=f"{sheet}_Averages", index=False)  #write averages sheet

    st.download_button(  #show a download button
        label="Download Results Excel",  #button label
        data=output.getvalue(),  #byte content of the workbook
        file_name="modality_calculations.xlsx",  #download filename
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"  #excel mime type
    )  #end download button

    st.success("Processed all sheets with daily/weekly/monthly/quarterly/yearly breakdowns and averages.")  #status message
