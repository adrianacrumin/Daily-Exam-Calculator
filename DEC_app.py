import streamlit as st
import numpy as np
import matplotlib.pyplot as plt


#-----------------------------
#Streamlit UI
#-----------------------------

st.set_page_config(page_title="CPI – Modality Averages", page_icon="pg", layout="wide")  #set page config
st.title("Modality Daily/Weekly/Monthly/Quarterly/Yearly Calculator")  #set page title

uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])  #upload excel file

#-----------------------------
#Processing logic
#-----------------------------

if uploaded_file:  #check if file uploaded
    file_bytes = BytesIO(uploaded_file.read())  #read uploaded file into memory

    all_sheets = pd.read_excel(file_bytes, sheet_name=None)  #load all sheets into dictionary
    results = {}  #store results for each sheet

    modality_days = {"US": 4}  #business rule: US modality has 4 days/week
    default_days = 5  #all other modalities default to 5 days/week

    for sheet_name, df in all_sheets.items():  #loop through sheets
        if df.empty:  #skip if sheet is empty
            continue

        date_col, modality_col = None, None  #placeholders for column names
        for c in df.columns:  #detect date and modality columns
            if "date" in str(c).lower():
                date_col = c
            if "modality" in str(c).lower() or "room" in str(c).lower():
                modality_col = c
        if not date_col or not modality_col:  #skip if columns not found
            continue

        df_clean = df.copy()  #make copy of data
        df_clean[date_col] = pd.to_datetime(df_clean[date_col], errors="coerce")  #convert dates
        df_clean = df_clean.dropna(subset=[date_col])  #drop rows without valid date
        df_clean["Volume"] = 1  #each row counts as 1

        daily = df_clean.groupby([modality_col, df_clean[date_col].dt.date], as_index=False)["Volume"].sum()  #daily totals

        daily["Week"] = pd.to_datetime(daily[date_col]).dt.to_period("W-MON")  #weekly period
        weekly = daily.groupby([modality_col, "Week"], as_index=False)["Volume"].sum()  #weekly totals

        daily["Month"] = pd.to_datetime(daily[date_col]).dt.to_period("M")  #monthly period
        monthly = daily.groupby([modality_col, "Month"], as_index=False)["Volume"].sum()  #monthly totals

        daily["Quarter"] = pd.to_datetime(daily[date_col]).dt.to_period("Q")  #quarterly period
        quarterly = daily.groupby([modality_col, "Quarter"], as_index=False)["Volume"].sum()  #quarterly totals

        daily["Year"] = pd.to_datetime(daily[date_col]).dt.to_period("Y")  #yearly period
        yearly = daily.groupby([modality_col, "Year"], as_index=False)["Volume"].sum()  #yearly totals

        avg_per_modality = []  #store averages per modality
        for modality in daily[modality_col].unique():  #loop through modalities
            days_per_week = modality_days.get(modality, default_days)  #get scheduled days/week
            mod_data = daily[daily[modality_col] == modality]  #filter modality data
            avg_day = mod_data["Volume"].mean()  #average per day
            avg_week = mod_data.groupby("Week")["Volume"].sum().mean() / days_per_week  #average per week
            avg_month = mod_data.groupby("Month")["Volume"].sum().mean()  #average per month
            avg_quarter = mod_data.groupby("Quarter")["Volume"].sum().mean()  #average per quarter
            avg_year = mod_data.groupby("Year")["Volume"].sum().mean()  #average per year
            avg_per_modality.append([modality, avg_day, avg_week, avg_month, avg_quarter, avg_year])

        avg_df = pd.DataFrame(  #make dataframe of averages
            avg_per_modality,
            columns=["Modality", "Avg/Day", "Avg/Week", "Avg/Month", "Avg/Quarter", "Avg/Year"]
        )

        results[sheet_name] = {  #store all tables for this sheet
            "Daily": daily,
            "Weekly": weekly,
            "Monthly": monthly,
            "Quarterly": quarterly,
            "Yearly": yearly,
            "Averages": avg_df
        }

    output = BytesIO()  #output file in memory
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:  #write results to Excel
        for sheet, tables in results.items():
            tables["Daily"].to_excel(writer, sheet_name=f"{sheet}_Daily", index=False)
            tables["Weekly"].to_excel(writer, sheet_name=f"{sheet}_Weekly", index=False)
            tables["Monthly"].to_excel(writer, sheet_name=f"{sheet}_Monthly", index=False)
            tables["Quarterly"].to_excel(writer, sheet_name=f"{sheet}_Quarterly", index=False)
            tables["Yearly"].to_excel(writer, sheet_name=f"{sheet}_Yearly", index=False)
            tables["Averages"].to_excel(writer, sheet_name=f"{sheet}_Averages", index=False)

    st.download_button(  #download button for results
        label="Download Results Excel",
        data=output.getvalue(),
        file_name="modality_calculations.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Processed all sheets with daily/weekly/monthly/quarterly/yearly breakdowns and averages.")
