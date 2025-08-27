import re
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

#UI
st.set_page_config(page_title="CPI – Exam Volumes Tables", page_icon="page", layout="wide")
st.title("Exam Volumes → Daily / Weekly / Monthly / Quarterly / Yearly Tables")
uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])

#helpers
def canonical_room(name:str)->str:
    return re.sub(r"[^A-Za-z0-9]+","",str(name).upper().strip())

US_PATTERN=re.compile(r"US",re.I)
def scheduled_days_for_room(room_name:str)->int:
    return 4 if US_PATTERN.search(canonical_room(room_name)) else 5

def _norm(x:str)->str:
    s=str(x).replace("\xa0"," ").strip().lower()
    return re.sub(r"\s+"," ",s)

def read_exam_volumes_two_cols(src:BytesIO,sheet_name:str="Exam Volumes"):
    src.seek(0)
    raw=pd.read_excel(src,sheet_name=sheet_name,header=None,dtype=str)
    if raw.empty:
        return None
    hdr_row=None
    #find row that has both "room" and ("study date" or "scheduled date")
    for i in range(min(15,len(raw))):
        vals=[_norm(v) for v in raw.iloc[i].tolist()]
        has_room=any(("room"==v) or ("room" in v) for v in vals)
        has_date=any(v in ("study date","scheduled date") or ("study" in v and "date" in v) for v in vals)
        if has_room and has_date:
            hdr_row=i
            break
    if hdr_row is None:
        return None
    #identify the exact columns for Room and Study/Scheduled Date
    headers=[str(x).replace("\xa0"," ").strip() for x in raw.iloc[hdr_row].tolist()]
    col_map={_norm(h):idx for idx,h in enumerate(headers)}
    room_idx=None
    date_idx=None
    for k,idx in col_map.items():
        if ("room"==k) or ("room" in k):
            room_idx=idx
            break
    date_idx=col_map.get("study date",None)
    if date_idx is None:
        date_idx=col_map.get("scheduled date",None)
    if date_idx is None:
        for k,idx in col_map.items():
            if ("study" in k) and ("date" in k):
                date_idx=idx
                break
    if room_idx is None or date_idx is None:
        return None
    data=raw.iloc[hdr_row+1:, [room_idx, date_idx]].copy()
    data.columns=["Room","Study Date"]
    return data

def make_wide_table(long_df:pd.DataFrame,index_col:str,period_name:str)->pd.DataFrame:
    pivot=long_df.pivot_table(index=index_col,columns="Room",values="Volume",aggfunc="sum").fillna(0)
    pivot["Total Exams"]=pivot.sum(axis=1)
    pivot=pivot.sort_index()
    pivot.index.name=period_name
    return pivot.reset_index()

def add_pct_change_table(wide_df:pd.DataFrame,period_col:str)->pd.DataFrame:
    df=wide_df.copy()
    num_cols=[c for c in df.columns if c!=period_col]
    pct=df[num_cols].pct_change().round(4)
    pct.insert(0,period_col,df[period_col])
    pct.columns=[period_col]+[f"{c} %Δ" for c in num_cols]
    return pct

def append_overall_average_row(wide_df:pd.DataFrame,period_col:str,label:str)->pd.DataFrame:
    df=wide_df.copy()
    num_cols=[c for c in df.columns if c!=period_col]
    avg_row={period_col:label}
    for c in num_cols:
        avg_row[c]=df[c].mean()
    return pd.concat([df,pd.DataFrame([avg_row])],ignore_index=True)

def round_numeric(df:pd.DataFrame,digits:int=1)->pd.DataFrame:
    out=df.copy()
    for c in out.columns:
        if pd.api.types.is_numeric_dtype(out[c]):
            out[c]=out[c].round(digits)
    return out

def business_days_range(start_date:pd.Timestamp,end_date:pd.Timestamp)->pd.DatetimeIndex:
    return pd.date_range(start_date,end_date,freq="B")  #Mon–Fri

def insert_weekly_avg_rows(daily_wide:pd.DataFrame)->pd.DataFrame:
    date_col=daily_wide.columns[0]
    df=daily_wide.copy()
    df[date_col]=pd.to_datetime(df[date_col],errors="coerce").dt.normalize()
    df=df.dropna(subset=[date_col])
    #fill missing weekdays with zeros
    all_bd=business_days_range(df[date_col].min(),df[date_col].max())
    df=df.set_index(date_col).reindex(all_bd).fillna(0.0).rename_axis(date_col).reset_index()
    #compute Mon..Fri window
    dow=df[date_col].dt.dayofweek
    wk_start=df[date_col]-pd.to_timedelta(dow,unit="D")
    wk_end=wk_start+pd.to_timedelta(4,unit="D")
    df["_ws"]=wk_start; df["_we"]=wk_end
    room_cols=[c for c in df.columns if c not in [date_col,"Total Exams","_ws","_we"]]
    blocks=[]
    for (ws,we),g in df.groupby(["_ws","_we"],sort=True):
        g2=g.drop(columns=["_ws","_we"])
        g2=g2[g2[date_col].dt.dayofweek<=4]
        if len(g2)==5:
            avg_row={date_col:f"Weekly Avg {ws.date()}→{we.date()}"}
            total=0.0
            for col in room_cols:
                denom=scheduled_days_for_room(col)
                val=g2[col].sum()/denom if denom else 0.0
                avg_row[col]=val; total+=val
            avg_row["Total Exams"]=total
            blocks.append(g2); blocks.append(pd.DataFrame([avg_row]))
        else:
            blocks.append(g2)  #partial week: no avg row
    out=pd.concat(blocks,ignore_index=True)
    out[date_col]=out[date_col].apply(lambda x: x.date() if isinstance(x,pd.Timestamp) else x)
    return out

#main
if uploaded_file:
    src=BytesIO(uploaded_file.read())
    #read ONLY the "Exam Volumes" sheet (case-insensitive)
    src.seek(0)
    xl=pd.ExcelFile(src)
    exam_sheet=None
    for nm in xl.sheet_names:
        if _norm(nm)=="exam volumes" or "exam volumes" in _norm(nm):
            exam_sheet=nm; break
    if not exam_sheet:
        st.error("Could not find a sheet named 'Exam Volumes'. Please upload a workbook that has that sheet.")
        st.stop()

    df=read_exam_volumes_two_cols(src,exam_sheet)
    if df is None or df.empty:
        st.error("Could not locate usable 'Room' and 'Study Date' columns in 'Exam Volumes'.")
        st.stop()

    #clean 2 columns
    df["Room"]=df["Room"].apply(canonical_room)
    df["StudyDate"]=pd.to_datetime(df["Study Date"],errors="coerce").dt.date
    df=df.dropna(subset=["StudyDate"])
    df["Volume"]=1

    #daily long
    daily_long=(df.groupby(["Room","StudyDate"],as_index=False)["Volume"].sum())

    #daily wide + weekly avg rows (Mon–Fri complete weeks only)
    daily_wide=make_wide_table(daily_long,"StudyDate","Date")
    daily_wide=insert_weekly_avg_rows(daily_wide)
    daily_wide=round_numeric(daily_wide,1)

    #weekly totals (Mon-start), plus bottom average row
    daily_long["Week"]=pd.to_datetime(daily_long["StudyDate"]).dt.to_period("W-MON")
    weekly_long=daily_long.groupby(["Room","Week"],as_index=False)["Volume"].sum()
    weekly_wide=make_wide_table(weekly_long,"Week","Week")
    weekly_wide=append_overall_average_row(weekly_wide,"Week","Average (Weekly)")
    weekly_wide=round_numeric(weekly_wide,1)

    #monthly
    daily_long["Month"]=pd.to_datetime(daily_long["StudyDate"]).dt.to_period("M")
    monthly_long=daily_long.groupby(["Room","Month"],as_index=False)["Volume"].sum()
    monthly_wide=make_wide_table(monthly_long,"Month","Month")
    monthly_wide=append_overall_average_row(monthly_wide,"Month","Average (Monthly)")
    monthly_wide=round_numeric(monthly_wide,1)
    monthly_mom=add_pct_change_table(monthly_wide.drop(monthly_wide.index[-1]),"Month")

    #quarterly
    daily_long["Quarter"]=pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Q")
    quarterly_long=daily_long.groupby(["Room","Quarter"],as_index=False)["Volume"].sum()
    quarterly_wide=make_wide_table(quarterly_long,"Quarter","Quarter")
    quarterly_wide=append_overall_average_row(quarterly_wide,"Quarter","Average (Quarterly)")
    quarterly_wide=round_numeric(quarterly_wide,1)

    #yearly
    daily_long["Year"]=pd.to_datetime(daily_long["StudyDate"]).dt.to_period("Y")
    yearly_long=daily_long.groupby(["Room","Year"],as_index=False)["Volume"].sum()
    yearly_wide=make_wide_table(yearly_long,"Year","Year")
    yearly_wide=append_overall_average_row(yearly_wide,"Year","Average (Yearly)")
    yearly_wide=round_numeric(yearly_wide,1)
    yearly_yoy=add_pct_change_table(yearly_wide.drop(yearly_wide.index[-1]),"Year")

    #show in app
    st.header(exam_sheet)
    st.subheader("Daily (with Weekly Avg rows; Mon–Fri complete weeks only)")
    st.dataframe(daily_wide,use_container_width=True)
    st.subheader("Weekly (with bottom Average)")
    st.dataframe(weekly_wide,use_container_width=True)
    st.subheader("Monthly (with bottom Average)")
    st.dataframe(monthly_wide,use_container_width=True)
    st.subheader("Monthly (MoM % change)")
    st.dataframe(monthly_mom,use_container_width=True)
    st.subheader("Quarterly (with bottom Average)")
    st.dataframe(quarterly_wide,use_container_width=True)
    st.subheader("Yearly (with bottom Average)")
    st.dataframe(yearly_wide,use_container_width=True)
    st.subheader("Yearly (YoY % change)")
    st.dataframe(yearly_yoy,use_container_width=True)

    #append sheets back into original workbook
    src.seek(0)
    wb=load_workbook(filename=src)
    for ws in list(wb.worksheets):
        if ws.title.endswith(("_Daily","_Weekly","_Monthly","_Monthly_MoM_%","_Quarterly","_Yearly","_Yearly_YoY_%")):
            wb.remove(ws)

    out=BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        writer.book=wb
        writer.sheets={ws.title:ws for ws in wb.worksheets}
        daily_wide.to_excel(writer,sheet_name=f"{exam_sheet}_Daily",index=False)
        weekly_wide.to_excel(writer,sheet_name=f"{exam_sheet}_Weekly",index=False)
        monthly_wide.to_excel(writer,sheet_name=f"{exam_sheet}_Monthly",index=False)
        monthly_mom.to_excel(writer,sheet_name=f"{exam_sheet}_Monthly_MoM_%",index=False)
        quarterly_wide.to_excel(writer,sheet_name=f"{exam_sheet}_Quarterly",index=False)
        yearly_wide.to_excel(writer,sheet_name=f"{exam_sheet}_Yearly",index=False)
        yearly_yoy.to_excel(writer,sheet_name=f"{exam_sheet}_Yearly_YoY_%",index=False)
        writer.book.save(out)

    st.download_button(
        label="Download Results Excel",
        data=out.getvalue(),
        file_name=uploaded_file.name.replace(".xlsx","_with_tables.xlsx"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
