import io, re, shutil, zipfile, tempfile, os
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pivot_refresh import refresh_workbook

st.set_page_config(page_title="Auto Pivot Report", page_icon="📊", layout="wide")

# Password protection is intentionally kept. Set the same password in
# Streamlit Secrets as APP_PASSWORD (or environment variable APP_PASSWORD).
APP_PASSWORD = st.secrets.get("APP_PASSWORD", os.getenv("APP_PASSWORD", ""))

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False


if APP_PASSWORD:
    if not st.session_state.authenticated:
        st.markdown("""
        <style>
        .login-wrap{max-width:430px;margin:8vh auto 0 auto;}
        .login-head{background:#000;color:#fff;padding:22px 26px;border-radius:14px 14px 0 0;font-size:24px;font-weight:800;}
        .login-body{background:#fff;border:1px solid #dfe3e8;border-top:0;padding:26px;border-radius:0 0 14px 14px;box-shadow:0 8px 28px rgba(0,0,0,.08);}
        .login-note{color:#667085;font-size:14px;margin:0 0 16px 0;}
        div[data-testid="stTextInput"] input{border-radius:9px !important;min-height:44px !important;}
        .login-btn div.stButton > button{background:#000 !important;color:#fff !important;border:0 !important;border-radius:9px !important;min-height:44px !important;font-weight:800 !important;}
        </style>
        <div class="login-wrap">
          <div class="login-head">🔐 Daily Report Login</div>
          <div class="login-body">
            <div class="login-note">Password ထည့်ပြီး Daily Report System ကို ဝင်ပါ</div>
          </div>
        </div>
        """, unsafe_allow_html=True)
        pw = st.text_input("Password", type="password", label_visibility="collapsed", placeholder="Enter password")
        st.markdown('<div class="login-btn">', unsafe_allow_html=True)
        login = st.button("LOGIN", type="primary", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        if login:
            if pw == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Password မှားနေပါတယ်")
        st.stop()
else:
    st.warning("APP_PASSWORD ကို Streamlit Secrets မှာ ထည့်ထားပါ။ Password protection ကို မဖယ်ထားပါ။")
    st.stop()

RAW_SHEETS = ["R1","R1S","R2","R6","R1 OC","R6 OC"]
R2S_TOWNSHIPS = {"kamaryut","dagon myothit north","sanchaung","kyeemyindaing","lanmadaw","pabedan","pazundaung","kyauktada","latha","botahtaung"}

def norm(v): return "" if pd.isna(v) else str(v).strip()
def key(v): return re.sub(r"\s+"," ",norm(v)).lower()

def task_type(v):
    t=norm(v).upper()
    return "TKT" if "TKT" in t else ("POI" if "POI" in t else "")

def classification(t, service, root):
    t,s,r=norm(t).upper(),norm(service).upper(),norm(root).upper()
    if t=="TKT":
        if "FR-SLA" in s or "BIZ" in s: return "FR-SLA / Biz TKT"
        if "INSTALLATION TYPE CHANGE" in r: return "FR-SLA / Biz TKT"
        if "FT-SBS" in r: return "FT-SBS"
        return "Other TKT"
    if t=="POI": return "Install Express" if "EXPRESS" in r else "Install POI"
    return ""

def process(df, is_r2=False):
    df=df.dropna(how="all").reset_index(drop=True).copy()
    for c in ["TKT / POI","Classification","RC","RE"]:
        if c in df.columns: df.drop(columns=c,inplace=True)

    service=next((c for c in df.columns if key(c)=="service area"),None)
    task=next((c for c in df.columns if key(c)=="task no"),None)
    root=next((c for c in df.columns if key(c)=="root cause"),None)
    if not task:
        raise ValueError("Task No column မတွေ့ပါ")
    tt=df[task].map(task_type)
    cc=[classification(t,s,r) for t,s,r in zip(tt,
        df[service].map(norm) if service else [""]*len(df),
        df[root].map(norm) if root else [""]*len(df))]
    pos=df.columns.get_loc(service)+1 if service else 0
    df.insert(pos,"TKT / POI",tt); df.insert(pos+1,"Classification",cc)

    car=next((c for c in df.columns if key(c)=="car id"),None)
    eng=next((c for c in df.columns if key(c)=="lan engineer name"),None)
    repay=next((c for c in df.columns if key(c)=="repay status"),None)
    if car and eng:
        tmp=pd.DataFrame({"car":df[car].map(norm),"eng":df[eng].map(norm)})
        if repay:
            excluded=df[repay].map(norm).str.lower().str.contains(r"engr\s*leave|route\s*cancel",regex=True,na=False)
            tmp=tmp[~excluded]
        tmp=tmp[(tmp.car!="")&(tmp.eng!="")]
        counts=tmp.groupby("car").eng.apply(lambda s:len(set(s))).to_dict()
        rc=[]
        for x in df[car].map(norm):
            n=counts.get(x,0)
            rc.append("" if not x else ("Car Share Plus" if n>=3 else "Car Share" if n==2 else "One Car" if n==1 else ""))
        df.insert(df.columns.get_loc(car),"RC",rc)

    if is_r2:
        township=next((c for c in df.columns if key(c)=="township"),None)
        if township:
            vals=[]
            for x in df[township]:
                t=key(x)
                if not t:
                    vals.append(""); continue
                base=re.split(r"\s*\(|\s+ward\b",t,maxsplit=1)[0].strip()
                vals.append("R2S" if (t in R2S_TOWNSHIPS or base in R2S_TOWNSHIPS) else "R2")
            df.insert(df.columns.get_loc(township)+1,"RE",vals)
    return df

def build_workbook(raw_bytes, template_path):
    raw_stream=io.BytesIO(raw_bytes)
    xls=pd.ExcelFile(raw_stream)
    out=Path(tempfile.mktemp(suffix=".xlsx"))
    shutil.copy2(template_path,out)
    wb=load_workbook(out)
    used={}

    aliases={"R1 OC":["R1 OC","R1OC"],"R6 OC":["R6 OC","R6OC"]}
    for logical in RAW_SHEETS:
        actual=logical
        if actual not in xls.sheet_names and logical in aliases:
            actual=next((a for a in aliases[logical] if a in xls.sheet_names),None)
        if actual not in xls.sheet_names or logical not in wb.sheetnames:
            continue
        raw_stream.seek(0)
        df=process(pd.read_excel(raw_stream,sheet_name=actual),is_r2=(logical=="R2"))
        ws=wb[logical]
        headers=[ws.cell(1,c).value for c in range(1,ws.max_column+1)]
        hmap={key(v):i+1 for i,v in enumerate(headers) if v is not None}
        # Clear every old data row first: no previous-day rows remain.
        if ws.max_row >= 2:
            for row in ws.iter_rows(min_row=2,max_row=ws.max_row,max_col=ws.max_column):
                for cell in row:
                    cell.value=None
        for ri,row in enumerate(df.to_dict("records"),start=2):
            for col,val in row.items():
                ci=hmap.get(key(col))
                if ci:
                    cell=ws.cell(ri,ci)
                    cell.value=val
                    if key(col)=="task status": cell.number_format="General"
        for hname,ci in hmap.items():
            if hname=="task status":
                for rr in range(2,max(2,len(df)+1)):
                    ws.cell(rr,ci).number_format="General"
        ws.auto_filter.ref=f"A1:{get_column_letter(ws.max_column)}{max(2,len(df)+1)}"
        used[logical]=len(df)

    wb.save(out); wb.close()

    # Patch Pivot cache source ranges and refreshOnLoad while preserving template Pivot layout.
    tmp=out.with_suffix(".patched.xlsx")
    with zipfile.ZipFile(out,"r") as zin, zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zout:
        cache_map={1:"R1",2:"R1S",3:"R2",4:"R6",5:"R1 OC",6:"R6 OC"}
        for item in zin.infolist():
            data=zin.read(item.filename)
            m=re.fullmatch(r"xl/pivotCache/pivotCacheDefinition([1-6])\.xml",item.filename)
            if m:
                s=cache_map[int(m.group(1))]
                wbx=load_workbook(out,read_only=True,data_only=False)
                wsx=wbx[s]
                ref=f"A1:{get_column_letter(wsx.max_column)}{wsx.max_row}"
                wbx.close()
                xml=data.decode("utf-8")
                xml=re.sub(r'<worksheetSource ref="[^"]+" sheet="([^"]+)"',lambda mm:f'<worksheetSource ref="{ref}" sheet="{mm.group(1)}"',xml,count=1)
                if "refreshOnLoad=" not in xml:
                    xml=xml.replace("<pivotCacheDefinition ","<pivotCacheDefinition refreshOnLoad=\"1\" ",1)
                data=xml.encode()
            zout.writestr(item,data)
    out.unlink(); tmp.replace(out)

    # Rebuild/refresh the real PivotTables with LibreOffice before returning the file.
    # This prevents Excel's "PivotTable report is invalid" state and keeps Count numeric.
    refresh_workbook(out)
    return out.read_bytes(), used


# -----------------------------
# UI ONLY — processing / Excel / Pivot code above is unchanged
# -----------------------------
import time

st.markdown("""
<style>
html,body,[data-testid="stAppViewContainer"]{background:#f7f8fa !important;}
.block-container{max-width:1180px !important;padding-top:2rem !important;padding-bottom:2.5rem !important;}
.app-hero{background:#000;color:#fff;border-radius:14px;padding:22px 28px;margin-bottom:18px;}
.app-hero .title{font-size:32px;line-height:1.15;font-weight:850;margin:0;}
.app-hero .sub{font-size:15px;color:#e6e6e6;margin-top:6px;}
.workflow{margin-top:15px;font-size:14px;font-weight:700;color:#fff;opacity:.9;}
.ui-section{background:#fff;border:1px solid #e1e4e8;border-radius:12px;padding:20px 22px;margin:0 0 16px 0;box-shadow:0 3px 12px rgba(0,0,0,.045);}
.section-title{font-size:18px;font-weight:850;color:#000;margin:0 0 5px 0;}
.section-sub{color:#667085;font-size:13px;margin-bottom:14px;}
.runner-wrap{background:#fff;border:1px solid #dfe3e8;border-radius:12px;padding:16px 18px 14px 18px;margin:12px 0 18px 0;box-shadow:0 3px 12px rgba(0,0,0,.045);}
.runner-label{font-size:14px;font-weight:800;color:#000;margin-bottom:10px;}
.runner-track{width:100%;height:18px;border-radius:999px;background:#eceff2;border:1px solid #d8dce1;overflow:hidden;position:relative;}
.runner{position:absolute;left:-44px;top:-8px;font-size:31px;line-height:31px;animation:runSlow 4.8s linear infinite;}
@keyframes runSlow{0%{left:-44px;}100%{left:calc(100% + 8px);}}
.runner-hint{color:#667085;font-size:12px;margin-top:8px;}
.sheet-card{background:#fff;border:1px solid #dfe3e8;border-radius:11px;overflow:hidden;margin-bottom:16px;box-shadow:0 3px 12px rgba(0,0,0,.045);}
.sheet-head{background:#000;color:#fff;padding:12px 16px;font-size:19px;font-weight:850;}
.task-head{background:#fff;color:#000;border-bottom:1px solid #e1e4e8;padding:10px 16px;font-size:13px;font-weight:850;letter-spacing:.4px;}
.task-row{display:flex;align-items:center;justify-content:space-between;min-height:42px;padding:8px 16px;border-bottom:1px solid #edf0f2;background:#fff;}
.task-name{color:#101828;font-size:14px;font-weight:700;padding-right:12px;}
.task-count{color:#000;font-size:15px;font-weight:850;min-width:34px;text-align:right;}
.task-total{display:flex;align-items:center;justify-content:space-between;min-height:42px;padding:8px 16px;background:#f1f3f5;color:#000;font-weight:850;}
.empty-note{color:#667085;padding:14px 16px;font-size:13px;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="app-hero">
  <div class="title">📊 DAILY REPORT AUTOMATION</div>
  <div class="sub">Automated Excel Reporting System</div>
  <div class="workflow">Raw Excel &nbsp;→&nbsp; Process &nbsp;→&nbsp; Report</div>
</div>
""", unsafe_allow_html=True)

template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx")
if not template.exists():
    st.error("Template Excel မတွေ့ပါ။ app.py နဲ့ template Excel ကို GitHub repo တစ်ခုထဲမှာထားပါ။")
    st.stop()

st.markdown("""
<div class="ui-section">
  <div class="section-title">📁 UPLOAD RAW EXCEL</div>
  <div class="section-sub">Upload today's Raw Excel file</div>
</div>
""", unsafe_allow_html=True)

uploaded=st.file_uploader("Choose Excel File",type=["xlsx"],label_visibility="collapsed")
generate=st.button("🚀 GENERATE REPORT",type="primary",use_container_width=True)

if uploaded and generate:
    runner=st.empty()
    runner.markdown("""
    <div class="runner-wrap">
      <div class="runner-label">Generating report… Please wait</div>
      <div class="runner-track"><div class="runner">🏃‍♂️</div></div>
      <div class="runner-hint">Processing Raw Excel and preparing the final report</div>
    </div>
    """,unsafe_allow_html=True)
    started=time.time()
    try:
        data,counts=build_workbook(uploaded.getvalue(),template)
        remaining=4.8-(time.time()-started)
        if remaining>0: time.sleep(remaining)
        st.session_state["out"]=data
        st.session_state["counts"]=counts
        st.session_state["raw_bytes"]=uploaded.getvalue()
        runner.empty()
        st.success("✅ Report Generated")
    except Exception as e:
        runner.empty()
        st.error(f"Error: {e}")

if "out" in st.session_state:
    status_by_sheet={}
    try:
        raw_bytes=st.session_state.get("raw_bytes")
        if raw_bytes:
            xls=pd.ExcelFile(io.BytesIO(raw_bytes))
            aliases={"R1 OC":["R1 OC","R1OC"],"R6 OC":["R6 OC","R6OC"]}
            for logical in RAW_SHEETS:
                actual=logical
                if actual not in xls.sheet_names and logical in aliases:
                    actual=next((a for a in aliases[logical] if a in xls.sheet_names),None)
                if actual and actual in xls.sheet_names:
                    dfx=process(pd.read_excel(io.BytesIO(raw_bytes),sheet_name=actual),is_r2=(logical=="R2"))
                    status_col=next((c for c in dfx.columns if key(c)=="task status"),None)
                    if status_col:
                        s=dfx[status_col].map(norm)
                        s=s[s!=""]
                        status_by_sheet[logical]=s.value_counts().to_dict()
                    else:
                        status_by_sheet[logical]={}
    except Exception:
        status_by_sheet={}

    st.markdown("""
    <div class="ui-section">
      <div class="section-title">📋 REPORT STATUS</div>
      <div class="section-sub">Current report is ready</div>
    </div>
    """,unsafe_allow_html=True)

    cols=st.columns(3,gap="medium")
    for idx,logical in enumerate(RAW_SHEETS):
        cmap=status_by_sheet.get(logical,{})
        total=int(sum(cmap.values()))
        rows=[]
        for status,count in cmap.items():
            if norm(status):
                rows.append(f"""
                <div class="task-row">
                  <div class="task-name">{status}</div>
                  <div class="task-count">{int(count)}</div>
                </div>
                """)
        if not rows:
            rows.append('<div class="empty-note">No non-blank Task Status</div>')
        card=f"""
        <div class="sheet-card">
          <div class="sheet-head">📊 {logical}</div>
          <div class="task-head">TASK STATUS</div>
          {"".join(rows)}
          <div class="task-total"><span>Total</span><span>{total}</span></div>
        </div>
        """
        with cols[idx%3]:
            st.markdown(card,unsafe_allow_html=True)

    st.download_button(
        "📥 DOWNLOAD FINAL REPORT",
        data=st.session_state["out"],
        file_name="Auto_Pivot_Live.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
