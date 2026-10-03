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
        st.title("🔐 Auto Pivot Report")
        st.caption("Password ထည့်ပြီး Report ကိုအသုံးပြုပါ")
        pw = st.text_input("Password", type="password")
        if st.button("Login", type="primary", use_container_width=True):
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


# ---------------- UI ONLY ----------------
# Processing / business / Pivot logic above this line is intentionally unchanged.

st.markdown("""
<style>
:root{
  --navy:#0b2342;
  --navy2:#12345d;
  --line:#d9e1ea;
  --soft:#f5f8fb;
  --text:#10243e;
}
.block-container{max-width:1180px;padding-top:1.2rem;padding-bottom:2rem;}
.main-title{color:var(--navy);font-size:2.15rem;font-weight:900;letter-spacing:-.5px;margin-bottom:0;}
.sub-title{color:#607086;font-size:1rem;margin-top:.15rem;margin-bottom:1.15rem;}
.workflow{
  display:flex;align-items:center;justify-content:center;gap:14px;
  color:var(--navy);font-weight:800;font-size:.95rem;
  background:#f5f8fb;border:1px solid var(--line);border-radius:12px;
  padding:10px 16px;margin:0 0 18px;
}
.section{
  color:var(--navy);font-size:1.05rem;font-weight:900;
  letter-spacing:.2px;margin:18px 0 8px;
}
.login-wrap{
  max-width:430px;margin:7vh auto 0;background:#fff;
  border:1px solid #dce3eb;border-radius:18px;padding:30px 30px 26px;
  box-shadow:0 12px 35px rgba(11,35,66,.10);
}
.login-title{color:var(--navy);font-size:1.55rem;font-weight:900;text-align:center;margin-bottom:3px;}
.login-sub{text-align:center;color:#718096;font-size:.92rem;margin-bottom:18px;}
div[data-testid="stTextInput"] label{font-weight:800;color:var(--navy);}
div[data-testid="stFileUploader"]{
  border:1px dashed #aebdce;border-radius:12px;padding:5px;background:#fbfcfe;
}
.run-wrap{
  margin:12px 0 18px;border:1px solid #d7e0ea;border-radius:14px;
  background:#fff;padding:14px 16px 16px;box-shadow:0 5px 18px rgba(11,35,66,.06);
}
.runner-track{
  position:relative;height:42px;overflow:hidden;border-radius:22px;
  background:linear-gradient(90deg,#eef3f8,#f8fafc);
  border:1px solid #d8e1eb;
}
.runner-road{
  position:absolute;left:14px;right:14px;top:20px;height:2px;
  background:#b9c7d7;
}
.runner{
  position:absolute;top:3px;left:-55px;font-size:28px;
  animation:runSlow 6.5s linear infinite;
  will-change:left;
}
@keyframes runSlow{
  0%{left:-55px;transform:scaleX(1);}
  48%{transform:scaleX(1);}
  50%{transform:scaleX(-1);}
  98%{transform:scaleX(-1);}
  100%{left:calc(100% + 55px);transform:scaleX(-1);}
}
.runner-text{text-align:center;color:var(--navy);font-weight:800;font-size:.9rem;margin-top:8px;}
.status-box{
  border:1px solid #d7e0ea;border-radius:14px;background:#fff;
  padding:14px 16px;margin-bottom:16px;box-shadow:0 5px 18px rgba(11,35,66,.05);
}
.status-ok{color:#16704a;font-weight:900;}
.report-date{color:#5e7085;font-size:.92rem;margin-top:2px;}
.sheet-grid{
  display:grid;grid-template-columns:repeat(3,minmax(0,1fr));
  gap:14px;align-items:start;
}
.sheet-card{
  border:1px solid #d3dde8;border-radius:12px;background:#fff;
  overflow:hidden;box-shadow:0 4px 14px rgba(11,35,66,.06);
  min-width:0;
}
.sheet-head{
  background:var(--navy);color:#fff;padding:11px 13px;
  font-size:1.02rem;font-weight:900;display:flex;align-items:center;gap:7px;
}
.sheet-label{
  padding:9px 13px 7px;color:var(--navy);font-size:.78rem;
  font-weight:900;letter-spacing:.7px;border-bottom:1px solid #e0e6ed;
}
.ts-row{
  display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;
  padding:8px 13px;border-bottom:1px solid #edf1f5;
  min-height:36px;align-items:center;
}
.ts-name{color:#162b45;font-size:.88rem;font-weight:700;line-height:1.2;}
.ts-count{color:var(--navy);font-size:.91rem;font-weight:900;text-align:right;}
.ts-total{
  display:grid;grid-template-columns:1fr auto;gap:8px;
  padding:10px 13px;background:#edf3f8;color:var(--navy);
  font-size:.92rem;font-weight:900;border-top:1px solid #ccd8e4;
}
.download-wrap{margin-top:20px;}
@media(max-width:900px){
  .sheet-grid{grid-template-columns:repeat(2,minmax(0,1fr));}
}
@media(max-width:620px){
  .sheet-grid{grid-template-columns:1fr;}
  .main-title{font-size:1.7rem;}
  .workflow{font-size:.82rem;gap:7px;}
  .login-wrap{margin-top:3vh;padding:24px 20px;}
}
</style>
""", unsafe_allow_html=True)

# Clean password page — authentication logic itself remains the same.
if APP_PASSWORD:
    if not st.session_state.authenticated:
        st.markdown("""
        <div class="login-wrap">
          <div class="login-title">🔐 Auto Pivot Report</div>
          <div class="login-sub">Secure access · Daily Reporting System</div>
        </div>
        """, unsafe_allow_html=True)
        pw = st.text_input("Password", type="password", label_visibility="visible", key="login_password")
        if st.button("🔓  LOGIN", type="primary", use_container_width=True):
            if pw == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Password မှားနေပါတယ်")
        st.stop()
else:
    st.warning("APP_PASSWORD ကို Streamlit Secrets မှာ ထည့်ထားပါ။ Password protection ကို မဖယ်ထားပါ။")
    st.stop()

st.markdown('<div class="main-title">📊 DAILY REPORT AUTOMATION</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Automated Excel Reporting System</div>', unsafe_allow_html=True)
st.markdown('<div class="workflow">Raw Excel&nbsp;&nbsp;→&nbsp;&nbsp; Process&nbsp;&nbsp;→&nbsp;&nbsp; Report</div>', unsafe_allow_html=True)

template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx")
if not template.exists():
    st.error("Template Excel မတွေ့ပါ။ app.py နဲ့ template Excel ကို GitHub repo တစ်ခုထဲမှာထားပါ။")
    st.stop()

st.markdown('<div class="section">📁 UPLOAD RAW EXCEL</div>', unsafe_allow_html=True)
st.caption("Upload today's Raw Excel file")
uploaded=st.file_uploader("Choose Excel File",type=["xlsx"],label_visibility="collapsed")

def _task_status_display(raw_bytes):
    """UI display only; uses the same existing process() rules without changing the workbook."""
    result={}
    xls=pd.ExcelFile(io.BytesIO(raw_bytes))
    aliases={"R1 OC":["R1 OC","R1OC"],"R6 OC":["R6 OC","R6OC"]}
    for logical in RAW_SHEETS:
        actual=logical
        if actual not in xls.sheet_names and logical in aliases:
            actual=next((a for a in aliases[logical] if a in xls.sheet_names),None)
        if not actual:
            result[logical]={}
            continue
        df=process(pd.read_excel(io.BytesIO(raw_bytes),sheet_name=actual),is_r2=(logical=="R2"))
        status_col=next((c for c in df.columns if key(c)=="task status"),None)
        if status_col is None:
            result[logical]={}
            continue
        vals=df[status_col].map(norm)
        vals=vals[vals!=""]
        result[logical]=vals.value_counts().to_dict()
    return result

if uploaded:
    if st.button("🚀  GENERATE REPORT",type="primary",use_container_width=True):
        runner = st.empty()
        runner.markdown("""
        <div class="run-wrap">
          <div class="runner-track">
            <div class="runner-road"></div>
            <div class="runner">🏃‍♂️</div>
          </div>
          <div class="runner-text">Generating report… please wait until the file is ready.</div>
        </div>
        """, unsafe_allow_html=True)
        try:
            # Keep the existing build_workbook() and Pivot logic untouched.
            data,counts=build_workbook(uploaded.getvalue(),template)
            task_counts=_task_status_display(uploaded.getvalue())
            st.session_state["out"]=data
            st.session_state["counts"]=counts
            st.session_state["task_counts"]=task_counts
            st.session_state["report_file_name"]=uploaded.name
            runner.empty()
            st.success("✅ Report Generated")
        except Exception as e:
            runner.empty()
            st.error(f"Error: {e}")

if "out" in st.session_state:
    st.markdown('<div class="section">📋 REPORT STATUS</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="status-box">
      <div class="status-ok">✅ Report Generated</div>
    </div>
    """, unsafe_allow_html=True)

    task_counts=st.session_state.get("task_counts",{})
    st.markdown('<div class="sheet-grid">', unsafe_allow_html=True)
    for sheet in RAW_SHEETS:
        items=task_counts.get(sheet,{})
        total=sum(items.values())
        rows=[]
        for status,count in items.items():
            rows.append(
                f'<div class="ts-row"><div class="ts-name">{status}</div>'
                f'<div class="ts-count">{int(count)}</div></div>'
            )
        if not rows:
            rows.append('<div class="ts-row"><div class="ts-name">No non-blank Task Status</div><div class="ts-count">0</div></div>')
        card=f"""
        <div class="sheet-card">
          <div class="sheet-head">📊 {sheet}</div>
          <div class="sheet-label">TASK STATUS</div>
          {''.join(rows)}
          <div class="ts-total"><div>Total</div><div>{int(total)}</div></div>
        </div>
        """
        st.markdown(card, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="download-wrap"></div>', unsafe_allow_html=True)
    st.download_button(
        "📥  DOWNLOAD FINAL REPORT",
        data=st.session_state["out"],
        file_name="Auto_Pivot_Live.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.caption("Excel ကိုဖွင့်ပြီး PivotTable မှာ Refresh လုပ်ပါ။ Pivot layout/formatting က template အတိုင်းပဲ ဖြစ်ပါတယ်။")

