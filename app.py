import io, re, shutil, zipfile, tempfile, os, time
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
# Processing / Excel / Pivot logic above is intentionally unchanged.
st.markdown("""
<style>
:root { --navy:#0d2746; --blue:#1f5f95; --line:#d9e2ec; --bg:#f5f7fa; }
.stApp { background:var(--bg); }
.block-container { max-width:1240px; padding-top:1.35rem; padding-bottom:2.5rem; }
.hero {
    background:white; border:1px solid var(--line); border-radius:16px;
    padding:22px 28px 20px; margin-bottom:14px;
    box-shadow:0 4px 14px rgba(18,53,91,.06);
}
.hero h1 { color:var(--navy); font-size:34px; margin:0; font-weight:850; letter-spacing:-.3px; }
.hero p { color:#52667a; font-size:15px; margin:5px 0 0; }
.flow { text-align:center; color:var(--blue); font-weight:750; font-size:14px; margin-top:13px; }
.section {
    background:white; border:1px solid var(--line); border-radius:14px;
    padding:18px 20px; margin:12px 0;
    box-shadow:0 3px 10px rgba(18,53,91,.045);
}
.section-title { color:var(--navy); font-size:21px; font-weight:850; margin-bottom:4px; }
.section-sub { color:#617386; font-size:13px; margin-bottom:10px; }

/* Generate button / runner: hidden until Generate is actually clicked. */
.run-box { margin:10px 0 14px; }
.runner-track {
    position:relative; width:100%; height:56px; overflow:hidden;
    background:#eef4fa; border:1px solid #cdd9e6; border-radius:12px;
    box-shadow:inset 0 1px 2px rgba(18,53,91,.05);
}
.runner-line {
    position:absolute; left:0; right:0; bottom:9px; height:2px;
    background:#c9d7e6;
}
.runner-person {
    position:absolute; left:-44px; bottom:12px; font-size:28px; line-height:1;
    animation:runAcross 1.35s linear infinite;
    filter:drop-shadow(0 2px 1px rgba(0,0,0,.12));
}
@keyframes runAcross {
    0% { left:-44px; transform:scaleX(1); }
    49% { left:calc(100% - 2px); transform:scaleX(1); }
    50% { transform:scaleX(-1); }
    99% { left:-44px; transform:scaleX(-1); }
    100% { left:-44px; transform:scaleX(1); }
}
.runner-label { margin-top:6px; text-align:center; color:var(--navy); font-weight:800; font-size:14px; }

/* Compact, aligned report grid. Each desktop row stretches cards to the same height. */
.report-grid {
    display:grid; grid-template-columns:repeat(3,minmax(0,1fr));
    gap:12px; align-items:stretch; margin-top:10px;
}
.sheet-card {
    background:white; border:1px solid #cfdbe7; border-radius:13px;
    overflow:hidden; display:flex; flex-direction:column; min-width:0; height:100%;
    box-shadow:0 3px 10px rgba(18,53,91,.045);
}
.sheet-head {
    background:var(--navy); color:white; padding:11px 15px;
    font-size:18px; font-weight:850; letter-spacing:.1px;
}
.sheet-body { padding:0; flex:1; display:flex; flex-direction:column; }
.task-head {
    color:var(--navy); font-size:12px; font-weight:850; letter-spacing:.45px;
    padding:10px 15px 8px; border-bottom:1px solid #dfe7ef;
}
.status-row {
    display:flex; justify-content:space-between; align-items:center; gap:12px;
    padding:8px 15px; border-bottom:1px solid #edf1f5; min-height:34px;
    font-size:14px; line-height:1.2;
}
.status-name { color:#263746; font-weight:650; overflow-wrap:anywhere; }
.status-count { color:var(--navy); font-size:15px; font-weight:850; flex:0 0 auto; }
.total-row {
    display:flex; justify-content:space-between; align-items:center;
    margin-top:auto; padding:9px 15px; background:#eef4fa;
    border-top:1px solid #b8cce0; color:var(--navy); font-size:15px; font-weight:850;
}
.no-status { padding:13px 15px; color:#617386; font-size:13px; }
.download-wrap { margin-top:14px; }
.note { color:#617386; font-size:12px; }
@media (max-width: 980px) {
  .report-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
}
@media (max-width: 680px) {
  .block-container { padding-left:1rem; padding-right:1rem; }
  .hero h1 { font-size:27px; }
  .report-grid { grid-template-columns:1fr; }
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>📊 DAILY REPORT AUTOMATION</h1>
  <p>Automated Excel Reporting System</p>
  <div class="flow">Raw Excel&nbsp;&nbsp; → &nbsp;&nbsp;Process&nbsp;&nbsp; → &nbsp;&nbsp;Report</div>
</div>
""", unsafe_allow_html=True)

template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx")
if not template.exists():
    st.error("Template Excel မတွေ့ပါ။ app.py နဲ့ template Excel ကို GitHub repo တစ်ခုထဲမှာထားပါ။")
    st.stop()

st.markdown("""
<div class="section">
  <div class="section-title">📁 UPLOAD RAW EXCEL</div>
  <div class="section-sub">Upload today's Raw Excel file</div>
</div>
""", unsafe_allow_html=True)

uploaded=st.file_uploader("Choose Excel File",type=["xlsx"],label_visibility="collapsed")

generate = st.button("🚀 GENERATE REPORT", type="primary", use_container_width=True, disabled=uploaded is None)

if generate:
    # Show the long-running animation only after the button is clicked.
    runner_placeholder = st.empty()
    runner_placeholder.markdown("""
    <div class="run-box">
      <div class="runner-track">
        <div class="runner-line"></div>
        <div class="runner-person">🏃‍♂️</div>
      </div>
      <div class="runner-label">Generating report… please wait until the Excel file is ready</div>
    </div>
    """, unsafe_allow_html=True)
    # Give Streamlit a moment to send the animation to the browser before the heavy work starts.
    time.sleep(0.15)
    try:
        data,counts=build_workbook(uploaded.getvalue(),template)
        st.session_state["out"]=data
        st.session_state["counts"]=counts

        # UI-only status summary from the same processed Raw data.
        status_counts={}
        raw_bytes=uploaded.getvalue()
        xls_ui=pd.ExcelFile(io.BytesIO(raw_bytes))
        aliases_ui={"R1 OC":["R1 OC","R1OC"],"R6 OC":["R6 OC","R6OC"]}
        for logical in RAW_SHEETS:
            actual=logical
            if actual not in xls_ui.sheet_names and logical in aliases_ui:
                actual=next((a for a in aliases_ui[logical] if a in xls_ui.sheet_names),None)
            if actual in xls_ui.sheet_names:
                rdf=pd.read_excel(io.BytesIO(raw_bytes),sheet_name=actual)
                pdf=process(rdf,is_r2=(logical=="R2"))
                ts=next((c for c in pdf.columns if key(c)=="task status"),None)
                if ts:
                    vals=pdf[ts].map(norm)
                    vals=vals[vals!=""]
                    status_counts[logical]=vals.value_counts().to_dict()
                else:
                    status_counts[logical]={}
        st.session_state["status_counts"]=status_counts
        st.session_state["report_done"]=True
        runner_placeholder.empty()
    except Exception as e:
        runner_placeholder.empty()
        st.error(f"Error: {e}")

if st.session_state.get("report_done"):
    st.markdown("""
    <div class="section">
      <div class="section-title">📋 REPORT STATUS</div>
      <div class="section-sub">Report generated successfully from the uploaded Raw Excel</div>
    </div>
    """, unsafe_allow_html=True)
    st.success("✅ Report Generated")

    status_counts=st.session_state.get("status_counts",{})
    cards=[]
    for sheet in RAW_SHEETS:
        items=status_counts.get(sheet,{})
        parts=[f'<div class="sheet-card"><div class="sheet-head">📊 {sheet}</div><div class="sheet-body"><div class="task-head">TASK STATUS</div>']
        if items:
            total=0
            for status,count in items.items():
                total += int(count)
                safe_status=(str(status).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;"))
                parts.append(f'<div class="status-row"><span class="status-name">{safe_status}</span><span class="status-count">{int(count)}</span></div>')
            parts.append(f'<div class="total-row"><span>Total</span><span>{total}</span></div>')
        else:
            parts.append('<div class="no-status">No non-blank Task Status found.</div>')
        parts.append('</div></div>')
        cards.append("".join(parts))
    st.markdown('<div class="report-grid">' + "".join(cards) + '</div>', unsafe_allow_html=True)

    st.markdown('<div class="download-wrap">', unsafe_allow_html=True)
    st.download_button(
        "📥 DOWNLOAD FINAL REPORT",
        data=st.session_state["out"],
        file_name="Auto_Pivot_Live.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.markdown('</div>', unsafe_allow_html=True)
    st.caption("Excel ကိုဖွင့်ပြီး PivotTable မှာ Refresh လုပ်ပါ။ Pivot layout/formatting က template အတိုင်းပဲ ဖြစ်ပါတယ်။")
