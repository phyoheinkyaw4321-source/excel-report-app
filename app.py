import io, re, shutil, zipfile, tempfile, os
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pivot_refresh import refresh_workbook

st.set_page_config(page_title="DAILY REPORT AUTOMATION", page_icon="📊", layout="wide")

# Secure access: password is ONLY read from Streamlit Secrets or environment variable.
def get_app_password():
    try:
        secret_password = st.secrets.get("APP_PASSWORD")
    except Exception:
        secret_password = None
    return secret_password or os.getenv("APP_PASSWORD")

APP_PASSWORD = get_app_password()

# Corporate UI only — no report/data-processing logic is changed below.
st.markdown("""
<style>
.stApp { background: #ffffff; }
.block-container { padding-top: 0.8rem; max-width: 1400px; }
.corp-header { background: #0b1220; color: #ffffff; padding: 24px 30px; border-radius: 0 0 12px 12px; margin-bottom: 24px; }
.corp-title { font-size: 30px; font-weight: 750; letter-spacing: .2px; margin: 0; }
.corp-subtitle { color: #cbd5e1; font-size: 15px; margin-top: 5px; }
.flow { margin-top: 18px; font-size: 14px; color: #e2e8f0; font-weight: 600; }
.secure-card { max-width: 520px; margin: 70px auto; padding: 32px; background: #ffffff; border: 1px solid #d9dee7; border-radius: 14px; box-shadow: 0 8px 30px rgba(15,23,42,.08); }
.section-card { background: #ffffff; border: 1px solid #d9dee7; border-radius: 12px; padding: 20px; margin: 14px 0; }
.runner-box { background: #ffffff; border: 1px solid #d9dee7; border-radius: 12px; padding: 18px 20px; margin: 14px 0; }
.runner-track { position: relative; height: 54px; border-bottom: 3px solid #1f2937; overflow: hidden; }
.runner { position: absolute; left: -10px; bottom: 5px; font-size: 34px; animation: run-across 5.5s linear infinite; }
@keyframes run-across { from { left: -10px; } to { left: calc(100% - 38px); } }
.runner-label { margin-top: 8px; font-weight: 650; color: #334155; }
.done-label { color: #166534; font-weight: 750; font-size: 18px; }
.status-grid { display: grid; grid-template-columns: repeat(3, minmax(0,1fr)); gap: 14px; }
.status-card { border: 1px solid #d9dee7; border-radius: 12px; background: #ffffff; padding: 16px; min-height: 150px; }
.status-name { font-weight: 750; font-size: 16px; margin-bottom: 10px; color: #0f172a; }
.status-row { display: flex; justify-content: space-between; gap: 16px; padding: 7px 0; border-bottom: 1px solid #eef1f5; font-size: 14px; color: #111827; font-weight: 700; }
.status-row:last-child { border-bottom: 0; }
.status-count { font-weight: 800; min-width: 55px; text-align: right; color: #111827; }
.status-total { margin-top: 10px; padding-top: 9px; border-top: 1px solid #d9dee7; display:flex; justify-content:space-between; font-weight: 750; }
@media (max-width: 900px) { .status-grid { grid-template-columns: 1fr; } }
</style>
""", unsafe_allow_html=True)

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not APP_PASSWORD:
    st.markdown("""<div class="secure-card"><div style="font-size:26px;font-weight:750">🔐 Secure Access</div><div style="margin-top:8px;color:#64748b">APP_PASSWORD is not configured. Add it to Streamlit Secrets or the environment variable before using the report.</div></div>""", unsafe_allow_html=True)
    st.stop()

if not st.session_state.authenticated:
    st.markdown("""<div class="secure-card"><div style="font-size:28px;font-weight:750">🔐 Secure Access</div><div style="margin-top:6px;color:#64748b">DAILY REPORT AUTOMATION</div><div style="margin-top:20px"></div></div>""", unsafe_allow_html=True)
    pw = st.text_input("Password", type="password", label_visibility="visible")
    if st.button("Login", type="primary", use_container_width=True):
        import hmac
        if hmac.compare_digest(pw, APP_PASSWORD):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Password မှားနေပါတယ်")
    st.stop()

st.markdown("""<div class="corp-header"><div class="corp-title">📊 DAILY REPORT AUTOMATION</div><div class="corp-subtitle">Automated Excel Reporting System</div><div class="flow">Raw Excel &nbsp;→&nbsp; Process &nbsp;→&nbsp; Report</div></div>""", unsafe_allow_html=True)

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

        # Keep Task Status as TEXT all the way into the output workbook.
        # This prevents Excel/PivotTable from interpreting status labels as dates.
        task_status_cols=[c for c in df.columns if key(c)=="task status"]
        for _ts_col in task_status_cols:
            df[_ts_col]=df[_ts_col].map(norm)

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
                # Tell Excel to rebuild the Pivot cache from the current TEXT source data.
                if "refreshOnLoad=" not in xml:
                    xml=xml.replace("<pivotCacheDefinition ","<pivotCacheDefinition refreshOnLoad=\"1\" ",1)
                if "enableRefresh=" not in xml:
                    xml=xml.replace("<pivotCacheDefinition ","<pivotCacheDefinition enableRefresh=\"1\" ",1)
                data=xml.encode()
            # Pivot count fields must be numeric/General, never a date format.
            # Use built-in General (numFmtId=0) for Count - Task No.
            # openpyxl can remap custom numFmt IDs during save; using built-in Excel format 3 (#,##0) avoids
            # Excel interpreting counts such as 13 as 13-Jan-00, including after Pivot filters/refresh.
            if re.fullmatch(r"xl/pivotTables/pivotTable[1-6]\.xml", item.filename):
                xml=data.decode("utf-8")
                xml=re.sub(r'(<dataField\b[^>]*\bnumFmtId=)"\d+"', r'\1"3"', xml)
                data=xml.encode("utf-8")
            zout.writestr(item,data)
    out.unlink(); tmp.replace(out)

    # IMPORTANT: Do not pass the workbook through LibreOffice here.
    # LibreOffice can rewrite Excel PivotTable structures and make the
    # PivotTable unusable in Microsoft Excel. The original Excel PivotTables
    # from the template are preserved, their cache source ranges are patched
    # above, and refreshOnLoad is enabled so Excel can refresh them safely.
    return out.read_bytes(), used

template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx")
if not template.exists():
    template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1)(1).xlsx")
if not template.exists():
    st.error("Template Excel မတွေ့ပါ။ app.py နဲ့ template Excel ကို GitHub repo တစ်ခုထဲမှာထားပါ။")
    st.stop()

# Task Status display is read-only from the uploaded Raw workbook.
# It does not modify build_workbook/process/pivot logic.
def task_status_summary(raw_bytes):
    raw_stream = io.BytesIO(raw_bytes)
    xls = pd.ExcelFile(raw_stream)
    result = {}
    for sheet in RAW_SHEETS:
        actual = sheet
        if actual not in xls.sheet_names and sheet in {"R1 OC", "R6 OC"}:
            actual = next((a for a in [sheet, sheet.replace(" ", "")] if a in xls.sheet_names), None)
        if not actual or actual not in xls.sheet_names:
            result[sheet] = {"counts": {}, "total": 0}
            continue
        raw_stream.seek(0)
        df = pd.read_excel(raw_stream, sheet_name=actual, dtype=object)
        status_col = next((c for c in df.columns if key(c) == "task status"), None)
        if status_col is None:
            result[sheet] = {"counts": {}, "total": 0}
            continue
        status = df[status_col].map(norm)
        status = status[status != ""]
        counts = status.value_counts().to_dict()
        result[sheet] = {"counts": counts, "total": int(status.size)}
    return result

def render_task_status(summary):
    cards = []
    for sheet in RAW_SHEETS:
        info = summary.get(sheet, {"counts": {}, "total": 0})
        rows = "".join(f'<div class="status-row"><span style="color:#000;font-weight:700">{str(status)}</span><span class="status-count" style="color:#000;font-weight:700">{int(count)}</span></div>' for status, count in info["counts"].items())
        if not rows:
            rows = '<div style="color:#94a3b8;font-size:14px;padding:8px 0">No nonblank Task Status</div>'
        cards.append(f'<div class="status-card"><div class="status-name">{sheet}</div>{rows}<div class="status-total"><span>Total</span><span>{info["total"]}</span></div></div>')
    st.markdown('<div class="status-grid">' + "".join(cards) + '</div>', unsafe_allow_html=True)

st.markdown('<div class="section-card"><h3 style="margin:0 0 10px 0">📁 Generate Report</h3><div style="color:#64748b">Upload one Raw Excel file, then generate the report.</div></div>', unsafe_allow_html=True)
uploaded=st.file_uploader("Raw Excel (.xlsx)",type=["xlsx"])

if uploaded and st.button("▶ GENERATE REPORT",type="primary",use_container_width=True):
    runner = st.empty()
    runner.markdown("""<div class="runner-box"><div class="runner-track"><span class="runner">🏃</span></div><div class="runner-label">Processing Raw Excel → Report + Pivot...</div></div>""", unsafe_allow_html=True)
    try:
        raw_bytes = uploaded.getvalue()
        data,counts=build_workbook(raw_bytes,template)
        status_summary = task_status_summary(raw_bytes)
        st.session_state["out"]=data
        st.session_state["counts"]=counts
        st.session_state["task_status"]=status_summary
        runner.markdown("""<div class="runner-box"><div class="done-label">✅ Report Generated</div></div>""", unsafe_allow_html=True)
        st.success("Report ပြီးပါပြီ။ PivotTable source data ကို TEXT အဖြစ်ထိန်းထားပြီး Excel ဖွင့်တဲ့အချိန် Refresh ဖြစ်အောင်ပြင်ထားပါတယ်။")

    except Exception as e:
        runner.empty()
        st.error(f"Error: {e}")

if "out" in st.session_state:
    st.markdown('<div class="section-card"><h3 style="margin:0 0 14px 0">📋 Task Status</h3><div style="color:#64748b">Processed report status · nonblank values only</div></div>', unsafe_allow_html=True)
    render_task_status(st.session_state.get("task_status", {}))
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.download_button("⬇️ Download Live Pivot Excel",data=st.session_state["out"],file_name="Auto_Pivot_Live.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
