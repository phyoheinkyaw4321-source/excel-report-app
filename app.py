import io, re, shutil, zipfile, tempfile, os
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pivot_refresh import refresh_workbook

st.set_page_config(page_title="Auto Pivot Report", page_icon="📊", layout="wide")

# Custom CSS for UI styling, running animation, cards, and password form
st.markdown("""
<style>
/* Modern styling for Password Card */
.login-card {
    max-width: 450px;
    margin: 60px auto;
    padding: 30px;
    border-radius: 16px;
    background: #ffffff;
    box-shadow: 0 10px 25px rgba(0,0,0,0.08);
    border: 1px solid #e0e0e0;
    text-align: center;
}

/* Running man track animation styling */
.runner-container {
    width: 100%;
    background-color: #f1f3f4;
    border-radius: 12px;
    height: 38px;
    position: relative;
    overflow: hidden;
    margin: 15px 0;
    border: 1px solid #e0e0e0;
}
.runner-track {
    width: 100%;
    height: 100%;
    position: relative;
}
.running-man {
    position: absolute;
    top: 3px;
    font-size: 22px;
    animation: runTrack 2.5s linear infinite;
}
@keyframes runTrack {
    0% { left: 0%; }
    100% { left: 95%; }
}

/* Card Styles matching screenshot */
.status-card {
    background: #ffffff;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
    margin-bottom: 20px;
    overflow: hidden;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}
.status-card-header {
    background: #0f2b48;
    color: #ffffff;
    padding: 10px 16px;
    font-weight: 700;
    font-size: 1.1rem;
    display: flex;
    align-items: center;
    gap: 8px;
}
.status-card-subheader {
    background: #f8fafc;
    color: #1e293b;
    padding: 8px 16px;
    font-size: 0.75rem;
    font-weight: 800;
    letter-spacing: 0.05em;
    border-bottom: 1px solid #e2e8f0;
    text-transform: uppercase;
}
.status-table {
    width: 100%;
    border-collapse: collapse;
}
.status-table tr {
    border-bottom: 1px solid #f1f5f9;
}
.status-table tr:last-child {
    border-bottom: none;
}
.status-table td {
    padding: 8px 16px;
    font-size: 0.88rem;
    color: #0f172a;
}
.status-name {
    font-weight: 600;
    text-align: left;
}
.status-count {
    font-weight: 800;
    text-align: right;
    color: #000000;
}
.status-total-row {
    background: #f1f5f9;
    font-weight: 700;
}
.status-total-row td {
    padding: 10px 16px;
    font-size: 0.95rem;
    color: #0f2b48;
}
</style>
""", unsafe_allow_html=True)

# Password protection
APP_PASSWORD = st.secrets.get("APP_PASSWORD", os.getenv("APP_PASSWORD", ""))

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if APP_PASSWORD:
    if not st.session_state.authenticated:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown("<div style='height: 40px;'></div>", unsafe_allow_html=True)
            st.markdown("### 🔐 Auto Pivot Report")
            st.caption("Password ထည့်ပြီး Login ဝင်ပါ")
            pw = st.text_input("Password", type="password", key="login_pw")
            if st.button("Login", type="primary", use_container_width=True):
                if pw == APP_PASSWORD:
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("❌ Password မှားနေပါတယ်")
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
    status_summaries={}

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

        # Extract Task Status breakdown for Card UI
        task_col=next((c for c in df.columns if key(c)=="task status"),None)
        if task_col:
            # Filter out blank / empty Task Status
            st_counts = df[task_col].map(norm).value_counts()
            st_counts = {k: v for k, v in st_counts.items() if k != ""}
            status_summaries[logical] = st_counts
        else:
            status_summaries[logical] = {}

    wb.save(out); wb.close()

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

    refresh_workbook(out)
    return out.read_bytes(), used, status_summaries

st.title("📊 Auto Pivot Report")
st.caption("Raw Excel → RUN → Daily/Historical report — previous-day rows are cleared first")
template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx")
if not template.exists():
    st.error("Template Excel မတွေ့ပါ။ app.py နဲ့ template Excel ကို GitHub repo တစ်ခုထဲမှာထားပါ။")
    st.stop()

uploaded=st.file_uploader("📁 Raw Excel (.xlsx)",type=["xlsx"])

# Show button only
if uploaded:
    if st.button("▶ GENERATE REPORT", type="primary", use_container_width=True):
        # Running animation placeholder
        runner_placeholder = st.empty()
        runner_placeholder.markdown("""
        <div class="runner-container">
            <div class="runner-track">
                <div class="running-man">🏃‍♂️</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.spinner("Report + Pivot ကို ပြင်ဆင်နေပါသည်..."):
            try:
                data, counts, status_summaries = build_workbook(uploaded.getvalue(), template)
                st.session_state["out"] = data
                st.session_state["counts"] = counts
                st.session_state["status_summaries"] = status_summaries
                st.success("✅ ပြီးပါပြီ — Report Generate လုပ်ခြင်း အောင်မြင်ပါသည်။")
            except Exception as e:
                st.error(f"Error: {e}")
            finally:
                # Clear running man animation when complete
                runner_placeholder.empty()

if "out" in st.session_state:
    st.download_button("⬇️ Download Live Pivot Excel", data=st.session_state["out"], file_name="Auto_Pivot_Live.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    st.info("Excel ကိုဖွင့်ပြီး PivotTable မှာ Refresh လုပ်ပါ။ Pivot layout/formatting က template အတိုင်းပဲ ဖြစ်ပါတယ်။")

    st.write("---")
    st.subheader("📋 Task Status Summary Cards")

    # Render 3-column responsive grid layout for the 6 sheets
    sheets = RAW_SHEETS
    summaries = st.session_state.get("status_summaries", {})

    cols = st.columns(3)
    for idx, sheet_name in enumerate(sheets):
        col = cols[idx % 3]
        sheet_summary = summaries.get(sheet_name, {})
        total_count = sum(sheet_summary.values())

        # Construct rows HTML
        rows_html = ""
        for status_name, count in sheet_summary.items():
            if status_name: # Ensure no blank status is shown
                rows_html += f"""
                <tr>
                    <td class="status-name">{status_name}</td>
                    <td class="status-count">{count:,}</td>
                </tr>
                """
        
        card_html = f"""
        <div class="status-card">
            <div class="status-card-header">
                📊 {sheet_name}
            </div>
            <div class="status-card-subheader">
                TASK STATUS
            </div>
            <table class="status-table">
                {rows_html}
                <tr class="status-total-row">
                    <td class="status-name">Total</td>
                    <td class="status-count">{total_count:,}</td>
                </tr>
            </table>
        </div>
        """
        with col:
            st.markdown(card_html, unsafe_allow_html=True)
