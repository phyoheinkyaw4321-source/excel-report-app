
import io, re, shutil, zipfile, tempfile
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Auto Pivot Report", page_icon="📊", layout="wide")
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
    if task:
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
            tmp=tmp[~df[repay].map(norm).str.lower().str.contains(r"engr\s*leave|route\s*cancel",regex=True,na=False)]
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
                t = key(x)
                # Blank Township must stay blank.
                if not t:
                    vals.append("")
                    continue
                # Match exact township or township text before a ward/parenthesis suffix,
                # e.g. "Sanchaung (North) Ward (SCHG)" -> Sanchaung.
                base = re.split(r"\s*\(|\s+ward\b", t, maxsplit=1)[0].strip()
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

    for s in RAW_SHEETS:
        if s not in xls.sheet_names or s not in wb.sheetnames:
            continue
        raw_stream.seek(0)
        df=process(pd.read_excel(raw_stream,sheet_name=s),is_r2=(s=="R2"))
        ws=wb[s]
        headers=[ws.cell(1,c).value for c in range(1,ws.max_column+1)]
        hmap={key(v):i+1 for i,v in enumerate(headers) if v is not None}
        # Clear values only; keep the template's formatting, widths, panes and pivot layout.
        for row in ws.iter_rows(min_row=2,max_row=ws.max_row,max_col=ws.max_column):
            for cell in row: cell.value=None
        for ri,row in enumerate(df.to_dict("records"),start=2):
            for col,val in row.items():
                ci=hmap.get(key(col))
                if ci:
                    cell = ws.cell(ri,ci)
                    cell.value = val
                    if key(col) == "task status":
                        cell.number_format = "General"
        # Keep Task Status header/data column as General; never convert it to a date.
        for hname, ci in hmap.items():
            if hname == "task status":
                for rr in range(2, max(2, len(df)+1)):
                    ws.cell(rr, ci).number_format = "General"
        ws.auto_filter.ref=f"A1:{get_column_letter(ws.max_column)}{max(2,len(df)+1)}"
        used[s]=len(df)

    wb.save(out)
    wb.close()

    # Update each existing PivotCache source range and ask Excel to refresh the cache
    # when the user opens the downloaded workbook. This preserves the PivotTable
    # definitions/layout from the template.
    tmp=out.with_suffix(".patched.xlsx")
    with zipfile.ZipFile(out,"r") as zin, zipfile.ZipFile(tmp,"w",zipfile.ZIP_DEFLATED) as zout:
        cache_map={1:"R1",2:"R1S",3:"R2",4:"R6",5:"R1 OC",6:"R6 OC"}
        for item in zin.infolist():
            data=zin.read(item.filename)
            m=re.fullmatch(r"xl/pivotCache/pivotCacheDefinition([1-6])\.xml",item.filename)
            if m:
                i=int(m.group(1)); s=cache_map[i]
                # Reopen workbook just for max_column/max_row.
                wbx=load_workbook(out,read_only=True,data_only=False)
                wsx=wbx[s]
                ref=f"A1:{get_column_letter(wsx.max_column)}{wsx.max_row}"
                wbx.close()
                xml=data.decode("utf-8")
                xml=re.sub(r'<worksheetSource ref="[^"]+" sheet="([^"]+)"',
                           lambda mm:f'<worksheetSource ref="{ref}" sheet="{mm.group(1)}"',xml,count=1)
                if 'refreshOnLoad=' not in xml:
                    xml=xml.replace('<pivotCacheDefinition ','<pivotCacheDefinition refreshOnLoad="1" ',1)
                data=xml.encode()
            # IMPORTANT: Pivot value fields such as "Count - Task No" must stay
            # numeric.  If the template's DataField carries a date numFmtId,
            # Excel can render counts such as 11 as 11-Jan-00 after refresh.
            # Force every PivotTable data field back to General/Number format.
            if re.fullmatch(r"xl/pivotTables/pivotTable[0-9]+\\.xml", item.filename):
                xml = data.decode("utf-8")
                xml = re.sub(r'(dataField\\b[^>]*?)\\snumFmtId="[^"]+"',
                             r'\\1 numFmtId="0"', xml)
                # If a dataField has no numFmtId, add General format.
                xml = re.sub(r'<dataField(?![^>]*\\bnumFmtId=)([^>]*)/',
                             r'<dataField\\1 numFmtId="0"/', xml)
                data = xml.encode("utf-8")
            zout.writestr(item,data)
    out.unlink()
    tmp.replace(out)
    return out.read_bytes(), used


# -------------------- Corporate UI --------------------
st.markdown("""
<style>
.stApp { background: #ffffff; }
.block-container { max-width: 1180px; padding-top: 2rem; padding-bottom: 3rem; }

.hero {
    background: #0b1f3a;
    color: white;
    padding: 28px 32px 24px 32px;
    border-radius: 12px;
    margin-bottom: 22px;
}
.hero-title { font-size: 34px; font-weight: 800; letter-spacing: .2px; margin: 0; }
.hero-sub { font-size: 18px; font-weight: 500; margin-top: 6px; opacity: .95; }
.hero-flow { text-align:center; font-size: 20px; font-weight: 700; margin-top: 22px; }

.section-title {
    color: #0b1f3a;
    font-size: 25px;
    font-weight: 800;
    border-bottom: 3px solid #0b1f3a;
    padding-bottom: 7px;
    margin-top: 24px;
    margin-bottom: 14px;
}
.sheet-title {
    color: #0b1f3a;
    font-size: 25px;
    font-weight: 800;
    margin-top: 25px;
    margin-bottom: 7px;
}
.sub-title {
    color: #0b1f3a;
    font-size: 19px;
    font-weight: 800;
    margin-bottom: 8px;
}

.status-card {
    border: 1px solid #c9d3df;
    border-left: 6px solid #0b1f3a;
    border-radius: 8px;
    padding: 14px 18px;
    margin-bottom: 16px;
    background: #f8fafc;
    font-size: 17px;
}
.status-ok { color: #0b1f3a; font-weight: 800; font-size: 20px; }
.status-date { font-size: 17px; margin-top: 5px; }

.upload-box {
    background: #f8fafc;
    border: 1px solid #c9d3df;
    border-radius: 10px;
    padding: 18px;
}

div[data-testid="stFileUploader"] label {
    font-size: 18px !important;
    font-weight: 700 !important;
}
div[data-testid="stFileUploader"] section {
    min-height: 86px;
}
.stButton button, .stDownloadButton button {
    min-height: 52px;
    font-size: 18px !important;
    font-weight: 800 !important;
    border-radius: 8px;
}
.stDownloadButton button {
    background: #0b1f3a !important;
    color: white !important;
    border: 1px solid #0b1f3a !important;
}

.report-table {
    width: 100%;
    border-collapse: collapse;
    margin-bottom: 20px;
    font-size: 17px;
}
.report-table th {
    background: #0b1f3a;
    color: white;
    font-weight: 800;
    padding: 13px 14px;
    border: 1px solid #9aa8b7;
    text-align: center;
    white-space: nowrap;
}
.report-table td {
    background: white;
    color: #111827;
    font-weight: 650;
    padding: 14px 14px;
    border: 1px solid #b8c2cc;
    text-align: center;
    min-height: 48px;
}
.report-table td:first-child {
    text-align: left;
    font-weight: 800;
}
.report-table .total {
    background: #eef3f8;
    font-weight: 850;
}
.note {
    color: #4b5563;
    font-size: 15px;
    margin-top: 6px;
}
</style>
""", unsafe_allow_html=True)

# Password / login
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.markdown("""
    <div style="max-width:560px;margin:70px auto 0 auto;text-align:center;">
      <div style="color:#0b1f3a;font-size:34px;font-weight:800;">📊 DAILY REPORT AUTOMATION</div>
      <div style="color:#4b5563;font-size:18px;margin-top:8px;">Secure access</div>
    </div>
    """, unsafe_allow_html=True)

    password = st.secrets.get("APP_PASSWORD", None)
    if password is None:
        import os
        password = os.environ.get("APP_PASSWORD", "")

    with st.container():
        st.markdown('<div style="max-width:560px;margin:24px auto;">', unsafe_allow_html=True)
        entered = st.text_input("🔐 Password", type="password", placeholder="Enter password")
        if st.button("🔓 LOGIN", type="primary", use_container_width=True):
            if password and entered == password:
                st.session_state["authenticated"] = True
                st.rerun()
            elif not password:
                st.error("APP_PASSWORD is not configured in Streamlit Secrets.")
            else:
                st.error("Incorrect password.")
        st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

st.markdown("""
<div class="hero">
  <div class="hero-title">📊 DAILY REPORT AUTOMATION</div>
  <div class="hero-sub">Automated Excel Reporting System</div>
  <div class="hero-flow">Raw Excel&nbsp;&nbsp; → &nbsp;&nbsp;Process&nbsp;&nbsp; → &nbsp;&nbsp;Report</div>
</div>
""", unsafe_allow_html=True)

template = Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx")
if not template.exists():
    st.error("Template Excel မတွေ့ပါ။ app.py နဲ့ template Excel ကို GitHub repo တစ်ခုထဲမှာထားပါ။")
    st.stop()

st.markdown('<div class="section-title">📁 UPLOAD RAW EXCEL</div>', unsafe_allow_html=True)
st.markdown('<div class="upload-box">Upload today\'s Raw Excel file</div>', unsafe_allow_html=True)
uploaded = st.file_uploader("Choose Excel File", type=["xlsx"], label_visibility="visible")

if uploaded and st.button("🚀 GENERATE REPORT", type="primary", use_container_width=True):
    with st.spinner("Report + Pivot ကိုပြင်နေပါတယ်..."):
        try:
            data, counts = build_workbook(uploaded.getvalue(), template)

            # Build UI-only Task Status summaries from the same processed Raw data.
            raw_stream = io.BytesIO(uploaded.getvalue())
            xls_ui = pd.ExcelFile(raw_stream)
            task_status_counts = {}
            report_date = ""
            for s in RAW_SHEETS:
                if s not in xls_ui.sheet_names:
                    continue
                raw_stream.seek(0)
                df_ui = process(pd.read_excel(raw_stream, sheet_name=s), is_r2=(s == "R2"))
                task_col = next((c for c in df_ui.columns if key(c) == "task status"), None)
                if task_col:
                    vals = df_ui[task_col].map(norm)
                    vc = vals[vals != ""].value_counts(dropna=False)
                    task_status_counts[s] = {str(k): int(v) for k, v in vc.items()}
                else:
                    task_status_counts[s] = {}

                if not report_date:
                    date_col = next((c for c in df_ui.columns if key(c) in {"schedule date", "date"}), None)
                    if date_col:
                        dates = pd.to_datetime(df_ui[date_col], errors="coerce").dropna()
                        if len(dates):
                            report_date = dates.min().strftime("%d-%b-%Y")

            st.session_state["out"] = data
            st.session_state["counts"] = counts
            st.session_state["task_status_counts"] = task_status_counts
            st.session_state["report_date"] = report_date
            st.success("✅ Report Generated")
        except Exception as e:
            st.error(f"Error: {e}")

if "out" in st.session_state:
    st.markdown('<div class="section-title">📋 REPORT STATUS</div>', unsafe_allow_html=True)
    rd = st.session_state.get("report_date", "")
    rd_text = rd if rd else "Current Raw Excel"
    st.markdown(
        f'<div class="status-card"><div class="status-ok">✅ Report Generated</div>'
        f'<div class="status-date">Reporting Date: <b>{rd_text}</b></div></div>',
        unsafe_allow_html=True
    )

    sheet_order = ["R1", "R1S", "R2", "R6", "R1 OC", "R6 OC"]
    all_statuses = set()
    for s in sheet_order:
        all_statuses.update(st.session_state.get("task_status_counts", {}).get(s, {}).keys())

    # Keep the familiar statuses first, then show any new non-blank statuses automatically.
    preferred = ["Resolved", "Done", "Resolved (Auto)", "Pending", "Cancelled"]
    statuses = [x for x in preferred if x in all_statuses]
    statuses += sorted(x for x in all_statuses if x not in statuses and x.strip())

    for s in sheet_order:
        counts_s = st.session_state.get("task_status_counts", {}).get(s, {})
        if not counts_s:
            continue

        st.markdown(f'<div class="sheet-title">📊 {s}</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-title">TASK STATUS</div>', unsafe_allow_html=True)

        cells = []
        for status in statuses:
            cells.append(f"<th>{status}</th>")
        total = sum(counts_s.values())
        values = [f"<td>{counts_s.get(status, 0):,}</td>" for status in statuses]
        cells_html = "".join(cells)
        values_html = "".join(values)

        st.markdown(
            f"""
            <table class="report-table">
              <thead>
                <tr><th>Sheet</th>{cells_html}<th>Total</th></tr>
              </thead>
              <tbody>
                <tr><td>{s}</td>{values_html}<td class="total">{total:,}</td></tr>
              </tbody>
            </table>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        "📥 DOWNLOAD FINAL REPORT",
        data=st.session_state["out"],
        file_name="Auto_Pivot_Live.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.markdown(
        '<div class="note">Excel file retains the existing Live Pivot layout. '
        'Use Excel → Data → Refresh All to refresh the PivotTables.</div>',
        unsafe_allow_html=True
    )
