import copy

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
    """Build a report from the uploaded workbook while preserving the native PivotTable parts.

    Rules locked for this app:
    - Raw sheets are copied by header name, not by column position.
    - Original raw cell values/styles are preserved; no pandas round-trip is used.
    - All six raw sheets are cleared first; missing sheets remain completely blank.
    - Blank rows in the uploaded sheet are retained at their original row positions.
    - Only TKT / POI, Classification, RC and (R2) RE are calculated.
    - Task Status and Task No are General/Number, never date-formatted.
    - Native PivotTable definitions are taken from the template unchanged.
    - Pivot cache 5 = R6 OC and cache 6 = R1 OC (matching the template).
    """
    raw_path = Path(tempfile.mktemp(suffix='.xlsx'))
    prepared = Path(tempfile.mktemp(suffix='.xlsx'))
    out = Path(tempfile.mktemp(suffix='.xlsx'))
    raw_path.write_bytes(raw_bytes)
    shutil.copy2(template_path, prepared)

    raw_wb = load_workbook(raw_path, data_only=False)
    wb = load_workbook(prepared, data_only=False)
    used = {}

    def clear_sheet(ws):
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
            for cell in row:
                cell.value = None

    def header_map(ws):
        return {key(ws.cell(1,c).value): c for c in range(1, ws.max_column+1)
                if ws.cell(1,c).value is not None}

    def copy_cell(src, dst):
        dst.value = src.value
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)
        if src.hyperlink:
            dst._hyperlink = copy(src.hyperlink)

    for s in RAW_SHEETS:
        if s not in wb.sheetnames:
            continue
        tws = wb[s]
        clear_sheet(tws)
        if s not in raw_wb.sheetnames:
            used[s] = 0
            continue

        rws = raw_wb[s]
        th = header_map(tws)
        rh = header_map(rws)

        # Copy every source row at its original row number. This intentionally
        # keeps blank Schedule Date rows and blank separator rows.
        for rr in range(2, rws.max_row + 1):
            for kh, rc in rh.items():
                tc = th.get(kh)
                if tc is not None:
                    copy_cell(rws.cell(rr, rc), tws.cell(rr, tc))

        # Derived fields use the already-copied raw values.
        task_c = th.get('task no')
        service_c = th.get('service area')
        root_c = th.get('root cause')
        tkt_c = th.get('tkt / poi')
        cls_c = th.get('classification')
        car_c = th.get('car id')
        eng_c = th.get('lan engineer name')
        repay_c = th.get('repay status')
        rc_c = th.get('rc')
        town_c = th.get('township')
        re_c = th.get('re')

        # RC distinct engineer count by Car ID, excluding Engr Leave / route cancel.
        car_eng = {}
        if car_c and eng_c:
            for rr in range(2, rws.max_row + 1):
                car = norm(tws.cell(rr, car_c).value)
                eng = norm(tws.cell(rr, eng_c).value)
                repay = norm(tws.cell(rr, repay_c).value) if repay_c else ''
                if not car or not eng:
                    continue
                if re.search(r'\bengr\s*leave\b|\broute\s*cancel\b', repay, flags=re.I):
                    continue
                car_eng.setdefault(car.lower(), set()).add(eng.lower())

        for rr in range(2, rws.max_row + 1):
            task = tws.cell(rr, task_c).value if task_c else None
            tt = task_type(task)
            if tkt_c:
                tws.cell(rr, tkt_c).value = tt
            if cls_c:
                tws.cell(rr, cls_c).value = classification(
                    tt,
                    tws.cell(rr, service_c).value if service_c else '',
                    tws.cell(rr, root_c).value if root_c else ''
                )
            if rc_c:
                car = norm(tws.cell(rr, car_c).value) if car_c else ''
                n = len(car_eng.get(car.lower(), set())) if car else 0
                tws.cell(rr, rc_c).value = (
                    '' if not car else
                    'Car Share Plus' if n >= 3 else
                    'Car Share' if n == 2 else
                    'One Car' if n == 1 else ''
                )
            if re_c:
                town = norm(tws.cell(rr, town_c).value) if town_c else ''
                if not town:
                    re_val = ''
                else:
                    base = re.split(r'\s*\(|\s+ward\b', town.lower(), maxsplit=1)[0].strip()
                    re_val = 'R2S' if base in R2S_TOWNSHIPS else 'R2'
                tws.cell(rr, re_c).value = re_val

            if th.get('task status'):
                tws.cell(rr, th['task status']).number_format = 'General'
            if task_c:
                tws.cell(rr, task_c).number_format = 'General'

        used[s] = rws.max_row - 1

    # Preserve native PivotTable definitions by using this workbook only to
    # generate the six worksheet XML files; the final ZIP starts from the
    # untouched template and replaces only raw sheet XML + cache metadata.
    wb.save(prepared)
    wb.close()
    raw_wb.close()

    raw_sheet_xml = {}
    with zipfile.ZipFile(template_path, 'r') as z:
        # Resolve worksheet part by sheet name from workbook.xml + rels.
        wbxml = z.read('xl/workbook.xml').decode('utf-8')
        relxml = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        rels = dict(re.findall(r'<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"', relxml))
        for s in RAW_SHEETS:
            m = re.search(r'<sheet[^>]+name="' + re.escape(s) + r'"[^>]+r:id="([^"]+)"', wbxml)
            if m and m.group(1) in rels:
                target = rels[m.group(1)]
                target = target.lstrip('/')
                if not target.startswith('xl/'):
                    target = 'xl/' + target
                raw_sheet_xml[s] = target

    with zipfile.ZipFile(template_path, 'r') as zin, \
         zipfile.ZipFile(prepared, 'r') as zprep, \
         zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        # IMPORTANT: this mapping follows the actual template PivotTable cache IDs.
        cache_map = {1:'R1', 2:'R1S', 3:'R2', 4:'R6', 5:'R6 OC', 6:'R1 OC'}
        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)
            sheet_to_replace = next((s for s, part in raw_sheet_xml.items() if part == name), None)
            if name == 'xl/styles.xml' and name in zprep.namelist():
                data = zprep.read(name)
            elif sheet_to_replace is not None and name in zprep.namelist():
                data = zprep.read(name)

            m = re.fullmatch(r'xl/pivotCache/pivotCacheDefinition([1-6])\.xml', name)
            if m:
                i = int(m.group(1))
                s = cache_map[i]
                # Keep the template's original source width; only change the row end.
                xml = data.decode('utf-8')
                src = re.search(r'<worksheetSource ref="A1:([A-Z]+)\d+" sheet="([^"]+)"', xml)
                if src:
                    col_end = src.group(1)
                    row_end = used.get(s, 0) + 1
                    ref = f'A1:{col_end}{max(1,row_end)}'
                    xml = re.sub(r'<worksheetSource ref="[^"]+" sheet="([^"]+)"',
                                 lambda mm: f'<worksheetSource ref="{ref}" sheet="{mm.group(1)}"', xml, count=1)
                if 'refreshOnLoad=' in xml:
                    xml = re.sub(r'refreshOnLoad="[^"]*"', 'refreshOnLoad="1"', xml, count=1)
                else:
                    xml = xml.replace('<pivotCacheDefinition ', '<pivotCacheDefinition refreshOnLoad="1" ', 1)
                if 'enableRefresh=' in xml:
                    xml = re.sub(r'enableRefresh="[^"]*"', 'enableRefresh="1"', xml, count=1)
                else:
                    xml = xml.replace('<pivotCacheDefinition ', '<pivotCacheDefinition enableRefresh="1" ', 1)
                data = xml.encode('utf-8')

            m = re.fullmatch(r'xl/pivotTables/pivotTable([1-6])\.xml', name)
            if m:
                xml = data.decode('utf-8')
                xml = re.sub(r'(<dataField\b[^>]*name="Count - Task No"[^>]*?)\snumFmtId="[^"]*"',
                             r'\1 numFmtId="164"', xml, count=1)
                data = xml.encode('utf-8')

            zout.writestr(item, data)

    for q in (raw_path, prepared):
        try: q.unlink()
        except Exception: pass
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
                df_ui = process(pd.read_excel(raw_stream, sheet_name=s, dtype=object), is_r2=(s == "R2"))
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
