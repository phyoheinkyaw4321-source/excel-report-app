
import io, re, shutil, zipfile, tempfile, time
from concurrent.futures import ThreadPoolExecutor
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
    """Build the report without letting openpyxl rewrite the native PivotTable parts.

    The template workbook is treated as the source of truth for all PivotTable,
    PivotCache, styles, filters and layout.  openpyxl is used only to prepare the
    six raw worksheets; their worksheet XML is then copied back into the original
    template ZIP.  This prevents PivotTable relationships/cache definitions from
    being damaged during save.
    """
    raw_stream = io.BytesIO(raw_bytes)
    xls = pd.ExcelFile(raw_stream)

    out = Path(tempfile.mktemp(suffix=".xlsx"))
    prepared = Path(tempfile.mktemp(suffix=".xlsx"))
    shutil.copy2(template_path, prepared)

    # Update only the raw sheets in a temporary copy.  Do not use this saved
    # workbook as the final workbook because openpyxl can rewrite native Pivot
    # XML parts.  We will copy only the six raw worksheet XML files below.
    wb = load_workbook(prepared)
    used = {}
    schedule_style_ids = {}
    for _s in RAW_SHEETS:
        if _s in wb.sheetnames:
            _ws0 = wb[_s]
            for _c in range(1, _ws0.max_column + 1):
                if key(_ws0.cell(1, _c).value) in {"schedule date", "date"}:
                    schedule_style_ids[_s] = _ws0.cell(2, _c).style_id if _ws0.max_row >= 2 else 0
                    break
    raw_sheet_xml = {
        "R1 OC": "xl/worksheets/sheet2.xml",
        "R6":    "xl/worksheets/sheet3.xml",
        "R6 OC": "xl/worksheets/sheet4.xml",
        "R1":    "xl/worksheets/sheet5.xml",
        "R1S":   "xl/worksheets/sheet6.xml",
        "R2":    "xl/worksheets/sheet7.xml",
    }

    for s in RAW_SHEETS:
        if s not in xls.sheet_names or s not in wb.sheetnames:
            continue

        raw_stream.seek(0)
        df = process(
            pd.read_excel(raw_stream, sheet_name=s, dtype=object),
            is_r2=(s == "R2")
        )
        ws = wb[s]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        hmap = {key(v): i + 1 for i, v in enumerate(headers) if v is not None}

        # Clear old values only.  Keep template styles, widths and panes.
        for row in ws.iter_rows(
            min_row=2, max_row=ws.max_row, max_col=ws.max_column
        ):
            for cell in row:
                cell.value = None

        for ri, row in enumerate(df.to_dict("records"), start=2):
            for col, val in row.items():
                ci = hmap.get(key(col))
                if ci:
                    cell = ws.cell(ri, ci)
                    cell.value = val
                    if key(col) == "task status":
                        cell.number_format = "General"

        # Task Status must never be interpreted/displayed as a date.
        for hname, ci in hmap.items():
            if hname == "task status":
                for rr in range(2, len(df) + 2):
                    ws.cell(rr, ci).number_format = "General"

        used[s] = len(df)

    wb.save(prepared)
    wb.close()

    # Start from the ORIGINAL template ZIP so all native PivotTable parts,
    # relationships, styles and layout remain byte-for-byte intact except for
    # the cache source/refresh settings intentionally patched below.
    with zipfile.ZipFile(template_path, "r") as zin, \
         zipfile.ZipFile(prepared, "r") as zprep, \
         zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:

        # Use the prepared workbook's styles.xml so worksheet style IDs and the
        # style table remain aligned. Then append only the Schedule Date styles.
        styles_xml = zprep.read("xl/styles.xml").decode("utf-8")
        mfmt = re.search(r'<numFmts\b[^>]*>', styles_xml)
        if mfmt and 'formatCode="dd-mmm-yyyy"' not in styles_xml:
            styles_xml = styles_xml[:mfmt.end()] + '<numFmt numFmtId="172" formatCode="dd-mmm-yyyy"/>' + styles_xml[mfmt.end():]
        mxf = re.search(r'<cellXfs\b[^>]*count="(\d+)"[^>]*>(.*?)</cellXfs>', styles_xml, re.S)
        date_style_ids = {}
        if mxf:
            old_count = int(mxf.group(1))
            body = mxf.group(2)
            entries = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', body, re.S)
            for _s in RAW_SHEETS:
                _sheet_xml = zprep.read(raw_sheet_xml[_s]).decode("utf-8")
                _m = re.search(r'<c\b[^>]*r="A2"[^>]*s="(\d+)"[^>]*>', _sheet_xml)
                if _m:
                    old_id = int(_m.group(1))
                    if old_id < len(entries):
                        base_xf = entries[old_id]
                        date_xf = re.sub(r'numFmtId="[^"]*"', 'numFmtId="172"', base_xf, count=1)
                        if 'numFmtId=' not in date_xf:
                            date_xf = date_xf.replace('<xf ', '<xf numFmtId="172" ', 1)
                        date_style_ids[_s] = old_count + len(date_style_ids)
                        body += date_xf
            if date_style_ids:
                styles_xml = styles_xml[:mxf.start(2)] + body + styles_xml[mxf.end(2):]
                styles_xml = re.sub(r'(<cellXfs\b[^>]*count=")\d+("[^>]*>)', rf'\g<1>{old_count+len(date_style_ids)}\g<2>', styles_xml, count=1)
        # Cache 1..6 -> raw sheet mapping in the source-of-truth template.
        cache_map = {
            1: "R1",
            2: "R1S",
            3: "R2",
            4: "R6",
            5: "R1 OC",
            6: "R6 OC",
        }

        # Copy every original template part, replacing only raw worksheet XML and
        # the six PivotCache definitions/data-field number format metadata.
        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)

            # Replace raw worksheet XML generated from the processed data.
            replace_sheet = next(
                (sheet for sheet, xml_name in raw_sheet_xml.items() if xml_name == name),
                None
            )
            if replace_sheet is not None and name in zprep.namelist():
                data = zprep.read(name)
                xml = data.decode("utf-8")
                if replace_sheet in date_style_ids:
                    new_sid = date_style_ids[replace_sheet]
                    def _date_style_cell(mm):
                        tag = re.sub(r'\s+s="\d+"', '', mm.group(0))
                        return tag[:-1] + f' s="{new_sid}">'
                    xml = re.sub(r'<c\b[^>]*r="A([2-9]\d*)"[^>]*>', _date_style_cell, xml)
                data = xml.encode("utf-8")

            if name == "xl/styles.xml":
                data = styles_xml.encode("utf-8")

            m = re.fullmatch(r"xl/pivotCache/pivotCacheDefinition([1-6])\.xml", name)
            if m:
                i = int(m.group(1))
                sheet = cache_map[i]
                nrows = used.get(sheet, 0) + 1  # header + data rows

                # Keep the template's original source columns exactly; only
                # change the row end to the current processed data length.
                src_match = re.search(r'<worksheetSource ref="([A-Z]+)\d+"', data.decode("utf-8"))
                max_col_letter = src_match.group(1) if src_match else get_column_letter(
                    load_workbook(prepared, read_only=True, data_only=False)[sheet].max_column
                )
                ref = f"A1:{max_col_letter}{max(1, nrows)}"

                xml = data.decode("utf-8")
                xml = re.sub(
                    r'<worksheetSource ref="[^"]+" sheet="([^"]+)"',
                    lambda mm: f'<worksheetSource ref="{ref}" sheet="{mm.group(1)}"',
                    xml,
                    count=1,
                )

                # Force refresh on open and keep cache enabled.
                if "refreshOnLoad=" in xml:
                    xml = re.sub(r'refreshOnLoad="[^"]*"', 'refreshOnLoad="1"', xml, count=1)
                else:
                    xml = xml.replace(
                        '<pivotCacheDefinition ',
                        '<pivotCacheDefinition refreshOnLoad="1" enableRefresh="1" ',
                        1,
                    )
                if "enableRefresh=" not in xml:
                    xml = xml.replace('<pivotCacheDefinition ', '<pivotCacheDefinition enableRefresh="1" ', 1)

                # Keep recordCount consistent with the new raw data count. Excel
                # will rebuild the cache on refresh; this avoids a stale count hint.
                xml = re.sub(r'recordCount="[^"]*"', f'recordCount="{max(0, nrows - 1)}"', xml, count=1)
                data = xml.encode("utf-8")

            # Explicitly force all six Pivot value fields to General/Number.
            mpt = re.fullmatch(r"xl/pivotTables/pivotTable([1-6])\.xml", name)
            if mpt:
                xml = data.decode("utf-8")
                xml = re.sub(
                    r'(<dataField\b[^>]*name="Count - Task No"[^>]*?)\snumFmtId="[^"]*"',
                    r'\1 numFmtId="0"',
                    xml,
                    count=1,
                )
                data = xml.encode("utf-8")

            zout.writestr(item, data)

    try:
        prepared.unlink()
    except Exception:
        pass

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
    runner = st.empty()
    status_line = st.empty()
    try:
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(build_workbook, uploaded.getvalue(), template)
            frames = [
                "🏃‍♂️  Generating report .  ",
                " 🏃‍♂️ Generating report .. ",
                "  🏃‍♂️ Generating report ...",
                "   🏃‍♂️ Generating report ....",
                "  🏃‍♂️ Generating report ... ",
                " 🏃‍♂️ Generating report ..  ",
            ]
            i = 0
            while not future.done():
                runner.markdown(
                    f'<div style="background:#eef3f8;border:1px solid #b9c6d4;border-radius:10px;padding:15px 18px;text-align:center;color:#0b1f3a;font-size:22px;font-weight:800;">{frames[i % len(frames)]}</div>',
                    unsafe_allow_html=True,
                )
                status_line.info("⏳ Raw Excel ကို process လုပ်ပြီး Live Pivot structure ကို ထိန်းသိမ်းနေပါတယ်…")
                i += 1
                time.sleep(0.35)
            data, counts = future.result()

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

        runner.empty()
        status_line.empty()
        st.session_state["out"] = data
        st.session_state["counts"] = counts
        st.session_state["task_status_counts"] = task_status_counts
        st.session_state["report_date"] = report_date
        st.success("✅ Report Generated — Final Excel is ready.")
    except Exception as e:
        runner.empty()
        status_line.empty()
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
