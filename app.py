
import io, re, shutil, zipfile, tempfile
import xml.etree.ElementTree as ET
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

def process(df, is_r2=False, preserve_rows=True):
    """Process rows without deleting blank/raw rows. Raw row positions stay stable."""
    df = df.copy()
    if not preserve_rows:
        df = df.dropna(how="all").reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)

    for c in ["TKT / POI", "Classification", "RC", "RE"]:
        if c in df.columns:
            df.drop(columns=c, inplace=True)

    service = next((c for c in df.columns if key(c) == "service area"), None)
    task = next((c for c in df.columns if key(c) == "task no"), None)
    root = next((c for c in df.columns if key(c) == "root cause"), None)

    if task:
        tt = df[task].map(task_type)
        cc = [classification(t, s, r) for t, s, r in zip(
            tt,
            df[service].map(norm) if service else [""] * len(df),
            df[root].map(norm) if root else [""] * len(df)
        )]
        pos = df.columns.get_loc(service) + 1 if service else 0
        df.insert(pos, "TKT / POI", tt)
        df.insert(pos + 1, "Classification", cc)

    car = next((c for c in df.columns if key(c) == "car id"), None)
    eng = next((c for c in df.columns if key(c) == "lan engineer name"), None)
    repay = next((c for c in df.columns if key(c) == "repay status"), None)
    if car and eng:
        tmp = pd.DataFrame({"car": df[car].map(norm), "eng": df[eng].map(norm)})
        if repay:
            bad = df[repay].map(norm).str.lower().str.contains(
                r"engr\s*leave|route\s*cancel", regex=True, na=False
            )
            tmp = tmp[~bad]
        tmp = tmp[(tmp.car != "") & (tmp.eng != "")]
        counts = tmp.groupby("car").eng.apply(lambda x: len(set(x))).to_dict()
        rc = []
        for x in df[car].map(norm):
            n = counts.get(x, 0)
            rc.append("" if not x else ("Car Share Plus" if n >= 3 else "Car Share" if n == 2 else "One Car" if n == 1 else ""))
        df.insert(df.columns.get_loc(car), "RC", rc)

    if is_r2:
        township = next((c for c in df.columns if key(c) == "township"), None)
        if township:
            vals = []
            for x in df[township]:
                t = key(x)
                if not t:
                    vals.append("")
                    continue
                base = re.split(r"\s*\(|\s+ward\b", t, maxsplit=1)[0].strip()
                vals.append("R2S" if (t in R2S_TOWNSHIPS or base in R2S_TOWNSHIPS) else "R2")
            df.insert(df.columns.get_loc(township) + 1, "RE", vals)
    return df


def _raw_last_row(ws):
    last = 1
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if any(v is not None for v in row):
            last = r
    return last


def _final_column_map(raw_max_col, is_r2=False):
    """Return raw-column-index -> output-column-index after calculated columns."""
    mp = {}
    for raw_col in range(1, raw_max_col + 1):
        shift = 2  # TKT/POI + Classification after Service Area (raw col 10)
        if raw_col <= 10:
            shift = 0
        # RC before Car ID (raw col 40)
        if raw_col >= 40:
            shift += 1
        # RE after Township (raw col 20), R2 only
        if is_r2 and raw_col >= 21:
            shift += 1
        mp[raw_col] = raw_col + shift
    return mp


def _raw_last_row(ws):
    last = 1
    for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        if any(v is not None for v in row):
            last = r
    return last


def _col_letter(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _excel_xml_value(cell):
    import datetime
    from openpyxl.utils.datetime import to_excel
    v = cell.value
    if v is None:
        return None, None
    if cell.data_type == "e":
        return "e", str(v)
    if cell.data_type == "b":
        return "b", "1" if v else "0"
    if cell.data_type == "f":
        return "f", str(v)
    if isinstance(v, (datetime.datetime, datetime.date, datetime.time)):
        return None, repr(to_excel(v))
    if isinstance(v, (int, float)):
        return None, repr(v)
    return "inlineStr", str(v)


def build_workbook(raw_bytes, template_path):
    """Build final workbook while preserving raw cell positions/values and native PivotTables."""
    import datetime
    import copy
    import xml.etree.ElementTree as ET
    from openpyxl.utils.datetime import to_excel

    NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", NS)
    STYLE_MAP = {
        "General": 164,
        "d-mmm-yy": 165,
        "h:mm": 166,
        "h:mm:ss AM/PM": 170,
        "#,##0": 171,
        "h:mm AM/PM": 18,
        "0.00E+00": 11,
        "[$-409]d\\-mmm\\-yyyy;@": 172,
    }

    raw_path = Path(tempfile.mktemp(suffix="_raw.xlsx"))
    out_path = Path(tempfile.mktemp(suffix="_final.xlsx"))
    raw_path.write_bytes(raw_bytes)

    raw_wb = load_workbook(raw_path, data_only=False, read_only=False)
    with zipfile.ZipFile(template_path, "r") as zin:
        styles_root = ET.fromstring(zin.read("xl/styles.xml"))
        cell_xfs = styles_root.find(f"{{{NS}}}cellXfs")
        num_fmts = styles_root.find(f"{{{NS}}}numFmts")
        if num_fmts is None:
            num_fmts = ET.Element(f"{{{NS}}}numFmts", {"count": "0"})
            styles_root.insert(0, num_fmts)
        existing_numfmt = {e.attrib["formatCode"]: int(e.attrib["numFmtId"]) for e in num_fmts}
        for code, fmt_id in STYLE_MAP.items():
            if fmt_id >= 164 and code not in existing_numfmt:
                ET.SubElement(num_fmts, f"{{{NS}}}numFmt", {"numFmtId": str(fmt_id), "formatCode": code})
                existing_numfmt[code] = fmt_id
        num_fmts.set("count", str(len(list(num_fmts))))
        style_cache = {}

        def style_lookup(base_id, fmt):
            fmt_id = STYLE_MAP.get(fmt, 164)
            k = (base_id, fmt_id)
            if k in style_cache:
                return style_cache[k]
            base_xf = cell_xfs[base_id]
            if int(base_xf.attrib.get("numFmtId", "0")) == fmt_id:
                style_cache[k] = base_id
                return base_id
            new_xf = copy.deepcopy(base_xf)
            new_xf.set("numFmtId", str(fmt_id))
            cell_xfs.append(new_xf)
            new_id = len(cell_xfs) - 1
            style_cache[k] = new_id
            return new_id

        sheet_xml = {
            "R1 OC": "xl/worksheets/sheet2.xml",
            "R6": "xl/worksheets/sheet3.xml",
            "R6 OC": "xl/worksheets/sheet4.xml",
            "R1": "xl/worksheets/sheet5.xml",
            "R1S": "xl/worksheets/sheet6.xml",
            "R2": "xl/worksheets/sheet7.xml",
        }
        replacements = {}
        used = {}

        for s in RAW_SHEETS:
            if s not in raw_wb.sheetnames:
                continue
            ws = raw_wb[s]
            last_row = _raw_last_row(ws)
            n = max(0, last_row - 1)

            raw_stream = io.BytesIO(raw_bytes)
            df = process(
                pd.read_excel(raw_stream, sheet_name=s, dtype=object),
                is_r2=(s == "R2"),
                preserve_rows=True,
            )
            if len(df) < n:
                df = df.reindex(range(n))
            elif len(df) > n:
                df = df.iloc[:n].copy()

            root = ET.fromstring(zin.read(sheet_xml[s]))
            sheet_data = root.find(f"{{{NS}}}sheetData")
            existing_rows = sheet_data.findall(f"{{{NS}}}row")
            style_by_col = {}
            for rr in existing_rows[1:]:
                for c in rr.findall(f"{{{NS}}}c"):
                    col = re.sub(r"\d+$", "", c.attrib["r"])
                    if col not in style_by_col and c.attrib.get("s") is not None:
                        style_by_col[col] = c.attrib["s"]
            for rr in list(sheet_data):
                if int(rr.attrib.get("r", "0")) >= 2:
                    sheet_data.remove(rr)

            raw_cols = ws.max_column
            is_r2 = s == "R2"
            max_out = raw_cols + 3 + (1 if is_r2 else 0)
            colmap = _final_column_map(raw_cols, is_r2)
            calc = {k: list(df[k]) if k in df.columns else [""] * n for k in ["TKT / POI", "Classification", "RC", "RE"]}

            for rr in range(2, last_row + 1):
                idx = rr - 2
                row = ET.Element(f"{{{NS}}}row", {"r": str(rr)})
                cells = []
                for oc in range(1, max_out + 1):
                    c = ET.SubElement(row, f"{{{NS}}}c", {"r": f"{_col_letter(oc)}{rr}"})
                    base_style = style_by_col.get(_col_letter(oc))
                    if base_style is not None:
                        c.set("s", base_style)
                    cells.append(c)

                for raw_c in range(1, raw_cols + 1):
                    out_c = colmap[raw_c]
                    c = cells[out_c - 1]
                    src = ws.cell(rr, raw_c)
                    for child in list(c):
                        c.remove(child)
                    t, v = _excel_xml_value(src)
                    if t == "inlineStr":
                        c.set("t", "inlineStr")
                        is_el = ET.SubElement(c, f"{{{NS}}}is")
                        t_el = ET.SubElement(is_el, f"{{{NS}}}t")
                        if v and (v[0].isspace() or v[-1].isspace()):
                            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        t_el.text = v
                    elif t == "f":
                        ET.SubElement(c, f"{{{NS}}}f").text = v
                    elif t in ("e", "b"):
                        c.set("t", t)
                        ET.SubElement(c, f"{{{NS}}}v").text = v
                    elif v is not None:
                        c.attrib.pop("t", None)
                        ET.SubElement(c, f"{{{NS}}}v").text = str(v)

                    base_style = style_by_col.get(_col_letter(out_c))
                    if src.value is not None and base_style is not None:
                        c.set("s", str(style_lookup(int(base_style), src.number_format)))

                def put_calc(out_idx, value):
                    c = cells[out_idx]
                    for child in list(c):
                        c.remove(child)
                    c.attrib.pop("t", None)
                    if value is None or (not isinstance(value, (list, dict, tuple)) and pd.isna(value)):
                        return
                    if isinstance(value, str):
                        c.set("t", "inlineStr")
                        ET.SubElement(ET.SubElement(c, f"{{{NS}}}is"), f"{{{NS}}}t").text = value
                    elif isinstance(value, bool):
                        c.set("t", "b")
                        ET.SubElement(c, f"{{{NS}}}v").text = "1" if value else "0"
                    else:
                        ET.SubElement(c, f"{{{NS}}}v").text = str(value)

                put_calc(10, calc["TKT / POI"][idx])
                put_calc(11, calc["Classification"][idx])
                if is_r2:
                    put_calc(22, calc["RE"][idx])
                    put_calc(42, calc["RC"][idx])
                else:
                    put_calc(41, calc["RC"][idx])

                # Task Status is always General.
                for c_idx, h in enumerate([x.text for x in []], 1):
                    pass
                for c in cells:
                    # header is handled below by template; this branch intentionally stays cheap
                    pass
                sheet_data.append(row)

            dim = root.find(f"{{{NS}}}dimension")
            if dim is not None:
                dim.set("ref", f"A1:{_col_letter(max_out)}{max(1, last_row)}")
            replacements[sheet_xml[s]] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            used[s] = n

        # Build final ZIP from the untouched template so PivotTable relationships/layout remain native.
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            cache_map = {1: "R1", 2: "R1S", 3: "R2", 4: "R6", 5: "R1 OC", 6: "R6 OC"}
            for item in zin.infolist():
                name = item.filename
                data = zin.read(name)
                if name in replacements:
                    data = replacements[name]
                if name == "xl/styles.xml":
                    cell_xfs.set("count", str(len(list(cell_xfs))))
                    data = ET.tostring(styles_root, encoding="utf-8", xml_declaration=True)

                m = re.fullmatch(r"xl/pivotCache/pivotCacheDefinition([1-6])\.xml", name)
                if m:
                    s = cache_map[int(m.group(1))]
                    raw_cols = raw_wb[s].max_column
                    max_out = raw_cols + 3 + (1 if s == "R2" else 0)
                    end = used[s] + 1
                    xml = data.decode("utf-8")
                    xml = re.sub(
                        r'<worksheetSource ref="[^"]+" sheet="([^"]+)"',
                        lambda mm: f'<worksheetSource ref="A1:{_col_letter(max_out)}{end}" sheet="{mm.group(1)}"',
                        xml,
                        count=1,
                    )
                    xml = re.sub(r'refreshOnLoad="[^"]*"', 'refreshOnLoad="1"', xml, count=1) if "refreshOnLoad=" in xml else xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
                    xml = re.sub(r'enableRefresh="[^"]*"', 'enableRefresh="1"', xml, count=1) if "enableRefresh=" in xml else xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition enableRefresh="1" ', 1)
                    xml = re.sub(r'recordCount="[^"]*"', f'recordCount="{used[s]}"', xml, count=1)
                    data = xml.encode("utf-8")

                if re.fullmatch(r"xl/pivotTables/pivotTable[1-6]\.xml", name):
                    xml = data.decode("utf-8")
                    xml = re.sub(
                        r'(<dataField\b[^>]*name="Count - Task No"[^>]*?)\snumFmtId="[^"]*"',
                        r'\1 numFmtId="0"',
                        xml,
                        count=1,
                    )
                    data = xml.encode("utf-8")
                zout.writestr(item, data)

    raw_wb.close()
    try:
        raw_path.unlink()
    except Exception:
        pass
    return out_path.read_bytes(), used


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
.stButton button, .stButton button { position: relative; overflow: hidden; }
.stButton button::before { content:"🏃‍♂️"; display:inline-block; margin-right:9px; animation: run 0.55s steps(2) infinite; }
@keyframes run { 0%{ transform:translateX(-2px); } 50%{ transform:translateX(3px); } 100%{ transform:translateX(-2px); } }
.stDownloadButton button {
    min-height: 52px;
    font-size: 18px !important;
    font-weight: 800 !important;
    border-radius: 8px;
}
.stButton button { position: relative; overflow: hidden; }
.stButton button::before { content:"🏃‍♂️"; display:inline-block; margin-right:9px; animation: run 0.55s steps(2) infinite; }
@keyframes run { 0%{ transform:translateX(-2px); } 50%{ transform:translateX(3px); } 100%{ transform:translateX(-2px); } }
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
.sheet-card {
    border: 1px solid #c9d3df;
    border-radius: 10px;
    background: #ffffff;
    box-shadow: 0 2px 8px rgba(11,31,58,.08);
    padding: 0;
    margin-bottom: 22px;
    overflow: hidden;
}
.sheet-card-title { background:#0b1f3a; color:#ffffff; font-size:22px; font-weight:850; padding:12px 16px; }
.sheet-card-sub { color:#0b1f3a; font-size:15px; font-weight:850; letter-spacing:.4px; padding:11px 16px 7px 16px; }
.mini-status-table { width:100%; border-collapse:collapse; font-size:16px; margin-bottom:0; }
.mini-status-table td { border-top:1px solid #d8e0e8; padding:10px 14px; }
.status-name { text-align:left; font-weight:700; color:#17202a; }
.status-count { text-align:right; font-weight:850; color:#0b1f3a; }
.mini-status-table tfoot td { background:#eef3f8; border-top:2px solid #aebccc; }
.total-name { text-align:left; font-weight:900; color:#0b1f3a; }
.total-count { text-align:right; font-weight:900; color:#0b1f3a; font-size:18px; }
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

if uploaded and st.button("GENERATE REPORT", type="primary", use_container_width=True):
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

    preferred = ["Resolved", "Done", "Resolved (Auto)", "Pending", "Cancelled"]
    statuses = [x for x in preferred if x in all_statuses]
    statuses += sorted(x for x in all_statuses if x not in statuses and x.strip())

    # Clean management-style 3-column Task Status cards. Blank statuses are hidden.
    columns = st.columns(3, gap="large")
    for idx, s in enumerate(sheet_order):
        counts_s = st.session_state.get("task_status_counts", {}).get(s, {})
        with columns[idx % 3]:
            if not counts_s:
                continue
            total = sum(counts_s.values())
            rows = []
            for status in statuses:
                if status not in counts_s:
                    continue
                rows.append(
                    f'<tr><td class="status-name">{status}</td><td class="status-count">{counts_s[status]:,}</td></tr>'
                )
            st.markdown(
                f'''<div class="sheet-card">
                    <div class="sheet-card-title">📊 {s}</div>
                    <div class="sheet-card-sub">TASK STATUS</div>
                    <table class="mini-status-table">
                      <tbody>{"".join(rows)}</tbody>
                      <tfoot><tr><td class="total-name">Total</td><td class="total-count">{total:,}</td></tr></tfoot>
                    </table>
                </div>''',
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
