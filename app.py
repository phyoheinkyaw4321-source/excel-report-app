import io, re, shutil, zipfile, tempfile, os
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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

RAW_SHEETS = ["R1","R2","R6","R1 OC","R6 OC"]
# R1 source sheet contains both R1 and R1S records; preserve the source Region in RE.
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

def process(df, is_r2=False, preserve_blank_rows=False, source_re=None):
    # R1 must retain every source row, including rows where Region is blank.
    if not preserve_blank_rows:
        df=df.dropna(how="all")
    df=df.reset_index(drop=True).copy()
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

    if not is_r2:
        # R1 and R1S source tabs are combined into the single output R1 sheet.
        # The RE tag comes from the source tab, not from Region (which may be blank).
        region=next((c for c in df.columns if key(c)=="region"),None)
        township=next((c for c in df.columns if key(c)=="township"),None)
        if township:
            if source_re is not None:
                re_vals=list(source_re)
            else:
                re_vals=df[region].map(norm) if region else [""]*len(df)
            df.insert(df.columns.get_loc(township)+1,"RE",re_vals)
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
    """Build report data while preserving the template's native PivotTable parts.

    The single raw R1 tab contains both R1 and R1S records; Region values are copied to RE.
    Only raw worksheet XML is copied into the original template package; the
    native PivotTable/cache/layout/format parts are retained and their sources
    are refreshed when Excel opens the generated workbook.
    """
    raw_stream = io.BytesIO(raw_bytes)
    xls = pd.ExcelFile(raw_stream)
    raw_wb = load_workbook(io.BytesIO(raw_bytes), read_only=True, data_only=True)
    out = Path(tempfile.mktemp(suffix=".xlsx"))
    prepared = Path(tempfile.mktemp(suffix=".xlsx"))
    shutil.copy2(template_path, prepared)
    wb = load_workbook(prepared)
    used = {}

    aliases = {"R1 OC": ["R1 OC", "R1OC"], "R6 OC": ["R6 OC", "R6OC"]}

    def read_source_sheet(sheet_name):
        if sheet_name not in xls.sheet_names:
            raise ValueError(f"Raw Excel ထဲမှာ '{sheet_name}' Sheet မတွေ့ပါ")
        raw_stream.seek(0)
        frame = pd.read_excel(raw_stream, sheet_name=sheet_name, dtype=object)
        # Preserve meaningful blank rows and the source row count for R1/R1S.
        expected = max(0, raw_wb[sheet_name].max_row - 1)
        if len(frame) < expected:
            extra = pd.DataFrame([{c: None for c in frame.columns} for _ in range(expected-len(frame))])
            frame = pd.concat([frame, extra], ignore_index=True)
        return frame

    # Raw workbook has one R1 sheet containing both R1 and R1S records.
    # Use its Region column to populate RE; do not look for a separate R1S sheet.
    combined_r1 = read_source_sheet("R1")
    region_col = next((c for c in combined_r1.columns if key(c) == "region"), None)
    if region_col is None:
        raise ValueError("Raw Excel ရဲ့ R1 Sheet ထဲမှာ 'Region' Column မတွေ့ပါ")
    r1_re = combined_r1[region_col].map(norm).tolist()

    for logical in RAW_SHEETS:
        if logical == "R1":
            source_df = combined_r1
            df = process(source_df, preserve_blank_rows=True, source_re=r1_re)
        else:
            actual = logical
            if actual not in xls.sheet_names and logical in aliases:
                actual = next((a for a in aliases[logical] if a in xls.sheet_names), None)
            if actual not in xls.sheet_names:
                raise ValueError(f"Raw Excel ထဲမှာ '{logical}' Sheet မတွေ့ပါ")
            raw_stream.seek(0)
            source_df = pd.read_excel(raw_stream, sheet_name=actual, dtype=object)
            df = process(source_df, is_r2=(logical == "R2"))

        if logical not in wb.sheetnames:
            raise ValueError(f"Template ထဲမှာ '{logical}' Sheet မတွေ့ပါ")
        ws = wb[logical]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        hmap = {key(v): i + 1 for i, v in enumerate(headers) if v is not None}

        # Clear old values only; preserve template formatting, widths and panes.
        if ws.max_row >= 2:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.value = None

        for ri, row in enumerate(df.to_dict("records"), start=2):
            for col, val in row.items():
                ci = hmap.get(key(col))
                if ci:
                    cell = ws.cell(ri, ci)
                    if key(col) == "task status":
                        # Force Task Status to text; never store it as an Excel date.
                        cell.value = None if val is None or (isinstance(val, float) and pd.isna(val)) else str(val)
                        cell.number_format = "General"
                    else:
                        cell.value = val

        # Explicitly keep every Task Status cell as General/text display.
        status_ci = hmap.get("task status")
        if status_ci:
            for rr in range(2, len(df) + 2):
                ws.cell(rr, status_ci).number_format = "General"
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(2, len(df)+1)}"
        used[logical] = len(df)

    raw_wb.close()
    wb.save(prepared)
    wb.close()

    # Resolve worksheet XML paths from the template's own workbook relationships.
    import xml.etree.ElementTree as ET
    main_ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg_rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
    with zipfile.ZipFile(template_path, "r") as ztemplate:
        workbook_root = ET.fromstring(ztemplate.read("xl/workbook.xml"))
        rel_root = ET.fromstring(ztemplate.read("xl/_rels/workbook.xml.rels"))
        targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rel_root.findall(f"{{{pkg_rel_ns}}}Relationship")}
        sheet_parts = {}
        for sh in workbook_root.find(f"{{{main_ns}}}sheets"):
            rid = sh.attrib[f"{{{rel_ns}}}id"]
            target = targets[rid].lstrip("/")
            sheet_parts[sh.attrib["name"]] = target if target.startswith("xl/") else "xl/" + target

    # Cache IDs in this template are 1, 3, 4, 5, 6 (there is no cache 2).
    cache_map = {1: "R1", 3: "R2", 4: "R6", 5: "R1 OC", 6: "R6 OC"}
    with zipfile.ZipFile(template_path, "r") as zin, \
         zipfile.ZipFile(prepared, "r") as zprep, \
         zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        prep_names = set(zprep.namelist())
        for item in zin.infolist():
            name = item.filename
            data = zin.read(name)

            # Replace only the five raw-data worksheet XML files.
            replace_sheet = next((sheet for sheet in RAW_SHEETS if sheet_parts.get(sheet) == name), None)
            if replace_sheet and name in prep_names:
                data = zprep.read(name)

            cache_match = re.fullmatch(r"xl/pivotCache/pivotCacheDefinition(\d+)\.xml", name)
            if cache_match:
                cache_id = int(cache_match.group(1))
                if cache_id in cache_map:
                    sheet = cache_map[cache_id]
                    ws_headers = [wb_header for wb_header in []]  # no workbook handles remain open
                    # The template worksheet header width is preserved; row count is based on new data.
                    temp_wb = load_workbook(prepared, read_only=True, data_only=False)
                    temp_ws = temp_wb[sheet]
                    last_col = get_column_letter(temp_ws.max_column)
                    last_row = used[sheet] + 1
                    temp_wb.close()
                    ref = f"A1:{last_col}{max(1, last_row)}"
                    xml = data.decode("utf-8")
                    xml = re.sub(r'<worksheetSource\b[^>]*/>',
                                 f'<worksheetSource ref="{ref}" sheet="{sheet}"/>',
                                 xml, count=1)
                    # Refresh from the new text-valued source data when opened in Excel.
                    if re.search(r'\brefreshOnLoad="[^"]*"', xml):
                        xml = re.sub(r'\brefreshOnLoad="[^"]*"', 'refreshOnLoad="1"', xml, count=1)
                    else:
                        xml = xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
                    if re.search(r'\benableRefresh="[^"]*"', xml):
                        xml = re.sub(r'\benableRefresh="[^"]*"', 'enableRefresh="1"', xml, count=1)
                    else:
                        xml = xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition enableRefresh="1" ', 1)
                    data = xml.encode("utf-8")

            # Add RE as a Report Filter to the combined R1 PivotTable only.
            # RE is column 23 in the R1 source sheet / cache field index 22.
            if name == "xl/pivotCache/pivotCacheDefinition1.xml":
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                ET.register_namespace("", ns)
                root = ET.fromstring(data)
                fields = root.find(f"{{{ns}}}cacheFields")
                if fields is not None and not any(f.attrib.get("name") == "RE" for f in fields):
                    new_field = ET.Element(f"{{{ns}}}cacheField", {"name": "RE"})
                    ET.SubElement(new_field, f"{{{ns}}}sharedItems")
                    fields.insert(22, new_field)
                    fields.set("count", str(len(fields)))
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            if name == "xl/pivotTables/pivotTable1.xml":
                ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                ET.register_namespace("", ns)
                root = ET.fromstring(data)
                pfields = root.find(f"{{{ns}}}pivotFields")
                if pfields is not None and len(pfields) == 100:
                    pfield = ET.Element(f"{{{ns}}}pivotField", {"axis": "axisPage", "showAll": "0", "compact": "0"})
                    pfields.insert(22, pfield)
                    pfields.set("count", str(len(pfields)))
                page_fields = root.find(f"{{{ns}}}pageFields")
                if page_fields is not None and not any(x.attrib.get("fld") == "22" for x in page_fields):
                    ET.SubElement(page_fields, f"{{{ns}}}pageField", {"fld": "22", "hier": "-1"})
                    page_fields.set("count", str(len(page_fields)))
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

            # Correct PivotTable count fields that otherwise display as dates (e.g. 5-Jan-00).
            if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", name):
                xml = data.decode("utf-8")
                xml = re.sub(r'(<dataField\b[^>]*\bnumFmtId=)"\d+"', r'\1"0"', xml)
                data = xml.encode("utf-8")

            zout.writestr(item, data)

    return out.read_bytes(), used


# ---------------- UI ONLY ----------------
# Processing / business / Pivot logic above this line is intentionally unchanged.

st.markdown("""
<style>
:root{
  --navy:#18212f;
  --navy2:#242f3f;
  --line:#d9dee6;
  --soft:#f1f3f6;
  --text:#2b3038;
}
.block-container{max-width:1180px;padding-top:1.2rem;padding-bottom:2rem;}
body,[data-testid="stAppViewContainer"]{background:#f1f3f6;color:#2b3038;}
[data-testid="stHeader"]{background:#18212f;}

.main-title{color:#18212f;font-size:2.15rem;font-weight:900;letter-spacing:-.5px;margin-bottom:0;}
.sub-title{color:#607086;font-size:1rem;margin-top:.15rem;margin-bottom:1.15rem;}
.workflow{
  display:flex;align-items:center;justify-content:center;gap:14px;
  color:var(--navy);font-weight:800;font-size:.95rem;
  background:#ffffff;border:1px solid var(--line);border-radius:12px;
  padding:10px 16px;margin:0 0 18px;
}
.section{
  color:#18212f;font-size:1.05rem;font-weight:900;
  letter-spacing:.2px;margin:18px 0 8px;
}
.login-wrap{
  max-width:430px;margin:7vh auto 0;background:#fff;
  border:1px solid #dce3eb;border-radius:18px;padding:30px 30px 26px;
  box-shadow:0 12px 35px rgba(11,35,66,.10);
}
.login-title{color:#18212f;font-size:1.55rem;font-weight:900;text-align:center;margin-bottom:3px;}
.login-sub{text-align:center;color:#718096;font-size:.92rem;margin-bottom:18px;}
div[data-testid="stTextInput"] label{font-weight:800;color:var(--navy);}
div[data-testid="stFileUploader"]{
  border:1px dashed #aebdce;border-radius:12px;padding:5px;background:#ffffff;
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
  background:#18212f;color:#fff;padding:11px 13px;
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
.ts-name{color:#2b3038;font-size:.88rem;font-weight:700;line-height:1.2;}
.ts-count{color:var(--navy);font-size:.91rem;font-weight:900;text-align:right;}
.ts-total{
  display:grid;grid-template-columns:1fr auto;gap:8px;
  padding:10px 13px;background:#f1f3f6;color:#2b3038;
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

template=Path(__file__).with_name("Auto_Pivot_Live_25-Nov_R1_Combined_FIXED.xlsx")
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
        df=process(pd.read_excel(io.BytesIO(raw_bytes),sheet_name=actual),is_r2=(logical=="R2"),preserve_blank_rows=(logical=="R1"))
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
            # Build combined R1/R1S data and preserve native PivotTable parts.
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

