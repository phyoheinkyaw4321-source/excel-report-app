import io, re, shutil, zipfile, tempfile, os, math, datetime as dt
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

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
.stApp { background: #f5f7fa; }
.block-container { padding-top: 0.8rem; max-width: 1400px; }
.stFileUploader, .stButton > button, .stDownloadButton > button { border-radius: 8px; }
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

def _sync_pivot_cache_with_libreoffice(workbook_path, used):
    """Refresh Pivot cache records without replacing the final Excel PivotTables."""
    import subprocess
    import shutil as _shutil

    office = _shutil.which("libreoffice") or _shutil.which("soffice")
    if not office:
        # Excel will still rebuild because refreshOnLoad is enabled below.
        return

    workdir = Path(tempfile.mkdtemp(prefix="pivot_cache_"))
    try:
        source = workdir / workbook_path.name
        shutil.copy2(workbook_path, source)

        # LibreOffice must see the exact current source ranges.
        pre = source.with_suffix(".pre.xlsx")
        with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(pre, "w", zipfile.ZIP_DEFLATED) as zout:
            cache_map = {1:"R1",2:"R1S",3:"R2",4:"R6",5:"R1 OC",6:"R6 OC"}
            for item in zin.infolist():
                data = zin.read(item.filename)
                m = re.fullmatch(r"xl/pivotCache/pivotCacheDefinition([1-6])\.xml", item.filename)
                if m:
                    sname = cache_map[int(m.group(1))]
                    # Use the actual current data row, not the worksheet's stale max_row.
                    # Column count is obtained from the existing cache/source definition.
                    xml = data.decode("utf-8")
                    old_src = re.search(r'<worksheetSource ref="([^"]+)" sheet="([^"]+)"', xml)
                    if old_src:
                        old_ref = old_src.group(1)
                        last_col = re.match(r"[A-Z]+", old_ref.split(":")[-1]).group(0)
                        ref = f"A1:{last_col}{used.get(sname, 0) + 1}"
                        xml = re.sub(r'<worksheetSource ref="[^"]+" sheet="([^"]+)"',
                                     lambda mm: f'<worksheetSource ref="{ref}" sheet="{mm.group(1)}"', xml, count=1)
                    if "refreshOnLoad=" not in xml:
                        xml = xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
                    if "enableRefresh=" not in xml:
                        xml = xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition enableRefresh="1" ', 1)
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
        source.unlink()
        pre.rename(source)

        refreshed_dir = workdir / "refreshed"
        refreshed_dir.mkdir()
        subprocess.run([office, "--headless", "--convert-to", "xlsx", "--outdir", str(refreshed_dir), str(source)],
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True, timeout=180)
        refreshed = refreshed_dir / source.name
        if not refreshed.exists():
            return

        # Copy ONLY cache definition + records parts into the original workbook.
        # Do not copy pivotTables, worksheets, styles, or workbook parts from LO.
        merged = workbook_path.with_suffix(".cachemerged.xlsx")
        with zipfile.ZipFile(workbook_path, "r") as zin, zipfile.ZipFile(refreshed, "r") as zr, zipfile.ZipFile(merged, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.fullmatch(r"xl/pivotCache/(pivotCacheDefinition[1-6]\.xml|pivotCacheRecords[1-6]\.xml)", item.filename):
                    data = zr.read(item.filename)
                    if re.fullmatch(r"xl/pivotCache/pivotCacheDefinition[1-6]\.xml", item.filename):
                        xml = data.decode("utf-8")
                        if "refreshOnLoad=" not in xml:
                            xml = xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition refreshOnLoad="1" ', 1)
                        if "enableRefresh=" not in xml:
                            xml = xml.replace("<pivotCacheDefinition ", '<pivotCacheDefinition enableRefresh="1" ', 1)
                        data = xml.encode("utf-8")
                zout.writestr(item, data)
        workbook_path.unlink()
        merged.replace(workbook_path)
    except Exception:
        # Never fail report generation because the optional cache-builder is unavailable.
        # The workbook still has corrected source ranges, numeric Pivot count formatting,
        # and refreshOnLoad enabled.
        try:
            if 'merged' in locals() and merged.exists():
                merged.unlink()
        except Exception:
            pass
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

def _make_editable_live_pivot_workbook(template_path, processed_path, used):
    """Build the final workbook from the original Excel Pivot template.

    openpyxl can rewrite PivotTable XML in ways that make Excel PivotTables
    appear dead.  Therefore the final file starts from the untouched template;
    only the six raw-data sheet contents and six Pivot cache parts are replaced.
    The original PivotTable definitions remain intact and editable.
    """
    out = processed_path.with_suffix(".live.xlsx")
    raw_sheets = ["R1", "R1S", "R2", "R6", "R1 OC", "R6 OC"]

    def sheet_paths(xlsx_path):
        with zipfile.ZipFile(xlsx_path, "r") as z:
            wbxml = z.read("xl/workbook.xml").decode("utf-8")
            relxml = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
            rid_to_target = {
                m.group(1): m.group(2)
                for m in re.finditer(
                    r'<Relationship\b(?=[^>]*\bId="([^"]+)")(?=[^>]*\bTarget="([^"]+)")[^>]*/>',
                    relxml,
                )
            }
            result = {}
            for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml):
                target = rid_to_target[m.group(2)].lstrip("/")
                result[m.group(1)] = target if target.startswith("xl/") else "xl/" + target
            return result

    template_paths = sheet_paths(template_path)
    processed_paths = sheet_paths(processed_path)

    with zipfile.ZipFile(template_path, "r") as tin, zipfile.ZipFile(processed_path, "r") as pin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in tin.infolist():
            name = item.filename
            data = tin.read(name)

            # The processed workbook may contain additional cell styles (for example
            # the uploaded Raw workbook's Schedule Date format).  Keep that styles
            # table together with the processed raw sheet XML so style IDs remain valid.
            if name == "xl/styles.xml":
                data = pin.read(name)

            # Raw sheets: keep the original sheet layout/header and replace only
            # the actual data rows with the current processed rows.
            elif name in [template_paths[s] for s in raw_sheets if s in template_paths]:
                sname = next(s for s in raw_sheets if template_paths.get(s) == name)
                old = data.decode("utf-8")
                new = pin.read(processed_paths[sname]).decode("utf-8")
                old_sd = re.search(r"<sheetData>.*?</sheetData>", old, re.S)
                new_sd = re.search(r"<sheetData>.*?</sheetData>", new, re.S)
                old_header = re.search(r"<row r=\"1\".*?</row>", old_sd.group(0), re.S).group(0)
                max_row = int(used.get(sname, 0)) + 1
                data_rows = []
                for rm in re.finditer(r"<row r=\"(\d+)\".*?</row>", new_sd.group(0), re.S):
                    if 2 <= int(rm.group(1)) <= max_row:
                        row_xml = rm.group(0)
                        # Keep cell style IDs from the processed workbook.
                        # These styles include the original source number formats,
                        # especially Schedule Date.
                        data_rows.append(row_xml)
                sheet_data = "<sheetData>" + old_header + "".join(data_rows) + "</sheetData>"

                dim = re.search(r'<dimension ref="([^"]+)"', old)
                if dim:
                    old_ref = dim.group(1)
                    last_col = re.match(r"[A-Z]+", old_ref.split(":")[-1]).group(0)
                    old = re.sub(r'<dimension ref="[^"]+"', f'<dimension ref="A1:{last_col}{max_row}"', old, count=1)
                old = re.sub(r"<sheetData>.*?</sheetData>", sheet_data, old, count=1, flags=re.S)
                old = re.sub(r'<autoFilter ref="[^"]+"', f'<autoFilter ref="A1:{last_col}{max_row}"', old, count=1) if '<autoFilter ' in old else old
                data = old.encode("utf-8")

            # New cache definition/records correspond to the processed current data.
            elif re.fullmatch(r"xl/pivotCache/(pivotCacheDefinition[1-6]\.xml|pivotCacheRecords[1-6]\.xml)", name):
                data = pin.read(name)

            # Keep the original Excel PivotTable XML intact except for the
            # Count - Task No display format.  Do NOT add unsupported
            # PivotTable attributes: the template already contains live
            # editable PivotTable definitions.
            elif re.fullmatch(r"xl/pivotTables/pivotTable[1-6]\.xml", name):
                xml = data.decode("utf-8")
                xml = re.sub(r'(<dataField\b[^>]*\bnumFmtId=)"\d+"', r'\1"3"', xml)
                data = xml.encode("utf-8")

            zout.writestr(item, data)

    return out

def _update_pivot_caches_from_current_raw(workbook_path, used):
    """Rebuild the six Pivot caches from current processed raw sheets without rewriting PivotTables."""
    from xml.etree import ElementTree as ET
    NS="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    def q(tag): return f"{{{NS}}}{tag}"
    cache_map={1:"R1",2:"R1S",3:"R2",4:"R6",5:"R1 OC",6:"R6 OC"}
    def norm(v):
        if v is None or (isinstance(v,float) and math.isnan(v)): return None
        if isinstance(v,str):
            v=v.strip(); return v if v else None
        return v
    def valkey(v):
        v=norm(v)
        if v is None:return ('blank','')
        if isinstance(v,(dt.datetime,dt.date)):return ('date',v.isoformat())
        if isinstance(v,bool):return ('num','1' if v else '0')
        if isinstance(v,(int,float)):return ('num',str(int(v)) if isinstance(v,float) and v.is_integer() else str(v))
        return ('str',str(v))
    def text(v):
        if isinstance(v,dt.datetime):return v.isoformat()
        if isinstance(v,dt.date):return dt.datetime.combine(v,dt.time()).isoformat()
        if isinstance(v,bool):return '1' if v else '0'
        if isinstance(v,float) and v.is_integer():return str(int(v))
        return str(v)

    tmp=workbook_path.with_suffix('.cachefix.xlsx')
    with zipfile.ZipFile(workbook_path,'r') as zin:
        infos=zin.infolist(); raw={i.filename:zin.read(i.filename) for i in infos}
    wb=load_workbook(workbook_path,read_only=True,data_only=False)
    new={}; maps={}
    for idx,sheet in cache_map.items():
        name=f'xl/pivotCache/pivotCacheDefinition{idx}.xml'
        if name not in raw: continue
        root=ET.fromstring(raw[name]); ws=wb[sheet]
        rows=list(ws.iter_rows(values_only=True)); headers=[str(v).strip() if v is not None else '' for v in rows[0]] if rows else []
        hmap={h:i for i,h in enumerate(headers) if h}
        fields=root.find(q('cacheFields')); fieldmaps=[]
        for cf in list(fields):
            fname=cf.attrib.get('name',''); ci=hmap.get(fname)
            vals=[r[ci] if ci is not None and ci<len(r) else None for r in rows[1:]]
            si=cf.find(q('sharedItems')); entries=[]
            if si is not None:
                for ch in list(si):
                    tag=ch.tag.rsplit('}',1)[-1]
                    if tag=='m': entries.append(None)
                    elif tag in ('s','n','d','e'): entries.append(({'s':'str','n':'num','d':'date','e':'err'}[tag],ch.attrib.get('v','')))
            mp={('blank',''):i for i,e in enumerate(entries) if e is None}
            for i,e in enumerate(entries):
                if e is not None: mp[e]=i
            rec=[]
            for v in vals:
                v=norm(v)
                if v is None:
                    if ('blank','') not in mp: mp[('blank','')]=len(entries); entries.append(None)
                    rec.append(mp[('blank','')]); continue
                k=valkey(v)
                if k in mp: rec.append(mp[k]); continue
                k2=(k[0],text(v))
                if k2 in mp: rec.append(mp[k2]); continue
                et=('date',text(v)) if isinstance(v,(dt.datetime,dt.date)) else (('num',text(v)) if isinstance(v,(int,float,bool)) else ('str',str(v)))
                mp[et]=len(entries); entries.append(et); rec.append(mp[et])
            if si is None: si=ET.SubElement(cf,q('sharedItems'))
            for ch in list(si): si.remove(ch)
            si.set('count',str(len(entries)))
            for a in ('containsBlank','containsString','containsNumber','containsDate','containsMixedTypes'): si.attrib.pop(a,None)
            types={e[0] for e in entries if e is not None}
            if any(e is None for e in entries):si.set('containsBlank','1')
            if 'str' in types:si.set('containsString','1')
            if 'num' in types:si.set('containsNumber','1')
            if 'date' in types:si.set('containsDate','1')
            if len(types)>1:si.set('containsMixedTypes','1')
            for e in entries:
                if e is None:ET.SubElement(si,q('m'))
                elif e[0]=='str':ET.SubElement(si,q('s'),{'v':e[1]})
                elif e[0]=='num':ET.SubElement(si,q('n'),{'v':e[1]})
                elif e[0]=='date':ET.SubElement(si,q('d'),{'v':e[1]})
                else:ET.SubElement(si,q('e'),{'v':e[1]})
            fieldmaps.append(rec)
        src=root.find(q('cacheSource')); wss=src.find(q('worksheetSource')) if src is not None else None
        if wss is not None:
            old=wss.get('ref','A1:A1'); last=re.match(r'[A-Z]+',old.split(':')[-1]).group(0); wss.set('ref',f'A1:{last}{used.get(sheet,0)+1}')
        root.set('recordCount',str(used.get(sheet,0))); root.set('refreshOnLoad','1'); root.set('enableRefresh','1')
        new[name]=ET.tostring(root,encoding='utf-8',xml_declaration=True); maps[idx]=fieldmaps
    wb.close()
    # Build records after ALL definitions are processed (ZIP ordering is not guaranteed).
    for idx,sheet in cache_map.items():
        name=f'xl/pivotCache/pivotCacheRecords{idx}.xml'; fieldmaps=maps.get(idx)
        if name not in raw or fieldmaps is None: continue
        n=used.get(sheet,0); root=ET.Element(q('pivotCacheRecords'),{'count':str(n)})
        for r in range(n):
            rr=ET.SubElement(root,q('r'))
            for fmap in fieldmaps:
                x=fmap[r] if r<len(fmap) else None
                if x is None:ET.SubElement(rr,q('m'))
                else:ET.SubElement(rr,q('x'),{'v':str(x)})
        new[name]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            data=new.get(info.filename,raw[info.filename])
            if re.fullmatch(r'xl/pivotTables/pivotTable[1-6]\.xml',info.filename):
                xml=data.decode('utf-8'); xml=re.sub(r'(<dataField\b[^>]*\bnumFmtId=)"\d+"',r'\1"3"',xml); data=xml.encode('utf-8')
            zout.writestr(info,data)
    workbook_path.unlink(); tmp.replace(workbook_path)


def _patch_native_cache_source_ranges(workbook_path, used):
    """Change only worksheetSource refs/refresh flags; leave native cache records untouched."""
    tmp=workbook_path.with_suffix('.srcpatch.xlsx')
    cache_map={1:"R1",2:"R1S",3:"R2",4:"R6",5:"R1 OC",6:"R6 OC"}
    with zipfile.ZipFile(workbook_path,'r') as zin, zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data=zin.read(info.filename)
            m=re.fullmatch(r'xl/pivotCache/pivotCacheDefinition([1-6])\.xml',info.filename)
            if m:
                idx=int(m.group(1)); sheet=cache_map[idx]
                xml=data.decode('utf-8')
                mm=re.search(r'<worksheetSource\b[^>]*ref="([^"]+)"[^>]*sheet="([^"]+)"',xml)
                if mm:
                    old_ref=mm.group(1)
                    last=re.match(r'[A-Z]+',old_ref.split(':')[-1]).group(0)
                    new_ref=f'A1:{last}{used.get(sheet,0)+1}'
                    xml=re.sub(r'(<worksheetSource\b[^>]*ref=")([^"]+)("[^>]*sheet=")',rf'\g<1>{new_ref}\g<3>',xml,count=1)
                if 'refreshOnLoad=' not in xml:
                    xml=xml.replace('<pivotCacheDefinition ','<pivotCacheDefinition refreshOnLoad="1" ',1)
                if 'enableRefresh=' not in xml:
                    xml=xml.replace('<pivotCacheDefinition ','<pivotCacheDefinition enableRefresh="1" ',1)
                data=xml.encode('utf-8')
            zout.writestr(info,data)
    workbook_path.unlink(); tmp.replace(workbook_path)


def _make_native_pivot_workbook(template_path, processed_path, used):
    """Start from the native template; replace only raw sheetData and cache source refs.
    PivotTable XML and PivotCache records remain byte-for-byte from the template."""
    out=processed_path.with_suffix('.native.xlsx')
    raw_sheets=["R1","R1S","R2","R6","R1 OC","R6 OC"]
    def sheet_paths(path):
        with zipfile.ZipFile(path,'r') as z:
            wbxml=z.read('xl/workbook.xml').decode('utf-8'); rel=z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
            rid={m.group(1):m.group(2) for m in re.finditer(r'<Relationship\b(?=[^>]*\bId="([^"]+)")(?=[^>]*\bTarget="([^"]+)")[^>]*/>',rel)}
            return {m.group(1): ('xl/'+rid[m.group(2)].lstrip('/').removeprefix('xl/')) for m in re.finditer(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"',wbxml)}
    tp=sheet_paths(template_path); pp=sheet_paths(processed_path)
    with zipfile.ZipFile(template_path,'r') as tin, zipfile.ZipFile(processed_path,'r') as pin, zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as zout:
        for info in tin.infolist():
            name=info.filename; data=tin.read(name)
            if name in [tp[s] for s in raw_sheets if s in tp]:
                sname=next(s for s in raw_sheets if tp.get(s)==name)
                old=data.decode('utf-8'); new=pin.read(pp[sname]).decode('utf-8')
                osd=re.search(r'<sheetData>.*?</sheetData>',old,re.S); nsd=re.search(r'<sheetData>.*?</sheetData>',new,re.S)
                header=re.search(r'<row r="1".*?</row>',osd.group(0),re.S).group(0)
                max_row=used.get(sname,0)+1
                rows=[]
                for rm in re.finditer(r'<row r="(\d+)".*?</row>',nsd.group(0),re.S):
                    if 2<=int(rm.group(1))<=max_row: rows.append(rm.group(0))
                old=re.sub(r'<sheetData>.*?</sheetData>','<sheetData>'+header+''.join(rows)+'</sheetData>',old,count=1,flags=re.S)
                old=re.sub(r'<dimension ref="[^"]+"',lambda m: re.sub(r':.*','',m.group(0))+f':{get_column_letter(0) if False else "A"}{max_row}',old,count=0) if False else old
                old=re.sub(r'<dimension ref="[^"]+"',lambda m: '<dimension ref="A1:'+re.match(r'[A-Z]+',re.search(r'<dimension ref="([^"]+)"',old).group(1).split(':')[-1]).group(0)+str(max_row)+'"',old,count=1)
                if '<autoFilter ' in old:
                    last=re.match(r'[A-Z]+',re.search(r'<dimension ref="([^"]+)"',old).group(1).split(':')[-1]).group(0)
                    old=re.sub(r'<autoFilter ref="[^"]+"',f'<autoFilter ref="A1:{last}{max_row}"',old,count=1)
                data=old.encode('utf-8')
            elif re.fullmatch(r'xl/pivotTables/pivotTable[1-6]\.xml',name):
                # Leave native PivotTable XML completely untouched.
                pass
            elif re.fullmatch(r'xl/pivotCache/pivotCache(Definition|Records)[1-6]\.xml',name):
                # Cache parts are kept from the template; only source ranges are patched below.
                pass
            zout.writestr(info,data)
    _patch_native_cache_source_ranges(out,used)
    return out

def build_workbook(raw_bytes, template_path):
    """Build from the original live-Pivot template while preserving source raw cells.

    Existing raw values are copied from the uploaded workbook directly, including
    Schedule Date values and their number formats.  Only the requested derived
    columns are calculated.  PivotTable XML itself is kept from the template.
    """
    raw_stream=io.BytesIO(raw_bytes)
    xls=pd.ExcelFile(raw_stream)
    out=Path(tempfile.mktemp(suffix=".xlsx"))
    shutil.copy2(template_path,out)
    wb=load_workbook(out)
    src_wb=load_workbook(io.BytesIO(raw_bytes),data_only=False)
    used={}
    aliases={"R1 OC":["R1 OC","R1OC"],"R6 OC":["R6 OC","R6OC"]}

    for logical in RAW_SHEETS:
        actual=logical
        if actual not in xls.sheet_names and logical in aliases:
            actual=next((a for a in aliases[logical] if a in xls.sheet_names),None)
        if actual not in xls.sheet_names or logical not in wb.sheetnames or actual not in src_wb.sheetnames:
            continue

        raw_stream.seek(0)
        source_df=pd.read_excel(raw_stream,sheet_name=actual)
        # Keep the same row-removal rule as the existing processing logic.
        keep_mask=~source_df.isna().all(axis=1)
        source_row_numbers=[i+2 for i,keep in enumerate(keep_mask.tolist()) if keep]
        df=process(source_df,is_r2=(logical=="R2"))

        task_status_cols=[c for c in df.columns if key(c)=="task status"]
        for _ts_col in task_status_cols:
            df[_ts_col]=df[_ts_col].map(norm)

        ws=wb[logical]
        src_ws=src_wb[actual]
        headers=[ws.cell(1,c).value for c in range(1,ws.max_column+1)]
        hmap={key(v):i+1 for i,v in enumerate(headers) if v is not None}
        src_headers=[src_ws.cell(1,c).value for c in range(1,src_ws.max_column+1)]
        src_hmap={key(v):i for i,v in enumerate(src_headers) if v is not None}

        # Clear old rows first; the uploaded Raw workbook remains the source of truth.
        if ws.max_row >= 2:
            for row in ws.iter_rows(min_row=2,max_row=ws.max_row,max_col=ws.max_column):
                for cell in row:
                    cell.value=None

        records=df.to_dict("records")
        for ri,(row,src_r) in enumerate(zip(records,source_row_numbers),start=2):
            for col,val in row.items():
                ci=hmap.get(key(col))
                if not ci:
                    continue
                cell=ws.cell(ri,ci)
                kcol=key(col)
                if kcol in src_hmap:
                    sc=src_ws.cell(src_r,src_hmap[kcol]+1)
                    # Preserve the original uploaded value exactly for existing fields.
                    cell.value=sc.value
                    # Preserve the original display format (critical for mixed/date Schedule Date).
                    cell.number_format=sc.number_format
                else:
                    cell.value=val
                if kcol=="task status":
                    cell.number_format="General"

        # Keep Task Status explicitly as General so status labels never become dates.
        ts_ci=hmap.get("task status")
        if ts_ci:
            for rr in range(2,len(records)+2):
                ws.cell(rr,ts_ci).number_format="General"
        ws.auto_filter.ref=f"A1:{get_column_letter(ws.max_column)}{max(2,len(records)+1)}"
        used[logical]=len(records)

    src_wb.close()
    wb.save(out); wb.close()

    # Keep the native PivotCache/PivotTable intact. Only update each cache source range
    # to the current raw row count; Excel can rebuild the cache on Refresh.
    _patch_native_cache_source_ranges(out, used)

    # Assemble from the untouched Excel Pivot template so all six PivotTables remain native/live.
    live_out=_make_native_pivot_workbook(template_path,out,used)
    data=live_out.read_bytes()
    try:
        out.unlink(); live_out.unlink()
    except Exception:
        pass
    return data,used

template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx")
if not template.exists():
    template=Path(__file__).with_name("Auto_Pivot_29-Sep_Final_LivePivot(1)(1).xlsx")
if not template.exists():
    template=Path(__file__).with_name("Auto_Pivot_Live_Template_v10.xlsx")
if not template.exists():
    template=Path(__file__).with_name("template.xlsx")
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
