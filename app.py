import io, re, shutil, zipfile, tempfile, subprocess, sys
from pathlib import Path
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='Auto Pivot Report', page_icon='📊', layout='wide')
RAW_SHEETS=['R1','R1S','R2','R6','R1 OC','R6 OC']
R2S_TOWNSHIPS={'kamaryut','dagon myothit north','sanchaung','kyeemyindaing','lanmadaw','pabedan','pazundaung','kyauktada','latha','botahtaung'}

def norm(v): return '' if pd.isna(v) else str(v).strip()
def key(v): return re.sub(r'\s+',' ',norm(v)).lower()
def task_type(v):
    t=norm(v).upper(); return 'TKT' if 'TKT' in t else ('POI' if 'POI' in t else '')
def classification(t,service,root):
    t,s,r=norm(t).upper(),norm(service).upper(),norm(root).upper()
    if t=='TKT':
        if 'FR-SLA' in s or 'BIZ' in s: return 'FR-SLA / Biz TKT'
        if 'INSTALLATION TYPE CHANGE' in r: return 'FR-SLA / Biz TKT'
        if 'FT-SBS' in r: return 'FT-SBS'
        return 'Other TKT'
    if t=='POI': return 'Install Express' if 'EXPRESS' in r else 'Install POI'
    return ''

def process(df,is_r2=False):
    df=df.dropna(how='all').reset_index(drop=True).copy()
    for c in ['TKT / POI','Classification','RC','RE']:
        if c in df.columns: df.drop(columns=c,inplace=True)
    service=next((c for c in df.columns if key(c)=='service area'),None)
    task=next((c for c in df.columns if key(c)=='task no'),None)
    root=next((c for c in df.columns if key(c)=='root cause'),None)
    if task:
        tt=df[task].map(task_type)
        cc=[classification(t,s,r) for t,s,r in zip(tt,df[service].map(norm) if service else ['']*len(df),df[root].map(norm) if root else ['']*len(df))]
        pos=df.columns.get_loc(service)+1 if service else 0
        df.insert(pos,'TKT / POI',tt); df.insert(pos+1,'Classification',cc)
    car=next((c for c in df.columns if key(c)=='car id'),None)
    eng=next((c for c in df.columns if key(c)=='lan engineer name'),None)
    repay=next((c for c in df.columns if key(c)=='repay status'),None)
    if car and eng:
        tmp=pd.DataFrame({'car':df[car].map(norm),'eng':df[eng].map(norm)})
        if repay: tmp=tmp[~df[repay].map(norm).str.lower().str.contains(r'engr\s*leave|route\s*cancel',regex=True,na=False)]
        tmp=tmp[(tmp.car!='')&(tmp.eng!='')]
        counts=tmp.groupby('car').eng.apply(lambda s:len(set(s))).to_dict()
        rc=[]
        for x in df[car].map(norm):
            n=counts.get(x,0)
            rc.append('' if not x else ('Car Share Plus' if n>=3 else 'Car Share' if n==2 else 'One Car' if n==1 else ''))
        df.insert(df.columns.get_loc(car),'RC',rc)
    if is_r2:
        township=next((c for c in df.columns if key(c)=='township'),None)
        if township:
            vals=[]
            for x in df[township]:
                t=key(x)
                if not t: vals.append(''); continue
                base=re.split(r'\s*\(|\s+ward\b',t,maxsplit=1)[0].strip()
                vals.append('R2S' if t in R2S_TOWNSHIPS or base in R2S_TOWNSHIPS else 'R2')
            df.insert(df.columns.get_loc(township)+1,'RE',vals)
    return df

def write_intermediate(raw_bytes,template,out_path):
    xls=pd.ExcelFile(io.BytesIO(raw_bytes))
    shutil.copy2(template,out_path)
    wb=load_workbook(out_path)
    used={}
    for s in RAW_SHEETS:
        if s not in xls.sheet_names or s not in wb.sheetnames: continue
        df=process(pd.read_excel(io.BytesIO(raw_bytes),sheet_name=s),s=='R2')
        used[s]={'rows':len(df),'cols':len(df.columns)}
        ws=wb[s]
        headers=[ws.cell(1,c).value for c in range(1,ws.max_column+1)]
        hmap={key(v):i+1 for i,v in enumerate(headers) if v is not None}
        for row in ws.iter_rows(min_row=2,max_row=ws.max_row,max_col=ws.max_column):
            for cell in row: cell.value=None
        for ri,row in enumerate(df.to_dict('records'),2):
            for col,val in row.items():
                ci=hmap.get(key(col))
                if ci:
                    cell=ws.cell(ri,ci); cell.value=None if pd.isna(val) else val
                    if key(col)=='task status': cell.number_format='General'
        ts=hmap.get('task status')
        if ts:
            for rr in range(2,len(df)+2): ws.cell(rr,ts).number_format='General'
        ws.auto_filter.ref=f'A1:{get_column_letter(ws.max_column)}{len(df)+1}'
    wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode='auto'
    wb.save(out_path); wb.close()
    return used

def patch_pivot_sources(path,used):
    # After LibreOffice has re-saved the workbook, patch the Excel PivotCache
    # source ranges to the exact processed data size and force refresh on open.
    cache_map={1:'R1',2:'R1S',3:'R2',4:'R6',5:'R1 OC',6:'R6 OC'}
    tmp=path.with_suffix('.patched.xlsx')
    with zipfile.ZipFile(path,'r') as zin, zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data=zin.read(item.filename)
            m=re.fullmatch(r'xl/pivotCache/pivotCacheDefinition([1-6])\.xml',item.filename)
            if m:
                idx=int(m.group(1)); s=cache_map[idx]; u=used.get(s)
                if u:
                    ref=f'A1:{get_column_letter(u["cols"])}{u["rows"]+1}'
                    xml=data.decode('utf-8')
                    xml=re.sub(r'<pivotCacheDefinition([^>]*)>',lambda mm: '<pivotCacheDefinition'+mm.group(1).replace(' refreshOnLoad="1"','').replace(' enableRefresh="1"','')+' refreshOnLoad="1" enableRefresh="1">',xml,count=1)
                    xml=re.sub(r'(<worksheetSource[^>]*\bref=")[^"]+("[^>]*\bsheet="'+re.escape(s)+r'"[^>]*/>)',r'\g<1>'+ref+r'\g<2>',xml,count=1)
                    xml=re.sub(r'(recordCount=")[^"]+("[^>]*>)',r'\g<1>'+str(u['rows'])+r'\g<2>',xml,count=1)
                    data=xml.encode('utf-8')
            elif re.fullmatch(r'xl/pivotTables/pivotTable[1-6]\.xml',item.filename):
                # Count fields must display as numbers, not dates.
                xml=data.decode('utf-8')
                xml=re.sub(r'(<dataField[^>]*\bname="Count - Task No"[^>]*\bnumFmtId=")[^"]+("[^>]*/>)',r'\g<1>164\g<2>',xml)
                data=xml.encode('utf-8')
            zout.writestr(item,data)
    path.unlink(); tmp.rename(path)

def make_report(raw_bytes,template):
    work=Path(tempfile.mktemp(suffix='.xlsx'))
    final=Path(tempfile.mktemp(suffix='.xlsx'))
    used=write_intermediate(raw_bytes,template,work)
    helper=Path(__file__).with_name('pivot_refresh.py')
    cmd=['/usr/bin/python3',str(helper),str(work),str(final)]
    if not Path('/usr/bin/python3').exists(): cmd=['python3',str(helper),str(work),str(final)]
    p=subprocess.run(cmd,capture_output=True,text=True,timeout=180)
    if p.returncode!=0:
        raise RuntimeError('Pivot refresh engine failed: '+(p.stderr[-1500:] or p.stdout[-1500:]))
    patch_pivot_sources(final,used)
    data=final.read_bytes()
    work.unlink(missing_ok=True); final.unlink(missing_ok=True)
    return data,used

template=Path(__file__).with_name('Auto_Pivot_29-Sep_Final_LivePivot(1).xlsx')
if not template.exists(): st.error('Live Pivot template မတွေ့ပါ။'); st.stop()
st.title('📊 Auto Pivot Report')
st.caption('Raw Excel Upload → True Excel PivotTable → Refresh / Filter / Field Selection')
uploaded=st.file_uploader('📁 Raw Excel (.xlsx)',type=['xlsx'])
if uploaded and st.button('▶ RUN REPORT',type='primary',use_container_width=True):
    with st.spinner('Raw data + True PivotTable ကို update လုပ်နေပါတယ်...'):
        try:
            data,used=make_report(uploaded.getvalue(),template)
            st.session_state['out']=data; st.session_state['used']=used
            st.success('✅ True PivotTable report ပြီးပါပြီ။')
        except Exception as e: st.error(str(e))
if 'out' in st.session_state:
    st.write('### Processed sheets'); st.write(st.session_state['used'])
    st.download_button('⬇️ Download Live Pivot Excel',data=st.session_state['out'],file_name='Auto_Pivot_Live.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
    st.info('Excel ဖွင့်ပြီး Pivot filter/dropdown, field selection, Refresh အားလုံးကို သုံးနိုင်ပါတယ်။')
