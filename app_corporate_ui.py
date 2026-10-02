import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Excel Report Dashboard', page_icon='📊', layout='wide')

RAW_SHEETS = ['R1', 'R1S', 'R2', 'R6', 'R1 OC', 'R6 OC']
SUMMARY_SHEETS = {
    'R1 V': 'R1', 'R1S V': 'R1S', 'R2 V': 'R2', 'R6 V': 'R6',
    'R1 OCV': 'R1 OC', 'R6 OCV': 'R6 OC'
}
CATEGORY_ORDER = ['FR-SLA / Biz TKT', 'FT-SBS', 'Install Express', 'Install POI', 'Other TKT']

st.markdown('''
<style>
html, body, [class*="css"] {font-size: 17px;}
.block-container {padding-top: 2rem; padding-bottom: 3rem; max-width: 1180px;}
.main-title {font-size: 3rem; line-height: 1.1; font-weight: 900; color:#123A63; margin-bottom:.35rem;}
.sub {font-size:1.25rem; color:#475569; margin-bottom:1rem;}
.upload-help {font-size:1.18rem; font-weight:600; color:#334155; margin:.35rem 0 .8rem;}
.section-card {border:1px solid #d7e0ea; border-radius:14px; padding:1.25rem 1.35rem; background:#fff; margin:1.1rem 0 1.4rem;}
.section-title {font-size:1.65rem; font-weight:900; color:#123A63; margin:1.25rem 0 .45rem;}
.sheet-title {border-top:1px solid #d7e0ea; padding-top:1.2rem; margin-top:1.6rem;}
.task-label {font-size:1.15rem; font-weight:850; color:#1e3a5f; letter-spacing:.04em; margin-bottom:.5rem;}
.generated {font-size:1.2rem; font-weight:800; color:#166534; margin-top:.35rem;}
.report-date {font-size:1.08rem; color:#334155; margin-top:.35rem;}
.selected-file {font-size:1.05rem; color:#334155; background:#f8fafc; border:1px solid #d7e0ea; border-radius:10px; padding:.7rem .9rem; margin:.5rem 0 1rem;}
.download-area {margin-top:2rem; padding-top:1rem; border-top:1px solid #d7e0ea;}
div[data-testid="stDataFrame"] {font-size:1.08rem;}
button[kind="primary"] {font-size:1.12rem !important; font-weight:800 !important;}
label[data-testid="stWidgetLabel"] p {font-size:1.05rem !important; font-weight:700 !important;}
@media (max-width: 800px) {
  .main-title {font-size:2.15rem;}
  .sub {font-size:1.05rem;}
  .section-title {font-size:1.35rem;}
  .block-container {padding-left:1rem; padding-right:1rem;}
}
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 DAILY REPORT AUTOMATION</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">Automated Excel Reporting System</div>', unsafe_allow_html=True)
st.markdown('<div style="text-align:center;font-size:1.25rem;font-weight:800;color:#123A63;margin:1rem 0 1.5rem;">Raw Excel &nbsp; → &nbsp; Process &nbsp; → &nbsp; Report</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📁 UPLOAD RAW EXCEL</div>', unsafe_allow_html=True)


def norm(v):
    if pd.isna(v):
        return ''
    return str(v).strip()


def find_col(df, name):
    if name in df.columns:
        return name
    low = {str(c).strip().lower(): c for c in df.columns}
    return low.get(name.lower())


def classify_tkt_poi(row):
    task = norm(row.get('Task No', ''))
    if re.search(r'\bTKT\b', task, flags=re.I):
        return 'TKT'
    if re.search(r'\bPOI\b', task, flags=re.I):
        return 'POI'
    return ''


def classify_extra(row):
    typ = norm(row.get('TKT / POI', '')).upper()
    service = norm(row.get('Service Area', ''))
    root = norm(row.get('Root Cause', ''))
    if typ == 'TKT':
        if ('FR-SLA' in service.upper()) or ('BIZ' in service.upper()) or ('INSTALLATION TYPE CHANGE' in root.upper()):
            return 'FR-SLA / Biz TKT'
        if 'FT-SBS' in root.upper():
            return 'FT-SBS'
        return 'Other TKT'
    if typ == 'POI':
        if 'EXPRESS' in root.upper():
            return 'Install Express'
        return 'Install POI'
    return ''


def classify_rc(df):
    car = find_col(df, 'Car ID')
    eng = find_col(df, 'LAN Engineer Name')
    if not car or not eng:
        return pd.Series([''] * len(df), index=df.index)
    temp = df[[car, eng]].copy()
    temp['_car'] = temp[car].map(norm)
    temp['_eng'] = temp[eng].map(norm)
    counts = temp[temp['_car'] != ''].groupby('_car')['_eng'].apply(lambda s: len({x for x in s if x}))
    def f(v):
        v = norm(v)
        if not v:
            return ''
        n = int(counts.get(v, 0))
        if n >= 3:
            return 'Car Share Plus'
        if n == 2:
            return 'Car Share'
        if n == 1:
            return 'One Car'
        return ''
    return temp['_car'].map(f)


def process_sheet(df):
    df = df.copy()
    # Keep original values, while ensuring required calculated columns exist.
    if 'Task No' not in df.columns:
        raise ValueError('Task No column မတွေ့ပါ')
    if 'TKT / POI' not in df.columns:
        df.insert(min(10, len(df.columns)), 'TKT / POI', '')
    if 'Extra' not in df.columns:
        df['Extra'] = ''
    df['TKT / POI'] = df.apply(classify_tkt_poi, axis=1)
    df['Extra'] = df.apply(classify_extra, axis=1)
    df['RC'] = classify_rc(df)
    return df


def read_workbook(uploaded):
    xls = pd.ExcelFile(uploaded)
    sheets = {}
    for raw in RAW_SHEETS:
        if raw in xls.sheet_names:
            sheets[raw] = process_sheet(pd.read_excel(uploaded, sheet_name=raw))
    if not sheets:
        # Accept common close spellings as a convenience.
        aliases = {s.replace(' ', '').lower(): s for s in xls.sheet_names}
        for raw in RAW_SHEETS:
            key = raw.replace(' ', '').lower()
            if key in aliases:
                actual = aliases[key]
                sheets[raw] = process_sheet(pd.read_excel(uploaded, sheet_name=actual))
    if not sheets:
        raise ValueError('R1 / R1S / R2 / R6 / R1 OC / R6 OC raw sheet မတွေ့ပါ')
    return sheets


def filtered_df(df, key):
    out = df.copy()
    cols = ['Repay Status', 'Task Status', 'LAN Engineer Name', 'Car ID', 'RC']
    for col in cols:
        if col in out.columns:
            vals = st.session_state.get(f'{key}_{col}', [])
            if vals:
                out = out[out[col].fillna('').astype(str).isin(vals)]
    return out


def summary_table(df):
    d = df.copy()
    task = d['Task No'].fillna('').astype(str).str.strip()
    d = d[task != '']
    counts = d['Extra'].fillna('').replace('', '(empty)').value_counts()
    rows = []
    for cat in CATEGORY_ORDER:
        rows.append((cat, int(counts.get(cat, 0))))
    # Preserve unexpected categories too.
    for cat, n in counts.items():
        if cat not in CATEGORY_ORDER and cat != '(empty)':
            rows.append((cat, int(n)))
    total = int(len(d))
    return pd.DataFrame(rows, columns=['Extra', 'Count - Task No']), total


def make_output_excel(sheets):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
        for summary, raw in SUMMARY_SHEETS.items():
            if raw not in sheets:
                continue
            stbl, total = summary_table(sheets[raw])
            s = pd.DataFrame([
                ['Filter', ''],
                ['Repay Status', '- all -'],
                ['Task Status', '- all -'],
                ['LAN Engineer Name', '- all -'],
                ['Car ID', '- all -'],
                ['RC', '- all -'],
                ['', ''],
                ['Extra', 'Count - Task No'],
            ], columns=['Filter', 'Value'])
            rows = pd.DataFrame(stbl.values, columns=['Filter', 'Value'])
            end = pd.DataFrame([['Total Result', total]], columns=['Filter', 'Value'])
            pd.concat([s, rows, end], ignore_index=True).to_excel(writer, sheet_name=summary[:31], index=False, header=False)
    bio.seek(0)
    return bio.getvalue()



def _reporting_date(sheets):
    for df in sheets.values():
        for col in ['Schedule Date', 'schedule date']:
            if col in df.columns:
                vals = pd.to_datetime(df[col], errors='coerce').dropna()
                if len(vals):
                    return vals.min().strftime('%d-%b-%Y')
    return ''


def _task_status_table(df):
    if 'Task Status' not in df.columns:
        return pd.DataFrame(columns=['Task Status', 'Count'])
    s = df['Task Status'].fillna('').astype(str).str.strip()
    s = s[s != '']
    counts = s.value_counts()
    rows = [{'Task Status': k, 'Count': int(v)} for k, v in counts.items()]
    rows.append({'Task Status': 'Total', 'Count': int(counts.sum())})
    return pd.DataFrame(rows)


uploaded = st.file_uploader('Choose Excel File', type=['xlsx'], accept_multiple_files=False, label_visibility='collapsed')

st.markdown('<div class="upload-help">Upload today\'s Raw Excel file</div>', unsafe_allow_html=True)

if uploaded:
    st.markdown(f'<div class="selected-file">📄 {uploaded.name}</div>', unsafe_allow_html=True)

if st.button('🚀 GENERATE REPORT', type='primary', use_container_width=True, disabled=(uploaded is None)):
    with st.spinner('Generating report...'):
        try:
            st.session_state['report_sheets'] = read_workbook(uploaded)
            st.session_state['source_name'] = uploaded.name
            st.session_state['output_xlsx'] = make_output_excel(st.session_state['report_sheets'])
            st.session_state['report_date'] = _reporting_date(st.session_state['report_sheets'])
            st.session_state['report_generated'] = True
        except Exception as e:
            st.error(f'Process မအောင်မြင်ပါ: {e}')

sheets = st.session_state.get('report_sheets')
if sheets:
    st.markdown('<div class="section-card report-status"><div class="section-title">📋 REPORT STATUS</div><div class="generated">✅ Report Generated</div><div class="report-date">Reporting Date: <b>' + (st.session_state.get('report_date') or '—') + '</b></div></div>', unsafe_allow_html=True)

    for name in RAW_SHEETS:
        if name not in sheets:
            continue
        st.markdown(f'<div class="section-title sheet-title">📊 {name}</div>', unsafe_allow_html=True)
        st.markdown('<div class="task-label">TASK STATUS</div>', unsafe_allow_html=True)
        table = _task_status_table(sheets[name])
        st.dataframe(table, use_container_width=True, hide_index=True, height=min(520, max(150, 72 * len(table) + 58)))

    st.markdown('<div class="download-area">', unsafe_allow_html=True)
    st.download_button('📥 DOWNLOAD FINAL REPORT', data=st.session_state['output_xlsx'], file_name=f"Processed_{st.session_state.get('source_name','report.xlsx')}", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
