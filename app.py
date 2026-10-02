import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Daily Report Automation', page_icon='📊', layout='wide')

RAW_SHEETS = ['R1', 'R1S', 'R2', 'R6', 'R1 OC', 'R6 OC']
SUMMARY_SHEETS = {
    'R1 V': 'R1', 'R1S V': 'R1S', 'R2 V': 'R2', 'R6 V': 'R6',
    'R1 OCV': 'R1 OC', 'R6 OCV': 'R6 OC'
}
CATEGORY_ORDER = ['FR-SLA / Biz TKT', 'FT-SBS', 'Install Express', 'Install POI', 'Other TKT']

# Corporate white + navy UI. Processing logic below is unchanged.
st.markdown('''
<style>
    :root { --navy:#123B6D; --navy2:#0B2E59; --line:#D9E2EC; --muted:#52606D; }
    .stApp { background:#FFFFFF; }
    .block-container { max-width:1200px; padding-top:2.0rem; padding-bottom:3rem; }
    .hero { border-bottom:3px solid var(--navy); padding:0.3rem 0 1.25rem 0; margin-bottom:1.5rem; }
    .main-title { color:var(--navy); font-size:2.45rem; line-height:1.15; font-weight:900; letter-spacing:-0.02em; margin:0; }
    .sub-title { color:#334E68; font-size:1.18rem; font-weight:600; margin-top:.45rem; }
    .flow { text-align:center; color:var(--navy); font-size:1.05rem; font-weight:800; margin-top:1rem; letter-spacing:.02em; }
    .section { color:var(--navy); font-size:1.42rem; font-weight:900; border-left:6px solid var(--navy); padding:.35rem .8rem; margin:1.2rem 0 .55rem 0; }
    .section-small { color:var(--navy); font-size:1.15rem; font-weight:900; margin:.25rem 0 .55rem 0; }
    .report-card { border:1px solid var(--line); border-radius:12px; padding:1rem 1rem 1.15rem 1rem; margin:1rem 0 1.2rem 0; background:#fff; box-shadow:0 2px 8px rgba(18,59,109,.06); }
    .status-head { color:var(--navy); font-size:1.08rem; font-weight:900; letter-spacing:.03em; margin-bottom:.45rem; }
    .report-date { color:#334E68; font-size:1.05rem; font-weight:700; }
    .success-box { border:1px solid #B7D8C0; background:#F4FBF6; border-radius:10px; padding:.8rem 1rem; color:#1F5E35; font-size:1.05rem; font-weight:800; }
    .hint { color:var(--muted); font-size:1rem; font-weight:600; }
    div[data-testid="stFileUploader"] { border:1px solid var(--line); border-radius:12px; padding:.6rem; background:#FAFCFF; }
    div[data-testid="stFileUploader"] label { font-size:1.05rem !important; font-weight:800 !important; color:#243B53 !important; }
    div.stButton > button { min-height:3.15rem; font-size:1.08rem; font-weight:900; border-radius:9px; }
    div.stDownloadButton > button { min-height:3.15rem; font-size:1.08rem; font-weight:900; border-radius:9px; }
    div[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:9px; overflow:hidden; }
    .stMetric label { font-size:1rem !important; font-weight:800 !important; }
    .stMetric [data-testid="stMetricValue"] { font-size:1.7rem !important; font-weight:900 !important; color:var(--navy) !important; }
    @media (max-width: 700px) {
        .block-container { padding:1rem .8rem 2rem .8rem; }
        .main-title { font-size:1.9rem; }
        .sub-title { font-size:1rem; }
        .flow { font-size:.9rem; }
        .section { font-size:1.2rem; }
    }
</style>
''', unsafe_allow_html=True)

st.markdown('''
<div class="hero">
  <div class="main-title">📊 DAILY REPORT AUTOMATION</div>
  <div class="sub-title">Automated Excel Reporting System</div>
  <div class="flow">Raw Excel&nbsp;&nbsp; → &nbsp;&nbsp;Process&nbsp;&nbsp; → &nbsp;&nbsp;Report</div>
</div>
''', unsafe_allow_html=True)


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


def task_status_table(df):
    if 'Task Status' not in df.columns:
        return pd.DataFrame(columns=['Task Status', 'Count']), 0
    vals = df['Task Status'].fillna('').astype(str).str.strip()
    # Show every exact non-blank Task Status separately; blanks are kept as a separate result when present.
    counts = vals.value_counts(dropna=False)
    rows = []
    for status, count in counts.items():
        label = status if status else '(Blank)'
        rows.append((label, int(count)))
    out = pd.DataFrame(rows, columns=['Task Status', 'Count'])
    return out, int(len(df))


def reporting_date(sheets):
    dates = []
    for df in sheets.values():
        col = find_col(df, 'Schedule Date')
        if col:
            parsed = pd.to_datetime(df[col], errors='coerce', dayfirst=True)
            dates.extend(parsed.dropna().tolist())
    if not dates:
        return '-'
    d = max(dates)
    return pd.Timestamp(d).strftime('%d-%b-%Y')


st.markdown('<div class="section">📁 UPLOAD RAW EXCEL</div>', unsafe_allow_html=True)
st.markdown('<div class="hint">Upload today\'s Raw Excel file</div>', unsafe_allow_html=True)
uploaded = st.file_uploader('Choose Excel File', type=['xlsx'], accept_multiple_files=False, label_visibility='visible')

if uploaded:
    st.info(f'📄 Selected file: **{uploaded.name}**')

if st.button('🚀 GENERATE REPORT', type='primary', use_container_width=True):
    if not uploaded:
        st.warning('ကျေးဇူးပြု၍ Raw Excel ဖိုင်ကို အရင်ရွေးပါ။')
    else:
        with st.spinner('Excel ကို process လုပ်နေပါတယ်...'):
            try:
                st.session_state['report_sheets'] = read_workbook(uploaded)
                st.session_state['source_name'] = uploaded.name
                st.session_state['output_xlsx'] = make_output_excel(st.session_state['report_sheets'])
                st.session_state['generated'] = True
            except Exception as e:
                st.error(f'Process မအောင်မြင်ပါ: {e}')

sheets = st.session_state.get('report_sheets')
if sheets:
    st.markdown('<div class="section">📋 REPORT STATUS</div>', unsafe_allow_html=True)
    st.markdown('<div class="success-box">✅ Report Generated</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="report-date">Reporting Date: {reporting_date(sheets)}</div>', unsafe_allow_html=True)

    for name in RAW_SHEETS:
        if name not in sheets:
            continue
        st.markdown(f'<div class="report-card"><div class="section-small">📊 {name}</div><div class="status-head">TASK STATUS</div></div>', unsafe_allow_html=True)
        table, total = task_status_table(sheets[name])
        st.dataframe(
            table,
            use_container_width=True,
            hide_index=True,
            column_config={
                'Task Status': st.column_config.TextColumn('Task Status', width='large'),
                'Count': st.column_config.NumberColumn('Count', format='%d', width='small'),
            },
            height=min(480, max(110, 58 + len(table) * 46)),
        )
        st.markdown(f'<div class="status-head" style="text-align:right; margin-top:-.15rem;">Total&nbsp;&nbsp; {total:,}</div>', unsafe_allow_html=True)

    st.markdown('<div style="height:.8rem"></div>', unsafe_allow_html=True)
    st.download_button(
        '📥 DOWNLOAD FINAL REPORT',
        data=st.session_state['output_xlsx'],
        file_name=f"Processed_{st.session_state.get('source_name','report.xlsx')}",
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
