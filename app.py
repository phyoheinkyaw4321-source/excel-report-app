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
.block-container {padding-top: 2rem; padding-bottom: 2rem;}
.main-title {font-size: 2rem; font-weight: 800; margin-bottom: .2rem;}
.sub {color:#667085; margin-bottom:1rem;}
.card {padding:1rem 1.1rem; border:1px solid #e5e7eb; border-radius:14px; background:#fff;}
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 Excel Report Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">Excel တစ်ဖိုင် Upload → RUN → Report အလိုအလျောက်တွက်မယ်</div>', unsafe_allow_html=True)


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


uploaded = st.file_uploader('📁 Excel ဖိုင်တစ်ဖိုင်ရွေးပါ (.xlsx)', type=['xlsx'], accept_multiple_files=False)

if uploaded:
    st.info(f'ရွေးထားတဲ့ဖိုင် — **{uploaded.name}**')
    if st.button('▶ RUN REPORT', type='primary', use_container_width=True):
        with st.spinner('Excel ကို process လုပ်နေပါတယ်...'):
            try:
                st.session_state['report_sheets'] = read_workbook(uploaded)
                st.session_state['source_name'] = uploaded.name
                st.session_state['output_xlsx'] = make_output_excel(st.session_state['report_sheets'])
                st.success('✅ Report ပြီးပါပြီ')
            except Exception as e:
                st.error(f'Process မအောင်မြင်ပါ: {e}')

sheets = st.session_state.get('report_sheets')
if sheets:
    st.divider()
    st.subheader('📌 Summary')
    tabs = st.tabs([x for x in SUMMARY_SHEETS if SUMMARY_SHEETS[x] in sheets])
    for tab, summary_name in zip(tabs, [x for x in SUMMARY_SHEETS if SUMMARY_SHEETS[x] in sheets]):
        raw = SUMMARY_SHEETS[summary_name]
        with tab:
            df = sheets[raw]
            st.caption(f'Raw sheet: {raw}  •  Rows: {len(df):,}')
            for col in ['Repay Status', 'Task Status', 'LAN Engineer Name', 'Car ID', 'RC']:
                if col in df.columns:
                    options = sorted([x for x in df[col].dropna().astype(str).unique() if x.strip()])
                    st.multiselect(col, options, key=f'{raw}_{col}')
            filtered = filtered_df(df, raw)
            stbl, total = summary_table(filtered)
            c1, c2, c3 = st.columns(3)
            c1.metric('Total Task No', f'{total:,}')
            c2.metric('Rows', f'{len(filtered):,}')
            c3.metric('Categories', f'{(stbl.iloc[:,0] != '(empty)').sum():,}')
            st.dataframe(stbl, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader('📄 Raw Data')
    raw_tabs = st.tabs([x for x in RAW_SHEETS if x in sheets])
    for tab, name in zip(raw_tabs, [x for x in RAW_SHEETS if x in sheets]):
        with tab:
            st.dataframe(sheets[name], use_container_width=True, height=500, hide_index=True)

    st.download_button('⬇️ Processed Excel Download', data=st.session_state['output_xlsx'], file_name=f"Processed_{st.session_state.get('source_name','report.xlsx')}", mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
