import io
import re
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Daily Report Automation', page_icon='📊', layout='wide')

import os

# ---------- Secure Corporate Login ----------
def _report_login():
    configured = st.secrets.get("APP_PASSWORD", os.getenv("APP_PASSWORD", ""))
    if not configured:
        st.warning("APP_PASSWORD မသတ်မှတ်ရသေးပါ။ Streamlit Secrets ထဲမှာ password ထည့်ပါ။")
        st.stop()
    if st.session_state.get("_report_logged_in"):
        return
    st.markdown("""
    <style>
      .stApp { background:#F4F7FB; }
      .block-container { max-width:1100px; padding-top:5rem; }
      .login-card { max-width:560px; margin:0 auto; background:#FFFFFF; border:1px solid #D9E2EC; border-radius:20px; padding:42px 46px 30px; box-shadow:0 14px 38px rgba(16,24,40,.10); text-align:center; }
      .login-icon { font-size:3rem; }
      .login-title { color:#0B2F52; font-size:1.8rem; font-weight:800; margin-top:8px; }
      .login-sub { color:#667085; margin-top:7px; line-height:1.6; }
      div.stButton > button[kind="primary"] { background:#123A63; border-color:#123A63; border-radius:10px; height:46px; font-weight:700; }
    </style>
    """, unsafe_allow_html=True)
    st.markdown("""<div class="login-card"><div class="login-icon">📊</div><div class="login-title">DAILY REPORT AUTOMATION</div><div class="login-sub"><b>Secure Report Access</b><br>Report အသုံးပြုရန် Password ထည့်ပါ</div></div>""", unsafe_allow_html=True)
    _, mid, _ = st.columns([1,2,1])
    with mid:
        password = st.text_input("🔐 Password", type="password", placeholder="Password ထည့်ပါ")
        if st.button("🔓  LOGIN TO REPORT", type="primary", use_container_width=True):
            if password == configured:
                st.session_state["_report_logged_in"] = True
                st.rerun()
            else:
                st.error("Password မှားနေပါတယ်။")
    st.stop()

_report_login()

st.markdown("""
<style>
  .stApp { background:#F4F7FB; color:#172B4D; }
  .block-container { max-width:1180px; padding:2rem 2rem 3rem; }
  .corporate-hero { background:linear-gradient(135deg,#0B2F52,#123A63); color:#fff; padding:28px 32px; border-radius:18px; box-shadow:0 8px 24px rgba(18,58,99,.14); margin-bottom:22px; }
  .corporate-hero h1 { margin:0; font-size:2rem; font-weight:800; }
  .corporate-hero p { margin:6px 0 0; opacity:.88; }
  .flow { margin-top:18px; display:flex; gap:10px; align-items:center; flex-wrap:wrap; }
  .flow span { padding:7px 13px; border-radius:999px; background:rgba(255,255,255,.12); border:1px solid rgba(255,255,255,.2); }
  div.stButton > button[kind="primary"] { background:#123A63; border-color:#123A63; border-radius:10px; font-weight:700; height:46px; }
  div.stDownloadButton > button { border-radius:10px; font-weight:700; height:46px; }
  [data-testid="stFileUploader"] { background:#fff; border:1px dashed #AFC1D3; border-radius:14px; padding:8px; }
</style>
""", unsafe_allow_html=True)
st.markdown("""<div class="corporate-hero"><h1>📊 DAILY REPORT AUTOMATION</h1><p>Automated Excel Reporting System</p><div class="flow"><span>📁 Raw Excel</span><b>→</b><span>⚙️ Process</span><b>→</b><span>📊 Pivot Report</span></div></div>""", unsafe_allow_html=True)

RAW_SHEETS = ['R1', 'R1S', 'R2', 'R6', 'R1 OC', 'R6 OC']
SUMMARY_SHEETS = {
    'R1 V': 'R1', 'R1S V': 'R1S', 'R2 V': 'R2', 'R6 V': 'R6',
    'R1 OCV': 'R1 OC', 'R6 OCV': 'R6 OC'
}
CATEGORY_ORDER = ['FR-SLA / Biz TKT', 'FT-SBS', 'Install Express', 'Install POI', 'Other TKT']

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
