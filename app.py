
import io
import re
from collections import Counter
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Excel Report Dashboard", page_icon="📊", layout="wide")

RAW_SHEETS = ["R1", "R1S", "R2", "R6", "R1 OC", "R6 OC"]
SHEET_ALIASES = {
    "R1": ["R1"],
    "R1S": ["R1S"],
    "R2": ["R2", "R2 "],
    "R6": ["R6"],
    "R1 OC": ["R1 OC", "R1OC", "R1 OC "],
    "R6 OC": ["R6 OC", "R6OC", "R6 ocv", "R6 OCV"],
}
SUMMARY_NAMES = ["R1 V", "R1S V", "R2 V", "R6 V", "R1 OCV", "R6 OCV"]
CATEGORY_ORDER = ["FR-SLA / Biz TKT", "FT-SBS", "Other TKT", "Install Express", "Install POI"]
R2S_TOWNSHIPS = {
    "kamaryut", "dagon myothit north", "sanchaung", "kyeemyindaing",
    "lanmadaw", "pabedan", "pazundaung", "kyauktada", "latha", "botahtaung"
}

st.markdown("""
<style>
.block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
h1 {margin-bottom: .2rem;}
.small {color:#667085;}
</style>
""", unsafe_allow_html=True)

st.title("📊 Excel Report Dashboard")
st.caption("Excel တစ်ဖိုင် → RUN REPORT → R1/R1S/R2/R6/R1 OC/R6 OC ကို အလိုအလျောက်တွက်မယ်")

# --- Access protection (keeps existing Streamlit Secrets password setup) ---
def _get_app_password():
    for k in ("PSW", "PASSWORD", "APP_PASSWORD", "REPORT_PASSWORD"):
        try:
            v = st.secrets.get(k)
            if v:
                return str(v)
        except Exception:
            pass
    return ""

def _login():
    expected = _get_app_password()
    if not expected:
        return True
    if st.session_state.get("_auth_ok"):
        return True
    st.markdown("### 🔐 Report Login")
    pw = st.text_input("Password", type="password", key="_login_password")
    if st.button("Login", type="primary", use_container_width=True):
        if pw == expected:
            st.session_state["_auth_ok"] = True
            st.rerun()
        else:
            st.error("Password မမှန်ပါ")
    return False

if not _login():
    st.stop()


def norm(v):
    if pd.isna(v):
        return ""
    return str(v).strip()

def key(v):
    return re.sub(r"\s+", " ", norm(v)).strip().lower()

def clean_headers(df):
    df = df.copy()
    seen = {}
    cols = []
    for c in df.columns:
        base = norm(c) or "Unnamed"
        n = seen.get(base, 0)
        seen[base] = n + 1
        cols.append(base if n == 0 else f"{base}.{n}")
    df.columns = cols
    return df

def find_col(df, name):
    target = key(name)
    for c in df.columns:
        if key(c) == target:
            return c
    return None

def task_type(task):
    t = norm(task).upper()
    if "TKT" in t:
        return "TKT"
    if "POI" in t:
        return "POI"
    return ""

def classification(task_type_value, service_area, root_cause):
    typ = norm(task_type_value).upper()
    service = norm(service_area).upper()
    root = norm(root_cause).upper()
    if typ == "TKT":
        if "FR-SLA" in service or "BIZ" in service:
            return "FR-SLA / Biz TKT"
        if "INSTALLATION TYPE CHANGE" in root:
            return "FR-SLA / Biz TKT"
        if "FT-SBS" in root:
            return "FT-SBS"
        return "Other TKT"
    if typ == "POI":
        if "EXPRESS" in root:
            return "Install Express"
        return "Install POI"
    return ""

def eligible_engineers_by_car(df):
    car_col = find_col(df, "Car ID")
    eng_col = find_col(df, "LAN Engineer Name")
    repay_col = find_col(df, "Repay Status")
    if not car_col or not eng_col:
        return {}

    temp = pd.DataFrame({
        "_car": df[car_col].map(norm),
        "_eng": df[eng_col].map(norm),
    })
    if repay_col:
        temp["_repay"] = df[repay_col].map(norm).str.lower()
        # Engr Leave / route cancel rows do NOT contribute an engineer to the Car Share count.
        excluded = temp["_repay"].str.contains(r"engr\s*leave|route\s*cancel", regex=True, na=False)
        temp = temp[~excluded]

    temp = temp[(temp["_car"] != "") & (temp["_eng"] != "")]
    return temp.groupby("_car")["_eng"].apply(lambda s: len(set(s))).to_dict()

def add_calculated_columns(df, is_r2=False):
    df = clean_headers(df)

    task_col = find_col(df, "Task No")
    service_col = find_col(df, "Service Area")
    root_col = find_col(df, "Root Cause")
    car_col = find_col(df, "Car ID")
    eng_col = find_col(df, "LAN Engineer Name")

    if not task_col:
        raise ValueError("Task No column မတွေ့ပါ")

    # TKT/POI and Classification are always calculated columns.
    task_values = df[task_col].map(task_type)
    service_values = df[service_col].map(norm) if service_col else pd.Series([""] * len(df), index=df.index)
    root_values = df[root_col].map(norm) if root_col else pd.Series([""] * len(df), index=df.index)
    class_values = [
        classification(t, s, r) for t, s, r in zip(task_values, service_values, root_values)
    ]

    # Replace existing calculated columns instead of duplicating them.
    for col in ["TKT / POI", "Classification", "RC"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Insert TKT / POI + Classification immediately after Service Area.
    insert_at = df.columns.get_loc(service_col) + 1 if service_col else min(10, len(df.columns))
    df.insert(insert_at, "TKT / POI", task_values)
    df.insert(insert_at + 1, "Classification", class_values)

    # RC is only added when the source has the columns needed to calculate it.
    if car_col and eng_col:
        counts = eligible_engineers_by_car(df)
        car_values = df[car_col].map(norm)
        rc_values = []
        for car in car_values:
            if not car:
                rc_values.append("")
            else:
                n = int(counts.get(car, 0))
                rc_values.append("Car Share Plus" if n >= 3 else "Car Share" if n == 2 else "One Car" if n == 1 else "")
        # Re-find Car ID after inserted columns.
        car_pos = df.columns.get_loc(car_col)
        df.insert(car_pos, "RC", rc_values)

    # R2 only: RE = R2S for the requested townships; all other nonblank townships = R2.
    if is_r2:
        township_col = find_col(df, "Township")
        if township_col:
            re_values = []
            for v in df[township_col]:
                t = key(v)
                re_values.append("" if not t else ("R2S" if t in R2S_TOWNSHIPS else "R2"))
            if "RE" in df.columns:
                df["RE"] = re_values
            else:
                # Put RE beside Township.
                pos = df.columns.get_loc(township_col) + 1
                df.insert(pos, "RE", re_values)

    return df

def resolve_sheet(xls, logical_name):
    actual = {key(s): s for s in xls.sheet_names}
    for alias in SHEET_ALIASES[logical_name]:
        if key(alias) in actual:
            return actual[key(alias)]
    return None

def read_and_process(uploaded):
    uploaded.seek(0)
    xls = pd.ExcelFile(uploaded)
    sheets = {}
    for logical in RAW_SHEETS:
        actual = resolve_sheet(xls, logical)
        if actual is None:
            continue
        uploaded.seek(0)
        df = pd.read_excel(uploaded, sheet_name=actual)
        # Drop completely empty rows only; preserve Task Status values exactly otherwise.
        df = df.dropna(how="all").reset_index(drop=True)
        sheets[logical] = add_calculated_columns(df, is_r2=(logical == "R2"))
    if not sheets:
        raise ValueError("R1 / R1S / R2 / R6 / R1 OC / R6 OC sheet မတွေ့ပါ")
    return sheets

def filtered(df, selections):
    out = df.copy()
    for col, vals in selections.items():
        if vals and col in out.columns:
            out = out[out[col].fillna("").astype(str).isin(vals)]
    return out

def metrics(df):
    task_col = find_col(df, "Task No")
    eng_col = find_col(df, "LAN Engineer Name")
    car_col = find_col(df, "Car ID")
    labour_col = find_col(df, "Contractor / Office Labour ID")
    task_mask = df[task_col].map(norm).ne("") if task_col else pd.Series([True]*len(df), index=df.index)
    d = df[task_mask]
    engineers = d[eng_col].map(norm)
    cars = d[car_col].map(norm) if car_col else pd.Series([], dtype=str)
    labour = d[labour_col].map(norm) if labour_col else pd.Series([], dtype=str)
    car_counts = cars[cars != ""].value_counts() if len(cars) else pd.Series(dtype=int)
    return {
        "tasks": len(d),
        "engineers": len(set(x for x in engineers if x)),
        "cars": len(set(x for x in cars if x)),
        "labour": len(set(x for x in labour if x)),
        "duplicate_cars": int((car_counts > 1).sum()),
    }

def make_download(sheets):
    # Download is a normal Excel workbook with AutoFilter + freeze panes.
    # The web dashboard remains the interactive pivot/filter view.
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.book[name[:31]]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

        for summary_name, raw_name in zip(SUMMARY_NAMES, RAW_SHEETS):
            if raw_name not in sheets:
                continue
            df = sheets[raw_name]
            m = metrics(df)
            task_col = find_col(df, "Task No")
            d = df[df[task_col].map(norm).ne("")] if task_col else df
            cat_counts = d["Classification"].fillna("").replace("", "(blank)").value_counts() if "Classification" in d else pd.Series(dtype=int)
            rows = [
                ["Metric", "Value"],
                ["Task No", m["tasks"]],
                ["LAN Engineer Name (distinct)", m["engineers"]],
                ["Car ID (distinct)", m["cars"]],
                ["Contractor / Office Labour ID (distinct)", m["labour"]],
                ["Duplicate Car ID", m["duplicate_cars"]],
                ["", ""],
                ["Classification", "Count - Task No"],
            ]
            for cat in CATEGORY_ORDER:
                rows.append([cat, int(cat_counts.get(cat, 0))])
            pd.DataFrame(rows).to_excel(writer, sheet_name=summary_name[:31], index=False, header=False)
            ws = writer.book[summary_name[:31]]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    bio.seek(0)
    return bio.getvalue()

uploaded = st.file_uploader("📁 Excel ဖိုင်တစ်ဖိုင်ရွေးပါ (.xlsx)", type=["xlsx"], accept_multiple_files=False)

if uploaded and st.button("▶ RUN REPORT", type="primary", use_container_width=True):
    with st.spinner("Excel ကို process လုပ်နေပါတယ်..."):
        try:
            sheets = read_and_process(uploaded)
            st.session_state["sheets"] = sheets
            st.session_state["source_name"] = uploaded.name
            st.session_state["download"] = make_download(sheets)
            st.success("✅ Report ပြီးပါပြီ")
        except Exception as e:
            st.error(f"Process မအောင်မြင်ပါ: {e}")

sheets = st.session_state.get("sheets")
if sheets:
    st.divider()
    st.subheader("📌 Summary / Pivot View")

    tab_names = [x for x, raw in zip(SUMMARY_NAMES, RAW_SHEETS) if raw in sheets]
    tabs = st.tabs(tab_names)

    for tab, summary_name in zip(tabs, tab_names):
        raw = SUMMARY_NAMES.index(summary_name)
        raw_name = RAW_SHEETS[raw]
        df = sheets[raw_name]

        with tab:
            available_filters = [
                c for c in ["Repay Status", "Task Status", "LAN Engineer Name", "Car ID", "RC", "TKT / POI", "Classification", "RE", "Township"]
                if c in df.columns
            ]
            selections = {}
            cols = st.columns(3)
            for i, col in enumerate(available_filters):
                opts = sorted({norm(x) for x in df[col] if norm(x)})
                if opts:
                    with cols[i % 3]:
                        selections[col] = st.multiselect(col, opts, key=f"{raw_name}_{col}")

            view = filtered(df, selections)
            m = metrics(view)
            c1,c2,c3,c4,c5 = st.columns(5)
            c1.metric("Task No", f"{m['tasks']:,}")
            c2.metric("LAN Engineers", f"{m['engineers']:,}")
            c3.metric("Car ID", f"{m['cars']:,}")
            c4.metric("Labour ID", f"{m['labour']:,}")
            c5.metric("Duplicate Car ID", f"{m['duplicate_cars']:,}")

            if "Classification" in view.columns:
                t = view[view[find_col(view,"Task No")].map(norm).ne("")]
                pivot = pd.crosstab(
                    t["Classification"].replace("", "(blank)"),
                    columns="Count - Task No"
                ).reindex(CATEGORY_ORDER + ["(blank)"], fill_value=0)
                pivot.index.name = "Classification"
                st.dataframe(pivot, use_container_width=True)

            # R2-specific RE view
            if raw_name == "R2" and "RE" in view.columns:
                st.markdown("**R2 / R2S (Township based)**")
                re_pivot = view[view[find_col(view,"Task No")].map(norm).ne("")]
                re_pivot = re_pivot.groupby(["RE"], dropna=False).size().reset_index(name="Count - Task No")
                st.dataframe(re_pivot, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("📄 Processed Raw Data")
    raw_tabs = st.tabs([x for x in RAW_SHEETS if x in sheets])
    for tab, name in zip(raw_tabs, [x for x in RAW_SHEETS if x in sheets]):
        with tab:
            st.dataframe(sheets[name], use_container_width=True, height=520, hide_index=True)

    st.download_button(
        "⬇️ Processed Excel Download",
        data=st.session_state["download"],
        file_name=f"Processed_{st.session_state.get('source_name','report.xlsx')}",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
