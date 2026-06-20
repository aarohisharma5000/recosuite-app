# ============================================================
# RecoSuite — Clean Customer App
# Simple, fast reconciliation tool for paying customers
# Auth gate + 100 row limit for free/demo users
# ============================================================

import io
import time
import re
import pandas as pd
import streamlit as st
from decimal import Decimal, InvalidOperation
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment
from pathlib import Path
import zipfile
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="RecoSuite — Reconciliation Tool",
    page_icon="📌",
    layout="wide"
)

# ============================================================
# AUTH GATE — must pass before anything loads
# ============================================================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds  = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
gc     = gspread.authorize(creds)
sheet  = gc.open_by_key(st.secrets["SHEET_ID"]).sheet1
rows   = sheet.get_all_records()

user_email = st.query_params.get("email", "")

if not user_email:
    st.markdown("""
    <div style="text-align:center; padding: 4rem 2rem;">
        <div style="font-size:3rem; margin-bottom:1rem;">📌</div>
        <h1 style="font-size:2rem; font-weight:900; color:#0a1628;">RecoSuite</h1>
        <p style="color:#666; font-size:1.1rem; margin-bottom:2rem;">
            Smart Reconciliation Tool — Access Restricted
        </p>
        <div style="background:#fff3cd; border:1px solid #ffc107; border-radius:12px;
                    padding:1.5rem; max-width:480px; margin:0 auto 2rem;">
            <p style="margin:0; color:#856404; font-size:1rem;">
                ❌ Please use your personal subscription link to access this tool.<br><br>
                📧 Contact <strong>aarohisharma5000@gmail.com</strong> to get your link.
            </p>
        </div>
        <a href="https://aarohisharma5000.github.io/recosuite"
           style="background:#00d4ff; color:#0a1628; padding:0.8rem 2rem;
                  border-radius:50px; font-weight:800; text-decoration:none; font-size:1rem;">
            View Plans & Subscribe →
        </a>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

user_row = next(
    (r for r in rows if str(r.get("email", "")).strip().lower() == user_email.strip().lower()),
    None
)

if not user_row:
    st.error(f"❌ **{user_email}** is not subscribed.")
    st.info("Subscribe at [aarohisharma5000.github.io/recosuite](https://aarohisharma5000.github.io/recosuite)")
    st.stop()

expiry = date.fromisoformat(str(user_row["expiry_date"]))
if expiry < date.today():
    st.error(f"⏰ Your subscription expired on **{expiry}**. Please renew.")
    st.info("Renew at [aarohisharma5000.github.io/recosuite](https://aarohisharma5000.github.io/recosuite)")
    st.stop()

# Auth passed — set plan
plan      = str(user_row.get("plan", "free")).strip().lower()
ROW_LIMIT = {"free": 100, "starter": 250000, "pro": 9999999}.get(plan, 100)
st.session_state["user_plan"]  = plan
st.session_state["user_email"] = user_email
st.session_state["row_limit"]  = ROW_LIMIT

# ============================================================
# CLEAN HEADER
# ============================================================
st.markdown(f"""
<style>
    .block-container {{ padding-top: 1.5rem; padding-bottom: 4rem; }}
    header[data-testid="stHeader"] {{ display:none; }}
    .reco-header {{
        display: flex; align-items: center; justify-content: space-between;
        background: linear-gradient(135deg, #0a1628, #0f2040);
        border-radius: 16px; padding: 1.2rem 2rem; margin-bottom: 1.5rem;
        border: 1px solid rgba(0,212,255,0.2);
    }}
    .reco-logo {{ font-size: 1.5rem; font-weight: 900; color: #00d4ff; }}
    .reco-logo span {{ color: white; }}
    .reco-plan {{
        background: rgba(0,212,255,0.15); border: 1px solid rgba(0,212,255,0.3);
        color: #00d4ff; padding: 0.3rem 1rem; border-radius: 50px;
        font-size: 0.85rem; font-weight: 700;
    }}
    .reco-user {{ color: rgba(255,255,255,0.6); font-size: 0.82rem; margin-top: 0.2rem; }}
    .upgrade-banner {{
        background: linear-gradient(135deg, #fff3cd, #fff8e1);
        border: 1px solid #ffc107; border-radius: 12px;
        padding: 1rem 1.5rem; margin: 1rem 0;
        display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 1rem;
    }}
    .upgrade-text {{ color: #856404; font-size: 0.95rem; }}
    .upgrade-btn {{
        background: #ffc107; color: #1a1a00; padding: 0.5rem 1.5rem;
        border-radius: 50px; font-weight: 800; text-decoration: none; font-size: 0.9rem;
    }}
</style>
<div class="reco-header">
    <div>
        <div class="reco-logo">Reco<span>Suite</span></div>
        <div class="reco-user">👤 {user_email}</div>
    </div>
    <div class="reco-plan">
        {'🆓 Free Plan — 100 rows' if plan == 'free' else ('⭐ Starter Plan' if plan == 'starter' else '🏆 Pro Plan')}
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# HELPERS (from app.py — kept identical)
# ============================================================
def normalize_cols(df):
    if df is None or df.empty:
        return df
    df = df.copy()
    def _norm(c):
        c = str(c)
        c = c.replace("\u00A0", " ").replace("\u200B", "")
        c = c.replace("\n", " ").replace("\r", " ").strip()
        c = re.sub(r"\s+", " ", c)
        return c
    df.columns = [_norm(c) for c in df.columns]
    seen = {}
    new_cols = []
    for c in df.columns:
        if c not in seen:
            seen[c] = 1
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}__{seen[c]}")
    df.columns = new_cols
    return df

def _dedup_columns(df):
    seen = {}
    new_cols = []
    for c in df.columns:
        if c not in seen:
            seen[c] = 1
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}__{seen[c]}")
    df.columns = new_cols
    return df

def _fix_mixed_types(df):
    for col in df.columns:
        try:
            import pyarrow as pa
            pa.array(df[col].values)
        except Exception:
            df[col] = df[col].astype(str)
    return df

def read_file_once(uploaded_file):
    name = uploaded_file.name.lower()
    b    = uploaded_file.getvalue()
    bio  = io.BytesIO(b)
    if name.endswith(".csv"):
        df = pd.read_csv(bio)
        df = _dedup_columns(df)
        df = _fix_mixed_types(df)
        df["_SHEET"] = "CSV"
        return df
    elif name.endswith((".xlsx", ".xls")):
        sheets = pd.read_excel(bio, sheet_name=None)
        best_df, best_rows, best_sheet = None, -1, None
        for sh, dfx in sheets.items():
            if dfx is not None and len(dfx) > best_rows:
                best_rows = len(dfx); best_df = dfx; best_sheet = sh
        best_df = best_df if best_df is not None else pd.DataFrame()
        best_df = _dedup_columns(best_df)
        best_df = _fix_mixed_types(best_df)
        best_df["_SHEET"] = str(best_sheet) if best_sheet else "UNKNOWN"
        return best_df
    elif name.endswith(".xlsb"):
        sheets = pd.read_excel(bio, engine="pyxlsb", sheet_name=None)
        best_df, best_rows, best_sheet = None, -1, None
        for sh, dfx in sheets.items():
            if dfx is not None and len(dfx) > best_rows:
                best_rows = len(dfx); best_df = dfx; best_sheet = sh
        best_df = best_df if best_df is not None else pd.DataFrame()
        best_df = _dedup_columns(best_df)
        best_df = _fix_mixed_types(best_df)
        best_df["_SHEET"] = str(best_sheet) if best_sheet else "UNKNOWN"
        return best_df
    else:
        raise ValueError("Unsupported file type. Upload CSV / XLSX / XLSB.")

def is_blank(x):
    if pd.isna(x): return True
    if isinstance(x, str):
        t = x.strip().lower()
        if t in {"", "-", "--", "na", "n/a", "null", "none"}: return True
    return False

SCI_RE = re.compile(r"^\s*-?\d+(\.\d+)?[eE][+-]?\d+\s*$")

def clean_key_series(s, treat_as_text=True):
    def _one(v):
        if pd.isna(v): return ""
        t = str(v).strip()
        if not t: return ""
        t = t.replace("\u00A0","").replace("\u200B","").strip()
        if t.startswith("="): t = t[1:].strip()
        t = t.strip('"').strip("'").strip()
        t = re.sub(r"[\s,]+","",t)
        if t in {"0","0.0","0.00"}: return ""
        t = re.sub(r"\.0+$","",t)
        if SCI_RE.match(t):
            try:
                d = Decimal(t)
                t = format(d.quantize(Decimal(1)),"f")
            except: pass
        if treat_as_text:
            return re.sub(r"[^0-9A-Za-z]+","",t).upper()
        return re.sub(r"\D+","",t)
    return s.apply(_one)

KEY_CANDIDATES = [
    "lender loan account number","loan account number",
    "loanaccountnumber","lenderloanaccountnumber",
    "account number","customer id","loan id","Transaction Id",
    "transaction id","TransactionId"
]

_KEY_HINTS = [
    "lenderloanaccountnumber","loanaccountnumber","lan",
    "accountnumber","loanid","customerid","transactionid"
]

def suggest_key_candidates(cols, top_k=5):
    scored = []
    for c in cols:
        n = re.sub(r"[^a-z0-9]+","",str(c).strip().lower())
        score = 0
        for h in _KEY_HINTS:
            if h in n: score += 10
        if n.endswith(("id","number","no","key","code","ref")): score += 3
        if n.startswith(("key","id","ref","reco","unique","primary")): score += 3
        if any(x in n for x in ["date","amount","sum","balance","rate","status","month","year"]): score -= 5
        scored.append((score,c))
    scored.sort(key=lambda x:(-x[0],str(x[1]).lower()))
    return [c for sc,c in scored if sc>=1][:top_k]

def _norm_col(x):
    x = str(x).replace("\u00A0"," ").replace("\u200B","")
    x = x.replace("\n"," ").replace("\r"," ").strip().lower()
    return re.sub(r"\s+","","".join(x.split()))

def build_effective_key(df, candidates):
    if df is None or df.empty:
        return pd.Series([], dtype=str), pd.Series([], dtype=str)
    norm_map = {_norm_col(c): c for c in df.columns}
    raw_key = pd.Series([""]*len(df), index=df.index, dtype=str)
    src_col = pd.Series([""]*len(df), index=df.index, dtype=str)
    for cand in candidates:
        actual = norm_map.get(_norm_col(cand))
        if not actual or actual not in df.columns: continue
        s = df[actual]
        valid = ~s.apply(is_blank)
        fill_mask = (raw_key=="") & valid
        if fill_mask.any():
            raw_key.loc[fill_mask] = s.loc[fill_mask].astype(str)
            src_col.loc[fill_mask] = actual
    return raw_key, src_col

def is_text_by_name(col):
    lc = str(col).strip().lower()
    return any(k in lc for k in ["loan account","lender loan","account number","lan","loan id","customer id"])

def is_numeric_column(df_both, col):
    if is_text_by_name(col): return False
    c1,c2 = f"{col}_f1", f"{col}_f2"
    if c1 not in df_both.columns or c2 not in df_both.columns: return False
    s1,s2 = df_both[c1], df_both[c2]
    nb1 = ~s1.apply(is_blank); nb2 = ~s2.apply(is_blank)
    n1 = pd.to_numeric(s1,errors="coerce"); n2 = pd.to_numeric(s2,errors="coerce")
    r1 = (n1.notna()&nb1).sum()/max(1,nb1.sum())
    r2 = (n2.notna()&nb2).sum()/max(1,nb2.sum())
    return (r1>=0.80) and (r2>=0.80)

def _first_non_blank(x):
    for v in x:
        if not is_blank(v): return v
    return x.iloc[0] if len(x)>0 else None

def _join_unique(x, sep="; "):
    vals,seen = [],set()
    for v in x:
        if is_blank(v): continue
        sv = str(v)
        if sv not in seen: seen.add(sv); vals.append(sv)
    return sep.join(vals)

def aggregate_by_key(df, key_col, numeric_cols, join_cols):
    df = df.copy()
    if key_col not in df.columns: return df
    dup_counts = df.groupby(key_col,dropna=False).size().rename("_DUP_COUNT").reset_index()
    agg_map = {}
    for c in df.columns:
        if c == key_col: continue
        if c in numeric_cols: agg_map[c] = "sum"
        elif c in join_cols: agg_map[c] = _join_unique
        else: agg_map[c] = _first_non_blank
    out = df.groupby(key_col,dropna=False).agg(agg_map).reset_index()
    return out.merge(dup_counts, on=key_col, how="left")

def uniq_list(seq):
    seen,out = set(),[]
    for x in seq:
        if x not in seen: seen.add(x); out.append(x)
    return out

def build_excel_bytes(sheets_dict):
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    HEADER_FILL = PatternFill("solid", fgColor="E7EEF8")
    BOLD = Font(bold=True)
    CENTER = Alignment(horizontal="center", vertical="center")
    for name, df in sheets_dict.items():
        ws = wb.create_sheet(str(name)[:31])
        if df is None or df.empty:
            ws.cell(row=1, column=1, value="No data")
            continue
        cols = list(df.columns)
        for j,col in enumerate(cols,1):
            cell = ws.cell(row=1,column=j,value=col)
            cell.font = BOLD; cell.fill = HEADER_FILL; cell.alignment = CENTER
        for i,row in enumerate(df.itertuples(index=False),2):
            for j,val in enumerate(row,1):
                ws.cell(row=i,column=j,value=val)
        ws.freeze_panes = "A2"
        for j,col in enumerate(cols,1):
            ws.column_dimensions[get_column_letter(j)].width = min(45,max(12,len(str(col))+2))
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# ============================================================
# MAIN APP UI
# ============================================================
st.markdown("### 📂 Upload Your Files")
st.caption("Upload CSV, XLSX, XLS, or XLSB files. Multiple files allowed on each side.")

col1, col2 = st.columns(2)
with col1:
    f1_list = st.file_uploader(
        "📄 File 1 — Source / System file",
        type=["csv","xlsx","xls","xlsb"],
        accept_multiple_files=True,
        key="file1"
    )
with col2:
    f2_list = st.file_uploader(
        "📄 File 2 — Target / Bank / Partner file",
        type=["csv","xlsx","xls","xlsb"],
        accept_multiple_files=True,
        key="file2"
    )

if not (f1_list and f2_list):
    st.info("⬆ Upload at least 1 file on both sides to begin reconciliation.")

    # Show upgrade banner for free users
    if plan == "free":
        st.markdown("""
        <div class="upgrade-banner">
            <div class="upgrade-text">
                🆓 <strong>Free Plan:</strong> Limited to 100 rows total input.
                Upgrade for unlimited reconciliation up to 250 MB files.
            </div>
            <a href="https://aarohisharma5000.github.io/recosuite" class="upgrade-btn" target="_blank">
                Upgrade ₹299 →
            </a>
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# ── Read files ──
sig1 = tuple((f.name, len(f.getvalue())) for f in f1_list)
sig2 = tuple((f.name, len(f.getvalue())) for f in f2_list)
sig  = (sig1, sig2)

if st.session_state.get("files_sig") != sig:
    with st.spinner("Reading files..."):
        try:
            dfs1 = [normalize_cols(read_file_once(f)) for f in f1_list]
            dfs2 = [normalize_cols(read_file_once(f)) for f in f2_list]
            df1  = pd.concat(dfs1, ignore_index=True) if dfs1 else pd.DataFrame()
            df2  = pd.concat(dfs2, ignore_index=True) if dfs2 else pd.DataFrame()
            for f in f1_list: df1["_SOURCE_FILE"] = f.name
            for f in f2_list: df2["_SOURCE_FILE"] = f.name
        except Exception as e:
            st.error(f"Error reading files: {e}")
            st.stop()
    st.session_state.update({
        "files_sig": sig, "df1": df1, "df2": df2,
        "run_done": False, "run_cache_ready": False
    })
else:
    df1 = st.session_state["df1"]
    df2 = st.session_state["df2"]

df1 = normalize_cols(df1)
df2 = normalize_cols(df2)
st.session_state["df1"] = df1
st.session_state["df2"] = df2

# ── FREE PLAN ROW LIMIT CHECK ──
if plan == "free":
    total_input_rows = len(df1) + len(df2)
    if total_input_rows > ROW_LIMIT:
        st.markdown(f"""
        <div class="upgrade-banner">
            <div class="upgrade-text">
                🔒 <strong>Free Plan limit:</strong> Your files have <strong>{total_input_rows:,} rows</strong>
                but Free plan allows max <strong>{ROW_LIMIT} rows</strong> total input.
                Upgrade to process unlimited rows.
            </div>
            <a href="https://aarohisharma5000.github.io/recosuite" class="upgrade-btn" target="_blank">
                Upgrade ₹299 →
            </a>
        </div>
        """, unsafe_allow_html=True)
        st.warning("⚠️ Showing reconciliation on first 50 rows of each file only for preview.")
        df1 = df1.head(50)
        df2 = df2.head(50)

# ── File summary ──
internal_cols = {"_SOURCE_FILE","_KEY","_merge","_SHEET"}
common_cols   = sorted(list(set(df1.columns).intersection(set(df2.columns))))
matched_cols  = [c for c in common_cols if c not in internal_cols]

c1m, c2m, c3m = st.columns(3)
c1m.metric("File 1 rows", f"{len(df1):,}")
c2m.metric("File 2 rows", f"{len(df2):,}")
c3m.metric("Common columns", len(matched_cols))

if not matched_cols:
    st.error("❌ No common columns found between the two files. Please check your headers match.")
    st.stop()

# ── Key column selection ──
st.markdown("### 🔑 Match Key Column")
_common_map = {_norm_col(c): c for c in common_cols}
existing_keys = []
for cand in KEY_CANDIDATES:
    k = _norm_col(cand)
    if k in _common_map and _common_map[k] not in existing_keys:
        existing_keys.append(_common_map[k])

auto_suggest = existing_keys if existing_keys else suggest_key_candidates(common_cols, top_k=3)

if existing_keys:
    st.success(f"✅ Auto-detected key: **{' / '.join(existing_keys[:2])}**")
else:
    st.warning("⚠️ Could not auto-detect key. Please select manually below.")

key_cols_sel = st.multiselect(
    "Select key column(s) for matching",
    options=uniq_list(common_cols),
    default=[c for c in auto_suggest if c in common_cols][:2]
)

if not key_cols_sel:
    st.warning("Please select at least one key column.")
    st.stop()

# ── Settings ──
st.markdown("### ⚙️ Settings")
s1, s2, s3 = st.columns(3)
with s1:
    compare_cols = st.multiselect(
        "Columns to compare",
        options=matched_cols,
        default=[c for c in matched_cols if c not in internal_cols and c not in key_cols_sel][:8]
    )
with s2:
    if compare_cols:
        focus_col = st.selectbox("Summary column (pick one)", options=compare_cols)
    else:
        st.warning("Select columns to compare first.")
        st.stop()
with s3:
    numeric_tol = st.number_input("Numeric tolerance (+/-)", min_value=0.0, value=5.0, step=1.0)
    treat_blanks = st.checkbox("Treat blanks as equal", value=True)

if not compare_cols:
    st.warning("Please select at least one column to compare.")
    st.stop()

# ── Run button ──
st.markdown("---")
if st.button("🚀 Run Reconciliation", type="primary", use_container_width=True):
    st.session_state["run_done"] = True

if not st.session_state.get("run_done", False):
    st.info("☝ Configure settings above then click **Run Reconciliation**")
    st.stop()

# ============================================================
# RUN RECONCILIATION
# ============================================================
with st.spinner("Running reconciliation..."):
    prog = st.progress(0)

    # Step 1: Build keys
    prog.progress(15, text="Building match keys...")
    df1_raw = df1.copy()
    df2_raw = df2.copy()

    raw1, _ = build_effective_key(df1_raw, key_cols_sel)
    raw2, _ = build_effective_key(df2_raw, key_cols_sel)
    df1_raw["_KEY"] = clean_key_series(raw1, treat_as_text=True)
    df2_raw["_KEY"] = clean_key_series(raw2, treat_as_text=True)
    df1_raw = df1_raw[df1_raw["_KEY"] != ""].copy()
    df2_raw = df2_raw[df2_raw["_KEY"] != ""].copy()

    # Duplicates
    c1_counts = df1_raw.groupby("_KEY").size()
    c2_counts = df2_raw.groupby("_KEY").size()
    dup_keys_1 = set(c1_counts[c1_counts>1].index)
    dup_keys_2 = set(c2_counts[c2_counts>1].index)
    dup_union  = dup_keys_1.union(dup_keys_2)

    # Step 2: Aggregate
    prog.progress(35, text="Aggregating duplicates...")
    safe_cols = [c for c in compare_cols if c in df1_raw.columns and c in df2_raw.columns]
    needed    = uniq_list(["_KEY","_SOURCE_FILE","_SHEET"] + safe_cols)

    df1_n = df1_raw[[c for c in needed if c in df1_raw.columns]].copy()
    df2_n = df2_raw[[c for c in needed if c in df2_raw.columns]].copy()

    numeric_cols = []
    for c in safe_cols:
        if is_text_by_name(c): continue
        n1 = pd.to_numeric(df1_n.get(c, pd.Series()), errors="coerce")
        n2 = pd.to_numeric(df2_n.get(c, pd.Series()), errors="coerce")
        if (n1.notna().sum() + n2.notna().sum()) > 0:
            numeric_cols.append(c)
            df1_n[c] = pd.to_numeric(df1_n[c].where(~df1_n[c].apply(is_blank), None), errors="coerce").fillna(0)
            df2_n[c] = pd.to_numeric(df2_n[c].where(~df2_n[c].apply(is_blank), None), errors="coerce").fillna(0)

    df1_ = aggregate_by_key(df1_n, "_KEY", numeric_cols=numeric_cols, join_cols=["_SOURCE_FILE","_SHEET"])
    df2_ = aggregate_by_key(df2_n, "_KEY", numeric_cols=numeric_cols, join_cols=["_SOURCE_FILE","_SHEET"])

    # Step 3: Merge
    prog.progress(55, text="Merging and comparing...")
    merged  = df1_.merge(df2_, on="_KEY", how="outer", suffixes=("_f1","_f2"), indicator=True)
    only_f1 = merged[merged["_merge"]=="left_only"].copy()
    only_f2 = merged[merged["_merge"]=="right_only"].copy()
    both    = merged[merged["_merge"]=="both"].copy()

    # Step 4: Match vs mismatch
    prog.progress(70, text="Identifying matches and mismatches...")
    focus    = focus_col
    c1n      = f"{focus}_f1"
    c2n      = f"{focus}_f2"
    s1s      = both.get(c1n, pd.Series())
    s2s      = both.get(c2n, pd.Series())
    is_num_focus = is_numeric_column(both, focus)

    if is_num_focus:
        n1 = pd.to_numeric(s1s, errors="coerce")
        n2 = pd.to_numeric(s2s, errors="coerce")
        mismatch = (n1-n2).abs() > (float(numeric_tol)+1e-9)
        if treat_blanks:
            mismatch = mismatch & (~(s1s.apply(is_blank)&s2s.apply(is_blank)))
    else:
        mismatch = s1s.astype(str) != s2s.astype(str)
        if treat_blanks:
            mismatch = mismatch & (~(s1s.apply(is_blank)&s2s.apply(is_blank)))

    both["_MATCH"] = ~mismatch
    matched    = both[both["_MATCH"]].copy()
    mismatched = both[~both["_MATCH"]].copy()

    # Add diff columns
    for col in safe_cols:
        if col in key_cols_sel: continue
        c1c = f"{col}_f1"; c2c = f"{col}_f2"
        if c1c not in both.columns or c2c not in both.columns: continue
        if is_numeric_column(both, col):
            for dfx in [matched, mismatched]:
                n1x = pd.to_numeric(dfx.get(c1c, pd.Series()), errors="coerce")
                n2x = pd.to_numeric(dfx.get(c2c, pd.Series()), errors="coerce")
                dfx[f"Diff_{col}"] = n1x - n2x

    # Step 5: Output columns
    prog.progress(85, text="Building output tables...")
    base  = ["_KEY"]
    extra = []
    for col in safe_cols:
        if col in key_cols_sel: continue
        c1c = f"{col}_f1"; c2c = f"{col}_f2"
        extra += [c1c, c2c]
        if f"Diff_{col}" in matched.columns:
            extra.append(f"Diff_{col}")

    keep = uniq_list(base + extra)
    matched_out    = matched.reindex(columns=keep).copy()
    mismatched_out = mismatched.reindex(columns=keep).copy()
    only_cols      = uniq_list(["_KEY", f"{focus}_f1", f"{focus}_f2"])
    only_f1_out    = only_f1.reindex(columns=only_cols).copy()
    only_f2_out    = only_f2.reindex(columns=only_cols).copy()

    # Step 6: Summary
    prog.progress(95, text="Building summary...")
    def _sum(dfx, col):
        return float(pd.to_numeric(dfx.get(col, pd.Series()), errors="coerce").fillna(0).sum())

    summary_rows = [
        {"Category":"✅ Matched",       "Count":len(matched_out),    f"{focus}_f1":_sum(matched_out,c1n),    f"{focus}_f2":_sum(matched_out,c2n),    "Diff":_sum(matched_out,c1n)-_sum(matched_out,c2n)},
        {"Category":"❌ Mismatched",    "Count":len(mismatched_out), f"{focus}_f1":_sum(mismatched_out,c1n), f"{focus}_f2":_sum(mismatched_out,c2n), "Diff":_sum(mismatched_out,c1n)-_sum(mismatched_out,c2n)},
        {"Category":"⬅ Only in File 1","Count":len(only_f1_out),    f"{focus}_f1":_sum(only_f1_out,f"{focus}_f1"), f"{focus}_f2":0.0, "Diff":_sum(only_f1_out,f"{focus}_f1")},
        {"Category":"➡ Only in File 2","Count":len(only_f2_out),    f"{focus}_f1":0.0, f"{focus}_f2":_sum(only_f2_out,f"{focus}_f2"), "Diff":-_sum(only_f2_out,f"{focus}_f2")},
        {"Category":"🔢 Duplicate Keys (info)","Count":len(dup_union), f"{focus}_f1":0.0, f"{focus}_f2":0.0, "Diff":0.0},
    ]
    total_row = {"Category":"TOTAL","Count":sum(r["Count"] for r in summary_rows[:4]),
                 f"{focus}_f1":sum(r[f"{focus}_f1"] for r in summary_rows[:4]),
                 f"{focus}_f2":sum(r[f"{focus}_f2"] for r in summary_rows[:4]),
                 "Diff":sum(r["Diff"] for r in summary_rows[:4])}
    summary_df = pd.concat([pd.DataFrame(summary_rows), pd.DataFrame([total_row])], ignore_index=True)

    prog.progress(100, text="✅ Done!")

# ============================================================
# RESULTS
# ============================================================

# ── FREE PLAN RESULT LIMIT ──
is_limited = False
if plan == "free":
    total_results = len(matched_out) + len(mismatched_out) + len(only_f1_out) + len(only_f2_out)
    if total_results > ROW_LIMIT:
        is_limited = True
        matched_out    = matched_out.head(ROW_LIMIT)
        mismatched_out = mismatched_out.head(ROW_LIMIT)
        only_f1_out    = only_f1_out.head(ROW_LIMIT)
        only_f2_out    = only_f2_out.head(ROW_LIMIT)

st.markdown("### 📊 Results")

# Summary metrics
m1, m2, m3, m4 = st.columns(4)
m1.metric("✅ Matched",        f"{summary_rows[0]['Count']:,}")
m2.metric("❌ Mismatched",     f"{summary_rows[1]['Count']:,}")
m3.metric("⬅ Only in File 1", f"{summary_rows[2]['Count']:,}")
m4.metric("➡ Only in File 2", f"{summary_rows[3]['Count']:,}")

# Summary table
st.dataframe(summary_df, use_container_width=True, hide_index=True)

# Upgrade banner for limited results
if is_limited:
    st.markdown(f"""
    <div class="upgrade-banner">
        <div class="upgrade-text">
            🔒 <strong>Free Plan:</strong> Showing first {ROW_LIMIT} rows only.
            Your full reconciliation has <strong>{total_results:,} rows</strong>.
            Upgrade to see all results and download.
        </div>
        <a href="https://aarohisharma5000.github.io/recosuite" class="upgrade-btn" target="_blank">
            Upgrade ₹299 →
        </a>
    </div>
    """, unsafe_allow_html=True)

# Result tabs
tabs = st.tabs([
    f"✅ Matched ({len(matched_out):,})",
    f"❌ Mismatched ({len(mismatched_out):,})",
    f"⬅ Only in File 1 ({len(only_f1_out):,})",
    f"➡ Only in File 2 ({len(only_f2_out):,})",
])
with tabs[0]: st.dataframe(matched_out, use_container_width=True, hide_index=True)
with tabs[1]: st.dataframe(mismatched_out, use_container_width=True, hide_index=True)
with tabs[2]: st.dataframe(only_f1_out, use_container_width=True, hide_index=True)
with tabs[3]: st.dataframe(only_f2_out, use_container_width=True, hide_index=True)

# ============================================================
# DOWNLOADS
# ============================================================
st.markdown("### ⬇ Download Results")

if plan == "free" and is_limited:
    st.warning("🔒 Downloads are limited to 100 rows on Free plan. Upgrade for full downloads.")

# CSV Downloads
d1, d2, d3, d4 = st.columns(4)
with d1:
    st.download_button(
        "⬇ Summary CSV",
        data=summary_df.to_csv(index=False).encode("utf-8"),
        file_name="reco_summary.csv", mime="text/csv"
    )
with d2:
    st.download_button(
        "⬇ Mismatched CSV",
        data=mismatched_out.to_csv(index=False).encode("utf-8"),
        file_name="reco_mismatched.csv", mime="text/csv"
    )
with d3:
    st.download_button(
        "⬇ Only in File 1",
        data=only_f1_out.to_csv(index=False).encode("utf-8"),
        file_name="reco_only_f1.csv", mime="text/csv"
    )
with d4:
    st.download_button(
        "⬇ Only in File 2",
        data=only_f2_out.to_csv(index=False).encode("utf-8"),
        file_name="reco_only_f2.csv", mime="text/csv"
    )

# Excel download
if st.button("📦 Prepare Full Excel (all sheets)"):
    with st.spinner("Building Excel..."):
        xl_bytes = build_excel_bytes({
            "Summary":       summary_df,
            "Matched":       matched_out,
            "Mismatched":    mismatched_out,
            "Only_in_File1": only_f1_out,
            "Only_in_File2": only_f2_out,
        })
    st.download_button(
        "⬇ Download Full Excel",
        data=xl_bytes,
        file_name="reco_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ── Footer ──
st.markdown("---")
st.markdown(
    f"<div style='text-align:center; color:#888; font-size:0.8rem;'>"
    f"📌 RecoSuite · {user_email} · {plan.title()} Plan · "
    f"<a href='https://aarohisharma5000.github.io/recosuite' style='color:#00d4ff;'>Upgrade</a> · "
    f"<a href='https://wa.me/919818799197' style='color:#25D366;'>WhatsApp Support</a>"
    f"</div>",
    unsafe_allow_html=True
)
