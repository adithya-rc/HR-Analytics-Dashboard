"""
HR Analytics Dashboard
=======================
Upload your Employee Master workbook and Attrition Tracker workbook
(drop both into the same uploader). Curated views only.

Run with:  streamlit run app.py
"""

import io
import datetime
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="HR Analytics Dashboard", layout="wide", page_icon="📊")

PALETTE = px.colors.qualitative.Set2

st.markdown("""
<style>
div[data-testid="stMetric"] {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 10px;
    padding: 14px 16px 8px 16px;
}
section[data-testid="stSidebar"] { width: 340px !important; }

/* Make dropdown/multiselect option lists (e.g. Exit Period range,
   PIP Review Status) scrollable instead of being cut off at the edge
   of the screen, and give the scrollbar visible styling so it's
   obvious there's more to scroll to (like the "Custom" option). */
div[data-baseweb="popover"] ul[role="listbox"] {
    max-height: 260px !important;
    overflow-y: auto !important;
}
div[data-baseweb="popover"] ul[role="listbox"]::-webkit-scrollbar {
    width: 8px;
}
div[data-baseweb="popover"] ul[role="listbox"]::-webkit-scrollbar-track {
    background: rgba(255,255,255,0.05);
}
div[data-baseweb="popover"] ul[role="listbox"]::-webkit-scrollbar-thumb {
    background: rgba(255,255,255,0.30);
    border-radius: 4px;
}

/* Sticky main header — pinned summary of the currently filtered data,
   stays visible while scrolling so you don't have to hunt for it. */
.main-header-bar {
    position: sticky;
    top: 0;
    z-index: 999;
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    background: #0e1117;
    padding: 12px 14px 10px 14px;
    border-bottom: 1px solid rgba(255,255,255,0.15);
    margin: 0 0 6px 0;
}
.mh-item {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 10px;
    padding: 6px 14px;
    min-width: 108px;
    display: block;
    text-decoration: none;
    color: inherit;
    cursor: pointer;
    transition: background 0.15s ease, border-color 0.15s ease, transform 0.1s ease;
}
a.mh-item:hover {
    background: rgba(255,255,255,0.09);
    border-color: rgba(255,255,255,0.30);
    transform: translateY(-1px);
}
a.mh-item:active {
    transform: translateY(0px);
}
.mh-label {
    font-size: 0.72rem;
    opacity: 0.65;
    white-space: nowrap;
}
.mh-value {
    font-size: 1.25rem;
    font-weight: 700;
    line-height: 1.3;
    white-space: nowrap;
}

/* Keep jump-to-section links from landing underneath the sticky header. */
h1, h2, h3 { scroll-margin-top: 92px; }
</style>
""", unsafe_allow_html=True)

st.title("📊 HR Analytics Dashboard")

# ==========================================================================
# Loading
# ==========================================================================

@st.cache_data(show_spinner=False)
def load_all_sheets(file_bytes: bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    out = {}
    for s in xl.sheet_names:
        try:
            d = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s)
            d.columns = [str(c).strip() for c in d.columns]
            d = d.dropna(axis=1, how="all").dropna(axis=0, how="all")
            if not d.empty:
                out[s] = d
        except Exception:
            continue
    return out


def auto_parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in df.columns:
        name = c.lower()
        if any(k in name for k in ["date", "doj", "dor", "lwd", "dob", "exit", "evaluation", "revoke"]):
            try:
                converted = pd.to_datetime(df[c], errors="coerce")
                if converted.notna().sum() >= max(1, int(df[c].notna().sum() * 0.5)):
                    df[c] = converted
            except Exception:
                pass
    return df


ROLE_KEYWORDS = {
    "id":            ["emp id", "employee id", "empid"],
    "name":          ["employee name", "emp name"],
    "gender":        ["gender", "sex"],
    "doj":           ["doj", "date of joining", "joining date", "hire date"],
    "dob":           ["date of birth", "dob", "birth date", "birthday"],
    "exit_date":     ["date of exit", "exit date", "dor", "lwd", "last working day", "separation date"],
    "status":        ["employee status", "employment status", "status"],
    "department":    ["department", "dept"],
    "division":      ["division"],
    "vertical":      ["vertical", "business unit"],
    "location":      ["location", "office", "site", "city"],
    "manager":       ["manager", "reporting manager", "l1 manager"],
    "designation":   ["designation", "title", "position"],
    "reason":        ["reason"],
    "attrition_type": ["attrition type"],
    "tenure_days":   ["tenure[days]", "tenure days", "tenure (days)"],
    "dor":           ["dor"],
    "lwd":           ["lwd", "last working day"],
    "pip_start":     ["pip start date"],
    "pip_end":       ["pip revoke date"],
    "pip_status":    ["review status", "pip status", "pip outcome", "outcome"],
}


def find_col(df, role):
    for kw in ROLE_KEYWORDS.get(role, []):
        for c in df.columns:
            if kw in c.lower():
                return c
    return None


def detect_roles(df):
    return {r: find_col(df, r) for r in ROLE_KEYWORDS if find_col(df, r)}


def pct(n, d):
    return 0.0 if not d else round(100 * n / d, 1)


def clean_series(df, col, min_opts=2, max_opts=400):
    if not col or col not in df.columns:
        return []
    vals = sorted(df[col].dropna().astype(str).unique())
    return vals if min_opts <= len(vals) <= max_opts else []


# ==========================================================================
# Sidebar — Upload
# ==========================================================================

st.sidebar.header("Upload Data")
uploaded_files = st.sidebar.file_uploader(
    "Employee Master + Attrition Tracker workbooks",
    type=["xlsx", "xls", "xlsm"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("⬅️ Upload your Employee Master and Attrition Tracker workbooks to begin (drop both into the same uploader).")
    st.stop()

all_sheets = {}
sheet_meta = []
for f in uploaded_files:
    for s_name, s_df in load_all_sheets(f.getvalue()).items():
        parsed = auto_parse_dates(s_df)
        key = f"{f.name} → {s_name}"
        all_sheets[key] = parsed
        sheet_meta.append({"key": key, "sheet": s_name, "file": f.name, "rows": parsed.shape[0]})

if not all_sheets:
    st.error("No readable data found in the uploaded file(s).")
    st.stop()

# ==========================================================================
# Sheet detection — strict: Employee Master must have BOTH a join-date and a
# gender column, since no tracker sheet (Notice/PIP/Exits/Summary) has either.
# ==========================================================================

TRACKER_KW = ["notice", "pip", "exit", "summary", "dropdown", "backfill"]

def score_sheet(name, sdf):
    r = detect_roles(sdf)
    is_master_shaped = ("doj" in r) and ("gender" in r)
    is_tracker_named = any(kw in name.lower() for kw in TRACKER_KW)
    return (is_master_shaped, len(r), not is_tracker_named, sdf.shape[0])

best_key = max(all_sheets, key=lambda k: score_sheet(k, all_sheets[k]))

def guess_role(name, sdf):
    n = name.lower()
    if "notice" in n: return "📋 Notice Period"
    if "pip" in n: return "📝 PIP"
    if "exit" in n: return "🚪 Exits"
    if "summary" in n: return "📈 Summary"
    r = detect_roles(sdf)
    if "doj" in r and "gender" in r: return "👤 Employee Master"
    return "— other —"

emp = all_sheets[best_key]
roles = detect_roles(emp)
master_ok = ("doj" in roles) and ("gender" in roles)

def locate_sheet(keyword):
    matches = [k for k in all_sheets if keyword in k.lower() and k != best_key]
    return all_sheets[matches[0]] if matches else pd.DataFrame()

notice_df = locate_sheet("notice")
pip_df = locate_sheet("pip")
exits_df = locate_sheet("exit")
summary_df = locate_sheet("summary")

with st.sidebar.expander("📁 Data source", expanded=not master_ok):
    if not master_ok:
        st.warning(
            "No sheet with both a joining-date and gender column was found — make sure your **Employee "
            "Master workbook** is included in the upload. Showing the closest match below; override if needed."
        )
    st.caption("Auto-detected from the uploaded file(s):")
    summary_tbl = pd.DataFrame([
        {"Sheet": m["sheet"], "File": m["file"], "Rows": m["rows"], "Detected as": guess_role(m["sheet"], all_sheets[m["key"]])}
        for m in sheet_meta
    ])
    st.dataframe(summary_tbl, use_container_width=True, hide_index=True)

    def format_sheet_option(key):
        m = next(m for m in sheet_meta if m["key"] == key)
        short_file = m["file"] if len(m["file"]) <= 22 else m["file"][:19] + "…"
        return f"{m['sheet']}  ({short_file}, {m['rows']} rows)"

    main_key = st.selectbox(
        "Sheet to use as Employee Master (override)",
        list(all_sheets.keys()),
        index=list(all_sheets.keys()).index(best_key),
        format_func=format_sheet_option,
    )

if main_key != best_key:
    emp = all_sheets[main_key]
    roles = detect_roles(emp)
    notice_df = locate_sheet("notice") if "notice" not in main_key.lower() else pd.DataFrame()
    pip_df = locate_sheet("pip") if "pip" not in main_key.lower() else pd.DataFrame()
    exits_df = locate_sheet("exit") if "exit" not in main_key.lower() else pd.DataFrame()
    summary_df = locate_sheet("summary") if "summary" not in main_key.lower() else pd.DataFrame()

# ==========================================================================
# Sidebar — Vertical filter (locked to Motivity Labs only — every table,
# chart, and download in this app is restricted to this business unit;
# other verticals are excluded entirely, not just unselected.)
# ==========================================================================

FIXED_VERTICAL = "Motivity Labs"

vertical_values = set()
for d in [emp, notice_df, pip_df, exits_df, summary_df]:
    if d.empty:
        continue
    vcol = find_col(d, "vertical")
    if vcol:
        vertical_values.update(d[vcol].dropna().astype(str).tolist())
all_verticals = sorted(vertical_values)

selected_verticals = [FIXED_VERTICAL]

# Employee-ID → Vertical lookup, built from whichever sheet(s) actually
# carry a Vertical column. This is needed because the Employee Master
# sheet often does NOT have its own Vertical column — without this
# fallback, apply_vertical() would have nothing to filter on for that
# sheet and would silently let every business unit through.
id_vertical_map = {}
for d in [emp, notice_df, pip_df, exits_df, summary_df]:
    if d.empty:
        continue
    idc = find_col(d, "id")
    vcol = find_col(d, "vertical")
    if idc and vcol:
        for _id, _v in zip(d[idc].astype(str).str.strip(), d[vcol].astype(str)):
            if _id and _id not in id_vertical_map:
                id_vertical_map[_id] = _v

st.sidebar.header("Vertical")
if FIXED_VERTICAL in all_verticals:
    st.sidebar.success(f"Hard-locked to **{FIXED_VERTICAL}** only — every table and download below excludes anything not confirmed as this business unit.")
else:
    st.sidebar.warning(
        f"Hard-locked to **{FIXED_VERTICAL}**, but it wasn't found in the uploaded data "
        f"(detected: {', '.join(all_verticals) if all_verticals else 'none'}). No rows will match."
    )
if not find_col(emp, "vertical"):
    st.sidebar.caption(
        "⚠️ Your Employee Master sheet has no Vertical/Business Unit column of its own, so it can only "
        "confirm an employee as Motivity Labs by cross-referencing their Employee ID against Notice, PIP, "
        "or Exit records. Active employees who've never appeared on any of those (which is most of your "
        "active headcount) can't be confirmed and are excluded under this hard lock. Add a Vertical column "
        "to the Employee Master to include your full active population accurately."
    )

def apply_vertical(tdf):
    if tdf.empty:
        return tdf
    col = find_col(tdf, "vertical")
    if col:
        return tdf[tdf[col].astype(str).isin(selected_verticals)]
    # Hard lock for sheets with no Vertical column of their own (commonly
    # the Employee Master): only rows we can POSITIVELY confirm as
    # Motivity Labs via an Employee ID cross-reference against
    # Notice/PIP/Exit are kept. Anything unconfirmed is excluded — no
    # benefit-of-the-doubt inclusion, per explicit instruction to only
    # ever show Motivity Labs data and nothing else.
    idc = find_col(tdf, "id")
    if idc and id_vertical_map:
        ids_series = tdf[idc].astype(str).str.strip()
        confirmed_ids = {i for i, v in id_vertical_map.items() if v in selected_verticals}
        return tdf[ids_series.isin(confirmed_ids)]
    # No Vertical column and no way to cross-reference at all — cannot
    # confirm anything as Motivity Labs, so exclude everything rather
    # than risk showing other verticals.
    return tdf.iloc[0:0]

# ==========================================================================
# Sidebar — segment filters
# ==========================================================================

st.sidebar.header("Filters")
df = apply_vertical(emp.copy())

filter_defs = [
    ("department", "Department"), ("location", "Location"),
    ("manager", "Reporting Manager"), ("gender", "Gender"), ("status", "Employee Status"),
    ("designation", "Designation"),
]
active_filters = {}
for role, label in filter_defs:
    col = roles.get(role)
    opts = clean_series(df, col)
    if opts:
        chosen = st.sidebar.multiselect(label, opts, key=f"filter_{role}")
        if chosen:
            df = df[df[col].astype(str).isin(chosen)]
            active_filters[role] = chosen

id_col_master = roles.get("id")

def filter_tracker(tdf):
    tdf = apply_vertical(tdf)
    if tdf.empty:
        return tdf
    out = tdf

    # Primary path: join on Employee ID against the already-filtered
    # Employee Master (df). This is what makes sidebar filters like Gender,
    # Designation, or Manager apply correctly to Notice/PIP/Exits sheets
    # even though those sheets don't carry those columns themselves —
    # previously such filters were silently skipped on those sheets.
    if (active_filters or selected_verticals) and id_col_master and id_col_master in df.columns:
        tid_col = find_col(out, "id")
        if tid_col:
            allowed_ids = set(df[id_col_master].astype(str).str.strip())
            out = out[out[tid_col].astype(str).str.strip().isin(allowed_ids)]

    # Belt-and-suspenders: also apply any filter directly by column name for
    # sheets that happen to carry their own copy of a filtered field.
    for role, _ in filter_defs:
        col = find_col(out, role)
        chosen = active_filters.get(role)
        if col and chosen:
            out = out[out[col].astype(str).isin(chosen)]
    return out

# ==========================================================================
# Sidebar — PIP Review Status filter (PIP sheet only)
# ==========================================================================

pip_status_col = find_col(pip_df, "pip_status")
pip_status_opts = clean_series(pip_df, pip_status_col) if not pip_df.empty else []
selected_pip_status = []
if pip_status_opts:
    st.sidebar.header("PIP Review Status")
    selected_pip_status = st.sidebar.multiselect(
        "Review Status", pip_status_opts, key="filter_pip_status",
        help="Leave empty to include all PIP review statuses.",
    )

def apply_pip_status(tdf):
    if tdf.empty or not selected_pip_status or not pip_status_col or pip_status_col not in tdf.columns:
        return tdf
    return tdf[tdf[pip_status_col].astype(str).isin(selected_pip_status)]

today = datetime.date.today()

def _point_in_range(tdf, date_col, range_start, range_end):
    if tdf.empty or range_start is None or not date_col or date_col not in tdf.columns:
        return tdf
    d = tdf[date_col]
    start_ts, end_ts = pd.Timestamp(range_start), pd.Timestamp(range_end) + pd.Timedelta(days=1)
    return tdf[(d >= start_ts) & (d < end_ts)]

# Base filtered segments (vertical + sidebar segment filters applied; no
# date-period filtering yet — each tracker below gets its own dedicated
# date-range filter instead of one blanket Report Period).
notice_seg = filter_tracker(notice_df)
notice_f = notice_seg

pip_seg = filter_tracker(pip_df)
pip_seg = apply_pip_status(pip_seg)
pip_f = pip_seg

exits_seg = filter_tracker(exits_df)
exits_event_col = find_col(exits_seg, "lwd") or find_col(exits_seg, "dor")

exit_col = roles.get("exit_date")
doj_col = roles.get("doj")
exited_all = df[df[exit_col].notna()] if exit_col and exit_col in df.columns else pd.DataFrame()

# ==========================================================================
# Sidebar — Joining Period (dedicated joining-date range — answers "how
# many employees joined between X and Y").
# ==========================================================================

st.sidebar.header("Joining Period")
use_joining_period = st.sidebar.checkbox(
    "Filter by a specific joining date range", key="use_joining_period",
    help="Shows exactly how many employees joined within the dates you pick.",
)

joining_period_start = joining_period_end = None
joined_custom = pd.DataFrame()
if use_joining_period:
    join_dates_avail = df[doj_col].dropna() if doj_col and doj_col in df.columns else pd.Series(dtype="datetime64[ns]")
    default_join_start = join_dates_avail.min().date() if not join_dates_avail.empty else today
    default_join_end = join_dates_avail.max().date() if not join_dates_avail.empty else today
    # Explicit wide bounds: st.date_input otherwise limits the year dropdown
    # to roughly ±10 years around the *value*, which can cut off the
    # current year if the earliest/latest joining date on record is old.
    join_bound_min = min(default_join_start, datetime.date(1990, 1, 1))
    join_bound_max = max(default_join_end, today)

    jc1, jc2 = st.sidebar.columns(2)
    joining_period_start = jc1.date_input("Join start date", value=default_join_start,
                                           min_value=join_bound_min, max_value=join_bound_max, key="joining_period_start")
    joining_period_end = jc2.date_input("Join end date", value=default_join_end,
                                         min_value=join_bound_min, max_value=join_bound_max, key="joining_period_end")

    if joining_period_start > joining_period_end:
        st.sidebar.error("Join start date is after join end date.")
    elif doj_col and doj_col in df.columns:
        joined_custom = _point_in_range(df, doj_col, joining_period_start, joining_period_end)

    st.sidebar.metric("Joiners in this range", len(joined_custom))

# ==========================================================================
# Sidebar — Exit Period (dedicated exit-date range — answers "how many
# employees left between X and Y". Also scopes the Exit Tracker sheet,
# exits_f, when active.)
# ==========================================================================

st.sidebar.header("Exit Period")
use_exit_period = st.sidebar.checkbox(
    "Filter by a specific exit date range", key="use_exit_period",
    help="Shows exactly how many employees left within the dates you pick.",
)

exit_period_start = exit_period_end = None
exited_custom = pd.DataFrame()
exits_f = exits_seg
if use_exit_period:
    exit_dates_avail = df[exit_col].dropna() if exit_col and exit_col in df.columns else pd.Series(dtype="datetime64[ns]")
    default_exit_start = exit_dates_avail.min().date() if not exit_dates_avail.empty else today
    default_exit_end = exit_dates_avail.max().date() if not exit_dates_avail.empty else today
    # Explicit wide bounds: st.date_input otherwise limits the year dropdown
    # to roughly ±10 years around the *value*, which can cut off the
    # current year if the earliest/latest exit on record is old.
    exit_bound_min = min(default_exit_start, datetime.date(1990, 1, 1))
    exit_bound_max = max(default_exit_end, today)

    ec1, ec2 = st.sidebar.columns(2)
    exit_period_start = ec1.date_input("Exit start date", value=default_exit_start,
                                        min_value=exit_bound_min, max_value=exit_bound_max, key="exit_period_start")
    exit_period_end = ec2.date_input("Exit end date", value=default_exit_end,
                                      min_value=exit_bound_min, max_value=exit_bound_max, key="exit_period_end")

    if exit_period_start > exit_period_end:
        st.sidebar.error("Exit start date is after exit end date.")
    else:
        if exit_col and exit_col in df.columns:
            exited_custom = _point_in_range(df, exit_col, exit_period_start, exit_period_end)
        exits_f = _point_in_range(exits_seg, exits_event_col, exit_period_start, exit_period_end)

    st.sidebar.metric("Exits in this range", len(exited_custom))

exited_period = exited_custom if use_exit_period else exited_all

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]

# ==========================================================================
# Sidebar — Birthdays (month-only filter, calculated from Date of Birth)
# ==========================================================================

dob_col = roles.get("dob")
st.sidebar.header("Birthdays")
if dob_col and dob_col in df.columns:
    birthday_month = st.sidebar.selectbox(
        "Birthday month", ["All"] + MONTH_NAMES, index=today.month, key="birthday_month",
    )
else:
    birthday_month = "All"
    st.sidebar.caption("No Date of Birth column detected in the Employee Master.")

birthday_df = pd.DataFrame()
if dob_col and dob_col in df.columns:
    valid_dob = df[dob_col].notna()
    if birthday_month == "All":
        birthday_df = df[valid_dob]
    else:
        month_num = MONTH_NAMES.index(birthday_month) + 1
        birthday_df = df[valid_dob & (df[dob_col].dt.month == month_num)]

# ==========================================================================
# Sidebar — Service Tenure (5/10/15-year anniversaries, month-only filter,
# calculated from Date of Joining)
# ==========================================================================

st.sidebar.header("Service Tenure")
if doj_col and doj_col in df.columns:
    tenure_month = st.sidebar.selectbox(
        "Anniversary month", ["All"] + MONTH_NAMES, index=today.month, key="tenure_month",
    )
else:
    tenure_month = "All"
    st.sidebar.caption("No Date of Joining column detected in the Employee Master.")

tenure_milestone_df = pd.DataFrame()
if doj_col and doj_col in df.columns:
    active_mask = df[exit_col].isna() if exit_col and exit_col in df.columns else pd.Series(True, index=df.index)
    tenure_base = df[active_mask & df[doj_col].notna()].copy()
    if tenure_month != "All":
        month_num = MONTH_NAMES.index(tenure_month) + 1
        tenure_base = tenure_base[tenure_base[doj_col].dt.month == month_num]
    tenure_base["Milestone (Years)"] = today.year - tenure_base[doj_col].dt.year
    tenure_milestone_df = tenure_base[tenure_base["Milestone (Years)"].isin([5, 10, 15])]

st.sidebar.divider()
st.sidebar.caption("Build: v7.1 — Fixed NaT date-comparison crash; hard-locked Vertical filter to Motivity Labs only, no exceptions")

# ==========================================================================
# Core figures (kept for use in section headers/captions below — the old
# 9-box KPI strip was removed since it would silently show 0/— whenever a
# role column wasn't detected, which read as "broken" even when filters
# were working fine).
# ==========================================================================

hc = len(df)

# Active headcount = filtered employees with no recorded exit date
exited_mask = df[exit_col].notna() if exit_col and exit_col in df.columns else pd.Series(False, index=df.index)
active_hc = int((~exited_mask).sum())

# Average tenure (years), open-ended employees measured to today
avg_tenure_years = None
if doj_col and doj_col in df.columns:
    end_dates = df[exit_col].copy() if exit_col and exit_col in df.columns else pd.Series(pd.NaT, index=df.index)
    end_dates = end_dates.fillna(pd.Timestamp(today))
    tenure_days = (end_dates - df[doj_col]).dt.days
    tenure_days = tenure_days[tenure_days.notna() & (tenure_days >= 0)]
    if not tenure_days.empty:
        avg_tenure_years = round(tenure_days.mean() / 365.25, 1)

# Early attrition — of the employees who exited within the current
# filter/period selection, how many left within their first 6 months?
# A far more actionable signal than a raw attrition % on its own.
early_attrition_count = None
early_attrition_pct_of_exits = None
early_attrition_df = pd.DataFrame()
if doj_col and exit_col and doj_col in exited_period.columns and exit_col in exited_period.columns and not exited_period.empty:
    tenure_at_exit_days = (exited_period[exit_col] - exited_period[doj_col]).dt.days
    early_mask = tenure_at_exit_days.notna() & (tenure_at_exit_days >= 0) & (tenure_at_exit_days <= 182)
    early_attrition_df = exited_period[early_mask]
    early_attrition_count = int(early_mask.sum())
    early_attrition_pct_of_exits = pct(early_attrition_count, len(exited_period))

# Actual attrition % YTD — exits from Jan 1 of this year to today, divided
# by the average headcount over that same window (headcount at the start
# of the year + headcount today, divided by two). This is independent of
# Exit Period/Joining Period toggles — it's always the
# calendar-year-to-date figure for whatever segment filters are active.
ytd_attrition_pct = None
ytd_exits = None
ytd_exits_df = pd.DataFrame()
if doj_col and exit_col and doj_col in df.columns and exit_col in df.columns:
    ytd_start = datetime.date(today.year, 1, 1)

    def _headcount_as_of(as_of_date):
        as_of_ts = pd.Timestamp(as_of_date)
        joined = df[doj_col] <= as_of_ts
        exited_by_then = df[exit_col] <= as_of_ts
        return int((joined & ~exited_by_then).sum())

    hc_start_of_year = _headcount_as_of(ytd_start - datetime.timedelta(days=1))
    hc_today = _headcount_as_of(today)
    ytd_avg_hc = (hc_start_of_year + hc_today) / 2

    ytd_exits_df = _point_in_range(df, exit_col, ytd_start, today)
    ytd_exits = len(ytd_exits_df)
    if ytd_avg_hc > 0:
        ytd_attrition_pct = round(100 * ytd_exits / ytd_avg_hc, 1)

# ==========================================================================
# Sticky main header — the primary at-a-glance output for whatever filters
# are currently selected (vertical, department/location/etc., PIP status,
# Exit Period, Joining Period). Recomputes on every filter
# change and stays pinned at the top of the page so it never needs
# scrolling to find.
# ==========================================================================

header_items = [("Headcount", hc, "overview"), ("Active", active_hc, "overview")]
if avg_tenure_years is not None:
    header_items.append(("Avg tenure", f"{avg_tenure_years} yrs", "overview"))
if early_attrition_count is not None:
    header_items.append(("Early attrition (≤6mo)", f"{early_attrition_count} ({early_attrition_pct_of_exits}% of exits)", "reasons"))
if ytd_attrition_pct is not None:
    header_items.append(("Attrition % YTD", f"{ytd_attrition_pct}% ({ytd_exits})", "reasons"))
if not notice_df.empty:
    header_items.append(("On notice", len(notice_f), "notice-pip"))
if not pip_df.empty:
    header_items.append(("On PIP", len(pip_f), "notice-pip"))
if use_exit_period:
    header_items.append((f"Exits {exit_period_start} → {exit_period_end}", len(exited_custom), "exits"))
if use_joining_period:
    header_items.append((f"Joiners {joining_period_start} → {joining_period_end}", len(joined_custom), "joiners"))
if dob_col:
    header_items.append((f"Birthdays ({birthday_month})", len(birthday_df), "birthdays"))
if doj_col:
    header_items.append((f"Service Tenure ({tenure_month})", len(tenure_milestone_df), "tenure"))

header_html = "".join(
    f'<a class="mh-item" href="#{anchor}"><div class="mh-label">{label}</div><div class="mh-value">{value}</div></a>'
    for label, value, anchor in header_items
)
st.markdown(f'<div class="main-header-bar">{header_html}</div>', unsafe_allow_html=True)

st.divider()

# ==========================================================================
# Single-page layout — every section stacked vertically instead of tabs,
# so everything is visible with one scroll instead of clicking around.
# Each section only renders its metric line for values that are actually
# available, instead of a fixed 9-box grid that shows 0/— for anything
# undetected. The header boxes above link straight to these sections via
# their anchors, and the Joiners section only exists at all when the
# Joining Period filter is switched on, to keep the page clutter-free.
# ==========================================================================

tab_overview = st.container()
st.divider()
tab_reasons = st.container()
st.divider()
tab_np = st.container()
st.divider()
if use_joining_period:
    tab_joiners = st.container()
    st.divider()
else:
    tab_joiners = None
tab_exits = st.container()
st.divider()
tab_birthdays = st.container()
st.divider()
tab_tenure = st.container()
st.divider()
tab_trends = st.container()

# ==========================================================================
# SECTION: Overview — Gender diversity + headcount breakdowns
# ==========================================================================

with tab_overview:
    st.header("🏠 Overview", anchor="overview")

    gender_col = roles.get("gender")
    dept_col = roles.get("department")
    loc_col = roles.get("location")

    cols_avail = [c for c in [gender_col, dept_col, loc_col] if c and clean_series(df, c)]
    if cols_avail:
        cols_ui = st.columns(len(cols_avail))
        for i, col in enumerate(cols_avail):
            vc = df[col].value_counts().head(10)
            if col == gender_col:
                fig = px.pie(values=vc.values, names=vc.index, hole=0.5, title="Gender Diversity",
                             color_discrete_sequence=PALETTE)
            else:
                fig = px.bar(vc, orientation="h", title=col, color_discrete_sequence=PALETTE)
                fig.update_layout(showlegend=False, yaxis_title=None, xaxis_title="Headcount")
            cols_ui[i].plotly_chart(fig, use_container_width=True, key=f"ov_{col}")
    else:
        st.info("No demographic fields detected to chart.")

    st.subheader("Employee List (filtered)")
    st.dataframe(df, use_container_width=True, height=350)
    st.download_button("Download filtered data (CSV)", df.to_csv(index=False).encode("utf-8"),
                        "filtered_employees.csv", "text/csv")

with tab_reasons:
    st.header("📌 Reasons & Managers", anchor="reasons")

    # --- Early Attrition: employees who left within 6 months of joining ---
    if early_attrition_count is not None:
        st.subheader("Early Attrition (≤ 6 months)")
        st.metric("Left within 6 months of joining", f"{early_attrition_count} ({early_attrition_pct_of_exits}% of exits)",
                   help=f"{early_attrition_count} of the {len(exited_period)} exits in the current filter/period "
                        f"selection left within 182 days of joining.")
        st.dataframe(early_attrition_df, use_container_width=True, height=300)
        st.download_button(
            "Download early leavers (CSV)",
            early_attrition_df.to_csv(index=False).encode("utf-8"),
            "early_attrition.csv", "text/csv", key="dl_early_attrition",
        )
        st.divider()

    # --- Attrition % YTD: total exits + rate from Jan 1 to today ---
    if ytd_attrition_pct is not None:
        st.subheader(f"Attrition % — Year to Date ({today.year})")
        st.metric("YTD attrition rate", f"{ytd_attrition_pct}% ({ytd_exits} exits)",
                   help=f"{ytd_exits} exits from Jan 1, {today.year} to today, divided by average headcount "
                        f"(headcount at start of year + headcount today, ÷ 2) for the currently selected segment filters. "
                        f"Independent of the Exit Period/Joining Period toggles.")
        st.dataframe(ytd_exits_df, use_container_width=True, height=300)
        st.download_button(
            "Download YTD leavers (CSV)",
            ytd_exits_df.to_csv(index=False).encode("utf-8"),
            "ytd_attrition.csv", "text/csv", key="dl_ytd_attrition",
        )
        st.divider()

    if exited_period.empty:
        st.info("No exited employees in the current filter/period selection.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            reason_col = roles.get("reason")
            if reason_col and exited_period[reason_col].notna().any():
                vc = exited_period[reason_col].value_counts().head(10)
                fig = px.bar(vc, orientation="h", title="Attrition Reasons", color_discrete_sequence=PALETTE)
                fig.update_layout(showlegend=False, yaxis_title=None, xaxis_title="Exits")
                st.plotly_chart(fig, use_container_width=True)
        with c2:
            type_col = roles.get("attrition_type")
            if type_col and exited_period[type_col].notna().any():
                vc = exited_period[type_col].value_counts()
                fig = px.pie(values=vc.values, names=vc.index, hole=0.5, title="Voluntary vs Involuntary",
                             color_discrete_sequence=PALETTE)
                st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Reporting Manager Breakdown")
        mgr_col = roles.get("manager")
        if mgr_col and mgr_col in df.columns and df[mgr_col].notna().any():
            hc_by_mgr = df[mgr_col].value_counts()
            exit_by_mgr = exited_period[mgr_col].value_counts() if not exited_period.empty else pd.Series(dtype=int)
            mgr_tbl = pd.DataFrame({"Headcount": hc_by_mgr}).join(
                pd.DataFrame({"Attrition Count": exit_by_mgr}), how="left"
            ).fillna(0)
            mgr_tbl["Attrition Count"] = mgr_tbl["Attrition Count"].astype(int)
            mgr_tbl["Attrition %"] = (mgr_tbl["Attrition Count"] / mgr_tbl["Headcount"] * 100).round(1)
            mgr_tbl = mgr_tbl.sort_values("Attrition %", ascending=False)
            mgr_tbl.index.name = "Reporting Manager"
            st.dataframe(mgr_tbl, use_container_width=True, height=380)
        else:
            st.info("No reporting-manager field detected in this dataset.")

    with c2:
        st.subheader("Department Attrition Rate")
        dept_col = roles.get("department")
        if dept_col and dept_col in df.columns and df[dept_col].notna().any():
            hc_by_dept = df[dept_col].value_counts()
            exit_by_dept = exited_period[dept_col].value_counts() if not exited_period.empty else pd.Series(dtype=int)
            dept_tbl = pd.DataFrame({"Headcount": hc_by_dept}).join(
                pd.DataFrame({"Attrition Count": exit_by_dept}), how="left"
            ).fillna(0)
            dept_tbl["Attrition Count"] = dept_tbl["Attrition Count"].astype(int)
            dept_tbl["Attrition %"] = (dept_tbl["Attrition Count"] / dept_tbl["Headcount"] * 100).round(1)
            dept_tbl = dept_tbl.sort_values("Attrition %", ascending=False)
            dept_tbl.index.name = "Department"
            st.dataframe(dept_tbl, use_container_width=True, height=380)
        else:
            st.info("No department field detected in this dataset.")

with tab_np:
    st.header("📋 Notice & PIP", anchor="notice-pip")
    if selected_verticals:
        st.caption(f"Vertical filter active: **{', '.join(selected_verticals)}**")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader(f"On Notice Period ({len(notice_f)})")
        if notice_f.empty:
            st.info("No notice-period records for the current selection.")
        else:
            st.dataframe(notice_f, use_container_width=True, height=420)
    with c2:
        st.subheader(f"On PIP ({len(pip_f)})")
        if pip_f.empty:
            st.info("No PIP records for the current selection.")
        else:
            if pip_status_col and pip_status_col in pip_f.columns and pip_f[pip_status_col].notna().any():
                vc = pip_f[pip_status_col].value_counts()
                fig = px.pie(values=vc.values, names=vc.index, hole=0.5, title="PIP Review Status",
                             color_discrete_sequence=PALETTE)
                st.plotly_chart(fig, use_container_width=True)
            st.dataframe(pip_f, use_container_width=True, height=420)

if tab_joiners is not None:
    with tab_joiners:
        st.header("🆕 New Joiners", anchor="joiners")
        st.subheader(f"Joined between {joining_period_start} and {joining_period_end}")
        st.metric("Employees who joined in this range", len(joined_custom))
        if joined_custom.empty:
            st.info("No employees joined in the selected range.")
        else:
            st.dataframe(joined_custom, use_container_width=True, height=350)
            st.download_button(
                "Download joiners in range (CSV)",
                joined_custom.to_csv(index=False).encode("utf-8"),
                "joiners_in_range.csv", "text/csv", key="dl_joining_period",
            )

with tab_exits:
    st.header("🚪 Exits", anchor="exits")

    if use_exit_period:
        st.subheader(f"Exits between {exit_period_start} and {exit_period_end}")
        st.metric("Employees who left in this range", len(exited_custom))
        if exited_custom.empty:
            st.info("No employees exited in the selected range.")
        else:
            st.dataframe(exited_custom, use_container_width=True, height=300)
            st.download_button(
                "Download exits in range (CSV)",
                exited_custom.to_csv(index=False).encode("utf-8"),
                "exits_in_range.csv", "text/csv", key="dl_exit_period",
            )
        st.divider()

    if selected_verticals:
        st.caption(f"Vertical filter active: **{', '.join(selected_verticals)}**")

    st.subheader(f"Exited Employees — Employee Master ({len(exited_period)})")
    if exited_period.empty:
        st.info("No exited employees for the current selection.")
    else:
        st.dataframe(exited_period, use_container_width=True, height=300)

    st.subheader(f"Exit Tracker ({len(exits_f)})")
    if exits_f.empty:
        st.info("No exit-tracker records for the current selection.")
    else:
        reason_col = find_col(exits_f, "reason")
        if reason_col and exits_f[reason_col].notna().any():
            vc = exits_f[reason_col].value_counts().head(10)
            fig = px.bar(vc, orientation="h", title="Exit Reasons (Tracker)", color_discrete_sequence=PALETTE)
            fig.update_layout(showlegend=False, yaxis_title=None, xaxis_title="Exits")
            st.plotly_chart(fig, use_container_width=True)
        st.dataframe(exits_f, use_container_width=True, height=420)

with tab_birthdays:
    st.header("🎂 Birthdays", anchor="birthdays")
    if not dob_col or dob_col not in df.columns:
        st.info("No Date of Birth column detected in the Employee Master, so birthdays can't be shown.")
    else:
        st.caption(f"Showing birthdays for: **{birthday_month}**")
        if birthday_df.empty:
            st.info("No birthdays found for the selected month.")
        else:
            st.subheader(f"Employees with a birthday in {birthday_month} ({len(birthday_df)})")
            st.dataframe(birthday_df, use_container_width=True, height=350)
            st.download_button(
                "Download birthdays (CSV)",
                birthday_df.to_csv(index=False).encode("utf-8"),
                "birthdays.csv", "text/csv", key="dl_birthdays",
            )

with tab_tenure:
    st.header("🎉 Service Tenure Milestones (5 / 10 / 15 Years)", anchor="tenure")
    if not doj_col or doj_col not in df.columns:
        st.info("No Date of Joining column detected in the Employee Master, so service tenure milestones can't be calculated.")
    else:
        st.caption(f"Showing 5/10/15-year anniversaries for: **{tenure_month}** (active employees only)")
        if tenure_milestone_df.empty:
            st.info("No employees hitting a 5, 10, or 15-year milestone in the selected month.")
        else:
            for milestone in (5, 10, 15):
                sub = tenure_milestone_df[tenure_milestone_df["Milestone (Years)"] == milestone]
                if not sub.empty:
                    st.subheader(f"{milestone}-Year Anniversaries ({len(sub)})")
                    st.dataframe(sub, use_container_width=True, height=250)
            st.download_button(
                "Download service tenure milestones (CSV)",
                tenure_milestone_df.to_csv(index=False).encode("utf-8"),
                "service_tenure_milestones.csv", "text/csv", key="dl_tenure",
            )

with tab_trends:
    st.header("📈 Trends", anchor="trends")
    st.subheader("Monthly Trend (Summary Sheet)")
    summary_seg = apply_vertical(summary_df) if not summary_df.empty else summary_df

    if summary_seg.empty:
        st.info("No **Summary** sheet was detected in the uploaded workbook(s), so no monthly trend is available.")
    else:
        # Identify the month/period column
        month_col = None
        for c in summary_seg.columns:
            cl = c.lower()
            if "month" in cl or "period" in cl:
                month_col = c
                break
        if month_col is None:
            for c in summary_seg.columns:
                if "date" in c.lower():
                    month_col = c
                    break
        if month_col is None:
            month_col = summary_seg.columns[0]

        value_cols = [
            c for c in summary_seg.columns
            if c != month_col and pd.api.types.is_numeric_dtype(summary_seg[c])
        ]

        if not value_cols:
            st.info("No numeric metric columns were found in the Summary sheet to trend.")
        else:
            trend_df = summary_seg[[month_col] + value_cols].copy()
            sort_key = pd.to_datetime(trend_df[month_col], errors="coerce")
            if sort_key.notna().sum() >= max(1, int(len(trend_df) * 0.5)):
                trend_df = trend_df.assign(_sort=sort_key).sort_values("_sort").drop(columns="_sort")

            metric_choices = st.multiselect(
                "Metrics to plot", value_cols, default=value_cols[: min(4, len(value_cols))],
            )
            plot_cols = metric_choices or value_cols
            melted = trend_df.melt(id_vars=month_col, value_vars=plot_cols,
                                    var_name="Metric", value_name="Value")
            melted[month_col] = melted[month_col].astype(str)
            fig = px.line(melted, x=month_col, y="Value", color="Metric", markers=True,
                          title="Monthly Trend", color_discrete_sequence=PALETTE)
            fig.update_layout(xaxis_title=month_col, yaxis_title="Value")
            st.plotly_chart(fig, use_container_width=True)

            with st.expander("View Summary sheet data"):
                st.dataframe(summary_seg, use_container_width=True, height=320)
