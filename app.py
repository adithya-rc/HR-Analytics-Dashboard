"""
ML HR Dashboard
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

st.set_page_config(page_title="ML HR Dashboard", layout="wide", page_icon="📊")

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

st.title("📊 ML HR Dashboard")

# ==========================================================================
# Loading
# ==========================================================================

def _detect_header_row(file_bytes: bytes, sheet_name: str, max_scan: int = 10) -> int:
    """Some exported workbooks have one or more title rows (e.g. a sheet
    name repeated as a cell) before the real column headers. Reading with
    header=0 in that case turns every column into 'Unnamed: N' and the
    sheet becomes unusable for role-detection. Scan the first few rows and
    pick the first one that looks like a real header: mostly non-blank
    across the row, unlike a single-cell title row."""
    try:
        raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, nrows=max_scan)
    except Exception:
        return 0
    n_cols = raw.shape[1]
    if n_cols == 0:
        return 0
    threshold = max(2, int(n_cols * 0.6))
    for i in range(len(raw)):
        if raw.iloc[i].notna().sum() >= threshold:
            return i
    return 0


@st.cache_data(show_spinner=False)
def load_all_sheets(file_bytes: bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    out = {}
    for s in xl.sheet_names:
        try:
            header_row = _detect_header_row(file_bytes, s)
            d = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=header_row)
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
    "id":            ["emp id", "employee id", "empid", "employee number", "employee no", "emp no", "emp code", "employee code"],
    "name":          ["employee name", "emp name", "full name"],
    "gender":        ["gender", "sex"],
    "doj":           ["doj", "date of joining", "joining date", "hire date"],
    "dob":           ["date of birth", "dob", "birth date", "birthday"],
    "exit_date":     ["date of exit", "exit date", "dor", "lwd", "last working day", "separation date"],
    "status":        ["employee status", "employment status", "status"],
    "department":    ["department", "dept"],
    "sub_department": ["sub department", "sub-department", "subdepartment"],
    "division":      ["division"],
    "vertical":      ["vertical", "business unit"],
    "location":      ["location", "office", "site", "city"],
    # Most specific first: "L1 Manager"/"Reporting Manager" should win over
    # a merely-generic "manager" substring match (which could otherwise
    # land on something like "Dotted Line Manager" or "Skip Level Manager"
    # if one happens to appear earlier in the column order).
    "manager":       ["l1 manager", "reporting manager", "manager", "reporting to"],
    "designation":   ["designation", "title", "position"],
    "reason":        ["reason"],
    "attrition_type": ["attrition type"],
    "tenure_days":   ["tenure[days]", "tenure days", "tenure (days)"],
    "dor":           ["dor"],
    "lwd":           ["lwd", "last working day"],
    "pip_start":     ["pip start date"],
    "pip_end":       ["pip revoke date"],
    "pip_status":    ["review status", "pip status", "pip outcome", "outcome"],
    "notice_date":   ["notice period", "notice date", "date of notice"],
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


def first_date_col(df, exclude=()):
    if df is None or df.empty:
        return None
    for c in df.columns:
        if c in exclude:
            continue
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            return c
    return None


def clickable_chart(fig, source_df, category_col, key, chart_type="bar"):
    """Render a Plotly chart with click-to-drill-down: clicking a bar or
    pie slice shows the underlying records for that category right below
    the chart. Click the same bar/slice again to clear the table."""
    event = st.plotly_chart(fig, use_container_width=True, key=key,
                             on_select="rerun", selection_mode="points")
    points = (event or {}).get("selection", {}).get("points", [])
    if points:
        pt = points[0]
        # Plotly's selection payload varies by trace type; try the keys
        # most likely to hold the clicked category name, in order.
        candidate_keys = ["label", "y", "x"] if chart_type == "pie" else ["y", "label", "x"]
        clicked_value = next((pt[k] for k in candidate_keys if pt.get(k) is not None), None)
        if clicked_value is not None and category_col and source_df is not None and category_col in source_df.columns:
            matches = source_df[source_df[category_col].astype(str) == str(clicked_value)]
            st.caption(f"🔎 {len(matches)} record(s) where **{category_col} = {clicked_value}** "
                        f"— click the same bar/slice again to clear.")
            st.dataframe(matches, use_container_width=True,
                         height=min(350, max(120, 38 * min(len(matches), 8) + 40)))


def clickable_table(display_df, source_df, category_col, key, height=380):
    """Render a summary table indexed by a category (e.g. Reporting Manager
    or Department) with click-to-drill-down: selecting a row shows the
    underlying employee records for that category right below the table.
    Click the same row again to clear."""
    event = st.dataframe(display_df, use_container_width=True, height=height, key=key,
                          on_select="rerun", selection_mode="single-row")
    sel_rows = (event or {}).get("selection", {}).get("rows", [])
    if sel_rows and category_col and source_df is not None and category_col in source_df.columns:
        clicked_value = display_df.index[sel_rows[0]]
        matches = source_df[source_df[category_col].astype(str) == str(clicked_value)]
        st.caption(f"🔎 {len(matches)} record(s) where **{category_col} = {clicked_value}** "
                    f"— click the same row again to clear.")
        st.dataframe(matches, use_container_width=True,
                     height=min(350, max(120, 38 * min(len(matches), 8) + 40)))


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
# "Backfill"/"Dropdown" sheets are already recognized by TRACKER_KW (so they
# never get mistaken for the Employee Master), but were never actually
# loaded anywhere. If present, they typically carry a full ID -> Vertical
# mapping for the WHOLE employee population (active + exited) — unlike
# Notice/PIP/Exit, which only cover people who are already leaving or at
# risk. Without this, the vertical hard-lock below could only "confirm" an
# employee via Notice/PIP/Exit records, which silently skewed every
# breakdown toward already-exited people (see id_vertical_map below).
backfill_df = locate_sheet("backfill")
dropdown_df = locate_sheet("dropdown")

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
    backfill_df = locate_sheet("backfill") if "backfill" not in main_key.lower() else pd.DataFrame()
    dropdown_df = locate_sheet("dropdown") if "dropdown" not in main_key.lower() else pd.DataFrame()

# ==========================================================================
# Derive Notice / PIP / Exit-tracker views from the Employee Master itself
# when no dedicated sheet was uploaded for them. Many workbooks (like this
# one) track "on notice" / "on PIP" as VALUES in an Employee Status column
# on one combined sheet rather than as separate sheets — without this, the
# Notice & PIP tab and the Exit Reasons chart would sit empty even though
# the underlying data is right there.
# ==========================================================================

_status_col = roles.get("status")
if notice_df.empty and _status_col and _status_col in emp.columns:
    _notice_mask = emp[_status_col].astype(str).str.contains("notice", case=False, na=False)
    if _notice_mask.any():
        notice_df = emp[_notice_mask].copy()

if pip_df.empty and _status_col and _status_col in emp.columns:
    _pip_mask = emp[_status_col].astype(str).str.contains(r"\bpip\b", case=False, na=False, regex=True)
    if _pip_mask.any():
        pip_df = emp[_pip_mask].copy()

if exits_df.empty:
    _exit_col_probe = roles.get("exit_date")
    if _exit_col_probe and _exit_col_probe in emp.columns:
        _exit_mask = emp[_exit_col_probe].notna()
        if _exit_mask.any():
            exits_df = emp[_exit_mask].copy()


# ==========================================================================
# Sidebar — Vertical filter (locked to Motivity Labs only — every table,
# chart, and download in this app is restricted to this business unit;
# other verticals are excluded entirely, not just unselected.)
# ==========================================================================

FIXED_VERTICAL = "Motivity Labs"

vertical_values = set()
for d in [backfill_df, dropdown_df, emp, notice_df, pip_df, exits_df, summary_df]:
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
for d in [backfill_df, dropdown_df, emp, notice_df, pip_df, exits_df, summary_df]:
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
    if backfill_df.empty and dropdown_df.empty:
        st.sidebar.caption(
            "⚠️ Your Employee Master sheet has no Vertical/Business Unit column of its own, and no "
            "Backfill/Dropdown mapping sheet was found either, so it can only confirm an employee as "
            "Motivity Labs by cross-referencing their Employee ID against Notice, PIP, or Exit records. "
            "Active employees who've never appeared on any of those (which is most of your active "
            "headcount) can't be confirmed and are excluded under this hard lock — this is why Headcount "
            "and Attrition Count can look identical in the breakdown tables below. Add a Vertical column "
            "to the Employee Master, or include a Backfill/Dropdown mapping sheet with the full ID→Vertical "
            "list, to include your full active population accurately."
        )
    else:
        st.sidebar.caption(
            "ℹ️ Your Employee Master sheet has no Vertical/Business Unit column of its own, so verticals are "
            "confirmed by cross-referencing Employee ID against your Backfill/Dropdown mapping sheet "
            "(preferred, full population) plus Notice/PIP/Exit records as a fallback."
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
    # Notice/PIP/Exit/Backfill are kept. Anything unconfirmed is excluded —
    # no benefit-of-the-doubt inclusion, per explicit instruction to only
    # ever show Motivity Labs data and nothing else.
    idc = find_col(tdf, "id")
    if idc and id_vertical_map:
        ids_series = tdf[idc].astype(str).str.strip()
        confirmed_ids = {i for i, v in id_vertical_map.items() if v in selected_verticals}
        return tdf[ids_series.isin(confirmed_ids)]
    if idc:
        # Has an ID column but nothing at all to cross-reference against —
        # still per-employee data, so stay cautious and exclude.
        return tdf.iloc[0:0]
    # No Vertical column AND no ID column either — this isn't a per-employee
    # row list at all (e.g. an already-aggregated monthly Summary sheet with
    # one row per month, not per person). There's no per-row "other vertical"
    # leakage risk here and nothing to cross-reference, so pass it through
    # unchanged instead of nuking a sheet that was never going to be scoped
    # by vertical in the first place.
    return tdf

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
pip_start_col = find_col(pip_df, "pip_start")
pip_end_col = find_col(pip_df, "pip_end")
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


def _pip_in_progress(tdf):
    """PIP rows that are actually still ongoing RIGHT NOW, per the Review
    Status column — this is what 'Currently on PIP' should mean, not every
    PIP record ever raised (which would include ones already Completed/
    Passed/Failed/Closed). Use _pip_active_as_of / _pip_active_in_range
    instead for any historical ("as of a date" / "during a range")
    question — status alone has no memory of when a PIP was actually open."""
    if tdf.empty or not pip_status_col or pip_status_col not in tdf.columns:
        return tdf
    in_progress_mask = tdf[pip_status_col].astype(str).str.contains("progress", case=False, na=False)
    if in_progress_mask.any():
        return tdf[in_progress_mask]
    # No status explicitly says "In Progress" — fall back to excluding
    # anything that reads as a closed/terminal outcome, so a workbook using
    # different wording (e.g. "Ongoing", "Active") doesn't just come back empty.
    closed_kw = "complete|closed|pass|fail|terminat|exit|revoke|end"
    closed_mask = tdf[pip_status_col].astype(str).str.contains(closed_kw, case=False, na=False, regex=True)
    return tdf[~closed_mask]


def _pip_effective_end(tdf):
    """Each row's PIP end date for interval math: the actual Revoke/End
    Date if recorded, otherwise 'still open' — represented as today, since
    a PIP with no recorded end hasn't been closed out yet."""
    today_ts = pd.Timestamp(datetime.date.today())
    if pip_end_col and pip_end_col in tdf.columns:
        return tdf[pip_end_col].fillna(today_ts)
    return pd.Series(today_ts, index=tdf.index)


def _pip_active_as_of(tdf, as_of_date):
    """Was this person actively on PIP on a specific date? Proper interval
    check: PIP Start Date <= as_of_date <= (Revoke Date or today, if still
    open). Falls back to the current Review Status if there's no PIP Start
    Date at all to build an interval from — which can only answer 'who is
    on PIP right now', not a historical date."""
    if tdf.empty:
        return tdf
    if not pip_start_col or pip_start_col not in tdf.columns:
        return _pip_in_progress(tdf)
    as_of_ts = pd.Timestamp(as_of_date)
    started_by = tdf[pip_start_col] <= as_of_ts
    not_yet_ended = _pip_effective_end(tdf) >= as_of_ts
    return tdf[started_by & not_yet_ended]


def _pip_active_in_range(tdf, range_start, range_end):
    """Was this person on PIP at ANY point during [range_start, range_end]?
    A true interval-overlap check — catches PIPs that started before the
    range but are still running through it, not just ones that happened to
    start inside the range (which was the previous, narrower behaviour).
    Falls back to current Review Status if there's no PIP Start Date."""
    if tdf.empty or range_start is None:
        return tdf
    if not pip_start_col or pip_start_col not in tdf.columns:
        return _pip_in_progress(tdf)
    range_start_ts, range_end_ts = pd.Timestamp(range_start), pd.Timestamp(range_end)
    overlaps = (tdf[pip_start_col] <= range_end_ts) & (_pip_effective_end(tdf) >= range_start_ts)
    return tdf[overlaps]

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
pip_seg_all_status = pip_seg
pip_seg = apply_pip_status(pip_seg)
pip_f = _pip_in_progress(pip_seg)

exits_seg = filter_tracker(exits_df)
exits_event_col = find_col(exits_seg, "lwd") or find_col(exits_seg, "dor")

exit_col = roles.get("exit_date")
doj_col = roles.get("doj")
status_col = roles.get("status")


def is_resigned_mask(tdf):
    """Authoritative 'has this employee actually exited' signal. Per the
    Employee Status field: status == 'Resigned' (or any status containing
    that word) means exited; every OTHER status — Active, Probation,
    Intern, PIP, On Notice Period, Consultant, etc. — counts as current
    headcount, regardless of whether an exit date happens to be filled in.
    Falls back to exit-date presence only when there's no status column at
    all to check against."""
    if status_col and status_col in tdf.columns:
        return tdf[status_col].astype(str).str.contains("resign", case=False, na=False)
    if exit_col and exit_col in tdf.columns:
        return tdf[exit_col].notna()
    return pd.Series(False, index=tdf.index)


VOLUNTARY_REASON_KEYWORDS = [
    "better opportunity", "personal reason", "health issue", "higher education",
    "relocation", "work life balance", "worklife balance", "moving abroad",
    "moved to us", "job dissatisfaction", "h1b transfer", "us transfer",
    "resignation", "resigned",
]
INVOLUNTARY_REASON_KEYWORDS = [
    "performance issue", "non-performance", "non performance", "pip",
    "termination", "terminated", "re-org", "reorg", "absconded", "bgv failure",
    "background verification", "layoff", "redundan", "death", "proxy",
    "contract end",
]


def classify_exit_type(tdf):
    """Voluntary / Involuntary / Unclassified for each row. Prefers the
    sheet's own Attrition Type column when present — that's HR-curated and
    per-row, so it correctly captures cases where the same reason text (e.g.
    'InterCompany Transfer') is Voluntary in some cases and Involuntary in
    others, which no reason-text rule could ever get right. Falls back to a
    conservative keyword match against the Reason column only when there's
    no Attrition Type column at all; anything unrecognized stays
    'Unclassified' rather than being silently guessed."""
    if tdf.empty:
        return pd.Series(dtype=object)
    attrition_type_col = roles.get("attrition_type")
    if attrition_type_col and attrition_type_col in tdf.columns and tdf[attrition_type_col].notna().any():
        return tdf[attrition_type_col].fillna("Unclassified")
    r_col = find_col(tdf, "reason")
    if not r_col or r_col not in tdf.columns:
        return pd.Series("Unclassified", index=tdf.index)

    def _classify(val):
        s = str(val).lower()
        if any(kw in s for kw in INVOLUNTARY_REASON_KEYWORDS):
            return "Involuntary"
        if any(kw in s for kw in VOLUNTARY_REASON_KEYWORDS):
            return "Voluntary"
        return "Unclassified"

    return tdf[r_col].apply(_classify)


def _resigned_in_range(tdf, range_start, range_end):
    """Resigned rows only, further scoped to those whose exit date falls in
    the given range."""
    if tdf.empty:
        return tdf
    resigned = tdf[is_resigned_mask(tdf)]
    return _point_in_range(resigned, exit_col, range_start, range_end)


def _active_as_of(tdf, as_of_date):
    """True point-in-time active roster on a given date: joined on/before
    that date, and not (Resigned with an exit date on/before that date).
    Unlike df_period (event-window) or 'currently Resigned' (today's live
    status), this reflects who was actually on the books on that specific
    day — the stable, validated number for a recurring headcount check."""
    if tdf.empty or not doj_col or not exit_col or doj_col not in tdf.columns or exit_col not in tdf.columns:
        return tdf.iloc[0:0]
    as_of_ts = pd.Timestamp(as_of_date)
    joined_by = tdf[doj_col] <= as_of_ts
    exited_by = is_resigned_mask(tdf) & (tdf[exit_col] <= as_of_ts)
    return tdf[joined_by & ~exited_by]


def _last_friday_on_or_before(a_date):
    return a_date - datetime.timedelta(days=(a_date.weekday() - 4) % 7)


exited_all = df[is_resigned_mask(df)] if not df.empty else pd.DataFrame()

# ==========================================================================
# Sidebar — Master Date Filter. Takes precedence over Joining Period and
# Exit Period below (both get locked to this range while it's on), and also
# scopes Notice/PIP counts and the attrition-rate metric to the same window.
# Turn this on once to make "exits, joiners, notice, PIP, attrition %" all
# reflect one selected window (e.g. the last 7 days) at the same time.
# ==========================================================================

st.sidebar.header("🗓️ Master Date Filter")
use_master_period = st.sidebar.checkbox(
    "Use one master date range for the whole dashboard",
    key="use_master_period",
    help="When on, this single range overrides Exit Period and Joining Period below, and also scopes "
         "Notice/PIP counts and the attrition-rate metric to the same window.",
)

master_period_start = master_period_end = None
if use_master_period:
    master_dates_avail = pd.concat([
        df[doj_col].dropna() if doj_col and doj_col in df.columns else pd.Series(dtype="datetime64[ns]"),
        df[exit_col].dropna() if exit_col and exit_col in df.columns else pd.Series(dtype="datetime64[ns]"),
    ])
    master_min = master_dates_avail.min().date() if not master_dates_avail.empty else today
    master_max = master_dates_avail.max().date() if not master_dates_avail.empty else today
    # Explicit wide bounds, same reasoning as Joining/Exit Period below.
    master_bound_min = min(master_min, datetime.date(1990, 1, 1))
    master_bound_max = max(master_max, today)

    preset = st.sidebar.selectbox(
        "Quick range", ["Last 7 days", "Last 30 days", "This month", "This quarter", "This year", "Custom"],
        index=0, key="master_preset",
    )
    if preset == "Last 7 days":
        preset_start, preset_end = today - datetime.timedelta(days=6), today
    elif preset == "Last 30 days":
        preset_start, preset_end = today - datetime.timedelta(days=29), today
    elif preset == "This month":
        preset_start, preset_end = today.replace(day=1), today
    elif preset == "This quarter":
        q_start_month = 3 * ((today.month - 1) // 3) + 1
        preset_start, preset_end = datetime.date(today.year, q_start_month, 1), today
    elif preset == "This year":
        preset_start, preset_end = datetime.date(today.year, 1, 1), today
    else:  # Custom
        preset_start, preset_end = master_min, master_max

    mc1, mc2 = st.sidebar.columns(2)
    # Key includes the preset name: switching presets must show that
    # preset's own dates. With a fixed key, Streamlit keeps whatever the
    # widget last held and silently ignores the new `value=`, which is why
    # picking a different Quick Range appeared to do nothing. "Custom"
    # keeps one stable key so manual edits persist while you stay on it.
    _preset_key = preset.replace(" ", "_").lower()
    master_period_start = mc1.date_input("Range start", value=preset_start,
                                          min_value=master_bound_min, max_value=master_bound_max,
                                          key=f"master_period_start_{_preset_key}")
    master_period_end = mc2.date_input("Range end", value=preset_end,
                                        min_value=master_bound_min, max_value=master_bound_max,
                                        key=f"master_period_end_{_preset_key}")

    if master_period_start > master_period_end:
        st.sidebar.error("Master range start date is after end date.")
        master_period_start = master_period_end = None
    else:
        st.sidebar.success(f"Master range active: **{master_period_start} → {master_period_end}**")
        # Notice and PIP are both live/current statuses — neither should
        # shrink or disappear just because an unrelated date range is
        # selected elsewhere. notice_f and pip_f intentionally stay as set
        # above (full current on-notice list; current In Progress PIPs),
        # un-scoped by the Master Date Filter. For a historical "who was on
        # PIP as of a past date" question, use the Data Snapshot section's
        # as-of-date control instead — that's built for exactly this.

use_master_active = use_master_period and master_period_start is not None and master_period_start <= master_period_end

# df_period: the population "belonging to" this window — employees who
# JOINED within the range, or who exited within the range (Resigned with an
# exit date inside it). This is what Headcount should mean whenever it's
# shown next to a period-scoped Attrition Count (mgr/dept breakdown
# tables), and what the Overview tab (Employee List, Gender Diversity,
# Department chart) should show while the Master Date Filter is on.
# Deliberately NOT "anyone still employed during this window" — that would
# pull in every long-tenured employee who joined years ago and simply
# hasn't left, which isn't "this period's data." Falls back to the full
# filtered df when the master filter is off, or when DOJ/Exit-date columns
# aren't available to compute it.
if use_master_active and doj_col and exit_col and doj_col in df.columns and exit_col in df.columns:
    _range_start_ts = pd.Timestamp(master_period_start)
    _range_end_ts = pd.Timestamp(master_period_end) + pd.Timedelta(days=1)
    _joined_in_window = (df[doj_col] >= _range_start_ts) & (df[doj_col] < _range_end_ts)
    _is_resigned = is_resigned_mask(df)
    _exited_in_window = _is_resigned & (df[exit_col] >= _range_start_ts) & (df[exit_col] < _range_end_ts)
    df_period = df[_joined_in_window | _exited_in_window]
    df_period_is_exit = _exited_in_window.loc[df_period.index]
else:
    df_period = df
    df_period_is_exit = pd.Series(False, index=df_period.index)

# ==========================================================================
# Sidebar — Joining Period (dedicated joining-date range — answers "how
# many employees joined between X and Y"). Locked to the Master Date
# Filter's range whenever that's switched on.
# ==========================================================================

st.sidebar.header("Joining Period")
joining_period_start = joining_period_end = None
joined_custom = pd.DataFrame()
if use_master_active:
    st.sidebar.caption("Controlled by the Master Date Filter above.")
    use_joining_period = True
    joining_period_start, joining_period_end = master_period_start, master_period_end
    if doj_col and doj_col in df.columns:
        joined_custom = _point_in_range(df, doj_col, joining_period_start, joining_period_end)
    st.sidebar.metric("Joiners in this range", len(joined_custom))
else:
    use_joining_period = st.sidebar.checkbox(
        "Filter by a specific joining date range", key="use_joining_period",
        help="Shows exactly how many employees joined within the dates you pick.",
    )

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
# exits_f, when active.) Locked to the Master Date Filter's range whenever
# that's switched on.
# ==========================================================================

st.sidebar.header("Exit Period")
exit_period_start = exit_period_end = None
exited_custom = pd.DataFrame()
exits_f = exits_seg
if use_master_active:
    st.sidebar.caption("Controlled by the Master Date Filter above.")
    use_exit_period = True
    exit_period_start, exit_period_end = master_period_start, master_period_end
    exited_custom = _resigned_in_range(df, exit_period_start, exit_period_end)
    exits_f = _point_in_range(exits_seg, exits_event_col, exit_period_start, exit_period_end)
    st.sidebar.metric("Exits in this range", len(exited_custom))
else:
    use_exit_period = st.sidebar.checkbox(
        "Filter by a specific exit date range", key="use_exit_period",
        help="Shows exactly how many employees left within the dates you pick.",
    )

    if use_exit_period:
        _resigned_only = df[is_resigned_mask(df)]
        exit_dates_avail = _resigned_only[exit_col].dropna() if exit_col and exit_col in _resigned_only.columns else pd.Series(dtype="datetime64[ns]")
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
            exited_custom = _resigned_in_range(df, exit_period_start, exit_period_end)
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
    active_mask = ~is_resigned_mask(df)
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

hc = len(df_period)

# Active headcount = a stable point-in-time count as of last Friday (not
# "today's live status", and not tied to df_period/Master Date Filter).
# Matches the number used every week for the validation meeting, so it
# doesn't drift day-to-day as new resignations get logged, and doesn't
# change just because a different Master Date Filter window is selected.
_hdr_last_friday = _last_friday_on_or_before(today)
active_hc = len(_active_as_of(df, _hdr_last_friday))

# Average tenure (years), open-ended employees measured to today
avg_tenure_years = None
if doj_col and doj_col in df.columns:
    if exit_col and exit_col in df.columns:
        end_dates = df[exit_col].where(is_resigned_mask(df))
    else:
        end_dates = pd.Series(pd.NaT, index=df.index)
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

# Attrition % — exits over average headcount for a reference window.
# Normally this is always calendar-year-to-date, independent of the Exit
# Period/Joining Period toggles. But when the Master Date Filter is active
# it takes precedence, per that filter's job of scoping *every* number on
# the dashboard (this one included) to the one selected range.
ytd_attrition_pct = None
ytd_exits = None
ytd_exits_df = pd.DataFrame()
ytd_is_master_scoped = use_master_active
if doj_col and exit_col and doj_col in df.columns and exit_col in df.columns:
    if use_master_active:
        ytd_start, ytd_end = master_period_start, master_period_end
    else:
        ytd_start, ytd_end = datetime.date(today.year, 1, 1), today

    _is_resigned_df = is_resigned_mask(df)

    def _headcount_as_of(as_of_date):
        as_of_ts = pd.Timestamp(as_of_date)
        joined = df[doj_col] <= as_of_ts
        exited_by_then = _is_resigned_df & (df[exit_col] <= as_of_ts)
        return int((joined & ~exited_by_then).sum())

    hc_start_of_year = _headcount_as_of(ytd_start - datetime.timedelta(days=1))
    hc_today = _headcount_as_of(ytd_end)
    ytd_avg_hc = (hc_start_of_year + hc_today) / 2

    ytd_exits_df = _resigned_in_range(df, ytd_start, ytd_end)
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

header_items = [("Headcount", hc, "overview"), (f"Active (as of {_hdr_last_friday})", active_hc, "overview")]
if avg_tenure_years is not None:
    header_items.append(("Avg tenure", f"{avg_tenure_years} yrs", "overview"))
if early_attrition_count is not None:
    header_items.append(("Early attrition (≤6mo)", f"{early_attrition_count} ({early_attrition_pct_of_exits}% of exits)", "reasons"))
if ytd_attrition_pct is not None:
    attr_label = "Attrition % (Range)" if ytd_is_master_scoped else "Attrition % YTD"
    header_items.append((attr_label, f"{ytd_attrition_pct}% ({ytd_exits})", "reasons"))
if not notice_df.empty:
    header_items.append(("On notice", len(notice_f), "notice-pip"))
if not pip_df.empty:
    header_items.append(("PIP", len(pip_f), "notice-pip"))
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

tab_snapshot = st.container()
st.divider()
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

with tab_snapshot:
    st.header("📋 Weekly Data Snapshot", anchor="snapshot")

    _default_friday = _last_friday_on_or_before(today)
    snap_c1, snap_c2 = st.columns([1, 2])
    with snap_c1:
        as_of_date = st.date_input(
            "Snapshot as-of date", value=_default_friday,
            help="Active Headcount is calculated as of this date. Defaults to the most recent Friday, "
                 "so opening this on a Tuesday needs no adjustment.",
            key="snapshot_as_of",
        )
    week_start = as_of_date - datetime.timedelta(days=6)
    with snap_c2:
        st.caption(f"'Last week' = **{week_start} → {as_of_date}** (the 7 days ending on the as-of date). "
                    f"Move the as-of date above if your reporting week runs differently.")

    if doj_col and exit_col and doj_col in df.columns and exit_col in df.columns:
        snap_active_df = _active_as_of(df, as_of_date)
        snap_active_hc = len(snap_active_df)
        snap_exits_df = _resigned_in_range(df, week_start, as_of_date)
        snap_joiners_df = _point_in_range(df, doj_col, week_start, as_of_date)
    else:
        snap_active_hc = None
        snap_exits_df = pd.DataFrame()
        snap_joiners_df = pd.DataFrame()

    snap_pip_df = _pip_in_progress(pip_seg_all_status)
    _pip_label = "On PIP (In Progress)" if pip_status_col else "On PIP (no Review Status column found)"

    m1, m2, m3, m4 = st.columns(4)
    m1.metric(f"Active Headcount (as of {as_of_date})", snap_active_hc if snap_active_hc is not None else "—")
    m2.metric(_pip_label, len(snap_pip_df))
    m3.metric("Exits last week", len(snap_exits_df))
    m4.metric("New Joiners last week", len(snap_joiners_df))

    with st.expander(f"Exits last week ({week_start} → {as_of_date}) — {len(snap_exits_df)}"):
        if snap_exits_df.empty:
            st.info("No exits in this window.")
        else:
            st.dataframe(snap_exits_df, use_container_width=True, height=250)

    with st.expander(f"New Joiners last week ({week_start} → {as_of_date}) — {len(snap_joiners_df)}"):
        if snap_joiners_df.empty:
            st.info("No joiners in this window.")
        else:
            st.dataframe(snap_joiners_df, use_container_width=True, height=250)

    with st.expander(f"{_pip_label} — {len(snap_pip_df)}"):
        if snap_pip_df.empty:
            st.info("No one currently on an In Progress PIP.")
        else:
            st.dataframe(snap_pip_df, use_container_width=True, height=250)

    _snap_exits_out = snap_exits_df.copy()
    if not _snap_exits_out.empty:
        _snap_exits_out.insert(0, "Snapshot Category", "Exit (last week)")
    _snap_joiners_out = snap_joiners_df.copy()
    if not _snap_joiners_out.empty:
        _snap_joiners_out.insert(0, "Snapshot Category", "Joiner (last week)")
    _snap_pip_out = snap_pip_df.copy()
    if not _snap_pip_out.empty:
        _snap_pip_out.insert(0, "Snapshot Category", _pip_label)
    _snap_combined = pd.concat([_snap_exits_out, _snap_joiners_out, _snap_pip_out], ignore_index=True, sort=False)
    if not _snap_combined.empty:
        st.download_button(
            "Download meeting pack (CSV) — exits, joiners, PIP",
            _snap_combined.to_csv(index=False).encode("utf-8"),
            f"weekly_snapshot_{as_of_date}.csv", "text/csv",
        )

with tab_overview:
    st.header("🏠 Overview", anchor="overview")
    if use_master_active:
        st.caption(f"Master date range active: **{master_period_start} → {master_period_end}** — showing only "
                    f"employees who joined or exited within this window.")

    gender_col = roles.get("gender")
    dept_col = roles.get("department")
    loc_col = roles.get("location")
    active_df_period = df_period[~is_resigned_mask(df_period)]

    cols_avail = [c for c in [gender_col, dept_col, loc_col] if c and clean_series(df_period, c)]
    if cols_avail:
        cols_ui = st.columns(len(cols_avail))
        for i, col in enumerate(cols_avail):
            if col == gender_col:
                vc = df_period[col].value_counts().head(10)
                fig = px.pie(values=vc.values, names=vc.index, hole=0.5, title="Gender Diversity",
                             color_discrete_sequence=PALETTE)
                # Plotly's default legend sits to the right of the chart,
                # which — in a narrow Streamlit column — can visually spill
                # into the NEXT column's chart, making it look like that
                # chart's legend belongs to it. Anchor it below the donut
                # instead, fully contained within this chart's own space.
                fig.update_layout(legend=dict(orientation="h", yanchor="top", y=-0.05,
                                               xanchor="center", x=0.5))
                _click_source, _click_type = df_period, "pie"
            elif col == dept_col and gender_col and gender_col in active_df_period.columns:
                # Active employees only, bifurcated by gender (stacked bar),
                # using the same colors as the Gender Diversity pie.
                dept_order = active_df_period[dept_col].value_counts().head(10).index.tolist()
                _plot_df = active_df_period[active_df_period[dept_col].isin(dept_order)]
                gender_vals = _plot_df[gender_col].dropna().unique().tolist()
                male_like = [v for v in gender_vals if "female" not in str(v).lower() and "male" in str(v).lower()]
                other_vals = [v for v in gender_vals if v not in male_like]
                gender_order = male_like + other_vals
                color_map = {v: PALETTE[j % len(PALETTE)] for j, v in enumerate(gender_order)}
                fig = px.bar(
                    _plot_df.groupby([dept_col, gender_col]).size().reset_index(name="Headcount"),
                    x="Headcount", y=dept_col, color=gender_col, orientation="h",
                    category_orders={dept_col: dept_order, gender_col: gender_order},
                    color_discrete_map=color_map, title="Department (Active Employees)",
                )
                fig.update_layout(yaxis_title=None, xaxis_title="Headcount", legend_title_text="")
                _click_source, _click_type = active_df_period, "bar"
            else:
                vc = df_period[col].value_counts().head(10)
                fig = px.bar(vc, orientation="h", title=col, color_discrete_sequence=PALETTE)
                fig.update_layout(showlegend=False, yaxis_title=None, xaxis_title="Headcount")
                _click_source, _click_type = df_period, "bar"
            with cols_ui[i]:
                clickable_chart(fig, _click_source, col, key=f"ov_{col}", chart_type=_click_type)
    else:
        st.info("No demographic fields detected to chart.")

    st.subheader("Employee List (filtered)")
    if use_master_active:
        _df_period_display = df_period.copy()
        _df_period_display.insert(0, "Event Type", df_period_is_exit.map({True: "Exit", False: "Joiner"}))
        _df_period_display = _df_period_display.sort_values(
            "Event Type", key=lambda s: s.map({"Joiner": 0, "Exit": 1})
        )
        st.caption("Sorted with joiners first, exits last.")
        st.dataframe(_df_period_display, use_container_width=True, height=350)
        st.download_button("Download filtered data (CSV)", _df_period_display.to_csv(index=False).encode("utf-8"),
                            "filtered_employees.csv", "text/csv")
    else:
        st.dataframe(df_period, use_container_width=True, height=350)
        st.download_button("Download filtered data (CSV)", df_period.to_csv(index=False).encode("utf-8"),
                            "filtered_employees.csv", "text/csv")

with tab_reasons:
    st.header("📌 Reasons & Managers", anchor="reasons")
    if use_master_active:
        st.caption(f"Master date range active: **{master_period_start} → {master_period_end}** — Headcount below "
                    f"is joiners + exits within this window; Attrition Count/% is the exits portion of that.")

    # --- Early Attrition: employees who left within 6 months of joining ---
    if early_attrition_count is not None:
        st.subheader("Early Attrition (≤ 6 months)")
        st.metric("Left within 6 months of joining", f"{early_attrition_count} ({early_attrition_pct_of_exits}% of exits)",
                   help=f"{early_attrition_count} of the {len(exited_period)} exits in the current filter/period "
                        f"selection left within 182 days of joining.")
        st.caption(f"The **{roles.get('exit_date')}** column is what's scoped to the selected range — "
                    f"**{doj_col}** can legitimately be earlier, since this table is about people who left "
                    f"within 6 months of *their own* join date, not people who joined within the range.")
        early_attrition_display = early_attrition_df.copy()
        if doj_col in early_attrition_display.columns and exit_col in early_attrition_display.columns:
            early_attrition_display.insert(
                early_attrition_display.columns.get_loc(exit_col) + 1, "Tenure at Exit (Days)",
                (early_attrition_display[exit_col] - early_attrition_display[doj_col]).dt.days,
            )
        st.dataframe(early_attrition_display, use_container_width=True, height=300)
        st.download_button(
            "Download early leavers (CSV)",
            early_attrition_display.to_csv(index=False).encode("utf-8"),
            "early_attrition.csv", "text/csv", key="dl_early_attrition",
        )
        st.divider()

    # --- Attrition %: YTD by default, or the Master Date Filter's range when active ---
    if ytd_attrition_pct is not None:
        if ytd_is_master_scoped:
            st.subheader(f"Attrition % — {master_period_start} to {master_period_end}")
            st.metric("Attrition rate (selected range)", f"{ytd_attrition_pct}% ({ytd_exits} exits)",
                       help=f"{ytd_exits} exits from {master_period_start} to {master_period_end}, divided by average "
                            f"headcount (headcount at range start + headcount at range end, ÷ 2), following the "
                            f"Master Date Filter.")
        else:
            st.subheader(f"Attrition % — Year to Date ({today.year})")
            st.metric("YTD attrition rate", f"{ytd_attrition_pct}% ({ytd_exits} exits)",
                       help=f"{ytd_exits} exits from Jan 1, {today.year} to today, divided by average headcount "
                            f"(headcount at start of year + headcount today, ÷ 2) for the currently selected segment filters. "
                            f"Independent of the Exit Period/Joining Period toggles. Turn on the Master Date Filter "
                            f"to scope this to a custom range instead.")
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
                clickable_chart(fig, exited_period, reason_col, key="reasons_bar", chart_type="bar")
        with c2:
            type_col = roles.get("attrition_type")
            if type_col and exited_period[type_col].notna().any():
                vc = exited_period[type_col].value_counts()
                fig = px.pie(values=vc.values, names=vc.index, hole=0.5, title="Voluntary vs Involuntary",
                             color_discrete_sequence=PALETTE)
                clickable_chart(fig, exited_period, type_col, key="voluntary_pie", chart_type="pie")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Reporting Manager Breakdown")
        mgr_col = roles.get("manager")
        if mgr_col and mgr_col in df_period.columns and df_period[mgr_col].notna().any():
            hc_by_mgr = df_period[mgr_col].value_counts()
            exit_by_mgr = exited_period[mgr_col].value_counts() if not exited_period.empty else pd.Series(dtype=int)
            mgr_tbl = pd.DataFrame({"Headcount": hc_by_mgr}).join(
                pd.DataFrame({"Attrition Count": exit_by_mgr}), how="left"
            ).fillna(0)
            mgr_tbl["Attrition Count"] = mgr_tbl["Attrition Count"].astype(int)
            mgr_tbl["Attrition %"] = (mgr_tbl["Attrition Count"] / mgr_tbl["Headcount"] * 100).round(1)
            mgr_tbl = mgr_tbl.sort_values("Attrition %", ascending=False)
            mgr_tbl.index.name = mgr_col
            clickable_table(mgr_tbl, df_period, mgr_col, key="mgr_tbl")
        else:
            st.info("No reporting-manager field detected in this dataset.")

    with c2:
        dept_col_for_table = roles.get("sub_department") or roles.get("department")
        st.subheader("Department Attrition Rate")
        if dept_col_for_table and dept_col_for_table in df_period.columns and df_period[dept_col_for_table].notna().any():
            hc_by_dept = df_period[dept_col_for_table].value_counts()
            exit_by_dept = exited_period[dept_col_for_table].value_counts() if not exited_period.empty else pd.Series(dtype=int)
            dept_tbl = pd.DataFrame({"Headcount": hc_by_dept}).join(
                pd.DataFrame({"Attrition Count": exit_by_dept}), how="left"
            ).fillna(0)
            dept_tbl["Attrition Count"] = dept_tbl["Attrition Count"].astype(int)
            dept_tbl["Attrition %"] = (dept_tbl["Attrition Count"] / dept_tbl["Headcount"] * 100).round(1)
            dept_tbl = dept_tbl.sort_values("Attrition %", ascending=False)
            dept_tbl.index.name = dept_col_for_table
            clickable_table(dept_tbl, df_period, dept_col_for_table, key="dept_tbl")
        else:
            st.info("No department field detected in this dataset.")

with tab_np:
    st.header("📋 Notice & PIP", anchor="notice-pip")
    if selected_verticals:
        st.caption(f"Vertical filter active: **{', '.join(selected_verticals)}**")
    if use_master_active:
        st.caption(f"Master date range active: **{master_period_start} → {master_period_end}** — this doesn't "
                    f"affect On Notice Period or PIP below; both always show the current list (PIP: In Progress "
                    f"only). For a historical PIP question ('who was on PIP 3 months ago'), use the Data "
                    f"Snapshot section's as-of-date control instead.")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader(f"On Notice Period ({len(notice_f)})")
        if notice_f.empty:
            st.info("No notice-period records for the current selection.")
        else:
            st.dataframe(notice_f, use_container_width=True, height=420)
    with c2:
        st.subheader(f"On PIP — In Progress ({len(pip_f)})")
        if pip_f.empty:
            st.info("No PIP records for the current selection.")
        else:
            if pip_status_col and pip_status_col in pip_f.columns and pip_f[pip_status_col].notna().any():
                vc = pip_f[pip_status_col].value_counts()
                fig = px.pie(values=vc.values, names=vc.index, hole=0.5, title="PIP Review Status",
                             color_discrete_sequence=PALETTE)
                clickable_chart(fig, pip_f, pip_status_col, key="pip_status_pie", chart_type="pie")
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
    if selected_verticals:
        st.caption(f"Vertical filter active: **{', '.join(selected_verticals)}**")

    exit_type_choice = st.radio(
        "Exit Type", ["All", "Voluntary", "Involuntary"], horizontal=True, key="exit_type_filter",
        help="Voluntary/Involuntary comes from the Attrition Type column when your data has one (per-row, "
             "HR-curated — the most reliable source). If a sheet doesn't have that column, it's inferred from "
             "the Reason text instead, and anything not recognized is labeled 'Unclassified' rather than guessed.",
    )

    def _apply_exit_type(tdf):
        if tdf.empty or exit_type_choice == "All":
            return tdf
        return tdf[classify_exit_type(tdf) == exit_type_choice]

    exited_custom_view = _apply_exit_type(exited_custom)
    exited_period_view = _apply_exit_type(exited_period)
    exits_f_view = _apply_exit_type(exits_f)

    # "Exits between X and Y" and "Exited Employees — Employee Master" used
    # to both render whenever Exit Period (or the Master Date Filter, which
    # forces it on) was active — but exited_period IS exited_custom in that
    # case, so it was the exact same table shown twice. Show whichever one
    # actually applies, not both.
    if use_exit_period:
        st.subheader(f"Exits between {exit_period_start} and {exit_period_end}")
        st.metric("Employees who left in this range", len(exited_custom_view))
        if exited_custom_view.empty:
            st.info("No employees exited in the selected range.")
        else:
            st.dataframe(exited_custom_view, use_container_width=True, height=300)
            st.download_button(
                "Download exits in range (CSV)",
                exited_custom_view.to_csv(index=False).encode("utf-8"),
                "exits_in_range.csv", "text/csv", key="dl_exit_period",
            )
    else:
        st.subheader(f"Exited Employees — Employee Master ({len(exited_period_view)})")
        if exited_period_view.empty:
            st.info("No exited employees for the current selection.")
        else:
            st.dataframe(exited_period_view, use_container_width=True, height=300)
            st.download_button(
                "Download exited employees (CSV)",
                exited_period_view.to_csv(index=False).encode("utf-8"),
                "exited_employees.csv", "text/csv", key="dl_exited_period",
            )

    st.divider()

    st.subheader(f"Exit Tracker ({len(exits_f_view)})")
    if exits_f_view.empty:
        st.info("No exit-tracker records for the current selection.")
    else:
        reason_col = find_col(exits_f_view, "reason")
        if reason_col and exits_f_view[reason_col].notna().any():
            vc = exits_f_view[reason_col].value_counts().head(10)
            fig = px.bar(vc, orientation="h", title="Exit Reasons (Tracker)", color_discrete_sequence=PALETTE)
            fig.update_layout(showlegend=False, yaxis_title=None, xaxis_title="Exits")
            clickable_chart(fig, exits_f_view, reason_col, key="exit_reasons_bar", chart_type="bar")
        st.dataframe(exits_f_view, use_container_width=True, height=420)

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
