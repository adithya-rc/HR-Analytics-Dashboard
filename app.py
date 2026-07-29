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
# Sidebar — Vertical filter
# ==========================================================================

vertical_values = set()
for d in [emp, notice_df, pip_df, exits_df, summary_df]:
    if d.empty:
        continue
    vcol = find_col(d, "vertical")
    if vcol:
        vertical_values.update(d[vcol].dropna().astype(str).tolist())
all_verticals = sorted(vertical_values)

st.sidebar.header("Vertical")
selected_verticals = st.sidebar.multiselect(
    "Vertical / Business Unit", all_verticals, key="filter_vertical",
    help="Leave empty to include all verticals.",
)

def apply_vertical(tdf):
    if tdf.empty or not selected_verticals:
        return tdf
    col = find_col(tdf, "vertical")
    return tdf[tdf[col].astype(str).isin(selected_verticals)] if col else tdf

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

# ==========================================================================
# Sidebar — Report Period (always-visible custom dates)
# ==========================================================================

st.sidebar.header("Report Period")
preset = st.sidebar.selectbox(
    "Quick range",
    ["All Time", "Today", "Last 7 Days", "Last 30 Days", "This Month", "Last Month", "Custom"],
    key="period_preset",
)

today = datetime.date.today()
if preset == "Today":
    default_start, default_end = today, today
elif preset == "Last 7 Days":
    default_start, default_end = today - datetime.timedelta(days=6), today
elif preset == "Last 30 Days":
    default_start, default_end = today - datetime.timedelta(days=29), today
elif preset == "This Month":
    default_start, default_end = today.replace(day=1), today
elif preset == "Last Month":
    last_month_end = today.replace(day=1) - datetime.timedelta(days=1)
    default_start, default_end = last_month_end.replace(day=1), last_month_end
elif preset == "Custom":
    default_start, default_end = today - datetime.timedelta(days=6), today
else:
    default_start = default_end = None

if preset == "All Time":
    period_start = period_end = None
else:
    c1, c2 = st.sidebar.columns(2)
    period_start = c1.date_input("Start date", value=default_start, key=f"period_start_{preset}")
    period_end = c2.date_input("End date", value=default_end, key=f"period_end_{preset}")
    if period_start > period_end:
        st.sidebar.error("Start date is after end date.")

def overlap_in_period(tdf, start_col, end_col):
    if tdf.empty or period_start is None or not start_col or start_col not in tdf.columns:
        return tdf
    s = tdf[start_col]
    e = tdf[end_col] if end_col and end_col in tdf.columns else pd.Series(pd.NaT, index=tdf.index)
    cond = s.notna() & (s.dt.date <= period_end) & (e.isna() | (e.dt.date >= period_start))
    return tdf[cond]

def point_in_period(tdf, date_col):
    if tdf.empty or period_start is None or not date_col or date_col not in tdf.columns:
        return tdf
    d = tdf[date_col]
    return tdf[d.notna() & (d.dt.date >= period_start) & (d.dt.date <= period_end)]

notice_seg = filter_tracker(notice_df)
notice_f = overlap_in_period(notice_seg, find_col(notice_seg, "dor"), find_col(notice_seg, "lwd"))

pip_seg = filter_tracker(pip_df)
pip_seg = apply_pip_status(pip_seg)
pip_start_col = find_col(pip_seg, "pip_start") or find_col(pip_seg, "dor")
pip_end_col = find_col(pip_seg, "pip_end") or find_col(pip_seg, "lwd")
pip_f = overlap_in_period(pip_seg, pip_start_col, pip_end_col)

exits_seg = filter_tracker(exits_df)
exits_event_col = find_col(exits_seg, "lwd") or find_col(exits_seg, "dor")
exits_f = point_in_period(exits_seg, exits_event_col)

exit_col = roles.get("exit_date")
exited_all = df[df[exit_col].notna()] if exit_col and exit_col in df.columns else pd.DataFrame()
exited_period = point_in_period(exited_all, exit_col) if period_start else exited_all

st.sidebar.divider()
st.sidebar.caption("Build: v6.0 — KPIs, PIP status filter, dept attrition, trends")

# ==========================================================================
# Core figures (kept for use in section headers/captions below — the old
# 9-box KPI strip was removed since it would silently show 0/— whenever a
# role column wasn't detected, which read as "broken" even when filters
# were working fine).
# ==========================================================================

hc = len(df)
attrition_count = len(exited_period)
attrition_rate = pct(attrition_count, hc)
period_label = f" ({preset})" if period_start else ""

# Active headcount = filtered employees with no recorded exit date
exited_mask = df[exit_col].notna() if exit_col and exit_col in df.columns else pd.Series(False, index=df.index)
active_hc = int((~exited_mask).sum())

# Average tenure (years), open-ended employees measured to today
doj_col = roles.get("doj")
avg_tenure_years = None
if doj_col and doj_col in df.columns:
    end_dates = df[exit_col].copy() if exit_col and exit_col in df.columns else pd.Series(pd.NaT, index=df.index)
    end_dates = end_dates.fillna(pd.Timestamp(today))
    tenure_days = (end_dates - df[doj_col]).dt.days
    tenure_days = tenure_days[tenure_days.notna() & (tenure_days >= 0)]
    if not tenure_days.empty:
        avg_tenure_years = round(tenure_days.mean() / 365.25, 1)

# New hires in the selected period, and net headcount change
new_hires_df = point_in_period(df, doj_col) if doj_col else pd.DataFrame()
new_hires = len(new_hires_df)
net_change = new_hires - attrition_count

st.divider()

# ==========================================================================
# Single-page layout — every section stacked vertically instead of tabs,
# so everything is visible with one scroll instead of clicking around.
# Each section only renders its metric line for values that are actually
# available, instead of a fixed 9-box grid that shows 0/— for anything
# undetected.
# ==========================================================================

tab_overview = st.container()
st.divider()
tab_reasons = st.container()
st.divider()
tab_np = st.container()
st.divider()
tab_exits = st.container()
st.divider()
tab_trends = st.container()

# ==========================================================================
# SECTION: Overview — Gender diversity + headcount breakdowns
# ==========================================================================

with tab_overview:
    st.header("🏠 Overview")

    glance_bits = [f"**Headcount:** {hc}", f"**Active:** {active_hc}",
                    f"**Attrition{period_label}:** {attrition_rate}% ({attrition_count})"]
    if avg_tenure_years is not None:
        glance_bits.append(f"**Avg tenure:** {avg_tenure_years} yrs")
    if doj_col:
        glance_bits.append(f"**New hires{period_label}:** {new_hires}")
        glance_bits.append(f"**Net change{period_label}:** {net_change:+d}")
    if not notice_df.empty:
        glance_bits.append(f"**On notice:** {len(notice_f)}")
    if not pip_df.empty:
        glance_bits.append(f"**On PIP:** {len(pip_f)}")
    st.caption(" · ".join(glance_bits))

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
    st.header("📌 Reasons & Managers")
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
    st.header("📋 Notice & PIP")
    if period_start:
        st.caption(f"Overlapping **{preset}** ({period_start} to {period_end}).")
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

with tab_exits:
    st.header("🚪 Exits")
    if period_start:
        st.caption(f"Exits within **{preset}** ({period_start} to {period_end}).")
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

with tab_trends:
    st.header("📈 Trends")
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
