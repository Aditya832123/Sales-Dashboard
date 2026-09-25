"""Combined CRM Sales Dashboard (Streamlit).

Client controls everything from ONE Google Sheet ("Dashboard Config"):
  - tab "Sources": Name (optional) | Sheet URL | Active (optional yes/no)
  - tab "Access" : one allowed Google email per row (header in row 1)
Add a row to Sources -> the new salesperson appears on the dashboard.
"""
import json
import re
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import altair as alt
import gspread
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from google.oauth2.service_account import Credentials

COURSE_ORDER = ["MW", "DSP", "PV", "CR", "CDM", "RA", "MC", "SAS"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
PALETTE = ["#378ADD", "#D4537E", "#1D9E75", "#BA7517", "#7F77DD", "#D85A30", "#5DCAA5", "#993C1D"]


def color_for(names, name):
    return PALETTE[names.index(name) % len(PALETTE)] if name in names else "#888780"


# ---------- helpers ----------
def inr(x):
    s = str(int(round(x)))
    if len(s) > 3:
        s = re.sub(r"(\d)(?=(\d\d)+$)", r"\1,", s[:-3]) + "," + s[-3:]
    return "Rs " + s


def money(s):
    return pd.to_numeric(s.astype(str).str.replace(r"[^\d.]", "", regex=True), errors="coerce")


def clean_course(x):
    t = re.sub(r"[&/+,]", " ", str(x).upper().replace(".", ""))
    parts = t.split()
    if not parts:
        return "Unknown"
    parts = sorted(set(parts), key=lambda p: COURSE_ORDER.index(p) if p in COURSE_ORDER else 99)
    return "+".join(parts)


def clean_title(t):
    t = re.sub(r"^\s*copy of\s+", "", t, flags=re.I)
    return re.sub(r"[^\x00-\x7F]+", "", t).strip()


def find_col(df, prefix):
    return next((c for c in df.columns if c.lower().startswith(prefix)), None)


def add_periods(df):
    df["month_start"] = df.date.dt.to_period("M").dt.to_timestamp()
    df["Month"] = df.month_start.dt.strftime("%b %Y")
    df["week_start"] = df.date - pd.to_timedelta(df.date.dt.weekday, unit="D")
    df["Week"] = df.week_start.dt.strftime("%d %b") + " - " + (df.week_start + pd.Timedelta(days=6)).dt.strftime("%d %b %Y")
    df["Day"] = df.date.dt.strftime("%d %b %Y")
    return df


# ---------- Google access ----------
def sa_info():
    """Service account key: paste the whole JSON as `gcp_service_account_json`, or use a [gcp_service_account] table."""
    if "gcp_service_account_json" in st.secrets:
        return json.loads(st.secrets["gcp_service_account_json"])
    return dict(st.secrets["gcp_service_account"])


@st.cache_resource
def get_client():
    return gspread.authorize(Credentials.from_service_account_info(sa_info(), scopes=SCOPES))


@st.cache_data(ttl=60, show_spinner=False)
def load_config():
    sh = get_client().open_by_key(st.secrets["app"]["config_sheet_id"])
    sources = []
    for r in sh.worksheet("Sources").get_all_records():
        url = str(r.get("Sheet URL", "")).strip()
        active = str(r.get("Active", "yes")).strip().lower() not in ("no", "n", "false", "0")
        if url and active:
            sources.append({"name": str(r.get("Name", "")).strip(), "url": url})
    rows = sh.worksheet("Access").get_all_values()
    emails = {r[0].strip().lower() for r in rows[1:] if r and "@" in r[0]}
    emails |= {e.lower() for e in st.secrets["app"].get("admin_emails", [])}
    return sources, emails


def read_source(gc, url, override=""):
    sh = gc.open_by_url(url)
    name = override or clean_title(sh.title)
    m = re.search(r"gid=(\d+)", url)
    ws = next((w for w in sh.worksheets() if m and str(w.id) == m.group(1)), sh.sheet1)
    rows = ws.get_all_values()
    h = next(i for i, r in enumerate(rows) if any(c.strip().lower() == "timestamp" for c in r))
    df = pd.DataFrame(rows[h + 1:], columns=[c.strip() for c in rows[h]])

    ts_raw = df[find_col(df, "timestamp")]
    ts = pd.to_datetime(ts_raw, dayfirst=True, format="mixed", errors="coerce")
    fee = money(df[find_col(df, "course fee")])
    total = money(df[find_col(df, "total sale")])
    dp = money(df[find_col(df, "payment/dp")]).fillna(0)
    fin = money(df[find_col(df, "disbursement")]).fillna(0)
    fallback = (dp + fin).where(lambda s: s > 0, fee)
    value = total.fillna(fallback)
    opted = df[find_col(df, "opted")].astype(str).str.strip().str.lower().eq("yes")

    out = pd.DataFrame({
        "Salesperson": name, "date": ts.dt.normalize(),
        "Course": df[find_col(df, "course name")].map(clean_course),
        "value": value, "estimated": total.isna(), "opted_out": opted,
    })
    skipped = int((ts.isna() & ts_raw.str.strip().ne("")).sum())
    out = out.dropna(subset=["date", "value"])
    notes = []
    if skipped:
        notes.append(f"{name}: {skipped} row(s) skipped (unreadable date)")
    if out.estimated.sum():
        notes.append(f"{name}: {int(out.estimated.sum())} row(s) had a broken Total Sale Amount (#REF!), value estimated from payments")
    return name, out, notes


@st.cache_data(ttl=300, show_spinner="Loading data from Google Sheets...")
def load_all(sources):
    gc, frames, issues, names = get_client(), [], [], []
    sa = sa_info()["client_email"]
    for s in sources:
        try:
            name, df, notes = read_source(gc, s["url"], s["name"])
            frames.append(df)
            names.append(name)
            issues += notes
        except (gspread.exceptions.SpreadsheetNotFound, gspread.exceptions.APIError, PermissionError) as e:
            issues.append(f"Could not open {s['name'] or s['url']}: share the sheet with {sa} (Viewer). Details: {e}")
        except Exception as e:  # bad layout, etc.
            issues.append(f"Could not read {s['name'] or s['url']}: {e}")
    data = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if not data.empty:
        data = add_periods(data)
    return data, names, issues, datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d %b %Y, %I:%M %p")


# ---------- tables & charts ----------
def breakdown(df, label, sort, asc):
    keys = list(dict.fromkeys([sort, label]))
    g = df.groupby(keys, as_index=False).agg(Sales=("value", "size"), Value=("value", "sum"))
    g = g.sort_values("Value", ascending=False) if sort == label else g.sort_values(sort, ascending=asc)
    g["Avg"] = g.Value / g.Sales
    g["Share"] = g.Value / g.Value.sum() * 100
    return g


def render_table(g, label, key, total_label="Subtotal"):
    t = pd.DataFrame({label: g[label], "No. of Sales": g.Sales, "Sale Value (Rs)": g.Value.map(inr),
                      "Avg. Sale (Rs)": g.Avg.map(inr), "Share %": g.Share.map("{:.1f}%".format)})
    n, v = int(g.Sales.sum()), g.Value.sum()
    t.loc[len(t)] = [total_label, n, inr(v), inr(v / n), "-"]
    st.dataframe(t, hide_index=True, use_container_width=True, key=key)


# ---------- look & feel ----------
CSS = """
<style>
section[data-testid="stSidebar"] {background-color:#0B1220;}
section[data-testid="stSidebar"] * {color:#E5E7EB !important;}
section[data-testid="stSidebar"] .stButton button {background:#1D9E75;border:none;color:#fff !important;}
section[data-testid="stSidebar"] hr {border-color:#1F2937;}
div[data-testid="stMetric"] {background:var(--surface-1, #F1EFE8);border-radius:12px;padding:14px 16px;}
.kpi-row {display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin-bottom:14px;}
.kpi-card {background:var(--surface-1, #F1EFE8);border-radius:12px;padding:14px;display:flex;gap:10px;align-items:center;}
.kpi-icon {width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;flex-shrink:0;font-size:18px;font-weight:600;color:#fff;}
.kpi-text .kpi-label {font-size:12px;color:#6B7280;margin:0;}
.kpi-text .kpi-value {font-size:18px;font-weight:600;margin:2px 0 0;}
.stTabs [data-baseweb="tab-list"] {gap:4px;}
.stTabs [data-baseweb="tab"] {background:var(--surface-1, #F1EFE8);border-radius:8px;padding:6px 14px;}
</style>
"""


def kpi_card(icon, bg, label, value):
    return (f'<div class="kpi-card"><div class="kpi-icon" style="background:{bg};">{icon}</div>'
            f'<div class="kpi-text"><p class="kpi-label">{label}</p><p class="kpi-value">{value}</p></div></div>')


def kpi_row(cards):
    st.markdown('<div class="kpi-row">' + "".join(cards) + "</div>", unsafe_allow_html=True)


def donut_chart(g, label, key):
    colors = [color_for(list(g[label]), n) for n in g[label]]
    fig = go.Figure(go.Pie(labels=g[label], values=g.Value, hole=0.6, marker=dict(colors=colors),
                            textinfo="none", hovertemplate="%{label}: Rs %{value:,.0f}<extra></extra>"))
    fig.update_layout(showlegend=True, margin=dict(l=0, r=0, t=0, b=0), height=220,
                       legend=dict(orientation="h", yanchor="bottom", y=-0.25))
    st.plotly_chart(fig, use_container_width=True, key=key)


def glance_table(df, label, key):
    """Small, always-sorted-by-value table for the Today/This Month snapshot (no chart, no filters)."""
    if df.empty:
        st.caption("No sales yet.")
        return
    g = breakdown(df, label, label, False)
    render_table(g, label, key, "Total")


def glance_section(base, names):
    """Fixed overview: today's and this month's sales, by salesperson and by course. Ignores all sidebar filters."""
    today = pd.Timestamp(datetime.now(ZoneInfo("Asia/Kolkata")).date())
    today_df = base[base.date == today]
    month_df = base[(base.date.dt.year == today.year) & (base.date.dt.month == today.month)]

    st.subheader("Today & This Month at a Glance")
    st.caption("Always shows today and this month, not affected by the filters below.")
    kpi_row([
        kpi_card("&#8377;", "#185FA5", "Today's value", inr(today_df.value.sum()) if len(today_df) else "Rs 0"),
        kpi_card("#", "#534AB7", "Today's sales", len(today_df)),
        kpi_card("&#8599;", "#0F6E56", "This month's value", inr(month_df.value.sum()) if len(month_df) else "Rs 0"),
        kpi_card("&#128100;", "#854F0B", "This month's sales", len(month_df)),
    ])

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Today - by Salesperson**")
        glance_table(today_df, "Salesperson", "glance-today-person")
    with c2:
        st.markdown("**Today - by Course**")
        glance_table(today_df, "Course", "glance-today-course")

    c3, c4 = st.columns(2)
    with c3:
        st.markdown("**This Month - by Salesperson**")
        glance_table(month_df, "Salesperson", "glance-month-person")
    with c4:
        st.markdown("**This Month - by Course**")
        glance_table(month_df, "Course", "glance-month-course")
    st.divider()


def person_view(df, key, asc):
    n, v = len(df), df.value.sum()
    a, b, c = st.columns(3)
    a.metric("Total Sales", n)
    b.metric("Total Value", inr(v))
    c.metric("Avg. Sale Value", inr(v / n) if n else "-")
    if not n:
        st.info("No sales in the selected filters.")
        return
    tabs = st.tabs(["Monthly", "Weekly", "Daily", "Courses"])
    for tab, (label, sort) in zip(tabs, [("Month", "month_start"), ("Week", "week_start"), ("Day", "date"), ("Course", "Course")]):
        with tab:
            g = breakdown(df, label, sort, asc)
            chart = alt.Chart(g).mark_bar().encode(
                x=alt.X(f"{label}:N", sort=None, title=None, axis=alt.Axis(labelAngle=-90, labelOverlap=False, labelFontSize=9)),
                y=alt.Y("Value:Q", title="Sale Value (Rs)"), tooltip=[label, "Sales", "Value"])
            st.altair_chart(chart, use_container_width=True, key=f"{key}-{label}-chart")
            render_table(g, label, f"{key}-{label}-table", "Total" if label == "Course" else "Subtotal")


# ---------- app ----------
def gate():
    if not st.user.is_logged_in:
        st.title("Sales Dashboard")
        st.write("Please sign in with your Google account.")
        st.button("Sign in with Google", on_click=st.login)
        st.stop()
    try:
        _, allowed = load_config()
    except Exception as e:
        st.error(f"Could not read the config sheet: {e}")
        st.stop()
    if st.user.email.lower() not in allowed:
        st.error(f"{st.user.email} does not have access. Ask the account owner to add your email.")
        st.button("Log out", on_click=st.logout)
        st.stop()


def main():
    st.set_page_config(page_title="Combined CRM Sales Report", page_icon="📊", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    gate()
    sources, _ = load_config()
    data, names, issues, loaded_at = load_all(sources)

    with st.sidebar:
        st.markdown("### &#128202; CRM Analytics")
        st.caption("Combined sales dashboard")
        st.divider()
        st.write(f"Signed in as **{st.user.email}**")
        st.button("Log out", on_click=st.logout)
        if st.button("Refresh data now"):
            st.cache_data.clear()
            st.rerun()

    st.title("Combined CRM Sales Report")
    st.caption(f"Data as of {loaded_at}  |  Sources: {', '.join(names) or 'none'}")

    if issues:
        with st.expander(f"Data checks ({len(issues)})", expanded=any("Could not" in i for i in issues)):
            for i in issues:
                st.write("- " + i)
    if data.empty:
        st.warning("No data loaded. Check the Sources tab in the config sheet.")
        st.stop()

    with st.sidebar:
        st.divider()
        st.header("Filters")
        incl_opt = st.checkbox("Include opted-out sales", value=False)

    base = data if incl_opt else data[~data.opted_out]
    glance_section(base, names)

    with st.sidebar:
        dmin, dmax = data.date.min().date(), data.date.max().date()
        today_date = datetime.now(ZoneInfo("Asia/Kolkata")).date()
        preset = st.radio("Date range", ["Today", "Last 2 days", "Last 3 days", "Custom"], index=3)
        if preset == "Today":
            rng = (today_date, today_date)
        elif preset == "Last 2 days":
            rng = (today_date - timedelta(days=1), today_date)
        elif preset == "Last 3 days":
            rng = (today_date - timedelta(days=2), today_date)
        else:
            rng = st.date_input("Custom range", (dmin, dmax), min_value=dmin, max_value=dmax)
        picked = st.multiselect("Salespersons", names, default=names)
        newest = st.radio("Sort by date", ["Oldest first", "Newest first"]) == "Newest first"
    if len(rng) != 2:
        st.info("Pick both a start and an end date.")
        st.stop()

    f = base[base.date.between(pd.Timestamp(rng[0]), pd.Timestamp(rng[1])) & base.Salesperson.isin(picked)]
    asc = not newest

    st.subheader("Filtered Report")
    st.caption(f"{preset} - {rng[0]:%d %b %Y} to {rng[1]:%d %b %Y}")
    s = f.groupby("Salesperson").agg(Sales=("value", "size"), Value=("value", "sum")).reindex(picked).fillna(0)
    tot = s.Value.sum()
    summary = pd.DataFrame({
        "Salesperson": s.index, "Total Sales": s.Sales.astype(int).values,
        "Total Value (Rs)": s.Value.map(inr).values,
        "Avg. Sale Value (Rs)": [inr(v / n) if n else "-" for v, n in zip(s.Value, s.Sales)],
        "Share of Total": [f"{v / tot * 100:.1f}%" if tot else "-" for v in s.Value]})
    n_all = int(s.Sales.sum())
    summary.loc[len(summary)] = ["TOTAL (All)", n_all, inr(tot), inr(tot / n_all) if n_all else "-", "100%"]

    t1, t2 = st.columns([2, 1])
    with t1:
        st.dataframe(summary, hide_index=True, use_container_width=True)
    with t2:
        if tot:
            donut_chart(s.reset_index().rename(columns={"index": "Salesperson"}), "Salesperson", "summary-donut")
        else:
            st.caption("No sales in this range yet.")

    tabs = st.tabs(["Overall (All)"] + picked)
    with tabs[0]:
        person_view(f, "overall", asc)
    for tab, name in zip(tabs[1:], picked):
        with tab:
            person_view(f[f.Salesperson == name], name, asc)


if __name__ == "__main__":
    main()
