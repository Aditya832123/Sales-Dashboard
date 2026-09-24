"""Combined CRM Sales Dashboard (Streamlit).

Client controls everything from ONE Google Sheet ("Dashboard Config"):
  - tab "Sources": Name (optional) | Sheet URL | Active (optional yes/no)
  - tab "Access" : one allowed Google email per row (header in row 1)
Add a row to Sources -> the new salesperson appears on the dashboard.
"""
import json
import re
from datetime import datetime
from zoneinfo import ZoneInfo

import altair as alt
import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

COURSE_ORDER = ["MW", "DSP", "PV", "CR", "CDM", "RA", "MC", "SAS"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


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
                x=alt.X(f"{label}:N", sort=None, title=None, axis=alt.Axis(labelAngle=-45)),
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
    gate()
    sources, _ = load_config()
    data, names, issues, loaded_at = load_all(sources)

    st.title("Combined CRM Sales Report")
    st.caption(f"Data as of {loaded_at}  |  Sources: {', '.join(names) or 'none'}")

    with st.sidebar:
        st.write(f"Signed in as **{st.user.email}**")
        st.button("Log out", on_click=st.logout)
        if st.button("Refresh data now"):
            st.cache_data.clear()
            st.rerun()

    if issues:
        with st.expander(f"Data checks ({len(issues)})", expanded=any("Could not" in i for i in issues)):
            for i in issues:
                st.write("- " + i)
    if data.empty:
        st.warning("No data loaded. Check the Sources tab in the config sheet.")
        st.stop()

    with st.sidebar:
        st.header("Filters")
        dmin, dmax = data.date.min().date(), data.date.max().date()
        rng = st.date_input("Date range", (dmin, dmax), min_value=dmin, max_value=dmax)
        picked = st.multiselect("Salespersons", names, default=names)
        newest = st.radio("Sort by date", ["Oldest first", "Newest first"]) == "Newest first"
        incl_opt = st.checkbox("Include opted-out sales", value=False)
    if len(rng) != 2:
        st.info("Pick both a start and an end date.")
        st.stop()

    f = data[data.date.between(pd.Timestamp(rng[0]), pd.Timestamp(rng[1])) & data.Salesperson.isin(picked)]
    if not incl_opt:
        f = f[~f.opted_out]
    asc = not newest

    st.subheader("Executive Summary")
    s = f.groupby("Salesperson").agg(Sales=("value", "size"), Value=("value", "sum")).reindex(picked).fillna(0)
    tot = s.Value.sum()
    summary = pd.DataFrame({
        "Salesperson": s.index, "Total Sales": s.Sales.astype(int).values,
        "Total Value (Rs)": s.Value.map(inr).values,
        "Avg. Sale Value (Rs)": [inr(v / n) if n else "-" for v, n in zip(s.Value, s.Sales)],
        "Share of Total": [f"{v / tot * 100:.1f}%" if tot else "-" for v in s.Value]})
    n_all = int(s.Sales.sum())
    summary.loc[len(summary)] = ["TOTAL (All)", n_all, inr(tot), inr(tot / n_all) if n_all else "-", "100%"]
    st.dataframe(summary, hide_index=True, use_container_width=True)

    tabs = st.tabs(["Overall (All)"] + picked)
    with tabs[0]:
        person_view(f, "overall", asc)
    for tab, name in zip(tabs[1:], picked):
        with tab:
            person_view(f[f.Salesperson == name], name, asc)


if __name__ == "__main__":
    main()
