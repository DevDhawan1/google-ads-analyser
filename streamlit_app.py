#!/usr/bin/env python3
"""
myHQ Google Ads Agent — Streamlit Web App
Deploy to Streamlit Cloud for org-wide access.

Setup:
  pip install streamlit google-ads openpyxl pandas
  streamlit run streamlit_app.py

Deploy:
  Push this file + myhq_google_ads_agent.py + requirements.txt to GitHub,
  then connect at share.streamlit.io.
"""

import io
import os
import re
from datetime import datetime
from collections import defaultdict

import streamlit as st

# Load .env file if present (requires python-dotenv, silently ignored if not installed)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

st.set_page_config(
    page_title="myHQ · Google Ads Agent",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

try:
    from myhq_google_ads_agent import (
        GoogleAdsClient, CONFIG, T,
        spend, impressions, clicks,
        ctr as _ctr, cpm as _cpm, avg_cpc as _avg_cpc,
        conversions, conv_value, imp_share as _imp_share,
        cpl, roas, classify,
        inr, auto_quick_wins, generate_html, generate_excel, OPENPYXL_OK,
    )
    import pandas as pd
except ImportError as exc:
    st.error(
        f"**Import error:** `{exc}`\n\n"
        "Make sure `myhq_google_ads_agent.py` is in the same folder as `streamlit_app.py`, "
        "and that all requirements are installed:\n\n"
        "```\npip install streamlit google-ads openpyxl pandas\n```"
    )
    st.stop()


# ── Custom CSS ─────────────────────────────────────────────────
st.markdown("""
<style>
  [data-testid="metric-container"] {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: 16px 20px 14px;
    box-shadow: 0 1px 3px rgba(0,0,0,.06);
  }
  [data-testid="metric-container"] label {
    font-size: 11px !important;
    text-transform: uppercase;
    letter-spacing: .5px;
    color: #6b7280 !important;
  }
  [data-testid="stMetricValue"] { font-size: 22px !important; font-weight: 800 !important; }
  .stDownloadButton > button { font-weight: 700; border-radius: 8px; }
  .alert-card {
    background: #fff8f8;
    border: 1px solid #fca5a5;
    border-left: 4px solid #dc2626;
    border-radius: 0 8px 8px 0;
    padding: 16px 20px;
    margin: 12px 0;
  }
  .alert-card h4 { color: #dc2626; margin: 0 0 8px; }
  .win-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 8px;
    padding: 14px 18px;
    margin: 10px 0;
  }
  .win-card .rank {
    display: inline-block;
    background: #1e2d3d;
    color: #fff;
    font-weight: 700;
    font-size: 13px;
    border-radius: 50%;
    width: 26px; height: 26px;
    text-align: center;
    line-height: 26px;
    margin-right: 10px;
    vertical-align: middle;
  }
  .stTabs [data-baseweb="tab"] { font-size: 14px; font-weight: 600; }
  section[data-testid="stSidebar"] { min-width: 320px; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════

def _excel_bytes(account, campaign_rows, adgroup_rows, ad_rows_map, days, ts):
    buf = io.BytesIO()
    generate_excel(account, campaign_rows, adgroup_rows, ad_rows_map, days, ts, buf)
    buf.seek(0)
    return buf.getvalue()


def _strip_html(text):
    return re.sub(r"<[^>]+>", "", text or "")


def _tier_icon(tier):
    return {"green": "🟢", "yellow": "🟡", "red": "🔴", "grey": "⚪"}.get(tier, "")


def _build_campaign_df(rows):
    data = []
    for r in rows:
        tier, label = classify(r)
        cp = cpl(r); ro = roas(r); ct = _ctr(r); ish = _imp_share(r)
        data.append({
            "Status":       f"{_tier_icon(tier)} {label}",
            "Campaign":     r.get("campaign_name", "—"),
            "Spend (₹)":   int(spend(r)),
            "Conversions":  int(conversions(r)),
            "CPL (₹)":     round(cp) if cp is not None else None,
            "ROAS":         round(ro, 2) if ro is not None else None,
            "CTR %":        round(ct, 2),
            "Imp. Share %": round(ish, 1) if ish > 0 else None,
            "_tier":        tier,
        })
    return pd.DataFrame(data)


def _style_campaign_df(df):
    TIER_BG = {
        "green":  "background-color:#f0fdf4; color:#111827",
        "yellow": "background-color:#fefce8; color:#111827",
        "red":    "background-color:#fff0f0; color:#111827",
        "grey":   "background-color:#f9fafb; color:#111827",
    }
    def row_style(row):
        bg = TIER_BG.get(df.at[row.name, "_tier"], "")
        return [bg] * len(row)

    display = df.drop(columns=["_tier"])
    return (
        display.style
        .apply(row_style, axis=1)
        .format({
            "Spend (₹)":   "₹{:,.0f}",
            "CPL (₹)":     lambda x: f"₹{int(x):,}" if pd.notna(x) and x else "—",
            "ROAS":        lambda x: f"{x:.2f}×" if pd.notna(x) and x else "—",
            "CTR %":       "{:.2f}%",
            "Imp. Share %": lambda x: f"{x:.1f}%" if pd.notna(x) and x else "—",
            "Conversions": "{:,}",
        })
    )


def _build_adgroup_df(rows):
    data = []
    for a in rows:
        cp = cpl(a); ro = roas(a); ct = _ctr(a); ish = _imp_share(a)
        if cp and cp > T["cpl_red"]:                               status = "🔴 Cap Spend"
        elif cp and cp <= T["cpl_green"] and ct >= T["ctr_green"]: status = "🟢 Scale"
        elif ish > 0 and ish < T["imp_share_red"]:                 status = "🟡 Watch"
        else:                                                       status = "🟡 Hold"
        data.append({
            "Campaign":     a.get("campaign_name", "—"),
            "Ad Group":     a.get("ad_group_name", "—"),
            "Spend (₹)":   int(spend(a)),
            "Conversions":  int(conversions(a)),
            "CPL (₹)":     round(cp) if cp is not None else None,
            "ROAS":         round(ro, 2) if ro is not None else None,
            "CTR %":        round(ct, 2),
            "Imp. Share %": round(ish, 1) if ish > 0 else None,
            "Status":       status,
        })
    return pd.DataFrame(data)


def _build_ad_df(ad_rows_map):
    data = []
    for cid, ad_rows in ad_rows_map.items():
        cpls = {a.get("ad_name"): cpl(a) for a in ad_rows if cpl(a)}
        best = min(cpls, key=cpls.get) if cpls else None
        for a in sorted(ad_rows, key=spend, reverse=True):
            ad_name = a.get("ad_name", "Unknown")
            cp = cpl(a)
            if ad_name == best:                      verdict = "⭐ Keep + Scale"
            elif cp and cp > T["cpl_amber"]:         verdict = "❌ Pause"
            else:                                    verdict = "🔍 Test further"
            data.append({
                "Campaign":     a.get("campaign_name", "—"),
                "Ad Group":     a.get("ad_group_name", "—"),
                "Ad Name":      ad_name,
                "Ad Type":      a.get("ad_type", "—"),
                "Spend (₹)":   int(spend(a)),
                "Impressions":  int(impressions(a)),
                "CTR %":        round(_ctr(a), 2),
                "Conversions":  int(conversions(a)),
                "CPL (₹)":     round(cp) if cp is not None else None,
                "Verdict":      verdict,
            })
    return pd.DataFrame(data)


def _run_analysis(developer_token, client_id, client_secret, refresh_token, customer_id, days):
    client = GoogleAdsClient(developer_token, client_id, client_secret, refresh_token, customer_id)

    with st.status("Connecting to Google Ads API…", expanded=True) as status:

        st.write("🔑 Fetching account info…")
        try:
            account = client.account_info()
        except Exception as e:
            raise RuntimeError(f"Failed to connect: {e}")
        st.write(f"✔ Connected to **{account.get('name')}** (`{account.get('id')}`)")

        st.write(f"📋 Fetching {days}-day campaign insights…")
        try:
            campaign_rows = client.campaign_insights(days)
        except Exception as e:
            raise RuntimeError(f"Campaign fetch failed: {e}")
        if not campaign_rows:
            raise RuntimeError(
                "No active campaigns found for this account in the selected time window."
            )
        st.write(f"✔ **{len(campaign_rows)}** campaign rows")

        st.write("🗺 Fetching ad group breakdown…")
        try:
            adgroup_rows = client.adgroup_insights(days)
        except Exception as e:
            raise RuntimeError(f"Ad group fetch failed: {e}")
        st.write(f"✔ **{len(adgroup_rows)}** ad group rows")

        classified = [(r, *classify(r)) for r in campaign_rows]
        red_cids   = [r.get("campaign_id") for r, t, _ in classified if t == "red"]
        ad_rows_map = {}

        if red_cids:
            st.write(f"🔴 Pulling ad-level data for **{len(red_cids)}** underperforming campaign(s)…")
            try:
                all_ads = client.ad_insights(red_cids, days)
                for a in all_ads:
                    cid = a.get("campaign_id")
                    if cid not in ad_rows_map:
                        ad_rows_map[cid] = []
                    ad_rows_map[cid].append(a)
            except Exception as e:
                st.warning(f"Ad-level fetch partially failed: {e}")
            total_ads = sum(len(v) for v in ad_rows_map.values())
            st.write(f"✔ **{total_ads}** ad rows")
        else:
            st.write("✔ No underperforming campaigns — ad-level pull skipped")

        st.write("⚙️ Building HTML report and Excel workbook…")
        now  = datetime.now()
        html = generate_html(account, campaign_rows, adgroup_rows, ad_rows_map, days, now)
        xlsx = _excel_bytes(account, campaign_rows, adgroup_rows, ad_rows_map, days, now)
        st.write("✔ Reports ready")

        status.update(label="✅ Analysis complete!", state="complete", expanded=False)

    return {
        "account":       account,
        "campaign_rows": campaign_rows,
        "adgroup_rows":  adgroup_rows,
        "ad_rows_map":   ad_rows_map,
        "days":          days,
        "generated_at":  now,
        "html":          html,
        "excel_bytes":   xlsx,
    }


# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown(
        "<div style='padding:4px 0 12px'>"
        "<span style='font-size:28px;font-weight:900;color:#1e2d3d;letter-spacing:-1px'>myHQ</span>"
        "<span style='font-size:13px;color:#6b7280;margin-left:8px;font-weight:500'>"
        "Google Ads Agent</span></div>",
        unsafe_allow_html=True,
    )

    st.markdown("#### ⚙️ Configuration")

    developer_token_input = st.text_input(
        "Developer Token",
        value=os.environ.get("GOOGLE_ADS_DEVELOPER_TOKEN", ""),
        type="password",
        placeholder="xxxxxxxxxxxxxxxxxxxx",
        help="From Google Ads UI → Tools → API Center. Tied to your MCC account.",
    )

    client_id_input = st.text_input(
        "OAuth2 Client ID",
        value=os.environ.get("GOOGLE_ADS_CLIENT_ID", ""),
        type="password",
        placeholder="1234567890-abc.apps.googleusercontent.com",
        help="Google Cloud Console → APIs & Services → Credentials → OAuth 2.0 Client ID",
    )

    client_secret_input = st.text_input(
        "OAuth2 Client Secret",
        value=os.environ.get("GOOGLE_ADS_CLIENT_SECRET", ""),
        type="password",
        placeholder="GOCSPX-xxxxxxxxxxxxxxxx",
        help="Same location as Client ID — the secret value for your Desktop app credential.",
    )

    refresh_token_input = st.text_input(
        "OAuth2 Refresh Token",
        value=os.environ.get("GOOGLE_ADS_REFRESH_TOKEN", ""),
        type="password",
        placeholder="1//0gxxxxxxxxxxxxxxxx",
        help="Generate once using generate_refresh_token.py from google-ads-python examples.",
    )

    customer_id_input = st.text_input(
        "Customer ID",
        value=os.environ.get("GOOGLE_ADS_CUSTOMER_ID", ""),
        placeholder="123-456-7890",
        help="The ad account to analyse. Find it in the top-right of Google Ads UI (format XXX-XXX-XXXX).",
    )

    days_input = st.select_slider(
        "Lookback Window",
        options=[7, 14, 30, 60, 90],
        value=30,
        format_func=lambda x: f"Last {x} days",
    )

    all_filled = all([
        developer_token_input.strip(),
        client_id_input.strip(),
        client_secret_input.strip(),
        refresh_token_input.strip(),
        customer_id_input.strip(),
    ])

    run_btn = st.button(
        "🚀  Run Analysis",
        type="primary",
        use_container_width=True,
        disabled=not all_filled,
    )

    if not all_filled:
        st.caption("Enter all five credentials above to enable.")

    st.divider()

    st.markdown("#### 📋 How to get credentials")

    with st.expander("Step-by-step instructions", expanded=False):
        st.markdown("""
**Step 1 — Enable Google Ads API**

Go to [Google Cloud Console](https://console.cloud.google.com/) → APIs & Services →
Library → search **Google Ads API** → Enable.

---

**Step 2 — Create OAuth2 credentials**

APIs & Services → Credentials → **Create Credentials** → OAuth client ID →
Application type: **Desktop app** → give it a name → Create.

Copy the **Client ID** and **Client Secret**.

---

**Step 3 — Generate a Refresh Token**

Download and run `generate_refresh_token.py` from the
[google-ads-python examples](https://github.com/googleads/google-ads-python/blob/main/examples/authentication/generate_refresh_token.py):

```bash
python generate_refresh_token.py \\
  --client_id YOUR_CLIENT_ID \\
  --client_secret YOUR_CLIENT_SECRET \\
  --scopes https://www.googleapis.com/auth/adwords
```

Copy the refresh token printed at the end.

---

**Step 4 — Find your Developer Token**

In Google Ads UI → Tools & Settings → **API Center**.
Your developer token is listed there (request test access first if needed).

---

**Step 5 — Find your Customer ID**

In Google Ads UI, look at the top-right corner — it shows your account ID in
`XXX-XXX-XXXX` format. Paste it with or without dashes.
""")

    st.divider()
    st.caption(
        f"Thresholds: CPL green ≤ ₹{T['cpl_green']:,} · "
        f"red > ₹{T['cpl_red']:,} · "
        f"ROAS green ≥ {T['roas_green']}×\n\n"
        "Edit `CONFIG` in `myhq_google_ads_agent.py` to customise."
    )


# ══════════════════════════════════════════════════════════════
# MAIN AREA
# ══════════════════════════════════════════════════════════════

if run_btn and all_filled:
    try:
        results = _run_analysis(
            developer_token_input.strip(),
            client_id_input.strip(),
            client_secret_input.strip(),
            refresh_token_input.strip(),
            customer_id_input.strip(),
            days_input,
        )
        st.session_state["results"] = results
    except RuntimeError as e:
        st.error(f"**Analysis failed:** {_strip_html(str(e))}")
        st.session_state.pop("results", None)
    except Exception as e:
        st.error(f"**Unexpected error:** {e}")
        st.session_state.pop("results", None)


# ══════════════════════════════════════════════════════════════
# RESULTS
# ══════════════════════════════════════════════════════════════

if "results" in st.session_state:
    res          = st.session_state["results"]
    campaign_rows = res["campaign_rows"]
    adgroup_rows  = res["adgroup_rows"]
    ad_rows_map   = res["ad_rows_map"]
    account       = res["account"]
    days          = res["days"]
    ts            = res["generated_at"]
    date_str      = ts.strftime("%Y-%m-%d")

    classified = [(r, *classify(r)) for r in campaign_rows]
    red_rows   = [r for r, t, _ in classified if t == "red"]

    # Header strip
    st.markdown(
        f"<div style='background:#1e2d3d;color:#fff;border-radius:10px;"
        f"padding:20px 28px 18px;margin-bottom:24px'>"
        f"<div style='font-size:13px;opacity:.6;letter-spacing:1px;text-transform:uppercase;"
        f"margin-bottom:4px'>myHQ · Performance Marketing</div>"
        f"<div style='font-size:22px;font-weight:800;letter-spacing:-.3px'>"
        f"Google Ads — Performance Review</div>"
        f"<div style='font-size:13px;opacity:.75;margin-top:6px'>"
        f"Last {days} days &nbsp;·&nbsp; {len(campaign_rows)} active campaigns "
        f"&nbsp;·&nbsp; {account.get('name')} ({account.get('id')}) "
        f"&nbsp;·&nbsp; Generated {ts.strftime('%d %b %Y %H:%M')}"
        f"</div></div>",
        unsafe_allow_html=True,
    )

    # Download buttons
    st.markdown("#### ⬇️ Download Reports")
    dl1, dl2, _ = st.columns([2, 2, 3])

    with dl1:
        st.download_button(
            label="📄  HTML Report",
            data=res["html"].encode("utf-8"),
            file_name=f"Google_Ads_Report_{date_str}.html",
            mime="text/html",
            use_container_width=True,
            type="primary",
            help="Full VP-ready tabbed report — open in any browser or print to PDF",
        )
    with dl2:
        st.download_button(
            label="📊  Excel Workbook",
            data=res["excel_bytes"],
            file_name=f"Google_Ads_Report_{date_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="5-sheet workbook: Summary, Campaigns, Ad Groups, Ads, Quick Wins",
        )

    st.markdown("---")

    # KPI strip
    total_spend_v = sum(spend(r) for r in campaign_rows)
    total_conv_v  = sum(conversions(r) for r in campaign_rows)
    best_cpl_v    = min((cpl(r) for r in campaign_rows if cpl(r)),  default=None)
    best_roas_v   = max((roas(r) for r in campaign_rows if roas(r)), default=None)
    avg_ish_v     = (sum(_imp_share(r) for r in campaign_rows if _imp_share(r) > 0)
                     / max(sum(1 for r in campaign_rows if _imp_share(r) > 0), 1))

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Total Spend",       f"₹{int(total_spend_v):,}")
    k2.metric("Total Conversions", f"{int(total_conv_v):,}")
    k3.metric("Best CPL",          f"₹{int(best_cpl_v):,}" if best_cpl_v else "N/A")
    k4.metric("Best ROAS",         f"{best_roas_v:.2f}×"   if best_roas_v else "N/A")
    k5.metric("Avg Imp. Share",    f"{avg_ish_v:.1f}%",
              delta="Low — check" if avg_ish_v < T["imp_share_red"] else None,
              delta_color="inverse")

    st.markdown("---")

    # Tabs
    tab_overview, tab_alerts, tab_adgroups, tab_ads = st.tabs([
        "📊 Overview",
        f"⚠️ Alerts & Actions  ({len(red_rows)})",
        "🗺 Ad Group Breakdown",
        "🎯 Ad Analysis",
    ])

    # ── Tab 1: Overview ──────────────────────────────────────
    with tab_overview:
        n_green = sum(1 for _, t, _ in classified if t == "green")
        n_amber = sum(1 for _, t, _ in classified if t == "yellow")
        n_red   = len(red_rows)
        n_grey  = sum(1 for _, t, _ in classified if t == "grey")

        h1, h2, h3, h4 = st.columns(4)
        h1.metric("🟢 Performing",      n_green)
        h2.metric("🟡 Watch",           n_amber)
        h3.metric("🔴 Underperforming", n_red)
        h4.metric("⚪ Too Early",        n_grey)

        st.markdown("##### Campaign Triage Table")
        st.caption(
            "🟢 Performing — CPL on target, CTR >1%. &nbsp;"
            "🟡 Watch — borderline metrics. &nbsp;"
            "🔴 Underperforming — CPL >₹500 or CTR <0.5%. &nbsp;"
            "⚪ Too Early — insufficient spend."
        )

        camp_df = _build_campaign_df(campaign_rows)
        styled  = _style_campaign_df(camp_df)
        st.dataframe(
            styled,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Spend (₹)":    st.column_config.NumberColumn("Spend (₹)"),
                "Conversions":  st.column_config.NumberColumn("Conversions"),
                "Imp. Share %": st.column_config.NumberColumn("Imp. Share %"),
            },
        )

    # ── Tab 2: Alerts & Actions ──────────────────────────────
    with tab_alerts:
        if not red_rows:
            st.success("No underperforming campaigns in this period.")
        else:
            st.markdown("##### Underperforming Campaigns")

            adgroup_by_cid = defaultdict(list)
            for a in adgroup_rows:
                adgroup_by_cid[a.get("campaign_id")].append(a)

            for r in red_rows:
                cid  = r.get("campaign_id")
                name = r.get("campaign_name", "Unknown")
                cp   = cpl(r); ro = roas(r); ct = _ctr(r); ish = _imp_share(r)

                with st.container():
                    st.markdown(
                        f"<div class='alert-card'><h4>🔴 {name}</h4></div>",
                        unsafe_allow_html=True,
                    )
                    issues = []
                    if cp and cp > T["cpl_red"]:
                        issues.append(f"CPL **₹{int(cp):,}** is {((cp/T['cpl_green'])-1)*100:.0f}% above the ₹{T['cpl_green']:,} target.")
                    if ro and ro < T["roas_red"]:
                        issues.append(f"ROAS **{ro:.2f}×** is below minimum viable {T['roas_red']}×.")
                    if ct < T["ctr_red"]:
                        issues.append(f"CTR **{ct:.2f}%** is below 0.5% — ad copy or targeting needs work.")
                    if ish > 0 and ish < T["imp_share_red"]:
                        issues.append(f"Search Impression Share **{ish:.1f}%** is below 30% — budget cap or quality issue.")
                    if not issues:
                        issues.append(f"Performance below account benchmarks. Spend: ₹{int(spend(r)):,}.")

                    for issue in issues:
                        st.warning(issue)

                    # Ad group gap signal
                    adgroups = adgroup_by_cid.get(cid, [])
                    cpls_ag  = [(a.get("ad_group_name"), cpl(a)) for a in adgroups if cpl(a)]
                    if len(cpls_ag) >= 2:
                        cpls_ag.sort(key=lambda x: x[1])
                        best_n, best_v   = cpls_ag[0]
                        worst_n, worst_v = cpls_ag[-1]
                        if worst_v > best_v * 1.8:
                            st.info(
                                f"**Ad Group Gap:** _{worst_n}_ costs ₹{int(worst_v):,} CPL vs "
                                f"₹{int(best_v):,} for _{best_n}_ — a {worst_v/best_v:.1f}× gap. "
                                "Reduce its budget and shift to the lower-CPL ad group."
                            )

        # Quick wins
        st.markdown("---")
        st.markdown("##### Quick Wins")
        wins = auto_quick_wins(campaign_rows, adgroup_rows, ad_rows_map, red_rows)
        if not wins:
            st.info("No quick wins generated — all campaigns are within targets.")
        for w in wins:
            with st.container():
                st.markdown(
                    f"<div class='win-card'>"
                    f"<span class='rank'>{w['rank']}</span>"
                    f"<strong>{w['title']}</strong>"
                    f"<br><small style='color:#6b7280'>{w.get('label','')}</small>"
                    f"<br>{_strip_html(w['body'])}"
                    f"</div>",
                    unsafe_allow_html=True,
                )

    # ── Tab 3: Ad Group Breakdown ────────────────────────────
    with tab_adgroups:
        adgroup_by_cid = defaultdict(list)
        for a in adgroup_rows:
            adgroup_by_cid[a.get("campaign_id")].append(a)

        for r, tier, label in classified:
            cid  = r.get("campaign_id")
            name = r.get("campaign_name", "Unknown")
            ags  = adgroup_by_cid.get(cid, [])

            with st.expander(
                f"{_tier_icon(tier)} {name}  ({len(ags)} ad groups)",
                expanded=(tier == "red"),
            ):
                if not ags:
                    st.caption("No ad group data for this campaign.")
                    continue
                ag_df = _build_adgroup_df(ags)
                st.dataframe(
                    ag_df.style.format({
                        "Spend (₹)":    "₹{:,.0f}",
                        "CPL (₹)":      lambda x: f"₹{int(x):,}" if pd.notna(x) and x else "—",
                        "ROAS":         lambda x: f"{x:.2f}×" if pd.notna(x) and x else "—",
                        "CTR %":        "{:.2f}%",
                        "Imp. Share %": lambda x: f"{x:.1f}%" if pd.notna(x) and x else "—",
                        "Conversions":  "{:,}",
                    }),
                    use_container_width=True,
                    hide_index=True,
                )

    # ── Tab 4: Ad Analysis ───────────────────────────────────
    with tab_ads:
        if not ad_rows_map:
            st.info("No underperforming campaigns — ad-level analysis skipped.")
        else:
            ad_df = _build_ad_df(ad_rows_map)
            if ad_df.empty:
                st.info("No ad data returned for underperforming campaigns.")
            else:
                st.markdown("##### Ad Performance — Underperforming Campaigns Only")
                st.caption(
                    "⭐ Keep + Scale — best CPL in its ad group. &nbsp;"
                    "❌ Pause — CPL >₹350. &nbsp;"
                    "🔍 Test further — insufficient data or mid-range CPL."
                )
                st.dataframe(
                    ad_df.style.format({
                        "Spend (₹)":   "₹{:,.0f}",
                        "CPL (₹)":     lambda x: f"₹{int(x):,}" if pd.notna(x) and x else "—",
                        "CTR %":       "{:.2f}%",
                        "Impressions": "{:,}",
                        "Conversions": "{:,}",
                    }),
                    use_container_width=True,
                    hide_index=True,
                )
