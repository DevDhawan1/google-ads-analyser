# Google Ads Analyser — Build Spec

Use this file to build a Google Ads performance analyser from scratch. Mirror the structure of the existing Meta Ads analyser (`streamlit_app.py` + `myhq_ads_agent.py`) but adapted for Google Ads.

---

## Purpose

A Streamlit web app that integrates with the Google Ads API to pull live campaign data and generate VP-ready performance reports. Same deployment target as the Meta analyser — Streamlit Cloud, org-wide, no data stored, credentials are session-only.

---

## Tech Stack

| Layer | Technology |
|---|---|
| UI | Streamlit (>=1.32.0) |
| API client | google-ads (>=23.0.0) |
| Report export | openpyxl (>=3.1.0), pandas (>=2.0.0) |
| API | Google Ads API v17 |
| Deployment | Streamlit Cloud (GitHub-based) |

---

## File Structure

Mirror the Meta analyser exactly:

| File | Purpose |
|---|---|
| `streamlit_app.py` | Streamlit UI — sidebar inputs, progress, 4-tab report, download buttons |
| `myhq_google_ads_agent.py` | Core logic — GoogleAdsClient, data helpers, HTML/Excel builders, CLI entry |
| `requirements.txt` | `streamlit`, `google-ads`, `openpyxl`, `pandas` |

---

## Authentication

Google Ads API requires four credentials. Collect all four in the Streamlit sidebar (use `type="password"` for secrets):

| Field | Sidebar label | Notes |
|---|---|---|
| Developer Token | Developer Token | From Google Ads API Center — tied to the MCC account |
| Client ID | OAuth2 Client ID | From Google Cloud Console → APIs & Services → Credentials |
| Client Secret | OAuth2 Client Secret | Same location as Client ID |
| Refresh Token | OAuth2 Refresh Token | Generate once via `google-auth-oauthlib` flow; user pastes it |
| Customer ID | Customer ID | The specific ad account to analyse, format `123-456-7890` or `1234567890` |

Build a `GoogleAdsClient` class (analogous to `MetaClient`) that constructs the client from these five values at runtime — do not use a `google-ads.yaml` file since credentials are session-only.

```python
from google.ads.googleads.client import GoogleAdsClient as _GAClient

class GoogleAdsClient:
    def __init__(self, developer_token, client_id, client_secret, refresh_token, customer_id):
        credentials = {
            "developer_token": developer_token,
            "client_id": client_id,
            "client_secret": client_secret,
            "refresh_token": refresh_token,
            "use_proto_plus": True,
        }
        self.client = _GAClient.load_from_dict(credentials)
        self.customer_id = customer_id.replace("-", "")  # strip dashes
        self.ga_service = self.client.get_service("GoogleAdsService")
```

---

## API Queries (GAQL)

All data is fetched via `GoogleAdsService.search_stream()` using GAQL. Use `date_preset` via `DURING LAST_N_DAYS` or a date range clause.

### Campaign-level insights
```sql
SELECT
  campaign.id, campaign.name, campaign.status,
  metrics.cost_micros, metrics.impressions, metrics.clicks,
  metrics.ctr, metrics.average_cpc, metrics.average_cpm,
  metrics.conversions, metrics.conversions_value,
  metrics.cost_per_conversion, metrics.search_impression_share,
  metrics.search_rank_lost_impression_share
FROM campaign
WHERE campaign.status = 'ENABLED'
  AND segments.date DURING LAST_30_DAYS
```

### Ad group-level insights
```sql
SELECT
  campaign.id, campaign.name,
  ad_group.id, ad_group.name, ad_group.status,
  metrics.cost_micros, metrics.impressions, metrics.clicks,
  metrics.ctr, metrics.average_cpm, metrics.average_cpc,
  metrics.conversions, metrics.conversions_value,
  metrics.cost_per_conversion, metrics.search_impression_share
FROM ad_group
WHERE campaign.status = 'ENABLED'
  AND ad_group.status = 'ENABLED'
  AND segments.date DURING LAST_30_DAYS
```

### Ad-level insights (red campaigns only)
```sql
SELECT
  campaign.id, campaign.name,
  ad_group.id, ad_group.name,
  ad_group_ad.ad.id, ad_group_ad.ad.name,
  ad_group_ad.ad.type,
  metrics.cost_micros, metrics.impressions, metrics.clicks,
  metrics.ctr, metrics.average_cpm,
  metrics.conversions, metrics.conversions_value,
  metrics.cost_per_conversion
FROM ad_group_ad
WHERE campaign.id IN ({campaign_ids})
  AND ad_group_ad.status = 'ENABLED'
  AND segments.date DURING LAST_30_DAYS
```

**Important:** `cost_micros` is cost in millionths of the account currency. Always divide by `1_000_000` to get the actual spend value.

---

## Data Helper Functions

Define module-level helpers that accept a row dict (normalised from the GAQL response). Mirror the Meta helpers exactly.

```python
def spend(r):       return r.get("cost_micros", 0) / 1_000_000
def impressions(r): return int(r.get("impressions", 0) or 0)
def clicks(r):      return int(r.get("clicks", 0) or 0)
def ctr(r):         return float(r.get("ctr", 0) or 0) * 100   # API returns 0–1, convert to %
def avg_cpc(r):     return r.get("average_cpc", 0) / 1_000_000
def cpm(r):         return r.get("average_cpm", 0) / 1_000_000
def conversions(r): return float(r.get("conversions", 0) or 0)
def conv_value(r):  return float(r.get("conversions_value", 0) or 0)
def imp_share(r):   return float(r.get("search_impression_share", 0) or 0) * 100  # to %

def cpl(r):
    c = conversions(r); s = spend(r)
    return s / c if c > 0 and s > 0 else None

def roas(r):
    v = conv_value(r); s = spend(r)
    return v / s if v > 0 and s > 0 else None
```

**Note on CTR:** The Google Ads API returns `ctr` as a decimal (e.g., `0.0312` = 3.12%). Multiply by 100 before storing so all downstream logic treats it as a percentage — same as the Meta analyser.

**Note on Frequency:** Google Ads Search campaigns do not have a frequency metric. For Display/YouTube campaigns, use `metrics.average_frequency`. Add it to the GAQL query only when the campaign type is Display or Video. For Search, leave frequency as `None` and omit the column from the report.

---

## Performance Thresholds (CONFIG)

Keep identical to the Meta analyser so the same mental model applies org-wide:

```python
CONFIG = {
    "api_version": "v17",
    "default_days": 30,
    "brand": "myHQ",
    "thresholds": {
        "cpl_green":  200,
        "cpl_amber":  350,
        "cpl_red":    500,
        "ctr_green":  1.0,    # % — Search CTR benchmark is higher than display
        "ctr_amber":  0.6,
        "ctr_red":    0.5,
        "roas_green": 3.0,
        "roas_red":   1.5,
        "imp_share_red": 30.0,   # Search Impression Share < 30% → worth flagging
        "min_spend":  3000,
    },
}
T = CONFIG["thresholds"]
```

---

## classify() Function

Same logic as Meta, with one addition — flag low impression share for Search campaigns:

```python
def classify(r):
    s = spend(r)
    if s < T["min_spend"]:
        return ("grey", "Too Early")

    cp = cpl(r);   ro = roas(r)
    ct = ctr(r);   ish = imp_share(r)

    # No conversion signal
    if conversions(r) == 0 and conv_value(r) == 0:
        return ("yellow", "Watch — No Conv. Attr.")

    # ROAS-primary
    if ro is not None:
        if ro < T["roas_red"]:
            return ("red", "Underperforming")
        if ro >= T["roas_green"]:
            return ("green", "Performing")

    # CPL-primary
    if cp is not None:
        if cp > T["cpl_red"] or ct < T["ctr_red"]:
            return ("red", "Underperforming")
        if cp <= T["cpl_green"] and ct >= T["ctr_green"]:
            return ("green", "Performing")

    if ish > 0 and ish < T["imp_share_red"]:
        return ("yellow", "Watch — Low Imp. Share")

    return ("yellow", "Watch")
```

---

## Report Sections (4 Tabs)

### Tab 1 — Overview
- KPI strip: Total Spend, Total Conversions, Best CPL, Best ROAS, Avg Impression Share
- Campaign triage table: Status · Campaign · Spend · Conversions · CPL · ROAS · CTR · Imp Share
- Row colours: green/amber/red/grey (same as Meta)

### Tab 2 — Alerts & Actions
- One alert card per red campaign, listing:
  - High CPL (vs ₹500 target)
  - Low ROAS (vs 1.5× minimum)
  - Low CTR (below 0.5%)
  - Low Search Impression Share (below 30% — budget or rank issue)
- Ad group gap signal: if best ad group CPL is <50% of worst, call it out
- Ranked quick wins list (auto-generated, same pattern as Meta)

### Tab 3 — Ad Group Breakdown
- Grouped under campaign expanders (red campaigns expanded by default)
- Columns: Ad Group · Spend · Budget % · Conversions · CPL · ROAS · CTR · Imp Share · Status
- Status badges: Scale / Cap Spend / Watch / Hold (same colour logic as Meta's ad set badges)

### Tab 4 — Ad Analysis
- Only for red campaigns (same as Meta creative analysis)
- Grouped by campaign → ad group → individual ads (one row per ad, direct API values)
- Columns: Ad Name · Ad Type · Spend · Impressions · CPM · CTR · Conversions · CPL · Verdict
- No frequency column for Search; include it only if Display/Video campaign
- Verdict: Keep + Scale ⭐ / Pause ❌ / Test further 🔍 (best CPL within ad group wins)

---

## Output Files

Same as Meta:
- `Google_Ads_Report_YYYY-MM-DD.html` — self-contained tabbed report
- `Google_Ads_Report_YYYY-MM-DD.xlsx` — 5 sheets: Summary, Campaigns, Ad Groups, Ads, Quick Wins

---

## Sidebar Inputs

```
Developer Token       [password]
OAuth2 Client ID      [password]
OAuth2 Client Secret  [password]
Refresh Token         [password]
Customer ID           [text]     e.g. 123-456-7890
Lookback Window       [slider]   7 / 14 / 30 / 60 / 90 days
[Run Analysis]        [primary button]
```

Add a collapsible "How to get credentials" expander (analogous to the Meta token instructions) covering:
1. Enable Google Ads API in Google Cloud Console
2. Create OAuth2 credentials (Desktop app type)
3. Generate refresh token via `generate_refresh_token.py` from the `google-ads-python` examples
4. Find your Customer ID in Google Ads UI (top-right corner, format XXX-XXX-XXXX)

---

## CLI Entry Point

```bash
python myhq_google_ads_agent.py \
  --developer-token <TOKEN> \
  --client-id <ID> \
  --client-secret <SECRET> \
  --refresh-token <TOKEN> \
  --customer-id <ID> \
  --days 30
```

Support env vars: `GOOGLE_ADS_DEVELOPER_TOKEN`, `GOOGLE_ADS_CLIENT_ID`, `GOOGLE_ADS_CLIENT_SECRET`, `GOOGLE_ADS_REFRESH_TOKEN`, `GOOGLE_ADS_CUSTOMER_ID`.

---

## Key Differences from Meta Analyser

| Aspect | Meta | Google Ads |
|---|---|---|
| Auth | Single access token | 4 credentials (dev token + OAuth2) |
| Query language | REST + field list | GAQL (SQL-like) |
| Cost field | `spend` (float, INR) | `cost_micros` (÷ 1,000,000) |
| CTR field | Already % | Decimal 0–1, multiply × 100 |
| Frequency | Available at all levels | Search: N/A · Display/Video: `average_frequency` |
| Ad sets | Ad Sets | Ad Groups |
| Creatives | Ad Name | Ad Name + Ad Type |
| Extra metric | — | Search Impression Share |
| Rate limiting | Error code 17, backoff | `RESOURCE_EXHAUSTED` gRPC status, backoff same pattern |
| Pagination | `paging.next` cursor | `search_stream()` streams all pages automatically |
