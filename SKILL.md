---
name: apple-ads-daily-delta
description: Analyze Apple Ads daily spend changes and multi-day keyword performance, then create an actionable Excel report with keyword bid and search-term screening recommendations. Use for Apple Ads keyword, search-term, or weekly performance reviews.
---

# Apple Ads Daily Delta

Analyze Apple Ads performance CSV exports and create a black-and-white Excel report. Cover day-over-day spend changes and, when multiple dates are present, daily trends, ad-group efficiency, and actionable keyword/search-term screening. Use report dates and export metadata; do not use the computer's date to infer the reporting period.

## Inputs and dimensions

- Required: Apple Ads performance CSV with daily date and spend. Keyword and search-term fields enable keyword/query analysis.
- Optional: an earlier or richer performance CSV with ad-group ID/name for historical matching. Label those matches as references unless the current report directly contains the same dimensions.
- Current performance exports may include impressions, clicks, total/click/view installs, new downloads, CTR, conversion rates, average CPT/CPA/CPM, ad-group ID/name/status, keyword, keyword match type/status, search term/source, keyword maximum CPT, default maximum bid, target CPA, bidding strategy/type, campaign end date, and remaining days.
- A campaign ID/name is not an ad-group ID/name. Only show a campaign association when a source field directly connects them; never join on similar names or start dates.

Preserve source attribution labels. “安装次数（总计）” includes click- and view-attributed installs; do not relabel it as click installs. An Apple low-volume placeholder is a masked query bucket, not a search term that can be added or negated.

## Build the workbook

Use the bundled Node.js runtime and `@oai/artifact-tool` when available. The reusable builder is [scripts/build_report.mjs](scripts/build_report.mjs):

```text
node scripts/build_report.mjs --current <performance.csv> --output <report.xlsx> [--previous <historical-performance.csv>] [--target-cpa <user-confirmed-common-target>]
```

Save the workbook in a user-visible outputs folder and open it in Codex when useful. It contains:

- **关键词汇总** and **搜索词增量**: latest two available dates; show gross positive increases and explain how decreases offset them to produce net account change.
- **每日趋势**: spend, impressions, clicks, total/click/view installs, CTR, click-install conversion rate, total-install CPA, click-install CPA, CPC, and CPM. Recompute ratios from summed numerators/denominators rather than averaging row ratios.
- **广告组表现**: aggregate by direct ad-group ID and name; separate click and view attribution.
- **关键词出价建议**: one row per actionable unit, grouped by ad-group ID/name, keyword ID/text, match type, and match source. Include the current reported keyword max CPT, target CPA and its source, spend, impressions, clicks, click/view installs, click CPA, total CPA, CTR, click-install conversion rate, same-group/same-match peer medians, principal search terms, recommendation, rationale, suggested test range, ad-group name/ID, and review/rollback condition.
- **搜索词筛选动作**: visible query candidates to promote to exact match or manually review for negation, with triggering keyword and ad-group ID/name.

For each daily upload, deliver both the daily analysis workbook and the dashboard link. The dashboard's historical data page maintains the appendable, chronological daily keyword master table from imported CSV history. It includes only keywords that qualify as performance leaders or actionable review candidates, with the same decision evidence and ad-group dimensions as the workbook. The dashboard can export the currently selected history range as CSV or an Excel-compatible .xls file. A newer import for a date replaces that date's old snapshot, so corrected reports update the historical table without duplicating rows. Keep the daily workbook as a separate downloadable artifact and use its existing worksheet structure.

Keep tables black and white; mark keyword names in red where useful. Sort actionable recommendations ahead of observation rows. If ad-group fields are missing, say the group cannot be confirmed; do not substitute a campaign ID or historical name as current attribution.

## Period and metric rules

- Use the latest two dates present in the current report for day-over-day deltas. Do not assume they are consecutive calendar days if a date is missing.
- For multi-day review, use up to the latest seven complete dates. If export metadata shows the latest row date matches the export date, flag it as possibly incomplete and exclude it from period-based keyword recommendations; keep it visible in the daily trend.
- Show the selected date window and currency/time-zone metadata when available. If only one date exists, skip day-over-day calculations and ask for another date or report.
- Calculate CTR = clicks / impressions; click-install conversion rate = click-attributed installs / clicks; total-install CPA = spend / total installs; click-install CPA = spend / click-attributed installs; CPC = spend / clicks; CPM = spend / impressions × 1,000. Return blank/— when denominators are zero.
- Show total-install and click-install efficiency separately. The target CPA field may not use the same attribution basis as click-install CPA; state this caveat and treat the comparison as a screening signal to confirm against the account's optimization goal.
- Use `每日预算` only at the level directly represented in the source. It is not a per-keyword budget. Do not invent keyword budget allocations from a repeated campaign budget.

## Keyword and search-term actions

Make recommendations actionable but non-executing: name the exact keyword/search term, match type, ad-group name and ID, current reported max CPT, proposed test range when calculable, evidence, confidence/limitation, and a review/rollback condition. Do not make account changes.

Use the following conservative defaults as review flags, not universal truths; adjust them when the user supplies account-specific thresholds:

- Require at least 10 clicks and 3 click-attributed installs before using CPA alone to recommend a bid increase or decrease.
- **Raise candidate:** enabled keyword in a confirmed manual-CPT ad group; target CPA and current keyword max CPT are present; click-install CPA is at or below 80% of target; CTR is not less than half the median for keywords in the same ad group and match type. Suggest a small +5–10% test, not a guaranteed volume increase.
- **Lower candidate:** enabled manual-CPT keyword with sufficient sample and click-install CPA at or above 120% of target. Suggest a −10–15% test and inspect search-term mix first.
- **Spend/no-install review:** zero click-attributed installs with at least 20 clicks and spend at least 1.5× target CPA. Flag for search-term and relevance review; a lower-bid or pause decision remains a human choice.
- **Promote query to exact:** a visible query from broad or search-match traffic has at least 10 clicks, 3 click-attributed installs, and click-install CPA at or below 80% of target. Recommend “人工确认相关性后新增精准匹配” in the triggering ad group; do not create it automatically.
- **Negation review:** a visible query has at least 15 clicks, zero click-attributed installs, and spend at least the target CPA. Label “否词复核”, not an automatic negative. Do not recommend a specific negative from a low-volume placeholder.
- If target CPA is absent from the report, accept `--target-cpa` only when the user confirms one common target applies. Otherwise do not invent an account-average target or give an absolute bid change. With at least three comparable peers, label a keyword ≥20% better/worse than same-group/same-match median as a **relative** expansion/search-term review candidate; state that this is not a target-based bid decision.
- If current max CPT or direct ad-group mapping is absent; the ad group/keyword is paused; the strategy is automated; or volume is below threshold, use an observation/review label rather than a forced bid change.
- Compare CTR and click-install conversion rate with peers in the same ad group and match type to help distinguish weak engagement from weak post-click conversion. Peer rates diagnose where to inspect; they do not establish causation.
- `平均每次点击付费` versus reported keyword max CPT can be shown as context, but without impression share/rank data do not claim that the bid cap caused low impressions or that raising it will recover volume.
- For search terms, label winners as exact-match candidates and poor performers as manual negative-review candidates. Semantic relevance requires user/domain context; do not declare a query irrelevant based only on low conversions.

## Attribution and validation

- Group deltas by match source, keyword ID/text, and search term. For actionable bid suggestions, also keep ad-group ID/name and match type in the key so the same keyword in different groups is never combined.
- Label low-volume Apple buckets as masked and exclude them from exact/negative query actions.
- Keep spend growth separate from efficiency changes. Show click and view attribution separately; do not use total-install CPA alone to recommend keyword bids.
- Validate that daily totals reconcile to the source, the latest-day net change equals current spend minus prior spend, cross-report overlap agrees when two reports are supplied, and the workbook exports without formula errors.

Keep the final answer concise: report net spend movement, main performance/action signals, attribution and sample-size caveats, then provide both the daily workbook and the dashboard link. If the CSV was uploaded only as a ChatGPT project attachment and has not been saved into the folder authorized by the dashboard, explain that the dashboard page cannot ingest that attachment directly; do not claim its local history has been updated. The user can save the CSV into the linked folder and let the dashboard sync it, or import it through the dashboard.

