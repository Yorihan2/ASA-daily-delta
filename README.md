# ASA Daily Delta

A Codex skill and public interactive dashboard for reviewing Apple Ads daily performance.

## Public dashboard

Open the [Apple Ads daily dashboard](https://yorihan2.github.io/ASA-daily-delta/). The same dashboard file is available as [apple-ads-daily-dashboard.html](apple-ads-daily-dashboard.html).

Upload an Apple Ads performance CSV; the dashboard automatically detects report columns, compares the latest available dates, and provides keyword spend and bidding strategy suggestions based on the report's target CPA where available. CSV processing runs in the browser on your device. Bidding suggestions are advisory and do not change Apple Ads settings.

## Skill

The skill compares Apple Ads reports and produces an Excel report of spend increases by keyword and search term.

- [Skill instructions](SKILL.md)
- [Excel report builder](scripts/build_report.mjs)

Install this skill in your Codex skills directory, then ask Codex to analyze an Apple Ads daily cost increase report. The bundled workbook builder can be run as follows:

```text
node scripts/build_report.mjs --current <current-report.csv> --output <report.xlsx> [--previous <previous-keyword-report.csv>] [--campaigns <campaign-summary.csv>]
```

The report requires a CSV with daily date, spend, keyword, and search-term fields. If the current export contains only one date, provide the previous day's report too.

See [SKILL.md](SKILL.md) for the complete analysis rules and verification steps.
