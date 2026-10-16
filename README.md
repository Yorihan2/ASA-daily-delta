# ASA Daily Delta

A Codex skill and local interactive dashboard for reviewing Apple Ads daily performance.

## Skill

The skill compares Apple Ads reports and produces an Excel report of spend increases by keyword and search term.

- [Skill instructions](SKILL.md)
- [Excel report builder](scripts/build_report.mjs)

Install the skill in your Codex skills directory, then ask Codex to analyze an Apple Ads daily cost increase report. The bundled workbook builder can be run as follows:

```text
node scripts/build_report.mjs --current <current-report.csv> --output <report.xlsx> [--previous <previous-keyword-report.csv>] [--campaigns <campaign-summary.csv>]
```

The report requires a CSV with daily date, spend, keyword, and search-term fields. If the current export contains only one date, provide the previous day's report too.

## Interactive dashboard

Download [apple-ads-daily-dashboard.html](apple-ads-daily-dashboard.html) and open it in a browser. Upload an Apple Ads performance CSV; the dashboard automatically detects report columns, compares the latest available dates, and provides keyword spend and bidding strategy suggestions based on the report's target CPA where available.

CSV processing runs in the browser on your device. Bidding suggestions are advisory and do not change Apple Ads settings.

See [SKILL.md](SKILL.md) for the complete analysis rules and verification steps.
