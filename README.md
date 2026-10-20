# ASA Daily Delta

A Codex skill and public interactive dashboard for reviewing Apple Ads daily performance.

## Public dashboard

Open the [Apple Ads daily dashboard](https://yorihan2.github.io/ASA-daily-delta/). The same dashboard file is available as [apple-ads-daily-dashboard.html](apple-ads-daily-dashboard.html).

The dashboard includes three views: overview, historical data, and import history. It can read multiple CSVs at once. To avoid selecting files every day, choose a local folder once; when the dashboard is opened it scans that folder, and while the tab is active it checks for new or updated CSVs every two minutes. Browser folder permission may need to be renewed if the browser revokes it.

Imported reports are saved in IndexedDB in that browser on that device. The dashboard deduplicates identical files and replaces a date's earlier snapshot when a newer report for that date is imported. The historical data view includes a chronological **daily keyword performance master table** with qualifying high performers and review/action candidates, including ad-group IDs, bids, CPA metrics, search-term clues, recommendation rationale, and suggested test ranges. Filter by date and export that view as CSV or an Excel-compatible workbook. Use the JSON backup and restore controls in the import history view to move raw history to another browser or device.

The GitHub Pages dashboard is a static web page and cannot directly access attachments uploaded to a ChatGPT project. Save those CSVs into the chosen local folder, or import them through the dashboard. The CSVs are processed in your browser and are not uploaded to GitHub. Bidding suggestions are advisory and do not change Apple Ads settings.

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
