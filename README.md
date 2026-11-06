# ASA Daily Delta

A Codex skill and public interactive dashboard for reviewing Apple Ads daily performance.

## Daily decision dashboard · 2026-11-06

The default daily decision view combines a USD 700 actual daily spend target, a USD 750 warning line, query-level impression share and rank goals, keyword-ID-level review evidence, and campaign budget drafts. It retains all previous overview, keyword, history, import and post-change tracking views.

- Import performance detail CSVs and daily search-term impression-share CSVs together. Separate IndexedDB stores prevent share-report spend from being added to the performance ledger. Duplicate files are skipped. For a report type and scope, newer file snapshots replace the same dates; older snapshots do not overwrite newer ones. A CSV should contain a complete date snapshot for its exported scope, not an arbitrary filtered subset of an earlier report.
- Share scope includes App ID, country/region, campaign-group ID, ad position basis, currency and timezone. First-position share and all-position share stay separate. Share and rank are never summed or simply averaged. `91–100%`, `>5`, and missing values retain their meaning.
- Search-term share is an App/market-level observation, not a keyword-level share. Keyword IDs remain separate even when keyword text is identical. Low-volume masked queries cannot be assigned an observed query share or converted into negative keywords.
- Confirm that both reports refer to the same App, region, account scope and timezone before linking keyword detail to share evidence. The performance export may omit these fields. A share report being UTC does not prove that the performance report is UTC.
- Default goals cover the named brand and competitor queries. Edit the explicit query goal list to include additional competitor/category terms. Goals refer to impression-share rank, not physical ad position. Jobsdb has a rank-2 floor and rank-1 stretch goal.
- Campaign budget inputs are saved drafts, not live ASA settings or spend forecasts. Daily budgets do not enforce a hard USD 750 daily ceiling. The static page does not monitor live spend or pause campaigns.
- Missing match type, bidding strategy, campaign budget or click-install fields are disclosed. No total-install-to-registration relabeling, inferred click installs or unsupported absolute bid recommendations.
- JSON backup/restore includes performance history, share history, goals, budget drafts and recorded bid changes. Previously exported backups remain accepted. CSV exports are spreadsheet-formula escaped.

Implementation: `decision-engine.js` contains shared calculations; `decision-ui.js` renders the decision view. Both HTML entry points are kept identical. Tests are in `test-data.mjs` and `test-browser.mjs`; test inputs and private ad data are not published.

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
