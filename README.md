# ASA Daily Delta

A Codex skill for comparing Apple Ads daily performance reports and producing an Excel report of spend increases by keyword and search term.

## What it does

- Compares the latest two dates present in the report.
- Summarizes keyword-level spend growth and keyword/search-term increases.
- Highlights rising keywords in red in a black-and-white workbook.
- Distinguishes gross increases from net account spend change.
- Marks historical campaign or ad group matches as references when the current report does not include those dimensions.

## Use

Install this skill in your Codex skills directory, then ask Codex to analyze an Apple Ads daily cost increase report. The bundled workbook builder can be run as follows:

```text
node scripts/build_report.mjs --current <current-report.csv> --output <report.xlsx> [--previous <previous-keyword-report.csv>] [--campaigns <campaign-summary.csv>]
```

The report requires a CSV with daily date, spend, keyword, and search-term fields. If the current export contains only one date, provide the previous day's report too.

See [SKILL.md](SKILL.md) for the complete analysis rules and verification steps.

