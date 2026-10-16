---
name: apple-ads-daily-delta
description: Compare daily Apple Ads spend reports and create a black-and-white Excel table of rising keywords, search terms, and spend changes. Use when the user asks for a daily Apple Ads cost increase or keyword delta report.
---

# Apple Ads Daily Delta

Create a reusable day-over-day spend report from Apple Ads CSV exports. Use the latest two dates present in the current report; do not assume the computer's date identifies the reporting day.

## Inputs

- Required: the current Apple Ads performance CSV with daily date, spend, keyword, and search-term fields.
- Optional: a previous or richer keyword report with ad group name and ID, for historical matching references.
- Optional: a campaign summary CSV with campaign names and IDs. Only use a direct, supported association; do not infer parent campaigns from similar names or start dates.

If the current CSV contains only one date, compare it manually with the previous daily export or ask for a report covering at least two dates. Preserve the report's actual attribution field (for example, total installs); do not relabel it as click installs.

## Build the report

Use the bundled Node.js runtime and `@oai/artifact-tool` when available. The reusable workbook builder is [scripts/build_report.mjs](scripts/build_report.mjs). Invoke it with the current file and output path; pass historical and campaign files when available:

```text
node scripts/build_report.mjs --current <current-report.csv> --output <report.xlsx> [--previous <previous-keyword-report.csv>] [--campaigns <campaign-summary.csv>]
```

In Codex desktop, use the workspace dependency paths for Node.js and its packages. Ensure the report output is saved in a user-visible outputs folder and open it in Codex when useful.

The workbook contains:

- **关键词汇总**: keywords with net spend growth, compared across the latest two report dates.
- **搜索词增量**: keyword/search-term combinations whose spend increased.

Keep the table black and white. Mark increased keyword text in red. Sort increases from largest to smallest. Include spend, change, clicks, installs, and the principal search terms contributing to growth. Explain the difference between gross positive increases and net account spend change when decreases offset increases.

## Attribution and interpretation

- Group rows by match source, keyword ID/text, and search term when those fields exist.
- Apple low-volume placeholders are masked query buckets, not real search terms; label them clearly.
- If the current report omits campaign or ad group dimensions, mark any previous-day keyword/search-term match as a **historical reference**, not a confirmed current-day assignment. Show ambiguous multiple-group matches as ambiguous.
- Show a campaign ID only when the source directly links it to the group or campaign. Leave it unconfirmed otherwise; never use naming similarity or campaign start date as proof.
- Separate spend growth from efficiency changes. Use only available metrics and name their attribution basis.
- Verify that daily totals match the source CSV, that the reported net increase equals current-day spend minus prior-day spend, and that the workbook opens without formula errors.

Keep the answer concise: state the day-over-day net spend change, the largest drivers, attribution limits, and link the workbook.

