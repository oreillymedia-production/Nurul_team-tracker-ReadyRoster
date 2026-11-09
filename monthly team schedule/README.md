# Monthly team schedule automation

This project builds the Nurul December schedule without changing either source Google Sheet. It:

- exports the two Google Sheets with read-only credentials;
- filters the December schedule to producer `NI`;
- adds matched instructor email addresses;
- checks each published live-event page;
- keeps unpublished events, highlighted in peach; and
- writes a readiness report for events exactly three calendar days ahead.

## One-time setup

The project uses Codex's bundled spreadsheet runtime. From this project directory, link its package directory:

```sh
ln -s /Users/nishak/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules node_modules
```

Never commit `node_modules`, downloaded source files, credentials, or generated reports.

## Refresh private Google Sheet sources

Create a short-lived, read-only Google access token with access to the two source Sheets, then run:

```sh
export GOOGLE_ACCESS_TOKEN='…'
npm run refresh-sources
```

This exports `data/schedule.xlsx` and `data/contacts.xlsx` locally. It does not edit either Google Sheet.

## Build the schedule and readiness report

```sh
export TEMPLATE_XLTX='/absolute/path/to/NI_Monthly_Team_Schedule_Template(AutoRecovered).xltx'
npm run build
```

The command writes the filled schedule and `artifacts/readiness-YYYY-MM-DD.json`. The readiness report selects only events scheduled exactly three calendar days after the run date.

For a reproducible check date:

```sh
npm run check-readiness -- --as-of 2026-11-28
```

## Automate in CI

Schedule the same commands daily in your CI provider. Store `GOOGLE_ACCESS_TOKEN` as a read-only secret and save the generated artifact. Rotate the token regularly. The project does not send email or modify the source schedules.
