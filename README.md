# Speaker outreach automation workflow

Reusable Power Automate companion for creating speaker-outreach and overdue-materials follow-up drafts from an Excel event tracker.

## Deliverable

`outputs/speaker_outreach/Speaker outreach tracker - flow ready.xlsx` contains:

- an `EventTracker` Excel table with stable event IDs, status fields, To/CC fields, course type, registration page, and draft timestamps;
- a `Power Automate setup` sheet with manual-flow configuration and separate wording for new and returning courses;
- a `Follow-up drafts` sheet with reviewable drafts for materials whose deadline has passed.

## Safety behavior

The documented flow uses Outlook's **Draft an email message** action only. It has no send-email action. Rows are eligible only after review, and records missing a contact or instructor, or marked as held, are excluded.

## Build notes

`build_flow_ready_tracker.mjs` regenerates the workbook from the source tracker at the local path configured near the top of that file. It requires the Codex workspace spreadsheet runtime.
