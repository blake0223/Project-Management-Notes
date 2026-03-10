# Changelog

All notable changes to the Review Meeting Functions project will be documented in this file.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [1.0.0] - 2026-03-10

### Added
- Initial repository analysis and extraction from `Review Meeting Functions.json`
- Organized source tree under `src/` with three branches:
  - `src/config/` - Project manifest (`appsscript.json`)
  - `src/server/` - Server-side Google Apps Script files (`.gs`)
  - `src/html/` - Client-side HTML dialog templates
- `ANALYSIS.md` - Comprehensive documentation of JSON structure, all 7 files, function inventory, and spreadsheet schema
- `CHANGELOG.md` - This file, to track all future modifications

### Source Files Extracted
- `src/config/appsscript.json` - V8 runtime manifest (7 lines)
- `src/server/Job_Review_Confirmation.gs` - Core review logic: job selection, review submission (completed/draft), email notifications, post-project notes (623 lines)
- `src/server/OnOpen.gs` - Custom menu trigger with 3 menu items (8 lines)
- `src/server/PMT_Import.gs` - Import pipeline from external "Job Tracking" spreadsheet into Data!M:U (146 lines)
- `src/html/reviewModal.html` - Job selection + reviewer picker dialog (134 lines)
- `src/html/reviewSummary.html` - Per-job review card with prev/next nav, draft/completed toggle (231 lines)
- `src/html/postProjectNotes.html` - Two-panel post-project notes editor (182 lines)

### Notes
- Original JSON is a Google Apps Script project export containing a `files` array with 7 entries
- Each entry has: `id` (UUID), `name`, `type` (json/server_js/html), and `source` (raw code string)
- The project powers a Google Sheets add-on for construction/installation job review workflows
- Key integrations: Google Sheets (SpreadsheetApp), Gmail (MailApp), external spreadsheet import via ID
