# Review Meeting Functions - JSON Analysis

## Overview

The file `Review Meeting Functions.json` is a **Google Apps Script project export** (`.clasp`-style) containing a complete Google Sheets add-on for managing **job review meetings** in a construction/installation project management workflow.

- **Format**: Google Apps Script project JSON (single `files` array with 7 entries)
- **Runtime**: V8 (modern JavaScript)
- **Timezone**: America/New_York
- **Total source**: ~1,335 lines across 7 files

---

## File Inventory

| # | Name | Type | Extension | Lines | Location |
|---|------|------|-----------|-------|----------|
| 0 | `appsscript` | json | `.json` | 7 | `src/config/` |
| 1 | `Job Review Confirmation` | server_js | `.gs` | 623 | `src/server/` |
| 2 | `OnOpen` | server_js | `.gs` | 8 | `src/server/` |
| 3 | `reviewModal` | html | `.html` | 134 | `src/html/` |
| 4 | `reviewSummary` | html | `.html` | 231 | `src/html/` |
| 5 | `PMT Import` | server_js | `.gs` | 146 | `src/server/` |
| 6 | `postProjectNotes` | html | `.html` | 182 | `src/html/` |

---

## Directory Structure (Extracted)

```
src/
├── config/
│   └── appsscript.json          # Project manifest (timezone, runtime, dependencies)
├── server/
│   ├── Job_Review_Confirmation.gs  # Core business logic (623 lines)
│   ├── OnOpen.gs                   # Menu trigger (8 lines)
│   └── PMT_Import.gs              # Data import from external spreadsheet (146 lines)
└── html/
    ├── reviewModal.html            # Job selection dialog (134 lines)
    ├── reviewSummary.html          # Review detail/submit dialog (231 lines)
    └── postProjectNotes.html       # Post-project notes dialog (182 lines)
```

---

## Functional Breakdown

### 1. Config (`appsscript.json`)
- Sets timezone to `America/New_York`
- Uses V8 runtime
- Logs exceptions to Stackdriver
- No external dependencies declared

### 2. Menu & Entry Points (`OnOpen.gs`)
Adds a custom **"Job Review Tools"** menu to Google Sheets with three items:
- **Begin Job Review** -> `openReviewModal()`
- **Add Post-Project Note to Job** -> `openPostProjectNotesModal()`
- **Refresh Jobs** -> `loadJobTrackingIntoData()`

### 3. Core Business Logic (`Job_Review_Confirmation.gs`)
The main file, containing 623 lines organized into these functional areas:

#### 3a. Configuration & Utilities (Lines 1-69)
- Sheet name constants: `REVIEW_SHEET = 'Review Queue'`, `DATA_SHEET = 'Data'`
- Helper functions: `htmlEscape_()`, `normalizeUrl_()`, `fmtDate_()`, `parseMoney_()`, `fmtCurrency_()`, `nameKey_()`

#### 3b. URL Extraction (Lines 71-102)
- `getEmbeddedUrl_()` - Extracts hyperlinks from rich text values, `HYPERLINK()` formulas, or display values

#### 3c. Row Management (Lines 104-199)
- `findHighestBlankRow_()` - Finds first blank row in A:L (main data area)
- `findHighestBlankRowDraft_()` - Finds first blank row in AA:AC (draft area)
- `getDraftNotesByJobName_()` - Retrieves draft notes from Data!AA:AC
- `removeDraftByJobName_()` - Clears draft row when finalized
- `upsertDraft_()` - Insert or update draft notes in Data!AA:AC

#### 3d. Data Cleanup (Lines 201-226)
- `clearDataMUByPipedriveIds_()` - Clears import staging area (M:U) for reviewed jobs by Pipedrive ID

#### 3e. Email Recipients (Lines 228-277)
- `collectRecipients_()` - Builds recipient list from:
  - All emails in Data!W (always CC'd)
  - Head of Sales in Data!Y2
  - Shop-specific emails from Data!Z (matched via salesman name in Data!X)

#### 3f. Initial Data Loading (Lines 279-326)
- `getInitialData()` - Builds the payload for the review modal:
  - Owner from Review Queue!A2
  - Jobs from Review Queue!B:J (filtered by owner/salesman match)
  - Reviewers from Data!V2:V

#### 3g. Job Details (Lines 328-372)
- `getJobDetailsByRows()` - Fetches full details for selected rows from Review Queue (B:J), including draft notes

#### 3h. Review Summary Dialog (Lines 374-386)
- `openReviewSummary()` - Templated HTML dialog with job data payload

#### 3i. Email Generation (Lines 388-450)
- `buildEmailBody_()` - Generates styled HTML email with:
  - Tabular layout per job
  - DRAFT/COMPLETED status badges
  - Material list links, cost data, mandays, notes

#### 3j. Job Submission (Lines 452-539)
- `processSingleJobReview()` - Core submission handler:
  - **Draft mode**: Saves to Data!AA:AC via `upsertDraft_()`
  - **Completed mode**: Writes to Data!A:L (main area), clears staging M:U, removes draft
  - Sends email notification to collected recipients

#### 3k. Post-Project Notes (Lines 541-623)
- `getPostProjectJobs()` - Lists all completed jobs from Data!A with existing post-notes from Data!AD
- `getPostProjectJobDetails()` - Full detail view (A through AD)
- `savePostProjectNote()` - Writes note text to Data!AD

### 4. Data Import (`PMT_Import.gs`)
- `loadJobTrackingIntoData()` - Imports from external spreadsheet ("Job Tracking" tab) into Data!M:U
- Source columns mapped: A (Job Name), B (Pipedrive ID), C (Start Date), E (Salesman), G (Deal Value), L (Est Material), O (Est Mandays), X (Issue Notes), Y (Material List Link)
- Skips jobs already present in Data!B (by Pipedrive ID)
- Preserves hyperlinks from rich text values in material list column

### 5. UI Dialogs (HTML files)

#### 5a. `reviewModal.html` - Job Selection
- Displays unreviewed jobs with checkboxes
- Reviewer selection via chip-style checkboxes (from Data!V)
- Validates at least one job + one reviewer selected
- Calls `openReviewSummary()` on submit

#### 5b. `reviewSummary.html` - Review & Submit
- Card-based per-job review interface with prev/next navigation
- Shows: project name, salesman, Pipedrive ID, dates, financials, material list
- Editable fields: PM Notes (textarea), material list URL override
- Mode toggle: Completed vs. Draft
- Submits one job at a time via `processSingleJobReview()`
- Auto-advances to next unsubmitted job

#### 5c. `postProjectNotes.html` - Post-Project Notes
- Two-panel layout: job list (left) + details/notes (right)
- Loads jobs from `getPostProjectJobs()`
- Saves notes to Data!AD via `savePostProjectNote()`

---

## Spreadsheet Schema

### "Review Queue" Sheet
| Column | Field |
|--------|-------|
| A | Owner (salesman filter) |
| B | Project Name |
| C | Pipedrive ID |
| D | Salesman |
| E | Start Date |
| F | Deal Value |
| G | Estimated Material Cost |
| H | Estimated Mandays |
| I | Material List Link |
| J | Material List Notes |

### "Data" Sheet
| Column(s) | Field |
|-----------|-------|
| A | Project Name |
| B | Pipedrive ID |
| C | Salesman |
| D | Reviewed By |
| E | Project Status |
| F | Start Date |
| G | Material List Link |
| H | Estimated Material Cost |
| J | Estimated Mandays |
| L | PM Notes |
| M-U | Import staging area (Job Tracking data) |
| V | Reviewers list |
| W | CC email addresses |
| X | Salesman names (for shop email lookup) |
| Y | Head of Sales email |
| Z | Shop-specific emails |
| AA-AC | Draft notes staging (Name, Deal Value, Notes) |
| AD | Post-Project Notes |
