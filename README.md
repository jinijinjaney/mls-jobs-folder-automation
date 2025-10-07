# Google Sheets Automation Script (MLS Jobs)

This Google Apps Script automates folder creation and status tracking for MLS job records stored in Google Sheets.  
It integrates with Google Drive to automatically generate organized folders, update job notes, and move completed jobs to a separate sheet.

---

## Configuration

Make sure your Google Sheet contains the following sheets and columns:

- **Info** — Main sheet where job records are stored.  
- **Config** — Configuration sheet with key values.  
- **Completed Jobs** — Automatically created when jobs are marked complete.

### Config Sheet Keys
| Key | Description |
|-----|--------------|
| `ROOT_FOLDER_ID` | The Google Drive folder ID where all job folders will be created |

---

## Main Features

- **Automatic Folder Creation** — Creates folders and subfolders (`Field`, `Drafting`, `Deliverables`) in Drive.  
- **Job Notes Generation** — Creates a `Job Notes.txt` file summarizing job info.  
- **Completion Flow** — When a job’s **Status** changes to “Complete”, it:
  - Adds a timestamp in `Date Completed`
  - Moves the row to the `Completed Jobs` sheet
  - Renames the Drive folder to `[COMPLETE] Job Name`
- **Revert Flow** — If status is changed from “Complete” back to another state, it:
  - Removes completion date and entry in `Completed Jobs`
  - Restores the folder name
- **Menu Options** — Adds a custom menu “MLS Jobs” with:
  - `Process all rows`
  - `Process current row`
- **Auto Trigger** — Responds to edits (`onEdit`) to automatically update folders and notes.

---

## Key Variables

| Constant | Description |
|-----------|-------------|
| `TARGET_SHEET_NAME` | The main working sheet (`Info`) |
| `CONFIG_SHEET_NAME` | Configuration sheet (`Config`) |
| `COMPLETED_SHEET_NAME` | Sheet for completed jobs (`Completed Jobs`) |
| `HEADER_ROW` | Row containing headers (default: 1) |
| `FOLDER_URL_COLUMN` | Column to store folder link (default: G / 7) |
| `DELIVERABLES_URL_COLUMN` | Column for Deliverables folder link (default: H / 8) |
| `FOLDER_ID_COLUMN` | Column for folder ID (default: I / 9) |

---

## How It Works

1. Edit a row in the **Info** sheet.  
2. If `Bid` and `Client` have values, the script:
   - Creates or updates the Drive folder.
   - Writes the folder and subfolder URLs back to the sheet.
   - Generates or replaces `Job Notes.txt`.
3. If the **Status** changes to `Complete`, the job is timestamped, archived, and renamed.
4. If reverted, all completion changes are undone.

---

## Setup Instructions

1. Open your Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Paste the full code from `Code.gs`.
4. Create a **Config** sheet with:
   - ROOT_FOLDER_ID | (your-drive-folder-id-here)
5. Reload the sheet — a custom menu **MLS Jobs** will appear.
6. Use `Process all rows` to generate folders for all entries or manually update one row.

---

## Notes

- Make sure the script has access to Google Drive and Spreadsheet services.
- It’s best to install an **installable trigger** for `onEdit` under:
- **Triggers → Add Trigger → onEdit → From spreadsheet → On edit**
- The script uses **LockService** to avoid race conditions.

---
