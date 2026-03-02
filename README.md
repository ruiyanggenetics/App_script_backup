# App_script_backup (YangLab Metabase backup)

This repo stores Google Apps Script code used to **auto-backup** the `YangLab_Metabase` Google Spreadsheet to a Drive folder on a fixed schedule, with retention cleanup.

---

## What this script does

### Backup schedule
- Runs **every 5 minutes on the clock** (America/Chicago): `:00, :05, :10, ... :55`
- Creates a **full Google Sheets file copy** into your backup folder
- Names backups like:
  - `Metabase_YYYY-MM-DD_HHmm` (example: `Metabase_2026-03-01_2345`)

### Retention policy
- Keep **all backups** within the last `KEEP_DAYS` days
- For backups **older** than `KEEP_DAYS`:
  - Keep **only** the midnight backup (`HHmm == 0000`)
  - Trash all other times
  - If multiple midnight backups exist for the same date, keep **one** and trash duplicates

---

## Configuration

In `Code.gs` set:

```js
const SOURCE_SPREADSHEET_ID = '...'; // spreadsheet ID (between /d/ and /edit)
const BACKUP_FOLDER_ID      = '...'; // Drive folder where backups are stored

const BACKUP_PREFIX = 'Metabase_';
const KEEP_DAYS = 365;
const TZ = 'America/Chicago';
