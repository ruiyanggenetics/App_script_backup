/*********************************
 * YangLab_Metabase_Backup (FULL Code.gs)
 *
 * What it does:
 * 1) Every 5 minutes on the clock (America/Chicago): 00,05,10,...55
 *    -> makes a full Google Sheets copy into BACKUP_FOLDER_ID
 *    -> names it Metabase_YYYY-MM-DD_HHmm
 * 2) Retention:
 *    - Keep everything within last KEEP_DAYS days
 *    - Older than KEEP_DAYS: keep only backups at 00:00; trash the rest
 *    - If multiple 00:00 backups exist for same date, keep one and trash duplicates
 *
 * How to use:
 * - Paste into a NEW standalone Apps Script project
 * - Set SOURCE_SPREADSHEET_ID and BACKUP_FOLDER_ID
 * - Run installBackupTriggers() once
 * - Run runBackupNow() to test immediately
 *********************************/

/************ CONFIG ************/
const SOURCE_SPREADSHEET_ID = '1soWIXL1usgLSI5YyHnw8gDvxf9i45PgAcbwTQFijc0I';
const BACKUP_FOLDER_ID      = '1zZqZQ-Yxp-ulEGgzNJUzZDkZWcdNKK7x';

const BACKUP_PREFIX = 'Metabase_';
const KEEP_DAYS = 365;
const TZ = 'America/Chicago';

// ScriptProperties keys
const PROP_LAST_SLOT = 'BACKUP_LAST_SLOT';

/**
 * Run ONCE manually to install triggers.
 * This creates a 1-minute trigger, then the script gates to exact 5-min slots (:00,:05,...).
 * This is the most reliable way to hit “clock times”.
 */
function installBackupTriggers() {
  uninstallBackupTriggers();

  ScriptApp.newTrigger('backupTick')
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log('Installed trigger: backupTick every 1 minute (runs only at 5-min slots).');
}

/**
 * Optional: remove triggers (run manually).
 */
function uninstallBackupTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'backupTick') {
      ScriptApp.deleteTrigger(t);
    }
  }
  Logger.log('Removed trigger(s) for backupTick.');
}

/**
 * Trigger entrypoint (runs every 1 minute, but gates to 5-minute marks).
 */
function backupTick() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const now = new Date();

    // Only run at minutes divisible by 5 in America/Chicago time
    const minute = Number(Utilities.formatDate(now, TZ, 'mm')); // "00".."59" -> Number ok
    if (minute % 5 !== 0) return;

    // Build a slot key like "2026-02-28_2315" (5-min slot aligned)
    const slot = Utilities.formatDate(now, TZ, 'yyyy-MM-dd_HHmm');

    // Dedupe: if already ran for this slot, skip
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(PROP_LAST_SLOT) === slot) return;
    props.setProperty(PROP_LAST_SLOT, slot);

    // Do the work
    backupOnce_(now);
    cleanupOldBackups_(now);

  } finally {
    lock.releaseLock();
  }
}

/**
 * Manual: force one backup now (ignores the 5-minute gate).
 * Use this to verify permissions + folder ID + sheet ID.
 */
function runBackupNow() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const now = new Date();
    backupOnce_(now);
    cleanupOldBackups_(now);
    Logger.log('runBackupNow complete.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Creates a full spreadsheet copy into BACKUP_FOLDER_ID.
 */
function backupOnce_(now) {
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const srcFile = DriveApp.getFileById(SOURCE_SPREADSHEET_ID);

  const stamp = Utilities.formatDate(now, TZ, 'yyyy-MM-dd_HHmm');
  const name = `${BACKUP_PREFIX}${stamp}`;

  const copy = srcFile.makeCopy(name, folder);
  Logger.log(`Backup created: ${name} (fileId=${copy.getId()})`);
}

/**
 * Retention policy:
 * - keep all backups newer than cutoff
 * - older than cutoff: keep only HHmm == 0000; trash others
 * - dedupe midnight backups per date: keep 1, trash extras
 */
function cleanupOldBackups_(now) {
  const folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  const files = folder.getFiles();

  // cutoffDate: now - KEEP_DAYS
  const cutoffMs = now.getTime() - KEEP_DAYS * 24 * 60 * 60 * 1000;

  // Track kept midnight backups by dateKey (yyyy-MM-dd)
  const keptMidnightByDate = {};

  let trashed = 0;
  let scanned = 0;

  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();

    if (!name.startsWith(BACKUP_PREFIX)) continue;

    const parsed = parseBackupTimestampFromName_(name);
    if (!parsed) continue;

    scanned++;

    const { dtChicagoMs, dateKey, hhmm } = parsed;

    // Keep everything within last KEEP_DAYS
    if (dtChicagoMs >= cutoffMs) continue;

    // Older than cutoff: keep only 00:00
    if (hhmm !== '0000') {
      f.setTrashed(true);
      trashed++;
      continue;
    }

    // hhmm == 0000: dedupe by date
    if (!keptMidnightByDate[dateKey]) {
      keptMidnightByDate[dateKey] = f.getId();
      continue;
    }

    // duplicate midnight for same date
    f.setTrashed(true);
    trashed++;
  }

  Logger.log(`cleanupOldBackups_: scanned=${scanned}, trashed=${trashed}`);
}

/**
 * Parse filename format: Metabase_YYYY-MM-DD_HHmm
 * Returns a Chicago-local timestamp converted into milliseconds for comparison.
 */
function parseBackupTimestampFromName_(name) {
  const re = new RegExp(`^${BACKUP_PREFIX}(\\d{4}-\\d{2}-\\d{2})_(\\d{4})$`);
  const m = name.match(re);
  if (!m) return null;

  const datePart = m[1]; // yyyy-MM-dd
  const hhmm = m[2];     // HHmm

  const yyyy = Number(datePart.slice(0, 4));
  const MM   = Number(datePart.slice(5, 7));
  const dd   = Number(datePart.slice(8, 10));
  const HH   = Number(hhmm.slice(0, 2));
  const mm   = Number(hhmm.slice(2, 4));

  // Build a Date as if it were in Chicago time:
  // We'll create a string in Chicago timezone and then parse back to Date via formatDate
  // We can’t directly construct Date in arbitrary TZ, so we approximate by using UTC base
  // and comparing relative ages. This is sufficient for retention.
  const dtUtc = Date.UTC(yyyy, MM - 1, dd, HH, mm, 0);

  // dtChicagoMs: keep as UTC ms derived from name; comparison vs cutoff is consistent
  return { dtChicagoMs: dtUtc, dateKey: datePart, hhmm };
}
