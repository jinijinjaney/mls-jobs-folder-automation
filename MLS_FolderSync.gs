/* CONFIG */
const TARGET_SHEET_NAME = 'Info';
const CONFIG_SHEET_NAME = 'Config';
const HEADER_ROW = 1;
const FOLDER_URL_COLUMN = 7;            // G
const DELIVERABLES_URL_COLUMN = 8;      // H
const FOLDER_ID_COLUMN = 9;             // I
const COMPLETED_SHEET_NAME = 'Completed Jobs';
const DATE_COMPLETED_HEADER = 'Date Completed';
const RENAME_COMPLETE_CONFIG_KEY = 'RENAME_FOLDER_ON_COMPLETE';

const COLUMNS = {
  BID: 1,       // A
  CLIENT: 2,    // B
  TMK: 3,       // C
  ADDRESS: 4,   // D
  STATUS: 5     // E
};

/* Helpers */
function cleanName(name) {
  if (!name) return 'unnamed';
  return name.toString().replace(/[\/\\\?\%\*\:\|\"<>\.]/g, '-').trim();
}

function getRootFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) throw new Error(`Config sheet "${CONFIG_SHEET_NAME}" not found`);

  const values = config.getDataRange().getValues();
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim() === 'ROOT_FOLDER_ID') {
      const folderId = values[r][1];
      if (!folderId) throw new Error("ROOT_FOLDER_ID is empty in Config sheet");
      return DriveApp.getFolderById(folderId);
    }
  }
  throw new Error("ROOT_FOLDER_ID not found in Config sheet");
}

function getConfigValue(key, defaultValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) return defaultValue;
  const values = config.getDataRange().getValues();
  for (let r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim() === key) {
      const v = values[r][1];
      if (typeof defaultValue === 'boolean') {
        if (String(v).toLowerCase() === 'true' || v === 1) return true;
        if (String(v).toLowerCase() === 'false' || v === 0) return false;
      }
      return v !== undefined && v !== null ? v : defaultValue;
    }
  }
  return defaultValue;
}

function getHeaderIndex(sheet, headerName) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === String(headerName).trim().toLowerCase()) return i + 1;
  }
  return -1;
}

function ensureHeaderExists(sheet, headerName) {
  let idx = getHeaderIndex(sheet, headerName);
  if (idx !== -1) return idx;
  const newCol = Math.max(1, sheet.getLastColumn()) + 1;
  sheet.getRange(HEADER_ROW, newCol).setValue(headerName);
  return newCol;
}

/* Drive helpers */
function findOrCreateFolderInParent(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function replaceFile(folder, fileName, content) {
  // delete all files with the same name, then create new
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    const f = files.next();
    try { f.setTrashed(true); } catch (e) { /* ignore */ }
  }
  folder.createFile(fileName, content);
}

/* Menu */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('MLS Jobs')
    .addItem('Process all rows', 'processAllRows')
    .addItem('Process current row', 'processCurrentRowMenu')
    .addToUi();
}

function processCurrentRowMenu() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const row = sheet.getActiveRange().getRow();
  if (row <= HEADER_ROW) {
    SpreadsheetApp.getUi().alert('Select a data row.');
    return;
  }
  createOrUpdateFolderForRow(sheet, row);
  SpreadsheetApp.getUi().alert('Processed row ' + row);
}

/* onEdit trigger (installable recommended) */
function onEdit(e) {
  try {
    if (!e) return;
    const sheet = e.source.getActiveSheet();
    if (!sheet || sheet.getName() !== TARGET_SHEET_NAME) return;

    const row = e.range.getRow();
    if (row <= HEADER_ROW) return;

    const editedCol = e.range.getColumn();

    // If user changed STATUS column, handle complete or revert
    if (editedCol === COLUMNS.STATUS) {
      const statusValue = String(sheet.getRange(row, COLUMNS.STATUS).getDisplayValue()).trim();
      if (statusValue === 'Complete') {
        handleCompleteFlow(sheet, row);
      } else {
        handleRevertFromComplete(sheet, row);
        createOrUpdateFolderForRow(sheet, row);
      }
      return;
    }

    // if user edited other key columns, update folder and notes
    const allowedCols = [COLUMNS.BID, COLUMNS.CLIENT, COLUMNS.TMK, COLUMNS.ADDRESS];
    if (!allowedCols.includes(editedCol)) return;

    const bid = String(sheet.getRange(row, COLUMNS.BID).getDisplayValue()).trim();
    const client = String(sheet.getRange(row, COLUMNS.CLIENT).getDisplayValue()).trim();
    if (!bid || !client) return;

    createOrUpdateFolderForRow(sheet, row);

  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

/* Main folder/notes logic */
function createOrUpdateFolderForRow(sheet, row) {
  try {
    function safeGet(r, c) {
      const cell = sheet.getRange(r, c);
      const d = cell.getDisplayValue();
      return (d !== undefined && d !== null && d !== '') ? d : cell.getValue();
    }

    const bid = safeGet(row, COLUMNS.BID);
    const client = safeGet(row, COLUMNS.CLIENT);
    const folderName = cleanName(bid + ' ' + client);

    // Get root MLS folder from Config
    const root = getRootFolder();

    // Ensure year folder exists inside MLS (2025)
    const yearFolderName = "2025";
    const yearFolders = root.getFoldersByName(yearFolderName);
    const yearFolder = yearFolders.hasNext() ? yearFolders.next() : root.createFolder(yearFolderName);

    // Reuse folder if ID already stored in row
    let mainFolder;
    let folderId = sheet.getRange(row, FOLDER_ID_COLUMN).getValue();

    if (folderId) {
      try {
        mainFolder = DriveApp.getFolderById(folderId);
      } catch (err) {
        Logger.log("Invalid folderId in row " + row + ", creating new.");
        mainFolder = findOrCreateFolderInParent(yearFolder, folderName);
        sheet.getRange(row, FOLDER_ID_COLUMN).setValue(mainFolder.getId());
      }
    } else {
      mainFolder = findOrCreateFolderInParent(yearFolder, folderName);
      sheet.getRange(row, FOLDER_ID_COLUMN).setValue(mainFolder.getId());
    }

    // Subfolders
    const subfolderNames = ['Field', 'Drafting', 'Deliverables'];
    const subfolders = {};
    subfolderNames.forEach(name => {
      subfolders[name] = findOrCreateFolderInParent(mainFolder, name);
    });

    // Notes file (ensure only 1 exists by using replaceFile)
    const notesContent = buildNotes(sheet, row);
    replaceFile(mainFolder, 'Job Notes.txt', notesContent);

    // Write URLs back to sheet
    try { sheet.getRange(row, FOLDER_URL_COLUMN).setValue(mainFolder.getUrl()); } catch (e) {}
    try { sheet.getRange(row, DELIVERABLES_URL_COLUMN).setValue(subfolders['Deliverables'].getUrl()); } catch (e) {}

  } catch (err) {
    Logger.log('createOrUpdateFolderForRow error: ' + err);
  }
}

/* Notes formatting */
function buildNotes(sheet, row) {
  function safeGet(r, c) {
    const cell = sheet.getRange(r, c);
    const d = cell.getDisplayValue();
    return (d !== undefined && d !== null && d !== '') ? d : cell.getValue();
  }

  const bid = safeGet(row, COLUMNS.BID);
  const client = safeGet(row, COLUMNS.CLIENT);
  const tmk = safeGet(row, COLUMNS.TMK);
  const address = safeGet(row, COLUMNS.ADDRESS);
  const status = safeGet(row, COLUMNS.STATUS);

  const dateCol = getHeaderIndex(sheet, DATE_COMPLETED_HEADER);
  const dateCompletedVal = dateCol > 0 ? safeGet(row, dateCol) : '';

  const updated = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  const lines = [
    '      JOB INFORMATION',
    '------------------------------',
    `Bid #:            ${bid || ''}`,
    `Client:           ${client || ''}`,
    `TMK:              ${tmk || ''}`,
    `Address:          ${address || ''}`,
    '',
    '      STATUS DETAILS',
    '------------------------------',
    `Status:           ${status || ''}`,
    `Date Completed:   ${dateCompletedVal || 'N/A'}`,
    '',
    '      SYSTEM INFORMATION',
    '------------------------------',
    `Row #:            ${row}`,
    `Last Updated:     ${updated}`,
  ];

  return lines.join('\n');
}

/* Completed sheet helpers */
function ensureCompletedSheetHeaders(sourceSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let comp = ss.getSheetByName(COMPLETED_SHEET_NAME);
  const sourceLastCol = Math.max(1, sourceSheet.getLastColumn());
  const sourceHeaders = sourceSheet.getRange(HEADER_ROW, 1, 1, sourceLastCol).getValues()[0].map(String);

  if (!comp) {
    comp = ss.insertSheet(COMPLETED_SHEET_NAME);
    // copy source headers
    comp.getRange(HEADER_ROW, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
    // ensure Date Completed exists in comp
    comp.getRange(HEADER_ROW, sourceHeaders.length + 1).setValue(DATE_COMPLETED_HEADER);
  } else {
    // ensure Date Completed header exists
    if (getHeaderIndex(comp, DATE_COMPLETED_HEADER) === -1) {
      comp.insertColumnAfter(comp.getLastColumn());
      comp.getRange(HEADER_ROW, comp.getLastColumn()).setValue(DATE_COMPLETED_HEADER);
    }
    // if comp has fewer cols than source, append missing headers so column positions match
    if (comp.getLastColumn() < sourceHeaders.length) {
      const startCol = comp.getLastColumn() + 1;
      const missing = sourceHeaders.slice(comp.getLastColumn());
      comp.getRange(HEADER_ROW, startCol, 1, missing.length).setValues([missing]);
    }
  }
  return comp;
}

function findRowsInSheetByFolderId(sheet, folderId, folderIdHeaderName) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return [];
  let colIndex = getHeaderIndex(sheet, folderIdHeaderName);
  if (colIndex === -1) colIndex = FOLDER_ID_COLUMN; // fallback
  const values = sheet.getRange(HEADER_ROW + 1, colIndex, lastRow - HEADER_ROW, 1).getValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(folderId)) rows.push(HEADER_ROW + 1 + i);
  }
  return rows;
}

/* Completion flow */
function handleCompleteFlow(sheet, row) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000); // wait up to 5s
  } catch (e) {
    Logger.log('Could not acquire lock for completion flow.');
    return;
  }

  try {
    // ensure folder and notes exist and folderId set
    createOrUpdateFolderForRow(sheet, row);

    const folderId = sheet.getRange(row, FOLDER_ID_COLUMN).getValue();
    if (!folderId) {
      Logger.log('No folderId for row ' + row + '. Aborting completion flow.');
      lock.releaseLock();
      return;
    }

    // ensure Date Completed header in source sheet
    const dateColIndex = ensureHeaderExists(sheet, DATE_COMPLETED_HEADER);

    // timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // write Date Completed in original sheet
    sheet.getRange(row, dateColIndex).setValue(timestamp);

    // Completed sheet setup
    const comp = ensureCompletedSheetHeaders(sheet);

    // get header names for mapping
    const sourceLastCol = Math.max(1, sheet.getLastColumn());
    const sourceHeaders = sheet.getRange(HEADER_ROW, 1, 1, sourceLastCol).getValues()[0].map(String);
    const compLastCol = comp.getLastColumn();
    const compHeaders = comp.getRange(HEADER_ROW, 1, 1, compLastCol).getValues()[0].map(String);

    // build a mapped row array to match comp headers
    const sourceValues = sheet.getRange(row, 1, 1, sourceLastCol).getValues()[0];
    const rowToPut = new Array(compLastCol).fill('');
    for (let c = 0; c < compHeaders.length; c++) {
      const h = String(compHeaders[c]).trim();
      // find index in source headers
      for (let s = 0; s < sourceHeaders.length; s++) {
        if (String(sourceHeaders[s]).trim().toLowerCase() === h.toLowerCase()) {
          rowToPut[c] = sourceValues[s];
          break;
        }
      }
    }
    // ensure Date Completed column in comp has timestamp
    const compDateCol = getHeaderIndex(comp, DATE_COMPLETED_HEADER);
    if (compDateCol > 0) rowToPut[compDateCol - 1] = timestamp;

    // find existing rows in comp by folderId
    const folderIdHeaderName = sourceHeaders[FOLDER_ID_COLUMN - 1] || 'Folder ID';
    const matches = findRowsInSheetByFolderId(comp, folderId, folderIdHeaderName);

    if (matches.length === 0) {
      // append new row
      comp.appendRow(rowToPut);
    } else {
      // update first match, delete other duplicates
      const firstRow = matches[0];
      comp.getRange(firstRow, 1, 1, rowToPut.length).setValues([rowToPut]);
      // delete duplicates from bottom to top
      for (let i = matches.length - 1; i >= 1; i--) {
        comp.deleteRow(matches[i]);
      }
    }

    // rename folder to add prefix if enabled
    const renameOnComplete = getConfigValue(RENAME_COMPLETE_CONFIG_KEY, true);
    if (renameOnComplete) {
      try {
        const folder = DriveApp.getFolderById(folderId);
        const currentName = folder.getName();
        if (!/^\[(complete|completed)\]/i.test(currentName)) {
          folder.setName('[COMPLETE] ' + currentName);
        }
      } catch (err) {
        Logger.log('Failed to rename folder for row ' + row + ': ' + err);
      }
    }

    // ensure only one Job Notes file (replaceFile will delete all then create)
    try {
      const folder = DriveApp.getFolderById(folderId);
      const notesContent = buildNotes(sheet, row);
      replaceFile(folder, 'Job Notes.txt', notesContent);
    } catch (err) {
      Logger.log('Failed to update job notes after completion: ' + err);
    }

  } catch (err) {
    Logger.log('handleCompleteFlow error: ' + err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* Revert flow: when status changed away from Complete */
function handleRevertFromComplete(sheet, row) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
  } catch (e) {
    Logger.log('Could not acquire lock for revert flow.');
    return;
  }

  try {
    const folderId = sheet.getRange(row, FOLDER_ID_COLUMN).getValue();
    if (!folderId) {
      // nothing to revert if no folder
      return;
    }

    // Remove Date Completed from source sheet if present
    const dateColIndex = getHeaderIndex(sheet, DATE_COMPLETED_HEADER);
    if (dateColIndex > 0) {
      const current = sheet.getRange(row, dateColIndex).getValue();
      if (current) sheet.getRange(row, dateColIndex).setValue('');
    }

    // Remove from Completed Jobs sheet (all rows with same folderId)
    const comp = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(COMPLETED_SHEET_NAME);
    if (comp) {
      // determine folderId header name from source sheet headers
      const sourceHeader = sheet.getRange(HEADER_ROW, FOLDER_ID_COLUMN).getDisplayValue() || 'Folder ID';
      let matches = findRowsInSheetByFolderId(comp, folderId, sourceHeader);
      // delete matches from bottom to top
      if (matches.length > 0) {
        matches.sort(function(a,b){return b - a;}); // descending
        matches.forEach(r => comp.deleteRow(r));
      }
    }

    // rename drive folder to remove '[COMPLETE]' or '[COMPLETED]' prefix
    try {
      const folder = DriveApp.getFolderById(folderId);
      const currentName = folder.getName();
      const newName = currentName.replace(/^\[(complete|completed)\]\s*/i, '');
      if (newName !== currentName) {
        folder.setName(newName);
      }
    } catch (err) {
      Logger.log('Failed to rename folder during revert for row ' + row + ': ' + err);
    }

    // Update Job Notes to reflect removed date
    try {
      createOrUpdateFolderForRow(sheet, row);
    } catch (err) {
      Logger.log('Failed to update notes after revert: ' + err);
    }

  } catch (err) {
    Logger.log('handleRevertFromComplete error: ' + err);
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* Bulk process */
function processAllRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
    const bid = sheet.getRange(r, COLUMNS.BID).getValue();
    const client = sheet.getRange(r, COLUMNS.CLIENT).getValue();
    if (bid && client) createOrUpdateFolderForRow(sheet, r);
  }
}
