// v1.0.4
// Source control: https://github.com/brandonscript/totk
// Standalone: copy this entire file into Code.gs. Globals.gs is not required.
const TRACKER_SHEET_NAME = 'Armor Tracker';
const ARMOR_SHEET_NAME = '_Armor';
const MATERIALS_SHEET_NAME = '_Materials';
const COSTS_SHEET_NAME = '_Costs';
const MATERIAL_GROUPS_SHEET_NAME = '_MaterialGroups';
const SUMMARY_COL = 5;
const PARENT_COL = 5;
const DEPENDENT_COL = 6;

function cleanMaterialKey(title) {
  return String(title || '').replace(/^.*?\-\s/i, '');
}

// Serialize complete refreshes, rather than individual nested functions.
// Locks release even on exceptions; no persistent "running" flags can get stuck.
function withTrackerLock(callback) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    try {
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
  }
}

function getArmorRows(book) {
  const sheet = book.getSheetByName(ARMOR_SHEET_NAME);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues()
    .filter(row => row[0] !== '');
}

function normalizeArmorRows(book, armor, editedRange) {
  const tracker = book.getSheetByName(TRACKER_SHEET_NAME);
  const first = editedRange ? Math.max(2, editedRange.getRow()) : 2;
  const last = editedRange ? Math.min(armor.length + 1, editedRange.getLastRow()) : armor.length + 1;
  const firstCol = editedRange ? Math.max(6, editedRange.getColumn()) : 6;
  const lastCol = editedRange ? Math.min(10, editedRange.getLastColumn()) : 10;
  if (last < first || lastCol < firstCol) return;
  const range = tracker.getRange(first, firstCol, last - first + 1, lastCol - firstCol + 1);
  const values = range.getValues();
  let changed = false;
  const normalized = values.map((row, index) => row.map((value, column) => {
    const col = firstCol + column;
    // _Armor columns H:K contain the four upgrade defense values.
    const next = col > 6 && !armor[first - 2 + index][col] ? 'N/A'
      : String(value).toLowerCase() === 'n/a' ? 'N/A'
      : String(value).trim() ? '✓' : value;
    if (next !== value) changed = true;
    return next;
  }));
  if (changed) range.setValues(normalized);
}

function getUpgradeLevel(upgradesRow) {
  return (upgradesRow || []).reduce((count, value) => count + (value === '✓' ? 1 : 0), 0);
}

// Read the cost matrix once, then sum in memory. Match materials by key rather
// than assuming the _Costs columns and _Materials rows are in the same order.
function writeMaterialSums(book, armor) {
  const costs = book.getSheetByName(COSTS_SHEET_NAME);
  const materials = book.getSheetByName(MATERIALS_SHEET_NAME);
  const numCols = costs.getLastColumn() - SUMMARY_COL;
  const headers = costs.getRange(1, 6, 1, numCols).getValues()[0].map(cleanMaterialKey);
  // Each armor piece owns four rows. Ignore thousands of empty formula spill rows.
  const rows = costs.getRange(2, 1, armor.length * 4, numCols + SUMMARY_COL).getValues();
  const upgrades = book.getSheetByName(TRACKER_SHEET_NAME)
    .getRange(2, 1, armor.length, 10).getValues();
  const levels = new Map(upgrades.map(row => [String(row[0]), getUpgradeLevel(row.slice(6, 10))]));
  const totals = Array(numCols).fill(0);
  const remaining = Array(numCols).fill(0);
  for (let row = 0; row < rows.length; row += 4) {
    const level = levels.get(String(rows[row][0])) || 0;
    for (let star = 0; star < 4; star++) {
      const values = rows[row + star];
      for (let col = 0; col < numCols; col++) {
        const quantity = Number(values[col + SUMMARY_COL]) || 0;
        totals[col] += quantity;
        if (star >= level) remaining[col] += quantity;
      }
    }
  }
  const byKey = new Map(headers.map((key, i) => [key, [totals[i], remaining[i]]]));
  const keys = materials.getRange(2, 2, materials.getLastRow() - 1, 1).getValues();
  // Trim trailing empty rows without disturbing unrelated cells or inventory.
  let length = keys.length;
  while (length && !keys[length - 1][0]) length--;
  if (length) materials.getRange(2, 8, length, 2).setValues(
    keys.slice(0, length).map(([key]) => byKey.get(cleanMaterialKey(key)) || [0, 0]));
}

function refreshArmorTracker() {
  return withTrackerLock(() => {
    const book = SpreadsheetApp.getActive();
    const armor = getArmorRows(book);
    normalizeArmorRows(book, armor);
    writeMaterialSums(book, armor);
  });
}

function calculateMaterialSums() {
  return withTrackerLock(() => {
    const book = SpreadsheetApp.getActive();
    writeMaterialSums(book, getArmorRows(book));
  });
}

function refreshMaterials() {
  return calculateMaterialSums();
}

function convertRangeToDictionary(range) {
  const dictionary = {};
  let key = '';
  for (const row of range.getValues()) {
    if (row[0] !== '') key = row[0];
    if (row[1] !== '') (dictionary[key] || (dictionary[key] = [])).push(row[1]);
  }
  return dictionary;
}

function updateDependentDropdowns(editedRange) {
  const sheet = editedRange ? editedRange.getSheet() : SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== MATERIALS_SHEET_NAME) return;
  const range = editedRange || sheet.getActiveRange();
  if (range.getColumn() > 6 || range.getLastColumn() < 5) return;
  const first = Math.max(2, range.getRow());
  const last = range.getLastRow();
  if (last < first) return;
  const groups = convertRangeToDictionary(SpreadsheetApp.getActive()
    .getSheetByName(MATERIAL_GROUPS_SHEET_NAME).getRange('A2:B500'));
  const values = sheet.getRange(first, 5, last - first + 1, 2).getValues();
  const validations = values.map(([parent]) => {
    const options = groups[parent] || [];
    return [options.length ? SpreadsheetApp.newDataValidation().requireValueInList(options).build() : null];
  });
  // A changed parent invalidates its dependent value. Editing the dependent
  // column alone retains the entered value, as in the original tracker.
  if (range.getColumn() <= 5) {
    sheet.getRange(first, 6, values.length, 1).setValues(values.map(([parent, value]) =>
      [(groups[parent] || []).includes(value) ? value : '']));
  }
  sheet.getRange(first, 6, values.length, 1).setDataValidations(validations);
}

function updateAllDynamicDropdowns() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(MATERIALS_SHEET_NAME);
  updateDependentDropdowns(sheet.getRange(2, 5, sheet.getLastRow() - 1, 2));
}

function writeCostSummaries(book, requestedRange) {
  const costs = book.getSheetByName(COSTS_SHEET_NAME);
  const armor = getArmorRows(book);
  const first = requestedRange ? Math.max(2, requestedRange.getRow()) : 2;
  const last = requestedRange ? Math.min(armor.length * 4 + 1, requestedRange.getLastRow()) : armor.length * 4 + 1;
  if (last < first) return;
  const materialSheet = book.getSheetByName(MATERIALS_SHEET_NAME);
  const names = new Map(materialSheet.getRange(2, 2, materialSheet.getLastRow() - 1, 3)
    .getValues().filter(row => row[0]).map(row => [cleanMaterialKey(row[0]), row[2]]));
  const width = costs.getLastColumn() - SUMMARY_COL;
  const keys = costs.getRange(1, 6, 1, width).getValues()[0].map(cleanMaterialKey);
  const values = costs.getRange(first, 6, last - first + 1, width).getValues();
  const summaries = values.map(row => [row.flatMap((quantity, index) => {
    const key = keys[index];
    if (!quantity || !key || (key === 'rupee' && [10, 50, 200, 500].includes(quantity))) return [];
    return [`${quantity} ${names.get(key) || key}`];
  }).join('\n')]);
  costs.getRange(first, SUMMARY_COL, summaries.length, 1).setValues(summaries);
}

function summarizeCostsRows(range) {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (!range && sheet.getName() !== COSTS_SHEET_NAME) return;
  return withTrackerLock(() => writeCostSummaries(SpreadsheetApp.getActive(), range || sheet.getActiveRange()));
}

function updateAllSummaries() {
  return withTrackerLock(() => writeCostSummaries(SpreadsheetApp.getActive()));
}

function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const name = range.getSheet().getName();
  if (name === TRACKER_SHEET_NAME && range.getLastRow() >= 2
      && range.getColumn() <= 10 && range.getLastColumn() >= 6) {
    withTrackerLock(() => {
      const book = e.source || SpreadsheetApp.getActive();
      const armor = getArmorRows(book);
      normalizeArmorRows(book, armor, range);
      writeMaterialSums(book, armor);
    });
  } else if (name === COSTS_SHEET_NAME && range.getLastColumn() > SUMMARY_COL) {
    withTrackerLock(() => {
      const book = e.source || SpreadsheetApp.getActive();
      writeCostSummaries(book, range);
      writeMaterialSums(book, getArmorRows(book));
    });
  } else if (name === MATERIALS_SHEET_NAME) {
    updateDependentDropdowns(range);
  }
}

function onSelectionChange() {
  // Validation updates on edit; moving the selection requires no spreadsheet calls.
}

function onOpen() {
  // Static cost summaries change only when _Costs changes. Avoid rebuilding them
  // on every open; the refresh button also repairs missed edit-trigger updates.
  refreshArmorTracker();
}
