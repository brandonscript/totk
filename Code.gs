// v1.0.3 (alpha)
// Source control maintained at https://github.com/brandonscript/totk

// Material groups
const PARENT_COL = 5;
const DEPENDENT_COL = 6;
const MATERIAL_GROUPS_SHEET_NAME = "_MaterialGroups"

// Cost summaries
const MATERIALS_SHEET_NAME = "_Materials";
const COSTS_SHEET_NAME = "_Costs"
const MATERIALS_DICT = {}
const SUMMARY_COL = 5
let NUM_COSTS_COLS = 0;

// Tracker
const TRACKER_SHEET_NAME = "Armor Tracker"
const ARMOR_SHEET_NAME = "_Armor"

function cleanMaterialKey(title) {
  return title?.replace(/^.*?\-\s/i, '')
}

function buildMaterialsDict() {
  var activeSheet = SpreadsheetApp.getActiveSheet().getName();
  if (activeSheet !== MATERIALS_SHEET_NAME && Object.keys(MATERIALS_DICT).length > 0) {
    return;
  }

  var materialsSheet = SpreadsheetApp.getActive().getSheetByName(MATERIALS_SHEET_NAME);
  var keysRange = materialsSheet.getRange("B2:B" + materialsSheet.getLastRow());
  var valuesRange = materialsSheet.getRange("D2:D" + materialsSheet.getLastRow());
  
  var keys = keysRange.getValues().flat();  // Get the keys as a flat array
  var values = valuesRange.getValues().flat();  // Get the values as a flat array
  
  for (var i = 0; i < keys.length; i++) {
    let key = cleanMaterialKey(keys[i])
    let value = values[i]
    MATERIALS_DICT[key] = value;
  }
}

function getCostsCols() {

  var activeSheet = SpreadsheetApp.getActiveSheet().getName();
  if (activeSheet !== COSTS_SHEET_NAME && NUM_COSTS_COLS > 0) {
    return;
  }

  var costsSheet = SpreadsheetApp.getActive().getSheetByName(COSTS_SHEET_NAME);

  if (costsSheet.getActiveCell().getRow() > 1 && NUM_COSTS_COLS > 0) {
    return
  }

  NUM_COSTS_COLS = costsSheet.getLastColumn() - SUMMARY_COL; 
}

function areArraysEqual(arr1, arr2) {
  if (!Array.isArray(arr1) || !Array.isArray(arr2) || arr1.length !== arr2.length) {
    return false;
  }

  for (var i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) {
      return false;
    }
  }

  return true;
}

function clearValidationAndValue(range) {
  range.clearDataValidations();
  range.clearContent();
}

function applyDataValidation(range, allowedValues) {
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(allowedValues).build();
  range.setDataValidation(rule);
}

function getParentValue(sheet, row, column) {
  return sheet.getRange(row, column).getValue();
}

function getCurrentValidationOptions(row) {
  var sheet = SpreadsheetApp.getActiveSheet();

  var activeCell = sheet.getRange(row, DEPENDENT_COL);
  var dataValidation = activeCell.getDataValidation();

  if (dataValidation !== null) {
    var validationCriteria = dataValidation.getCriteriaType();

    if (validationCriteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      var validationValues = dataValidation.getCriteriaValues();
      var options = validationValues[0];

      return options;
    }
  }

  return null; // Return null if the cell in column E doesn't have value validation
}

function convertRangeToDictionary(range) {
  var dictionary = {};
  var numRows = range.getNumRows();
  var values = range.getValues();
  var currentKey = '';
  
  for (var i = 0; i < numRows; i++) {
    var key = values[i][0];
    var value = values[i][1];
    
    if (key !== '') {
      currentKey = key;
    }
    if (value !== '') {
      if (!dictionary.hasOwnProperty(currentKey)) {
        dictionary[currentKey] = [];
      }
      dictionary[currentKey].push(value);
    }
  }
  
  return dictionary;
}

function updateDependentDropdowns() {
  var sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== MATERIALS_SHEET_NAME) {
    return;
  }
  
  var selectedCell = sheet.getActiveCell();
  var row = selectedCell.getRow();
  var column = selectedCell.getColumn();

  var parentValue = sheet.getRange(row, PARENT_COL).getValue();
  var currentValue = sheet.getRange(row, DEPENDENT_COL).getValue();

  var isParentColumn = column === PARENT_COL;
  var isDependentColumn = column === DEPENDENT_COL;

  var dataSheet = SpreadsheetApp.getActive().getSheetByName(MATERIAL_GROUPS_SHEET_NAME);
  var dataRange = dataSheet.getRange("A2:B500");
  var dictionary = convertRangeToDictionary(dataRange);
  var expectedOptions = dictionary?.[parentValue] ?? [];
  var dependentRange = sheet.getRange(row, DEPENDENT_COL);
  var currentOptions = getCurrentValidationOptions(row)

  // if (isParentColumn && !areArraysEqual(expectedOptions, currentOptions)) {
  if (isParentColumn && !expectedOptions.includes(currentValue)) {
    var parentValue = getParentValue(sheet, row, PARENT_COL);
    clearValidationAndValue(dependentRange);
    applyDataValidation(dependentRange, expectedOptions);
  } else if (isDependentColumn) {
    applyDataValidation(dependentRange, expectedOptions);
  }
}


function summarizeCostsRows(range) {

  Singleton(() => {

    buildMaterialsDict();
    getCostsCols();
    var colOffset = 5

    var sheet = SpreadsheetApp.getActiveSheet();
    var sheetName = sheet.getName();

    if (sheetName !== COSTS_SHEET_NAME) {
      return
    }

    var currentRange = range ?? sheet.getActiveRange();
    var startRow = currentRange.getRow();
    var numRows = currentRange.getNumRows();
    var endRow = startRow + numRows

    let summaries = []

    const batchSize = Math.min(4, endRow - startRow)

    for (var row = startRow; row < endRow; row++) {
      let summary = [];

      var partRange = sheet.getRange(row, SUMMARY_COL + 1, 1, NUM_COSTS_COLS)
      var parts = partRange.getValues();

      for (var col = 1; col <= NUM_COSTS_COLS; col++) {

        var value = parts[0][col - 1];
        if (!value) {
          continue 
        }
        var colTitle = cleanMaterialKey(sheet.getRange(1, col + colOffset).getValue());

        if (colTitle === "rupee" && [10, 50, 200, 500].includes(value)) {
          // ignore default pricing
          continue;
        }
        var colDisplayTitle = MATERIALS_DICT?.[colTitle]

        // var qtyStr = value < 10 ? value.toString().padStart(3) : value.toString();
        summary.push(`${value} ${colDisplayTitle}`)
      }

      if (row > 1) {
        summaries.push([summary.join('\n')])
      }

      if (summaries.length === batchSize) {
        console.log(`Updating summary for ${batchSize} rows (${row-batchSize+1}-${row})`)
        sheet.getRange(row-batchSize+1, SUMMARY_COL, summaries.length, 1).setValues(summaries)
        summaries = []
      }
    }
    // sheet.getRange(row, SUMMARY_COL).setValue()
  }, 'summarizeCostRows');
}

function updateAllDynamicDropdowns() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var dataSheet = SpreadsheetApp.getActive().getSheetByName(MATERIAL_GROUPS_SHEET_NAME);
  var dataRange = dataSheet.getRange("A2:B500");
  var dictionary = convertRangeToDictionary(dataRange);

  var numRows = sheet.getLastRow();

  for (var row = 2; row <= numRows; row++) {
    var parentValue = getParentValue(sheet, row, PARENT_COL);
    var expectedOptions = dictionary?.[parentValue] ?? [];

    var dependentRange = sheet.getRange(row, DEPENDENT_COL);
    var currentValue = dependentRange.getValue(); // Get the current value before clearing validation

    clearValidationAndValue(dependentRange);
    applyDataValidation(dependentRange, expectedOptions);

    // Restore the current value if it exists in the new options
    if (expectedOptions.includes(currentValue)) {
      dependentRange.setValue(currentValue);
    }
  }
}

function updateAllSummaries() {
  var costsSheet = SpreadsheetApp.getActive().getSheetByName(COSTS_SHEET_NAME);
  var dataRange = costsSheet.getRange("A2:A553")
  summarizeCostsRows(dataRange)
}

function refreshArmorTracker() {

  Singleton(() => {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var trackerSheet = SpreadsheetApp.getActive().getSheetByName(TRACKER_SHEET_NAME);
    var armorSheet = SpreadsheetApp.getActive().getSheetByName(ARMOR_SHEET_NAME);

    if (activeSheet.getName() !== TRACKER_SHEET_NAME) {
      return;
    }

    var startRow = 2;
    var lastRow = trackerSheet.getLastRow();
    var range = trackerSheet.getRange(startRow, 6, lastRow - startRow + 1, 5);
    var armorRange = armorSheet.getRange(startRow, 7, lastRow - startRow + 1, 4); // Adjust column index to match ARMOR_SHEET_NAME
    var values = range.getValues();
    var armorValues = armorRange.getValues();

    var updatedValues = [];
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var armorRow = armorValues[i];
      var updatedRow = [];

      for (var j = 0; j < row.length; j++) {
        var currentValue = row[j];
        var armorValue = armorRow[j - 1]; // Adjust column index to match ARMOR_SHEET_NAME

        if (j > 0 && (!armorValue || currentValue.toString().toLowerCase() === "n/a")) {
          updatedRow.push("N/A");
        } else if (currentValue && currentValue.toString().trim() !== "") {
          updatedRow.push("✓");
        } else {
          updatedRow.push(currentValue);
        }
      }

      updatedValues.push(updatedRow);
    }

    range.setValues(updatedValues);
  }, 'refreshArmorTracker');
}

function calculateMaterialSums() {

  Singleton(() => {
    const costsSheet = SpreadsheetApp.getActive().getSheetByName(COSTS_SHEET_NAME);
    const materialsSheet = SpreadsheetApp.getActive().getSheetByName(MATERIALS_SHEET_NAME);
    const trackerSheet = SpreadsheetApp.getActive().getSheetByName(TRACKER_SHEET_NAME);
    const startRow = 2;
    const startCol = 5;
    const lastRow = costsSheet.getLastRow();

    const materialKeys = costsSheet.getRange(1, 1, 1, 256).getValues()[0].slice(startCol).map(cleanMaterialKey);
    const materialValues = costsSheet.getRange(2, 1, lastRow - 1, 256).getValues(); // Assuming A:IW range
    const materialsDict = {}; // Dictionary to store the totals and remaining values

    const upgradesRange = trackerSheet.getRange(2, 7, trackerSheet.getLastRow() - 1, 4); // Assuming columns G to J contain the upgrades in Tracker sheet
    const upgradesValues = upgradesRange.getValues();

    for (let row = 0; row < lastRow - startRow; row += 4) {
      const currentItem = materialValues[row];
      const currentItemName = currentItem[3];
      const itemId = currentItem[0];
      const upgradeLevel = itemId ? getUpgradeLevel(upgradesValues[itemId - 1]) : null;

      // sub-loop through this + the next three rows to get upgrade materials for this item
      for (let subRow = 0; subRow < 4; subRow++) {
        const itemRow = materialValues[row + subRow];
        
        // for each col from F to IW
        for (let col = 0; col < materialKeys.length; col++) {
          const key = materialKeys[col];

          // add or increment in materialsDict { [key]: total }
          const total = itemRow.slice(startCol)[col] || 0;
          
          // if current subrow is greater than upgradeLevel, add or increment { [key]: remaining } as well
          const remaining = subRow >= upgradeLevel ? total : 0;

          if (key && !materialsDict?.[key]) {
            materialsDict[key] = {
              total: total,
              remaining: remaining
            };
          } else if (key && materialsDict?.[key]) {
            materialsDict[key].total += total;
            materialsDict[key].remaining += remaining;
          }
        }
      }
    }
    
    // Write materialsDict totals to _Materials column H
    const totalsRange = materialsSheet.getRange(2, 8, materialKeys.length, 1);
    const totalsValues = materialKeys.map((key) => [materialsDict[key]?.total || 0]);
    totalsRange.setValues(totalsValues);

    // Write materialsDict remaining values to _Materials column I
    const remainingRange = materialsSheet.getRange(2, 9, materialKeys.length, 1);
    const remainingValues = materialKeys.map((key) => [materialsDict[key]?.remaining || 0]);
    remainingRange.setValues(remainingValues);
  }, "calculateMaterialSums");
}


function getUpgradeLevel(upgradesRow) {
  return upgradesRow.reduce(function(count, value) {
    return count + (value === "✓" ? 1 : 0);
  }, 0);
}


////////////////////////////

function refreshMaterials() {
  updateAllDynamicDropdowns();
  calculateMaterialSums();
}

function onEdit() {
  updateDependentDropdowns();
  refreshArmorTracker();
  summarizeCostsRows();
}

function onSelectionChange() {
  updateDependentDropdowns();
}

function onOpen() {
  buildMaterialsDict();
  updateAllSummaries();
  refreshArmorTracker();
  calculateMaterialSums();
}




