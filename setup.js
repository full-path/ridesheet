/**
 * @fileoverview Data import utilities for migrating data into a RideSheet instance.
 *
 * Provides tools to import trip, run, customer, and reference data from another
 * RideSheet spreadsheet. The import is column-header-driven: data is matched by
 * header name, so schema differences between source and target are handled
 * gracefully (extra source columns are logged and skipped; pipe-prefixed `|
 * formula columns are always skipped). After import, sheet formats and data
 * validation rules are reapplied via `applySheetFormatsAndValidation()`.
 */

/**
 * Imports data from another Google Sheet into the current RideSheet instance,
 * replacing the data in all standard sheets. Prompts the user for a source
 * file ID if one is not supplied, and optionally imports Document Properties
 * as well. Validates that the source spreadsheet looks like a RideSheet
 * instance and that Trip Review / Run Review are empty before proceeding.
 *
 * @param {string} [fileId=null] - The ID of the Google Sheet to import data
 *   from. If not provided, the user is prompted to enter it.
 * @param {boolean} [showWarning=true] - Whether to show a confirmation
 *   warning before proceeding with the import.
 */
function importDataFromSheet(fileId = null, showWarning = true) {
  const ui = SpreadsheetApp.getUi();

  if (showWarning) {
    const response = ui.alert(
      "Warning!",
      "This operation will delete and replace the data in this sheet. Continue?",
      ui.ButtonSet.YES_NO
    );

    if (response != ui.Button.YES) {
      ui.alert("Operation cancelled.");
      return;
    }
  }

  if (!fileId) {
    fileId = ui
      .prompt(
        "Enter the ID of the Google Sheet you want to import data from:",
        ui.ButtonSet.OK_CANCEL
      )
      .getResponseText();

    if (!fileId) {
      ui.alert("No file selected. Operation cancelled.");
      return;
    }
  }

  const importDocProps = ui.alert(
    "Do you want to import and overwrite Document Properties?",
    ui.ButtonSet.YES_NO
  );

  const file = DriveApp.getFileById(fileId);
  const importSpreadsheet = SpreadsheetApp.open(file);

  const importSpreadsheetName = importSpreadsheet.getName();
  log("Importing data from sheet", importSpreadsheetName);

  const requiredSheets = [
    "Customers",
    "Trips",
    "Runs",
    "Trip Review",
    "Run Review",
    "Trip Archive",
    "Run Archive",
    "Services",
    "Drivers",
    "Vehicles",
  ];
  const sheets = importSpreadsheet.getSheets().map((sheet) => sheet.getName());
  const missingSheets = requiredSheets.filter(
    (sheetName) => !sheets.includes(sheetName)
  );

  if (missingSheets.length > 0) {
    ui.alert(
      "Can't import data. Does not appear to be a valid instance of RideSheet. Missing sheets: " +
        missingSheets.join(", ")
    );
    return;
  }

  const tripReviewSheet = importSpreadsheet.getSheetByName("Trip Review");
  const runReviewSheet = importSpreadsheet.getSheetByName("Run Review");

  if (tripReviewSheet.getLastRow() > 1 || runReviewSheet.getLastRow() > 1) {
    ui.alert(
      "Can't import data. Please review and archive all data in Trip Review and Run Review before proceeding."
    );
    return;
  }

  try {
    const sheetsToImport = [
      "Customers",
      "Trips",
      "Runs",
      "Trip Archive",
      "Run Archive",
      "Services",
      "Drivers",
      "Vehicles",
    ];
    for (const sheetName of sheetsToImport) {
      importSheet(importSpreadsheet, sheetName);
    }
    
    if (importDocProps === ui.Button.YES) {
      importDocumentProperties(importSpreadsheet, SpreadsheetApp.getActiveSpreadsheet());
    }
    
    ui.alert("Data import completed successfully.");

  } catch (error) {
    ui.alert("Data import failed: " + error.message);
    logError(error);
  }
}

/**
 * Imports data from a specific sheet in the source spreadsheet into the
 * corresponding sheet in the active spreadsheet. Matches rows to columns
 * by header name, skipping pipe-prefixed (`|`) formula columns. Clears the
 * target sheet's data range before writing, then reapplies formats and
 * validation via `applySheetFormatsAndValidation()`. Logs any source columns
 * that have no match in the target.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet - The
 *   spreadsheet to import data from.
 * @param {string} sheetName - The name of the sheet to import.
 */
function importSheet(sourceSpreadsheet, sheetName) {
  const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
  if (!sourceSheet) {
    log(`Skipping import for ${sheetName}`, 'Source sheet not found.');
    return;
  }
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = targetSpreadsheet.getSheetByName(sheetName);
  if (!targetSheet) {
    log(`Skipping import for ${sheetName}`, 'Target sheet not found.');
    return;
  }
  const sourceData = sourceSheet.getDataRange().getValues();
  const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getMaxColumns()).getValues()[0];
  const sourceHeaders = sourceData[0];

  const sourceHeaderMap = {};
  sourceHeaders.forEach((header, index) => {
    sourceHeaderMap[header] = index;
  });

  const targetHeaderMap = {};
  targetHeaders.forEach((header, index) => {
    targetHeaderMap[header] = index;
  });
  const rowsToImport = [];

  const missingInTarget = sourceHeaders.filter(header => !targetHeaders.includes(header));
  if (missingInTarget.length > 0) {
    log(`Columns in source but not in target for ${sheetName}`, missingInTarget.join(', '));
  }
  
  for (let i = 1; i < sourceData.length; i++) {
    const sourceRow = sourceData[i];
    const targetRow = new Array(targetHeaders.length).fill(null);

    sourceHeaders.forEach((sourceHeader, sourceIndex) => {
      if (!sourceHeader.startsWith('|') && targetHeaderMap.hasOwnProperty(sourceHeader)) {
        const targetIndex = targetHeaderMap[sourceHeader];
        targetRow[targetIndex] = sourceRow[sourceIndex];
      }
    });

    rowsToImport.push(targetRow);
  }

  const dataRange = targetSheet.getRange(2, 1, targetSheet.getMaxRows() - 1, targetSheet.getMaxColumns());
  dataRange.clearContent().clearDataValidations();

  if (rowsToImport.length > 0) {
    const dataRange = targetSheet.getRange(2, 1, rowsToImport.length, rowsToImport[0].length);
    dataRange.setValues(rowsToImport);
  }

  applySheetFormatsAndValidation(targetSheet);

  log(`Imported ${rowsToImport.length} rows into sheet ${sheetName}`);
}

/**
 * Imports document property values from the Document Properties sheet of
 * `sourceSpreadsheet` into the Document Properties sheet of `targetSpreadsheet`.
 * Only properties whose keys already exist in the target are updated; unknown
 * source keys are logged and skipped. Calls `buildDocumentPropertiesFromSheet()`
 * after writing to reload the in-memory properties cache.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet - The
 *   spreadsheet to import properties from.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} targetSpreadsheet - The
 *   spreadsheet to import properties into.
 */
function importDocumentProperties(sourceSpreadsheet, targetSpreadsheet) {
  const sourceSheet = sourceSpreadsheet.getSheetByName("Document Properties");
  const targetSheet = targetSpreadsheet.getSheetByName("Document Properties");
  const sourceData = sourceSheet.getDataRange().getValues();
  const targetData = targetSheet.getDataRange().getValues();
  
  const targetProps = new Map(targetData.slice(1).map(row => [row[0], row[1]]));
  
  sourceData.slice(1).forEach(row => {
    const [key, value] = row;
    if (key && value !== undefined) {
      if (targetProps.has(key)) {
        targetProps.set(key, value);
      } else {
        log("Key found in source but not in target Document Properties", key);
      }
    }
  });
  
  const updatedData = [targetData[0].slice(0, 2), ...Array.from(targetProps)];
  
  targetSheet.getRange(1, 1, targetSheet.getLastRow(), 2).clearContent();
  targetSheet.getRange(1, 1, updatedData.length, 2).setValues(updatedData);
  
  // Rebuild document properties from the updated sheet
  buildDocumentPropertiesFromSheet();
}
