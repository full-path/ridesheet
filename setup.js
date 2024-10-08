function importDataFromSheet(fileId = null, showWarning = true) {
  const ui = SpreadsheetApp.getUi();

  // Show the warning modal if showWarning is true.
  if (showWarning) {
    const response = ui.alert(
      "Warning!",
      "This operation will delete and replace the data in this sheet. Continue?",
      ui.ButtonSet.YES_NO
    );

    // If the user clicks 'No', stop the function.
    if (response != ui.Button.YES) {
      ui.alert("Operation cancelled.");
      return;
    }
  }

  // If no fileId is provided, prompt the user to enter one.
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

  const file = DriveApp.getFileById(fileId);
  const importSpreadsheet = SpreadsheetApp.open(file);

  // For testing
  const importSpreadsheetName = importSpreadsheet.getName();
  log("Importing data from sheet", importSpreadsheetName);

  // Check that the spreadsheet has the required sheets.
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

  // Check that Trip Review and Run Review sheets are empty apart from the header row.
  const tripReviewSheet = importSpreadsheet.getSheetByName("Trip Review");
  const runReviewSheet = importSpreadsheet.getSheetByName("Run Review");

  if (tripReviewSheet.getLastRow() > 1 || runReviewSheet.getLastRow() > 1) {
    ui.alert(
      "Can't import data. Please review and archive all data in Trip Review and Run Review before proceeding."
    );
    return;
  }

  // Attempt to import data from each required sheet.
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
    ui.alert("Data import completed successfully.");
  } catch (error) {
    ui.alert("Data import failed: " + error.message);
  }
}

function importSheet(sourceSpreadsheet, sheetName) {
    const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    if (!sourceSheet) {
      log('Skipping import for', sheetName, '- Source sheet not found.');
      return;
    }
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = targetSpreadsheet.getSheetByName(sheetName);
    if (!targetSheet) {
      log('Skipping import for', sheetName, '- Target sheet not found.');
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
  
    // Process each row from the source sheet, starting from the second row (skipping headers).
    for (let i = 1; i < sourceData.length; i++) {
      const sourceRow = sourceData[i];
      const targetRow = new Array(targetHeaders.length).fill(''); // Initialize the target row with empty values.
  
      targetHeaders.forEach((targetHeader, targetIndex) => {
        // Skip columns that start with the "|" character.
        if (targetHeader.startsWith('|')) {
          return;
        }
  
        // Check if the target header exists in the source headers.
        if (sourceHeaderMap.hasOwnProperty(targetHeader)) {
          const sourceIndex = sourceHeaderMap[targetHeader];
          targetRow[targetIndex] = sourceRow[sourceIndex];
        } else {
          // Log missing columns in the source sheet.
          log('Missing column in source for import:', sheetName, targetHeader);
        }
      });
  
      rowsToImport.push(targetRow);
    }
  
    // Clear values in the target sheet before importing data, keeping the header intact.
    const headerRange = targetSheet.getRange(1, 1, 1, targetSheet.getMaxColumns());
    const headerValues = headerRange.getValues();
    const numRows = targetSheet.getMaxRows();
    const numCols = targetSheet.getMaxColumns();
  
    // Clear all but the first row (header).
    targetSheet.getRange(2, 1, numRows - 1, numCols).clearContent();
  
    // Copy the mapped data to the target sheet.
    if (rowsToImport.length > 0) {
      const dataRange = targetSheet.getRange(2, 1, rowsToImport.length, rowsToImport[0].length);
      dataRange.setValues(rowsToImport);
    }
  
    log('Imported data into sheet', sheetName);
  }  
