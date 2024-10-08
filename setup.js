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

// Stub function to import a single sheet.
function importSheet(importSpreadsheet, sheetName) {
  // Get the sheet from the import spreadsheet.
  const importSheet = importSpreadsheet.getSheetByName(sheetName);
  if (!importSheet) {
    throw new Error(
      'Sheet "' + sheetName + '" not found in the source spreadsheet.'
    );
  }

  // Get the corresponding sheet in the current spreadsheet.
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = targetSpreadsheet.getSheetByName(sheetName);
  if (!targetSheet) {
    throw new Error(
      'Sheet "' + sheetName + '" not found in the target spreadsheet.'
    );
  }

  // Clear the target sheet before importing data.

  // Copy data from the import sheet to the target sheet.

  // Log the completion of importing this sheet.
  log("Successfully imported data into sheet", sheetName);
}
