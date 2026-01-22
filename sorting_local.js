function addSortingMenuItems() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('Trip Sorting')
  const config = getSortConfig()
  for (let i = 0; i < Math.min(config.length, 9); i++) {
    menu.addItem(config[i].name, `sortingMenuItem${i}`)
  }
  menu.addToUi()
}

function sortTrips(index) {
  const config = getSortConfig()

  const sortingByColumnName = config[index].sort

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  const dataRange = sheet.getDataRange()
  const numRows = dataRange.getNumRows()
  const numCols = dataRange.getNumColumns()
  if (numRows <= 1) return
  const rangeToSort = sheet.getRange(2, 1, numRows - 1, numCols)
  const sheetHeaderNames = getSheetHeaderNames(sheet)

  let missingHeaders = []
  const sortingByColumnNumber = sortingByColumnName.map(item => {
    const columnNumber = sheetHeaderNames.indexOf(item.column) + 1
    if (!columnNumber) missingHeaders.push(item.column)
    return {column: columnNumber, ascending: item.ascending}
  })
  if (missingHeaders.length) {
    ss.toast(`Sorting failed. Columns missing: ${missingHeaders.join(", ")}`)
    return
  }

  rangeToSort.sort(sortingByColumnNumber)
}

function sortingMenuItem0() { sortTrips(0) }
function sortingMenuItem1() { sortTrips(1) }
function sortingMenuItem2() { sortTrips(2) }
function sortingMenuItem3() { sortTrips(3) }
function sortingMenuItem4() { sortTrips(4) }
function sortingMenuItem5() { sortTrips(5) }
function sortingMenuItem6() { sortTrips(6) }
function sortingMenuItem7() { sortTrips(7) }
function sortingMenuItem8() { sortTrips(8) }
function sortingMenuItem9() { sortTrips(9) }

function getSortConfig() {
  return [
    {
      name: "Date, PU Time",
      sort: [
        {
          column: "Trip Date",
          ascending: false
        },
        {
          column: "PU Time",
          ascending: true
        },
      ]
    },
    {
      name: "Date, Vehicle, PU Time",
      sort: [
        {
          column: "Trip Date",
          ascending: false
        },
        {
          column: "Vehicle ID",
          ascending: true
        },
        {
          column: "PU Time",
          ascending: true
        },
      ]
    },
    {
      name: "Customer, Date, PU Time",
      sort: [
        {
          column: "Customer Name and ID",
          ascending: true
        },
        {
          column: "Trip Date",
          ascending: false
        },
        {
          column: "PU Time",
          ascending: true
        },
      ]
    }
  ]
}

/**
 * Creates or updates a filter view on the active sheet.
 * This function is designed to be a companion to sortTrips() and uses the
 * same config object from getSortConfig().
 *
 * @param {number} index - The index of the config object from getSortConfig().
 */
function createOrUpdateFilterView(index) {
  // --- 1. Get Config and Sheet Details ---
  if (!index) index = 2
  const allConfigs = getSortConfig();
  if (!allConfigs[index]) {
    Logger.log(`Error: No config found at index ${index}.`);
    return;
  }
  
  const config = allConfigs[index];
  const filterViewTitle = config.name; // Assumes config has a "name" property
  const sortSettings = config.sort;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetId = ss.getId();
  const sheet = ss.getSheetByName("Trip Review")
  const sheetId = sheet.getSheetId();
  
  // --- 2. Define the Filter View's Range ---
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow <= 1) return; // Don't create a filter for just a header

  // Define the GridRange for the filter view
  const filterRange = {
    sheetId: sheetId,
    startRowIndex: 0, // 0-indexed, so 0 is row 1 (header)
    endRowIndex: lastRow,
    startColumnIndex: 0, // 0 is column A
    endColumnIndex: lastCol
  };
  
  // --- 3. Convert Sort Config to API "SortSpecs" ---
  
  const sheetHeaderNames = getSheetHeaderNames(sheet);
  let missingHeaders = [];
  const sortSpecs = [];

  sortSettings.forEach(item => {
    const columnNumber = sheetHeaderNames.indexOf(item.column) + 1;
    if (!columnNumber) {
      missingHeaders.push(item.column);
    } else {
      sortSpecs.push({
        dimensionIndex: columnNumber - 1, // API uses 0-indexed column
        sortOrder: item.ascending ? "ASCENDING" : "DESCENDING"
      });
    }
  });

  if (missingHeaders.length) {
    ss.toast(`Filter view failed. Columns missing: ${missingHeaders.join(", ")}`);
    return;
  }

  // --- 4. Check for an Existing Filter View ---
  
  let existingFilterViewId = null;
  try {
    // Get all filter views on the spreadsheet
    const allFilterViews = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: 'sheets.filterViews(filterViewId,title,range.sheetId)'
    }).sheets
      .flatMap(s => s.filterViews || []); // Get all filter views from all sheets

    // Find a filter view on *this* sheet with a matching title
    const existingView = allFilterViews.find(fv => 
      fv.range.sheetId === sheetId && fv.title === filterViewTitle
    );

    if (existingView) {
      existingFilterViewId = existingView.filterViewId;
    }

  } catch (e) {
    ss.toast(`Error checking for filter views: ${e.message}`);
    return;
  }

  // --- 5. Create or Update the Filter View ---
  
  let request;
  if (existingFilterViewId) {
    // --- UPDATE existing filter view ---
    request = {
      updateFilterView: {
        filter: {
          filterViewId: existingFilterViewId,
          title: filterViewTitle,
          range: filterRange,
          sortSpecs: sortSpecs
        },
        fields: "title,range,sortSpecs" // Fields to update
      }
    };
    ss.toast(`Filter view "${filterViewTitle}" updated.`);
  } else {
    // --- ADD new filter view ---
    request = {
      addFilterView: {
        filter: {
          // No filterViewId needed for creation
          title: filterViewTitle,
          range: filterRange,
          sortSpecs: sortSpecs
        }
      }
    };
    ss.toast(`Filter view "${filterViewTitle}" created.`);
  }

  // --- 6. Execute the Request ---
  
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: [request] }, spreadsheetId);
  } catch (e) {
    ss.toast(`Error saving filter view: ${e.message}`);
  }
}

/**
 * Updates all filter views on the active sheet to match the sheet's
 * current data range (A1:LastColumnLastRow).
 * * This requires the Advanced Sheets Service to be enabled.
 */
function updateFilterViewRanges() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetId = sheet.getSheetId();
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // Get the full range of data on the sheet (A1 notation)
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  // Do nothing if the sheet is empty
  if (lastRow === 0 || lastCol === 0) {
    SpreadsheetApp.getUi().alert('Sheet is empty. No filters to update.');
    return;
  }
  
  // Define the new range that all filters should cover.
  // This starts from row 1, column 1.
  const newRange = {
    sheetId: sheetId,
    startRowIndex: 0, // 0-indexed (row 1 is 0)
    endRowIndex: lastRow,
    startColumnIndex: 0, // 0-indexed (col A is 0)
    endColumnIndex: lastCol
  };

  try {
    // 1. Get the spreadsheet's metadata to find all filter views
    const spreadsheet = Sheets.Spreadsheets.get(spreadsheetId, {
      fields: 'sheets(filterViews(filterViewId,range))'
    });

    // 2. Find all filter views that are on the *current* sheet
    const sheetInfo = spreadsheet.sheets.find(s => s.filterViews && s.filterViews.some(fv => fv.range.sheetId === sheetId));
    
    if (!sheetInfo || !sheetInfo.filterViews) {
      SpreadsheetApp.getUi().alert('No filter views found on this sheet.');
      return;
    }

    const requests = [];
    const filterViewsOnThisSheet = sheetInfo.filterViews.filter(fv => fv.range.sheetId === sheetId);

    // 3. Create an "update" request for each filter view on this sheet
    filterViewsOnThisSheet.forEach(fv => {
      requests.push({
        updateFilterView: {
          filter: {
            filterViewId: fv.filterViewId,
            range: newRange 
          },
          fields: 'range' // This tells the API to *only* update the range
        }
      });
    });

    // 4. Send all updates to the API in a single "batch" request
    if (requests.length > 0) {
      Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
      SpreadsheetApp.getUi().alert(`Successfully updated ${requests.length} filter view(s).`);
    } else {
      SpreadsheetApp.getUi().alert('No filter views found on this sheet.');
    }

  } catch (e) {
    Logger.log(e);
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}