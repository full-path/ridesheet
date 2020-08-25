function moveTripsToReview() {
  const ss              = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet       = ss.getSheetByName("Trips")
  const tripReviewSheet = ss.getSheetByName("Trip Review")
  const tripFilter      = function(row) { return row["Trip Date"] && row["Trip Date"] < dateToday() }
  moveRows(tripSheet, tripReviewSheet, tripFilter)  
  
  const runSheet        = ss.getSheetByName("Runs")
  const runReviewSheet  = ss.getSheetByName("Run Review")
  const runFilter       = function(row) { return row["Run Date"] && row["Run Date"] < dateToday() }
  moveRows(runSheet, runReviewSheet, runFilter)  
}

function moveTripsToArchive() {
  const ss               = SpreadsheetApp.getActiveSpreadsheet()
  const tripReviewSheet  = ss.getSheetByName("Trip Review")
  const tripArchiveSheet = ss.getSheetByName("Trip Archive")
  const tripFilter       = function(row) { 
    const columns = getDocProp("tripReviewRequiredFields")
    blankColumns = columns.filter(column => !row[column])
    return blankColumns.length === 0
  }
  moveRows(tripReviewSheet, tripArchiveSheet, tripFilter)  

  const runReviewSheet  = ss.getSheetByName("Run Review")
  const runArchiveSheet = ss.getSheetByName("Run Archive")
  const runFilter       = function(row) { 
    const columns = getDocProp("runReviewRequiredFields")
    blankColumns = columns.filter(column => !row[column])
    return blankColumns.length === 0
  }
  moveRows(runReviewSheet, runArchiveSheet, runFilter)  
}