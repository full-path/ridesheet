function createReturnTrip() {
  const ss              = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet       = ss.getActiveSheet()
  const sourceTripRange = getFullRow(tripSheet.getActiveCell())
  const sourceTripRow   = sourceTripRange.getRow()
  const sourceTripData  = getRangeValuesAsTable(sourceTripRange)[0]
  let   returnTripData  = {...sourceTripData}
  if (tripSheet.getName() === "Trips" && isCompleteTrip(sourceTripData)) {
    returnTripData["PU Address"] = sourceTripData["DO Address"]
    returnTripData["DO Address"] = sourceTripData["PU Address"]
    if (sourceTripData["Appt Time"]) {
      returnTripData["PU Time"] = timeAdd(sourceTripData["Appt Time"], 60*60*1000)
    } else if (sourceTripData["DO Time"]) {
      returnTripData["PU Time"] = timeAdd(sourceTripData["DO Time"], 60*60*1000)
    } else {
      returnTripData["PU Time"] = null
    }
    returnTripData["DO Time"]     = null
    returnTripData["Appt Time"]   = null
    returnTripData["Est Hours"]   = null
    returnTripData["Est Miles"]   = null
    returnTripData["Trip ID"]     = Utilities.getUuid()
    returnTripData["Calendar ID"] = null
    tripSheet.insertRowAfter(sourceTripRow)
    let returnTripRange = getFullRow(tripSheet.getRange(sourceTripRow + 1, 1))
    setValuesByHeaderNames([returnTripData],returnTripRange)
    fillHoursAndMilesOnEdit(returnTripRange)
    updateTripTimesOnEdit(returnTripRange)
  } else {
    ss.toast("Select a cell in a trip to create its return trip.")
  }
}

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

function isCompleteTrip(trip) {
  return (trip["Trip Date"] && trip["Customer Name and ID"])
}
