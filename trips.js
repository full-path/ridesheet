function copyTrip(sourceTripRange, isReturnTrip) {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet       = ss.getActiveSheet()
    if (!sourceTripRange) sourceTripRange = getFullRow(tripSheet.getActiveCell())
    const sourceTripRow   = sourceTripRange.getRow()
    const sourceTripData  = getRangeValuesAsTable(sourceTripRange)[0]
    let   newTripData     = {...sourceTripData}
    if (tripSheet.getName() === "Trips" && isCompleteTrip(sourceTripData)) {
      newTripData["PU Address"] = sourceTripData["DO Address"]
      newTripData["DO Address"] = (isReturnTrip ? sourceTripData["PU Address"] : null)
      if (sourceTripData["Appt Time"]) {
        newTripData["PU Time"] = timeAdd(sourceTripData["Appt Time"], 60*60*1000)
      } else if (sourceTripData["DO Time"]) {
        newTripData["PU Time"] = timeAdd(sourceTripData["DO Time"], 60*60*1000)
      } else {
        newTripData["PU Time"] = null
      }
      newTripData["Earliest PU Time"] = null
      newTripData["Latest PU Time"]   = null
      newTripData["DO Time"]          = null
      newTripData["Appt Time"]        = null
      newTripData["Est Hours"]        = null
      newTripData["Est Miles"]        = null
      newTripData["Trip ID"]          = Utilities.getUuid()
      newTripData["Calendar ID"]      = null
      tripSheet.insertRowAfter(sourceTripRow)
      let newTripRange = getFullRow(tripSheet.getRange(sourceTripRow + 1, 1))
      setValuesByHeaderNames([newTripData],newTripRange)
      if (isReturnTrip) {
        fillHoursAndMilesOnEdit(newTripRange)
        updateTripTimesOnEdit(newTripRange)
      }
    } else {
      ss.toast("Select a cell in a trip to create its return trip.","Trip Creation Failed")
    }
  } catch(e) { logError(e) }
}

function createReturnTrip() { copyTrip(null, true) }

function addStop() { copyTrip(null, false) }

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
