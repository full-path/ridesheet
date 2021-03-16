// Once a customer is selected for a trip, fill in trip data with essential data to all trips
// (customer ID and Trip ID), and then any default values from the custome record.
// Home address is a special case if there isn't a designated default PU address.
function fillTripCells(range) {
  try {
    if (range.getValue()) {
      const flagForDefaultValue = "Default "
      const flagLength = flagForDefaultValue.length
      const ss = SpreadsheetApp.getActiveSpreadsheet()
      const tripRow = getFullRow(range)
      const tripValues = getRangeValuesAsTable(tripRow)[0]
      const filter = function(row) { return row["Customer Name and ID"] === tripValues["Customer Name and ID"] }
      const customerRow = findFirstRowByHeaderNames(ss.getSheetByName("Customers"), filter)
      const defaultValueHeaderNames = Object.keys(customerRow).filter(fieldName => fieldName.slice(0,flagLength) === "Default ")
      let valuesToChange = {}
      valuesToChange["Customer ID"] = customerRow["Customer ID"]
      if (tripValues["Trip ID"] == '') { valuesToChange["Trip ID"] = Utilities.getUuid() }
      defaultValueHeaderNames.forEach (defaultValueHeaderName => {
        const tripHeaderName = defaultValueHeaderName.slice(flagLength)
        if (tripValues[tripHeaderName] == '') { valuesToChange[tripHeaderName] = customerRow[defaultValueHeaderName] }
      })
      if (tripValues["PU Address"] == '' && defaultValueHeaderNames.indexOf(flagForDefaultValue + "PU Address") == -1) {
        valuesToChange["PU Address"] = customerRow["Home Address"]
      }
      setValuesByHeaderNames([valuesToChange], tripRow)
      if (valuesToChange["PU Address"] || valuesToChange["DO Address"]) {
        fillHoursAndMilesOnEdit(range)
      }
    }
  } catch(e) { logError(e) }
}

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

function createReturnTrip() {
  try {
    copyTrip(null, true)
  } catch(e) { logError(e) }
}

function addStop() {
  try {
    copyTrip(null, false)
  } catch(e) { logError(e) }
}

function moveTripsToReview() {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet       = ss.getSheetByName("Trips")
    const tripReviewSheet = ss.getSheetByName("Trip Review")
    const tripFilter      = function(row) { return row["Trip Date"] && row["Trip Date"] < dateToday() }
    moveRows(tripSheet, tripReviewSheet, tripFilter)

    const runSheet        = ss.getSheetByName("Runs")
    const runReviewSheet  = ss.getSheetByName("Run Review")
    const runFilter       = function(row) { return row["Run Date"] && row["Run Date"] < dateToday() }
    moveRows(runSheet, runReviewSheet, runFilter)
  } catch(e) { logError(e) }
}

function moveTripsToArchive() {
  try {
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
  } catch(e) { logError(e) }
}

function isCompleteTrip(trip) {
  try {
    return (trip["Trip Date"] && trip["Customer Name and ID"])
  } catch(e) { logError(e) }
}

function isTripWithValidTimes(trip) {
  try {
    return (
      trip["PU Time"] &&
      trip["DO Time"] &&
      Number.isFinite(trip["PU Time"].valueOf()) &&
      Number.isFinite(trip["DO Time"].valueOf()) &&
      trip["DO Time"].valueOf() - trip["PU Time"].valueOf() < 24*60*60*1000 &&
      trip["DO Time"].valueOf() - trip["PU Time"].valueOf() > 0
    )
  } catch(e) {
    logError(e)
    return false
  }
}