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

// Takes an existing trip and makes a copy of it, using the DO address of the existing trip as the
// PU address of the new one. If this isReturnTrip is true, the DO address of the new trip is the
// PU address of the earliest trip for the customer and day
// Function can be called from the Action/Go trigger or from the menu bar. If called from the
// menu bar, the source row is taken from taken from whatever row the active cell is in
function copyTrip(sourceTripRange, isReturnTrip) {
  try {
    const ss                  = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet           = ss.getActiveSheet()
    if (!sourceTripRange) sourceTripRange = getFullRow(tripSheet.getActiveCell())
    const sourceTripRow       = sourceTripRange.getRow()
    const sourceTripData      = getRangeValuesAsTable(sourceTripRange,{includeFormulaValues: false})[0]
    const defaultStayDuration = getDocProp("defaultStayDuration")
    if (!isCompleteTrip(sourceTripData)) {
      ss.toast("Select a cell in a trip to create a subsequent trip.","Trip Creation Failed")
      return
    }
    let DoAddress
    if (isReturnTrip) {
      const allTrips = getRangeValuesAsTable(tripSheet.getDataRange())
      const customerTripsThisDay = allTrips.
        filter((row) => row["Customer ID"] === sourceTripData["Customer ID"] &&
        row["Trip Date"].getTime() === sourceTripData["Trip Date"].getTime())
      const firstCustomerTripThisDay = customerTripsThisDay.
        reduce((earliestRow, row) => timeOnlyAsMilliseconds(row["PU Time"]) < timeOnlyAsMilliseconds(earliestRow["PU Time"]) ? row : earliestRow)
        DoAddress = firstCustomerTripThisDay["PU Address"]
    } else {
      DoAddress = null
    }
    let   newTripData     = {...sourceTripData}
    newTripData["PU Address"] = sourceTripData["DO Address"]
    newTripData["DO Address"] = DoAddress
    if (defaultStayDuration === -1) {
      newTripData["PU Time"] = null
    } else if (sourceTripData["Appt Time"]) {
      newTripData["PU Time"] = timeAdd(sourceTripData["Appt Time"], defaultStayDuration*60*1000)
    } else if (sourceTripData["DO Time"]) {
      newTripData["PU Time"] = timeAdd(sourceTripData["DO Time"], defaultStayDuration*60*1000)
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

// When a trip has enough information so that it can be associated with certainty with a run,
// Fill in the missing data
function completeTripRunValues(e) {
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const range = e.range
    const tripRow = getFullRow(range)
    const tripValues = getRangeValuesAsTable(tripRow)[0]
    if (tripValues["|Run OK?|"] === 1) {
      if (tripValues["Driver ID"] && !tripValues["Vehicle ID"]) {
        const filter = function(row) {
          return row["Run Date"].valueOf() === tripValues["Trip Date"].valueOf() &&
          row["Driver ID"] === tripValues["Driver ID"] &&
          row["Run ID"] === tripValues["Run ID"]
        }
        const runRow = findFirstRowByHeaderNames(ss.getSheetByName("Runs"), filter)
        if (runRow) {
          let valuesToChange = {}
          valuesToChange["Vehicle ID"] = runRow["Vehicle ID"]
          setValuesByHeaderNames([valuesToChange], tripRow)
        }
      } else if (!tripValues["Driver ID"] && tripValues["Vehicle ID"]) {
        const filter = function(row) {
          return row["Run Date"].valueOf() === tripValues["Trip Date"].valueOf() &&
          row["Vehicle ID"] === tripValues["Vehicle ID"] &&
          row["Run ID"] === tripValues["Run ID"]
        }
        const runRow = findFirstRowByHeaderNames(ss.getSheetByName("Runs"), filter)
        if (runRow) {
          let valuesToChange = {}
          valuesToChange["Driver ID"] = runRow["Driver ID"]
          setValuesByHeaderNames([valuesToChange], tripRow)
        }
      }
    }
  } catch(e) { logError(e) }
}
