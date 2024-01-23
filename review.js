function addDeadheadDataToRunsForDate(date) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const vehicleSheet = ss.getSheetByName("Vehicles")
    const vehicles = getRangeValuesAsTable(vehicleSheet.getDataRange())
    const runsSheet = ss.getSheetByName("Run Review")
    const tripSheet = ss.getSheetByName("Trip Review")
    const trips = getRangeValuesAsTable(tripSheet.getDataRange())
    const tripReviewCompletedTripResults = getDocProp("tripReviewCompletedTripResults")
    const completedTripsThisDay = trips.filter((row) => {
      return row["Trip Date"].valueOf() === date.valueOf() &&
        tripReviewCompletedTripResults.includes(row["Trip Result"])
    })
    let runs = getRangeValuesAsTable(runsSheet.getDataRange())
    let runsThisDay = runs.filter((row) => row["Run Date"].valueOf() === date.valueOf())

    const dataErrorMessages = getDeadheadDataErrorMessages(completedTripsThisDay, runsThisDay)
    if (dataErrorMessages.length) {
      SpreadsheetApp.getUi().alert(
        `Deadhead data could not be added for ${formatDate(date)} to the due to the following error${dataErrorMessages.length === 1 ? '' : 's'}:\n\n- ${dataErrorMessages.join("\n\n- ")}`
      )
    } else {
      runsThisDay.forEach((run) => {
        const deadheadData = getDeadheadDataForRun(run, completedTripsThisDay, vehicles)
        run = {...run, ...deadheadData}
      })
      setValuesByHeaderNames(runs, runsSheet.getDataRange())
    }
  } catch(e) { logError(e) }
}

function getDeadheadDataForRun(run, tripsThisDay, vehicles) {
  const tripsThisRun = tripsThisDay.
    filter((row) => 
      row["Driver ID"] === run["Driver ID"] &&
      row["Vehicle ID"] === run["Vehicle ID"] &&
      row["Run ID"] === run["Run ID"]
    )
  const firstTrip = tripsThisRun.
    reduce((earliestTrip, row) => 
      timeOnlyAsMilliseconds(row["PU Time"]) < timeOnlyAsMilliseconds(earliestTrip["PU Time"]) ? 
      row : earliestTrip
    )
  const lastTrip = tripsThisRun.
    reduce((latestRow, row) => 
      timeOnlyAsMilliseconds(row["PU Time"]) > timeOnlyAsMilliseconds(latestRow["PU Time"]) ? 
      row : latestRow
    )
  const vehicle = vehicles.find((row) => row["Vehicle ID"] === vehicleId)

  result = {}
  result["First PU Address"] = parseAddress(firstTrip["PU Address"]).geocodeAddress
  result["Last DO Address"] = parseAddress(lastTrip["DO Address"]).geocodeAddress
  result["Vehicle Garage Address"] = parseAddress(vehicle["Garage Address"]).geocodeAddress
  result["Starting Deadhead Miles"] =
      getTripEstimate(result["Vehicle Garage Address"], result["First PU Address"], "miles")
  result["Ending Deadhead Miles"] =
      getTripEstimate(result["Last DO Address"], result["Vehicle Garage Address"], "miles")
  return result
}

function getDeadheadDataErrorMessages(tripsThisDay, runsThisDay) {
  return [
    ...hasOrphans(tripsThisDay, runsThisDay),
    hasDuplicateRuns(runsThisDay),
    hasIncompleteTrips(tripsThisDay),
    hasIncompleteRuns(runsThisDay)
  ].filter((msg) => msg.length > 0)
}

function hasOrphans(tripsThisDay, runsThisDay) {
  const runKeys = runsThisDay.map((row) => getRunKey(row))
  const tripKeys = tripsThisDay.map((row) => getRunKey(row))
  let runKeyErrors = []
  let tripKeyErrors = []
  let runErrorMessage = ""
  let tripErrorMessage = ""
  runKeys.forEach((runKey) => {
    if (!tripKeys.indexOf(tripKey) === -1) runKeyErrors.push(runKey)
  })
  tripKeys.forEach((tripKey) => {
    if (runKeys.indexOf(tripKey) === -1) tripKeyErrors.push(runKey)
  })
  if (runKeyErrors.length) {
    runErrorMessage = (runKeyErrors.length === 1 ? "1 run" : runKeyErrors.length + " runs") +
      " with no matching trips:\n" + 
      runKeyErrors.join("\n")
  }
  if (tripKeyErrors.length) {
    tripErrorMessage = (tripKeyErrors.length === 1 ? "1 trip" : tripKeyErrors.length + " trips") +
      " with no matching runs:\n" + 
      tripKeyErrors.join("\n")
  }
  return [runErrorMessage, tripErrorMessage]
}

function hasDuplicateRuns(runsThisDay) {
  const runKeys = runsThisDay.map((row) => getRunKey(row))
  if (new Set(runKeys).size !== runKeys.length) {
    return "There are duplicate runs for this day."
  } else {
    return
  }
}

function hasIncompleteTrips(tripsThisDay) {
  const incompleteTrips = tripsThisDay.filter(!isReviewedTrip)
  if (incompleteTrips.length) {
    return "There are " + incompleteTrips.length + " trips with incomplete data."
  } else {
    return ""
  }
}

function hasIncompleteRuns(runsThisDay) {
  const incompleteRuns = runsThisDay.filter(!isUserReviewedRun)
  if (incompleteRuns.length) {
    return "There are " + incompleteRuns.length + " runs with incomplete data."
  } else {
    return ""
  }
}

function isReviewedTrip(trip) {
    const tripReviewRequiredFields       = getDocProp("tripReviewRequiredFields")
    const tripReviewCompletedTripResults = getDocProp("tripReviewCompletedTripResults")
    if (!trip["Trip Result"]) {
      return false
    } else if (tripReviewCompletedTripResults.includes(trip["Trip Result"])) {
      blankColumns = tripReviewRequiredFields.filter(column => !trip[column])
      return blankColumns.length === 0
    } else {
      return true
    }
}

function isUserReviewedRun(run) {
  const runReviewRequiredFields = getDocProp("runUserReviewRequiredFields")
  blankColumns = runReviewRequiredFields.filter(column => !run[column])
  return blankColumns.length === 0
}

function isFullyReviewedRun(run) {
  const runReviewRequiredFields = getDocProp("runFullReviewRequiredFields")
  blankColumns = runReviewRequiredFields.filter(column => !run[column])
  return blankColumns.length === 0
}

function getRunKey(runOrTrip) {
  return [
    "Driver ID:" + runOrTrip["Driver ID"],
    "Vehicle ID:" + runOrTrip["Vehicle ID"],
    "Run ID:" + (runOrTrip["Run ID"] ? runOrTrip["Run ID"] : "<Blank>")
  ].join()
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
    const ss                      = SpreadsheetApp.getActiveSpreadsheet()
    const tripReviewSheet         = ss.getSheetByName("Trip Review")
    const runReviewSheet          = ss.getSheetByName("Run Review")
    const tripArchiveSheet        = ss.getSheetByName("Trip Archive")
    const runArchiveSheet         = ss.getSheetByName("Run Archive")

    let trips = getRangeValuesAsTable(tripArchiveSheet.getDataRange(),{includeFormulaValues: false})
    let runs  = getRangeValuesAsTable(runReviewSheet.getDataRange(),{includeFormulaValues: false})
    let allDates = new Set([...trips.map((row) => row["Trip Date"]), ...runs.map((row) => row["Run Date"])])
    let moveDates = []

    allDates.forEach((date) => {
      const theseTrips = trips.filter((row) => row["Trip Date"].valueOf() === date.valueOf())
      const theseRuns = runs.filter((row) => row["Run Date"].valueOf() === date.valueOf())
      const incompleteTrips = theseTrips.filter((row) => !isReviewedTrip(row))
      const incompleteRuns = theseRuns.filter((row) => !isFullyReviewedRun(row))
      if (!incompleteTrips.length && !incompleteRuns.length) moveDates.push(date)
    })

    moveRows(tripReviewSheet, tripArchiveSheet, function(row){
      return moveDates.find(thisDate => thisDate.valueOf() === row["Trip Date"].valueOf())
    })
    moveRows(runReviewSheet, runArchiveSheet, function(row){
      return moveDates.find(thisDate => thisDate.valueOf() === row["Run Date"].valueOf())
    })
  } catch(e) { logError(e) }
}

function moveRows(sourceSheet, destSheet, filter) {
  try {
    const sourceData = getRangeValuesAsTable(sourceSheet.getDataRange(), {includeFormulaValues: false})
    const rowsToMove = sourceData.filter(row => filter(row))
    if (rowsToMove.length < 1) {
      log('moveRows', 'No data returned by filter. No rows moved.')
      return
    }
    const rowsMovedSuccessfully = createRows(destSheet, rowsToMove)
    if (rowsMovedSuccessfully) {
      safelyDeleteRows(sourceSheet, rowsToMove)
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('Error moving data. Please check for duplicate entries.')
    }
  } catch(e) { logError(e) }
}

function moveRow(sourceRange, destSheet, {extraFields = {}} = {}) {
  try {
    const sourceSheet = sourceRange.getSheet()
    const sourceData = getRangeValuesAsTable(sourceRange, {includeFormulaValues: false})[0]
    Object.keys(extraFields).forEach(key => sourceData[key] = extraFields[key])
    if (createRow(destSheet, sourceData)) {
      safelyDeleteRow(sourceSheet, sourceData)
    }
  } catch(e) { logError(e) }
}