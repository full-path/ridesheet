function addDataToRunsInReview() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const ui = SpreadsheetApp.getUi()
    const vehicleSheet = ss.getSheetByName("Vehicles")
    const vehicles = getRangeValuesAsTable(vehicleSheet.getDataRange())
    const runsSheet = ss.getSheetByName("Run Review")
    const tripSheet = ss.getSheetByName("Trip Review")
    const runs = getRangeValuesAsTable(runsSheet.getDataRange())
    const trips = getRangeValuesAsTable(tripSheet.getDataRange())
    const tripReviewCompletedTripResults = getDocProp("tripReviewCompletedTripResults")
    const earliestRunDateReadyForAddingData = runs.filter((row) => {
      return isUserReviewedRun(row) && !isFullyReviewedRun(row)
    }).map((row) => {
      return row["Run Date"].valueOf()
    }).reduce((earliest, thisDate) => {
      return thisDate < earliest ? thisDate : earliest
    })

    const promptResult = ui.prompt("Add Data to Runs in Review",
        "Enter date for runs to add data to. Leave blank for " + formatDate(earliestRunDateReadyForAddingData),
        ui.ButtonSet.OK_CANCEL)
    let date
    if (promptResult.getResponseText() == "") {
      date = new Date(earliestRunDateReadyForAddingData)
    } else {
      date = parseDate(promptResult.getResponseText(),"Invalid Date")
    }

    if (!isValidDate(date)) {
      ui.alert("Invalid date, action cancelled.")
      return
    } else if (promptResult.getSelectedButton() !== ui.Button.OK) {
      ui.alert("Action cancelled as requested.")
      return
    }

    const completedTripsThisDay = trips.filter((row) => {
      return row["Trip Date"].valueOf() === date.valueOf() &&
        tripReviewCompletedTripResults.includes(row["Trip Result"])
    })
    let runsThisDay = runs.filter((row) => row["Run Date"].valueOf() === date.valueOf())
    if (!runsThisDay.length) {
      ui.alert(`No runs found for ${formatDate(date)}. No action taken.`)
      return
    }

    const dataErrorMessages = getDeadheadDataErrorMessages(completedTripsThisDay, runsThisDay)
    if (dataErrorMessages.length) {
      SpreadsheetApp.getUi().alert(
        `Deadhead data could not be added for ${formatDate(date)} to the due to the following ${pluralize(dataErrorMessages.length,"error")}:\n\n- ${dataErrorMessages.join("\n\n- ")}`
      )
    } else {
      newRunData = runs.map((row) => {
        return { _rowPosition: row._rowPosition, _rowIndex: row._rowIndex }
      })
      runsThisDay.forEach((run) => {
        const deadheadData = getDeadheadDataForRun(run, completedTripsThisDay, vehicles)
        let newRunDataRow = newRunData.find((row) => row._rowPosition === run._rowPosition)
        Object.assign(newRunDataRow, deadheadData)
      })
      setValuesByHeaderNames(newRunData, runsSheet.getDataRange())
      ss.toast(`Data successfully added for ${runsThisDay.length} runs`)
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
  const vehicle = vehicles.find((row) => row["Vehicle ID"] === run["Vehicle ID"])

  let result = {}
  result["First PU Address"] = parseAddress(firstTrip["PU Address"]).geocodeAddress
  result["Last DO Address"] = parseAddress(lastTrip["DO Address"]).geocodeAddress
  result["Vehicle Garage Address"] = parseAddress(vehicle["Garage Address"]).geocodeAddress
  const startingDeadheadData = getTripEstimate(result["Vehicle Garage Address"],
        result["First PU Address"], "milesAndDays")
  const endingDeadheadData = getTripEstimate(result["Last DO Address"],
        result["Vehicle Garage Address"], "milesAndDays")
  result["Starting Deadhead Miles"] = startingDeadheadData.miles
  result["Starting Deadhead Hours"] = startingDeadheadData.days
  result["Ending Deadhead Miles"] = endingDeadheadData.miles
  result["Ending Deadhead Hours"] = endingDeadheadData.days
  return result
}

function getDeadheadDataErrorMessages(tripsThisDay, runsThisDay) {
  const result = [
    ...hasOrphans(tripsThisDay, runsThisDay),
    hasDuplicateTrips(tripsThisDay),
    hasDuplicateRuns(runsThisDay),
    hasIncompleteTrips(tripsThisDay),
    hasIncompleteRuns(runsThisDay),
    hasNegativeRunDistance(runsThisDay)
  ].filter((msg) => msg.length > 0)
  return result
}

function hasOrphans(tripsThisDay, runsThisDay) {
  const runKeys = runsThisDay.map((row) => getRunKey(row))
  const runForeignKeys = tripsThisDay.map((row) => getRunKey(row))
  let runKeyErrors = []
  let tripKeyErrors = []
  let runErrorMessage = ""
  let tripErrorMessage = ""
  runKeys.forEach((runKey) => {
    if (runForeignKeys.indexOf(runKey) === -1) runKeyErrors.push(runKey)
  })
  runForeignKeys.forEach((runKey, index) => {
    if (runKeys.indexOf(runKey) === -1) tripKeyErrors.push(getTripKey(tripsThisDay[index]))
  })
  if (runKeyErrors.length) {
    runErrorMessage = `${pluralize(runKeyErrors.length,"run")} with no matching trip:\n-- ${runKeyErrors.join("\n-- ")}`
  }
  if (tripKeyErrors.length) {
    tripErrorMessage = `${pluralize(tripKeyErrors.length,"trip")} with no matching run:\n-- ${tripKeyErrors.join("\n-- ")}`
  }
  return [runErrorMessage, tripErrorMessage]
}

function hasDuplicateRuns(runsThisDay) {
  const runKeys = runsThisDay.map((row) => getRunKey(row))
  const dupeRunKeysWithCount = Object.entries(getDupesWithCount(runKeys)).map(([dupe, count]) => {
    return `${dupe} (${count} occurances)`
  })
  if (dupeRunKeysWithCount.length) {
    return `Duplicate runs:\n-- ${dupeRunKeysWithCount.join("\n-- ")}`
  } else {
    return ""
  }
}

function hasDuplicateTrips(tripsThisDay) {
  const tripKeys = tripsThisDay.map((row) => getTripKey(row))
  const dupeTripKeysWithCount = Object.entries(getDupesWithCount(tripKeys)).map(([dupe, count]) => {
    return `${dupe} (${count} occurances)`
  })
  if (dupeTripKeysWithCount.length) {
    return `Duplicate trips:\n-- ${dupeTripKeysWithCount.join("\n-- ")}`
  } else {
    return ""
  }
}

function hasIncompleteTrips(tripsThisDay) {
  const incompleteTrips = tripsThisDay.filter((row) => !isReviewedTrip(row))
  if (incompleteTrips.length) {
    const incompleteTripKeys = incompleteTrips.map((row) => getTripKey(row))
    return `${pluralize(incompleteTrips.length,"trip")} with incomplete data:\n-- ${incompleteTripKeys.join("\n-- ")}`
  } else {
    return ""
  }
}

function hasIncompleteRuns(runsThisDay) {
  const incompleteRuns = runsThisDay.filter((row) => !isUserReviewedRun(row))
  if (incompleteRuns.length) {
    const incompleteRunKeys = incompleteRuns.map((row) => getRunKey(row))
    return `${pluralize(incompleteRuns.length,"run")} with incomplete data:\n-- ${incompleteRunKeys.join("\n-- ")}`
  } else {
    return ""
  }
}

function hasNegativeRunDistance(runsThisDay) {
  const badRuns = runsThisDay.filter((row) => {
    return (row["Odometer Start"] > row["Odometer End"])
  })
  if (badRuns.length) {
    const badRunKeys = badRuns.map((row) => getRunKey(row))
    return `${pluralize(badRuns.length,"run")} with a negative distance traveled:\n-- ${badRunKeys.join("\n-- ")}`
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
  const blankColumns = runReviewRequiredFields.filter(column => {
    return run[column] === 0 ? false : !run[column]
  })
  return blankColumns.length === 0
}

function isFullyReviewedRun(run) {
  const runReviewRequiredFields = getDocProp("runFullReviewRequiredFields")
  const blankColumns = runReviewRequiredFields.filter(column => {
    return run[column] === 0 ? false : !run[column]
  })
  return blankColumns.length === 0
}

function getRunKey(runOrTrip) {
  return [
    "Driver ID: " + runOrTrip["Driver ID"],
    "Vehicle ID: " + runOrTrip["Vehicle ID"],
    "Run ID: " + (runOrTrip["Run ID"] ? runOrTrip["Run ID"] : "<Blank>")
  ].join(", ")
}

function getTripKey(trip) {
  const tz = getDocProp("localTimeZone")
  return [
    trip["Customer Name and ID"],
    "PU Time: " + (trip["PU Time"] ? Utilities.formatDate(trip["PU Time"],tz,"h:mm a") : "<Blank>")
  ].join(", ")
}

function moveTripsToReview() {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet       = ss.getSheetByName("Trips")
    const tripReviewSheet = ss.getSheetByName("Trip Review")
    const tripFilter      = function(row) { return row["Trip Date"] && row["Trip Date"] < dateToday() }
    moveRows(tripSheet, tripReviewSheet, tripFilter, "Review TS")

    const runSheet        = ss.getSheetByName("Runs")
    const runReviewSheet  = ss.getSheetByName("Run Review")
    const runFilter       = function(row) { return row["Run Date"] && row["Run Date"] < dateToday() }
    moveRows(runSheet, runReviewSheet, runFilter, "Review TS")
  } catch(e) { logError(e) }
}

function moveTripsToArchive() {
  try {
    const ss                      = SpreadsheetApp.getActiveSpreadsheet()
    const tripReviewSheet         = ss.getSheetByName("Trip Review")
    const runReviewSheet          = ss.getSheetByName("Run Review")
    const tripArchiveSheet        = ss.getSheetByName("Trip Archive")
    const runArchiveSheet         = ss.getSheetByName("Run Archive")

    let trips = getRangeValuesAsTable(tripReviewSheet.getDataRange(),{includeFormulaValues: false})
    let runs  = getRangeValuesAsTable(runReviewSheet.getDataRange(),{includeFormulaValues: false})
    let allDates = Array.from(new Set([...trips.map((row) => row["Trip Date"].valueOf()), ...runs.map((row) => row["Run Date"].valueOf())]))
    let moveDates = []

    allDates.forEach((date) => {
      const theseTrips = trips.filter((row) => row["Trip Date"].valueOf() === date)
      const theseRuns = runs.filter((row) => row["Run Date"].valueOf() === date)
      const incompleteTrips = theseTrips.filter((row) => !isReviewedTrip(row))
      const incompleteRuns = theseRuns.filter((row) => !isFullyReviewedRun(row))
      if (theseTrips.length &&
          theseRuns.length &&
          !incompleteTrips.length &&
          !incompleteRuns.length
      ) moveDates.push(date)
    })

    moveRows(tripReviewSheet, tripArchiveSheet, function(row){
      return moveDates.find(thisDate => thisDate.valueOf() === row["Trip Date"].valueOf())
    }, "Archive TS")
    moveRows(runReviewSheet, runArchiveSheet, function(row){
      return moveDates.find(thisDate => thisDate.valueOf() === row["Run Date"].valueOf())
    }, "Archive TS")
  } catch(e) { logError(e) }
}

function moveRows(sourceSheet, destSheet, filter, timestampColName) {
  try {
    const sourceData = getRangeValuesAsTable(sourceSheet.getDataRange(), {includeFormulaValues: false})
    const rowsToMove = sourceData.filter(row => filter(row))
    if (rowsToMove.length < 1) {
      return
    }
    const rowsMovedSuccessfully = createRows(destSheet, rowsToMove, timestampColName)
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

function getDupesWithCount(arr) {
  const counts = {}
  const dupes = {}
  arr.forEach((value) => {
    counts[value] = (counts[value] || 0) + 1
    if (counts[value] > 1) dupes[value] = counts[value]
  })
  return dupes
}