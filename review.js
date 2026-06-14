/**
 * @fileoverview Trip and run review workflow for RideSheet.
 *
 * Manages the two-stage lifecycle of trip and run data:
 *
 * 1. **Move to review** (`moveTripsToReview`) — moves past-date trips from the
 *    Trips sheet to Trip Review. In `"default"` run mode, past-date runs are
 *    also moved from Runs to Run Review. In `"auto"` mode, runs are instead
 *    generated from the moved trips by `createRunsInReview()` (runs.js).
 *
 * 2. **Move to archive** (`moveTripsToArchive`) — moves fully-reviewed dates
 *    from the Review sheets to the Archive sheets. A date is archivable when
 *    all its trips and runs are complete, or when every trip is cancelled and
 *    there are no completed trip results (e.g. weather cancellation).
 *
 * Also provides `addDataToRunsInReview` for computing and writing deadhead
 *  mileage/hour fields to run rows in Run Review.
 *
 * Validation helpers (`hasOrphans`, `hasDuplicateRuns`, `hasDuplicateTrips`,
 * `hasIncompleteTrips`, `hasIncompleteRuns`, `hasNegativeRunDistance`) return
 * human-readable error message strings (empty string `""` if no problem found).
 *
 * State-check predicates:
 * - `isReviewedTrip`              — trip has a result and all required fields filled
 * - `isTripWithCompletedTripResult` — trip result is one of the completed values
 * - `isUserReviewedRun`           — run has all user-review required fields filled
 * - `isFullyReviewedRun`          — run has all full-review required fields filled
 */

/**
 * Prompts the user for a date, then computes and writes deadhead mileage and
 * hour fields to all runs on that date in the Run Review sheet.
 *
 * The default date is the earliest run date that is user-reviewed but not yet
 * fully reviewed. Validation checks (`getDeadheadDataErrorMessages`) are run
 * first; if any errors are found an alert is shown and no data is written.
 *
 * On success, writes the following fields to each matching run row:
 * `"First PU Address"`, `"Last DO Address"`, `"Vehicle Garage Address"`,
 * `"Starting Deadhead Miles"`, `"Starting Deadhead Hours"`,
 * `"Ending Deadhead Miles"`, `"Ending Deadhead Hours"`.
 */
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

    // Consider trips that have a blank trip result and that are marked completed in some way.
    // Completed trips will be checked further to ensure that they have any additional required
    // fields filled in.
    const tripsToConsiderThisDay = trips.filter((row) => {
      return row["Trip Date"].valueOf() === date.valueOf() &&
        (!row["Trip Result"] || tripReviewCompletedTripResults.includes(row["Trip Result"]))
    })
    let runsThisDay = runs.filter((row) => row["Run Date"].valueOf() === date.valueOf())
    if (!runsThisDay.length) {
      ui.alert(`No runs found for ${formatDate(date)}. No action taken.`)
      return
    }

    const dataErrorMessages = getDeadheadDataErrorMessages(tripsToConsiderThisDay, runsThisDay)
    if (dataErrorMessages.length) {
      SpreadsheetApp.getUi().alert(
        `Deadhead data could not be added for ${formatDate(date)} to the due to the following ${pluralize(dataErrorMessages.length,"error")}:\n\n- ${dataErrorMessages.join("\n\n- ")}`
      )
    } else {
      let newRunData = runs.map((row) => {
        return { _rowPosition: row._rowPosition, _rowIndex: row._rowIndex }
      })
      runsThisDay.forEach((run) => {
        const deadheadData = getDeadheadDataForRun(run, tripsToConsiderThisDay, vehicles)
        let newRunDataRow = newRunData.find((row) => row._rowPosition === run._rowPosition)
        Object.assign(newRunDataRow, deadheadData)
      })
      setValuesByHeaderNames(newRunData, runsSheet.getDataRange())
      ss.toast(`Data successfully added for ${runsThisDay.length} runs`)
    }
  } catch(e) { logError(e) }
}

/**
 * Computes deadhead distance and duration data for a single run by looking up
 * the first pickup address, last drop-off address, and vehicle garage address,
 * then calling `getTripEstimate()` for the start and end deadhead legs.
 *
 * Addresses are passed through `parseAddress().geocodeAddress` before being
 * sent to the Maps API.
 *
 * @param {Object} run - A run row object from the Run Review sheet.
 * @param {Object[]} tripsThisDay - All trip row objects for the same date.
 * @param {Object[]} vehicles - All vehicle row objects from the Vehicles sheet.
 * @returns {{"First PU Address": string, "Last DO Address": string,
 *   "Vehicle Garage Address": string, "Starting Deadhead Miles": number,
 *   "Starting Deadhead Hours": number, "Ending Deadhead Miles": number,
 *   "Ending Deadhead Hours": number}} Deadhead data object ready to be merged
 *   into the run row.
 */
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
  result["Starting Deadhead Hours"] = Math.round(startingDeadheadData.days * 1440) / 1440
  result["Ending Deadhead Miles"] = endingDeadheadData.miles
  result["Ending Deadhead Hours"] = Math.round(endingDeadheadData.days * 1440) / 1440
  return result
}

/**
 * Runs all validation checks against the trips and runs for a single date and
 * returns an array of non-empty error message strings. Used by
 * `addDataToRunsInReview()` to gate whether deadhead data can be written.
 * @param {Object[]} tripsThisDay - Trip row objects for the date being validated.
 * @param {Object[]} runsThisDay - Run row objects for the date being validated.
 * @returns {string[]} Array of error message strings; empty if all checks pass.
 */
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

/**
 * Checks for orphaned runs (runs with no matching trip) and orphaned trips
 * (trips with no matching run) by comparing run keys across both sets.
 * A "run key" is a composite of Driver ID + Vehicle ID + Run ID.
 * @param {Object[]} tripsThisDay - Trip row objects for the date being checked.
 * @param {Object[]} runsThisDay - Run row objects for the date being checked.
 * @returns {[string, string]} A two-element array: the first element is an error
 *   message for orphaned runs (empty string if none), the second for orphaned
 *   trips (empty string if none).
 */
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

/**
 * Checks for runs that share the same composite run key (Driver ID + Vehicle ID
 * + Run ID) on the same date.
 * @param {Object[]} runsThisDay - Run row objects for the date being checked.
 * @returns {string} An error message listing duplicate run keys, or `""` if none.
 */
function hasDuplicateRuns(runsThisDay) {
  const runKeys = runsThisDay.map((row) => getRunKey(row))
  const dupeRunKeysWithCount = Object.entries(getDupesWithCount(runKeys)).map(([dupe, count]) => {
    const msg = `${dupe} (${count} occurances)`
    return msg
  })
  if (dupeRunKeysWithCount.length) {
    const msg = `Duplicate runs:\n-- ${dupeRunKeysWithCount.join("\n-- ")}`
    return msg
  } else {
    return ""
  }
}

/**
 * Checks for trips that share the same composite trip key (Customer Name and ID
 * + PU Time) on the same date.
 * @param {Object[]} tripsThisDay - Trip row objects for the date being checked.
 * @returns {string} An error message listing duplicate trip keys, or `""` if none.
 */
function hasDuplicateTrips(tripsThisDay) {
  const tripKeys = tripsThisDay.map((row) => getTripKey(row))
  const dupeTripKeysWithCount = Object.entries(getDupesWithCount(tripKeys)).map(([dupe, count]) => {
    const msg = `${dupe} (${count} occurances)`
    return msg
  })
  if (dupeTripKeysWithCount.length) {
    const msg = `Duplicate trips:\n-- ${dupeTripKeysWithCount.join("\n-- ")}`
    return msg
  } else {
    return ""
  }
}

/**
 * Checks whether any trips in the set fail `isReviewedTrip()`.
 * @param {Object[]} tripsThisDay - Trip row objects for the date being checked.
 * @returns {string} An error message listing incomplete trip keys, or `""` if none.
 */
function hasIncompleteTrips(tripsThisDay) {
  const incompleteTrips = tripsThisDay.filter((row) => !isReviewedTrip(row))
  if (incompleteTrips.length) {
    const incompleteTripKeys = incompleteTrips.map((row) => getTripKey(row))
    const msg = `${pluralize(incompleteTrips.length,"trip")} with incomplete data:\n-- ${incompleteTripKeys.join("\n-- ")}`
    return msg
  } else {
    return ""
  }
}

/**
 * Checks whether any runs in the set fail `isUserReviewedRun()`.
 * @param {Object[]} runsThisDay - Run row objects for the date being checked.
 * @returns {string} An error message listing incomplete run keys, or `""` if none.
 */
function hasIncompleteRuns(runsThisDay) {
  const incompleteRuns = runsThisDay.filter((row) => !isUserReviewedRun(row))
  if (incompleteRuns.length) {
    const incompleteRunKeys = incompleteRuns.map((row) => getRunKey(row))
    const msg = `${pluralize(incompleteRuns.length,"run")} with incomplete data:\n-- ${incompleteRunKeys.join("\n-- ")}`
    return msg
  } else {
    return ""
  }
}

/**
 * Checks whether any runs have an odometer start value greater than the
 * odometer end value, indicating a data entry error.
 * @param {Object[]} runsThisDay - Run row objects for the date being checked.
 * @returns {string} An error message listing affected run keys, or `""` if none.
 */
function hasNegativeRunDistance(runsThisDay) {
  const badRuns = runsThisDay.filter((row) => {
    return (row["Odometer Start"] > row["Odometer End"])
  })
  if (badRuns.length) {
    const badRunKeys = badRuns.map((row) => getRunKey(row))
    const msg = `${pluralize(badRuns.length,"run")} with a negative distance traveled:\n-- ${badRunKeys.join("\n-- ")}`
    return msg
  } else {
    return ""
  }
}

/**
 * Returns `true` if a trip row is considered fully reviewed and ready to archive.
 *
 * Logic:
 * - If `"Trip Result"` is blank → `false`.
 * - If `"Trip Result"` is one of the `tripReviewCompletedTripResults` values,
 *   checks that every field in `tripReviewRequiredFields` is non-blank → `true`
 *   only if all required fields are filled.
 * - Otherwise (any other trip result, e.g. a cancellation code) → `true`.
 *
 * @param {Object} trip - A trip row object from the Trip Review sheet.
 * @returns {boolean}
 */
function isReviewedTrip(trip) {
    const tripReviewRequiredFields       = getDocProp("tripReviewRequiredFields")
    const tripReviewCompletedTripResults = getDocProp("tripReviewCompletedTripResults")
    if (!trip["Trip Result"]) {
      return false
    } else if (tripReviewCompletedTripResults.includes(trip["Trip Result"])) {
      const blankColumns = tripReviewRequiredFields.filter(column => !trip[column])
      return blankColumns.length === 0
    } else {
      return true
    }
}

/**
 * Returns `true` if the trip's `"Trip Result"` value is one of the
 * `tripReviewCompletedTripResults` document property values (e.g. `"Completed"`).
 * Returns `false` if `"Trip Result"` is blank or is a non-completed value
 * (e.g. a cancellation code).
 * @param {Object} trip - A trip row object.
 * @returns {boolean}
 */
function isTripWithCompletedTripResult(trip) {
  const tripReviewCompletedTripResults = getDocProp("tripReviewCompletedTripResults")
  if (!trip["Trip Result"]) {
    return false
  } else if (tripReviewCompletedTripResults.includes(trip["Trip Result"])) {
    return true
  } else {
    return false
  }
}

/**
 * Returns `true` if all fields listed in the `runUserReviewRequiredFields`
 * document property are non-blank on the run row. Numeric `0` is treated as
 * a valid (non-blank) value.
 * @param {Object} run - A run row object from the Run Review sheet.
 * @returns {boolean}
 */
function isUserReviewedRun(run) {
  const runReviewRequiredFields = getDocProp("runUserReviewRequiredFields")
  const blankColumns = runReviewRequiredFields.filter(column => {
    return run[column] === 0 ? false : !run[column]
  })
  return blankColumns.length === 0
}

/**
 * Returns `true` if all fields listed in the `runFullReviewRequiredFields`
 * document property are non-blank on the run row. Numeric `0` is treated as
 * a valid (non-blank) value.
 * @param {Object} run - A run row object from the Run Review sheet.
 * @returns {boolean}
 */
function isFullyReviewedRun(run) {
  const runReviewRequiredFields = getDocProp("runFullReviewRequiredFields")
  const blankColumns = runReviewRequiredFields.filter(column => {
    return run[column] === 0 ? false : !run[column]
  })
  return blankColumns.length === 0
}

/**
 * Builds a human-readable composite key string for a run or trip row, used
 * to identify a run when reporting validation errors. The key is formatted as:
 * `"Driver ID: <value>, Vehicle ID: <value>, Run ID: <value|<Blank>>"`.
 * @param {Object} runOrTrip - A run or trip row object with `"Driver ID"`,
 *   `"Vehicle ID"`, and `"Run ID"` fields.
 * @returns {string}
 */
function getRunKey(runOrTrip) {
  return [
    "Driver ID: " + runOrTrip["Driver ID"],
    "Vehicle ID: " + runOrTrip["Vehicle ID"],
    "Run ID: " + (runOrTrip["Run ID"] ? runOrTrip["Run ID"] : "<Blank>")
  ].join(", ")
}

/**
 * Builds a human-readable composite key string for a trip row, used to
 * identify a trip when reporting validation errors. The key is formatted as:
 * `"<Customer Name and ID>, PU Time: <H:MM a|<Blank>>"`.
 * @param {Object} trip - A trip row object with `"Customer Name and ID"` and
 *   `"PU Time"` fields.
 * @returns {string}
 */
function getTripKey(trip) {
  const tz = getDocProp("localTimeZone")
  return [
    trip["Customer Name and ID"],
    "PU Time: " + (trip["PU Time"] ? Utilities.formatDate(trip["PU Time"],tz,"h:mm a") : "<Blank>")
  ].join(", ")
}

/**
 * Moves past-date trips from the Trips sheet to Trip Review, and handles
 * runs according to the `createRunMode` document property:
 * - **`"default"`** — moves past-date runs from the Runs sheet to Run Review.
 * - **`"auto"`** — calls `createRunsInReview()` to generate run records from
 *   the moved trips.
 * - Any other value — shows a toast error and logs the problem.
 *
 * A trip is eligible if its `"Trip Date"` is before today, or if it has any
 * value in `"Trip Result"`.
 * A run is eligible if its combined `"Run Date"` + `"Scheduled End Time"` is
 * before now; if `"Scheduled End Time"` is blank, falls back to date-only
 * comparison using `dateToday()`.
 */
function moveTripsToReview() {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet       = ss.getSheetByName("Trips")
    const tripReviewSheet = ss.getSheetByName("Trip Review")
    const tripFilter      = function(row) { return row["Trip Date"] && (row["Trip Date"] < dateToday() || row["Trip Result"]) }
    const movedTrips      = moveRows(tripSheet, tripReviewSheet, tripFilter, "Review TS")

    const runMode = getDocProp("createRunMode")
    if (runMode === "default") {
      const runSheet        = ss.getSheetByName("Runs")
      const runReviewSheet  = ss.getSheetByName("Run Review")
      const runFilter       = function(row) {
        if (!row["Run Date"]) return false
        if (row["Scheduled End Time"]) return combineDateAndTime(row["Run Date"], row["Scheduled End Time"]) < new Date()
        return row["Run Date"] < dateToday()
      }
      moveRows(runSheet, runReviewSheet, runFilter, "Review TS")
    }
    else if (runMode === "auto") {
      createRunsInReview(movedTrips)
    } else {
      const errorMessage = `Invalid runMode setting: ${runMode}`
      ss.toast(errorMessage)
      logError(errorMessage)
    }
  } catch(e) { logError(e) }
}

/**
 * Moves fully-reviewed dates from Trip Review and Run Review to their
 * respective Archive sheets.
 *
 * A date is eligible to archive if it meets one of two conditions:
 * - **Normal completion**: the date has at least one trip and at least one run,
 *   and no trip fails `isReviewedTrip()` and no run fails `isFullyReviewedRun()`.
 * - **All-cancelled**: the date has trips, no runs, no incomplete trips, and no
 *   trip has a completed trip result (e.g. an all-cancellation/weather day).
 *
 * Rows are moved via `moveRows()`, which appends to the archive sheet and
 * deletes from the review sheet. An `"Archive TS"` timestamp is set on each
 * moved row.
 */
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
      const tripsWithCompletedTripResults = theseTrips.filter((row) => isTripWithCompletedTripResult(row))
      const incompleteRuns = theseRuns.filter((row) => !isFullyReviewedRun(row))
      // Where there are trips and runs, and no trip or run is incomplete
      if (theseTrips.length &&
          theseRuns.length &&
          !incompleteTrips.length &&
          !incompleteRuns.length) {
        moveDates.push(date)
      // Where every trip is cancelled, so there's no run info. Good for weather events.
      } else if (theseTrips.length &&
          !theseRuns.length &&
          !incompleteTrips.length &&
          !tripsWithCompletedTripResults.length
          ) {
        moveDates.push(date)
      }
    })

    moveRows(tripReviewSheet, tripArchiveSheet, function(row){
      return moveDates.find(thisDate => thisDate.valueOf() === row["Trip Date"].valueOf())
    }, "Archive TS")
    moveRows(runReviewSheet, runArchiveSheet, function(row){
      return moveDates.find(thisDate => thisDate.valueOf() === row["Run Date"].valueOf())
    }, "Archive TS")
  } catch(e) { logError(e) }
}

/**
 * Moves rows matching a filter from one sheet to another, optionally
 * stamping a timestamp column, then deletes the moved rows from the source.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sourceSheet - The sheet to move rows from.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destSheet - The sheet to move rows to.
 * @param {function(Object): boolean} filter - Predicate function; rows for which
 *   it returns truthy are moved.
 * @param {string} timestampColName - Column name to stamp with the current
 *   date/time on each moved row.
 * @returns {Object[]} The array of row objects that were moved (or `[]` on error
 *   or if no rows matched).
 */
function moveRows(sourceSheet, destSheet, filter, timestampColName) {
  try {
    const sourceData = getRangeValuesAsTable(sourceSheet.getDataRange(), {includeFormulaValues: false})
    const rowsToMove = sourceData.filter(row => filter(row))
    if (rowsToMove.length < 1) {
      return []
    }
    const rowsMovedSuccessfully = createRows(destSheet, rowsToMove, timestampColName)
    if (rowsMovedSuccessfully) {
      safelyDeleteRows(sourceSheet, rowsToMove)
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('Error moving data. Please check for duplicate entries.')
    }
    return rowsToMove
  } catch(e) {
    logError(e)
    return []
  }
}

/**
 * Moves a single row from its current sheet to a destination sheet using
 * `createRow()` / `safelyDeleteRow()`. Optionally merges extra field values
 * into the row data before writing.
 * @param {GoogleAppsScript.Spreadsheet.Range} sourceRange - A range representing
 *   the row to move (typically a full-row range).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destSheet - The sheet to move the row to.
 * @param {Object} [options]
 * @param {Object} [options.extraFields={}] - Additional key/value pairs to merge
 *   into the row data before it is written to the destination sheet.
 */
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

/**
 * Counts duplicate values in an array and returns an object containing only
 * the values that appear more than once, mapped to their occurrence count.
 * @param {Array} arr - The array to check for duplicates.
 * @returns {Object.<string, number>} Map of duplicate value → count.
 *   Only entries with count > 1 are included.
 */
function getDupesWithCount(arr) {
  const counts = {}
  const dupes = {}
  arr.forEach((value) => {
    counts[value] = (counts[value] || 0) + 1
    if (counts[value] > 1) dupes[value] = counts[value]
  })
  return dupes
}
