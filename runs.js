/**
 * @fileoverview Run record management for RideSheet.
 *
 * A "run" represents a driver + vehicle assignment for a date, potentially
 * covering multiple trips. Two run-creation strategies exist, controlled by
 * the `createRunMode` document property:
 *
 * - **`"default"`** — Runs are pre-created by staff in the Runs sheet either manually 
 *    or by using `buildRunsFromTemplate()`. Trips are later associated with runs
 *    manually or auto-matched by `completeTripRunValues()` (trips.js).
 * - **`"auto"`** — Runs are generated automatically at review time by
 *   `createRunsInReview()`, which groups the trips being moved to review by
 *   date + driver + vehicle and synthesises a run record for each group.
 *
 * The shared `runsObject` intermediate format used by `createRunsInReview()`
 * and `updateRunDetails()` is a plain object keyed by a composite trip key:
 * ```
 * {
 *   [tripKey: string]: {
 *     run:   Object,    // run row data (e.g. Run Date, Driver ID, Vehicle ID)
 *     trips: Object[]  // trip row objects that belong to this run
 *   }
 * }
 * ```
 * `tripKey` is constructed as `JSON.stringify(tripDate) + driverId + vehicleId`.
 */

/**
 * Calculates aggregate timing fields on each run entry in a `runsObject` and
 * returns a flat array of run row objects ready to write to a sheet.
 *
 * For each run entry the following fields are computed from its `trips` array:
 * - `"First PU Time"` / `"Scheduled Start Time"` — earliest `"PU Time"` across all
 *   trips that have one, or `null` if no trips have a PU time.
 * - `"Last DO Time"` / `"Scheduled End Time"` — latest `"DO Time"` across all
 *   trips that have one, or `null` if no trips have a DO time.
 * - `"Review TS"` — Timestamp set to the current date/time.
 *
 * @param {Object.<string, {run: Object, trips: Object[]}>} runsObject - The
 *   intermediate runs map (see file overview for structure).
 * @returns {Object[]} Array of run row objects, one per entry in `runsObject`,
 *   with timing fields populated.
 * @throws {Error} Re-throws any internal error with a descriptive message.
 */
function updateRunDetails(runsObject) {
  try {
    for (let tripKey in runsObject) {
      let runEntry = runsObject[tripKey]
      let tripsArray = runEntry.trips

      // Calculate First PU Time
      let puTrips = tripsArray.filter(trip => trip["PU Time"])
      if (puTrips.length === 0) {
        runEntry.run["First PU Time"] = null
        runEntry.run["Scheduled Start Time"] = null
      } else {
        runEntry.run["First PU Time"] = puTrips.reduce(
          (min, trip) => trip["PU Time"] < min ? trip["PU Time"] : min, 
          puTrips[0]["PU Time"]
        )
        runEntry.run["Scheduled Start Time"] = runEntry.run["First PU Time"]
      }

      // Calculate Last DO Time
      let doTrips = tripsArray.filter(trip => trip["DO Time"])
      if (doTrips.length === 0) {
        runEntry.run["Last DO Time"] = null
        runEntry.run["Scheduled End Time"] = null
      } else {
        runEntry.run["Last DO Time"] = doTrips.reduce(
          (max, trip) => trip["DO Time"] > max ? trip["DO Time"] : max, 
          doTrips[0]["DO Time"]
        )
        runEntry.run["Scheduled End Time"] = runEntry.run["Last DO Time"]
      }
      runEntry.run["Review TS"] = new Date()
    }

    return Object.values(runsObject).map(entry => entry.run)

  } catch(e) { 
    logError(e)
    throw new Error(`Failed to update run details: ${e.message}`)
  }
}

/**
 * Fills the Runs sheet with one week of run entries generated from the
 * "Run Template" sheet. Used in `"default"` run mode.
 *
 * The default start date is the day after the latest existing `"Run Date"`
 * in the Runs sheet (or today if the sheet is empty). The user is prompted
 * to confirm this date or enter a custom one.
 *
 * For each of the 7 days starting from the chosen date, every template row
 * whose `"Days of Week"` field includes that day's name is matched. Matching
 * templates are sorted by `"Scheduled Start Time"` and appended to the Runs
 * sheet as new rows with `"Run Date"`, `"Driver ID"`, `"Vehicle ID"`,
 * `"Scheduled Start Time"`, and `"Scheduled End Time"` populated.
 * `applySheetFormatsAndValidation()` is called on the new rows when done.
 */
function buildRunsFromTemplate() {
  const weekday = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let runsSheet = ss.getSheetByName("Runs") 
  let runTemplateSheet = ss.getSheetByName("Run Template")
  let runs = getRangeValuesAsTable(runsSheet.getDataRange())
  let runTemplates = getRangeValuesAsTable(runTemplateSheet.getDataRange())
  let runDateRaw = runs.reduce((latest, run) => 
    run["Run Date"] > latest ? run["Run Date"] : latest, 
    runs[0]["Run Date"]
  )

  let runDate
  if (!runDateRaw) {
    runDate = new Date()
  } else {
    runDate = new Date(runDateRaw)
  }

  runDate.setDate(runDate.getDate() + 1) 
  
  const ui = SpreadsheetApp.getUi()
  const response = ui.alert(
    'Generate Runs', 
    'Would you like to generate runs starting from ' + formatDate(runDate) + '? Click No to enter a custom date.',
    ui.ButtonSet.YES_NO_CANCEL
  )

  let startDate
  if (response == ui.Button.YES) {
    startDate = runDate
  } else if (response == ui.Button.NO) {
    const promptResult = ui.prompt(
      'Enter Start Date',
      'Enter the date to start generating runs from (MM/DD/YYYY):',
      ui.ButtonSet.OK_CANCEL
    )
    if (promptResult.getSelectedButton() == ui.Button.OK) {
      startDate = parseDate(promptResult.getResponseText())
      if (!isValidDate(startDate)) {
        ui.alert('Invalid date entered. Operation cancelled.')
        return
      }
    } else {
      ss.toast('Action cancelled')
      return
    }
  } else {
    ss.toast('Action cancelled') 
    return
  }

  const lastRow = runsSheet.getLastRow()

  for (let i = 0; i < 7; i++) {
    let currentDate = new Date(startDate)
    currentDate.setDate(startDate.getDate() + i)
    let currentDayOfWeek = weekday[currentDate.getDay()]
    
    const matchingTemplates = runTemplates.filter(row => 
      row["Days of Week"] && row["Days of Week"].includes(currentDayOfWeek)
    ).sort((a,b) => {
      return a["Scheduled Start Time"] - b["Scheduled Start Time"]
    })

    matchingTemplates.forEach(template => {
      const newRun = {
        "Run Date": formatDate(currentDate),
        "Driver ID": template["Driver ID"],
        "Vehicle ID": template["Vehicle ID"],
        "Scheduled Start Time": template["Scheduled Start Time"],
        "Scheduled End Time": template["Scheduled End Time"]
      }
      createRow(runsSheet, newRun)
    })
  }
  applySheetFormatsAndValidation(runsSheet, lastRow + 1)
}

/**
 * Generates run records in the Run Review sheet from an array of trips being
 * moved to review. Used in `"auto"` run mode (see `moveTripsToReview()` in
 * review.js).
 *
 * Trips are grouped by a composite key of `Trip Date + Driver ID + Vehicle ID`.
 * Each unique combination becomes one run. After grouping, `updateRunDetails()`
 * is called to compute timing fields, and the resulting run rows are appended
 * to the Run Review sheet sorted by `"Run Date"`.
 * `applySheetFormatsAndValidation()` is called on the newly added rows.
 *
 * @param {Object[]} trips - Array of trip row objects (as returned by
 *   `getRangeValuesAsTable()`) that are being moved to Trip Review.
 * @returns {void}
 * @throws {Error} Re-throws any internal error with a descriptive message.
 */
function createRunsInReview(trips) {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet()
    let runReviewSheet = ss.getSheetByName("Run Review")
    
    let newRunsOut = {}
    
    trips.forEach(tripRow => {
      let tripKey = JSON.stringify(tripRow["Trip Date"]) + 
                    tripRow["Driver ID"] + 
                    tripRow["Vehicle ID"]
      
      if (tripKey in newRunsOut) {
        newRunsOut[tripKey].trips.push(tripRow)
      } else {
        let newRun = {
          "Run Date": tripRow["Trip Date"],
          "Driver ID": tripRow["Driver ID"],
          "Vehicle ID": tripRow["Vehicle ID"]
        }
        newRunsOut[tripKey] = {
          run: newRun,
          trips: [tripRow]
        }
      }
    })

    const runsToCreate = updateRunDetails(newRunsOut)
      .sort((a,b) => a["Run Date"] - b["Run Date"])

    if (runsToCreate.length > 0) {
      runsToCreate.forEach(run => {
        createRow(runReviewSheet, run)
      })
      
      const lastRow = runReviewSheet.getLastRow()
      const startRow = lastRow - runsToCreate.length + 1
      applySheetFormatsAndValidation(runReviewSheet, startRow)
    }

  } catch(e) { 
    logError(e)
    throw new Error(`Failed to create runs in review: ${e.message}`)
  }
}