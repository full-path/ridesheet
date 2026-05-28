/**
 * @fileoverview Trip record helpers for RideSheet.
 *
 * Provides functions for creating, copying, and validating trip rows in the
 * Trips sheet, and for auto-filling run association data when a trip has
 * enough information to be matched to a run.
 *
 * Key functions:
 * - `fillTripCells()`         — populate a new trip row with customer defaults when
 *                               a customer is selected.
 * - `copyTrip()`              — duplicate an existing trip, swapping PU/DO addresses,
 *                               and optionally setting the DO address to the customer's
 *                               earliest PU address of the day (return-trip mode).
 * - `createReturnTrip()`      — menu/trigger entry point for return-trip copy.
 * - `addStop()`               — menu/trigger entry point for a forward copy (next stop).
 * - `isCompleteTrip()`        — minimal validity check for a trip row object.
 * - `isTripWithValidTimes()`  — stricter check that PU and DO times are valid and ordered.
 * - `completeTripRunValues()` — auto-fill Vehicle ID / Driver ID / Run ID when a trip
 *                               can be unambiguously matched to a run.
 */

/**
 * When a customer is selected on a new trip row, fills in essential and default
 * trip cell values from the matching customer record.
 *
 * Behaviour:
 * - Always sets `"Customer ID"` from the customer record.
 * - Generates a new UUID for `"Trip ID"` if the cell is currently empty.
 * - Reads every column in the customer record whose header starts with `"Default "`
 *   and copies its value into the corresponding trip column (with the `"Default "`
 *   prefix stripped), but only if that trip cell is currently empty.
 * - If `"PU Address"` is still empty after the above step and the customer record
 *   has no `"Default PU Address"` column, falls back to the customer's
 *   `"Home Address"`.
 * - Triggers `fillHoursAndMilesOnEdit()` if either `"PU Address"` or `"DO Address"`
 *   was set.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - A cell in the trip row that
 *   contains the newly selected `"Customer Name and ID"` value.
 */
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

/**
 * Creates a copy of an existing trip row, swapping the PU and DO addresses,
 * and inserts the new row immediately after the source row.
 *
 * In standard mode (`isReturnTrip = false`):
 * - The new trip's `"PU Address"` is the source trip's `"DO Address"`.
 * - The new trip's `"DO Address"` is `null` (left blank for the user to fill).
 * - `"PU Time"` is calculated by adding `defaultStayDuration` (minutes) to the
 *   source trip's `"Appt Time"` or `"DO Time"` (whichever is available). If
 *   `defaultStayDuration` is `-1`, `"PU Time"` is set to `null`.
 *
 * In return-trip mode (`isReturnTrip = true`):
 * - The new trip's `"DO Address"` is the `"PU Address"` of the **earliest** trip
 *   for the same customer on the same date.
 * - `fillHoursAndMilesOnEdit()` and `updateTripTimesOnEdit()` are called
 *   automatically after the row is inserted.
 *
 * In both modes, `"Earliest PU Time"`, `"Latest PU Time"`, `"DO Time"`,
 * `"Appt Time"`, `"Est Hours"`, `"Est Miles"`, `"Trip ID"`, and `"Calendar ID"`
 * are always reset to `null` on the new trip.
 *
 * Can be called directly (e.g. from a named-range trigger) with a `sourceTripRange`,
 * or called with `null` to use the active cell's row.
 *
 * Aborts with a toast if the source trip fails `isCompleteTrip()`.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range|null} sourceTripRange - A range in the
 *   source trip row, or `null` to use the active cell.
 * @param {boolean} isReturnTrip - When `true`, sets the DO address to the earliest
 *   same-day PU address for the customer.
 */
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
        filter((row) => {
          return row["Customer ID"] === sourceTripData["Customer ID"] &&
            row["Trip Date"] && row["PU Time"] && row["PU Address"] &&
            row["Trip Date"].getTime() === sourceTripData["Trip Date"].getTime()
        })
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

/**
 * Entry point for the "Create Return Trip" menu action and named-range trigger.
 * Calls `copyTrip(null, true)` using the active cell's row as the source.
 */
function createReturnTrip() {
  try {
    copyTrip(null, true)
  } catch(e) { logError(e) }
}

/**
 * Entry point for the "Add Stop" menu action and named-range trigger.
 * Calls `copyTrip(null, false)` using the active cell's row as the source,
 * creating a forward copy (next stop) without setting the DO address.
 */
function addStop() {
  try {
    copyTrip(null, false)
  } catch(e) { logError(e) }
}

/**
 * Returns `true` if the trip row object has the minimum fields required to be
 * considered a valid trip: a `"Trip Date"` value and a `"Customer Name and ID"` value.
 * @param {Object} trip - A trip row object as returned by `getRangeValuesAsTable()`.
 * @returns {boolean}
 */
function isCompleteTrip(trip) {
  try {
    return (trip["Trip Date"] && trip["Customer Name and ID"])
  } catch(e) { logError(e) }
}

/**
 * Returns `true` if the trip has both a `"PU Time"` and a `"DO Time"` that are
 * finite Date values, the DO time is after the PU time, and the difference is
 * less than 24 hours.
 * @param {Object} trip - A trip row object as returned by `getRangeValuesAsTable()`.
 * @returns {boolean}
 */
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

/**
 * Attempts to auto-fill run association fields (`"Vehicle ID"`, `"Driver ID"`,
 * `"Run ID"`) on a trip row when the trip can be unambiguously matched to a run.
 *
 * Matching logic (both branches require `"Trip Date"` to be set):
 * - **Driver ID set, Vehicle ID empty**: searches the Runs sheet for a run on
 *   the same date with the same Driver ID. If exactly one run is found (filtered
 *   further by scheduled start/end times when both PU and DO times are present),
 *   sets `"Vehicle ID"` and `"Run ID"`. If no unique run is found, falls back to
 *   the driver's `"Default Vehicle ID"` from the Drivers sheet.
 * - **Vehicle ID set, Driver ID empty**: searches the Runs sheet for a run on
 *   the same date with the same Vehicle ID (optionally filtered by time). If
 *   exactly one run is found, sets `"Driver ID"` and `"Run ID"`.
 *
 * This function is called on every edit to the Trips sheet via `tripSheetTrigger`
 * regardless of `createRunMode`. It is only meaningful when `createRunMode` is
 * `"default"` (runs are pre-created in the Runs sheet). In `"auto"` mode, no runs
 * exist in the Runs sheet at trip-entry time, so the function always returns `false`
 * without making any changes.
 *
 * Accepts either a `SheetsOnEdit` event object (uses `e.range`) or a Range
 * directly (for non-event callers).
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit|GoogleAppsScript.Spreadsheet.Range} e
 *   An onEdit event object whose `range` property is a cell in the trip row,
 *   or a Range object directly.
 * @returns {boolean} `true` if run values were successfully filled; `false` otherwise.
 */
function completeTripRunValues(e) {
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let range
    try {
      range = e.range
    } catch(error) {
      range = e
    }
    const tripRow = getFullRow(range)
    const tripValues = getRangeValuesAsTable(tripRow)[0]
    const tripDate = tripValues["Trip Date"]
    const tripPuTime = tripValues["PU Time"]
    const tripDoTime = tripValues["DO Time"]
    const tripDriverId = tripValues["Driver ID"]
    const tripVehicleId = tripValues["Vehicle ID"]
    let timeFilter = function(row) { return true }
    let runRows = getRangeValuesAsTable(ss.getSheetByName("Runs").getDataRange())

    if (tripDate) {
      if (tripPuTime && tripDoTime) {
        timeFilter = function(row) {
          return row["Scheduled Start Time"] && row["Scheduled End Time"] &&
            row["Scheduled Start Time"].valueOf() <= tripPuTime.valueOf() &&
            row["Scheduled End Time"].valueOf() >= tripDoTime.valueOf()
        }
      }
      if (tripDriverId && !tripVehicleId) {
        const filter = function(row) {
          return row["Run Date"] &&
            row["Run Date"].valueOf() === tripDate.valueOf() &&
            row["Driver ID"] === tripDriverId
        }
        const filteredRunRows = runRows.filter(filter).filter(timeFilter)
        if (filteredRunRows.length === 1) {
          const runRow = filteredRunRows[0]
          let valuesToChange = {}
          valuesToChange["Vehicle ID"] = runRow["Vehicle ID"]
          valuesToChange["Run ID"] = runRow["Run ID"]
          setValuesByHeaderNames([valuesToChange], tripRow)
          return true
        } else {
          const filter = function(row) { return row["Driver ID"] === tripDriverId && row["Default Vehicle ID"] }
          const driverRow = findFirstRowByHeaderNames(ss.getSheetByName("Drivers"), filter)
          if (driverRow && driverRow["Default Vehicle ID"]) {
            let valuesToChange = {}
            valuesToChange["Vehicle ID"] = driverRow["Default Vehicle ID"]
            setValuesByHeaderNames([valuesToChange], tripRow)
            return true
          }
        }
      } else if (!tripValues["Driver ID"] && tripValues["Vehicle ID"]) {
        const filter = function(row) {
          return row["Run Date"] &&
            row["Run Date"].valueOf() === tripDate.valueOf() &&
            row["Vehicle ID"] === tripVehicleId
        }
        const filteredRunRows = runRows.filter(filter).filter(timeFilter)
        if (filteredRunRows.length === 1) {
          const runRow = filteredRunRows[0]
          let valuesToChange = {}
          valuesToChange["Driver ID"] = runRow["Driver ID"]
          valuesToChange["Run ID"] = runRow["Run ID"]
          setValuesByHeaderNames([valuesToChange], tripRow)
          return true
        }
      }
    }
    return false
  } catch(e) { logError(e) }
}
