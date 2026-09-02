/**
 * @fileoverview Spreadsheet onEdit trigger dispatch for RideSheet.
 *
 * Implements a two-stage dispatch architecture so that edits are handled in a
 * predictable order and sheet-specific triggers fire before and after cell
 * triggers:
 *
 * ```
 * onEdit
 *   └─ callLocalSheetTriggers  (initialLocalSheetTriggers — from on_edit_local.js)
 *   └─ callSheetTriggers       (initialSheetTriggers)
 *   └─ callLocalCellTriggers   (from on_edit_local.js)
 *   └─ callCellTriggers        (rangeTriggers — named ranges prefixed with "code")
 *   └─ callSheetTriggers       (finalSheetTriggers)
 *   └─ callLocalSheetTriggers  (finalLocalSheetTriggers — from on_edit_local.js)
 * ```
 *
 * **Sheet triggers** (`initialSheetTriggers`, `finalSheetTriggers`) are plain
 * objects mapping sheet name → handler function. The handler is called with the
 * `onEdit` event object `e`.
 *
 * **Cell triggers** (`rangeTriggers`) are keyed by named range name (must start
 * with `"code"`). Each entry has:
 * - `functionCall` {function} — called with the edited `Range` object.
 * - `callOncePerRow` {boolean} — when `true`, the function is called at most
 *   once per row even if multiple cells in that row are in the named range.
 *
 * Named ranges used as cell triggers and their handlers:
 * - `codeTripActionButton`   → `tripActionButton`    (callOncePerRow: true)
 * - `codeFillRequestCells`   → `fillTripCellsOnEdit` (callOncePerRow: true)
 * - `codeFormatAddress`      → `formatAddressOnEdit` (callOncePerRow: false)
 * - `codeFillHoursAndMiles`  → `fillHoursAndMilesOnEdit` (callOncePerRow: true)
 * - `codeSetCustomerKey`     → `setCustomerKeyOnEdit` (callOncePerRow: true)
 * - `codeScanForDuplicates`  → `scanForDuplicatesOnEdit` (callOncePerRow: false)
 * - `codeUpdateTripTimes`    → `updateTripTimesOnEdit` (callOncePerRow: true)
 * - `codeExpandAddress`      → `expandAddressOnEdit` (callOncePerRow: true)
 */

/**
 * Sheet triggers called at the start of `onEdit`, before cell triggers.
 * Maps sheet name to a handler function receiving the onEdit event `e`.
 * Local overrides/additions are in `initialLocalSheetTriggers` (on_edit_local.js).
 * @type {Object.<string, function(GoogleAppsScript.Events.SheetsOnEdit): void>}
 */
const initialSheetTriggers = {
  "Document Properties": updatePropertiesOnEdit
}

/**
 * Sheet triggers called at the end of `onEdit`, after cell triggers.
 * Maps sheet name to a handler function receiving the onEdit event `e`.
 * Local overrides/additions are in `finalLocalSheetTriggers` (on_edit_local.js).
 * @type {Object.<string, function(GoogleAppsScript.Events.SheetsOnEdit): void>}
 */
const finalSheetTriggers = {
  "Trips": tripSheetTrigger,
  "Runs":  runSheetTrigger
}

/**
 * Cell-level triggers keyed by named range name (prefix `"code"`).
 * Each entry maps to a handler function and a `callOncePerRow` flag.
 * See the file overview for the full list of entries.
 * @type {Object.<string, {functionCall: function(GoogleAppsScript.Spreadsheet.Range): void, callOncePerRow: boolean}>}
 */
const rangeTriggers = {
  codeTripActionButton: {
    functionCall: tripActionButton,
    callOncePerRow: true
  },
  codeFillRequestCells: {
    functionCall: fillTripCellsOnEdit,
    callOncePerRow: true
  },
  codeFormatAddress: {
    functionCall: formatAddressOnEdit,
    callOncePerRow: false
  },
  codeFillHoursAndMiles: {
    functionCall: fillHoursAndMilesOnEdit,
    callOncePerRow: true
  },
  codeSetCustomerKey: {
    functionCall: setCustomerKeyOnEdit,
    callOncePerRow: true
  },
  codeScanForDuplicates: {
    functionCall: scanForDuplicatesOnEdit,
    callOncePerRow: false
  },
  codeUpdateTripTimes: {
    functionCall: updateTripTimesOnEdit,
    callOncePerRow: true
  },
  codeExpandAddress: {
    functionCall: expandAddressOnEdit,
    callOncePerRow: true
  }
}

/**
 * The Google Apps Script onEdit trigger entry point.
 * Dispatches the edit event through the full trigger pipeline in order:
 * initial local sheet triggers, initial sheet triggers, local cell triggers,
 * cell triggers, final sheet triggers, final local sheet triggers.
 * Logs total execution time when `debugLogging` is enabled.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
function onEdit(e) {
  try {
    const startTime = new Date()
    const sheetName = e.range.getSheet().getName()
    if (e.range.getRow() === 1) {
      const hasHeader = e.range.getSheet().createDeveloperMetadataFinder().withKey("hasHeader").find()?.[0]?.getValue()
      if (hasHeader  === "true") fixHeaderNames(e.range)
    }
    callLocalSheetTriggers(e, sheetName, initialLocalSheetTriggers)
    callSheetTriggers(e, sheetName, initialSheetTriggers)
    callLocalCellTriggers(e)
    callCellTriggers(e)
    callSheetTriggers(e, sheetName, finalSheetTriggers)
    callLocalSheetTriggers(e, sheetName, finalLocalSheetTriggers)
  } catch(e) {
    logError(e)
  } finally {
    if (debugLogging) log("onEdit duration:", new Date().getTime() - startTime.getTime())
  }
}

/**
 * Calls the handler for the given sheet name if one exists in `triggers`.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 * @param {string} sheetName - The name of the edited sheet.
 * @param {Object.<string, function>} triggers - A sheet-trigger map.
 */
function callSheetTriggers(e, sheetName, triggers) {
  if (Object.keys(triggers).indexOf(sheetName) !== -1) {
    triggers[sheetName](e)
  }
}

/**
 * Evaluates all `code`-prefixed named ranges that overlap the edited cell(s)
 * and calls the corresponding handler functions from `rangeTriggers`.
 *
 * For multi-cell edits, each cell is evaluated individually (header row cells
 * are skipped). When a trigger has `callOncePerRow: true`, its handler is
 * called at most once per row regardless of how many cells in that row match.
 *
 * Handlers are called in the order their trigger names appear in `rangeTriggers`,
 * not in the order cells were edited.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
function callCellTriggers(e) {
  try {
    const spreadsheet = e.source
    const sheet = e.range.getSheet()
    const allNamedRanges = sheet.getNamedRanges().filter(namedRange =>
      namedRange.getName().indexOf("code") === 0 && rangesOverlap(e.range, namedRange.getRange())
    )
    if (allNamedRanges.length === 0) return

    const isMultiColumnRange = (e.range.getWidth() > 1)
    const isMultiRowRange = (e.range.getHeight() > 1)
    let triggeredRows = {}
    let ranges = []
    let callsToMake = {}
    Object.keys(rangeTriggers).forEach(rangeTrigger => callsToMake[rangeTrigger] = [])

    // Set up the tracking to prevent running some code from running multiple times per row.
    Object.keys(rangeTriggers).forEach(key => {
      if (rangeTriggers[key].callOncePerRow) triggeredRows[key] = []
    })

    // If we're working with multiple rows or columns, collect all the 1-cell ranges we'll be looking at.
    if (isMultiRowRange || isMultiColumnRange) {
      for (let y = e.range.getColumn(); y <= e.range.getLastColumn(); y++) {
        for (let x = e.range.getRow(); x <= e.range.getLastRow(); x++) {
          if (x > 1) ranges.push(sheet.getRange(x,y))
        }
      }
    } else if (e.range.getRow() > 1) {
      ranges.push(e.range)
    }

    // Proceed through the array of 1-cell ranges
    ranges.forEach(range => {
      // For this 1-cell range, collect all the triggers to be triggered.
      let involvedTriggerNames = []
      allNamedRanges.forEach(namedRange => {
        if (isInRange(range, namedRange.getRange())) {
          involvedTriggerNames.push(convertNamedRangeToTriggerName(namedRange))
        }
      })

      // Call all the functions for the triggers involved with this 1-cell range
      involvedTriggerNames.forEach(triggerName => {
        // Check to see if this trigger has a one-call-per-row constraint on it
        if (triggeredRows[triggerName]) {
          // if it hasn't been triggered for this row, trigger and record it.
          if (triggeredRows[triggerName].indexOf(range.getRow()) === -1) {
            callsToMake[triggerName].push(range)
            triggeredRows[triggerName].push(range.getRow())
          }
        } else {
          callsToMake[triggerName]?.push(range)
        }
      })
    })

    Object.keys(callsToMake).forEach(rangeTrigger => {
      callsToMake[rangeTrigger].forEach(range => {
        rangeTriggers[rangeTrigger]["functionCall"](range)
      })
    })
  } catch(e) { logError(e) }
}

/**
 * Formats an address cell when its value changes. First attempts to resolve
 * the value as a short name via the Addresses sheet (`setAddressByShortName()`);
 * if that fails, geocodes it via the Maps API (`setAddressByApi()`). Clears
 * the cell note and background if the cell is blank.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The edited address cell.
 */
function formatAddressOnEdit(range) {
  try {
    if (range.getValue() && range.getValue().toString().trim()) {
      if (!setAddressByShortName(range)) {
        setAddressByApi(range)
      }
    } else {
      range.setNote("")
      range.setBackground(null)
    }
  } catch(e) { logError(e) }
}

/**
 * Expands a short name typed into a `|PU|`/`|DO|` helper cell into the full
 * address in the cell to its right, looked up via `getAddressByShortName()`.
 * On success, clears the helper cell and the address cell's note/background,
 * then recalculates hours/miles for the trip row.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The edited `|PU|`/`|DO|` helper cell.
 */
function expandAddressOnEdit(range) {
  const shortName = range.getValue()
  if (shortName?.toString().trim()) {
    try {
      const targetSheet = range.getSheet()
      const targetRange = targetSheet.getRange(range.getRow(), range.getColumn(), 1, 2)
      const result = getAddressByShortName(shortName)
      if (result) {
        targetRange.setValues([["",result]]).setNotes([["",""]]).setBackground(null)
        fillHoursAndMilesOnEdit(range)
        return true
      }
      return false
    } catch(e) {
      logError(e)
      return false
    }
  }
}

/**
 * Thin wrapper that calls `fillTripCells()` for a cell trigger invocation.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The edited cell.
 */
function fillTripCellsOnEdit(range) {
  try {
    fillTripCells(range)
  } catch(e) { logError(e) }
}

/**
 * Calculates estimated trip hours and miles from the PU and DO addresses in
 * the trip row and writes them back to `"Est Hours"` and `"Est Miles"`. Clears
 * both fields if either address is blank. On success, calls
 * `updateTripTimesOnEdit()` to derive missing PU/DO times from the estimate.
 * Shows a toast on a successful estimate.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - Any cell in the trip row.
 */
function fillHoursAndMilesOnEdit(range) {
  try {
    const tripRow = getFullRow(range)
    const tripValues = getRangeValuesAsTable(tripRow)[0]
    if (tripValues["PU Address"] && tripValues["DO Address"]) {
      const PUAddress = parseAddress(tripValues["PU Address"]).geocodeAddress
      const DOAddress = parseAddress(tripValues["DO Address"]).geocodeAddress
      const tripEstimate = getTripEstimate(PUAddress, DOAddress, "milesAndHours")
      setValuesByHeaderNames([{"Est Hours": tripEstimate.hours, "Est Miles": tripEstimate.miles}], tripRow)
      if (tripEstimate.hours) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Travel estimate saved")
      }
      updateTripTimesOnEdit(range)
    } else {
      setValuesByHeaderNames([{"Est Hours": null, "Est Miles": null}], tripRow)
    }
  } catch(e) { logError(e) }
}

/**
 * Manages customer record setup when the customer's name fields are edited.
 * - Trims first and last name.
 * - Generates a numeric `"Customer ID"` if not already set, using the
 *   `lastCustomerID_` private document property as a counter seed (falling
 *   back to scanning the ID column for the current max).
 * - Constructs `"Customer Name and ID"` from the three key fields.
 * - Updates `lastCustomerID_` when a new or higher ID is encountered.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - Any cell in the customer row.
 */
function setCustomerKeyOnEdit(range) {
  try {
    const customerRow = getFullRow(range)
    const customerValues = getRangeValuesAsTable(customerRow)[0]
    let newValues = {}
    if (customerValues["Customer First Name"] && customerValues["Customer Last Name"]) {
      let lastCustomerID = getDocProp("lastCustomerID_")
      if (!Number.isFinite(lastCustomerID) || lastCustomerID === 0) {
        const sheet = range.getSheet()
        const idColumnPosition = getSheetHeaderNames(sheet).indexOf("Customer ID") + 1
        const idRange = sheet.getRange(1, idColumnPosition, sheet.getLastRow())
        let maxID = getMaxValueInRange(idRange)
        lastCustomerID = Number.isFinite(maxID) ? maxID : 0
      }
      let nextCustomerID = Math.ceil(lastCustomerID) + 1
      // There is no ID. Set one and update the lastCustomerID property
      if (!customerValues["Customer ID"]) {
        newValues["Customer ID"] = nextCustomerID
        newValues["Customer First Name"] = customerValues["Customer First Name"].trim()
        newValues["Customer Last Name"] = customerValues["Customer Last Name"].trim()
        newValues["Customer Name and ID"] = getCustomerNameAndId(newValues["Customer First Name"], newValues["Customer Last Name"], newValues["Customer ID"])
        setDocProp("lastCustomerID_", nextCustomerID)
        // There is an ID value present, and it's numeric.
        // Update the lastCustomerID property if the new ID is greater than the current lastCustomerID property
      } else if (Number.isFinite(customerValues["Customer ID"])) {
        newValues["Customer ID"] = (customerValues["Customer ID"])
        newValues["Customer First Name"] = customerValues["Customer First Name"].trim()
        newValues["Customer Last Name"] = customerValues["Customer Last Name"].trim()
        newValues["Customer Name and ID"] = getCustomerNameAndId(newValues["Customer First Name"], newValues["Customer Last Name"], newValues["Customer ID"])
        if (customerValues["Customer ID"] >= nextCustomerID) { setDocProp("lastCustomerID_", customerValues["Customer ID"]) }
        // There is an ID value, and it's not numeric. Allow this, but don't track it as the lastCustomerID
      } else {
        newValues["Customer First Name"] = customerValues["Customer First Name"].trim()
        newValues["Customer Last Name"] = customerValues["Customer Last Name"].trim()
        newValues["Customer Name and ID"] = getCustomerNameAndId(newValues["Customer First Name"], newValues["Customer Last Name"], newValues["Customer ID"])
      }
      setValuesByHeaderNames([newValues], customerRow)
    }
  } catch(e) { logError(e) }
}

/**
 * Derives missing PU or DO times from `"Est Hours"` and document-property
 * padding/dwell settings. Handles three cases (for each row in the range):
 * - PU Time set, DO Time blank, no Appt Time → computes DO Time.
 * - DO Time set, PU Time blank, no Appt Time → computes PU Time.
 * - Appt Time set, neither PU nor DO Time set → computes DO Time from Appt
 *   Time minus `dropOffToAppointmentTimeInMinutes`, then PU Time from that.
 * Rows without `"Est Hours"` are passed through as empty objects (no change).
 * @param {GoogleAppsScript.Spreadsheet.Range} range - A cell or multi-row range
 *   in the Trips sheet.
 */
function updateTripTimesOnEdit(range) {
  try {
    const tripRows = getFullRows(range)
    const tripValues = getRangeValuesAsTable(tripRows)
    let newValues = []
    tripValues.forEach(row => {
      let newRowValues = {}
      if (row["Est Hours"] && isFinite(row["Est Hours"])) {
        const estMilliseconds = (row["Est Hours"] * 60 * 60 * 1000)
        const estHours = estMilliseconds / 3600000
        const padding = getDocProp("tripPaddingPerHourInMinutes") * estHours * 60000
        const apptPadding = getDocProp("dropOffToAppointmentTimeInMinutes") * 60000
        const dwellTime = getDocProp("dwellTimeInMinutes") * 60000
        const journeyTime = estMilliseconds + padding + dwellTime
        if (row["PU Time"] && !row["DO Time"] && !row["Appt Time"]) {
          newRowValues["DO Time"] = timeAdd(row["PU Time"], journeyTime)
        } else if (!row["PU Time"] && row["DO Time"] && !row["Appt Time"]) {
          newRowValues["PU Time"] = timeAdd(row["DO Time"], -journeyTime)
        } else if (!row["PU Time"] && !row["DO Time"] && row["Appt Time"]) {
          newRowValues["DO Time"] = timeAdd(row["Appt Time"], -apptPadding)
          newRowValues["PU Time"] = timeAdd(newRowValues["DO Time"], -journeyTime)
        }
      }
      newValues.push(newRowValues)
    })
    setValuesByHeaderNames(newValues, tripRows)
  } catch(e) { logError(e) }
}

/**
 * Scans the entire column of the edited cell for duplicate values and sets a
 * cell note listing the row numbers of any duplicates found. Clears the note
 * if no duplicates exist.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The edited cell.
 */
function scanForDuplicatesOnEdit(range) {
  try {
    const thisValue = range.getValue()
    const thisRowNumber = range.getRow()
    const fullRange = range.getSheet().getRange(1, range.getColumn(), range.getSheet().getLastRow())
    const values = fullRange.getValues().flat()

    let duplicateRows = []
    values.forEach((value, i) => {
      if (value == thisValue && (i + 1) != thisRowNumber) duplicateRows.push(i + 1)
    })
    if (duplicateRows.length == 1) range.setNote("This value is already used in row "  + duplicateRows[0])
    if (duplicateRows.length > 1)  range.setNote("This value is already used in rows " + duplicateRows.join(", "))
    if (duplicateRows.length == 0) range.clearNote()
  } catch(e) { logError(e) }
}

/**
 * When a full-width row paste is detected (the edited range spans all columns
 * from column 1), replaces any existing `"Trip ID"` values with new UUIDs to
 * prevent duplicate trip IDs from being introduced via paste.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
// When a trip is pasted in, change the Trip ID to avoid duplicate IDs
function updateTripID(e) {
  try {
    if (e.range.getColumn() === 1 &&
        e.range.getLastColumn() === e.range.getSheet().getMaxColumns()) {
      let tripValues = getRangeValuesAsTable(e.range)
      tripValues.forEach(row => {
        if (row["Trip ID"]) { row["Trip ID"] = Utilities.getUuid() }
      })
      setValuesByHeaderNames(tripValues, e.range)
    }
  } catch(e) { logError(e) }
}

/**
 * Handles the Action/Go checkbox trigger on trip rows. Reads the action text
 * from the cell to the left of the checkbox and dispatches to `createReturnTrip()`
 * or `addStop()`. Clears both the checkbox and action cell after dispatching,
 * or if no valid action is found.
 * @param {GoogleAppsScript.Spreadsheet.Range} goCheckBoxRange - The range of the
 *   Go checkbox cell that was checked.
 */
function tripActionButton(goCheckBoxRange) {
  try {
    const goCheckboxValue = goCheckBoxRange.getValue()
    const actionCell = goCheckBoxRange.getSheet().getRange(goCheckBoxRange.getRow(), goCheckBoxRange.getColumn()-1)
    const actionText = actionCell.getValue()
    if (goCheckboxValue && actionText) {
      if (actionText == "Add return trip") {
        goCheckBoxRange.setValue(null)
        actionCell.setValue(null)
        createReturnTrip()
      } else if (actionText == "Add stop") {
        goCheckBoxRange.setValue(null)
        actionCell.setValue(null)
        addStop()
      } else {
        goCheckBoxRange.setValue(null)
        actionCell.setValue(null)
      }
    } else {
      goCheckBoxRange.setValue(null)
      actionCell.setValue(null)
    }
  } catch(e) { logError(e) }
}

/**
 * Thin wrapper that calls `updateProperties()` for a sheet-trigger invocation
 * on the "Document Properties" sheet.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
function updatePropertiesOnEdit(e) {
  try {
    updateProperties(e)
  } catch(e) { logError(e) }
}

/**
 * Final sheet trigger for the Trips sheet. Runs after cell triggers.
 * Calls `updateTripID()` (paste detection), `completeTripRunValues()` (run
 * association auto-fill), and `clearSpillBlockages()` (array formula upkeep).
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
function tripSheetTrigger(e) {
  try {
    updateTripID(e)
    completeTripRunValues(e)
    clearSpillBlockages(e)
  } catch(e) { logError(e) }
}

/**
 * Final sheet trigger for the Runs sheet. Runs after cell triggers.
 * Calls `clearSpillBlockages()` to repair any array-formula spill blockages.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
function runSheetTrigger(e) {
  try {
    clearSpillBlockages(e)
  } catch(e) { logError(e) }
}
