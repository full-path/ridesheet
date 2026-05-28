/**
 * @fileoverview Local overrides and extensions for RideSheet's onEdit triggers.
 *
 * This file is the intended customization point for org-specific forks.
 * Add local sheet triggers and cell triggers here without touching the core
 * `on_edit.js` file.
 *
 * **Sheet triggers**: add entries to `initialLocalSheetTriggers` or
 * `finalLocalSheetTriggers` mapping sheet name → handler function. These
 * are dispatched by `callLocalSheetTriggers()` before and after the core
 * cell triggers respectively.
 *
 * **Cell triggers**: add entries to `rangeTriggersLocal` keyed by named range
 * name (must start with `"localCode"`). Each entry must have:
 * - `functionCall` {function} — called with the edited `Range`.
 * - `callOncePerRow` {boolean} — deduplicate calls within the same row.
 *
 * `callLocalCellTriggers()` mirrors the logic of `callCellTriggers()` in
 * `on_edit.js` but looks for `"localCode"`-prefixed named ranges and
 * dispatches from `rangeTriggersLocal`.
 */

// Any on_edit actions that are local to a specific RideSheet instance would
// be put here.
// cell-based triggers should be prefixed with "localCode"

/**
 * Local sheet triggers called at the start of `onEdit`, before cell triggers.
 * Maps sheet name to a handler function receiving the onEdit event `e`.
 * @type {Object.<string, function(GoogleAppsScript.Events.SheetsOnEdit): void>}
 */
const initialLocalSheetTriggers = {}

/**
 * Local sheet triggers called at the end of `onEdit`, after cell triggers.
 * Maps sheet name to a handler function receiving the onEdit event `e`.
 * @type {Object.<string, function(GoogleAppsScript.Events.SheetsOnEdit): void>}
 */
const finalLocalSheetTriggers  = {}

/**
 * Local cell-level triggers keyed by named range name (prefix `"localCode"`).
 * Each entry maps to a handler function and a `callOncePerRow` flag.
 * @type {Object.<string, {functionCall: function(GoogleAppsScript.Spreadsheet.Range): void, callOncePerRow: boolean}>}
 */
const rangeTriggersLocal = {}

/**
 * Dispatches a sheet-level trigger from a local trigger map.
 * Mirrors `callSheetTriggers()` in `on_edit.js` but operates on local triggers.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 * @param {string} sheetName - The name of the edited sheet.
 * @param {Object.<string, function>} triggers - A local sheet-trigger map.
 */
function callLocalSheetTriggers(e, sheetName, triggers) {
  if (Object.keys(triggers).indexOf(sheetName) !== -1) {
    triggers[sheetName](e)
  }
}

/**
 * Evaluates all `localCode`-prefixed named ranges that overlap the edited
 * cell(s) and calls the corresponding handler functions from
 * `rangeTriggersLocal`. Returns immediately if `rangeTriggersLocal` is empty.
 * Otherwise mirrors the logic of `callCellTriggers()` in `on_edit.js`.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
function callLocalCellTriggers(e) {
  try {
    if (!Object.keys(rangeTriggersLocal).length) return
    const spreadsheet = e.source
    const sheet = e.range.getSheet()
    const allNamedRanges = sheet.getNamedRanges().filter(namedRange =>
      namedRange.getName().indexOf("localCode") === 0 && rangesOverlap(e.range, namedRange.getRange())
    )
    if (allNamedRanges.length === 0) return

    const isMultiColumnRange = (e.range.getWidth() > 1)
    const isMultiRowRange = (e.range.getHeight() > 1)
    let triggeredRows = {}
    let ranges = []
    let callsToMake = {}
    Object.keys(rangeTriggersLocal).forEach(rangeTrigger => callsToMake[rangeTrigger] = [])

    // Set up the tracking to prevent running some code from running multiple times per row.
    Object.keys(rangeTriggersLocal).forEach(key => {
      if (rangeTriggersLocal[key].callOncePerRow) triggeredRows[key] = []
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
          callsToMake[triggerName].push(range)
        }
      })
    })

    Object.keys(callsToMake).forEach(rangeTrigger => {
      callsToMake[rangeTrigger].forEach(range => {
        rangeTriggersLocal[rangeTrigger]["functionCall"](range)
      })
    })
  } catch(e) { logError(e) }
}
