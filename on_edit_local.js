// Any on_edit actions that are local to a specific RideSheet instance would
// be put here.
// cell-based triggers should be prefixed with "localCode"

const initialLocalSheetTriggers = {}

const finalLocalSheetTriggers  = {}

const rangeTriggersLocal = {
  localCodeExpandAddress: {
    functionCall: localExpandAddress,
    callOncePerRow: true
  }
}

function callLocalSheetTriggers(e, sheetName, triggers) {
  if (Object.keys(triggers).indexOf(sheetName) !== -1) {
    triggers[sheetName](e)
  }
}

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

function localExpandAddress(sourceRange) {
  const shortName = sourceRange.getValue()
  if (shortName?.toString().trim()) {
    try {
      const targetSheet = sourceRange.getSheet()
      const targetRange = targetSheet.getRange(sourceRange.getRow(), sourceRange.getColumn(), 1, 2)
      const result = getAddressByShortName(shortName)
      if (result) {
        targetRange.setValues([["",result]]).setNotes([["",""]]).setBackground(null)
        fillHoursAndMilesOnEdit(sourceRange)
        return true
      }
      return false
    } catch(e) {
      logError(e)
      return false
    }
  }
}