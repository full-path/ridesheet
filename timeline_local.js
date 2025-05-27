/**
 * Checks whether the data in the "Trips" and "Runs" sheets has changed since the last check.
 * If any changes are detected, rebuilds the Timeline report. Primarily intended to be called
 * from an onEdit or onOpen trigger, or manually if desired.
 */
function refreshTimelineIfChanged(e) {
  //const startTime = new Date()
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sourceTripRange = ss.getSheetByName("Trips").getDataRange()
  const sourceDispatchRange = ss.getSheetByName("Dispatch").getDataRange()
  const sourceRunRange = ss.getSheetByName("Runs").getDataRange()
  const tripFilter = function(row) {
    return row["Trip Date"] instanceof Date &&
      row["Driver ID"] &&
      row["Vehicle ID"] && (
        !row["Trip Result"] || row["Trip Result"] == "Completed"
      )
  }
  const runFilter = function(row) {
    return row["Run Date"] instanceof Date && row["Driver ID"] && row["Vehicle ID"]
  }
  const tripColsToCheck = [
    "Trip Date",
    "Driver ID",
    "Vehicle ID",
    "Run ID",
    "Customer Name and ID",
    "PU Time",
    "DO Time",
    "Trip Result"
  ]
  const runColsToCheck = [
    "Run Date",
    "Driver ID",
    "Vehicle ID",
    "Run ID",
    "Scheduled Start Time",
    "Scheduled End Time"
  ]
  
  const haveTripsChanged = hasRangeChanged(sourceTripRange,"tripsDigest", tripColsToCheck, tripFilter) 
  const haveDispatchTripsChanged = hasRangeChanged(sourceDispatchRange,"dispatchDigest", tripColsToCheck, tripFilter)
  const haveRunsChanged = hasRangeChanged(sourceRunRange,"RunsDigest", runColsToCheck, runFilter)

  if (haveTripsChanged || haveRunsChanged || haveDispatchTripsChanged) {
    buildTimelineReport(sourceTripRange, sourceRunRange, sourceDispatchRange)
  }
  //Logger.log(["refreshTimelineIfChanged duration:",(new Date()) - startTime])
  //log("refreshTimelineIfChanged duration:",(new Date()) - startTime)
}

/**
 * Checks if the edited cell is the "timelineManualRefresh" named range. If so,
 * it rebuilds the Timeline report and then unchecks the associated checkbox.
 * This function is designed to be triggered by an onEdit() event.
 */
function timelineManualRefresh(e) {
  try {
    ss = e.range.getSheet().getParent()
    const manualRefreshCheckboxRange = ss.getRangeByName("timelineManualRefresh")
    if (e.range.getSheet().getName() == manualRefreshCheckboxRange.getSheet().getName() &&
        e.range.getA1Notation() == manualRefreshCheckboxRange.getA1Notation()) {
      buildTimelineReport()
      e.range.setValue(false)
      ss.toast("Manual refresh of timeline complete","Success")

    }
  } catch(e) { logError(e) }
}

/**
 * Builds or refreshes a "Timeline" report in the "Timeline" sheet of the active spreadsheet.
 * This function processes data from the "Trips" and "Runs" sheets (or from the provided ranges)
 * to generate a time-slot-based overview. The function aggregates pick-ups, drop-offs, 
 * and on-board counts, then writes the resulting data and cell notes to the "Timeline" sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} [sourceTripRange] 
 *        Optional. A range object representing the Trips data. If not provided, the function 
 *        defaults to using all data in the "Trips" sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} [sourceRunRange]
 *        Optional. A range object representing the Runs data. If not provided, the function 
 *        defaults to using all data in the "Runs" sheet.
 *
 * @returns {void}
 *
 * @throws {Error}
 *         Throws an error if anything goes wrong while fetching or processing the data.
 */
function buildTimelineReport(sourceTripRange, sourceRunRange, sourceDispatchRange) {
  const startTime = new Date()  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    if (!sourceTripRange) sourceTripRange = ss.getSheetByName("Trips").getDataRange()
    if (!sourceRunRange) sourceRunRange = ss.getSheetByName("Runs").getDataRange()
    if (!sourceDispatchRange) sourceDispatchRange = ss.getSheetByName("Dispatch").getDataRange()
    const mainSheet = ss.getSheetByName("Timeline")
    const sourceTripData = getRangeValuesAsTable(sourceTripRange).
            concat(getRangeValuesAsTable(sourceDispatchRange))
    const sourceRunData = getRangeValuesAsTable(sourceRunRange)
    const timeIncrement = 15
    const timeZone = getDocProp("localTimeZone")
    const startRow = 3 // This should be the header row

    ss.getRangeByName("timelineLastRefreshed")?.setValue(new Date())
    data = {}
    sourceRunData.filter(row => {
      return row["Run Date"] instanceof Date && row["Driver ID"] && row["Vehicle ID"]
    }).forEach(row => {
      const runKey = getTimelineRunKey(row["Run Date"], row, timeZone)
      data[runKey] = {}
      data[runKey]["run"] = row
      data[runKey]["trips"] = []
    })

    sourceTripData.filter(row => {
      return row["Trip Date"] instanceof Date &&
      row["Driver ID"] &&
      row["Vehicle ID"] && (
        !row["Trip Result"] || row["Trip Result"] == "Completed"
      )
    }).forEach(row => {
      const runKey = getTimelineRunKey(row["Trip Date"], row, timeZone)
      if (Object.hasOwn(data,runKey)) {
        data[runKey]["trips"].push(row)
      }
    })

    const runKeys = Object.keys(data).sort()
    runKeys.forEach(runKey => {
      data[runKey]["run-start-slot"] = roundDownMinutes(data[runKey]["run"]["Scheduled Start Time"],timeIncrement)
      data[runKey]["after-run-end-slot"] = roundUpMinutes(data[runKey]["run"]["Scheduled End Time"],timeIncrement)
      data[runKey]["first-pu-slot"] = roundDownMinutes(data[runKey]["trips"].map(row => row["PU Time"]).sort().at(0), timeIncrement)
      data[runKey]["after-last-do-slot"] = roundUpMinutes(data[runKey]["trips"].map(row => row["DO Time"]).sort().at(-1), timeIncrement)
      data[runKey]["tl-start-slot"] = getEarlier(data[runKey]["run-start-slot"], data[runKey]["first-pu-slot"])
      data[runKey]["tl-after-end-slot"] = getLater(data[runKey]["after-run-end-slot"],data[runKey]["after-last-do-slot"])

      //Logger.log([runKey,data[runKey]["run-start-slot"].getTime(), data[runKey]["first-pu-slot"].getTime(),data[runKey]["tl-start-slot"].getTime()])
      data[runKey]["slots"] = {}
      const slots = data[runKey]["slots"]
      for(let i = data[runKey]["tl-start-slot"]; i < data[runKey]["tl-after-end-slot"]; i.setMinutes(i.getMinutes() + timeIncrement)) {
        const slotKey = serializeTime(i, timeZone)
        slots[slotKey] = {}
        const slot = slots[slotKey]
        slot["time-start"] = new Date(i)
        const afterTimeEnd = new Date(i).setMinutes(i.getMinutes() + timeIncrement)
        slot["after-time-end"] = new Date(afterTimeEnd)
        slot["pickup-count"] = 0
        slot["dropoff-count"] = 0
        slot["onboard-count"] = 0
        slot["pickup-descriptions"] = []
        slot["dropoff-descriptions"] = [] 
        slot["onboard-descriptions"] = []
      }

      const slotKeys = Object.keys(slots).sort()
      data[runKey]["trips"].forEach(trip => {
        slotKeys.forEach(slotKey => {
          const slot = slots[slotKey]
          const guestCount = Number(trip["Guests"] ?? 0)
          const guestText = guestCount > 0 ? ` and ${guestCount} guest${guestCount == 1 ? "" : "s"} ` : ""
          const lastMinuteOfSlot = new Date(slot["after-time-end"].getTime() - 60000)
          let pickedUpOrDroppedOff = false
          if (trip["PU Time"] >= slot["time-start"] && trip["PU Time"] < slot["after-time-end"]) {
            slot["pickup-count"] = slot["pickup-count"] + guestCount + 1
            slot["pickup-descriptions"].push(`${serializeTime(trip["PU Time"], timeZone)} ⬆ ${trip["Customer Name and ID"]}${guestText}`)
            pickedUpOrDroppedOff = true
          }
          if (trip["DO Time"] >= slot["time-start"] && trip["DO Time"] < slot["after-time-end"]) {
            slot["dropoff-count"] = slot["dropoff-count"] + guestCount + 1
            slot["dropoff-descriptions"].push(`${serializeTime(trip["DO Time"], timeZone)} ⬇ ${trip["Customer Name and ID"]}${guestText}`)
            pickedUpOrDroppedOff = true
          }
          if (trip["PU Time"] <= slot["time-start"] && trip["DO Time"] >= lastMinuteOfSlot) {
            slot["onboard-count"] = slot["onboard-count"] + guestCount + 1
            if (!pickedUpOrDroppedOff) {
              slot["onboard-descriptions"].push(`${trip["Customer Name and ID"]}${guestText} on board`)
            }
          }
        })
      })
    })
    
    const allSlotKeys = [...new Set(runKeys.map(key => Object.keys(data[key]["slots"])).flat())].sort()
    let displayArray = [["Date","Driver","Vehicle",...allSlotKeys]]
    let noteArray = [["","","",...allSlotKeys.map(key => "")]]
    let previousDate = data[runKeys[0]]["run"]["Run Date"].getTime()
    runKeys.forEach(runKey => {
      const run = data[runKey]["run"]
      let thisDisplayRow = []
      let thisNoteRow = []
      if (run["Run Date"].getTime() != previousDate) {
        // Add a blank row between date groups
        displayArray.push([...Array(displayArray[0].length).fill("")])
        noteArray.push([...Array(noteArray[0].length).fill("")])
      }
      thisDisplayRow.push(run["Run Date"], run["Driver ID"], run["Vehicle ID"])
      thisNoteRow.push("", "", "")
      allSlotKeys.forEach(slotKey => {
        let slot = data[runKey]["slots"][slotKey] || {}
        thisDisplayRow.push(getFinalSlotValue(slot))
        thisNoteRow.push(getFinalSlotDescriptionValue(slot))
      })
      displayArray.push(thisDisplayRow)
      noteArray.push(thisNoteRow)
      data[runKey]["display-array"] = displayArray
      previousDate = data[runKey]["run"]["Run Date"].getTime()
    })

    const rowsToClear = mainSheet.getLastRow() - startRow + 1
    const colsToClear = mainSheet.getLastColumn()
    if (rowsToClear > 0) {
      const noteResetArray = Array.from({ length: rowsToClear }, () => Array(colsToClear).fill(""))
      mainSheet.getRange(startRow,1,rowsToClear,colsToClear).clearContent()
      mainSheet.getRange(startRow,1,rowsToClear,colsToClear).setNotes(noteResetArray)
    }
    const rangeToUpdate = mainSheet.getRange(startRow,1,displayArray.length,displayArray[0].length)
    rangeToUpdate.setValues(displayArray)
    rangeToUpdate.setNotes(noteArray)
  } catch(e) { logError(e) }
  //log("buildTimeline duration:",(new Date()) - startTime)
  //Logger.log(["buildTimeline duration:",(new Date()) - startTime])
}

function getTimelineRunKey(date, row, timeZone) {
  if (row["Run ID"] == "") {
    return `${Utilities.formatDate(date,timeZone,"yyyy-MM-dd")} ${row["Driver ID"]}:${row["Vehicle ID"]}`
  } else {
    return `${Utilities.formatDate(date,timeZone,"yyyy-MM-dd")} ${row["Driver ID"]}:${row["Vehicle ID"]}-${row["Run ID"]}`
  }
}

function serializeTime(time, timeZone) {
  return Utilities.formatDate(time,timeZone,"HH:mm") 
}

function roundDownMinutes(thisDate, increment) {
  let roundedDate = new Date(thisDate)
  const minutes = roundedDate.getMinutes()
  roundedDate.setMinutes(minutes - (minutes % increment), 0, 0) 
  return roundedDate
}

function roundUpMinutes(date, increment) {
  let roundedDate = new Date(date)
  const minutes = roundedDate.getMinutes()
  if (minutes % increment > 0) {
    roundedDate.setMinutes(minutes - (minutes % increment) + increment, 0, 0) 
  }
  return roundedDate
}

function getEarlier(date1, date2) {
  if (isNaN(date1)) {
    return new Date(date2)
  } else if (isNaN(date2)) {
    return new Date(date1)
  }
  if (date1 < date2) {
    return new Date(date1)
  } else {
    return new Date(date2)
  }
}

function getLater(date1, date2) {
  if (isNaN(date1)) {
    return new Date(date2)
  } else if (isNaN(date2)) {
    return new Date(date1)
  }
  if (date1 > date2) {
    return new Date(date1)
  } else {
    return new Date(date2)
  }
}

function getOneMinuteEarlier(date) {
  return new Date(date.getTime() - 60000)
}

function getFinalSlotValue(slot) {
  if (slot["onboard-count"]) { 
    return slot["onboard-count"]
  } else if (slot["pickup-count"] || slot["dropoff-count"]){
    return -(slot["pickup-count"] + slot["dropoff-count"])
  } else if (slot["onboard-count"] == 0) {
    return 0
  } else {
    return null
  }
}

function getFinalSlotDescriptionValue(slot) {
  const onboard = slot["onboard-descriptions"] || []
  const pickup = slot["pickup-descriptions"] || []
  const dropoff = slot["dropoff-descriptions"] || []
  const allDescriptions = [...onboard,...pickup,...dropoff].sort().join("\n")
  return allDescriptions
}

/**
 * Checks if a specified range of data has changed since the last time the function was called,
 * by comparing an MD5 digest of filtered/selected columns.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Range} range 
 *        The range to evaluate. Must have at least one row (header row) and at least one column.
 * @param {string} digestName 
 *        A unique key identifying the digest in the cache. Used to retrieve and store the last known digest.
 * @param {string[]} [columnsToCheck] 
 *        An array of column headers specifying which columns should be included in the comparison.
 *        If omitted or empty, all columns are included.
 * @param {Function} [filterFunction] 
 *        A function that receives a row object (header-value pairs) and returns a boolean.
 *        Only rows for which the function returns `true` will be considered. If omitted, all rows are included.
 *
 * @returns {boolean}
 *        Returns `true` if the range data has changed since the last check or if it is the first time checking.
 *        Returns `false` if no changes were detected.
 *
 * @throws {Error}
 *        Throws an error if any column name in `columnsToCheck` does not exist in the header row.
 */
function hasRangeChanged(range, digestName, columnsToCheck, filterFunction) {
  try {
    const cacheExpiration = 21600 // 6 hours, the maximum
    const currentValues = range.getValues()
    const header = currentValues[0]
    const dataRows = currentValues.slice(1)

    // columnsToCheck and filterFunction can be left empty. Handle those cases.
    let columnIndices = []
    if (columnsToCheck) {
      columnIndices = columnsToCheck.map(function(colName) {
        const idx = header.indexOf(colName)
        if (idx === -1) {
          throw new Error(`Column '${colName}' not found in header.`)
        }
        return idx
      })
    } else {
      columnIndices = header.map((_, i) => i)
    }
    if (!filterFunction) { filterFunction = function() { return true } }

    // Filter the rows based on the provided filter function.
    const filteredRows = dataRows.filter(function(row) {
      // Create an object with column names as keys and row values as values.
      let rowData = {}
      header.forEach(function(header, index) {
        rowData[header] = row[index]
      })
      // Apply the filter function to the row data.
      return filterFunction(rowData)
    })

    // Filter the data to include only the specified columns.
    let filteredValues = filteredRows.map(function(row) {
      return columnIndices.map(function(colIndex) {
        return row[colIndex]
      })
    })

    // Sort the rows in the filtered values array by all values in the row.
    filteredValues.sort(function(a, b) {
      for (let i = 0; i < a.length; i++) {
        if (a[i] > b[i]) return 1
        if (a[i] < b[i]) return -1
      }
      return 0
    })

    // Calculate the MD5 digest of the sorted and filtered values.
    const currentDigestBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(filteredValues))
    const currentDigest = Utilities.base64EncodeWebSafe(currentDigestBytes)
    //Logger.log([digestName,"currentDigest",currentDigest])

    // Get the stored digest from PropertiesService.
    const cache = CacheService.getDocumentCache()
    const storedDigest = cache.get(digestName)
    //Logger.log([digestName,"storedDigest",storedDigest])

    // If no stored digest exists, store the current digest and return true.
    if (!storedDigest) {
      cache.put(digestName, currentDigest, cacheExpiration)
      return true
    }

    // Compare the current digest with the stored digest.
    if (currentDigest !== storedDigest) {
      // If the digests are different, the range has changed.
      cache.put(digestName, currentDigest, cacheExpiration)
      return true
    }

    // If no changes were detected, return false.
    return false
  } catch(e) { logError(e) }
}
