const initialSheetTriggers = {
  "Document Properties": updatePropertiesOnEdit
}

const finalSheetTriggers = {
  "Trips": tripSheetTrigger
}

const rangeTriggers = {
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
  codeUpdateTripVehicle: {
    functionCall: updateTripVehicleOnEdit,
    callOncePerRow: true
  }
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {event} e The onEdit event.
 */
function onEdit(e) {
  try {  
    const startTime = new Date()
    const sheetName = e.range.getSheet().getName()
    callSheetTriggers(e, sheetName, initialSheetTriggers)
    callCellTriggers(e)
    callSheetTriggers(e, sheetName, finalSheetTriggers)
  } catch(e) {
    logError(e)
  } finally {
    if (debugLogging) log("onEdit duration:", new Date().getTime() - startTime.getTime())
  }
}

function callSheetTriggers(e, sheetName, triggers) {
  if (Object.keys(triggers).indexOf(sheetName) !== -1) {
    triggers[sheetName](e)
  }
}

function protectHeaders(e, sheetName) {
  if (sheetsWithHeaders.indexOf(sheetName) !== -1 && 
      e.range.getRow() === 1 &&
      e.range.getHeight() === 1 && 
      e.range.getWidth() === 1) {
    let ui = SpreadsheetApp.getUi()
    let response = ui.alert("Whoa There", 
                            "Did you really mean to change this column name? That can mess things up. If you really meant to, click 'Yes'", 
                            ui.ButtonSet.YES_NO_CANCEL)
    if (response != ui.Button.YES) {
      e.range.setValue(e.oldValue)
      return true
    }
  }
}

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
          callsToMake[triggerName].push(range)
        }
      })
    })
  
    //  const serializableCallsToMake = {}
    //  Object.keys(callsToMake).forEach(rangeTrigger => {
    //    serializableCallsToMake[rangeTrigger] = callsToMake[rangeTrigger].map(range => range.getA1Notation())
    //  })
    //  log(JSON.stringify(serializableCallsToMake))
    Object.keys(callsToMake).forEach(rangeTrigger => {
      callsToMake[rangeTrigger].forEach(range => {
        rangeTriggers[rangeTrigger]["functionCall"](range)
      })
    })
  } catch(e) {
    logError(e)
  }
}

function formatAddressOnEdit(range) {
  try {
    const app = SpreadsheetApp
    let backgroundColor = app.newColor()
    if (range.getValue() && range.getValue().trim()) {
      addressParts = parseAddress(range.getValue())
      let formattedAddress = getGeocode(addressParts.geocodeAddress, "formatted_address")
      if (addressParts.parenText) formattedAddress = formattedAddress + " " + addressParts.parenText
      if (formattedAddress.startsWith("Error")) {
        const msg = "Address " + formattedAddress
        range.setNote(msg)
        app.getActiveSpreadsheet().toast(msg)
        backgroundColor.setRgbColor(errorBackgroundColor)
      } else {
        range.setValue(formattedAddress)
        range.setNote("")
        backgroundColor.setRgbColor(defaultBackgroundColor)
      } 
    } else {
      range.setNote("")
      backgroundColor.setRgbColor(defaultBackgroundColor)
    }
    range.setBackgroundObject(backgroundColor.build())
  } catch(e) {
    logError(e)
  }
}

function fillTripCellsOnEdit(range) {
  try {
    if (range.getValue()) {
      const ss = SpreadsheetApp.getActiveSpreadsheet()
      const tripRow = getFullRow(range)
      const tripValues = getRangeValuesAsTable(tripRow)[0]
      const filter = function(row) { return row["Customer Name and ID"] === tripValues["Customer Name and ID"] }
      const customerRow = findFirstRowByHeaderNames(ss.getSheetByName("Customers"), filter)
      let valuesToChange = {}
      valuesToChange["Customer ID"] = customerRow["Customer ID"]
      if (tripValues["PU Address"] == '') { valuesToChange["PU Address"] = customerRow["Home Address"] }
      if (tripValues["DO Address"] == '') { valuesToChange["DO Address"] = customerRow["Default Destination"] }
      if (tripValues["Service ID"] == '') { valuesToChange["Service ID"] = customerRow["Default Service ID"] }
      if (tripValues["Trip ID"] == '')    { valuesToChange["Trip ID"]    = Utilities.getUuid() }
      setValuesByHeaderNames([valuesToChange], tripRow)
      if (valuesToChange["PU Address"] || valuesToChange["DO Address"]) { 
        //log("fillTripCellsOnEdit called")
        fillHoursAndMilesOnEdit(range) 
      }
    }
  } catch(e) {
    logError(e)
  }
}

function fillHoursAndMilesOnEdit(range) {
  try {
    const tripRow = getFullRow(range)
    const tripValues = getRangeValuesAsTable(tripRow)[0]
    if (tripValues["PU Address"] && tripValues["DO Address"]) {
      const PUAddress = parseAddress(tripValues["PU Address"]).geocodeAddress
      const DOAddress = parseAddress(tripValues["DO Address"]).geocodeAddress
      const tripEstimate = getTripEstimate(PUAddress, DOAddress, "milesAndDays")
      setValuesByHeaderNames([{"Est Hours": tripEstimate.days, "Est Miles": tripEstimate.miles}], tripRow)
      if (tripEstimate["days"]) {
        SpreadsheetApp.getActiveSpreadsheet().toast("Travel estimate saved")
      }
    } else {
      setValuesByHeaderNames([{"Est Hours": "", "Est Miles": ""}], tripRow)
    }
  } catch(e) {
    logError(e)
  }
}

/**
 * Manage setup of a new customer record. The goals here are to:
 * - Trim the customer name as needed 
 * _ Generate a customer ID when it's missing and there's a first and last name present
 * - Autofill the "Customer Name and ID" field when the first name, last name, and ID are present. 
 *   This will be the field used to identify the customer in trip records
 * - Keep track of the current highest customer ID in document properties, seeding data when needed
 */
function setCustomerKeyOnEdit(range) {
  try {
    const customerRow = getFullRow(range)
    const customerValues = getRangeValuesAsTable(customerRow)[0]
    let newValues = {}
    if (customerValues["Customer First Name"] && customerValues["Customer Last Name"]) {
      let lastCustomerID = getDocProp("lastCustomerID_")
      if (!Number.isFinite(lastCustomerID)) {
        const sheet = range.getSheet()
        const idColumnPosition = getSheetHeaderNames(sheet).indexOf("Customer ID") + 1
        const idRange = sheet.getRange(1, idColumnPosition, sheet.getLastRow())
        let maxID = getMaxValueInRange(idRange)
        lastCustomerID = Number.isFinite(maxID) ? maxID : 1
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
  } catch(e) {
    logError(e)
  }
}

function updateTripTimesOnEdit(range) {
  try {
    const tripRows = getFullRows(range)
    const tripValues = getRangeValuesAsTable(tripRows)
    let newValues = []
    tripValues.forEach(row => {
      let newRowValues = {}
      if (isFinite(row["Est Hours"])) {
        const estMilliseconds = timeOnlyAsMilliseconds(row["Est Hours"])
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
  } catch(e) {
    logError(e)
  }
}

function updateTripVehicleOnEdit(range) {
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const tripRow = getFullRow(range)
    const tripValues = getRangeValuesAsTable(tripRow)[0]
    if (tripValues["Driver ID"] && !tripValues["Vehicle ID"]) {
      const filter = function(row) { return row["Driver ID"] === tripValues["Driver ID"] && row["Default Vehicle ID"] }
      const driverRow = findFirstRowByHeaderNames(ss.getSheetByName("Drivers"), filter)
      if (driverRow) {
        let valuesToChange = {}
        valuesToChange["Vehicle ID"] = driverRow["Default Vehicle ID"]
        setValuesByHeaderNames([valuesToChange], tripRow)
      }
    }  
  } catch(e) {
    logError(e)
  }
}

function scanForDuplicatesOnEdit(range) {
  try {
    thisValue = range.getValue()
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
  } catch(e) {
    logError(e)
  }
}

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
  } catch(e) {
    logError(e)
  }
}

function updatePropertiesOnEdit(e) {
  try {
    updateProperties(e)
  } catch(e) {
    logError(e)
  }
}

function tripSheetTrigger(e) {
  try {
    updateTripID(e)
    updateRuns(e)
  } catch(e) {
    logError(e)
  }
}