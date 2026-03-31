function createManifestsByRunForDate() {
  try {
    const templateDocId = getDocProp("driverManifestTemplateDocId")
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeSheet = ss.getActiveSheet()
    const ui = safeGetUi()
    let defaultDate
    let date
    let runDate
    if (activeSheet.getName() == "Trips") {
      runDate = getValueByHeaderName("Trip Date", getFullRows(activeSheet.getActiveCell()))
    } else if (activeSheet.getName() == "Runs") {
      runDate = getValueByHeaderName("Run Date", getFullRows(activeSheet.getActiveCell()))
    } else {
      // tomorrow, at midnight
      runDate = dateOnly(dateAdd(new Date(), 1))
    }

    if (isValidDate(runDate)) {
      defaultDate = runDate
    } else {
      // tomorrow, at midnight
      defaultDate = dateOnly(dateAdd(new Date(), 1))
    }

    if (ui) {
      let promptResult = ui.prompt("Create Manifests",
          "Enter date for manifests. Leave blank for " + formatDate(defaultDate, null, null),
          ui.ButtonSet.OK_CANCEL)
      startTime = new Date()
      if (promptResult.getSelectedButton() !== ui.Button.OK) {
        ss.toast("Action cancelled as requested.")
        return
      }
      if (promptResult.getResponseText() == "") {
        date = defaultDate
      } else {
        date = parseDate(promptResult.getResponseText(),"Invalid Date")
      }
      if (!isValidDate(date)) {
        ss.toast("Invalid date, action cancelled.")
        return
      }
    } else {
      date = defaultDate
    }

    const dateFilter = createDateFilterForManifestData(date)
    const manifestData = getManifestData(dateFilter)
    const groupedManifestData = groupManifestDataByRun(manifestData)
    const manifestCount = createManifests(templateDocId, groupedManifestData, getManifestFileNameByRun)
    ss.toast(manifestCount + " created.","Manifest creation complete.")
  } catch(e) { logError(e) }
}

function createSelectedManifestsByRun() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeSheet = ss.getActiveSheet()
    const rangeList = activeSheet.getActiveRangeList().getRanges()
    const templateDocId = getDocProp("driverManifestTemplateDocId")
    let selectedRows = []
    rangeList.forEach(range => {selectedRows.push(...getRangeValuesAsTable(getFullRows(range)))})

    if (!selectedRows.length ||
        !Object.hasOwn(selectedRows[0], "Trip Date") ||
        !Object.hasOwn(selectedRows[0], "Driver ID") ||
        !Object.hasOwn(selectedRows[0], "Vehicle ID")) {
      ss.toast("No trips selected, no manifests created.")
      return
    }

    let runList = []
    selectedRows.forEach(row => {
      let thisRun = {}
      thisRun["Trip Date"] = row["Trip Date"]
      thisRun["Driver ID"] = row["Driver ID"]
      thisRun["Vehicle ID"] = row["Vehicle ID"]
      runList.push(thisRun)
    })

    const runFilter = createRunFilterForManifestData(runList)
    const manifestData = getManifestData(runFilter, activeSheet.getName())
    const groupedManifestData = groupManifestDataByRun(manifestData)
    const manifestCount = createManifests(templateDocId, groupedManifestData, getManifestFileNameByRun)
    ss.toast(manifestCount + " created.","Manifest creation complete.")
  } catch(e) { logError(e) }
}

function createManifests(templateDocId, groupedManifestData, fileNameFunction) {
  try {
    const manifestFolderId = getDocProp("driverManifestFolderId")
    const templateDoc = DocumentApp.openById(templateDocId)
    prepareTemplate(templateDocId)

    let manifestCount = 0
    groupedManifestData.forEach(manifestGroup => {
      const manifestFileName = fileNameFunction(manifestGroup)
      const manifestDocId = createManifest(manifestGroup, templateDoc, manifestFileName, manifestFolderId)
      if (getDocProp("createManifestPdf")) {
        createPdfFromDocFile(manifestDocId, manifestFileName, manifestFolderId)
      }
      if (!getDocProp("keepManifestDoc")) {
        Drive.Files.update({ trashed: true }, manifestDocId, null, { supportsAllDrives: true })
      }
      manifestCount++
    })
    return manifestCount
  } catch(e) {
    logError(e)
  }
}

function createManifest(manifestGroup, templateDoc, manifestFileName, folderId) {
  const templateBody    = templateDoc.getBody()
  const tempText        = "{{Temporary Text}}"
  const manifestDocId   = createDoc(manifestFileName, folderId, tempText, "text/plain")
  const manifestDoc     = DocumentApp.openById(manifestDocId)
  const manifestBody    = manifestDoc.getBody()

  // Update page settings
  manifestBody.setMarginTop(templateBody.getMarginTop())
  manifestBody.setMarginRight(templateBody.getMarginRight())
  manifestBody.setMarginBottom(templateBody.getMarginBottom())
  manifestBody.setMarginLeft(templateBody.getMarginLeft())
  manifestBody.setPageHeight(templateBody.getPageHeight())
  manifestBody.setMarginLeft(templateBody.getMarginLeft())

  // Update page header and page footer
  const templateHeader = templateDoc.getHeader()
  if (templateHeader) {
    manifestHeader = manifestDoc.addHeader()
    for (let i = 0, c = templateHeader.getNumChildren(); i < c; i++) {
      appendElement(manifestHeader, templateHeader.getChild(i).copy())
    }
    replaceElementText(manifestHeader, manifestGroup["Events"][0])
  }
  const templateFooter = templateDoc.getFooter()
  if (templateFooter) {
    manifestFooter = manifestDoc.addFooter()
    for (let i = 0, c = templateFooter.getNumChildren(); i < c; i++) {
      appendElement(manifestFooter, templateFooter.getChild(i).copy())
    }
    replaceElementText(manifestFooter, manifestGroup["Events"][manifestGroup["Events"].length - 1])
  }

  // Add the document header elements
  appendTemplateRange(templateDoc.getNamedRanges("HEADER")[0]?.getRange(), manifestBody, manifestGroup["Events"][0])

  // Add all the PU and DO elements. Use the section name of each event to decide whether to add a PU or DO range.
  manifestGroup["Events"].forEach((event, i) => {
    appendTemplateRange(templateDoc.getNamedRanges(event["Section Name"])[0]?.getRange(), manifestBody, event)
  })

  // Add the footer elements
  appendTemplateRange(templateDoc.getNamedRanges("FOOTER")[0]?.getRange(), manifestBody, manifestGroup["Events"][manifestGroup["Events"].length - 1])

  // Remove the tempText needed to create the file
  manifestBody.removeChild(manifestBody.getChild(0))
  manifestDoc.saveAndClose()
  return manifestDocId
}

function createPdfFromDocFile(manifestDocId, manifestFileName, manifestFolderId) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + manifestDocId + '/export?mimeType=application/pdf'
  const options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  }
  const pdfFileName = manifestFileName + ".pdf"
  const pdfBlob = UrlFetchApp.fetch(url, options).getBlob().setName(pdfFileName)

  const createdPdfFile = Drive.Files.create(
    {
      name: pdfFileName,
      mimeType: 'application/pdf',
      parents: [manifestFolderId]
    },
    pdfBlob,
    {
      supportsAllDrives: true
    })
  return createdPdfFile.id
}

function createDoc(fileName, folderId, content, contentType) {
  try {
    const blob = Utilities.newBlob(content, contentType)
    const file = Drive.Files.create(
      {
        name: fileName,
        mimeType: 'application/vnd.google-apps.document',
        parents: [folderId]
      },
      blob,
      {
        supportsAllDrives: true
      }
    )
    return file.id
  } catch(e) {
    logError(e)
    // Re-throw to allow callers to handle it
    throw e
  }
}

function appendTemplateRange(range, docSection, data) {
  if (!range) return
  const rangeElements = range.getRangeElements()
  rangeElements.forEach(rangeElement => {
    const templateElement = rangeElement.getElement()
    const newElement = templateElement.copy()
    if (data) {
      const tempText = replaceText(templateElement.getText(), data)
      // Append the element if it will ultimately have text or
      // if the element has no fields to populate (e.g., it's just a blank line)
      if (tempText.trim() || elementFieldCount(templateElement) === 0) {
        appendElement(docSection, newElement)
        replaceElementText(newElement, data)
      }
    } else {
      appendElement(docSection, newElement)
    }
  })
}

function replaceElementText(element, data) {
  let elementText = element.getText()

  // First pass: Process conditional fields {?field}...{field}
  const conditionalPattern = /\{\?([^}]+)\}(.*?)\{\1\}/g
  const conditionalMatches = [...elementText.matchAll(conditionalPattern)]

  conditionalMatches.forEach(match => {
    const fullMatch = match[0]
    const fieldName = match[1]
    if (Object.keys(data).indexOf(fieldName) != -1) {
      const hasValue = data[fieldName] !== null &&
                       data[fieldName] !== undefined &&
                       data[fieldName] !== ''
      if (hasValue) {
        // Field has value - remove just the conditional marker {?field}
        element.replaceText(escapeRegex('{?' + fieldName + '}'), '')
      } else {
        // Field is empty - remove the entire conditional block
        element.replaceText(escapeRegex(fullMatch), '')
      }
    }
  })

  // Second pass: Process regular fields {field}
  elementText = element.getText()
  const pattern = /{(.*?)}/g
  const innerMatches = [...elementText.matchAll(pattern)].map(match => match[1])
  let datum
  innerMatches.forEach(fieldName => {
    if (isValidDate(data[fieldName])) {
      if (fieldName.match(/\bdate\b/i)) {
        datum = formatDate(data[fieldName])
      } else if (fieldName.match(/\btime\b/i)) {
        datum = formatDate(data[fieldName], null, "hh:mm aa")
      } else {
        datum = formatDate(data[fieldName], null, "hh:mm aa M/d/yy")
      }
    } else {
      datum = data[fieldName]
    }
    if (Object.keys(data).indexOf(fieldName) != -1) {
      element.replaceText("{" + fieldName + "}", datum)
      if (fieldName.match(/\baddress\b/i) && datum && getDocProp("addManifestAddressLinks")) {
        const url = createGoogleMapsDirectionsURL(datum)
        const text = element.asText()
        let addressRange = text.findText(escapeRegex(datum))
        if (addressRange) {
          do {
            text.setLinkUrl(addressRange.getStartOffset(), addressRange.getEndOffsetInclusive(), url)
            addressRange = text.findText(escapeRegex(datum), addressRange)
          } while (addressRange)
        }
      }
    }
  })
}

function elementFieldCount(element) {
  const elementText = element.getText()
  const pattern = /{(.*?)}/g
  const matches = [...elementText.matchAll(pattern)]
  return matches.length
}

function getManifestData(filterFunction, tripsSheetName = "Trips") {
  // Get all the raw data
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const drivers = getRangeValuesAsTable(ss.getSheetByName("Drivers").getDataRange())
  const vehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())
  const customers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
  const trips = getRangeValuesAsTable(ss.getSheetByName(tripsSheetName).getDataRange())
  let manifestTrips = trips.filter(filterFunction)

  // Pull in the lookup table data
  manifestTrips.forEach(tripRow => {
    mergeAttributes(tripRow, drivers,   "Driver ID"  )
    mergeAttributes(tripRow, vehicles,  "Vehicle ID" )
    mergeAttributes(tripRow, customers, "Customer ID")
    tripRow["Manifest Creation Time"] = new Date()
    tripRow["Manifest Creation Date"] = new Date()
  })

  // For events, create two rows for each trip -- one for PU, and one for DO
  let pickups = manifestTrips.map(tripRow => {
    let newRow = Object.assign({},tripRow)
    newRow["Section Name"] = "PICKUP"
    newRow["Event Name"]   = "pickup"
    newRow["Event Time"]   = tripRow["PU Time"]
    newRow["Sort Field"]   = formatDate(new Date(tripRow["PU Time"]), null, "HH:mm") + " PU"
    return newRow
  })
  let dropOffs = manifestTrips.map(tripRow => {
    let newRow = Object.assign({},tripRow)
    newRow["Section Name"] = "DROP OFF"
    newRow["Event Name"]   = "drop off"
    newRow["Event Time"]   = tripRow["DO Time"]
    newRow["Sort Field"]   = formatDate(new Date(tripRow["DO Time"]), null, "HH:mm") + " DO"
    return newRow
  })
  let manifestEvents = pickups.concat(dropOffs).sort((a,b) => {
    if (a["Sort Field"] < b["Sort Field"]) { return -1 }
    if (a["Sort Field"] > b["Sort Field"]) { return  1 }
    return 0
  })

  const result = {
    "trips": manifestTrips,
    "events": manifestEvents
  }
  return result
}

function groupManifestDataByRun(manifestData) {
  // Group the trips into runs -- A run is a collection of trips on the same day
  // with the same driver, vehicle, and run id
  let manifestGroups = []
  manifestData.trips.forEach(trip => {
    let runIndex = manifestGroups.findIndex(r =>
      r["Trip Date"].getTime() == trip["Trip Date"].getTime() &&
      r["Driver ID"] == trip["Driver ID"] &&
      r["Vehicle ID"] == trip["Vehicle ID"] &&
      r["Run ID"] == trip["Run ID"]
    )
    if (runIndex == -1) {
      let newRun = {}
      newRun["Driver ID"]    = trip["Driver ID"]
      newRun["Vehicle ID"]   = trip["Vehicle ID"]
      newRun["Run ID"]       = trip["Run ID"]
      newRun["Driver Name"]  = trip["Driver Name"]
      newRun["Driver Email"] = trip["Driver Email"]
      newRun["Vehicle Name"] = trip["Vehicle Name"]
      newRun["Trip Date"]    = trip["Trip Date"]
      newRun["Trips"]        = [trip]
      newRun["Events"]       = []
      manifestGroups.push(newRun)
    } else {
      manifestGroups[runIndex]["Trips"].push(trip)
    }
  })
  // Group the manifest events into the same runs
  manifestData.events.forEach(event => {
    let matchedRun = manifestGroups.find(run =>
      run["Trip Date"].getTime() == event["Trip Date"].getTime() &&
      run["Driver ID"] == event["Driver ID"] &&
      run["Vehicle ID"] == event["Vehicle ID"] &&
      run["Run ID"] == event["Run ID"]
    )
    matchedRun["Events"].push(event)
  })
  return manifestGroups
}

function createDriverManifest(manifestDate, driverId) {
  const driverManifestFolderId = getDocProp("driverManifestFolderId")
  const driverManifestTemplateDocId = getDocProp("driverManifestTemplateDocId")
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const filter = function(row) { return row["Driver ID"] == driverId }
  const driverRow = findFirstRowByHeaderNames(ss.getSheetByName("Drivers"), filter)
  const driverName = driverRow["Driver Name"]
  const manifestFileName = formatDate(manifestDate, null, "yyyy-MM-dd") + " Manifest for " + driverName
  const manifestFolder = DriveApp.getFolderById(driverManifestFolderId)
  const manifestFile = DriveApp.getFileById(driverManifestTemplateDocId).makeCopy(manifestFolder).setName(manifestFileName)
  const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  const manifestDoc = DocumentApp.openById(manifestFile.getId())
  deleteAllNamedRanges(templateDoc)
  const manifestSections = templateDoc.getBody().getParent()
}

// All these element types support the getText() and editAsText() methods
// Element types without text: COMMENT_SECTION, DOCUMENT, EQUATION_FUNCTION_ARGUMENT_SEPARATOR, EQUATION_SYMBOL, FOOTNOTE HORIZONTAL_RULE,
//                             INLINE_DRAWING, INLINE_IMAGE, PAGE_BREAK, UNSUPPORTED
function elementHasText(element) {
  elementTypesWithText = [DocumentApp.ElementType.BODY_SECTION,
                          DocumentApp.ElementType.EQUATION,
                          DocumentApp.ElementType.EQUATION_FUNCTION,
                          DocumentApp.ElementType.FOOTER_SECTION,
                          DocumentApp.ElementType.FOOTNOTE_SECTION,
                          DocumentApp.ElementType.HEADER_SECTION,
                          DocumentApp.ElementType.LIST_ITEM,
                          DocumentApp.ElementType.PARAGRAPH,
                          DocumentApp.ElementType.TABLE,
                          DocumentApp.ElementType.TABLE_CELL,
                          DocumentApp.ElementType.TABLE_OF_CONTENTS,
                          DocumentApp.ElementType.TABLE_ROW,
                          DocumentApp.ElementType.TEXT]
  return (elementTypesWithText.indexOf(element.getType()) > -1)
}

function deleteAllNamedRanges(doc) {
  doc.getNamedRanges().forEach(range => {
    range.remove()
  })
}

function prepareTemplate(driverManifestTemplateDocId) {
  lastUpdated = getFileLastUpdated(driverManifestTemplateDocId)

  if (lastUpdated > getDocProp("manifestTemplateLastUpdated_")) {
    const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
    deleteAllNamedRanges(templateDoc)
    const body = templateDoc.getBody()
    for (let i = 0, c = body.getNumChildren(); i < c; i++) {
      let element = body.getChild(i)
      if (elementHasText(element)) {
        const match = element.getText().match(/^\s*\[BEGIN (?<sectionName>.+?)\]\s*$/)
        if (match) {
          // We're at the beginning of a section. 
          // We want to:
          // - Jump to the next element
          // - Set up a named range with a name matching the section name
          // - Begin a loop where we:
          //   - Check each element to see if it's the closing element
          //   - If it's not the closing element, add that element to a named range
          //   - If it's the closing element, complete the building of the named range and and exit the loop
          const outerRangeBuilder = templateDoc.newRange()
          const innerRangeBuilder = templateDoc.newRange()
          outerRangeBuilder.addElement(element)
          if (i < c) { i++ }
          const sectionName = match.groups["sectionName"]
          const regex = new RegExp(`^\\s*\\[END ${sectionName}\\]\\s*$`)
          for (let stayInLoop = true; stayInLoop && i < c ; i++) {
            element = body.getChild(i)
            if (elementHasText(element) && element.getText().match(regex)) {
              outerRangeBuilder.addElement(element)
              templateDoc.addNamedRange(`OUTER_${sectionName}`, outerRangeBuilder.build())
              if (innerRangeBuilder.getRangeElements().length > 0) {
                templateDoc.addNamedRange(sectionName, innerRangeBuilder.build())
              }
              stayInLoop = false
            } else {
              outerRangeBuilder.addElement(element)
              innerRangeBuilder.addElement(element)
            }
          }
        }
      }
      // Always increment the main counter to proceed through the document.
      i++
    }
    templateDoc.saveAndClose()
    setDocProp("manifestTemplateLastUpdated_", getFileLastUpdated(driverManifestTemplateDocId))
  }
}

function copyNamedRanges(source, destination) {
  ss = SpreadsheetApp.getActiveSpreadsheet()
  const driverManifestFolderId = getDocProp("driverManifestFolderId")
  const driverManifestTemplateDocId = getDocProp("driverManifestTemplateDocId")
  const manifestFileName = formatDate(new Date(), null, "yyyy-MM-dd") + " Test Manifest"
  const manifestFolder = DriveApp.getFolderById(driverManifestFolderId)
  const manifestFile = DriveApp.getFileById(driverManifestTemplateDocId).makeCopy(manifestFolder).setName(manifestFileName)
  const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  const manifestDoc = DocumentApp.openById(manifestFile.getId())
  const manifestBody = manifestDoc.getBody()

  templateDoc.getNamedRanges().forEach(range => {
    //range = templateDoc.getNamedRanges()[0]
    range.getRange().getRangeElements().forEach(element => {
      //element = range.getRange().getRangeElements()[0]
      appendElement(manifestBody, element.getElement().copy())
    })
  })
}

function appendElement(body, element) {
  let type = element.getType()
  if (type == DocumentApp.ElementType.PARAGRAPH) {
    body.appendParagraph(element)
  } else if (type == DocumentApp.ElementType.TABLE) {
    body.appendTable(element)
  } else if (type == DocumentApp.ElementType.LIST_ITEM) {
    body.appendListItem(element)
  } else if (type == DocumentApp.ElementType.HORIZONTAL_RULE) {
    body.appendHorizontalRule(element)
  } else if (type == DocumentApp.ElementType.INLINE_IMAGE) {
    body.appendImage(element)
  } else if (type == DocumentApp.ElementType.PAGE_BREAK) {
    body.appendPageBreak(element)
  }
}

function mergeAttributes(primaryRow, secondaryTable, primaryKeyName, secondaryKeyName) {
  secondaryKeyName = secondaryKeyName || primaryKeyName
  if (primaryRow[primaryKeyName]) {
    let matchingSecondaryRow = secondaryTable.find(secondaryRow => primaryRow[primaryKeyName] == secondaryRow[secondaryKeyName])
    if (matchingSecondaryRow) {
      primaryRow["_" + primaryKeyName + "-attributes"] = matchingSecondaryRow
      Object.keys(matchingSecondaryRow).forEach(key => {
        if (!primaryRow.hasOwnProperty(key)) {
          primaryRow[key] = matchingSecondaryRow[key]
        }
      })
    }
  }
}

function getManifestFileNameByRun(manifestGroup) {
  const manifestFileName = `${formatDate(manifestGroup["Trip Date"], null, "yyyy-MM-dd")} manifest for ${manifestGroup["Driver Name"]} on ${manifestGroup["Vehicle Name"]}`
  return manifestFileName
}

function createDateFilterForManifestData(date) {
  return function(trip) {
    return new Date(trip["Trip Date"]).valueOf() === date.valueOf()
  }
}

function createRunFilterForManifestData(runs) {
  return function (trip) {
    return runs.some(run => {
      return trip["Trip Date"] instanceof Date &&
              trip["Trip Date"].getTime() === run["Trip Date"].getTime() &&
              trip["Driver ID"] === run["Driver ID"] &&
              trip["Vehicle ID"] === run["Vehicle ID"]
    })
  }
}
