const tempText = "{{Temporary Text}}"

function createManifestsByRunForDate() {
  try {
    const templateDocId = getDocProp("driverManifestTemplateDocId")
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeSheet = ss.getActiveSheet()
    const ui = SpreadsheetApp.getUi()
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

    let promptResult = ui.prompt("Create Manifests",
        "Enter date for manifests. Leave blank for " + formatDate(defaultDate, null, null),
        ui.ButtonSet.OK_CANCEL)
    startTime = new Date()
    if (promptResult.getResponseText() == "") {
      date = defaultDate
    } else {
      date = parseDate(promptResult.getResponseText(),"Invalid Date")
    }
    if (!isValidDate(date)) {
      ui.alert("Invalid date, action cancelled.")
      return
    } else if (promptResult.getSelectedButton() !== ui.Button.OK) {
      ss.toast("Action cancelled as requested.")
      return
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
    const activeSheetName = activeSheet.getName()
    if (activeSheetName != "Trips" && activeSheetName != "Runs") {
      ss.toast("Nothing selected, no manifests created.")
      return
    }
    const rangeList = activeSheet.getActiveRangeList().getRanges()
    const templateDocId = getDocProp("driverManifestTemplateDocId")
    let selectedRows = []
    rangeList.forEach(range => {selectedRows.push(...getRangeValuesAsTable(getFullRows(range)))})
    let runList = []
    selectedRows.forEach(row => {
      let thisRun = {}
      thisRun["Trip Date"] = (activeSheetName == "Trips" ? row["Trip Date"] : row["Run Date"])
      thisRun["Driver ID"] = row["Driver ID"]
      thisRun["Vehicle ID"] = row["Vehicle ID"]
      runList.push(thisRun)
    })

    const runFilter = createRunFilterForManifestData(runList)
    const manifestData = getManifestData(runFilter)
    const groupedManifestData = groupManifestDataByRun(manifestData)
    const manifestCount = createManifests(templateDocId, groupedManifestData, getManifestFileNameByRun)
    ss.toast(manifestCount + " created.","Manifest creation complete.")
  } catch(e) { logError(e) }
}

function createManifests(templateDocId, groupedManifestData, fileNameFunction) {
  try {
    const templateDoc = DocumentApp.openById(templateDocId)
    const folder = DriveApp.getFolderById(getDocProp("driverManifestFolderId"))
    prepareTemplate(templateDoc)

    let manifestCount = 0
    groupedManifestData.forEach(manifestGroup => {
      const manifestFileName = fileNameFunction(manifestGroup)
      const manifestDocFile = createManifest(manifestGroup, templateDoc, manifestFileName, folder)
      createPdfFromDocFile(manifestDocFile, folder)
      manifestCount++
    })
    return manifestCount
  } catch(e) {
    logError(e)
  }
}

function createManifest(manifestGroup, templateDoc, manifestFileName, folder) {
  const templateFileId = getDocProp("driverManifestTemplateDocId")
  const manifestFile   = DriveApp.getFileById(templateFileId).makeCopy(folder).setName(manifestFileName)
  const manifestDoc    = DocumentApp.openById(manifestFile.getId())

  emptyBody(manifestDoc)
  populateManifest(manifestDoc, templateDoc, manifestGroup)
  removeTempElement(manifestDoc)
  manifestDoc.saveAndClose()
  return manifestFile
}

function createPdfFromDocFile(manifestFile, folder) {
  const PdfFile = folder.createFile(manifestFile.getBlob().getAs("application/pdf"))
  PdfFile.setName(manifestFile + ".pdf")
  return PdfFile
}
                                                  
function populateManifest(manifestDoc, templateDoc, manifestGroup) {
  //templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  
  const manifestBody   = manifestDoc.getBody()
  const manifestParent = manifestBody.getParent()
  
  // Replace the fields in the document header and footer sections
  // There may be up to four, if there are different headers or footers for the first page
  for (let i = 0, c = manifestParent.getNumChildren(); i < c; i++) {
    let section = manifestParent.getChild(i)
    if (section.getType() === DocumentApp.ElementType.HEADER_SECTION) {
      replaceElementText(section, manifestGroup["Events"][0])
    } else if (section.getType() === DocumentApp.ElementType.FOOTER_SECTION) {
      replaceElementText(section, manifestGroup["Events"][manifestGroup["Events"].length - 1])
    }  
  }

  // Add the header elements
  replaceTextInRange(templateDoc.getNamedRanges("HEADER")[0].getRange(), manifestBody, manifestGroup["Events"][0])
  // Add all the PU and DO elements. Use the section name of each event to decide whether to add a PU or DO range.
  manifestGroup["Events"].forEach((event, i) => {
    replaceTextInRange(templateDoc.getNamedRanges(event["Section Name"])[0].getRange(), manifestBody, event)
  })
  // Add the footer elements
  replaceTextInRange(templateDoc.getNamedRanges("FOOTER")[0].getRange(), manifestBody, manifestGroup["Events"][manifestGroup["Events"].length - 1])
}

function replaceTextInRange(range, docSection, data) {
  let elements = range.getRangeElements()
  elements.forEach(element => {
    newElement = element.getElement().copy()
    replaceElementText(newElement, data)
    appendElement(docSection, newElement)
  })
}

function replaceElementText(element, data) {
  let text = element.getText()
  //text = "This is {a} test {with} words in {braces]"
  let pattern = /{(.*?)}/g
  let innerMatches = [...text.matchAll(pattern)].map(match => match[1])
  innerMatches.forEach(field => {
    if (isValidDate(data[field])) {
      if (field.match(/\bdate\b/i)) {
        datum = formatDate(data[field])
      } else if (field.match(/\btime\b/i)) {
        datum = formatDate(data[field], null, "hh:mm aa")
      } else {
        datum = formatDate(data[field], null, "hh:mm aa M/d/yy")
      }
    } else {
      datum = data[field]
    }
    if (Object.keys(data).indexOf(field) != -1) {
      element.replaceText("{" + field + "}", datum)
      if (field.match(/\baddress\b/i) && datum) {
        let url = createGoogleMapsDirectionsURL(datum)
        let text = element.asText()
        let addressRange = text.findText(escapeRegex(datum))
        if (addressRange) {
          do {
            text.setLinkUrl(addressRange.getStartOffset(), addressRange.getEndOffsetInclusive(), url)
            addressRange = text.findText(datum, addressRange)
          } while (addressRange)
        }
      } 
    }
  })
}

// Remove all the original elements from the template, leaving just one temporary element,
// since the body has to have at least one element or there will be an error
function emptyBody(doc) {
  let body = doc.getBody()
  let tempParagraph = body.appendParagraph(tempText)
  while (body.getNumChildren() > 1) body.removeChild(body.getChild(0))
}

// Remove the temporary element created by emptyBody
function removeTempElement(doc) {
  let body = doc.getBody()
  if (body.getChild(0).asText().getText() === tempText) body.removeChild(body.getChild(0))
}

function getManifestData(filterFunction) {
  // Get all the raw data
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const drivers = getRangeValuesAsTable(ss.getSheetByName("Drivers").getDataRange())
  const vehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())
  const customers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
  const trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange())
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
  // TODO add date and run ID to code
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
  elementTypesWithText = [DocumentApp.ElementType.BODY_SECTION,DocumentApp.ElementType.EQUATION,DocumentApp.ElementType.EQUATION_FUNCTION,
                          DocumentApp.ElementType.FOOTER_SECTION,DocumentApp.ElementType.FOOTNOTE_SECTION,DocumentApp.ElementType.HEADER_SECTION,
                          DocumentApp.ElementType.LIST_ITEM,DocumentApp.ElementType.PARAGRAPH,DocumentApp.ElementType.TABLE,
                          DocumentApp.ElementType.TABLE_CELL,DocumentApp.ElementType.TABLE_OF_CONTENTS,DocumentApp.ElementType.TABLE_ROW,
                          DocumentApp.ElementType.TEXT]
  return (elementTypesWithText.indexOf(element.getType()) > -1)
}

function deleteAllNamedRanges(doc) {
  doc.getNamedRanges().forEach(range => {
    range.remove()
  })
}

function prepareTemplate(doc) {
  //let doc = DocumentApp.openById(driverManifestTemplateDocId)
  deleteAllNamedRanges(doc)
  let body = doc.getBody()
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
        if (i < c) { i++ }
        let sectionName = match.groups["sectionName"]
        let regex = new RegExp(`^\\s*\\[END ${sectionName}\\]\\s*$`)
        let rangeBuilder = doc.newRange()
        for (let stayInLoop = true; stayInLoop && i < c ; i++) {
          element = body.getChild(i) 
          if (elementHasText(element) && element.getText().match(regex)) {
            if (rangeBuilder.getRangeElements().length > 0) {
              doc.addNamedRange(sectionName, rangeBuilder.build())
            }
            stayInLoop = false
          } else {
            rangeBuilder.addElement(element)
          }
        }
      }
    }
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