const tempText = "{{Temporary Text}}"

function createManifests() {
  let startTime = new Date()
  ss = SpreadsheetApp.getActiveSpreadsheet()
  activeSheet = ss.getActiveSheet()
  ui = SpreadsheetApp.getUi()
  let defaultDate
  let manifestDate
  if (activeSheet.getName() == "Trips") {
    tripRow = getFullRows(activeSheet.getActiveCell())
    tripDate = getValueByHeaderName("Trip Date",tripRow)
    if (isValidDate(tripDate)) {
      defaultDate = tripDate
    } else {
      // tomorrow, at midnight
      defaultDate = dateOnly(dateAdd(new Date(), 1))
    }
  } else {
    // tomorrow, at midnight
    defaultDate = dateOnly(dateAdd(new Date(), 1))
  }
  log("Prompt ready:",(new Date()) - startTime)
  let promptResult = ui.prompt("Create Manifests", "Enter date for manifests. Leave blank for " + formatDate(defaultDate, null, null), ui.ButtonSet.OK_CANCEL)
  startTime = new Date()
  if (promptResult.getResponseText() == "") {
    manifestDate = defaultDate
  } else {
    manifestDate = parseDate(promptResult.getResponseText(),"Invalid Date")
  }
  if (!isValidDate(manifestDate)) {
    ui.alert("Invalid date, action cancelled.")
    return
  } else if (promptResult.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Action cancelled.")
    return
  }
  const templateDoc = DocumentApp.openById(getDocProp("driverManifestTemplateDocId"))
  prepareTemplate(templateDoc)
  log("Template prepared:",(new Date()) - startTime)
  
  let runs = getManifestData(manifestDate)
  log("Data acquired:",(new Date()) - startTime)

  runs.forEach((run, i) => {
    const driverManifestFolderId = getDocProp("driverManifestFolderId")
    const templateFileId = getDocProp("driverManifestTemplateDocId")
    const manifestDoc = createManifest(run, templateFileId, driverManifestFolderId)
    log("Created manifest " + i + ":",(new Date()) - startTime)
    emptyBody(manifestDoc)
    populateManifest(manifestDoc, templateDoc, run)
    removeTempElement(manifestDoc)
    manifestDoc.saveAndClose()
    const manifestFile = DriveApp.getFileById(manifestDoc.getId())
    let manifestPDFFile = DriveApp.getFolderById(driverManifestFolderId).createFile(manifestFile.getBlob().getAs("application/pdf"))
    manifestPDFFile.setName(manifestFile + ".pdf")
    log("Populated manifest " + i + ":",(new Date()) - startTime)
  })
  log(`All manifests created:`,(new Date()) - startTime)
}

function createManifest(run, templateFileId, driverManifestFolderId) {
  const manifestFileName = `${formatDate(run["Trip Date"], null, "yyyy-MM-dd")} manifest for ${run["Driver Name"]} on ${run["Vehicle Name"]}`
  const manifestFolder   = DriveApp.getFolderById(driverManifestFolderId)
  const manifestFile     = DriveApp.getFileById(templateFileId).makeCopy(manifestFolder).setName(manifestFileName)
  const manifestDoc      = DocumentApp.openById(manifestFile.getId())
  return manifestDoc
}
                                                  
function testCreateManifest() {
  //const driverManifestTemplateDocId = "https://docs.google.com/document/d/1SPjf_8oVA2BM6wTzSf3Ww521Slet1p6NOorKtDn_HOQ/edit"
  const driverManifestTemplateDocId = "1SPjf_8oVA2BM6wTzSf3Ww521Slet1p6NOorKtDn_HOQ"
  //const driverManifestFolderId = "https://drive.google.com/drive/folders/1MuBNXLE-ziUEDY56gWR-OZ1Y9fKYild0"
  const driverManifestFolderId = "1MuBNXLE-ziUEDY56gWR-OZ1Y9fKYild0"
  //const manifestFileName = `${formatDate(run["Trip Date"], null, "yyyy-MM-dd")} manifest for ${run["Driver Name"]} on ${run["Vehicle Name"]}`
  const manifestFileName = "Test Manifest"
  const manifestFolder   = DriveApp.getFolderById(driverManifestFolderId)
  const manifestFile     = DriveApp.getFileById(driverManifestTemplateDocId).makeCopy(manifestFolder).setName(manifestFileName)
  const manifestDoc      = DocumentApp.openByUrl() (manifestFile.getId())
  log(manifestDoc.getName())
}
                                                  
function populateManifest(manifestDoc, templateDoc, run) {
  //templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  
  const manifestBody   = manifestDoc.getBody()
  const manifestParent = manifestBody.getParent()
  
  // Replace the fields in the document header and footer sections
  // There may be up to four, if there are different headers or footers for the first page
  for (let i = 0, c = manifestParent.getNumChildren(); i < c; i++) {
    let section = manifestParent.getChild(i)
    if (section.getType() === DocumentApp.ElementType.HEADER_SECTION) {
      replaceText(section, run["Events"][0])
    } else if (section.getType() === DocumentApp.ElementType.FOOTER_SECTION) {
      replaceText(section, run["Events"][run["Events"].length - 1])
    }  
  }

  // Add the header elements
  //log ("Header")
  replaceTextInRange(templateDoc.getNamedRanges("HEADER")[0].getRange(), manifestBody, run["Events"][0])
  // Add all the PU and DO elements. Use the section name of each event to decide whether to add a PU or DO range.
  run["Events"].forEach((event, i) => {
    //log ("Event " + i)
    replaceTextInRange(templateDoc.getNamedRanges(event["Section Name"])[0].getRange(), manifestBody, event)
  })
  // Add the footer elements
 // log ("Footer")
  replaceTextInRange(templateDoc.getNamedRanges("FOOTER")[0].getRange(), manifestBody, run["Events"][run["Events"].length - 1])
}

function replaceTextInRange(range, docSection, data) {
  let elements = range.getRangeElements()
  elements.forEach(element => {
    newElement = element.getElement().copy()
    replaceText(newElement, data)
    appendElement(docSection, newElement)
  })
}

function replaceText(element, data) {
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
    //log("Field: " + field, "Data: " + data[field], "Type: " + Object.prototype.toString.call(data[field]))
    if (Object.keys(data).indexOf(field) != -1) {
      element.replaceText("{" + field + "}", datum)
      if (field.match(/\baddress\b/i)) {
        let url = createGoogleMapsDirectionsURL(datum)
        let text = element.asText()
        let addressRange = text.findText(escapeRegex(datum))
        if (addressRange) {
          do {
            text.setLinkUrl(addressRange.getStartOffset(), addressRange.getEndOffsetInclusive(), url)
            addressRange = text.findText(datum, addressRange)
          } while (addressRange)
        } else {
          //log(text.getText(), datum)
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

function getManifestData(date) {
  if (!date) date = new Date(2020, 5, 1)
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  
  // Get all the raw data
  const drivers = getRangeValuesAsTable(ss.getSheetByName("Drivers").getDataRange())
  const vehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())
  const customers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
  const trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange())
  const manifestTrips = trips.filter(row => new Date(row["Trip Date"]).valueOf() == date.valueOf())
  
  // Pull in the lookup table data
  manifestTrips.forEach(tripRow => {
    mergeAttributes(tripRow, drivers,   "Driver ID"  )
    mergeAttributes(tripRow, vehicles,  "Vehicle ID" )
    mergeAttributes(tripRow, customers, "Customer ID")
    tripRow["Manifest Creation Time"] = new Date()
    tripRow["Manifest Creation Date"] = new Date()
  })
  
  // For events, create to rows for each trip -- one for PU, and one for DO
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
  
  // Group the trips into runs -- A run is a collection of trips on the same day with the same driver and vehicle
  let manifestRuns = []
  manifestTrips.forEach(trip => {
    let runIndex = manifestRuns.findIndex(r => r["Driver ID"] == trip["Driver ID"] && r["Vehicle ID"] == trip["Vehicle ID"])
    if (runIndex == -1) {
      let newRun = {}
      newRun["Driver ID"]    = trip["Driver ID"]
      newRun["Vehicle ID"]   = trip["Vehicle ID"]
      newRun["Driver Name"]  = trip["Driver Name"]
      newRun["Driver Email"]  = trip["Driver Email"]
      newRun["Vehicle Name"] = trip["Vehicle Name"]
      newRun["Trip Date"]    = trip["Trip Date"]
      newRun["Trips"]        = [trip]
      newRun["Events"]       = []
      manifestRuns.push(newRun)
    } else {
      manifestRuns[runIndex]["Trips"].push(trip)
    }
  })
  manifestEvents.forEach(event => {
    let matchedRun = manifestRuns.find(run => run["Driver ID"] === event["Driver ID"] && run["Vehicle ID"] === event["Vehicle ID"])
    matchedRun["Events"].push(event)
  })
  
  return manifestRuns
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
  //clearLog()
  //let doc = DocumentApp.openById(driverManifestTemplateDocId)
  deleteAllNamedRanges(doc)
  let body = doc.getBody()
  for (let i = 0, c = body.getNumChildren(); i < c; i++) {
    let element = body.getChild(i) 
    //log("Got element " + i, element.getType(), elementHasText(element))
    if (elementHasText(element)) {
      //log(i,"Not in a section, searching for a BEGIN statement")
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
        //log(i, "Found start of section " + sectionName + ", starting to build named range.")
        //log(i, "Will search for " + regex.toString())
        for (let stayInLoop = true; stayInLoop && i < c ; i++) {
          element = body.getChild(i) 
          //log(i, element.getText(), (element.getText().match(regex)))
          if (elementHasText(element) && element.getText().match(regex)) {
            if (rangeBuilder.getRangeElements().length > 0) {
              doc.addNamedRange(sectionName, rangeBuilder.build())
              //log(i, "Found end of section " + sectionName + ". Saved named range " + sectionName)
            } else {
              //log(i, "Found end of section " + sectionName + ". No elements found. No range saved.")
            }
            stayInLoop = false
          } else {
            rangeBuilder.addElement(element)
            //log(i, "Adding element to " + sectionName)
          }
        }
      }
    }
  }
  //log("End of document reached.")
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
      Object.keys(matchingSecondaryRow).forEach(key => { 
        if (!primaryRow.hasOwnProperty()) { 
          primaryRow[key] = matchingSecondaryRow[key] 
        } 
      }) 
    }
  }
}