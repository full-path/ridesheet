function getManifestData(date) {
  if (!date) date = new Date(2020, 5, 1)
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  
  // Get all the raw data
  const drivers = getDataRangeAsTable(ss.getSheetByName("Drivers").getDataRange().getValues())
  const vehicles = getDataRangeAsTable(ss.getSheetByName("Vehicles").getDataRange().getValues())
  const customers = getDataRangeAsTable(ss.getSheetByName("Customers").getDataRange().getValues())
  const trips = getDataRangeAsTable(ss.getSheetByName("Trips").getDataRange().getValues())
  const manifestTrips = trips.filter(row => new Date(row["Trip Date"]).valueOf() == date.valueOf())
  
  // Pull in the lookup table data
  manifestTrips.forEach(tripRow => {
    mergeAttributes(tripRow, drivers,   "Driver ID"  )
    mergeAttributes(tripRow, vehicles,  "Vehicle ID" )
    mergeAttributes(tripRow, customers, "Customer ID")
  })
  
  // For events, create to rows for each trip -- one for PU, and one for DO
  let pickups = manifestTrips.map(tripRow => {
    let newRow = Object.assign({},tripRow)
    newRow["Section Name"] = "PICKUP"
    newRow["Event Name"]   = "pickup"
    newRow["Event Time"]   = tripRow["PU Time"]
    newRow["Sort Field"]   = Utilities.formatDate(new Date(tripRow["PU Time"]), timeZone, "HH:mm") + " PU"
    return newRow
  })
  let dropOffs = manifestTrips.map(tripRow => {
    let newRow = Object.assign({},tripRow)
    newRow["Section Name"] = "DROP OFF"
    newRow["Event Name"]   = "drop off"
    newRow["Event Time"]   = tripRow["DO Time"]
    newRow["Sort Field"]   = Utilities.formatDate(new Date(tripRow["DO Time"]), timeZone, "HH:mm") + " DO"
    return newRow
  })
  let manifestEvents = pickups.concat(dropOffs).sort((a,b) => {
    if (a["Sort Field"] < b["Sort Field"]) { return -1 }
    if (a["Sort Field"] > b["Sort Field"]) { return  1 }
    return 0
  })
  
  // Group the trips into runs -- all the trips with the same driver and vehicle
  let manifestRuns = manifestEvents.map(event => {
    let runAttrs = {}
    runAttrs["Driver ID"]    = event["Driver ID"]
    runAttrs["Vehicle ID"]   = event["Vehicle ID"]
    runAttrs["Driver Name"]  = event["Driver Name"]
    runAttrs["Vehicle Name"] = event["Vehicle Name"]
    runAttrs["Trip Date"]    = event["Trip Date"]
    runAttrs["Trips"]        = []
    runAttrs["Events"]       = []
    return runAttrs
  })
  manifestRuns.filter((v,i,a) => a.findIndex(t => (t["Driver ID"] === v["Driver ID"] && t["Vehicle ID"] === v["Vehicle ID"])) === i )
  manifestEvents.forEach(event => {
    let matchedRun = manifestRuns.find(run => run["Driver ID"] === event["Driver ID"] && run["Vehicle ID"] === event["Vehicle ID"])
    matchedRun["Events"].push(event)
  })
  manifestTrips.forEach(trip => {
    let matchedRun = manifestRuns.find(run => run["Driver ID"] === trip["Driver ID"] && run["Vehicle ID"] === trip["Vehicle ID"])
    matchedRun["Trips"].push(trip)
  })
  
  return manifestRuns
}

function createDriverManifest(manifestDate, driverId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const driverRow = findFirstRowByHeaderNames({"Driver ID": driverId},ss.getSheetByName("Drivers"))
  const driverName = getValueByHeaderName("Driver Name",driverRow)
  const manifestFileName = Utilities.formatDate(manifestDate, timeZone, "yyyy-MM-dd") + " Manifest for " + driverName
  const manifestFolder = DriveApp.getFolderById(driverManifestFolderId)
  const manifestFile = DriveApp.getFileById(driverManifestTemplateDocId).makeCopy(manifestFolder).setName(manifestFileName)
  const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  const manifestDoc = DocumentApp.openById(manifestFile.getId())
  deleteAllNamedRanges(templateDoc)
  const manifestSections = templateDoc.getBody().getParent()
  for (let i = 0; i < manifestSections.getNumChildren(); i += 1) {
    log(manifestSections.getChild(i).getText())
  }
}

function simpleDocTest() {
  log("Getting Doc")
  const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  log("Got Doc")
  const templateBody = templateDoc.getBody()
  log("Got Body")
  for (let i = 0; i < templateBody.getNumChildren(); i += 1) {
    log(i, templateBody.getChild(i).getType(), elementHasText(templateBody.getChild(i)))
  }
}

function testCreateDriverManifest() {
  //log(PropertiesService.getDocumentProperties().getProperty("Drivers" + sheetPropertySuffix))
  //log("Started")
  //createDriverManifest(new Date(),"DD")
  createNamedRanges()
}

function logProperties() {
  docProps = PropertiesService.getDocumentProperties()
  docProps.getKeys().forEach(prop => {
    log(prop,docProps.getProperty(prop))
  })
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

function createNamedRanges() {
  clearLog()
  let doc = DocumentApp.openById(driverManifestTemplateDocId)
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
  log("End of document reached.")
}

function copyNamedRanges(source, destination) {
  ss = SpreadsheetApp.getActiveSpreadsheet()
  const manifestFileName = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd") + " Test Manifest"
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

function verifyNamedRanges() {
  let doc = DocumentApp.openById(driverManifestTemplateDocId)
  doc.getNamedRanges().forEach(range => {
    log(range.getName())                             
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
