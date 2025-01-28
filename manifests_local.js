// TODO:
// - add to menu
// - add custom docProp for the single manifest template? Or add to constants_local?

// Use regular manifest functions
// - Use different template
// - Grab *all* runs for day, ignore whether vehicle/driver are set?
// - To Discuss: Will this show all trips, or only unassigned trips?
const TIMELINE_MANIFEST_TEMPLATE_ID = "16NvrVzqC1lMY0e5SQ2sPwMAa1625IigLpup8pDVTImY"

function createTimelineManifest() {
    let startTime = new Date()
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeSheet = ss.getActiveSheet()
    const ui = SpreadsheetApp.getUi()
    let defaultDate
    let manifestDate
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
        "Enter date for manifest. Leave blank for " + formatDate(defaultDate, null, null),
        ui.ButtonSet.OK_CANCEL)
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
      ui.alert("Action cancelled as requested.")
      return
    }
  
    const templateDoc = DocumentApp.openById(TIMELINE_MANIFEST_TEMPLATE_ID)
    prepareTemplate(templateDoc)
    let runs = getManifestData(manifestDate)
    buildTimelineManifest(runs[0], templateDoc)
    log('All manifests created',(new Date()).getTime() - startTime.getTime())
  }

  // TODO: Could make a change to the main branch createManifest to accept the templateDocId
  function buildTimelineManifest(run, templateDoc) {
    const driverManifestFolderId = getDocProp("driverManifestFolderId")
    const templateFileId = TIMELINE_MANIFEST_TEMPLATE_ID
    const manifestDoc = copyTimelineTemplateManifest(run, templateFileId, driverManifestFolderId)
    emptyBody(manifestDoc)
    populateManifest(manifestDoc, templateDoc, run)
    removeTempElement(manifestDoc)
    manifestDoc.saveAndClose()
    const manifestFile = DriveApp.getFileById(manifestDoc.getId())
    let manifestPDFFile = DriveApp.getFolderById(driverManifestFolderId).createFile(manifestFile.getBlob().getAs("application/pdf"))
    manifestPDFFile.setName(manifestFile + ".pdf")
}

function copyTimelineTemplateManifest(run, templateFileId, driverManifestFolderId) {
    const manifestFileName = `${formatDate(run["Trip Date"], null, "yyyy-MM-dd")} manifest for all trips`
    const manifestFolder   = DriveApp.getFolderById(driverManifestFolderId)
    const manifestFile     = DriveApp.getFileById(templateFileId).makeCopy(manifestFolder).setName(manifestFileName)
    const manifestDoc      = DocumentApp.openById(manifestFile.getId())
    return manifestDoc
  }