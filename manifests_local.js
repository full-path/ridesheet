function createLocalManifestByDay() {
  try {
    const templateDocId = getDocProp("driverManifestTemplateDocId")
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeSheet = ss.getActiveSheet()
    let defaultDate
    let date
    let runDate
    let tripSheetName = "Trips"

    const ui = SpreadsheetApp.getUi()
    if (activeSheet.getName() == "Trips" || activeSheet.getName() == "Dispatch") {
      runDate = getValueByHeaderName("Trip Date", getFullRows(activeSheet.getActiveCell()))
      tripSheetName = activeSheet.getName()
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

    let promptResult = ui.prompt("Create Manifest",
        "Enter date for manifest. Leave blank for " + formatDate(defaultDate, null, null),
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
    const manifestData = getManifestData(dateFilter, tripSheetName)
    const groupedManifestData = groupManifestDataByDay(manifestData)
    const manifestCount = createManifests(templateDocId, groupedManifestData, getManifestFileNameByDay)
    ss.toast(manifestCount + " created.","Manifest creation complete.")
  } catch(e) { logError(e) }
}

function groupManifestDataByDay(manifestData) {
  // Group the trips by day
  let manifestGroups = []
  manifestData.trips.forEach(trip => {
    const runIndex = manifestGroups.findIndex(row => 
      row["Trip Date"].getTime() === trip["Trip Date"].getTime()
    )
    if (runIndex == -1) {
      let newGroup = {}
      newGroup["Trip Date"] = trip["Trip Date"]
      newGroup["Trips"]     = [trip]
      newGroup["Events"]    = []
      manifestGroups.push(newGroup)
    } else {
      manifestGroups[runIndex]["Trips"].push(trip)
    }
  })
  // Group the manifest events into the same groups
  manifestData.events.forEach(event => {
    let matchedRun = manifestGroups.find(row => 
      row["Trip Date"].getTime() === event["Trip Date"].getTime()
    )
    matchedRun["Events"].push(event)
  })
  return manifestGroups
}

function getManifestFileNameByDay(manifestGroup) {
  const manifestFileName = `${formatDate(manifestGroup["Trip Date"], null, "yyyy-MM-dd")} all trips`
  return manifestFileName
}
