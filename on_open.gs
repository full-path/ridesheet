function onOpen(e) {
  const startTime = new Date()
  try {
    repairProps()
  } catch(e) {
    log("repairProps", e.name + ': ' + e.message)
  }
  try {
    let ui = SpreadsheetApp.getUi()
    let menu = ui.createMenu('RideSheet')
    menu.addItem('Create return trip', 'createReturnTrip')
    menu.addItem('Create manifests', 'createManifests')
    menu.addItem('Send past trips and runs to review', 'moveTripsToReview')
    menu.addItem('Send reviewed trips and runs to archive', 'moveTripsToArchive')
    menu.addSeparator()
    let settingsMenu = ui.createMenu('Settings')
    settingsMenu.addItem('Update properties', 'presentProperties')
    menu.addSubMenu(settingsMenu)
    menu.addToUi()
  } catch(e) {
    log("Custom Menu", e.name + ': ' + e.message)
  }
  try {
    addAndPopulateTripIdColumn()
  } catch(e) {
    log("Custom Menu", e.name + ': ' + e.message)
  }
  log("onOpen duration:",(new Date()) - startTime)
}

//function onSelectionChange(e) { 
//  e.source.toast(e.range.getSheet().getName() + '!' + e.range.getA1Notation())
//}

function addAndPopulateTripIdColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  if (getSheetHeaderNames(tripSheet).indexOf("Trip ID") === -1) {
    const lastRowPosition = tripSheet.getLastRow()
    let lastColumnPosition = tripSheet.getLastColumn()
    tripSheet.insertColumnAfter(lastColumnPosition)
    lastColumnPosition++
    tripSheet.getRange(1, lastColumnPosition).setValue("Trip ID")
    getSheetHeaderNames(tripSheet, {forceRefresh: true})
    const dataRange = tripSheet.getRange(2, 1, lastRowPosition - 1, lastColumnPosition)
    let values = getRangeValuesAsTable(dataRange)
    values.forEach((row) => {
      if (row["Trip Date"] && row["Customer ID"]) {
        row["Trip ID"] = Utilities.getUuid()
      }
    })
    setValuesByHeaderNames(values, dataRange)
  }
}