function onOpen(e) {
  const startTime = new Date()
  
  repairProps()
  let ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu('RideSheet')
  menu.addItem('Create manifests', 'createManifests')
  menu.addItem('Send past trips and runs to review', 'moveTripsToReview')
  menu.addItem('Send reviewed trips and runs to archive', 'moveTripsToArchive')
  menu.addSeparator()
  let settingsMenu = ui.createMenu('Settings')
  settingsMenu.addItem('Update properties', 'presentProperties')
  menu.addSubMenu(settingsMenu)
  menu.addToUi()

  log("onOpen duration:",(new Date()) - startTime)
}

//function onSelectionChange(e) { 
//  e.source.toast(e.range.getSheet().getName() + '!' + e.range.getA1Notation())
//}