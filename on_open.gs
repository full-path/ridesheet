function onOpen(e) {
  const startTime = new Date()
  //storeHeaderInformation(e)
  let ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu('RideSheet')
  menu.addItem('Create Manifests', 'createManifests')
  menu.addItem('Update Properties', 'presentProperties')
  menu.addToUi()
  
  log("onOpen duration:",(new Date()) - startTime)
}

//function onSelectionChange(e) { 
//  e.source.toast(e.range.getSheet().getName() + '!' + e.range.getA1Notation())
//}

