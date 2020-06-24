function onOpen(e) {
  const startTime = new Date()
  //storeHeaderInformation(e)
  let ui = SpreadsheetApp.getUi()
  ui.createMenu('RideSheet').addItem('Create Manifests', 'createManifests').addToUi()
  log("onOpen duration:",(new Date()) - startTime)
}