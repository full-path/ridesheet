function buildLocalMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('Marshall County')
  menu.addItem('Create all-day manifest', 'createLocalManifestByDay')
  menu.addItem('Refresh Dispatch Sheet', 'refreshDispatchSheetLocal')
  menu.addToUi()
}
