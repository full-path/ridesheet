function buildLocalMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('RS Add-ons')
  menu.addItem('Refresh Dispatch Sheet', 'refreshDispatchSheetLocal')
  menu.addToUi()
}
