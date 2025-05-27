function buildLocalMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('Clay County')
  menu.addItem('Refresh Dispatch Sheet', 'refreshDispatchSheetLocal')
  menu.addToUi()
}
