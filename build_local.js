function buildLocalMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('CallSheet')
  menu.addItem('Send referrals', 'sendCustomerReferrals')
  menu.addToUi()
}
