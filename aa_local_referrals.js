function createReferral(sourceRow) {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()
    const destSheet       = ss.getSheetByName("TDS Referrals")
    const sourceData      = getRangeValuesAsTable(sourceRow)[0]
    const matchingColumns = [
      'Customer First Name',
      'Customer Last Name',
      'Phone Number',
      'Email',
      'Home Address',
      'Call ID'
    ]

    log(JSON.stringify(sourceData))
    let newReferral = {}
    matchingColumns.map((col) => newReferral[col] = sourceData[col])
    newReferral['Referral Date'] = dateToday()
    newReferral['Agency'] = sourceData['Make TDS Referral To']
    if (createRow(destSheet,newReferral)) {
      ss.toast("Referral created")
    }
  } catch(e) { logError(e) }
}
