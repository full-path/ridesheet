function createReferral(sourceRow) {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()
    const destSheet       = ss.getSheetByName("TDS Referrals")
    const sourceData      = getRangeValuesAsTable(sourceRow)[0]
    const columnMapping = {
      'Customer First Name': 'Customer First Name',
      'Customer Last Name': 'Customer Last Name',
      'Phone Number': 'Mobile Phone Number',
      'Email': 'Email',
      'Home Address': 'Home Address',
      'Veteran?': 'Veteran?',
      'Call ID': 'Call ID',
    }

    if (sourceData["Call ID"]) {
      let newReferral = {}
      Object.keys(columnMapping).map((sourceCol) => newReferral[columnMapping[sourceCol]] = sourceData[sourceCol])
      newReferral['Language'] = sourceData['Other Language'] || sourceData['Language']
      newReferral['Referral Date'] = dateToday()
      newReferral['Agency'] = sourceData['Make TDS Referral To']
      if (createRow(destSheet,newReferral)) {
        ss.toast("Referral created")
      }
    } else {
      ss.toast("Call ID is missing.","Referral not created" )
    }
  } catch(e) { logError(e) }
}
