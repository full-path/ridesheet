// Pull notification email(s) from Doc Props
function sendNotification(subject, body) {
  let defaultEmail = getDocProp('notificationEmail')
  MailApp.sendEmail(defaultEmail, subject, body)
}

function sendSharedTripNotifications() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sharedTrips = getRangeValuesAsTable(ss.getSheetByName("Trips")
    .getDataRange())
    .filter(tripRow =>( tripRow['Share With'] === true) && (!tripRow['Shared']))
  const numShared = sharedTrips.length
  if (numShared > 0) {
    let agencyName = getDocProp('providerName')
    let subject = agencyName + ' has shared trips with you'
    let body = 'You have ' + numShared.toString() + ' shared trips waiting for review. To view shared trips, open the Outside Trips sheet in RideSheet. If the available trips are not visible, use the API menu to "Get Trip Requests"'
    sendNotification(subject, body)
    markTripsAsShared(sharedTrips)
  }
}

function markTripsAsShared(sharedTrips) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const colUpdate = {'Shared': true}
  sharedTrips.forEach(sharedTrip => {
    let rowPosition = sharedTrip._rowPosition
    let currentRow = tripSheet.getRange("A" + rowPosition + ":" + rowPosition)
    setValuesByHeaderNames([colUpdate], currentRow)
  })
}