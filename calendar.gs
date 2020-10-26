function createEvent(calendarID) {
  const calendar = CalendarApp.getCalendarById("fullpath.io_vlmbs3tfapdnpmnn6a9vsjfdrg@group.calendar.google.com")
  //event = calendar.createEvent("My test event", new Date, timeAdd(new Date, 3600000))
  let event = calendar.getEventById("a8af0j0t32cbq5ba36b6me4jfo@google.com")
  event.setTime(timeAdd(event.getStartTime(), 3600000), timeAdd(event.getEndTime(), 3600000))
  log(event.getId())
}

function createTripCalendarEvent(tripRange) {
  let tripValues = getRangeValuesAsTable(tripRange)[0]
  if (tripValues["Driver ID"]) {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const driverSheet = ss.getSheetByName("Drivers")
    const driverValues = findFirstRowByHeaderNames(driverSheet, function(row) { return row["Driver ID"] == tripValues["Driver ID"] })
    if (driverValues["Driver Calendar ID"]) {
      const calendar = getCalendarById(driverValues["Driver Calendar ID"])
      if (calendar) {
        const startTime = tripValues["Trip Date"].getTime() + tripValues["PU Time"].getTime()
        const endTime = tripValues["Trip Date"].getTime() + tripValues["DO Time"].getTime()
        const event = calendar.createEvent(tripValues["Customer Name and ID"], startTime, endTime)
        tripValues["Driver Calendar ID"] = driverValues["Driver Calendar ID"]
        tripValues["Trip Event ID"] = event.getId()
        setValuesByHeaderNames([tripValues], tripRange)
      } else {
        ss.toast("Could not connect with driver calendar.")
      }
    }
  }
}

function updateTripCalendarEvent(tripRange) {
  let tripValues = getRangeValuesAsTable(tripRange)[0]
  if (tripValues["Driver ID"] && tripValues["Calendar ID"]) {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const driverSheet = ss.getSheetByName("Drivers")
    const driverValues = findFirstRowByHeaderNames(driverSheet, function(row) { return row["Driver ID"] == tripValues["Driver ID"] })
    if (driverValues["Driver Calendar ID"] != tripValues["Driver Calendar ID"]) {
      // Need to delete old event under old driver calendar and create new event under the new calendar
    } else {
      // Update attributesof existing trip
      const event = CalendarApp.getEventById(tripValues["Trip Event ID"])
    }
  }
}

function deleteTripCalendarEvent(tripRange) {
  let tripValues = getRangeValuesAsTable(tripRange)[0]
  
}

function syncTripCalendarEvents() {
  
}

function getCalendarById(calendarId) {
  try {
    return CalendarApp.getCalendarById(calendarId)
  } catch(e) {
    return false
  }
}
  