function updateDriverCalendars() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips")
  let range = sheet.getDataRange()
  let values = getRangeValuesAsTable(range)
  let newValues = []
  values.forEach(row => {
    newValues.push(updateTripCalendarEvent(row))
  })
  log(JSON.stringify(newValues))
  setValuesByHeaderNames(newValues,range)
}

function updateTripCalendarEvent(tripValues) {
  let newTripValues = {}
  if (tripValues["Trip Date"] >= dateToday() && tripValues["PU Time"] && tripValues["DO Time"] && tripValues["Customer Name and ID"]) {
    if (tripValues["Driver ID"] && tripValues["Driver Calendar ID"]) {
      const ss = SpreadsheetApp.getActiveSpreadsheet()
      const driverSheet = ss.getSheetByName("Drivers")
      const driverValues = findFirstRowByHeaderNames(driverSheet, function(row) { return row["Driver ID"] == tripValues["Driver ID"] })
      if (driverValues["Driver Calendar ID"] != tripValues["Driver Calendar ID"]) {
        // Need to delete old event under old driver calendar and create new event under the new calendar
        log(1)
        deleteTripCalendarEvent(tripValues)
        newTripValues = createTripCalendarEvent(tripValues)
      } else {
        // Update attributes of existing trip
        const startTime = new Date(tripValues["Trip Date"].getTime() + timeOnly(tripValues["PU Time"]))
        const endTime = new Date(tripValues["Trip Date"].getTime() + timeOnly(tripValues["DO Time"]))
        log(1.5)
        const calendar = getCalendarById(tripValues["Driver Calendar ID"])
        if (calendar) {
          log(2)
          let event = calendar.getEventById(tripValues["Trip Event ID"])
          event.setTime(startTime, endTime)
          event.setTitle(tripValues["Customer Name and ID"])
          event.setLocation(tripValues["PU Address"])
        }
      }
    } else {
      newTripValues = createTripCalendarEvent(tripValues)
    }
  } else if (tripValues["Driver Calendar ID"] && tripValues["Trip Event ID"]) {
    newTripValues = deleteTripCalendarEvent(tripValues)
  }
  return newTripValues
}

function createTripCalendarEvent(tripValues) {
  try{
    let newTripValues = {}
    const startTime = new Date(tripValues["Trip Date"].getTime() + timeOnly(tripValues["PU Time"]))
    const endTime = new Date(tripValues["Trip Date"].getTime() + timeOnly(tripValues["DO Time"]))
    if (tripValues["Driver ID"]) {
      const ss = SpreadsheetApp.getActiveSpreadsheet()
      const driverSheet = ss.getSheetByName("Drivers")
      const driverValues = findFirstRowByHeaderNames(driverSheet, function(row) { return row["Driver ID"] == tripValues["Driver ID"] })
      if (driverValues["Driver Calendar ID"]) {
        log(3)
        const calendar = getCalendarById(driverValues["Driver Calendar ID"])
        if (calendar) {
          const event = calendar.createEvent(tripValues["Customer Name and ID"], startTime, endTime, {location: tripValues["PU Address"]})
          newTripValues["Driver Calendar ID"] = driverValues["Driver Calendar ID"]
          newTripValues["Trip Event ID"] = event.getId()
        }
      }
    } else {
      log(4)
      const calendar = getCalendarById(getDocProp("calendarIdForUnassignedTrips"))
      if (calendar) {
        const event = calendar.createEvent(tripValues["Customer Name and ID"], startTime, endTime, {location: tripValues["PU Address"]})
        tripValues["Driver Calendar ID"] = getDocProp("calendarIdForUnassignedTrips")
        tripValues["Trip Event ID"] = event.getId()
      }
    }
    return newTripValues
  } catch(e) {
    log("createTripCalendarEvent", e.name + ': ' + e.message)
  }
}

function deleteTripCalendarEvent(tripValues) {
  let newTripValues = {}
  if (tripValues["Calendar ID"] && tripValues["Trip Event ID"]) {
    log(5)
    const calendar = getCalendarById(tripValues["Driver Calendar ID"])
    if (calendar) {
      let event = calendar.getEventById(tripValues["Trip Event ID"])
      event.deleteEvent()
      newTripValues["Driver Calendar ID"] = null
      newTripValues["Trip Event ID"] = null
    }
  }
  return newTripValues
}

function getCalendarById(calendarId) {
  try {
    return CalendarApp.getCalendarById(calendarId)
  } catch(e) {
    msg = "Calendar error: " + e.name + ' — ' + e.message
    SpreadsheetApp.getActiveSpreadsheet().toast(msg)
    log(calendarId, msg)
  }
}

function testCal() {
  try {
    const calendar = CalendarApp.getCalendarById(getDocProp("calendarIdForUnassignedTrips"))
    const event = calendar.createAllDayEvent("Test", dateToday())
    log(event.getId())
  } catch(e) {
    msg = "Calendar error: " + e.name + ' — ' + e.message
    log(msg)
  }
}