function updateDriverCalendars() {
  const startTime = new Date()
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips")
    
    // Update all trips with matching calendar events
    let range = sheet.getDataRange()
    let values = getRangeValuesAsTable(range)
    let newValues = []
    values.forEach(row => newValues.push(updateTripCalendarEvent(row)))
    setValuesByHeaderNames(newValues,range)
    
    // Remove any calendar events that might have gotten left around, unattached to trips
    let calendarIds = getDriverCalendarIds()
    calendarIds.push(getDocProp("calendarIdForUnassignedTrips"))
    let tripEvents = {}
    calendarIds.forEach(calendarId => tripEvents[calendarId] = [])
    values = getRangeValuesAsTable(range)
    values.forEach(row => {
      if (row["Trip Date"] >= dateToday() && row["Driver Calendar ID"] && row["Trip Event ID"]) {
        tripEvents[row["Driver Calendar ID"]].push(row["Trip Event ID"])
      }
    })
    calendarIds.forEach(calendarId => {
      const calendar = getCalendarById(calendarId)
      if (calendar) {
        const events = calendar.getEvents(dateToday(), dateAdd(dateToday(), 30))
        events.forEach(event => {
          if (tripEvents[calendarId].indexOf(event.getId()) === -1 && event.getTag("RideSheet") === "RideSheet") event.deleteEvent()
        })
      }
    })
  } catch(e) {
    logError(e)
  } finally {
    log("updateDriverCalendars duration:",(new Date()) - startTime)
  }
}

function updateTripCalendarEvent(tripValues) {
  try {
    let newTripValues = {}
    let calendarId
    if (tripValues["Trip Date"] >= dateToday()) {
      if (tripValues["PU Time"] && tripValues["DO Time"] && tripValues["Customer Name and ID"]) {
        if (tripValues["Driver ID"]) {
          const ss = SpreadsheetApp.getActiveSpreadsheet()
          const driverSheet = ss.getSheetByName("Drivers")
          const driverValues = findFirstRowByHeaderNames(driverSheet, function(row) { return row["Driver ID"] === tripValues["Driver ID"] })
          if (driverValues) calendarId = driverValues["Driver Calendar ID"]
        } else {
          calendarId = getDocProp("calendarIdForUnassignedTrips")
        }      
        if (calendarId) {
          newTripValues["Driver Calendar ID"] = calendarId
          newTripValues["Trip Event ID"] = createTripCalendarEvent(calendarId, tripValues)
        } else {
          newTripValues["Driver Calendar ID"] = null
          newTripValues["Trip Event ID"] = null
        }
      } else {
        newTripValues["Driver Calendar ID"] = null
        newTripValues["Trip Event ID"] = null
      }
    }
    return newTripValues
  } catch(e) {
    logError(e)
  }
}

function createTripCalendarEvent(calendarId, tripValues) {
  try {
    let newTripValues = {}
    const startTime = new Date(tripValues["Trip Date"].getTime() + timeOnly(tripValues["PU Time"]))
    const endTime = new Date(tripValues["Trip Date"].getTime() + timeOnly(tripValues["DO Time"]))
    const calendar = getCalendarById(calendarId)
    if (calendar) {
      const event = calendar.createEvent(tripValues["Customer Name and ID"], startTime, endTime, {
        location: tripValues["PU Address"],
        description: "Generated automatically by RideSheet"
      })
      event.setTag("RideSheet","RideSheet")
      return event.getId()
    }
  } catch(e) {
    log("createTripCalendarEvent", e.name + ': ' + e.message)
  }
}

function deleteTripCalendarEvent(tripValues) {
  try {
    let newTripValues = {}
    if (tripValues["Driver Calendar ID"] && tripValues["Trip Event ID"]) {
      const calendar = getCalendarById(tripValues["Driver Calendar ID"])
      if (calendar) {
        let event = calendar.getEventById(tripValues["Trip Event ID"])
        event.deleteEvent()
        newTripValues["Driver Calendar ID"] = null
        newTripValues["Trip Event ID"] = null
      }
    }
    return newTripValues
  } catch(e) {
    logError(e)
  }
}

function getDriverCalendarIds() {
  let driverSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drivers")
  let driverValues = getRangeValuesAsTable(driverSheet.getDataRange())
  let result = []
  driverValues.forEach(row => {
    if (row["Driver Calendar ID"]) result.push(row["Driver Calendar ID"])
  })
  return result
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
//    const calendar = CalendarApp.getCalendarById(getDocProp("calendarIdForUnassignedTrips"))
//    const event = calendar.createAllDayEvent("Test", dateToday())
//    log(event.getId())
    log(JSON.stringify(getDriverCalendarIds()))
  } catch(e) {
    msg = "Calendar error: " + e.name + ' — ' + e.message
    log(msg)
  }
}