function updateDriverCalendars() {
  const startTime = new Date()
  try {
    // Update all trips with matching calendar events
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const range = ss.getSheetByName("Trips").getDataRange()
    const trips = getRangeValuesAsTable(range)
    const drivers = getRangeValuesAsTable(ss.getSheetByName("Drivers").getDataRange())
    const vehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())
    const customers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
    let newValuesToSave = [] // This gets saved back to the Trips sheet
    let newValues = []       // This is used to create the list of trips not to delete later
    trips.forEach(row => {
      mergeAttributes(row, drivers,   "Driver ID"  )
      mergeAttributes(row, vehicles,  "Vehicle ID" )
      mergeAttributes(row, customers, "Customer ID")
      let rowValuesToSave = updateTripCalendarEvent(row)
      let rowNewValues = {...rowValuesToSave}
      rowNewValues["Trip Date"] = row["Trip Date"]
      newValues.push(rowNewValues)
      newValuesToSave.push(rowValuesToSave)
    })
    setValuesByHeaderNames(newValuesToSave,range)
    
    // Remove any calendar events that might have gotten left around, unattached to trips
    // First, gather all the events that where just created so we don't delete those ones
    let calendarIds = getDriverCalendarIds()
    calendarIds.push(getDocProp("calendarIdForUnassignedTrips"))
    let tripEvents = {}
    calendarIds.forEach(calendarId => tripEvents[calendarId] = [])
    newValues.forEach(row => {
      if (row["Trip Date"] >= dateToday() && row["Driver Calendar ID"] && row["Trip Event ID"]) {
        tripEvents[row["Driver Calendar ID"]].push(row["Trip Event ID"])
      }
    })
    // Next go through all the events from today forward and delete the ones that aren't in our list.
    calendarIds.forEach(calendarId => {
      const calendar = getCalendarById(calendarId)
      if (calendar) {
        const events = calendar.getEvents(dateToday(), dateAdd(dateToday(), 30))
        events.forEach(event => {
          if (tripEvents[calendarId].indexOf(event.getId()) === -1 && event.getTag("RideSheet") === "RideSheet") {
            event.deleteEvent()
          }
        })
      }
    })
  } catch(e) { logError(e) } finally {
    log("updateDriverCalendars duration:",(new Date()) - startTime)
  }
}

function updateTripCalendarEvent(tripValues) {
  try {
    let newTripValues = {}
    let calendarId = "Unset"
    if (tripValues["Trip Date"] >= dateToday()) {
      if (tripValues["PU Time"] && tripValues["DO Time"] && tripValues["Customer Name and ID"]) {
        if (tripValues["Driver ID"]) {
          calendarId = tripValues["Driver Calendar ID"]
        } else {
          calendarId = getDocProp("calendarIdForUnassignedTrips")
        }
        if (calendarId) {
          newTripValues["Driver Calendar ID"] = calendarId
          newTripValues["Trip Event ID"] = createTripCalendarEvent(calendarId, tripValues)
          //log("Created new calendar event with ID",newTripValues["Trip Event ID"])
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
  } catch(e) { logError(e) }
}

function createTripCalendarEvent(calendarId, tripValues) {
  try {
    let newTripValues = {}
    const startTime = new Date(tripValues["Trip Date"].getTime() + timeOnlyAsMilliseconds(tripValues["PU Time"]))
    const endTime = new Date(tripValues["Trip Date"].getTime() + timeOnlyAsMilliseconds(tripValues["DO Time"]))
    const calendar = getCalendarById(calendarId)
    if (calendar) {
      const eventTitle = replaceText(getDocProp("tripCalendarEntryTitleTemplate"),tripValues)
      const event = calendar.createEvent(eventTitle, startTime, endTime, {
        description: "Generated automatically by RideSheet"
      })
      event.setTag("RideSheet","RideSheet")
      return event.getId()
    } else {
      //log("Didn't get calendar object for " + calendarId)
    }
  } catch(e) { logError(e) }
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
  } catch(e) { logError(e) }
}

function getDriverCalendarIds() {
  let driverSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drivers")
  let driverValues = getRangeValuesAsTable(driverSheet.getDataRange())
  let result = []
  driverValues.forEach(row => {
    if (row["Driver Calendar ID"] && result.indexOf(row["Driver Calendar ID"]) == -1) result.push(row["Driver Calendar ID"])
  })
  return result
}

function isCalendarEventActive(calendar, eventId) {
  try {
    const event = calendar.getEventById(eventId)
    if (event) {
      const events = calendar.getEvents(event.getStartTime(),event.getEndTime())
      return events.some((thisEvent) => thisEvent.getId() == eventId)
    } else {
      return false
    }
  } catch(e) { logError(e) }
}

function getCalendarById(calendarId) {
  try {
    return CalendarApp.getCalendarById(calendarId)
  } catch(e) { logError(e) }
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