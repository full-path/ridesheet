function updateDriverCalendars() {
  const startTime = new Date()
  let deleteOrCreateEvents = true
  let eventDeleteOrCreateCount = 0
  try {
    // GATHER DATA
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const range = ss.getSheetByName("Trips").getDataRange()
    const trips = getRangeValuesAsTable(range)
    const drivers = getRangeValuesAsTable(ss.getSheetByName("Drivers").getDataRange())
    const calendarIdForUnassignedTrips = getDocProp("calendarIdForUnassignedTrips")
    let tripChanges = trips.map(row => { return {} })

    console.log("Trips", trips.length)
    // Set up our trip filters
    const isTodayOrFutureTrip = function(row) { return row["Trip Date"].valueOf() >= dateToday().valueOf() }
    const isValidTrip = function(row) {
      return (
        row["Customer Name and ID"] &&
        row["PU Time"] &&
        row["DO Time"] &&
        timeOnlyAsMilliseconds(row["PU Time"]) < timeOnlyAsMilliseconds(row["DO Time"]) &&
        (
          !row["Trip Result"] ||
          row["Trip Result"] === "Completed"
        )
      )
    }
    const hasCalendarLinkData = function(row) {
      return (
        row["Driver Calendar ID"] ||
        row["Trip Event ID"]
      )
    }
    const hasValidLinkToEvent = function(row) {
      return (
        row["Driver Calendar ID"] &&
        row["Trip Event ID"] &&
        calendarIds.includes(row["Driver Calendar ID"])
      )
    }
    const hasNoValidLinkToEvent = function(row) {
      return (
        !row["Driver Calendar ID"] ||
        !row["Trip Event ID"] ||
        !initialCalendarEvents[row["Driver Calendar ID"]] ||
        !Object.keys(initialCalendarEvents[row["Driver Calendar ID"]]).includes(row["Trip Event ID"])
      )
    }

    // Get a list of all the calendars we'll be working with, both drivers and unassigned
    const calendarIds = getDriverCalendarIds()
    calendarIds.push(calendarIdForUnassignedTrips)
    // Get our furthest out trip date. Add a day for extra padding.
    const furthestTripDate = dateAdd(new Date(Math.max.apply(null, trips.map(trip => trip["Trip Date"]))), 1)
    // Gather all the current calendar events for all the calendars: drivers and unassigned
    let initialCalendarEvents = {}
    calendarIds.forEach(calendarId => {
      initialCalendarEvents[calendarId] = {}
      const events = getCalendarById(calendarId).getEvents(dateToday(),furthestTripDate)
      events.forEach(event => {
        if (event.getTag("RideSheet") === "RideSheet") {
          initialCalendarEvents[calendarId][event.getId()] = event
        }
      })
    })

    // Go through all the trips and clear out the [Driver Calendar ID] for those where the [Driver Calendar ID] for the trip
    // doesn't match the [Driver Calendar ID] of the selected driver (presumably because the driver has changed or been removed).
    // This will trigger deletion of the orphaned calendar event and creation of a new one under the appropriate new 
    // calendar
    trips.filter(isTodayOrFutureTrip).filter(isValidTrip).filter(hasValidLinkToEvent).forEach(row => {
      if (
        (
          !row["Driver ID"] &&
          row["Driver Calendar ID"] != calendarIdForUnassignedTrips
        ) || (
          row["Driver ID"] &&
          row["Driver Calendar ID"] != drivers.find(driver => driver["Driver ID"] == row["Driver ID"])["Driver Calendar ID"]
        )
      ) {
        row["Driver Calendar ID"] = null
      }
    })

    // Gather all the current valid calendar data in Trips
    let initialCalendarEventLinks = {}
    calendarIds.forEach(calendarId => initialCalendarEventLinks[calendarId] = [])
    trips.filter(isTodayOrFutureTrip).filter(isValidTrip).filter(hasValidLinkToEvent).forEach(row => {
      initialCalendarEventLinks[row["Driver Calendar ID"]].push(row["Trip Event ID"])
    })
    console.log("initialCalendarEventLinks",Object.keys(initialCalendarEventLinks, initialCalendarEventLinks).map(k => initialCalendarEventLinks[k].length).reduce((a, b) => a + b))

    // Gather all the calendar entries that need to be deleted because the trip
    // (or at least the link to a calendar entry) is cancelled or otherwise gone
    let eventsToDelete = {}
    calendarIds.forEach(calendarId => {
      eventsToDelete[calendarId] = Object.keys(initialCalendarEvents[calendarId]).filter(eventId => {
        return !initialCalendarEventLinks[calendarId].includes(eventId)
      })
    })
    console.log("eventsToDelete",Object.keys(eventsToDelete, eventsToDelete).map(k => eventsToDelete[k].length).reduce((a, b) => a + b))

    // Get a list of all the matches between entries and data in trips. These will need to be checked for possible updates.
    let eventsToUpdate = {}
    calendarIds.forEach(calendarId => {
      eventsToUpdate[calendarId] = Object.keys(initialCalendarEvents[calendarId]).filter(eventId => {
        return initialCalendarEventLinks[calendarId].includes(eventId)
      })
    })
    const countOfEventsToUpdate = Object.keys(eventsToUpdate).map(k => eventsToUpdate[k].length).reduce((a, b) => a + b)
    console.log("countOfEventsToUpdate",countOfEventsToUpdate)

    // Get a list of trip that are invalid for being having a calendar event but have calendar data nonetheless
    let invalidTripsWithCalendarData = trips.filter(isTodayOrFutureTrip).filter(row => !isValidTrip(row)).filter(hasCalendarLinkData)
    console.log("invalidTripsWithCalendarData",invalidTripsWithCalendarData.length)

    // Get a list of valid trips for which entries that need to be created because the trips are new or the calender entry cannot be found
    let tripsForNewEvents = trips.filter(isTodayOrFutureTrip).filter(isValidTrip).filter(hasNoValidLinkToEvent)
    console.log("tripsForNewEvents", tripsForNewEvents.length)
    //return

    // TAKE ACTION
    // For each calendar, delete the calendar entries that need that
    calendarIds.forEach(calendarId => {
      eventsToDelete[calendarId].forEach(eventId => {
        if (deleteOrCreateEvents) {
          try {
            initialCalendarEvents[calendarId][eventId].deleteEvent()
          } catch(e) {
            ss.toast("Not all events could be deleted / created. " + eventDeleteOrCreateCount +
              " successful. Please try again later for the rest.","Error")
            deleteOrCreateEvents = false
          }
          eventDeleteOrCreateCount++
        }
      })
    })

    // Update trips that should have their calendar links deleted
    invalidTripsWithCalendarData.forEach(trip => {
      let newTripValues = {}
      if (trip["Driver Calendar ID"]) newTripValues["Driver Calendar ID"] = null
      if (trip["Trip Event ID"]) newTripValues["Trip Event ID"] = null
      tripChanges[trip.rowIndex] = newTripValues
    })

    // 
    if (countOfEventsToUpdate + tripsForNewEvents.length) {
      const vehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())
      const customers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())

      calendarIds.forEach(calendarId => {
        // Check every matching calendar entry to see if it needs to be updated, and update if it is needed
        eventsToUpdate[calendarId].forEach(eventId => {
          const event = initialCalendarEvents[calendarId][eventId]
          const trip = trips.find(row => row["Trip Event ID"] === eventId)
          mergeAttributes(trip, drivers,   "Driver ID"  )
          mergeAttributes(trip, vehicles,  "Vehicle ID" )
          mergeAttributes(trip, customers, "Customer ID")
          const eventTitle = replaceText(getDocProp("tripCalendarEntryTitleTemplate"), trip)
          const startTime  = new Date(trip["Trip Date"].getTime() + timeOnlyAsMilliseconds(trip["PU Time"]))
          const endTime    = new Date(trip["Trip Date"].getTime() + timeOnlyAsMilliseconds(trip["DO Time"]))
          if (startTime.valueOf() != event.getStartTime().valueOf() || endTime.valueOf() != event.getEndTime().valueOf()) {
            event.setTime(startTime, endTime)
          }
          if (eventTitle != event.getTitle()) {
            event.setTitle(eventTitle)
          }
        })
      })

      tripsForNewEvents.forEach(trip => {
        if (deleteOrCreateEvents) {
          mergeAttributes(trip, drivers,   "Driver ID"  )
          mergeAttributes(trip, vehicles,  "Vehicle ID" )
          mergeAttributes(trip, customers, "Customer ID")
          let tripValuesToSave = updateTripCalendarEvent(trip)
          if (tripValuesToSave[Object.keys(tripValuesToSave)[0]] === -1) {
            deleteOrCreateEvents = false
            ss.toast("Not all events could be created. " + eventDeleteOrCreateCount +
              " successful. Please try again later for the rest.","Error")
          } else {
            eventDeleteOrCreateCount++
            tripChanges[trip.rowIndex] = tripValuesToSave
          }
        }
      })
    }
    setValuesByHeaderNames(tripChanges, range)
  } catch(e) {
    logError(e)
  } finally {
    if (eventDeleteOrCreateCount) log("Event deletions and creations:", eventDeleteOrCreateCount)
    log("updateDriverCalendars duration:",(new Date()) - startTime)
  }
}

function updateTripCalendarEvent(tripValues) {
  try {
    let newTripValues = {}
    let calendarId = "Unset"
    if (tripValues["Trip Date"] >= dateToday()) {
      if (
        tripValues["PU Time"] &&
        tripValues["DO Time"] &&
        timeOnlyAsMilliseconds(tripValues["PU Time"]) < timeOnlyAsMilliseconds(tripValues["DO Time"]) &&
        tripValues["Customer Name and ID"]
      ) {
        if (tripValues["Driver ID"]) {
          calendarId = tripValues["_Driver ID-attributes"]["Driver Calendar ID"]
        } else {
          calendarId = getDocProp("calendarIdForUnassignedTrips")
        }
        if (calendarId) {
          const eventID = createTripCalendarEvent(calendarId, tripValues)
          if (eventID === -1) {
            newTripValues["Driver Calendar ID"] = -1
            newTripValues["Trip Event ID"] = -1
          } else {
            newTripValues["Driver Calendar ID"] = calendarId
            newTripValues["Trip Event ID"] = eventID
          }
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
      let event
      try {
        event = calendar.createEvent(eventTitle, startTime, endTime, {
          description: "Generated automatically by RideSheet"
        })
      } catch(e) {
        console.log(e)
        return -1
      }
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