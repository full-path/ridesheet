function createEvent(calendarID) {
  const calendar = CalendarApp.getCalendarById("fullpath.io_vlmbs3tfapdnpmnn6a9vsjfdrg@group.calendar.google.com")
  //event = calendar.createEvent("My test event", new Date, timeAdd(new Date, 3600000))
  let event = calendar.getEventById("a8af0j0t32cbq5ba36b6me4jfo@google.com")
  event.setTime(timeAdd(event.getStartTime(), 3600000), timeAdd(event.getEndTime(), 3600000))
  log(event.getId())
}
