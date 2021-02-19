
function presentCalendarTrigger() {
  const calendarTrigger = getCalendarTrigger()
  const ui = SpreadsheetApp.getUi()
  if (calendarTrigger) {
    const response = ui.alert("Updates on","Schedule calendar updates are running. Keep running updates?",ui.ButtonSet.YES_NO)
    if (response === ui.Button.NO) {
      ScriptApp.deleteTrigger(calendarTrigger)
      ui.alert("Success","Scheduled updates stopped.",ui.ButtonSet.OK)
    } else {
      ui.alert("No change","Scheduled updates still running.",ui.ButtonSet.OK)
    }
  } else {
    const response = ui.alert("Updates off","Schedule calendar updates are not running. Start running updates?",ui.ButtonSet.YES_NO)
    if (response === ui.Button.YES) {
      ScriptApp.newTrigger("updateDriverCalendars").timeBased().nearMinute(0).everyHours(1).create()
      ui.alert("Success","Scheduled updates started.",ui.ButtonSet.OK)
    } else {
      ui.alert("No change","Scheduled updates not running.",ui.ButtonSet.OK)
    }
  }
}

function getCalendarTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers()
  return allTriggers.find(trigger => {
     return (trigger.getHandlerFunction() === "updateDriverCalendars" && trigger.getEventType() === ScriptApp.EventType.CLOCK) 
  })
}