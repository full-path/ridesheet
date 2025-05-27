function refreshDispatchSheetLocal() {
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()

    const dispatchSheet       = ss.getSheetByName("Dispatch")
    const tripReviewSheet = ss.getSheetByName("Trip Review")
    const dispatchFilter      = function(row) { return row["Trip Date"] && row["Trip Date"] < dateToday() }
    moveRows(dispatchSheet, tripReviewSheet, dispatchFilter, "Review TS")

    const runSheet        = ss.getSheetByName("Runs")
    const runReviewSheet  = ss.getSheetByName("Run Review")
    const runFilter       = function(row) { return row["Run Date"] && row["Run Date"] < dateToday() }
    moveRows(runSheet, runReviewSheet, runFilter, "Review TS")

    const tripSheet       = ss.getSheetByName("Trips")
    const tripFilter      = function(row) { 
      return row["Trip Date"] && row["Trip Date"].getTime() === dateToday().getTime() 
    }
    moveRows(tripSheet, dispatchSheet, tripFilter, undefined)
  } catch(e) { logError(e) }
}