/**
 * Updates run details based on associated trips and returns array of runs
 * @param {Object} runsObject - Object with run data in format {tripKey: {run: {...}, trips: [...]}}
 * @returns {Array} Array of run objects with calculated details
 */
function updateRunDetails(runsObject) {
  try {
    for (let tripKey in runsObject) {
      let runEntry = runsObject[tripKey]
      let tripsArray = runEntry.trips

      // Calculate First PU Time
      let puTrips = tripsArray.filter(trip => trip["PU Time"])
      if (puTrips.length === 0) {
        runEntry.run["First PU Time"] = null
        runEntry.run["Scheduled Start Time"] = null
      } else {
        runEntry.run["First PU Time"] = puTrips.reduce(
          (min, trip) => trip["PU Time"] < min ? trip["PU Time"] : min, 
          puTrips[0]["PU Time"]
        )
        runEntry.run["Scheduled Start Time"] = runEntry.run["First PU Time"]
      }

      // Calculate Last DO Time
      let doTrips = tripsArray.filter(trip => trip["DO Time"])
      if (doTrips.length === 0) {
        runEntry.run["Last DO Time"] = null
        runEntry.run["Scheduled End Time"] = null
      } else {
        runEntry.run["Last DO Time"] = doTrips.reduce(
          (max, trip) => trip["DO Time"] > max ? trip["DO Time"] : max, 
          doTrips[0]["DO Time"]
        )
        runEntry.run["Scheduled End Time"] = runEntry.run["Last DO Time"]
      }
      runEntry.run["Review TS"] = new Date()
    }

    return Object.values(runsObject).map(entry => entry.run)

  } catch(e) { 
    logError(e)
    throw new Error(`Failed to update run details: ${e.message}`)
  }
}

// Fill in a week of entries in the "Runs" sheet using the schedule information in "Run Template"
function buildRunsFromTemplate() {
  buildRecordsFromTemplate("Run Template","Runs","Run Date")
}

/**
 * Creates run entries in the Run Review sheet based on trips being moved to review
 * @param {Array} trips - Array of trip objects being moved to review
 * @returns {void}
 */
function createRunsInReview(trips) {
  try {
    let ss = SpreadsheetApp.getActiveSpreadsheet()
    let runReviewSheet = ss.getSheetByName("Run Review")
    
    let newRunsOut = {}
    
    trips.forEach(tripRow => {
      let tripKey = JSON.stringify(tripRow["Trip Date"]) + 
                    tripRow["Driver ID"] + 
                    tripRow["Vehicle ID"]
      
      if (tripKey in newRunsOut) {
        newRunsOut[tripKey].trips.push(tripRow)
      } else {
        let newRun = {
          "Run Date": tripRow["Trip Date"],
          "Driver ID": tripRow["Driver ID"],
          "Vehicle ID": tripRow["Vehicle ID"]
        }
        newRunsOut[tripKey] = {
          run: newRun,
          trips: [tripRow]
        }
      }
    })

    const runsToCreate = updateRunDetails(newRunsOut)
      .sort((a,b) => a["Run Date"] - b["Run Date"])

    if (runsToCreate.length > 0) {
      runsToCreate.forEach(run => {
        createRow(runReviewSheet, run)
      })
      
      const lastRow = runReviewSheet.getLastRow()
      const startRow = lastRow - runsToCreate.length + 1
      applySheetFormatsAndValidation(runReviewSheet, startRow)
    }

  } catch(e) { 
    logError(e)
    throw new Error(`Failed to create runs in review: ${e.message}`)
  }
}