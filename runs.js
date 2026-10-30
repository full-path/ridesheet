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
  const weekday = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"]
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let runsSheet = ss.getSheetByName("Runs") 
  let runTemplateSheet = ss.getSheetByName("Run Template")
  let runs = getRangeValuesAsTable(runsSheet.getDataRange())
  let runTemplates = getRangeValuesAsTable(runTemplateSheet.getDataRange())
  let runDateRaw = runs.reduce((latest, run) => 
    run["Run Date"] > latest ? run["Run Date"] : latest, 
    runs[0]["Run Date"]
  )

  let runDate
  if (!runDateRaw) {
    runDate = new Date()
  } else {
    runDate = new Date(runDateRaw)
  }

  runDate.setDate(runDate.getDate() + 1) 
  
  const ui = SpreadsheetApp.getUi()
  const response = ui.alert(
    'Generate Runs', 
    'Would you like to generate runs starting from ' + formatDate(runDate) + '? Click No to enter a custom date.',
    ui.ButtonSet.YES_NO_CANCEL
  )

  let startDate
  if (response == ui.Button.YES) {
    startDate = runDate
  } else if (response == ui.Button.NO) {
    const promptResult = ui.prompt(
      'Enter Start Date',
      'Enter the date to start generating runs from (MM/DD/YYYY):',
      ui.ButtonSet.OK_CANCEL
    )
    if (promptResult.getSelectedButton() == ui.Button.OK) {
      startDate = parseDate(promptResult.getResponseText())
      if (!isValidDate(startDate)) {
        ui.alert('Invalid date entered. Operation cancelled.')
        return
      }
    } else {
      ss.toast('Action cancelled')
      return
    }
  } else {
    ss.toast('Action cancelled') 
    return
  }

  const lastRow = runsSheet.getLastRow()

  for (let i = 0; i < 7; i++) {
    let currentDate = new Date(startDate)
    currentDate.setDate(startDate.getDate() + i)
    let currentDayOfWeek = weekday[currentDate.getDay()]
    
    const matchingTemplates = runTemplates.filter(row => 
      row["Days of Week"] && row["Days of Week"].includes(currentDayOfWeek)
    ).sort((a,b) => {
      return a["Scheduled Start Time"] - b["Scheduled Start Time"]
    })

    matchingTemplates.forEach(template => {
      const newRun = {
        "Run Date": formatDate(currentDate),
        "Driver ID": template["Driver ID"],
        "Vehicle ID": template["Vehicle ID"],
        "Scheduled Start Time": template["Scheduled Start Time"],
        "Scheduled End Time": template["Scheduled End Time"]
      }
      createRow(runsSheet, newRun)
    })
  }
  applySheetFormatsAndValidation(runsSheet, lastRow + 1)
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