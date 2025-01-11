function updateRuns(e) {
  try {
    let newRun
    let runsRange
    let runsData
    let ss = SpreadsheetApp.getActiveSpreadsheet()

    // Gather all needed trip data
    let trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange()).filter(tripRow => {
      return tripRow["Trip Date"] && tripRow["Driver ID"] && tripRow["Vehicle ID"]
    })
    // Gather all run data
    let runsSheet = ss.getSheetByName("Runs")
    let runsLastRow = runsSheet.getLastRow()
    let runsLastColumn = runsSheet.getLastColumn()
    if (runsLastRow > 1) {
      runsRange = runsSheet.getDataRange()
      runsData = getRangeValuesAsTable(runsRange)
    } else {
      runsData = []
    }

    // Iterate through trip data, associating it with the matching run row
    // Using a map here rather than an object, as the map allows us to use the entire
    // run row as a key.
    // The result here is two maps: one for runs already present, and one for runs that need to be
    // added
    let runsOut = new Map()
    let newRunsOut = new Map()
    let uniqueRuns = []
    runsData.forEach(runRow => {
      let runFingerprint = JSON.stringify(runRow["Run Date"]) + runRow["Driver ID"] + runRow["Vehicle ID"]
      if (!uniqueRuns.includes(runFingerprint)) {
        uniqueRuns.push(runFingerprint)
        runsOut.set(runRow,[])
      }
    })
    trips.forEach((tripRow, tripIndex) => {
      let found = false
      runsOut.forEach((value, runRow) => {
        if (tripRow["Trip Date"]  >=  runRow["Run Date"] &&
            tripRow["Trip Date"]  <=  runRow["Run Date"] &&
            tripRow["Driver ID"]  === runRow["Driver ID"] &&
            tripRow["Vehicle ID"] === runRow["Vehicle ID"]) {
          found = true
          runsOut.get(runRow).push(tripRow)
        }
      })
      if (!found) {
        newRunsOut.forEach((value, runRow) => {
          if (tripRow["Trip Date"]  >=  runRow["Run Date"] &&
              tripRow["Trip Date"]  <=  runRow["Run Date"] &&
              tripRow["Driver ID"]  === runRow["Driver ID"] &&
              tripRow["Vehicle ID"] === runRow["Vehicle ID"]) {
            found = true
            newRunsOut.get(runRow).push(tripRow)
          }
        })
      }
      if (!found) {
        newRun = {}
        newRun["Run Date"] = tripRow["Trip Date"]
        newRun["Driver ID"] = tripRow["Driver ID"]
        newRun["Vehicle ID"] = tripRow["Vehicle ID"]
        newRunsOut.set(newRun,[tripRow])
      }
    })

    // Merge together the existing and new runs, sorting the merged data set by Run Date
    const existingRuns = updateRunDetails(runsOut)
    const newRuns = updateRunDetails(newRunsOut)
    let runsToSave = [...existingRuns, ...newRuns].sort((a,b) => {
      return a["Run Date"] - b["Run Date"]
    })

    // If the resulting run list is longer than the original one, get a larger range.
    // If the resulting run list is shorter than the original one, clear out the values in the last rows
    const countOfRowsUnfilled = runsData.length - runsToSave.length
    let runsRangeOut = runsRange
    if (countOfRowsUnfilled < 0) {
      runsRangeOut = runsSheet.getRange(1, 1, runsToSave.length + 1, runsLastColumn)
    } else if (countOfRowsUnfilled > 0) {
      runsToSave = [...runsToSave, ...Array(countOfRowsUnfilled).fill({})]
    }
    // log("runsToSave",JSON.stringify(runsToSave))
    // log("runsRangeOut",runsRangeOut.getA1Notation())
    if (runsRangeOut) setValuesByHeaderNames(runsToSave,runsRangeOut,{overwriteAll: true})
  } catch(e) { logError(e) }
}

function updateRunDetails(runsMap) {
  try {
    let runsArray = Array.from(runsMap.keys()).filter(runRow => runsMap.get(runRow).length > 0)
    runsArray.forEach(run => {
      let tripsArray = runsMap.get(run).filter(trip => trip["PU Time"])
      if (tripsArray.length === 0) {
        run["First PU Time"] = null
      } else {
        run["First PU Time"] = tripsArray.reduce((min, p) => p["PU Time"] < min ? p["PU Time"] : min, tripsArray[0]["PU Time"])
      }
    })

    runsArray.forEach(run => {
      let tripsArray = runsMap.get(run).filter(trip => trip["DO Time"])
      if (tripsArray.length === 0) {
        run["Last DO Time"] = null
      } else {
        run["Last DO Time"] = tripsArray.reduce((max, p) => p["DO Time"] > max ? p["DO Time"] : max, tripsArray[0]["DO Time"])
      }
    })
    return runsArray
  } catch(e) { logError(e) }
}

function getStopText(riderChange) {
  if (riderChange > 0) {
    if (riderChange > 1) return riderChange + ' riders board'
    return '1 rider boards'
  } else {
    if (riderChange < -1) return Math.abs(riderChange) + ' riders alight'
    return '1 rider alights'
  }
}

function mergeStop(newStop, runStops) {
  let matchingStop = runStops.find(s => s.time === newStop.time && s.city === newStop.city)
  if (matchingStop) {
    matchingStop.riderChange = matchingStop.riderChange + newStop.riderChange
  } else {
    runStops.push(newStop)
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
  let runDate = runs.reduce((latest, run) => 
    run["Run Date"] > latest ? run["Run Date"] : latest, 
    runs[0]["Run Date"]
  )
  runDate.setDate(runDate.getDate() + 1) // Add one day to the latest run date
  
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