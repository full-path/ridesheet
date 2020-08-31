function updateRuns(e) {
  let newRun
  let runsRange
  let runsData
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange()).filter(tripRow => {
    return tripRow["Trip Date"] && tripRow["Driver ID"] && tripRow["Vehicle ID"]
  })
  let runsSheet = ss.getSheetByName("Runs")
  let runsLastRow = runsSheet.getLastRow()
  if (runsLastRow > 1) {
    runsRange = runsSheet.getRange(2, 1, runsLastRow-1, runsSheet.getLastColumn())
    runsData = getRangeValuesAsTable(runsRange)
  } else {
    runsData = []
  }
  let runsMap = new Map()
  let newRunsMap = new Map()
  runsData.forEach(runRow => runsMap.set(runRow,[]))
  trips.forEach((tripRow, tripIndex) => {
    let found = false
    runsMap.forEach((value, runRow) => {
      if (tripRow["Trip Date"]  >=  runRow["Run Date"] &&
          tripRow["Trip Date"]  <=  runRow["Run Date"] &&
          tripRow["Driver ID"]  === runRow["Driver ID"] &&
          tripRow["Vehicle ID"] === runRow["Vehicle ID"]) {
        found = true
        runRow["Test"] = runRow["Test"] + 1
        runsMap.get(runRow).push(tripRow)
      }
    })
    if (!found) {
      newRunsMap.forEach((value, runRow) => {
        if (tripRow["Trip Date"]  >=  runRow["Run Date"] &&
            tripRow["Trip Date"]  <=  runRow["Run Date"] &&
            tripRow["Driver ID"]  === runRow["Driver ID"] &&
            tripRow["Vehicle ID"] === runRow["Vehicle ID"]) {
          log(2)
          found = true
          newRunsMap.get(runRow).push(tripRow)
        }
      })
    }
    if (!found) {
      newRun = {}
      newRun["Run Date"] = tripRow["Trip Date"]
      newRun["Driver ID"] = tripRow["Driver ID"]
      newRun["Vehicle ID"] = tripRow["Vehicle ID"]
      newRun["Test"] = 0
      newRunsMap.set(newRun,[tripRow])
    }
  })
  if (runsMap.size) {
    let existingRuns = updateRunDetails(runsMap)
    setValuesByHeaderNames(existingRuns, runsRange)
  }
  if (newRunsMap.size) {
    let newRuns = updateRunDetails(newRunsMap)
    appendValuesByHeaderNames(newRuns, runsSheet)
  }
}

function updateRunDetails(runsMap) {
  let runsArray = Array.from(runsMap.keys())
  runsArray.forEach(run => {
    if (runsMap.get(run).length === 1) {
      log(6, JSON.stringify(runsMap.get(run)))
      run["First PU Time"] = runsMap.get(run)[0]["PU Time"]
      run["Last DO Time"] = runsMap.get(run)[0]["DO Time"]
    } else if (runsMap.get(run).length > 1) {
      log(7, JSON.stringify(runsMap.get(run)))
      run["First PU Time"] = runsMap.get(run).reduce((a,b) => {return new Date(Math.min(a["PU Time"], b["PU Time"]))})
      run["Last DO Time"] = runsMap.get(run).reduce((a,b) => {return new Date(Math.max(a["DO Time"], b["DO Time"]))})
    } else {
      run["First PU Time"] = null
      run["Last DO Time"] = null
    }
  })
  return runsArray
}

