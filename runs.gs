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
    let tripsArray = runsMap.get(run).filter(trip => trip["PU Time"])
    if (tripsArray.length === 0) {
      run["First PU Time"] = null
    } else if (tripsArray.length === 1) {
      //log(6, JSON.stringify(tripsArray))
      run["First PU Time"] = tripsArray[0]["PU Time"]
    } else {
      //log(7, JSON.stringify(tripsArray))
      run["First PU Time"] = tripsArray.reduce((min, p) => p["PU Time"] < min ? p["PU Time"] : min, tripsArray[0]["PU Time"])
    }
  })
  
  runsArray.forEach(run => {
    let tripsArray = runsMap.get(run).filter(trip => trip["DO Time"])
    if (tripsArray.length === 0) {
      run["Last DO Time"] = null
    } else if (tripsArray.length === 1) {
      //log(8, JSON.stringify(tripsArray))
      run["Last DO Time"] = tripsArray[0]["DO Time"]
    } else {
      //log(9, JSON.stringify(tripsArray))
      run["Last DO Time"] = tripsArray.reduce((max, p) => p["DO Time"] > max ? p["DO Time"] : max, tripsArray[0]["DO Time"])
    }
  })
  return runsArray
}

