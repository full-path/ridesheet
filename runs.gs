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
  try {
    let runsArray = Array.from(runsMap.keys())
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
  } catch(e) {
    logError(e)
  }
}

function shareRuns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const runs = getRangeValuesAsTable(ss.getSheetByName("Runs").getDataRange()).filter(runRow => {
      return runRow["Run Date"] >= dateToday()
    })
    const trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange()).filter(tripRow => {
      return tripRow["Trip Date"] >= dateToday() && tripRow["Driver ID"] && tripRow["Vehicle ID"]
    })
    const vehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())

    let result = runs.map(runIn => {
      let runOut = {}
      let vehicle = vehicles.find(v => v["Vehicle ID"] === runIn["Vehicle ID"]) || {}
      let runTrips = trips.filter(t => {
                           return t["Trip Date"].getTime()  === runIn["Run Date"].getTime() && 
                                  t["Vehicle ID"] === runIn["Vehicle ID"] && 
                                  t["Driver ID"]  === runIn["Driver ID"] 
                                  })

      runOut.runDate = runIn["Run Date"]    
      runOut.ambulatorySpacePoints = vehicle["Seating Capacity"]
      runOut.standardWheelchairSpacePoints = vehicle["Seating Capacity"]
      runOut.ambulatorySpacePoints = vehicle["Seating Capacity"]
      runOut.hasLift = !!vehicle["Has Lift"]
      runOut.hasRamp = !!vehicle["Has Ramp"]
      runOut.startLocation = runIn["Start Location"]
      runOut.endLocation = runIn["End Location"]
      let stops = []
      runTrips.forEach(trip => {
        let puStop = {}
        puStop.time = new Date(trip["Trip Date"].getTime() + timeOnly(trip["PU Time"]))
        puStop.city = extractCity(trip["PU Address"])
        puStop.riderChange =  1 + trip["Guests"]
        mergeStop(puStop, stops)

        let doStop = {}
        doStop.time = new Date(trip["Trip Date"].getTime() + timeOnly(trip["DO Time"]))
        doStop.city = extractCity(trip["DO Address"])
        doStop.riderChange = -1 - trip["Guests"]
        mergeStop(doStop, stops)
      })
      stops.sort((a, b) => a.time.getTime() - b.time.getTime())
      runOut.stops = stops
      return runOut
    })
    result.sort((a, b) => a.runDate.getTime() - b.runDate.getTime())
    log(JSON.stringify(result))
    return result
  } catch(e) {
    logError(e)
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