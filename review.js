function addDeadheadDataToRunsForDate(date) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const vehicleSheet = ss.getSheetByName("Vehicles")
    const vehicles = getRangeValuesAsTable(vehicleSheet.getDataRange())
    const runsSheet = ss.getSheetByName("Run Review")
    let runs = getRangeValuesAsTable(runsSheet.getDataRange())
    let runsThisDay = runs.filter( (row) => {
      return isInDay(row["Run Date"], date)
    })
    const tripSheet = ss.getSheetByName("Trip Review")
    const trips = getRangeValuesAsTable(tripSheet.getDataRange())
    const tripReviewCompletedTripResults = getDocProp("tripReviewCompletedTripResults")
    const tripsThisDay = trips.filter( (row) => {
      return isInDay(row["Trip Date"], date) && 
        tripReviewCompletedTripResults.includes(row["Trip Result"])
    })

    const dataErrors = preDeadheadDataErrors(tripsThisDay, runsThisDay)
    if (dataErrors) {
      SpreadsheetApp.getUi().alert(
        "Deadhead data could not be added to the due to the following error(s):\n\n" +
        dataErrors.filter(msg => msg.length > 0).join("\n\n")
      )
      // TODO: notify user
    } else {
      runsThisDay.forEach((run) => {
        const deadheadData = getDeadheadDataForRun(
          tripsThisDay, vehicles, run["Driver ID"], run["Vehicle ID"], run["Run ID"]
        )
        run = {...run, ...deadheadData}
      })
      setValuesByHeaderNames(runs, runsSheet.getDataRange())
    }
  } catch(e) { logError(e) }
}

function getDeadheadDataForRun(tripsThisDay, vehicles, driverId, vehicleId, runId) {
  const tripsThisRun = tripsThisDay.
    filter((row) => 
      row["Driver ID"] === driverId &&
      row["Vehicle ID"] === vehicleId &&
      row["Run ID"] === runID
    )
  const firstTrip = tripsThisRun.
    reduce((earliestTrip, row) => 
      timeOnlyAsMilliseconds(row["PU Time"]) < timeOnlyAsMilliseconds(earliestTrip["PU Time"]) ? 
      row : earliestTrip
    )
  const lastTrip = tripsThisRun.
    reduce((latestRow, row) => 
      timeOnlyAsMilliseconds(row["PU Time"]) > timeOnlyAsMilliseconds(latestRow["PU Time"]) ? 
      row : latestRow
    )
  const vehicle = vehicles.find((row) => row["Vehicle ID"] === vehicleId)

  result = {}
  result["First PU Address"] = parseAddress(firstTrip["PU Address"]).geocodeAddress
  result["Last DO Address"] = parseAddress(lastTrip["DO Address"]).geocodeAddress
  result["Vehicle Garage Address"] = parseAddress(vehicle["Garage Address"]).geocodeAddress
  result["Starting Deadhead Distance"] = 
      getTripEstimate(result["Vehicle Garage Address"], result["First PU Address"], "miles")
  result["Ending Deadhead Distance"] = 
      getTripEstimate(result["Last DO Address"], result["Vehicle Garage Address"], "miles")
  return result
}

function preDeadheadDataErrors(tripsThisDay, runsThisDay) {
  return [
    hasDuplicateRuns(runsThisDay),
    hasOrphans(tripsThisDay, runsThisDay),
    hasIncompleteTrips(tripsThisDay),
    hasOnlyCompleteRuns(runsThisDay)
  ]
}

function hasDuplicateRuns(runsThisDay) {
  const runKeys = runsThisDay.map((row) => getRunKey(row))
  if (new Set(runKeys).size !== runKeys.length) {
    return "There are duplicate runs for this day."
  } else {
    return
  }
}

function hasOrphans(tripsThisDay, runsThisDay) {
  const runKeys = runsThisDay.map((row) => getRunKey(row))
  const tripKeys = tripsThisDay.map((row) => getRunKey(row))
  let runKeyErrors = []
  let tripKeyErrors = []
  let runErrorMessage = ""
  let tripErrorMessage = ""
  runKeys.forEach((runKey) => {
    if (!tripKeys.indexOf(tripKey) === -1) runKeyErrors.push(runKey)
  })
  tripKeys.forEach((tripKey) => {
    if (runKeys.indexOf(tripKey) === -1) tripKeyErrors.push(runKey)
  })
  if (runKeyErrors.length) {
    runErrorMessage = (runKeyErrors.length === 1 ? "1 run" : runKeyErrors.length + " runs") +
      " with no matching trips:\n" + 
      runKeyErrors.join("\n")
  }
  if (tripKeyErrors.length) {
    tripErrorMessage = (tripKeyErrors.length === 1 ? "1 trip" : tripKeyErrors.length + " trips") +
      " with no matching runs:\n" + 
      tripKeyErrors.join("\n")
  }
  return [runErrorMessage, tripErrorMessage].filter(msg => msg.length > 0).join("\n\n")
}

function hasIncompleteTrips(tripsThisDay) {
  const incompleteTrips = tripsThisDay.filter(!isReviewedTrip)
  if (incompleteTrips.length) {
    return "There are " + incompleteTrips.length + " trips with incomplete data."
  } else {
    return ""
  }
}

function hasOnlyCompleteRuns(runsThisDay) {
  const incompleteRuns = runsThisDay.filter(!isUserReviewedRun)
  if (incompleteRuns.length) {
    return "There are " + incompleteRuns.length + " runs with incomplete data."
  } else {
    return ""
  }
}

function getRunKey(runOrTrip) {
  return [
    "Driver ID: " + runOrTrip["Driver ID"],
    "Vehicle ID: " + runOrTrip["Vehicle ID"],
    "Run ID: " + (runOrTrip["Run ID"] ? runOrTrip["Run ID"] : "<Blank>")
  ].join()
}