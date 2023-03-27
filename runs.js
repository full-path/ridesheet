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

function sendRequestForRuns() {
  try {
    let grid = [["Last Updated:", new Date(), new Date(), null]]
    let currentRow = 2
    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasRuns) {
        let params = {resource: "runs"}
        let response = getResource(endPoint, params)
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
        } catch(e) {
          logError(e)
          responseObject = {status: "LOCAL_ERROR::" + e.name}
        }

        let formatGroups = {
          header: {
            ranges: ["A1:D1"],
            formats: function(rl) {
              rl.setBackground(headerBackgroundColor)
              rl.setFontWeight("bold")
              rl.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
            }
          },
          date: {
            ranges: ["C1"],
            formats: function(rl) {
              rl.setNumberFormat("mm/dd/yyyy")
              rl.setHorizontalAlignment("left")
            }
          },
          time: {
            ranges: ["B1"],
            formats: function(rl) {
              rl.setNumberFormat("h:mm am/pm")
              rl.setFontWeight("bold")
            }
          },
          runAttributes: {
            ranges: [],
            formats: function(rl) {
              rl.setBackground(headerBackgroundColor)
              rl.setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
            }
          },
          label: {
            ranges: [],
            formats: function(rl) {
              rl.setFontWeight("bold")
            }
          },
          seat: {
            ranges: [],
            formats: function(rl) {
              rl.setNumberFormat("0")
              rl.setHorizontalAlignment("left")
            }
          },
          availableSeats: {
            ranges: [],
            formats: function(rl) {
              rl.setNumberFormat("0")
            }
          },
          mergeCells: {
            ranges: [],
            formats: function(rl) {
              rl.getRanges().forEach(range => range.merge())
            }
          }
        }

        if (responseObject.status !== "OK") {
          grid.push(
            [null, null, null, null],
            [responseObject.status, null, null, null]
          )
          const thisRange = "A" + (currentRow + 1) + ":D" + (currentRow + 1)
          formatGroups.header.ranges.push(thisRange)
          formatGroups.mergeCells.ranges.push(thisRange)
          currentRow += 2
        } else if (responseObject.results && responseObject.results.length) {
          responseObject.results.forEach(row => {
            grid.push(
              [null, null, null, null],
              ["Source", endPoint.name, "Run Date", new Date(row.runDate)],
              ["Seats", row.ambulatorySpacePoints ? row.ambulatorySpacePoints : "Unknown", "Wheelchair spaces", row.standardWheelchairSpacePoints ? row.standardWheelchairSpacePoints : "Unknown"],
              ["Lift", row.hasLift ? "Yes" : "No", "Ramp", row.hasRamp ? "Yes" : "No"],
              ["Start location", row.startLocation, "End location", row.endLocation],
              [null, "Stop Time", "Stop City", "Available Seats"]
            )
            formatGroups.label.ranges.push("A" + (currentRow + 1) + ":A" + (currentRow + 4), "C" + (currentRow + 1) + ":C" + (currentRow + 4))
            formatGroups.date.ranges.push("D" + (currentRow + 1))
            formatGroups.seat.ranges.push("B" + (currentRow + 2), "D" + (currentRow + 2))
            formatGroups.runAttributes.ranges.push("A" + (currentRow + 1) + ":D" + (currentRow + 5))
            formatGroups.header.ranges.push("A" + (currentRow + 5) + ":D" + (currentRow + 5))
            currentRow += 6
            let currentSeats = row.ambulatorySpacePoints ? parseInt(row.ambulatorySpacePoints) : 0
            row.stops.forEach(stop => {
              let riderChange = parseInt(stop.riderChange)
              currentSeats = row.ambulatorySpacePoints ? (currentSeats - riderChange) : ''
              let text = getStopText(riderChange)
              grid.push([text, new Date(stop.time), stop.city, currentSeats])
            })
            formatGroups.time.ranges.push("B" + currentRow + ":B" + (currentRow + row.stops.length -1))
            formatGroups.availableSeats.ranges.push("D" + currentRow + ":D" + (currentRow + row.stops.length -1))
            currentRow += row.stops.length
          })
        } else {
          grid.push(
            [null, null, null, null],
            [endPoint.name + " responded with no runs", null, null, null]
          )
          const thisRange = "A" + (currentRow + 1) + ":D" + (currentRow + 1)
          formatGroups.header.ranges.push(thisRange)
          formatGroups.mergeCells.ranges.push(thisRange)
          currentRow += 2
        }
       const ss = SpreadsheetApp.getActiveSpreadsheet()
       const runSheet = ss.getSheetByName("Outside Runs") || ss.insertSheet("Outside Runs")
       runSheet.getDataRange().clear().breakApart()
       let range = runSheet.getRange(1,1,grid.length,grid[0].length)
       range.setValues(grid)
       range.clearFormat()
       applyFormats(formatGroups, runSheet)
       runSheet.setFrozenRows(1)
       runSheet.autoResizeColumns(1,grid[0].length)
      }
    })
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

function receiveRequestForRuns() {
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
        return  t["Trip Date"].getTime()  === runIn["Run Date"].getTime() && 
                t["Vehicle ID"] === runIn["Vehicle ID"] && 
                t["Driver ID"]  === runIn["Driver ID"] 
      })

      runOut.runDate = runIn["Run Date"]    
      runOut.standardWheelchairSpacePoints = vehicle["Wheelchair Capacity"]
      runOut.ambulatorySpacePoints = vehicle["Seating Capacity"]
      runOut.hasLift = !!vehicle["Has Lift"]
      runOut.hasRamp = !!vehicle["Has Ramp"]
      runOut.startLocation = runIn["Start Location"]
      runOut.endLocation = runIn["End Location"]
      let stops = []
      runTrips.forEach(trip => {
        let puStop = {}
        puStop.time = new Date(trip["Trip Date"].getTime() + timeOnlyAsMilliseconds(trip["PU Time"]))
        puStop.city = extractCity(trip["PU Address"])
        puStop.riderChange =  1 + trip["Guests"]
        mergeStop(puStop, stops)

        let doStop = {}
        doStop.time = new Date(trip["Trip Date"].getTime() + timeOnlyAsMilliseconds(trip["DO Time"]))
        doStop.city = extractCity(trip["DO Address"])
        doStop.riderChange = -1 - trip["Guests"]
        mergeStop(doStop, stops)
      })
      stops.sort((a, b) => a.time.getTime() - b.time.getTime())
      runOut.stops = stops
      return runOut
    })
    result.sort((a, b) => a.runDate.getTime() - b.runDate.getTime())
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