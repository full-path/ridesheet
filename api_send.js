function getRuns() {
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
          riderChange: {
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
              ["Seats", row.ambulatorySpacePoints, "Wheelchair spaces", row.standardWheelchairSpacePoints],
              ["Lift", row.hasLift ? "Yes" : "No", "Ramp", row.hasLift ? "Yes" : "No"],
              ["Start location", row.startLocation, "End location", row.endLocation],
              [null, "Stop Time", "Stop City", "Rider Change"]
            )
            formatGroups.label.ranges.push("A" + (currentRow + 1) + ":A" + (currentRow + 4), "C" + (currentRow + 1) + ":C" + (currentRow + 4))
            formatGroups.date.ranges.push("D" + (currentRow + 1))
            formatGroups.seat.ranges.push("B" + (currentRow + 2), "D" + (currentRow + 2))
            formatGroups.runAttributes.ranges.push("A" + (currentRow + 1) + ":D" + (currentRow + 5))
            formatGroups.header.ranges.push("A" + (currentRow + 5) + ":D" + (currentRow + 5))
            currentRow += 6
            row.stops.forEach(stop => {
              grid.push([null, new Date(stop.time), stop.city, stop.riderChange])
            })
            formatGroups.time.ranges.push("B" + currentRow + ":B" + (currentRow + row.stops.length -1))
            formatGroups.riderChange.ranges.push("D" + currentRow + ":D" + (currentRow + row.stops.length -1))
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

function getTrips() {
  try {
    const lastColumnLetter = "R"
    const headers = [
      "Scheduled PU Time",
      "Decline",
      "Claim",
      "Source",
      "Trip Date",
      "Earliest PU Time",
      "Requested PU Time",
      "Latest PU Time",
      "Requested DO Time",
      "Appt Time",
      "PU Address",
      "DO Address",
      "Guests",
      "Mobility Factors",
      "Notes",
      "Est Hours",
      "Est Miles",
      "Trip ID"
    ]
    let grid = [
      ["Last Updated:", new Date(), new Date()].concat(Array(headers.length-3).fill(null)),
      headers
    ]
    let currentRow = 3
    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasTrips) {
        let params = {resource: "tripRequests"}
        let response = getResource(endPoint, params)
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
        } catch(e) {
          logError(e)
          responseObject = {status: "LOCAL_ERROR:" + e.name}
        }
        let formatGroups = {
          mergeCells: {
            ranges: [],
            formats: function(rl) {
              rl.getRanges().forEach(range => range.merge())
            }
          },       
          header: {
            ranges: ["A1:" + lastColumnLetter + "2"],
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
            }
          },
          duration: {
            ranges: [],
            formats: function(rl) {
              rl.setNumberFormat("h:mm")
            }
          },
          distance: {
            ranges: [],
            formats: function(rl) {
              rl.setNumberFormat("0.00")
            }
          },
          integer: {
            ranges: [],
            formats: function(rl) {
              rl.setNumberFormat("0")
              rl.setHorizontalAlignment("right")
            }
          },
          checkbox: {
            ranges: [],
            formats: function(rl) {
              rl.insertCheckboxes()
            }
          }
        }
        if (responseObject.status !== "OK") {
          grid.push(
            [responseObject.status].concat(Array(headers.length-1).fill(null))
          )
          const thisRange = "D" + currentRow + ":" + lastColumnLetter + currentRow
          formatGroups.mergeCells.ranges.push(thisRange)
          currentRow += 1
        } else if (responseObject.results && responseObject.results.length) {
          responseObject.results.forEach(item => {
            const row = item.tripRequest
            const openAttributes = JSON.parse(row["@openAttribute"])
            grid.push([
              null,
              false,
              false,
              endPoint.name,
              new Date(row.pickupTime["@time"]),
              row.pickupWindowStartTime ? new Date(row.pickupWindowStartTime["@time"]) : null,
              new Date(row.pickupTime["@time"]),
              row.pickupWindowEndTime ? new Date(row.pickupWindowEndTime["@time"]) : null,
              new Date(row.dropoffTime["@time"]),
              row.appointmentTime ? new Date(row.appointmentTime["@time"]) : null,
              buildAddressFromSpec(row.pickupAddress),
              buildAddressFromSpec(row.dropoffAddress),
              openAttributes.guestCount,
              openAttributes.mobilityFactors,
              openAttributes.notes,
              openAttributes.estimatedTripDurationInSeconds / 86400,
              openAttributes.estimatedTripDistanceInMiles,
              openAttributes.tripTicketId
            ])
          })
          formatGroups.checkbox.ranges.push("B" + currentRow + ":C" + (currentRow + responseObject.results.length - 1))
          formatGroups.date.ranges.push("E" + currentRow + ":E" + (currentRow + responseObject.results.length - 1))
          formatGroups.time.ranges.push("F" + currentRow + ":I" + (currentRow + responseObject.results.length - 1))
          formatGroups.integer.ranges.push("M" + currentRow + ":M" + (currentRow + responseObject.results.length - 1))
          formatGroups.duration.ranges.push("P" + currentRow + ":P" + (currentRow + responseObject.results.length - 1))
          formatGroups.distance.ranges.push("Q" + currentRow + ":Q" + (currentRow + responseObject.results.length - 1))
          currentRow += responseObject.results.length

          const ss = SpreadsheetApp.getActiveSpreadsheet()
          const sheet = ss.getSheetByName("Outside Trips") || ss.insertSheet("Outside Trips")
          sheet.getDataRange().clear().breakApart()
          let range = sheet.getRange(1, 1, grid.length, grid[0].length)
          range.clearFormat()
          range.setValues(grid)
          applyFormats(formatGroups, sheet)
          sheet.setFrozenRows(2)
          sheet.setFrozenColumns(3)
          sheet.autoResizeColumns(1,grid[0].length)
        } else {
          grid.push(
            [endPoint.name + " responded with no trip requests"].concat(Array(14).fill(null))
          )
          const thisRange = "A" + currentRow + ":" + lastColumnLetter + currentRow
          formatGroups.mergeCells.ranges.push(thisRange)
          currentRow += 1
        }
      }
    })
  } catch(e) { logError(e) }
}