
function getRuns() {
  try {
    let runArray = [["Last Updated:", new Date(), new Date(), null]]
    let currentRow = 2
    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasRuns) {
        let params = {}
        params.resource = "runs"
        params.version = endPoint.version
        params.apiKey = endPoint.apiKey
        let hmac = {}
        hmac.nonce = Utilities.getUuid()
        hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
        hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)
        
        const options = {method: 'GET', contentType: 'application/json'}
        let response = UrlFetchApp.fetch(endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac), options)
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
              rl.setNumberFormat("mm-dd-yyyy")
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
          runArray.push(
            [null, null, null, null],
            [responseObject.status, null, null, null]
          )
          const thisRange = "A" + (currentRow + 1) + ":D" + (currentRow + 1)
          formatGroups.header.ranges.push(thisRange)
          formatGroups.mergeCells.ranges.push(thisRange)

          currentRow += 2

        } else if (responseObject.runs && responseObject.runs.length) {
          responseObject.runs.forEach(run => {
            runArray.push(
              [null, null, null, null],
              ["Source", endPoint.name, "Run Date", new Date(run.runDate)],
              ["Seats", run.ambulatorySpacePoints, "Wheelchair spaces", run.standardWheelchairSpacePoints],
              ["Lift", run.hasLift ? "Yes" : "No", "Ramp", run.hasLift ? "Yes" : "No"],
              ["Start location", run.startLocation, "End location", run.endLocation],
              [null, "Stop Time", "Stop City", "Rider Change"]
            )
            formatGroups.label.ranges.push("A" + (currentRow + 1) + ":A" + (currentRow + 4), "C" + (currentRow + 1) + ":C" + (currentRow + 4))
            formatGroups.date.ranges.push("D" + (currentRow + 1))
            formatGroups.seat.ranges.push("B" + (currentRow + 2), "D" + (currentRow + 2))
            formatGroups.runAttributes.ranges.push("A" + (currentRow + 1) + ":D" + (currentRow + 5))
            formatGroups.header.ranges.push("A" + (currentRow + 5) + ":D" + (currentRow + 5))
            currentRow += 6
            run.stops.forEach(stop => {
              runArray.push([null, new Date(stop.time), stop.city, stop.riderChange])
            })
            formatGroups.time.ranges.push("B" + currentRow + ":B" + (currentRow + run.stops.length -1))
            formatGroups.riderChange.ranges.push("D" + currentRow + ":D" + (currentRow + run.stops.length -1))
            currentRow += run.stops.length           
          })
        } else {
          runArray.push(
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
       let range = runSheet.getRange(1,1,runArray.length,4)
       range.setValues(runArray)
       range.clearFormat()
       applyFormats(formatGroups, runSheet)
       runSheet.autoResizeColumns(1,4)
      }
    })
  } catch(e) { logError(e) }
}

function getTrips() {
  try {
    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasTrips) {
        let params = {}
        params.resource = "trips"
        params.version = endPoint.version
        params.apiKey = endPoint.apiKey
        let hmac = {}
        hmac.nonce = Utilities.getUuid()
        hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
        hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)
        
        const options = {method: 'GET', contentType: 'application/json'}
        let response = UrlFetchApp.fetch(endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac), options)
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
        } catch(e) {
          responseObject = {status: "LOCAL_ERROR:" + e.name}
        }
          
        let tripArray = []
        if (responseObject.trips) {
          responseObject.trips.forEach(trip => {
          })
        }
        log("Received from GET:",JSON.stringify(responseObject))
      }
    })
  } catch(e) { logError(e) }
}