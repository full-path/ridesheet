// Provider (this RideSheet instance) sends request for tripRequests to ordering client.
// Received format is an array JSON objects, each element of which complies with
// Telegram 1A of TCRP 210 Transactional Data Spec.
function sendRequestForTripRequests() {
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
      log("Name",endPoint.name)
      if (endPoint.hasTrips) {
        let params = {resource: "tripRequests"}
        let response = getResource(endPoint, params)
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
          log(responseObject)
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
          sheet.autoResizeRows(1,2)
        } else {
          log("No trip requests")
          grid.push(
            [null, null, null, endPoint.name + " responded with no trip requests"].concat(Array(14).fill(null))
          )
          const thisRange = "D" + currentRow + ":" + lastColumnLetter + currentRow
          formatGroups.mergeCells.ranges.push(thisRange)

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
          sheet.autoResizeRows(1,2)
          currentRow += 1
        }
      }
    })
  } catch(e) { logError(e) }
}

// Provider (this RideSheet instance) sends tripRequestResponses
// (whether service for each tripRequest is available or not) to ordering client.
// Send an array of tripRequestResponse JSON objects which comply with
// Telegram 1B of TCRP 210 Transactional Data Spec.
// Receive in response an array of clientOrderConfirmation JSON objects which comply with
// Telegram 2A of TCRP 210 Transactional Data Spec.
function sendTripRequestResponses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()

    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasTrips) {
        let params = {resource: "tripRequestResponses"}

        const trips = getRangeValuesAsTable(ss.getSheetByName("Outside Trips").getDataRange(), {headerRowPosition: 2}).filter(tripRow => {
          return (
            tripRow["Trip Date"] >=  dateToday() &&
            ( tripRow["Decline"]   === true || tripRow["Claim"] === true) &&
            !(tripRow["Decline"]   === true && tripRow["Claim"] === true)
          )
        })
        let payload = trips.map(tripIn => {
          let tripOut = {}
          tripOut.tripAvailable = (tripIn["Claim"] === true)
          if (tripOut.tripAvailable && trips["Scheduled PU Time"]) {
            tripOut["scheduledPickupTime"] = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["Scheduled PU Time"])}
          }

          let openAttributes = {}
          openAttributes.tripTicketId = tripIn["Trip ID"]
          tripOut["@openAttribute"] = JSON.stringify(openAttributes)

          return {tripRequestResponse: tripOut}
        })

        // Here we're sending the tripRequestResponses in the payload.
        // responseObject, if validated, contains clientOrderConfirmations
        let response = postResource(endPoint, params, JSON.stringify(payload))
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
        } catch(e) {
          logError(e)
          responseObject = {status: "LOCAL_ERROR:" + e.name}
        }
        if (responseObject.status === "OK") {
          // Process clientOrderConfirmations

          // confirm the confirmations
          sendProviderOrderConfirmations(responseObject, endPoint)
        }
      }
    })  

    return result
  } catch(e) { logError(e) }
}

// Send an array of providerOrderConfirmation JSON objects which comply with
// Telegram 2B of TCRP 210 Transactional Data Spec.
// Receive in response an array of customerInfo JSON objects which comply with
// Telegram 2A1 of TCRP 210 Transactional Data Spec.
function sendProviderOrderConfirmations(orderingClientResponses, endPoint) {
  if (endPoint.hasTrips) {
    let params = {resource: "providerOrderConfirmations"}
    let payload = {status: "OK"}

    let response = postResource(endPoint, params, payload)
    let responseObject
    try {
      responseObject = JSON.parse(response.getContentText())
    } catch(e) {
      logError(e)
      responseObject = {status: "LOCAL_ERROR:" + e.name}
    }
    if (responseObject.status === "OK") {
      // Process the customerInfo objects
      
    }
  }
}
