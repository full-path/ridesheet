function createReturnTrip() {
  const ss              = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet       = ss.getActiveSheet()
  const sourceTripRange = getFullRow(tripSheet.getActiveCell())
  const sourceTripRow   = sourceTripRange.getRow()
  const sourceTripData  = getRangeValuesAsTable(sourceTripRange)[0]
  let   returnTripData  = {...sourceTripData}
  if (tripSheet.getName() === "Trips" && isCompleteTrip(sourceTripData)) {
    returnTripData["PU Address"] = sourceTripData["DO Address"]
    returnTripData["DO Address"] = sourceTripData["PU Address"]
    if (sourceTripData["Appt Time"]) {
      returnTripData["PU Time"] = timeAdd(sourceTripData["Appt Time"], 60*60*1000)
    } else if (sourceTripData["DO Time"]) {
      returnTripData["PU Time"] = timeAdd(sourceTripData["DO Time"], 60*60*1000)
    } else {
      returnTripData["PU Time"] = null
    }
    returnTripData["DO Time"]     = null
    returnTripData["Appt Time"]   = null
    returnTripData["Est Hours"]   = null
    returnTripData["Est Miles"]   = null
    returnTripData["Trip ID"]     = Utilities.getUuid()
    returnTripData["Calendar ID"] = null
    tripSheet.insertRowAfter(sourceTripRow)
    let returnTripRange = getFullRow(tripSheet.getRange(sourceTripRow + 1, 1))
    setValuesByHeaderNames([returnTripData],returnTripRange)
    fillHoursAndMilesOnEdit(returnTripRange)
    updateTripTimesOnEdit(returnTripRange)
  } else {
    ss.toast("Select a cell in a trip to create its return trip.")
  }
}

function moveTripsToReview() {
  const ss              = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet       = ss.getSheetByName("Trips")
  const tripReviewSheet = ss.getSheetByName("Trip Review")
  const tripFilter      = function(row) { return row["Trip Date"] && row["Trip Date"] < dateToday() }
  moveRows(tripSheet, tripReviewSheet, tripFilter)  
  
  const runSheet        = ss.getSheetByName("Runs")
  const runReviewSheet  = ss.getSheetByName("Run Review")
  const runFilter       = function(row) { return row["Run Date"] && row["Run Date"] < dateToday() }
  moveRows(runSheet, runReviewSheet, runFilter)  
}

function moveTripsToArchive() {
  const ss               = SpreadsheetApp.getActiveSpreadsheet()
  const tripReviewSheet  = ss.getSheetByName("Trip Review")
  const tripArchiveSheet = ss.getSheetByName("Trip Archive")
  const tripFilter       = function(row) { 
    const columns = getDocProp("tripReviewRequiredFields")
    blankColumns = columns.filter(column => !row[column])
    return blankColumns.length === 0
  }
  moveRows(tripReviewSheet, tripArchiveSheet, tripFilter)  

  const runReviewSheet  = ss.getSheetByName("Run Review")
  const runArchiveSheet = ss.getSheetByName("Run Archive")
  const runFilter       = function(row) { 
    const columns = getDocProp("runReviewRequiredFields")
    blankColumns = columns.filter(column => !row[column])
    return blankColumns.length === 0
  }
  moveRows(runReviewSheet, runArchiveSheet, runFilter)  
}
  
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
          sheet.autoResizeRows(1,2)
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

// Ordering client (this RideSheet instance) receives request for tripRequests from provider and
// returns an array JSON objects, each element of which complies with
// Telegram 1A of TCRP 210 Transactional Data Spec.
function receiveRequestForTripRequests() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange()).filter(tripRow => {
      return tripRow["Trip Date"] >= dateToday() && tripRow["Share"] === true && tripRow["Source"] === ""
    })
    let result = trips.map(tripIn => {
      let tripOut = {}
      tripOut.pickupAddress = buildAddressToSpec(tripIn["PU Address"])
      tripOut.dropoffAddress = buildAddressToSpec(tripIn["DO Address"])
      if (tripIn["Earliest PU Time"]) {
        tripOut.pickupWindowStartTime = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["Earliest PU Time"])}
      }
      if (tripIn["Latest PU Time"]) {
        tripOut.pickupWindowEndTime = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["Latest PU Time"])}
      }
      tripOut.pickupTime = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["PU Time"])}
      tripOut.dropoffTime = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["DO Time"])}
      if (tripIn["Appt Time"]) {
        tripOut.appointmentTime = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["Appt Time"])}
      }

      let openAttributes = {}
      openAttributes.tripTicketId = tripIn["Trip ID"]
      openAttributes.estimatedTripDurationInSeconds = timeOnlyAsMilliseconds(tripIn["Est Hours"] || 0)/1000
      openAttributes.estimatedTripDistanceInMiles = tripIn["Est Miles"]
      if (tripIn["Guests"]) openAttributes.guestCount = tripIn["Guests"]
      if (tripIn["Mobility Factors"]) openAttributes.mobilityFactors = tripIn["Mobility Factors"]
      if (tripIn["Notes"]) openAttributes.notes = tripIn["Notes"]
      tripOut["@openAttribute"] = JSON.stringify(openAttributes)

      return {tripRequest: tripOut}
    })
    result.sort((a, b) => a.tripRequest.pickupTime["@time"].getTime() - b.tripRequest.pickupTime["@time"].getTime())
    console.log(JSON.stringify(result[1]))
    console.log(JSON.stringify(result[0]))
    return result
  } catch(e) { logError(e) }
}

// Provider (this RideSheet instance) sends tripRequestResponses
// (whether service for each tripRequest is available or not) to ordering client.
// Message is an array JSON objects, each element of which complies with
// Telegram 1:B of TCRP 210 Transactional Data Spec.
function sendTripRequestResponses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const trips = getRangeValuesAsTable(ss.getSheetByName("Outside Trips").getDataRange()).filter(tripRow => {
      return (
          tripRow["Trip Date"] >=  dateToday() &&
        ( tripRow["Decline"]   === true || tripRow["Claim"] === true) &&
        !(tripRow["Decline"]   === true && tripRow["Claim"] === true)
      )
    })
    let result = trips.map(tripIn => {
      let tripOut = {}
      tripOut.tripAvailable = (tripRow["Claim"] === true)
      if (tripOut.tripAvailable && trips["Scheduled PU Time"]) {
        tripOut["scheduledPickupTime"] = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["Scheduled PU Time"])}
      }

      let openAttributes = {}
      openAttributes.tripTicketId = tripIn["Trip ID"]
      tripOut["@openAttribute"] = JSON.stringify(openAttributes)

      return {tripRequestResponse: tripOut}
    })
    return result
  } catch(e) { logError(e) }
}

// Ordering client (this RideSheet instance) receives tripRequestResponses
// (whether service for each tripRequest is available or not) from provider.
// Message is an array JSON objects, each element of which complies with
// Telegram 1B of TCRP 210 Transactional Data Spec.
function receiveTripRequestResponses() {

}

function handleReceivedTripClaims() {

}

function handleReceivedTripDenials() {

}

// Ordering client (this RideSheet instance) sends clientOrderConfirmations
// to provider.
// Message is an array JSON objects, each element of which complies with
// Telegram 2A of TCRP 210 Transactional Data Spec.
// This will include "confirmations" that returns rescinded orders.
function sendClientOrderConfirmations() {

}

// Provider (this RideSheet instance) receives clientOrderConfirmations
// from ordering client.
// Message is an array JSON objects, each element of which complies with
// Telegram 2A of TCRP 210 Transactional Data Spec.
// This will include "confirmations" that returns rescinded orders.
function receiveClientOrderConfirmations() {

}

// Provider (this RideSheet instance) sends providerOrderConfirmations
// to ordering client.
// Message is an array JSON objects, each element of which complies with
// Telegram 2B of TCRP 210 Transactional Data Spec.
function sendProviderOrderConfirmations() {

}

// Ordering client (this RideSheet instance) receives providerOrderConfirmations
// from provider.
// Message is an array JSON objects, each element of which complies with
// Telegram 2B of TCRP 210 Transactional Data Spec.
function receiveProviderOrderConfirmations() {

}

// Ordering client (this RideSheet instance) sends customerInfo records
// to provider.
// Message is an array JSON objects, each element of which complies with
// Telegram 2A1 of TCRP 210 Transactional Data Spec.
function sendCustomerInfo() {

}

// Provider (this RideSheet instance) receives customerInfo records
// from ordering client.
// Message is an array JSON objects, each element of which complies with
// Telegram 2A1 of TCRP 210 Transactional Data Spec.
function receiveCustomerInfo() {

}

function isCompleteTrip(trip) {
  return (trip["Trip Date"] && trip["Customer Name and ID"])
}

function buildAddressToSpec(address) {
  try {
    let result = {}
    const parsedAddress = parseAddress(address)
    result["@addressName"] = parsedAddress.geocodeAddress
    if (parsedAddress.parenText) {
      let manualAddr = {}
      manualAddr["@manualText"] = parsedAddress.parenText
      manualAddr["@sendtoInvoice"] = true
      manualAddr["@sendtoVehicle"] = true
      manualAddr["@sendtoOperator"] = true
      manualAddr["@vehicleConfirmation"] = false
      result.manualDescriptionAddress = manualAddr
    }
    return result
  } catch(e) { logError(e) }
}

function buildAddressFromSpec(address) {
  try {
    let result = address["@addressName"]
    if (address.manualDescriptionAddress) {
      result = result + " (" + address.manualDescriptionAddress["@manualText"] + ")"
    }
    return result
  } catch(e) { logError(e) }
}