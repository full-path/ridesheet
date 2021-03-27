// Ordering client (this RideSheet instance) receives request for tripRequests from provider and
// returns an array JSON objects, each element of which complies with
// Telegram 1A of TCRP 210 Transactional Data Spec.
function receiveRequestForTripRequestsReturnTripRequests() {
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

// Ordering client (this RideSheet instance) receives tripRequestResponses
// (whether service for each tripRequest is available or not) from provider.
// Message in is an array of JSON tripRequestResponses objects, each element of which complies with
// Telegram 1B of TCRP 210 Transactional Data Spec.
// Response out is an array of JSON clientOrderConfirmations objects, each element of which complies with
// Telegram 2A of TCRP 210 Transactional Data Spec.
// This will include "confirmations" that returns rescinded orders.
function receiveTripRequestResponsesReturnClientOrderConfirmations(payload, apiAccount) {
  try {
    // Get the payload into a format that's easier to work with internally
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet = ss.getSheetByName("Trips")
    const sentTripSheet = ss.getSheetByName("Sent Trips")
    const allTrips = getRangeValuesAsTable(tripSheet.getDataRange())
    const claimTime = new Date()
    let response = {}
    response.status = "OK"
    response.results = []

    const tripRequestResponses = payload.map(trip => {
      let result = {}
      const openAttributes = JSON.parse(trip["@openAttribute"])
      result.tripID = openAttributes.tripTicketId
      result.tripAvailable = trip.tripRequestResponse.tripAvailable
      if (trip.tripRequestResponse.scheduledPickupTime) result.scheduledPickupTime = trip.tripRequestResponse.scheduledPickupTime
      return result
    })

    // With each tripRequestResponse, the possible actions are:
    // 1. Accept the claim
    //    Response: Additional trip information
    //    Local action: move trip to "Sent Trips"
    // 2. Rescind the initial offer
    //    Response: Send the rescission
    //    Local action: none
    // 3. Accept the decline
    //    Response: none (?)
    //    Local action: log the decline if the trip is still in a shared status so that it's useful to the user and also prevents 
    //                  resending the tripRequest to the same provider again.

    const filterTripRequestResponseClaims          = function(row) { return row.tripAvailable }
    const filterTripRequestResponseClaimsToAccept  = function(row) { return row.tripAvailable && allTrips.filter(tripsToBeClaimed).includes(row.tripID) }
    const tripRequestResponseClaimsToRescind = function(row) { return row.tripAvailable && !allTrips.filter(tripsToBeClaimed).includes(row.tripID) }
    const tripRequestResponseDeclines        = function(row) { return !row.tripAvailable }
    const tripsToBeClaimed = function(row) {
      return (
        tripRequestResponses.filter(filterTripRequestResponseClaims).map(t => t.tripID).includes(row["Trip ID"]) && 
        row["Trip Date"] >= dateToday() && 
        row["Share"] === true && 
        row["Source"] === ""
      )
    }

    const tripRequestResponseClaimsToAccept = tripRequestResponses.filter(filterTripRequestResponseClaimsToAccept)
    let allCustomers
    if (tripRequestResponseClaimsToAccept > 1) allCustomers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
    tripRequestResponseClaimsToAccept.forEach(claim => {
      trip = allTrips.find(row => row["Trip ID"] === claim.tripTicketId)
      customer = allCustomers.find(row => row["Customer ID"] === trip["Customer ID"])
      let tripInfo = {}
      tripInfo.tripTicketId = claim.tripTicketId
      tripInfo["@customerId"] = 1
      tripInfo.customerName = customer["Customer Last Name"] + ", " + customer["Customer First Name"]
      if (customer["Phone Number"]) {
        tripInfo.customerMobilePhone = customer["Phone Number"]
      }
      tripInfo.numOtherReservedPassengers = trip["Riders"] - 1
      if (tripIn["Earliest PU Time"]) {
        tripInfo.pickupWindowStartTime = buildTimeFromSpec(trip["Trip Date"], trip["Earliest PU Time"])
      }
      if (tripIn["Latest PU Time"]) {
        tripInfo.pickupWindowEndTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
      }
      if (claim.scheduledPickupTime) {
        tripInfo.negotiatedPickupTime = claim.scheduledPickupTime
      }
      tripInfo.pickupTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
      tripInfo.dropoffTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest DO Time"])
      if (trip["Appt Time"]) tripInfo.appointmentTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
      tripInfo.dropoffAddress = buildAddressToSpec(trip["PU Address"])
      tripInfo.pickupAddress = buildAddressToSpec(trip["DO Address"])
      response.results.push({clientOrderConfirmations: tripInfo})
      let sentTripData = {"Claim Time": claimTime, "Claiming Agency": apiAccount.name, "Sched PU Time": claim.scheduledPickupTime}
      moveRow(tripSheet.getRange("A" + trip._rowPosition + ":" + trip._rowPosition), sentTripSheet, {extraFields: sentTripData})
    })

    return response
  } catch(e) { logError(e) }
}

// Ordering client (this RideSheet instance) receives providerOrderConfirmations
// from provider.
// Message is an array JSON objects, each element of which complies with
// Telegram 2B of TCRP 210 Transactional Data Spec.
// Response out is an array of JSON customerInfo objects, each element of which complies with
// Telegram 2A1 of TCRP 210 Transactional Data Spec.
function receiveProviderOrderConfirmationsReturnCustomerInformation(payload, apiAccount) {
  try {
    let response = {}
    response.status = "OK"
    return response
  } catch(e) { logError(e) }
}