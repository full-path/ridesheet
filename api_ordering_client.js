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
    const tripRange = tripSheet.getDataRange()
    const allTrips = getRangeValuesAsTable(tripRange)
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
    //    Response: none
    //    Local action: log the decline if the trip is still in a shared status so that it's useful to the user and also prevents 
    //                  resending the tripRequest to the same provider again.

    const filterTripRequestResponseClaims          = row => row.tripAvailable
    const filterTripRequestResponseClaimsToAccept  = row => row.tripAvailable && allTrips.filter(filterTripsToBeClaimed).map(trip => trip["Trip ID"]).includes(row.tripID)
    const filterTripRequestResponseClaimsToRescind = row => row.tripAvailable && !allTrips.filter(filterTripsToBeClaimed).map(trip => trip["Trip ID"]).includes(row.tripID)
    const filterTripRequestResponseDeclines        = row => !row.tripAvailable
    const filterTripsToBeClaimed = row => {
      return (
        tripRequestResponses.filter(filterTripRequestResponseClaims).map(t => t.tripID).includes(row["Trip ID"]) && 
        row["Trip Date"] >= dateToday() && 
        row["Share"] === true && 
        row["Source"] === ""
      )
    }

    const tripRequestResponseClaimsToAccept = tripRequestResponses.filter(filterTripRequestResponseClaimsToAccept)
    const tripRequestResponseClaimsToRescind = tripRequestResponses.filter(filterTripRequestResponseClaimsToRescind)
    const tripRequestResponseDeclines = tripRequestResponses.filter(filterTripRequestResponseDeclines)

    let allCustomers
    if (tripRequestResponseClaimsToAccept.length) {
      allCustomers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
    }
    tripRequestResponseClaimsToAccept.forEach(claim => {
      let trip = allTrips.find(row => row["Trip ID"] === claim.tripTicketId)
      let customer = allCustomers.find(row => row["Customer ID"] === trip["Customer ID"])
      let confirmationOut = {}
      confirmationOut.tripAvailable = false
      confirmationOut.tripTicketId = claim.tripTicketId
      confirmationOut["@customerId"] = trip["Customer ID"]
      confirmationOut.customerName = customer["Customer Last Name"] + ", " + customer["Customer First Name"]
      if (customer["Mobile Phone"]) {
        confirmationOut.customerMobilePhone = customer["Phone Number"]
      }
      confirmationOut.numOtherReservedPassengers = trip["Riders"] - 1
      if (trip["Earliest PU Time"]) {
        confirmationOut.pickupWindowStartTime = buildTimeFromSpec(trip["Trip Date"], trip["Earliest PU Time"])
      }
      if (trip["Latest PU Time"]) {
        confirmationOut.pickupWindowEndTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
      }
      if (claim.scheduledPickupTime) {
        confirmationOut.negotiatedPickupTime = claim.scheduledPickupTime
      }
      confirmationOut.pickupTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
      confirmationOut.dropoffTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest DO Time"])
      if (trip["Appt Time"]) confirmationOut.appointmentTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
      confirmationOut.dropoffAddress = buildAddressToSpec(trip["PU Address"])
      confirmationOut.pickupAddress = buildAddressToSpec(trip["DO Address"])
      let openAttributes = {}
      openAttributes.estimatedTripDurationInSeconds = timeOnlyAsMilliseconds(trip["Est Hours"] || 0)/1000
      openAttributes.estimatedTripDistanceInMiles = trip["Est Miles"]
      if (trip["Mobility Factors"]) openAttributes.mobilityFactors = trip["Mobility Factors"]
      if (trip["Notes"]) openAttributes.notes = trip["Notes"]
      confirmationOut["@openAttribute"] = JSON.stringify(openAttributes)
      response.results.push({clientOrderConfirmations: confirmationOut})

      let customerInfoOut = {}
      customerInfoOut.customerId = trip["Customer ID"]
      customerInfoOut.customerName = customer["Customer Last Name"] + ", " + customer["Customer First Name"]
      if (customer["Home Address"]) customerInfoOut.customerAddress = buildAddressToSpec(customer["Home Address"])
      if (customer["Phone Number"]) customerInfoOut.customerPhone = buildPhoneNumberToSpec(customer["Phone Number"])
      if (customer["Mobile Phone"]) customerInfoOut.customerMobilePhone = buildPhoneNumberToSpec(customer["Mobile Phone"])
      if (customer["Gender"]) customerInfoOut.gender = customer["Gender"]
      if (customer["Caregiver Contact Info"]) customerInfoOut.caregiverContactInformation = customer["Caregiver Contact Info"]
      if (customer["Emergency Phone Number"]) customerInfoOut.customerEmergencyPhoneNumber = customer["Emergency Phone Number"]
      if (customer["Emergency Contact Name"]) customerInfoOut.customerEmergencyContactName = customer["Emergency Contact Name"]
      if (customer["Required Care Comments"]) customerInfoOut.requiredCareComments = customer["Required Care Comments"]
      if (customer["Notes"]) customerInfoOut.notesForDriver = customer["Notes"]
      response.results.push({customerInfo: customerInfoOut})

      let sentTripData = {"Claim Time": claimTime, "Claiming Agency": apiAccount.name, "Sched PU Time": claim.scheduledPickupTime}
      moveRow(tripSheet.getRange("A" + trip._rowPosition + ":" + trip._rowPosition), sentTripSheet, {extraFields: sentTripData})
    })

    tripRequestResponseClaimsToRescind.forEach(claim => {
      let confirmationOut = {}
      confirmationOut.tripTicketId = claim.tripTicketId
      confirmationOut.tripAvailable = false
      response.results.push({clientOrderConfirmations: confirmationOut})
    })

    if (tripRequestResponseDeclines.length) {
      let tripUpdates = allTrips.map(row => {return {}})
      tripRequestResponseDeclines.forEach(decline => {
        let trip = allTrips.find(row => row["Trip ID"] === decline.tripTicketId)
        tripUpdates[trip._rowIndex] = {}
        if (trip["Declined By"]) {
          let declinedBy = JSON.parse(trip["Declined By"])
          if (!declinedBy.includes(apiAccount.name)) {
            declinedBy.push(apiAccount.name)
            tripUpdates["Declined By"] = declinedBy
          }
        } else {
          tripUpdates["Declined By"] = JSON.stringify([apiAccount.name])
        }
      })
      setValuesByHeaderNames(tripUpdates, tripRange)
    }

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