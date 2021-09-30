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

function receiveTripRequestResponses(payload) { 
  const tripRequestResponses = formatTripRequestResponses(payload)
  const allTrips = getAllTrips()
  const filteredResponses = filterTripRequestResponses(tripRequestResponses, allTrips)
  return filteredResponses
}

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
function returnClientOrderConfirmations(filteredResponses, apiAccount) {
  const orderConfirmations = processTripRequestResponses(filteredResponses)
  const response = {}
  response.status = "OK"
  response.results = orderConfirmations

  // Performance improvement: do local upkeep after sending the 
  // response back to the provider
  const {accept, rescind, decline} = filteredResponses
  // TODO: These still don't work
  //moveAcceptedClaimsToSentTrips(accept, apiAccount)
  //logDeclinedTripRequests(decline, apiAccount)

  log('Response', response)

  return response
}

function testReceiveTripRequestResponses() {
  const payload = [{"tripRequestResponse":{"tripAvailable":true,"@openAttribute":"{\"tripTicketId\":\"0c4d6861-43e8-4146-9613-97d62ea84bff\"}"}},{"tripRequestResponse":{"tripAvailable":false,"@openAttribute":"{\"tripTicketId\":\"f082d0b6-8409-4025-b8e7-f28f6762bfce\"}"}}]
  const apiAccount = { name:"Agency B"}
  const processedResponses = receiveTripRequestResponses(payload)
  const response = returnClientOrderConfirmations(processedResponses, apiAccount)
  console.log(response)
}

function formatTripRequestResponses(payload) {
  const tripRequestResponses = payload.map(trip => {
      const response = trip["tripRequestResponse"]
      const openAttributes = JSON.parse(response["@openAttribute"])
      let result = {}
      result.tripID = openAttributes["tripTicketId"]
      result.tripAvailable = response["tripAvailable"]
      if (response.scheduledPickupTime) {
        result.scheduledPickupTime = response.scheduledPickupTime
      }
      return result
    })
    return tripRequestResponses
}

function getAllTrips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const tripRange = tripSheet.getDataRange()
  const allTrips = getRangeValuesAsTable(tripRange) 
  return allTrips
}

function filterTripRequestResponses(responses, allTrips) {
  const filterTripRequestResponseClaims = row => row.tripAvailable
  const filterTripsToBeClaimed = row => {
      return (
        responses.filter(filterTripRequestResponseClaims).map(t => t.tripID).includes(row["Trip ID"]) && 
        row["Trip Date"] >= dateToday() && 
        row["Share"] === true && 
        row["Source"] === ""
      )
  }
  const filterTripRequestResponseClaimsToAccept  = row => row.tripAvailable && allTrips.filter(filterTripsToBeClaimed).map(trip => trip["Trip ID"]).includes(row.tripID)
  const filterTripRequestResponseClaimsToRescind = row => row.tripAvailable && !allTrips.filter(filterTripsToBeClaimed).map(trip => trip["Trip ID"]).includes(row.tripID)
  const filterTripRequestResponseDeclines = row => !row.tripAvailable

  const accept = responses.filter(filterTripRequestResponseClaimsToAccept)
  const rescind = responses.filter(filterTripRequestResponseClaimsToRescind)
  const decline = responses.filter(filterTripRequestResponseDeclines)
  return {accept, rescind, decline}
}

function processTripRequestResponses(filteredTripRequestResponses) {
  const {accept, rescind} = filteredTripRequestResponses
  const acceptedClaimResponses = processAcceptedClaims(accept)
  const rescindedClaimResponses = processRescindedClaims(rescind)
  let orderConfirmations = acceptedClaimResponses.concat(rescindedClaimResponses)
  return orderConfirmations
}

function processAcceptedClaims(acceptedClaims) {
  let results = []
  if (!acceptedClaims || acceptedClaims.length < 1) {
    return results
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const allCustomers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
  const allTrips = getAllTrips()
  acceptedClaims.forEach(claim => {
    let trip = allTrips.find(row => row["Trip ID"] === claim.tripID)
    let customer = allCustomers.find(row => row["Customer ID"] === trip["Customer ID"])
    let customerName = customer["Customer Last Name"] + ", " + customer["Customer First Name"]
    let pickupWindowStartTime = null, pickupWindowEndTime = null, negotiatedPickupTime = null, appointmentTime = null
    if (trip["Earliest PU Time"]) {
      pickupWindowStartTime = buildTimeFromSpec(trip["Trip Date"], trip["Earliest PU Time"])
    }
    if (trip["Latest PU Time"]) {
      pickupWindowEndTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
    }
    if (claim.scheduledPickupTime) {
      negotiatedPickupTime = claim.scheduledPickupTime
    }
    let pickupTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"])
    let dropoffTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest DO Time"])
    if (trip["Appt Time"]) { 
      appointmentTime = buildTimeFromSpec(trip["Trip Date"], trip["Latest PU Time"]) 
    }
    let dropoffAddress = buildAddressToSpec(trip["PU Address"])
    let pickupAddress = buildAddressToSpec(trip["DO Address"])
    let openAttributes = {}
    openAttributes.estimatedTripDurationInSeconds = timeOnlyAsMilliseconds(trip["Est Hours"] || 0)/1000
    openAttributes.estimatedTripDistanceInMiles = trip["Est Miles"]
    if (trip["Mobility Factors"]) openAttributes.mobilityFactors = trip["Mobility Factors"]
    if (trip["Notes"]) openAttributes.notes = trip["Notes"]
    openAttributes = JSON.stringify(openAttributes)
    let confirmationOut = {
      tripAvailable: true,
      tripTicketId: claim.tripID,
      "@customerId": trip["Customer ID"],
      "@openAttributes": openAttributes,
      customerMobilePhone: customer["Phone Number"],
      customerName,
      pickupWindowStartTime,
      pickupWindowEndTime,
      negotiatedPickupTime,
      pickupTime,
      dropoffTime,
      appointmentTime,
      dropoffAddress,
      pickupAddress,
    }
    results.push({clientOrderConfirmations: confirmationOut})

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
    results.push({customerInfo: customerInfoOut})
  })
  return results
}

function processRescindedClaims(rescindedClaims) {
  let results = []
  rescindedClaims.forEach(claim => {
      let confirmationOut = {}
      confirmationOut.tripTicketId = claim.tripTicketId
      confirmationOut.tripAvailable = false
      results.push({clientOrderConfirmations: confirmationOut})
  })
  return results
}

function moveAcceptedClaimsToSentTrips(acceptedClaims, apiAccount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sentTripSheet = ss.getSheetByName("Sent Trips")
  const tripSheet = ss.getSheetByName("Trips")
  const allTrips = getAllTrips()
  const claimTime = new Date()
  acceptedClaims.forEach(claim => {
    let trip = allTrips.find(row => row["Trip ID"] === claim.tripID)
    let sentTripData = {"Claim Time": claimTime, "Claiming Agency": apiAccount.name, "Sched PU Time": claim.scheduledPickupTime}
    moveRow(tripSheet.getRange("A" + trip._rowPosition + ":" + trip._rowPosition), sentTripSheet, {extraFields: sentTripData})
  })
}

function logDeclinedTripRequests(declinedTripRequests, apiAccount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const allTrips = getAllTrips()
  const tripSheet = ss.getSheetByName("Trips")
  const tripRange = tripSheet.getDataRange()
  if (declinedTripRequests.length) {
      let tripUpdates = allTrips.map(row => {return {}})
      declinedTripRequests.forEach(decline => {
        let trip = allTrips.find(row => row["Trip ID"] === decline.tripID)
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