// Ordering client (this RideSheet instance) receives request for tripRequests from provider and
// returns an array JSON objects, each element of which complies with
// Telegram 1A of TCRP 210 Transactional Data Spec.


// Notes: must send each trip request individually - is there a way to do this more efficiently in 
// sheets, rather than waiting synchronously on each one? 
// Also - could be updated to *only* send new shared trips
function sendTripRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  endPoints = getDocProp("apiGetAccess")
  endPoints.forEach(endPoint => {
    if (endPoint.hasTrips) {
      const params = {endpointPath: "/v1/TripRequest"}
      const trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange()).filter(tripRow => {
        if (tripRow["Declined By"]) {
            const declinedBy = JSON.parse(tripRow["Declined By"])
            if (declinedBy.includes(apiAccount.name)) {
              return false
            }
          }
        return tripRow["Trip Date"] >= dateToday() && tripRow["Share"] === true && tripRow["Source"] === ""
      })
      // set necessary params: HMAC headers, resource (endpoint), ??
      trips.forEach(trip => {
        let payload = formatTripRequest(trip)
        let response = postResource(endpoint, params, JSON.stringify(payload))
        try {
          let responseObject = JSON.parse(response.getContentText())
          log('#1A response', responseObject)
          if (responseObject.status === "OK") {
            // mark "Shared" as true
          }
          // TODO: Check for TDS status codes
          // if share is successful, mark 'Shared' as true
          // What to do if we don't receive a 200?
        } catch(e) {
          logError(e)
        }
      })
    }
  })
}

// Following telegram #1A https://app.swaggerhub.com/apis/full-path/RideNoCo-TDS/0.5.a3#/Multiple%20Endpoints/post_v1_TripRequest
// Currently Unsupported fields:
// *NOTE* should pretty print in notes field
// - detailedPickupLocationDescription
// - detailedDropoffLocationDescription
// - tripPurpose
// - specialAttributes
// - detoursPermissible
// - negotiatedPickupTime
// - hardConstraintOnPickupTime
// - hardConstraintOnDropoff Time
// - transportServices
// - tripTransfer (Must support!)
function formatTripRequest(trip) {
  const formattedTrip = {
    tripTicketId: trip["Trip ID"],
    pickupAddress: buildAddressToSpec(trip["PU Address"]),
    dropoffAddress: buildAddressToSpec(trip["DO Address"]),
    pickupTime: combineDateAndTime(trip["Trip Date"], trip["PU Time"]),
    dropoffTime:  combineDateAndTime(trip["Trip Date"], trip["DO Time"]),
    customerInfo: getCustomerInfo(trip),
    openAttributes: {
      estimatedTripDurationInSeconds: timeOnlyAsMilliseconds(trip["Est Hours"] || 0)/1000,
      estimatedTripDistanceInMiles: trip["Est Miles"]
    }
  }
  // To add conditionally: pickupWindowStartTime, pickupWindowEndTime,
  // appointmentTime, numOtherReservedPassengers, notesForDriver
  return formattedTrip
}

function getCustomerInfo(trip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const allCustomers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
  const customer = allCustomers.find(row => row["Customer ID"] === trip["Customer ID"])
  const formattedCustomer = {
    firstLegalName: customer["Customer First Name"],
    lastName: customer["Customer Last Name"],
    address: buildAddressToSpec(customer["Home Address"]),
    phone: buildPhoneNumberToSpec(customer["Phone Number"]),
    customerId: customer["Customer ID"]
  }
  return formattedCustomer
}

function receiveRequestForTripRequestsReturnTripRequests(apiAccount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const trips = getRangeValuesAsTable(ss.getSheetByName("Trips").getDataRange()).filter(tripRow => {
      if (tripRow["Declined By"]) {
          let declinedBy = JSON.parse(tripRow["Declined By"])
          if (declinedBy.includes(apiAccount.name)) {
            return false
          }
        }
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
//    resending the tripRequest to the same provider again.
function returnClientOrderConfirmations(filteredResponses, apiAccount) {
  const orderConfirmations = processTripRequestResponses(filteredResponses)
  const response = {}
  response.status = "OK"
  response.results = orderConfirmations

  // Performance improvement: do local upkeep after sending the 
  // response back to the provider
  const {accept, rescind, decline} = filteredResponses

  moveAcceptedClaimsToSentTrips(accept, apiAccount)
  logDeclinedTripRequests(decline, apiAccount)
  return response
}

// To use, must update payload to valid tripTicketIds
function testReceiveTripRequestResponses() {
  const payload = [{"tripRequestResponse":{"tripAvailable":false,"@openAttribute":"{\"tripTicketId\":\"e0c7a017-ee04-48f2-8dc6-6924938cad6e\"}"}},{"tripRequestResponse":{"tripAvailable":true,"@openAttribute":"{\"tripTicketId\":\"08eb4058-eae2-4dff-bb2b-f481b8cdf876\"}"}},{"tripRequestResponse":{"tripAvailable":true,"@openAttribute":"{\"tripTicketId\":\"f72122f0-fd42-45fc-9866-519b4bca8356\"}"}}]
  const apiAccount = { name:"Agency B"}
  const processedResponses = receiveTripRequestResponses(payload)
  const response = returnClientOrderConfirmations(processedResponses, apiAccount)
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
    let pickupTime = buildTimeFromSpec(trip["Trip Date"], trip["PU Time"])
    let dropoffTime = buildTimeFromSpec(trip["Trip Date"], trip["DO Time"])
    if (trip["Appt Time"]) { 
      appointmentTime = buildTimeFromSpec(trip["Trip Date"], trip["Appt Time"]) 
    }
    let dropoffAddress = buildAddressToSpec(trip["PU Address"])
    let pickupAddress = buildAddressToSpec(trip["DO Address"])
    let openAttributes = {}
    openAttributes.estimatedTripDurationInSeconds = timeOnlyAsMilliseconds(trip["Est Hours"] || 0)/1000
    openAttributes.estimatedTripDistanceInMiles = trip["Est Miles"]
    if (trip["Mobility Factors"]) openAttributes.mobilityFactors = trip["Mobility Factors"]
    if (trip["Notes"]) openAttributes.notes = trip["Notes"]
    if (trip["Guests"]) openAttributes.guests = trip["Guests"]
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
    if (customer["Customer Manifest Notes"]) customerInfoOut.notesForDriver = customer["Customer Manifest Notes"]
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

  // Remove certain fields from the trip, while leaving in any custom
  // columns that may have been created
  let tripColumnNames = getSheetHeaderNames(tripSheet)
  let ignoredFields = ["Action", "Go", "Share", "Trip Result", "Driver ID", "Vehicle ID", "Driver Calendar ID", "Trip Event ID", "Declined By", "Shared"]
  let sentTripFields = tripColumnNames.filter(col => !(ignoredFields.includes(col)))
  acceptedClaims.reverse()
  acceptedClaims.forEach(claim => {
    let trip = allTrips.find(row => row["Trip ID"] === claim.tripID)
    let sentTripData = {
      "Claimed By" : apiAccount.name,
      "Claim Time" : claimTime,
      "Sched PU Time" : claim.scheduledPickupTime
    }
    sentTripFields.forEach(key => {
      sentTripData[key] = trip[key]
    });
    createRow(sentTripSheet, sentTripData)
    tripSheet.deleteRow(trip._rowPosition)
  })
}

// TODO: change "Declined By" into comma separated values, in order to be more human-readable
function logDeclinedTripRequests(declinedTripRequests, apiAccount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const allTrips = getAllTrips()
  const tripSheet = ss.getSheetByName("Trips")
  const tripRange = tripSheet.getDataRange()
  if (declinedTripRequests.length) {
      let tripUpdates = allTrips.map(row => {return {}})
      declinedTripRequests.forEach(decline => {
        let trip = allTrips.find(row => row["Trip ID"] === decline.tripID)
        if (!trip) {
          log('Warning: Provider declining invalid trip', JSON.stringify(decline))
          return
        }
        tripUpdates[trip._rowIndex] = {}
        let declinedBy = trip["Declined By"] ? JSON.parse(trip["Declined By"]) : []
        if (trip["Declined By"]) {
          if (!declinedBy.includes(apiAccount.name)) {
            declinedBy.push(apiAccount.name)
            tripUpdates[trip._rowIndex]["Declined By"] = declinedBy
          }
        } else {
          tripUpdates[trip._rowIndex]["Declined By"] = JSON.stringify([apiAccount.name])
          declinedBy.push(apiAccount.name)
        }
        // if all attached api accounts have decline trip, highlight the row
        // todo: what if get and give access are not symmetrical? POST request will always be from someone
        // we GIVE access to, but Get access is in a more convenient array
        const apiAccounts = getDocProp("apiGetAccess")
        if (apiAccounts.length === declinedBy.length) {
          let rowPosition = trip._rowPosition
          let currentRow = tripSheet.getRange("A" + rowPosition + ":" + rowPosition)  
          currentRow.setBackgroundRGB(255,255,204)
          currentRow.getCell(1,1).setNote('Notice: Trip has been declined by all agencies')
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