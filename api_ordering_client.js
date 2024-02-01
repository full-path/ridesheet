// Ordering client (this RideSheet instance) receives request for tripRequests from provider and
// returns an array JSON objects, each element of which complies with
// Telegram 1A of TCRP 210 Transactional Data Spec.
function sendTripRequests() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const endPoints = getDocProp("apiGetAccess")
  const tripSheet = ss.getSheetByName("Trips")
  endPoints.forEach(endPoint => {
    if (endPoint.hasTrips) {
      const params = {endpointPath: "/v1/TripRequest"}
      const trips = getRangeValuesAsTable(tripSheet.getDataRange()).filter(tripRow => {
        if (tripRow["Declined By"]) {
            const declinedBy = JSON.parse(tripRow["Declined By"])
            if (declinedBy.includes(apiAccount.name)) {
              return false
            }
          }
        return tripRow["Trip Date"] >= dateToday() && tripRow["Share"] === true && tripRow["Source"] === "" && tripRow["Shared"] === ""
      })
      // set necessary params: HMAC headers, resource (endpoint), ??
      trips.forEach(trip => {
        let payload = formatTripRequest(trip)
        let response = postResource(endPoint, params, JSON.stringify(payload))
        try {
          let responseObject = JSON.parse(response.getContentText())
          log('#1A response', responseObject)
          const rowPosition = trip._rowPosition
          const currentRow = tripSheet.getRange("A" + rowPosition + ":" + rowPosition)
          const headers = getSheetHeaderNames(tripSheet)
          if (responseObject.status === "OK") {
            const colPosition = headers.indexOf("Shared") + 1
            currentRow.getCell(1, colPosition).setValue("True")
          } else {
            currentRow.setBackgroundRGB(255,221,153)
            currentRow.getCell(1,1).setNote('Failed to share trip with 1 or more providers. Check logs for more details.')
            logError(`Failure to share trip with ${endPoint.name}`, responseObject)
          }
        } catch(e) {
          logError(e)
        }
      })
    }
  })
}

// Following telegram #1A https://app.swaggerhub.com/apis/full-path/RideNoCo-TDS/0.5.a3#/Multiple%20Endpoints/post_v1_TripRequest
// TODO: figure out how to support:
// - detailedPickupLocationDescription
// - detailedDropoffLocationDescription
// - tripPurpose
// - specialAttributes
// - detoursPermissible
// - negotiatedPickupTime
// - hardConstraintOnPickupTime
// - hardConstraintOnDropoffTime
// - transportServices
// - tripTransfer
function formatTripRequest(trip) {
  const formattedTrip = {
    tripTicketId: trip["Trip ID"],
    pickupAddress: buildAddressToSpec(trip["PU Address"]),
    dropoffAddress: buildAddressToSpec(trip["DO Address"]),
    pickupTime: {time: combineDateAndTime(trip["Trip Date"], trip["PU Time"])},
    dropoffTime:  {time: combineDateAndTime(trip["Trip Date"], trip["DO Time"])},
    customerInfo: getCustomerInfo(trip), 
    appointmentTime: trip["Appt Time"] ? {time: combineDateAndTime(trip["Trip Date"], trip["Appt Time"])} : "",
    notesForDriver: trip["Notes"],
    numOtherReservedPassengers: trip["Guests"],
    openAttributes: {
      estimatedTripDurationInSeconds: timeOnlyAsMilliseconds(trip["Est Hours"] || 0)/1000,
      estimatedTripDistanceInMiles: trip["Est Miles"]
    }
  }
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
    phone: customer["Phone Number"],
    customerId: customer["Customer ID"]
  }
  return formattedCustomer
}

function receiveTripRequestResponse(response, senderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const allTrips = getAllTrips()
  const tripSheet = ss.getSheetByName("Trips")
  const apiAccounts = getDocProp("apiGiveAccess")
  const senderAccount = apiAccounts[senderId]
  const trip = allTrips.find(row => row["Trip ID"] === response.tripTicketId)
  const headers = getSheetHeaderNames(tripSheet)

  if (!response.tripAvailable) {
    if (!trip) {
      log(`${senderAccount.name} attempted to decline invalid trip`, JSON.stringify(response))
    } else {
      const rowPosition = trip._rowPosition
      const currentRow = tripSheet.getRange("A" + rowPosition + ":" + rowPosition)
      const declinedBy = trip["Declined By"] ? JSON.parse(trip["Declined By"]) : []
      declinedBy.push(senderAccount.name)
      if (apiAccounts.length === declinedBy.length) {
        currentRow.setBackgroundRGB(255,255,204)
        currentRow.getCell(1,1).setNote('Notice: Trip has been declined by all agencies')
      }
      const declinedIndex = headers.indexOf("Declined By") + 1
      currentRow.getCell(1, declinedIndex).setValue(JSON.stringify(declinedBy))
    }
    return {status: "OK"}
  } else {
    if (!trip) {
      log(`${senderAccount.name} attempted to claim invalid trip`, JSON.stringify(response))
      return {status: "400", message: "Trip no longer available"}
    }
    const rowPosition = trip._rowPosition
    const currentRow = tripSheet.getRange("A" + rowPosition + ":" + rowPosition)
    const shared = trip["Share"]
    const claimed = trip["Claim Pending"]
    if (claimed || (!shared)) {
      log(`${senderAccount.name} attempted to claim unavailable trip`, JSON.stringify(response))
      return {status: "400", message: "Trip no longer available"}
    }
    // Process successfully pending claim
    const pendingIndex = headers.indexOf("Claim Pending") + 1
    currentRow.getCell(1, pendingIndex).setValue(senderAccount.name)
    currentRow.setBackgroundRGB(0,230,153)
    currentRow.getCell(1,1).setNote('Trip claim pending. Please approve/deny')
    return {status: "OK"}
  }
}

function sendClientOrderConfirmation(sourceTripRange = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const currentTrip = sourceTripRange ? sourceTripRange : getFullRow(tripSheet.getActiveCell())
  const trip = getRangeValuesAsTable(currentTrip,{includeFormulaValues: false})
  const endPoints = getDocProp("apiGetAccess")
  const endPoint = endPoints.find(endpoint => endpoint.name === trip["Claim Pending"])
  const params = {endpointPath: "/v1/ClientOrderConfirmation"}
  const telegram = {
    tripTicketId: trip["Trip ID"],
    tripConfirmed: true
  }
  const response = postResource(endPoint, params, JSON.stringify(telegram))
  try {
    const responseObject = JSON.parse(response.getContentText())
    if (responseObject.status === "OK") {
      // move to sent trips
      log('Trip successfully confirmed', telegram)
      const sentTripSheet = ss.getSheetByName("Sent Trips")
      const claimTime = new Date()
      const tripColumnNames = getSheetHeaderNames(tripSheet)
      const ignoredFields = ["Action", "Go", "Share", "Trip Result", "Driver ID", "Vehicle ID", "Driver Calendar ID", "Trip Event ID", "Declined By", "Shared"]
      const sentTripFields = tripColumnNames.filter(col => !(ignoredFields.includes(col)))
      const sentTripData = {
        "Claimed By" : apiAccount.name,
        "Claim Time" : claimTime,
        "Sched PU Time" : claim.scheduledPickupTime
      }
      sentTripFields.forEach(key => {
        sentTripData[key] = trip[key]
      });
      createRow(sentTripSheet, sentTripData)
      tripSheet.deleteRow(trip._rowPosition)
    } else {
      logError(`Failure to confirm trip with ${endPoint.name}`, responseObject)
    }
  } catch(e) {
    logError(e)
  }
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