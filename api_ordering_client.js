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
        try {
          let response = postResource(endPoint, params, JSON.stringify(payload))
          let responseObject = JSON.parse(response.getContentText())
          log('#1A response', responseObject)
          const rowPosition = trip._rowPosition
          const currentRow = tripSheet.getRange("A" + rowPosition + ":" + rowPosition)
          const headers = getSheetHeaderNames(tripSheet)
          if (responseObject.status && responseObject.status !== "OK") {
            currentRow.setBackgroundRGB(255,221,153)
            currentRow.getCell(1,1).setNote('Failed to share trip with 1 or more providers. Check logs for more details.')
            logError(`Failure to share trip with ${endPoint.name}`, responseObject)
          } else {
            const colPosition = headers.indexOf("Shared") + 1
            currentRow.getCell(1, colPosition).setValue("True")
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
    appointmentTime: trip["Appt Time"] ? {time: combineDateAndTime(trip["Trip Date"], trip["Appt Time"])} : null,
    notesForDriver: trip["Notes"],
    numOtherReservedPassengers: trip["Guests"] ? trip["Guests"] : 0,
    openAttributes: {
      estimatedTripDurationInSeconds: timeOnlyAsMilliseconds(trip["Est Hours"] || 0)/1000,
      estimatedTripDistanceInMiles: trip["Est Miles"]
    }
  }
  return formattedTrip
}

// TODO: Add more fields, support optional callsheet fields
function getCustomerInfo(trip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const allCustomers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
  const customer = allCustomers.find(row => row["Customer ID"] === trip["Customer ID"])
  const formattedCustomer = {
    firstLegalName: customer["Customer First Name"],
    lastName: customer["Customer Last Name"],
    address: buildAddressToSpec(customer["Home Address"]),
    phone: buildPhoneNumberToSpec(customer["Phone Number"]),
    customerId: customer["Customer ID"].toString()
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
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()

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
    return {status: "OK", message: "OK", referenceId}
  } else {
    if (!trip) {
      log(`${senderAccount.name} attempted to claim invalid trip`, JSON.stringify(response))
      return {status: "400", message: "Trip no longer available", referenceId}
    }
    const rowPosition = trip._rowPosition
    const currentRow = tripSheet.getRange("A" + rowPosition + ":" + rowPosition)
    const shared = trip["Share"]
    const claimed = trip["Claim Pending"]
    if (claimed || (!shared)) {
      log(`${senderAccount.name} attempted to claim unavailable trip`, JSON.stringify(response))
      return {status: "400", message: "Trip no longer available", referenceId}
    }
    // Process successfully pending claim
    const pendingIndex = headers.indexOf("Claim Pending") + 1
    currentRow.getCell(1, pendingIndex).setValue(senderAccount.name)
    currentRow.setBackgroundRGB(0,230,153)
    currentRow.getCell(1,1).setNote('Trip claim pending. Please approve/deny')
    return {status: "OK", message: "OK", referenceId}
  }
}

function sendClientOrderConfirmation(sourceTripRange = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const currentTrip = sourceTripRange ? sourceTripRange : getFullRow(tripSheet.getActiveCell())
  const trip = getRangeValuesAsTable(currentTrip,{includeFormulaValues: false})[0]
  const endPoints = getDocProp("apiGetAccess")
  const endPoint = endPoints.find(endpoint => endpoint.name === trip["Claim Pending"])
  const params = {endpointPath: "/v1/ClientOrderConfirmation"}
  const telegram = {
    tripTicketId: trip["Trip ID"],
    tripConfirmed: true
  }
  try {
    const response = postResource(endPoint, params, JSON.stringify(telegram))
    const responseObject = JSON.parse(response.getContentText())
    if (responseObject.status && responseObject.status !== "OK") {
      logError(`Failure to confirm trip with ${endPoint.name}`, responseObject)
    }
    else {
      // move to sent trips
      log('Trip successfully confirmed', telegram)
      const sentTripSheet = ss.getSheetByName("Sent Trips")
      const claimTime = new Date()
      const tripColumnNames = getSheetHeaderNames(tripSheet)
      const ignoredFields = ["Action", "Go", "Share", "Trip Result", "Driver ID", "Vehicle ID", "Driver Calendar ID", "Trip Event ID", "Declined By", "Shared"]
      const sentTripFields = tripColumnNames.filter(col => !(ignoredFields.includes(col)))
      const sentTripData = {
        "Claimed By" : endPoint.name,
        "Claim Time" : claimTime,
      }
      // "Sched PU Time" : claim.scheduledPickupTime
      sentTripFields.forEach(key => {
        sentTripData[key] = trip[key]
      });
      createRow(sentTripSheet, sentTripData)
      tripSheet.deleteRow(trip._rowPosition)
    } 
  } catch(e) {
    logError(e)
  }
}

// TODO: actually make this work with callsheet fields, ensure it uses the correct sheet
// and gets real data for contact date and note
function sendCustomerReferral(sourceRow = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const customerSheet = ss.getSheetByName("Customers")
  const currentCustomer = sourceRow ? sourceRow : getFullRow(customerSheet.getActiveCell())
  const customer = getRangeValuesAsTable(currentCustomer,{includeFormulaValues: false})[0]
  const customerId = customer["Customer ID"]
  const dateTime = new Date().toISOString();
  const referralId = customerId + ":" + dateTime
  const params = {endpointPath: "/v1/CustomerReferral"}
  const telegram = {
    customerReferralId: referralId,
    customerContactDate: dateTime,
    note: "This is just a test",
    customerInfo: {
      firstLegalName: customer["Customer First Name"],
      lastName: customer["Customer Last Name"],
      address: buildAddressToSpec(customer["Home Address"]),
      phone: buildPhoneNumberToSpec(customer["Phone Number"]),
      customerId: customer["Customer ID"].toString()
    }
  }
  // Get the endpoint (referral provider) from the sheet
  const endPoints = getDocProp("apiGetAccess")
  const endPoint = endPoints[0]
  try {
    const response = postResource(endPoint, params, JSON.stringify(telegram))
    const responseObject = JSON.parse(response.getContentText())
    if (responseObject.status && responseObject.status !== "OK") {
      logError(`Failure to send referral to ${endPoint.name}`, responseObject)
    }
    else {
      log('Referral success', telegram)
    } 
  } catch(e) {
    logError(e)
  }
}

function receiveCustomerReferralResponse(response, senderId) {
  log('Telegram #0B', response)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  return {status: "OK", message: "OK", referenceId} 
}

function getAllTrips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const tripRange = tripSheet.getDataRange()
  const allTrips = getRangeValuesAsTable(tripRange) 
  return allTrips
}