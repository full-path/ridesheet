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
        return tripRow["Trip Date"] >= dateToday() && 
               tripRow["Source"] === "" && 
               tripRow["Shared"] === "" &&
               (tripRow["Share With"] === endPoint.name || tripRow["Share With"] === 'All')
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
  if (trip["Earliest PU Time"]) {
    formattedTrip.pickupWindowStartTime = {time: combineDateAndTime(trip["Trip Date"], trip["Earliest PU Time"])}
  }
  if (trip["Latest PU Time"]) {
    formattedTrip.pickupWindowEndTime = {time: combineDateAndTime(trip["Trip Date"], trip["Latest PU Time"])}
  }
  if (trip["Transfer Trip"]) {
    formattedTrip.tripTransfer = trip["Transfer Trip"]
  }
  if (trip["Pickup Location Notes"]) {
    formattedTrip.detailedPickupLocationDescription = trip["Pickup Location Notes"]
  }
  if (trip["Dropoff Location Notes"]) {
    formattedTrip.detailedDropoffLocationDescription = trip["Dropoff Location Notes"]
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
  if (customer["Date of Birth"]) {
    formattedCustomer.dateOfBirth = customer["Date of Birth"]
  } 
  if (customer["Customer Manifest Notes"]) {
    formattedCustomer.notesForDriver = customer["Customer Manifest Notes"]
  }
  if (customer["Default Service ID"]) {
    formattedCustomer.fundingEntityId = customer["Default Service ID"]
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
    const shared = trip["Share With"]
    const claimed = trip["Claim Pending"]
    if (claimed || !(shared === senderAccount.name)) {
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

// Handle sending all confirmations from menu trigger
function sendClientOrderConfirmations() {
  try { 
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet = ss.getSheetByName("Trips")
    const trips = getRangeValuesAsTable(tripSheet.getDataRange()).filter(tripRow => {
      return (
        tripRow["TDS Actions"] ===  'Confirm pending claim' ||
        tripRow["TDS Actions"] === 'Deny pending claim'
      )
    })
    trips.forEach((trip) => {
      sendClientOrderConfirmation(trip)
    })
  } catch (e) {
    logError(e)
  }
}

// Telegram #2A
function sendClientOrderConfirmation(sourceTrip = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const trip = sourceTrip ? sourceTrip : getRangeValuesAsTable(getFullRow(tripSheet.getActiveCell()),{includeFormulaValues: false})[0]
  const endPoints = getDocProp("apiGetAccess")
  const endPoint = endPoints.find(endpoint => endpoint.name === trip["Claim Pending"])
  const params = {endpointPath: "/v1/ClientOrderConfirmation"}
  const telegram = {
    tripTicketId: trip["Trip ID"],
  }
  if (trip["TDS Actions"] === "Confirm pending claim") {
    telegram.tripConfirmed = true
  } else if (trip["TDS Actions" === "Deny pending claim"]) {
    telegram.tripConfirmed = false
  } else {
    ss.toast("Attempting to send client order confirmation for invalid trip")
    logError("Invalid client order confirmation", trip)
    return
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
      const ignoredFields = ["Action", "Go", "Trip Result", "Driver ID", "Vehicle ID", "Driver Calendar ID", "Trip Event ID", "Declined By"]
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

// Handle sending all confirmations from menu trigger
function sendTripCancelations() {
  try { 
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet = ss.getSheetByName("Trips")
    const trips = getRangeValuesAsTable(tripSheet.getDataRange()).filter(tripRow => {
      return (
        tripRow["TDS Actions"] ===  'Cancel Trip'
      )
    })
    trips.forEach((trip) => {
      sendTripStatusChange(trip)
    })
  } catch (e) {
    logError(e)
  }
}

function sendTripStatusChange(sourceTrip = null, tripResult = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const trip = sourceTrip ? sourceTrip : getRangeValuesAsTable(getFullRow(tripSheet.getActiveCell()),{includeFormulaValues: false})[0]
  const endPoints = getDocProp("apiGetAccess")
  // We can send either as ordering client or provider
  const endPoint = endPoints.find(endpoint =>
     endpoint.name === trip["Claim Pending"] ||
     endpoint.name === trip["Source"]
     )
  if (!endPoint) {
    ss.toast('Canceled trip is not a referral')
    log('Attempted to send a trip status change for invalid trip', trip)
    return
  }
  // TODO: find a way in the UI to indicate who is canceling. Need the option of "Rider"
  const params = {endpointPath: "/v1/TripStatusChange"}
  const telegram = {
    tripTicketId: trip["Trip ID"],
    status: "Cancel"
  }
  if (tripResult) {
    telegram.canceledBy = "Rider"
    telegram.reasonDescription = tripResult
  } else if (trip["Claim Pending"]) {
    telegram.canceledBy = "Ordering Client"
  } else {
    telegram.canceledBy = "Provider"
  }
  try {
    const response = postResource(endPoint, params, JSON.stringify(telegram))
    const responseObject = JSON.parse(response.getContentText())
    if (responseObject.status && responseObject.status !== "OK") {
      logError(`Failure to cancel trip with ${endPoint.name}`, responseObject)
    }
    else {
      log(`Trip canceled with ${endPoint.name}`, telegram)
    } 
  } catch(e) {
    logError(e)
  }
}

function sendCustomerReferrals() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const refSheet = ss.getSheetByName("TDS Referrals")
    const referrals = getRangeValuesAsTable(refSheet.getDataRange()).filter(row => {
      return (
        (!!row["Agency"]) && !row["Referral ID"]
      )
    })
    referrals.forEach((referral) => {
      sendCustomerReferral(referral)
    })
  } catch (e) {
    logError(e)
  }
}

function sendCustomerReferral(sourceRow = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const referralSheet = ss.getSheetByName("TDS Referrals")
  const referral = sourceRow ? sourceRow : getRangeValuesAsTable(getFullRow(referralSheet.getActiveCell()),{includeFormulaValues: false})[0]
  const customerId = referral["Customer ID"] || getCustomerId(referral)
  const referralDateString = formatDate(referral["Referral Date"],null,"yyyy-MM-dd")
  const agencyCode = referral["Agency"].replace(/[^A-Za-z]/g, '').toLowerCase()
  const referralId = `${customerId}:${agencyCode}:${referralDateString}`
  const params = {endpointPath: "/v1/CustomerReferral"}
  const telegram = {
    customerReferralId: referralId,
    customerContactDate: referral["Referral Date"],
    note: referral["Referral Notes"],
    customerInfo: {
      firstLegalName: referral["Customer First Name"],
      middleName: referral["Customer Middle Name"],
      lastName: referral["Customer Last Name"],
      nickName: referral["Customer Nickname"],
      address: buildAddressToSpec(referral["Home Address"]),
      phone: buildPhoneNumberToSpec(referral["Home Phone Number"]),
      mobilePhone: buildPhoneNumberToSpec(referral["Cell Phone Number"]),
      mailingBillingAddress: buildAddressToSpec(referral["Mailing Address"]),
      fundingEntityBillingInformation: referral["Billing Information"],
      fundingType: valueToBoolean(referral["Funding Type?"]),
      fundingEntityId: referral["Funding Entity"],
      gender: referral["Gender"],
      lowIncome: valueToBoolean(referral["Low Income?"]),
      disability: valueToBoolean(referral["Disability?"]),
      language: referral["Language"],
      race: referral["Race"],
      ethnicity: referral["Ethnicity"] || null,
      emailAddress: referral["Email"],
      veteran: valueToBoolean(referral["Veteran?"]),
      caregiverContactInformation: referral["Caregiver Contact Information"],
      emergencyPhoneNumber: referral["Emergency Phone Number"],
      emergencyContactName: referral["Emergency Contact Name"],
      emergencyContactRelationship: referral["Emergency Contact Relationship"]  ,
      requiredCareComments: referral["Comments About Care Required"],
      dateOfBirth: referral["Date of Birth"],
      customerId: customerId
    }
  }

  // Get the endpoint (referral provider) from the sheet
  const endPoints = getDocProp("apiGetAccess")
  const endPoint = endPoints.find(endpoint => endpoint.name === referral["Agency"])
  try {
    const response = postResource(endPoint, params, JSON.stringify(telegram))
    const responseObject = JSON.parse(response.getContentText())
    if (responseObject.status && responseObject.status !== "OK") {
      ss.toast(`Failure to send referral to ${endPoint.name}`, "Failure")
      logError(`Failure to send referral to ${endPoint.name}`, responseObject)
      return false
    } else {
      setValuesForRow({
        "Customer ID": customerId,
        "Referral ID": referralId,
        "Referral Sent Timestamp": new Date()
      }, referral._rowPosition, referralSheet)
      ss.toast(`Referral sent to ${endPoint.name}`, "Success")
      log('Referral success', telegram)
      return true
    } 
  } catch(e) {
    logError(e)
    return false
  }
}

function receiveCustomerReferralResponse(response, senderId) {
  log('Telegram #0B', response)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  try {
    const { customerReferralId, referralResponseType, note } = response
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const referralSheet = ss.getSheetByName("TDS Referrals")
    const referrals = getRangeValuesAsTable(referralSheet.getDataRange())
    const referral = referrals.find(row => row["Referral ID"] === customerReferralId)
    setValuesForRow({
      "Referral Response Timestamp": new Date(),
      "Referral Response": referralResponseType,
      "Response Notes": note
    }, referral._rowPosition, referralSheet)
    return {status: "OK", message: "OK", referenceId}
  } catch(e) {
    return {status: "400", message: "Unknown error", referenceId}
  }
}

function receiveTripTaskCompletion(tripTaskCompletion) {
  log('Telegram #4A', tripTaskCompletion)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  const { tripTicketId } = tripTaskCompletion
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sentTrips = ss.getSheetByName("Sent Trips")
  const trips = getRangeValuesAsTable(sentTrips.getDataRange())
  const trip = trips.find(row => row["Trip ID"] === tripTicketId)
  const rowPosition = trip._rowPosition
  const currentRow = sentTrips.getRange("A" + rowPosition + ":" + rowPosition)
  const headers = getSheetHeaderNames(sentTrips)
  const statusIndex = headers.indexOf("Status") + 1
  currentRow.getCell(1, statusIndex).setValue("Completed")
  return {status: "OK", message: "OK", referenceId} 
}

// Do we have any reason to track the other possible returned values (driver and vehicle info?)
function receiveProviderOrderConfirmation(response, senderId) {
  log('Telegram #2B', response)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  const { tripTicketId } = response
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sentTrips = ss.getSheetByName("Sent Trips")
  const trips = getRangeValuesAsTable(sentTrips.getDataRange())
  const trip = trips.find(row => row["Trip ID"] === tripTicketId)
  const rowPosition = trip._rowPosition
  const currentRow = sentTrips.getRange("A" + rowPosition + ":" + rowPosition)
  const headers = getSheetHeaderNames(sentTrips)
  const statusIndex = headers.indexOf("Status") + 1
  currentRow.getCell(1, statusIndex).setValue("Scheduled")
  return {status: "OK", message: "OK", referenceId}
}

function getAllTrips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const tripRange = tripSheet.getDataRange()
  const allTrips = getRangeValuesAsTable(tripRange) 
  return allTrips
}
