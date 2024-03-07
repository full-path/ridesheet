// Provider (this RideSheet instance) sends request for tripRequests to ordering client.
// Received format is an array JSON objects, each element of which complies with
// Telegram 1A of TCRP 210 Transactional Data Spec.
function receiveTripRequest(tripRequest, senderId) {
  log('Received TripRequest', JSON.stringify(tripRequest))
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet = ss.getSheetByName("Outside Trips")
    const apiAccounts = getDocProp("apiGiveAccess")
    const senderAccount = apiAccounts[senderId]

    const supportedFields = [
      "tripTicketId",
      "pickupAddress",
      "dropoffAddress",
      "pickupTime",
      "dropoffTime",
      "customerInfo",
      "openAttributes",
      "appointmentTime",
      "notesForDriver",
      "numOtherReservedPassengers"
    ]
    const tripRequestKeys = Object.keys(tripRequest)
    const extraDataFields = tripRequestKeys.filter(key => !supportedFields.includes(key));
    const extraInfo = extraDataFields.reduce((obj, key) => {
      obj[key] = tripRequest[key];
      return obj;
    }, {});
    // TODO: Format hours and miles? Either here or in the sheet
    const tripData = {
      'Source' : senderAccount.name,
      'PU Address' : buildAddressFromSpec(tripRequest.pickupAddress),
      'DO Address' : buildAddressFromSpec(tripRequest.dropoffAddress),
      'Trip ID' : tripRequest.tripTicketId,
      'Notes' : tripRequest.notesForDriver,
      'Customer Info' : JSON.stringify(tripRequest.customerInfo),
      'Extra Fields' : JSON.stringify(extraInfo),
      'Guests' : tripRequest.numOtherReservedPassengers,
      'Decline' : false,
      'Claim' : false
    }
    let dateFlag = false
    // Time fields are not required, check for at least one to get the trip date from
    if (tripRequest.pickupTime && tripRequest.pickupTime.time) {
      tripData["Trip Date"] = formatDate(tripRequest.pickupTime.time, null, 'M/d/yyyy')
      tripData["Requested PU Time"] = formatDate(tripRequest.pickupTime.time, null, 'h:mm a')
      dateFlag = true
    }
    if (tripRequest.dropoffTime && tripRequest.dropoffTime.time) {
      if (!dateFlag) {
        tripData["Trip Date"] = formatDate(tripRequest.dropoffTime.time, null, 'M/d/yyyy')
        dateFlag = true
      }
      tripData["Requested DO Time"] = formatDate(tripRequest.dropoffTime.time, null, 'h:mm a')
    }
    if (tripRequest.appointmentTime && tripRequest.appointmentTime.time) {
      if (!dateFlag) {
        tripData["Appt Time"] = formatDate(tripRequest.appointmentTime.time, null, 'M/d/yyyy')
        dateFlag = true
      }
      tripData["Requested DO Time"] = formatDate(tripRequest.appointmentTime.time, null, 'h:mm a')
    }

    if (tripRequest.openAttributes) {
      if (tripRequest.openAttributes.estimatedTripDurationInSeconds) {
        tripData["Est Hours"] = tripRequest.openAttributes.estimatedTripDurationInSeconds / (60 * 60 * 24)
      }
      if (tripRequest.openAttributes.estimatedTripDistanceInMiles) {
        tripData["Est Miles"] = tripRequest.openAttributes.estimatedTripDistanceInMiles
      }
    }
    createRow(tripSheet, tripData)
    return {status: "OK", message: "OK", referenceId}
  } catch(e) {
    logError(e)
    return {status: "400", message: "Unknown Error: Check Logs", referenceId}
  } 
}

function sendTripRequestResponses() {
  try { 
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const outsideTrips = ss.getSheetByName("Outside Trips")
    const trips = getRangeValuesAsTable(outsideTrips.getDataRange()).filter(tripRow => {
      return (
        tripRow["Trip Date"] >=  dateToday() &&
        ( tripRow["Decline"]  === true || tripRow["Claim"] === true) &&
        !(tripRow["Decline"]  === true && tripRow["Claim"] === true) &&
        !(tripRow["Pending"] === 'TRUE') 
      )
    })
    // For each trip: Check the 'Source column' and find the appropriate agency
    const endPoints = getDocProp("apiGetAccess")
    trips.forEach((trip) => {
      const source = trip["Source"]
      const endpoint = endPoints.find(endpoint => endpoint.name === source)
      const params = {endpointPath: "/v1/TripRequestResponse"}
      const claimed = trip["Claim"] === true
      const telegram = {
        tripTicketId: trip["Trip ID"],
        tripAvailable: claimed
      }
      if (trip["Scheduled PU Time"]) {
        telegram.scheduledPickupTime = {
          time: combineDateAndTime(trip["Trip Date"], trip["Scheduled PU Time"])
        }
      }
      try {
        const response = postResource(endpoint, params, JSON.stringify(telegram))
        const responseObject = JSON.parse(response.getContentText())
          log('#1B response', responseObject)
          // TODO: Handle the case where the 400-error is on a trip decline
          if (responseObject.status && responseObject.status !== "OK") {
            log('Claim Trip Rejected', JSON.stringify(responseObject))
            markTripAsFailure(outsideTrips, trip)
          } else {
            if (!claimed) {
              safelyDeleteRow(outsideTrips, trip)
            } else {
              markTripAsPending(outsideTrips, trip)
            }
          }
      } catch (e) {
        // TODO: Figure out logic to determine if "e" contains a 400-level error
        logError(e)
      }
    })
  } catch (e) {
    logError(e)
  }
}

function markTripAsPending(sheet, tripRow) {
  const headers = getSheetHeaderNames(sheet)
  const headerPosition = headers.indexOf("Pending") + 1
  const rowPosition = tripRow._rowPosition
  const currentRow = sheet.getRange("A" + rowPosition + ":" + rowPosition)
  currentRow.setBackgroundRGB(20,204,204)
  currentRow.getCell(1,1).setNote('Trip claim pending')
  currentRow.getCell(1, headerPosition).setValue("TRUE")
}

function markTripAsFailure(sheet, tripRow) {
  const rowPosition = tripRow._rowPosition
  const currentRow = sheet.getRange("A" + rowPosition + ":" + rowPosition)
  currentRow.setBackgroundRGB(255,102,102)
  currentRow.getCell(1,1).setNote('Claim rejected. See log for details.')
}

// Receive telegram #2A from ordering client
function receiveClientOrderConfirmation(confirmation) {
  const {tripTicketId, tripConfirmed} = confirmation
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const outsideTrips = ss.getSheetByName("Outside Trips")
  const trips = getRangeValuesAsTable(outsideTrips.getDataRange())
  const trip = trips.find(row => row["Trip ID"] === tripTicketId)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  if (!tripConfirmed) {
    markTripAsFailure(outsideTrips, trip)
    log('Trip not confirmed', confirmation)
    return {status: "OK", message: "OK", referenceId}
  }
  const customerInfo = JSON.parse(trip["Customer Info"])
  const customerSuccess = addCustomer(customerInfo)
  if (!customerSuccess.status) {
    logError('Error confirming trip. Customer Info invalid.', trip)
    return {status: "400", message: customerSuccess.message, referenceId}
  }
  // Move into trips
  const tripSheet = ss.getSheetByName("Trips")
  const tripColumnNames = getSheetHeaderNames(outsideTrips)
  const ignoredFields = ["Scheduled PU Time", "Requested PU Time", "Requested DO Time", "Decline", "Claim", "Customer Info", "Pending", "Extra Fields"]
  const tripFields = tripColumnNames.filter(col => !(ignoredFields.includes(col)))
  const tripData = {}
  tripFields.forEach(key => {
   tripData[key] = trip[key]
  });
  tripData['Customer Name and ID'] = customerSuccess.customerNameAndID
  tripData['Customer ID'] = customerSuccess.customerId
  tripData['PU Time'] = trip['Requested PU Time']
  tripData['DO Time'] = trip['Requested DO Time']
  if (trip['Scheduled PU Time']) {
    tripData['PU Time'] = trip['Scheduled PU Time']
  }
  // TODO: incorrect/incomplete data is moved over, correct trip is not deleted
  createRow(tripSheet, tripData)
  outsideTrips.deleteRow(trip._rowPosition)
  return {status: "OK", message: "OK", referenceId}
}

function addCustomer(customerInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const customers = ss.getSheetByName('Customers')
  const customerRange = customers.getDataRange()
  const allCustomers = getRangeValuesAsTable(customerRange)
  const { customerId, firstLegalName, lastName } = customerInfo
  const referralId = `Ref:${customerId}`
  let customer = allCustomers.find(row => row["Customer ID"] === referralId)
  let customerNameAndID = lastName + ', ' + firstLegalName + ' (' + customerId + ')'
  if (!customer) {
    const newCustomer = {
      'Customer ID': referralId,
      'Customer First Name': firstLegalName,
      'Customer Last Name': lastName,
      'Home Address': buildAddressFromSpec(customerInfo['address']),
      'Customer Name and ID' : customerNameAndID
    }
    // TODO: add all the fields!!
    if (customerInfo['mobilePhone'] && customerInfo['phone']) {
      newCustomer['Phone Number'] = buildPhoneNumberFromSpec(customerInfo['mobilePhone'])
      newCustomer['Alt. Phone'] = buildPhoneNumberFromSpec(customerInfo['phone'])
    } else if (customerInfo['mobilePhone']) {
      newCustomer['Phone Number'] = buildPhoneNumberFromSpec(customerInfo['mobilePhone'])
    } else if (customerInfo['phone']) {
      newCustomer['Phone Number'] = buildPhoneNumberFromSpec(customerInfo['phone'])
    } else {
      return {status: false, message: 'Missing phone number'}
    }
    createRow(customers, newCustomer)
  } else {
    log('Customer ID already exists', customerInfo)
  }
  return {status: true, customerId: referralId, customerNameAndID}
}

// Handle sending all confirmations from menu trigger
function sendProviderOrderConfirmations() {
  try { 
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet = ss.getSheetByName("Trips")
    const trips = getRangeValuesAsTable(tripSheet.getDataRange()).filter(tripRow => {
      return (
        tripRow["TDS Actions"] ===  'Confirm scheduled trip'
      )
    })
    trips.forEach((trip) => {
      sendProviderOrderConfirmation(trip)
    })
  } catch (e) {
    logError(e)
  }
}

// Telegram #2B
function sendProviderOrderConfirmation(sourceTrip = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  const trip = sourceTrip ? sourceTrip : getRangeValuesAsTable(getFullRow(tripSheet.getActiveCell()),{includeFormulaValues: false})[0]
  const endPoints = getDocProp("apiGetAccess")
  const endPoint = endPoints.find(endpoint => endpoint.name === trip["Source"])
  const params = {endpointPath: "/v1/ProviderOrderConfirmation"}
  if (!endpoint) {
    ss.toast("Attempting to send confirmation without valid source")
    logError("Invalid provider order confirmation", trip)
    return
  }
  const telegram = {
    tripTicketId: trip["Trip ID"],
    scheduledPickupTime: {time: combineDateAndTime(trip["Trip Date"], trip["PU Time"])},
    scheduledDropoffTime: {time: combineDateAndTime(trip["Trip Date"], trip["DO Time"])},
    scheduledPickupPoint: buildAddressToSpec(trip["PU Address"]),
    scheduledDropoffPoint: buildAddressToSpec(trip["DO Address"]),
    driverName: trip['Driver']
  }
  const allVehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())
  const vehicle = allVehicles.find(row => row["Vehicle ID"] === trip["Vehicle ID"])
  telegram.vehicleInformation = vehicle['Vehicle Name']
  telegram.hasRamp = vehicle['Has Ramp']
  telegram.hasLift = vehicle['Has Lift']
  try {
    const response = postResource(endPoint, params, JSON.stringify(telegram))
    const responseObject = JSON.parse(response.getContentText())
    if (responseObject.status && responseObject.status !== "OK") {
      logError(`Failure to send trip confirmation to ${endPoint.name}`, responseObject)
    }
  } catch(e) {
    // TODO: figure out how to get message from 400 response
    logError(e)
  }
}

// TODO: Add support for all fields, in particular, add place in sheet to handle
// customerReferralId, customerContactDate, and customerInfo
// Question: Are referrals going to their own tab, rather than into the main customers sheet
// initially?
function receiveCustomerReferral(customerReferral, senderId) {
  log('Telegram #0A', customerReferral)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  const {customerReferralId, customerContactDate, customerInfo} = customerReferral
  const customerSuccess = addCustomer(customerInfo)
  if (!customerSuccess.status) {
    logError('Error confirming trip. Customer Info invalid.', trip)
    return {status: "400", message: customerSuccess.message, referenceId}
  }
  return {status: "OK", message: "OK", referenceId} 
}

// TODO: Talk over what we want to happen here. I assume we will
// be adding a referrals sheet, which can have options
function sendCustomerReferralResponse(customerRow = null) {
  // Get response value
  // Get correct endpoint (original sender)
  // 
}

// TODO: Finish function
function receiveTripStatusChange(tripStatusChange, senderId) {
  log('Telegram #1C', tripStatusChange)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  return {status: "OK", message: "OK", referenceId} 
}

function removeDeclinedTrips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const outsideTrips = ss.getSheetByName('Outside Trips')
  const trips = getRangeValuesAsTable(outsideTrips.getDataRange(), {headerRowPosition: 2}).filter(tripRow => tripRow["Decline"] === true)
  safelyDeleteRows(outsideTrips, trips)
}
