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
  const tripColumnNames = getSheetHeaderNames(OutisdeTrips)
  const ignoredFields = ["Scheduled PU Time", "Decline", "Claim", "Customer Info", "Pending", "Extra Fields"]
  const tripFields = tripColumnNames.filter(col => !(ignoredFields.includes(col)))
  const tripData = {}
  tripFields.forEach(key => {
   tripData[key] = trip[key]
  });
  if (trip['Scheduled PU Time']) {
    tripData['PU Time'] = trip['Scheduled PU Time']
  }
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
      newCustomer['Phone Number'] = customerInfo['mobilePhone']
      newCustomer['Alt. Phone'] = customerInfo['phone']
    } else if (customerInfo['mobilePhone']) {
      newCustomer['Phone Number'] = customerInfo['mobilePhone']
    } else if (customerInfo['phone']) {
      newCustomer['Phone Number'] = customerInfo['phone']
    } else {
      return {status: false, message: 'Missing phone number'}
    }
    createRow(customers, newCustomer)
  } else {
    log('Customer ID already exists', customerInfo)
  }
  return {status: true}
}

function receiveCustomerReferral(customerReferral, senderId) {
  log('Telegram #0A', customerReferral)
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  return {status: "OK", message: "OK", referenceId} 
}

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
