// Provider (this RideSheet instance) sends request for tripRequests to ordering client.
// Received format is an array JSON objects, each element of which complies with
// Telegram 1A of TCRP 210 Transactional Data Spec.
function receiveTripRequest(tripRequest, senderId) {
  log('Received TripRequest', JSON.stringify(tripRequest))
  // Check/Validate Trip Request
  // Add Trip Request to sheet
  // if all goes correctly, return 200 resp
  // if anything goes wrong, send appropriate error
  try {
    const ss              = SpreadsheetApp.getActiveSpreadsheet()
    const tripSheet       = ss.getSheetByName("Outside Trips")
    const apiAccounts = getDocProp("apiGiveAccess")
    const senderAccount = apiAccounts[senderId]
    // Get Header values, set fields accordingly. Log any fields that 
    // are in the tripRequest but not in ridesheet
    // Note: put customerInfo and any fields not in Ridesheet into a string stored in a field
    const supportedFields = [
      tripTicketId,
      pickupAddress,
      dropoffAddress,
      pickupTime,
      dropoffTime,
      customerInfo,
      openAttributes,
      appointmentTime,
      notesForDriver
    ]
    const tripRequestKeys = Object.keys(tripRequest)
    const extraDataFields = tripRequestKeys.filter(key => !supportedFields.includes(key));
    const extraInfo = extraDataFields.reduce((obj, key) => {
      obj[key] = tripRequest[key];
      return obj;
    }, {});
    const tripData = {
      'Trip Date' : formatDateFromTrip(tripRequest.pickupTime, 'M/d/yyyy'),
      'Source' : senderAccount.name,
      'Requested PU Time' : formatDateFromTrip(tripRequest.pickupTime, 'h:mm a'),
      'Requested DO Time' : formatDateFromTrip(tripRequest.dropoffTime, 'h:mm a'),
      'Appt Time' : tripRequest.appointmentTime ? formatDateFromTrip(tripRequest.appointmentTime, 'h:mm a') : '',
      'PU Address' : buildAddressFromSpec(tripRequest.pickupAddress),
      'DO Address' : buildAddressFromSpec(tripRequest.dropoffAddress),
      'Trip ID' : tripRequest.tripTicketId,
      'Est Hours' : tripRequest.openAttributes.estimatedTripDurationInSeconds / (60 * 60 * 24),
      'Est Miles' : tripRequest.openAttributes.estimatedTripDistanceInMiles,
      'Notes' : tripRequest.notesForDriver,
      'Customer Info' : JSON.stringify(tripRequest.customerInfo),
      'Extra Fields' : JSON.stringify(extraInfo),
      'Decline' : false,
      'Claim' : false
    }
    createRow(tripSheet, tripData)
    return {status: "OK"}
  } catch(e) {
    logError(e)
    return {status: "400"}
  } 
}

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
      log("Name",endPoint.name)
      if (endPoint.hasTrips) {
        let params = {resource: "tripRequests"}
        let response = getResource(endPoint, params)
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
          log(responseObject)
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
          formatGroups.time.ranges.push("F" + currentRow + ":J" + (currentRow + responseObject.results.length - 1))
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
          log("No trip requests")
          grid.push(
            [null, null, null, endPoint.name + " responded with no trip requests"].concat(Array(14).fill(null))
          )
          const thisRange = "D" + currentRow + ":" + lastColumnLetter + currentRow
          formatGroups.mergeCells.ranges.push(thisRange)

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
          currentRow += 1
        }
      }
    })
  } catch(e) { logError(e) }
}

// Provider (this RideSheet instance) sends tripRequestResponses
// (whether service for each tripRequest is available or not) to ordering client.
// Send an array of tripRequestResponse JSON objects which comply with
// Telegram 1B of TCRP 210 Transactional Data Spec.
// Receive in response an array of clientOrderConfirmation JSON objects which comply with
// Telegram 2A of TCRP 210 Transactional Data Spec.
function sendTripRequestResponses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    //TODO: Update so that only the relevant endpoint (the source) receives the responses
    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasTrips) {
        let params = {resource: "tripRequestResponses"}

        const trips = getRangeValuesAsTable(ss.getSheetByName("Outside Trips").getDataRange(), {headerRowPosition: 2}).filter(tripRow => {
          return (
            tripRow["Trip Date"] >=  dateToday() &&
            ( tripRow["Decline"]   === true || tripRow["Claim"] === true) &&
            !(tripRow["Decline"]   === true && tripRow["Claim"] === true)
          )
        })
        let payload = trips.map(tripIn => {
          let tripOut = {}
          tripOut.tripAvailable = (tripIn["Claim"] === true)
          if (tripOut.tripAvailable && trips["Scheduled PU Time"]) {
            tripOut["scheduledPickupTime"] = {"@time": combineDateAndTime(tripIn["Trip Date"], tripIn["Scheduled PU Time"])}
          }

          let openAttributes = {}
          openAttributes.tripTicketId = tripIn["Trip ID"]
          tripOut["@openAttribute"] = JSON.stringify(openAttributes)

          return {tripRequestResponse: tripOut}
        })
        // Here we're sending the tripRequestResponses in the payload.
        // responseObject, if validated, contains clientOrderConfirmations
        let response = postResource(endPoint, params, JSON.stringify(payload))
        let responseObject
        try {
          log('Response', response.getContentText())
          responseObject = JSON.parse(response.getContentText())
        } catch(e) {
          logError(e)
          responseObject = {status: "LOCAL_ERROR:" + e.name}
        }
        if (responseObject.status === "OK") {
          // Process clientOrderConfirmations
          let {results} = responseObject
          removeDeclinedTrips()
          processClientOrderResponse(results, endPoint.name)
          sendProviderOrderConfirmations(responseObject, endPoint)
        }
      }
    })  

    return result
  } catch(e) { logError(e) }
}

function testOrderConfirmations() {
  let response = [ { clientOrderConfirmations: 
     { tripAvailable: true,
       tripTicketId: '08eb4058-eae2-4dff-bb2b-f481b8cdf876',
       '@customerId': 31,
       '@openAttributes': '{"estimatedTripDurationInSeconds":10497,"estimatedTripDistanceInMiles":154.316245108}',
       customerMobilePhone: '',
       customerName: 'The Grey, Gandolf',
       pickupWindowStartTime: null,
       pickupWindowEndTime: null,
       negotiatedPickupTime: null,
       pickupTime: { '@time': '2021-10-30T20:31:36.107Z' },
       dropoffTime: { '@time': '2021-10-30T20:31:36.108Z' },
       appointmentTime: { '@time': '2021-10-30T20:31:36.108Z' },
       dropoffAddress: { '@addressName': 'The Castle, Washington 98361, USA' },
       pickupAddress: { '@addressName': '1005 W Burnside St, Portland, OR 97209, USA' } } },
  { customerInfo: 
     { customerId: 31,
       customerName: 'The Grey, Gandolf',
       customerAddress: { '@addressName': '18 S G St, Lakeview, OR 97630, USA' } 
      }  
  }]
  let endpoint = {name: "Agency A"}
  processClientOrderResponse(response, endpoint.name)
}

function removeDeclinedTrips() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const outsideTrips = ss.getSheetByName('Outside Trips')
  const trips = getRangeValuesAsTable(outsideTrips.getDataRange(), {headerRowPosition: 2}).filter(tripRow => tripRow["Decline"] === true)
  safelyDeleteRows(outsideTrips, trips)
}

function processClientOrderResponse(payload, source) {
  if (typeof payload === 'object' && payload.length) {
    payload.forEach(telegram => {
      if (telegram.clientOrderConfirmations) {
        processClientOrderConfirmations(telegram.clientOrderConfirmations, source)
      } else if (telegram.customerInfo) {
        processCustomerInfo(telegram.customerInfo, source)
      } else {
        log('Warning: response type not recognized', telegram )
      }
    })
  }
}

function processClientOrderConfirmations(confirmation, source) {
  // If this trip is a decline (tripAvailable: false), remove it from "Outside trips"
  // Otherwise, move it from Outside Trips to Trips,
  // and add customer if they don't exist yet in the customers table
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const outsideTrips = ss.getSheetByName('Outside Trips')
  const trips = ss.getSheetByName('Trips')
  const tripRange = outsideTrips.getDataRange()
  const allTrips = getRangeValuesAsTable(tripRange, {headerRowPosition: 2}) 
  let trip = allTrips.find(row => row["Trip ID"] === confirmation.tripTicketId)
  let {
    tripAvailable,
    tripTicketId,
    customerMobilePhone,
    customerName,
    pickupWindowStartTime,
    pickupWindowEndTime,
    negotiatedPickupTime,
    pickupTime,
    dropoffTime,
    appointmentTime,
    dropoffAddress,
    pickupAddress
  } = confirmation
  if (!tripAvailable) {
    safelyDeleteRow(outsideTrips, trip)
  }
  // If customer doesn't exist, create them first
  // Create customer ID with source agency appended, to avoid ID conflicts
  // But what if the same customer exists at both agencies in their databases?
  let openAttributes = JSON.parse(confirmation["@openAttributes"])
  let customerID = getOutsideCustomerID(confirmation['@customerId'], source)
  const customers = ss.getSheetByName('Customers')
  const customerRange = customers.getDataRange()
  const allCustomers = getRangeValuesAsTable(customerRange)
  let customer = allCustomers.find(row => row["Customer ID"] === customerID) 
  let customerNames = customerName.split(', ')
  let customerNameAndID = customerName + ' (' + customerID + ')'
  let customerData = {
    'Customer ID' : customerID,
    'Customer First Name' : customerNames[1],
    'Customer Last Name' : customerNames[0],
    'Phone Number' : customerMobilePhone,
    'Customer Name and ID' : customerNameAndID
  }
  if (!customer) {
    createRow(customers, customerData)
  } else {
    let headers = getSheetHeaderNames(customers)
    Object.keys(customerData).forEach(header => {
      let row = customer._rowPosition
      let col = headers.indexOf(header) + 1
      customers.getRange(row, col).setValue(customerData[header])
    });
  }
  pickupTime = negotiatedPickupTime ? negotiatedPickupTime : pickupTime
  let tripData = {
    'Trip Date' : formatDateFromTrip(pickupTime, 'M/d/yyyy'),
    'Source' : source,
    'PU Time' : formatDateFromTrip(pickupTime, 'h:mm a'),
    'DO Time' : formatDateFromTrip(dropoffTime, 'h:mm a'),
    'PU Address' : buildAddressFromSpec(pickupAddress),
    'DO Address' : buildAddressFromSpec(dropoffAddress),
    'Trip ID' : tripTicketId,
    'Customer Name and ID' : customerNameAndID,
    'Earliest PU Time' : pickupWindowStartTime ? formatDateFromTrip(pickupWindowStartTime, 'h:mm a') : '',
    'Latest PU Time' : pickupWindowEndTime ? formatDateFromTrip(pickupWindowEndTime, 'h:mm a') : '',
    'Appt Time' : appointmentTime ? formatDateFromTrip(appointmentTime, 'h:mm a') : '',
    'Est Hours' : openAttributes.estimatedTripDurationInSeconds / (60 * 60 * 24),
    'Est Miles' : openAttributes.estimatedTripDistanceInMiles,
    'Guests' : openAttributes.guests,
    'Notes' : openAttributes.notes,
    'Mobility Factors' : openAttributes.mobilityFactors
  }
  createRow(trips, tripData)
  safelyDeleteRow(outsideTrips, trip)
}

//TODO: Make a better version of this function
// Takes an ID from another agency and prefixes
// Question: does Customer ID need to be an integer?
function getOutsideCustomerID(id, source) {
  return getAgencyPrefix(source) + ':' + id.toString()
}

function getAgencyPrefix(name) {
  let agencyWords = name.split(' ')
  if (agencyWords.length > 1) {
    return agencyWords.map(word => word[0]).join('')
  } 
  if (name.length > 2) {
    return name.slice(0,3).toUpperCase()
  }
  return name.toUpperCase()
}

// TODO: coordinate on additional customer info. 
function processCustomerInfo(customerInfo, source) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const customers = ss.getSheetByName('Customers')
  const customerRange = customers.getDataRange()
  const allCustomers = getRangeValuesAsTable(customerRange)
  let customerID = getOutsideCustomerID(customerInfo.customerId, source)
  let customer = allCustomers.find(row => row["Customer ID"] === customerID) 
  let {
    customerName,
    customerAddress,
    customerPhone,
    customerMobilePhone,
    gender,
    caregiverContactInformation,
    customerEmergencyPhoneNumber,
    customerEmergencyContactName,
    requiredCareComments,
    notesForDriver
  } = customerInfo
  let customerNameAndID = customerName + ' (' + customerID + ')'
  let customerNames = customerName.split(', ')
  let customerData = {
    'Customer Name and ID' : customerNameAndID,
    'Customer ID' : customerID,
    'Customer First Name' :  customerNames[1],
    'Customer Last Name' : customerNames[0],
    'Phone Number' : customerPhone ? buildPhoneNumberFromSpec(customerPhone) : "",
    'Home Address' : customerAddress? buildAddressFromSpec(customerAddress) : "",
    'Customer Manifest Notes' : notesForDriver
  }
  if (!customer) {
    createRow(customers, customerData)
  } else {
    let headers = getSheetHeaderNames(customers)
    Object.keys(customerData).forEach(header => {
      let row = customer._rowPosition
      let col = headers.indexOf(header) + 1
      customers.getRange(row, col).setValue(customerData[header])
    });
  }
}

// Send an array of providerOrderConfirmation JSON objects which comply with
// Telegram 2B of TCRP 210 Transactional Data Spec.
// Receive in response an array of customerInfo JSON objects which comply with
// Telegram 2A1 of TCRP 210 Transactional Data Spec.
function sendProviderOrderConfirmations(orderingClientResponses, endPoint) {
  if (endPoint.hasTrips) {
    let params = {resource: "providerOrderConfirmations"}
    let payload = {status: "OK"}

    let response = postResource(endPoint, params, payload)
    let responseObject
    try {
      responseObject = JSON.parse(response.getContentText())
    } catch(e) {
      logError(e)
      responseObject = {status: "LOCAL_ERROR:" + e.name}
    }
    if (responseObject.status === "OK") {
      // Process the customerInfo objects
      
    }
  }
}
