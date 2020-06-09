const triggers = {
  "codeFillRequestCells": fillRequestCells,
  "codeFormatAddress": formatAddress,
  "codeFillHoursAndMiles": fillHoursAndMiles,
  "codeSetCustomerKey": setCustomerKey,
  "codeScanForDuplicates": scanForDuplicates
}

/**
 * The event handler triggered when editing the spreadsheet.
 * @param {event} e The onEdit event.
 */
function onEdit(e) {
  let startTime = new Date()
  let allNamedRanges = e.source.getNamedRanges()
  let involvedNamedRanges = []
  
  allNamedRanges.forEach((namedRange) => {
    if (namedRange.getName().indexOf("code") === 0) {
      if (isInRange(e.range, namedRange.getRange())) {
        involvedNamedRanges.push(namedRange)
      }
    }
  })

  involvedNamedRanges.forEach(
    (namedRange) => {
    // log("Entering namedRange " + namedRange.getName())
    Object.keys(triggers).forEach(
      (triggerName) => {
    // log("Entering triggerName " + triggerName)
    if (namedRange.getName().indexOf(triggerName) > -1) {
      //log("Triggering " + triggerName)
      triggers[triggerName](e)
      //log("Triggered " + triggerName)
    }
    // log("Exiting triggerName " + triggerName)
    })
    // log("Exiting namedRange " + namedRange)
  })
  log("onEdit duration:",(new Date()) - startTime)
}

function formatAddress(e) {
  if (e.value) {
    let formattedAddress = getGeocode(e.value,"formatted_address")
    if (formattedAddress.includes("Error")) {
      e.range.setNote("Address " + formattedAddress)
    } else {
      e.range.setValue(formattedAddress)
    } 
  }
}

function fillRequestCells(e) {
  if (e.value) {
    const tripRow = getFullRow(e.range)
    const tripValues = getValuesByHeaderNames(["Customer Name and ID","PU Address","DO Address","Service ID"], tripRow)
    const customerRow = findFirstRowByHeaderNames({"Customer Name and ID": tripValues["Customer Name and ID"]}, e.source.getSheetByName("Customers"))
    const customerAddresses = getValuesByHeaderNames(["Customer ID","Home Address","Default Destination","Default Service ID"], customerRow)
    let valuesToChange = {}
    valuesToChange["Customer ID"] = customerAddresses["Customer ID"]
    if (tripValues["PU Address"] == '') { valuesToChange["PU Address"] = customerAddresses["Home Address"] }
    if (tripValues["DO Address"] == '') { valuesToChange["DO Address"] = customerAddresses["Default Destination"] }
    if (tripValues["Service ID"] == '') { valuesToChange["Service ID"] = customerAddresses["Default Service ID"] }
    setValuesByHeaderNames(valuesToChange, tripRow)
    if (valuesToChange["PU Address"] || valuesToChange["DO Address"]) { fillHoursAndMiles(e) }
  }
}

function fillHoursAndMiles(e) {
  const tripRow = getFullRow(e.range)
  const values = getValuesByHeaderNames(["PU Address", "DO Address"], tripRow)
  let tripEstimate
  if (values["PU Address"] && values["DO Address"]) {
    tripEstimate = getTripEstimate(values["PU Address"], values["DO Address"], "milesAndDays")
    setValuesByHeaderNames({"Est Hours": tripEstimate["days"], "Est Miles": tripEstimate["miles"]}, tripRow)
  }
}

/**
 * Manage setup of a new customer record. The goals here are to:
 * - Trim the customer name as needed 
 * _ Generate a customer ID when it's missing and there's a first and last name present
 * - Autofill the "Customer Name and ID" field when the first name, last name, and ID are present. 
 *   This will be the field used to identify the customer in trip records
 * - Keep track of the current highest customer ID in document properties, seeding data when needed
 */
function setCustomerKey(e) {
  const customerRow = getFullRow(e.range)
  const customerValues = getValuesByHeaderNames(["Customer First Name", "Customer Last Name", "ID", "Customer Name and ID"], customerRow)
  let newValues = {}
  if (customerValues["Customer First Name"] && customerValues["Customer Last Name"]) {
    const docProperties = PropertiesService.getDocumentProperties()
    const lastCustomerID = docProperties.getProperty("lastCustomerID")
    let nextCustomerID = ((lastCustomerID && (+lastCustomerID)) ? (Math.ceil(+lastCustomerID) + 1) : 1 )
    if (!customerValues["Customer ID"]) { // There is no ID. Set one and update the lastCustomerID property
      newValues["Customer ID"] = nextCustomerID
      newValues["Customer First Name"] = customerValues["Customer First Name"].trim()
      newValues["Customer Last Name"] = customerValues["Customer Last Name"].trim()
      newValues["Customer Name and ID"] = getCustomerNameAndId(newValues["Customer First Name"], newValues["Customer Last Name"], newValues["Customer ID"])
      docProperties.setProperty("lastCustomerID", nextCustomerID)
    } else if (+customerValues["Customer ID"]) { // There is an ID value, and it's numeric
      newValues["Customer ID"] = (+customerValues["ID"])
      newValues["Customer First Name"] = customerValues["Customer First Name"].trim()
      newValues["Customer Last Name"] = customerValues["Customer Last Name"].trim()
      newValues["Customer Name and ID"] = getCustomerNameAndId(newValues["Customer First Name"], newValues["Customer Last Name"], newValues["Customer ID"])
      if ((+customerValues["Customer ID"]) >= nextCustomerID) { docProperties.setProperty("lastCustomerID", customerValues["ID"]) }
    } else { // There is an ID value, and it's not numeric. Allow this, but don't track it as the lastCustomerID
      newValues["Customer First Name"] = customerValues["Customer First Name"].trim()
      newValues["Customer Last Name"] = customerValues["Customer Last Name"].trim()
      newValues["Customer Name and ID"] = getCustomerNameAndId(newValues["Customer First Name"], newValues["Customer Last Name"], newValues["Customer ID"])
    }
    setValuesByHeaderNames(newValues, customerRow)
  }
}

function scanForDuplicates(e) {
  const thisRowNumber = e.range.getRow()
  const range = e.range.getSheet().getRange(1, e.range.getColumn(), e.range.getSheet().getLastRow())
  const values = range.getValues().map(row => row[0])
  let duplicateRows = []
  values.forEach((value, i) => {
    if (value == e.value && (i + 1) != thisRowNumber) { duplicateRows.push(i + 1) }
  })
  if (duplicateRows.length == 1) { e.range.setNote("This value is already used in row " + duplicateRows[0]) }
  if (duplicateRows.length > 1) { e.range.setNote("This value is already used in rows " + duplicateRows.join(", ")) }
  if (duplicateRows.length == 0) { e.range.clearNote() }
}

function getCustomerNameAndId(first, last, id) {
  return `${last}, ${first} (${id})`
}