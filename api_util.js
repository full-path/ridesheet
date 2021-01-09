function buildAddressToSpec(address) {
  try {
    let result = {}
    const parsedAddress = parseAddress(address)
    result["@addressName"] = parsedAddress.geocodeAddress
    if (parsedAddress.parenText) {
      let manualAddr = {}
      manualAddr["@manualText"] = parsedAddress.parenText
      manualAddr["@sendtoInvoice"] = true
      manualAddr["@sendtoVehicle"] = true
      manualAddr["@sendtoOperator"] = true
      manualAddr["@vehicleConfirmation"] = false
      result.manualDescriptionAddress = manualAddr
    }
    return result
  } catch(e) { logError(e) }
}

function buildAddressFromSpec(address) {
  try {
    let result = address["@addressName"]
    if (address.manualDescriptionAddress) {
      result = result + " (" + address.manualDescriptionAddress["@manualText"] + ")"
    }
    return result
  } catch(e) { logError(e) }
}