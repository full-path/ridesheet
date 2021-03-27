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

function buildTimeFromSpec(date, time) {
  return {"@time": combineDateAndTime(date, time)}
}

function urlQueryString(params) {
  try {
    const keys = Object.keys(params)
    result = keys.reduce((a, key, i) => {
      if (i === 0) {
        return key + "=" + params[key]
      } else {
        return a + "&" + key + "=" + params[key]
      }
    },"")
    return result
  } catch(e) { logError(e) }
}

function byteArrayToHexString(bytes) {
  try {
    return bytes.map(byte => {
      return ("0" + (byte & 0xFF).toString(16)).slice(-2)
    }).join('')
  } catch(e) { logError(e) }
}

function generateHmacHexString(secret, nonce, timestamp, params) {
  try {
    let orderedParams = {}
    Object.keys(params).sort().forEach(key => {
      orderedParams[key] = params[key]
    })
    const value = [nonce, timestamp, JSON.stringify(orderedParams)].join(':')
    const sigAsByteArray = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, value, secret)
    return byteArrayToHexString(sigAsByteArray)
  } catch(e) { logError(e) }
}

function getResource(endPoint, params) {
  try {
    params.version = endPoint.version
    params.apiKey = endPoint.apiKey
    let hmac = {}
    hmac.nonce = Utilities.getUuid()
    hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
    hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)

    const options = {method: 'GET', contentType: 'application/json'}
    const response = UrlFetchApp.fetch(endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac), options)
    return response
  } catch(e) { logError(e) }
}

function postResource(endPoint, params, payload) {
  try {
    params.version = endPoint.version
    params.apiKey = endPoint.apiKey
    let hmac = {}
    hmac.nonce = Utilities.getUuid()
    hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
    hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)

    const options = {method: 'POST', contentType: 'application/json', payload: payload}
    const response = UrlFetchApp.fetch(endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac), options)
    return response
  } catch(e) { logError(e) }
}