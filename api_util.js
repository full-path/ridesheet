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
  try {
    return {"@time": combineDateAndTime(date, time)}
  } catch(e) { logError(e) }
}

// Takes a value and removes all characters that are not 0-9 and converts to integer
function buildPhoneNumberToSpec(value) {
  try {
    return parseInt([...value.toString()].map(char => isNaN(+char) ? '' : char.trim()).join(''))
  } catch(e) { logError(e) }
}

function buildPhoneNumberFromSpec(value) {
  try {
    if (!value) return ""
    const v = value.toString()

    if (v.length === 7) return `${v.slice(0,3)}-${v.slice(-4)}`
    if (v.length === 10) return `(${v.slice(0,3)})${v.slice(3,6)}-${v.slice(-4)}`
    if (v.length === 11 && v.slice(0,1) === '1') return `1(${v.slice(1,4)})${v.slice(4,7)}-${v.slice(-4)}`
    if (v.length > 11 && v.slice(0,1) === '1') return `1(${v.slice(1,4)})${v.slice(4,7)}-${v.slice(7,11)} x${v.slice(11)}`
    if (v.length > 10) return `(${v.slice(0,3)})${v.slice(3,6)}-${v.slice(6,10)} x${v.slice(10)}`
    return v
  } catch(e) { logError(e) }
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

function generateHmacHexStringOld(secret, nonce, timestamp, params) {
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

function generateHmacHexString(secret, senderId, receiverId, timestamp, nonce, method, body, url) {
  try {
    const value = senderId + receiverId + timestamp + nonce + method + body + url;
    const sigAsByteArray = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, value, secret)
    return byteArrayToHexString(sigAsByteArray);
  } catch (e) {
    logError(e);
  }
}

function getResourceOld(endPoint, params) {
  try {
    params.version = endPoint.version
    params.apiKey = endPoint.apiKey
    let hmac = {}
    hmac.nonce = Utilities.getUuid()
    hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
    hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)

    const options = {method: 'GET', contentType: 'application/json'}
    const fetchUrl = endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac)
    log(fetchUrl)
    const response = UrlFetchApp.fetch(fetchUrl, options)
    return response
  } catch(e) { logError(e) }
}

// Need to have a senderId -- right now the apiKey is both the sender/receiver ID
// TODO: format URL as just the /path not the full URL
function getResource(endPoint, params) {
  try {
    const version = endPoint.version
    const receiverId = endPoint.apiKey
    const senderId = endPoint.apiKey
    const nonce = Utilities.getUuid()
    const timestamp = new Date().getTime().toString()
    const endpointPath = extractPathFromUrl(endPoint.url)

    const signature = generateHmacHexString(
      endPoint.secret,
      senderId,
      receiverId,
      timestamp,
      nonce,
      'GET',
      '', // For GET requests, the body will be empty
      endpointPath
    );

    const authHeader = `${signature}:${senderId}:${receiverId}:${timestamp}:${nonce}`

    const options = {
      method: 'GET',
      contentType: 'application/json',
      headers: {
        Authorization: authHeader,
      }
    }

    params.authorization = authHeader
    params.endpointPath = endpointPath

    const fetchUrl = endPoint.url + "?" + urlQueryString(params)
    log(fetchUrl)

    const response = UrlFetchApp.fetch(fetchUrl, options)
    return response
  } catch (e) {
    logError(e)
  }
}

function postResourceOld(endPoint, payload) {
  try {
    params.version = endPoint.version
    params.apiKey = endPoint.apiKey
    let hmac = {}
    hmac.nonce = Utilities.getUuid()
    hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
    hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)

    const options = {method: 'POST', contentType: 'application/json', payload: payload}
    const fetchUrl = endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac)
    log(fetchUrl, JSON.stringify(payload))
    const response = UrlFetchApp.fetch(fetchUrl, options)
    return response
  } catch(e) { logError(e) }
}

// Need to have a senderId -- right now the apiKey is both the sender/receiver ID
// TODO: format URL as just the /path not the full URL
function postResource(endPoint, payload) {
  try {
    const version = endPoint.version
    const receiverId = endPoint.apiKey
    const senderId = endPoint.apiKey
    const nonce = Utilities.getUuid()
    const timestamp = new Date().getTime().toString()
    const endpointPath = extractPathFromUrl(endPoint.url)

    const signature = generateHmacHexString(
      endPoint.secret,
      senderId,
      receiverId,
      timestamp,
      nonce,
      'POST',
      JSON.stringify(payload),
      endpointPath
    );

    const authHeader = `${signature}:${senderId}:${receiverId}:${timestamp}:${nonce}`

    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: payload,
      headers: {
        Authorization: authHeader
      }
    };

    const params = {
      authorization: authHeader,
      endpointPath: endpointPath
    }

    const fetchUrl = endPoint.url + "?" + urlQueryString(params);
    log(fetchUrl);

    const response = UrlFetchApp.fetch(fetchUrl, options);
    return response;
  } catch (e) {
    logError(e);
  }
}

// Return empty string when working with other ridesheet instances
function extractPathFromUrl(url) {
  const regex = /\/v1\/?(.*)/; // Regex to match "/v1" followed by anything
  const match = url.match(regex);

  if (match && match.length > 1) {
    return "/v1" + match[1]; 
  } else {
    return '';
  }
}