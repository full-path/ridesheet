// TODO: Add more error handling and resilience. Add ability for addressName and other details
// check components[i].types to make sure each is the correct component
// This code is just for testing purposes!!!
// Make errors do something a little more useful--return a partial address or similar? Try to parse the address manually?
function buildAddressToSpec(address) {
  try {
    const geocoder = Maps.newGeocoder()
    const result = geocoder.geocode(address)
    if (result["status"] === "OK") {
      const components = result.results[0].address_components
      const addressObj = {
        addressName: "",
        street2: "",
        notes: "",
        formattedAddress: address,
        lat: null,
        long: null,
      }
      if (components.length >= 8) {
        addressObj.street = components[0].short_name + " " + components[1].short_name
        addressObj.city = components[3].long_name
        addressObj.state = components[5].short_name
        addressObj.postalCode = components[7].short_name
        addressObj.country = components[6].short_name
      }
      if (components.length === 6 || components.length === 7 ) {
        addressObj.street = components[0].short_name + " " + components[1].short_name
        addressObj.city = components[2].long_name
        addressObj.state = components[4].short_name
        addressObj.postalCode = components[6].short_name
        addressObj.country = components[5].short_name
      }
      if (components.length === 5) {
        addressObj.street = components[0].long_name
        addressObj.city = components[1].long_name
        addressObj.state = components[3].short_name
        addressObj.country = components[4].short_name
        addressObj.postalCode = ""
      }
      return addressObj
    }
  } catch(e) { logError(e) }
}

function buildAddressFromSpec(address) {
  try {
    if (address.formattedAddress) {
      return address.formattedAddress
    } 
    return address.street + ", " + address.city + ", " + address.state + " " + address.postalCode + ", " + address.country
  } catch(e) { logError(e) }
}

// TODO: support other time fields than just "time"
function buildTimeFromSpec(date, time) {
  try {
    return {"@time": combineDateAndTime(date, time)}
  } catch(e) { logError(e) }
}

//TODO: This is not very robust since the country code can be two digits
function buildPhoneNumberToSpec(phoneNumber, countryCode = "+1") {
  // Remove all non-digit characters from the input
  const cleaned = phoneNumber.replace(/\D/g, '')
  if (cleaned.length === 10) {
    return `${countryCode}${cleaned}`
  } else if (cleaned.length === 11) {
    return `+${cleaned}`
  } else {
    logError("Unsupported phone format")
    return ""
  }
}

function buildPhoneNumberFromSpec(e164Number) {
  if (!e164Number.startsWith('+')) {
    logError('Invalid E.164 phone number format - missing plus sign');
    return e164Number
  }
  const cleaned = e164Number.replace(/\D/g, '').slice(-10);
  if (cleaned.length === 10) {
    // Reformat the number to '123-123-1234' format
    return `${cleaned.slice(0, 3)}-${cleaned.slice(3, 6)}-${cleaned.slice(6)}`;
  } else {
    logError('Invalid E.164 phone number format - incorrect number of digits');
    return e164Number
  }
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
    const result = keys.map((key) => {
      return key + "=" + encodeURIComponent(params[key])
    }).join("&")
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

function generateHmacHexString(secret, senderId, receiverId, timestamp, nonce, method, body, url) {
  try {
    const value = senderId + receiverId + timestamp + nonce + method + body + url;
    const sigAsByteArray = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, value, secret)
    return byteArrayToHexString(sigAsByteArray);
  } catch (e) {
    logError(e);
  }
}

// Need to have a senderId -- right now the apiKey is both the sender/receiver ID
// GET requests will not be used at all with the standard TDS any longer; this is Ridesheet only 
// and likely soon to be deprecated
function getResource(endPoint, params) {
  try {
    const version = endPoint.version
    const receiverId = endPoint.receiverId
    const senderId = getDocProp("apiSenderId")
    const nonce = Utilities.getUuid()
    const timestamp = new Date().getTime().toString()
    const endpointPath = params.endpointPath || ''

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

// Need to have a senderId -- right now the apiKey is both the sender/receiver ID
// TODO: format URL as just the /path not the full URL
function postResource(endPoint, params, payload) {
  try {
    const version = endPoint.version
    const receiverId = endPoint.receiverId
    const senderId = getDocProp("apiSenderId")
    const nonce = Utilities.getUuid()
    const timestamp = new Date().getTime().toString()
    const endpointPath = params.endpointPath || ''

    const signature = generateHmacHexString(
      endPoint.secret,
      senderId,
      receiverId,
      timestamp,
      nonce,
      'POST',
      payload,
      endpointPath
    )

    const authHeader = `${signature}:${senderId}:${receiverId}:${timestamp}:${nonce}`

    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: payload,
      headers: {
        Authorization: authHeader
      }
    }

    params.authorization = authHeader

    const fetchUrl = buildUrl(endPoint.url, endpointPath, params)
    log(fetchUrl)

    const response = UrlFetchApp.fetch(fetchUrl, options)
    return response
  } catch (e) {
    logError(e)
  }
}

function buildUrl(baseURL, path, params) {
  if (baseURL.includes("script.google.com")) {
    return baseURL + "?" + urlQueryString(params)
  } else {
    return baseURL + path
  }
}

// Function to validate HMAC signature
function validateHmacSignature(signature, senderId, receiverId, timestamp, nonce, method, body, urlEndpoint) {
  try {
    // Fetch the secret associated with the receiverId (apiKey)
    const apiAccounts = getDocProp("apiGiveAccess")
    const statedApiAccount = apiAccounts[senderId]

    let receivedTimestamp = new Date(timestamp)
    let timeNow = new Date()
    let timePassed = timeNow.getTime() - receivedTimestamp.getTime()

    if (timePassed > 300000) return false

    // Generate the expected HMAC signature based on the retrieved secret and other parameters
    const expectedSignature = generateHmacHexString(
      statedApiAccount.secret,
      senderId,
      receiverId,
      timestamp,
      nonce,
      method,
      body,
      urlEndpoint
    )

    // Compare the expected signature with the received signature
    return signature === expectedSignature
  } catch (e) {
    logError(e)
    return false
  }
}

function createErrorResponse(status, message) {
  const response = ContentService.createTextOutput()
  const referenceId = (Math.floor(Math.random() * 10000000)).toString()
  response.setMimeType(ContentService.MimeType.JSON)
  const errorContent = { status: status, message: message, referenceId}
  response.setContent(JSON.stringify(errorContent))
  return response
}

function valueToBoolean(value) {
  if (
    value === null ||
    value === undefined ||
    value === "" ||
    value.toString().toLowerCase().slice(0, 3) === "unk"
  ) {
    return null
  } else if (
    !value ||
    value.toString().toLowerCase() === "false" ||
    value.toString().toLowerCase() === "no"
  ) {
    return false
  } else {
    return true
  }
}