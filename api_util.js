// TODO: Add more error handling and resilience. Add ability for addressName and other details
// check components[i].types to make sure each is the correct component
// This code is just for testing purposes!!!
// Make errors do something a little more useful--return a partial address or similar? Try to parse the address manually?
function buildAddressToSpec(address) {
  try {
    const geocoder = Maps.newGeocoder()
    const result = geocoder.geocode(address)
    if (result["status"] === "OK") {
      const mainResult = result.results[0]
      const components = mainResult.address_components
      let addressObj = {
        addressName: "",
        fullResult: JSON.stringify(result),
        addressName: "",
        street: "",
        street2: "",
        city: "",
        state: "",
        zip_code: "",
        country: "",
        notes: "",
        lat: mainResult.geometry.location.lat,
        long: mainResult.geometry.location.lng,
        formattedAddress: mainResult.formatted_address,
      }
      let street_number
      let route
      components.forEach((component) => {
        if (component.types.includes('street_number')) street_number = component.short_name
        if (component.types.includes('route')) route = component.short_name
        if (component.types.includes('subpremise')) {
          if (isNaN(+component.short_name)) {
            addressObj.street2 = component.short_name
          } else {
            addressObj.street2 = `#${component.short_name}`
          }
        }
        if (component.types.includes('locality')) addressObj.city = component.short_name
        if (component.types.includes('administrative_area_level_1')) {
          addressObj.state = component.short_name
        }
        if (component.types.includes('postal_code')) addressObj.zip_code = component.short_name
        if (component.types.includes('country')) addressObj.country = component.short_name
      })
      if (!street_number && !route) {
        addressObj.street = "Refer to lat/long coordinates"
      } else {
        addressObj.street = `${street_number} ${route}`
      }
      return addressObj
    } else {
      addressObj = {
        fullResult: JSON.stringify(result)
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

function createErrorResponse(status) {
  const response = ContentService.createTextOutput()
  response.setMimeType(ContentService.MimeType.JSON)
  const errorContent = { status: status }
  response.setContent(JSON.stringify(errorContent))
  return response
}