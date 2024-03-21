// Breaks a raw address down into the required TDS attributes.
// If the source address includes a plus code, then the plus code will
// be the source of lat/long data.
// If only a plus code is provided, then the function will return
// as much information as can supplied from Google Maps based on that
// lat/long (usually just city, state and country)
function buildAddressToSpec(address) {
  if (!address) {
    return null
  }
  try {
    const rawAddressParts = parseAddress(address)
    let result = {
      addressName: "",
      street: "",
      street2: "",
      city: "",
      state: "",
      postalCode: "",
      country: "",
      notes: "",
      lat: "",
      long: "",
    }
    if (rawAddressParts.addressToFormat) {
      const addressObj = getGeocode(rawAddressParts.addressToFormat,"object")
      if (addressObj.status === "OK") {
        Object.keys(result).forEach(key => result[key] = addressObj[key] || "")
      }
      if (rawAddressParts.globalPlusCode) {
        const plusCodeObj = getGeocode(rawAddressParts.globalPlusCode,"object")
        if (plusCodeObj.status === "OK") {
          result.lat = plusCodeObj.lat
          result.long = plusCodeObj.long
        }
      }
    } else if (rawAddressParts.globalPlusCode) {
      const plusCodeObj = getGeocode(rawAddressParts.globalPlusCode,"object")
      if (plusCodeObj.status === "OK") {
        Object.keys(result).forEach(key => result[key] = plusCodeObj[key] || "")
      }
    }
    return result
  } catch(e) { logError(e) }
}

function buildAddressFromSpec(address) {
  if (!address) {
    return null
  }
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
  if (!phoneNumber) {
    return null
  }
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
  if (!e164Number) {
    return null
  }
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

function getCustomerId(referral) {
  const theDate = referral["Date of Birth"] || new Date()
  return `${referral["Customer First Name"][0]}-${referral["Customer Last Name"]}-${formatDate(theDate,null,"yyyy-MM-dd")}`
}