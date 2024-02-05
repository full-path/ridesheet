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
    const fetchUrl = endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac)
    log(fetchUrl)
    const response = UrlFetchApp.fetch(fetchUrl, options)
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
    const fetchUrl = endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac)
    log(fetchUrl, JSON.stringify(payload))
    const response = UrlFetchApp.fetch(fetchUrl, options)
    return response
  } catch(e) { logError(e) }
}