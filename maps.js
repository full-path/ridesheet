/**
* Get latitude, longitude, or a formatted street address for a given
* @param address Address as string Ex. "300 N LaSalles St, Chicago, IL"
* @param return_type Return type as string "lat", "long", or "formatted_address"
* @customfunction
*/
function getGeocode(address,returnType) {
  try {
    const bounds = getDocProps([
      {name: "geocoderBoundSwLatitude"},
      {name: "geocoderBoundSwLongitude"},
      {name: "geocoderBoundNeLatitude"},
      {name: "geocoderBoundNeLongitude"}
      ])
    let mapGeo = Maps.newGeocoder().setBounds(
      bounds["geocoderBoundSwLatitude"],
      bounds["geocoderBoundSwLongitude"],
      bounds["geocoderBoundNeLatitude"],
      bounds["geocoderBoundNeLongitude"]
    )
    let result = mapGeo.geocode(address)
    if (returnType === "raw") {
      return JSON.stringify(result).slice(0,50000)
    } else if (result["status"] != "OK") {
      if (returnType === "object") {
        return {status: result.status}
      } else {
        return "Error: " + result.status
      }
    } else if (isPartialMatch(result)) {
      if (returnType === "object") {
        return {status: "partial_match"}
      } else {
        return "Error: partial match: " + result["results"][0]["formatted_address"]
      }
    } else {
      const mainResult = result.results[0]
      switch(returnType){
        case "lat":               return mainResult.geometry.location.lat
        case "lng":               return mainResult.geometry.location.lng
        case "formatted_address": return mainResult.formatted_address
        case "global_plus_code":  return mainResult.plus_code.global_code
        case "object": {
          const components = mainResult.address_components
          let street_number
          let route
          let addressObj = {}
          addressObj.status = "OK"
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
          addressObj.lat = mainResult.geometry.location.lat
          addressObj.lng = mainResult.geometry.location.lng
          addressObj.long = mainResult.geometry.location.lng
          if (mainResult.hasOwnProperty("plus_code")) {
            addressObj.global_plus_code = mainResult.plus_code.global_code
            addressObj.compound_plus_code = mainResult.plus_code.compound_code
          }
          return addressObj
        }
        default: return "Error: Invalid Return Type"
      }
    }
  } catch(e) { logError(e) }
}

function isPartialMatch(geocodeResults) {
  if (geocodeResults["results"][0]["partial_match"]) {
    let locationType = geocodeResults["results"][0]["geometry"]["location_type"]
    let types = geocodeResults["results"][0]["types"]
    if (locationType === 'APPROXIMATE') {
      return true
    }
  }
  return false
}

function setAddressByShortName(app, range) {
  try {
    const ss = app.getActiveSpreadsheet()
    const sheet = ss.getSheetByName('Addresses')
    const dataRange = sheet.getDataRange()
    const data = getRangeValuesAsTable(dataRange)
    const searchTerm = range.getValue().toString().toLowerCase().trim()
    const result = data.find((row) => row["Short Name"].toString().toLowerCase().trim() === searchTerm)["Address"].trim()
    if (!result) { throw new Error('No address found by short name') }
    range.setValue(result)
    range.setNote("")
    range.setBackground(null)
    return true
  } catch(e) {
    return false
  }
}

function setAddressByApi(app, range) {
  try {
    const rawAddressParts = parseAddress(range.getValue())
    let globalPlusCode = ""
    let formattedAddress = ""
    if (rawAddressParts.compoundPlusCode) {
      globalPlusCode = getGeocode(rawAddressParts.compoundPlusCode, "global_plus_code")
    } else if (rawAddressParts.globalPlusCode) {
      globalPlusCode = rawAddressParts.globalPlusCode
    }
    if (rawAddressParts.addressToFormat) {
      formattedAddress = getGeocode(rawAddressParts.addressToFormat, "formatted_address")
    }

    let errorMsgs = []
    if (globalPlusCode.startsWith("Error")) errorMsgs.push("Plus Code " + globalPlusCode)
    if (formattedAddress.startsWith("Error")) errorMsgs.push("Address " + formattedAddress)

    if (errorMsgs.length) {
      const backgroundColor = app.newColor()
      const msg = errorMsgs.join("\n")
      range.setNote(msg)
      app.getActiveSpreadsheet().toast(msg)
      backgroundColor.setRgbColor(errorBackgroundColor)
      range.setBackgroundObject(backgroundColor.build())
      return false
    } else {
      let resultParts = []
      if (globalPlusCode) resultParts.push(globalPlusCode)
      if (formattedAddress) resultParts.push(formattedAddress)
      let result = resultParts.join("; ")
      if (rawAddressParts.parenText) result = `${result} (${rawAddressParts.parenText})`
      range.setValue(result)
      range.setNote("")
      range.setBackground(null)
      return true
    }
  } catch(e) {
    logError(e)
    return false
  }
}

function getTripEstimate(origin, destination, returnType) {
  const mapObj = Maps.newDirectionFinder()
  mapObj.setOrigin(origin)
  mapObj.setDestination(destination)
  const result = mapObj.getDirections()

  if (returnType === "raw") {
    return JSON.stringify(result).slice(0,50000)
  } else if (result["status"] != "OK") {
    return "Error: " + result["status"]
  } else {
    const distanceInMeters  = result["routes"][0]["legs"][0]["distance"]["value"]
    const durationInSeconds = result["routes"][0]["legs"][0]["duration"]["value"]
    switch(returnType){
      case "meters":
        return  distanceInMeters
      case "kilometers":
        return (distanceInMeters / 1000)
      case "miles":
        return (distanceInMeters * 0.000621371)
      case "seconds":
        return  durationInSeconds
      case "minutes":
        return (durationInSeconds / 60)
      case "hours":
        return (durationInSeconds / 360)
      case "days":
        return (durationInSeconds / 86400)
      case "milesAndDays":
        return {miles: (distanceInMeters * 0.000621371), days: (durationInSeconds / 86400)}
      case "milesAndHours":
        return {miles: (distanceInMeters * 0.000621371), hours: (durationInSeconds / 3600)}
      default:
        return "Error: Invalid Unit Type"
    }
  }
}

// parenText = Whatever user puts in parentheses. It's not further evaluated or sent to Maps
// geocodeAddress = The string that can be directly passed to Google Maps for driving directions
//     This can be a plus code or street address
// compoundPlusCode = example: CWC7+RW Mountain View, California
// globalPlusCode = example: 849VCWC7+RW
// addressToFormat = If plus code is found, this value is whatever text remains after plus code
//     and parenText is removed. This allows a rawAddress to contain both a manually
//     entered plus code and a street address that can be parsed by the Maps geocoding API
//     to get the elements needed for TDS compliance. This combination is useful for
//     rural scenarios where Maps return lat/longs that can be off by a significant distance.
function parseAddress(rawAddress) {
  let result = {}
  let remainingAddress = rawAddress.toString()
  const parenText = remainingAddress.match(/\(([^)]*)\)/)
  if (parenText) {
    result.parenText  = parenText[1]
    remainingAddress = remainingAddress.replace(parenText[0],"").trim()
  }
  const globalPlusCode = remainingAddress.match(/(^|\s)(([23456789C][23456789CFGHJMPQRV][23456789CFGHJMPQRVWX]{6}\+[23456789CFGHJMPQRVWX]{2,3})\s*;?)(\s|$)/)
  if (globalPlusCode) {
    result.geocodeAddress = globalPlusCode[3]
    result.globalPlusCode = globalPlusCode[3]
    remainingAddress = remainingAddress.replace(globalPlusCode[0],"").trim()
    if (remainingAddress) {
      result.addressToFormat = remainingAddress
    }
  } else {
    const compoundPlusCode = remainingAddress.match(/(^|\s)(([23456789CFGHJMPQRVWX]{4,6}\+[23456789CFGHJMPQRVWX]{2,3}.*);)(\s|$)/)
    if (compoundPlusCode) {
      remainingAddress = remainingAddress.replace(compoundPlusCode[0],"").trim()
      if (remainingAddress) {
        result.compoundPlusCode = compoundPlusCode[3].trim()
        result.addressToFormat = remainingAddress
      } else {
        result.geocodeAddress = compoundPlusCode[3].trim()
      }
    } else {
      result.geocodeAddress = remainingAddress
      result.addressToFormat = remainingAddress
    }
  }
  return result
}

function createGoogleMapsDirectionsURL(address) {
  const baseURL  = "https://www.google.com/maps/dir/?api=1"
  const travelMode  = "&travelmode=driving"
  const destAddress = parseAddress(address).geocodeAddress
  const destination = "&destination=" + encodeURIComponent(destAddress)
  return baseURL + travelMode + destination
}

function extractCity(address) {
  let noParens = parseAddress(address).geocodeAddress
  let parsed = noParens.match(/.*, (.*, .*) \d{5}, USA/)
  if (parsed) return parsed[1]
  parsed = noParens.match(/[A-Z0-9]{4}\+[A-Z0-9]{2,3},? (.*, .*), USA/)
  if (parsed) return parsed[1]
  let isPlusCode = noParens.match(/.*\+.*/)
  if (isPlusCode) {
    let cache = CacheService.getScriptCache()
    let cachedCity = cache.get(noParens)
    if (cachedCity) {
      return cachedCity
    } else {
      let geocodeResult = Maps.newGeocoder().geocode(noParens)
      if (geocodeResult.status === 'OK') {
        let location = geocodeResult.results[0].geometry.location
        let locationInformation = Maps.newGeocoder().reverseGeocode(location.lat, location.lng)
        if (locationInformation.status === 'OK') {
          let approxAddress = locationInformation.results[0].formatted_address
          let city = extractCity(approxAddress)
          cache.put(noParens, city, 259200) // cache for 72 hours
          return city
        }
      }
    }
  }
  return "Unspecified area"
}
