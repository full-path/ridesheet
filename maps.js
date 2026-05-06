/**
 * @fileoverview Google Maps integration for RideSheet.
 *
 * Provides geocoding, address parsing, trip estimation, and driving-directions
 * URL generation via the Google Maps Service (`Maps`) available in Apps Script.
 *
 * Key functions:
 * - `getGeocode()`        — geocode an address to lat/lng, formatted address, plus code, or a
 *                            structured address object. Also usable as a sheet custom function.
 * - `parseAddress()`      — splits a raw address string into its constituent parts (paren notes,
 *                            global plus code, compound plus code, street address).
 * - `getTripEstimate()`   — fetch driving distance and/or duration between two points.
 * - `setAddressByApi()`   — geocode a cell's raw address value and write the canonical result
 *                            back, highlighting errors in red with a cell note.
 * - `setAddressByShortName()` — look up a short name in the Addresses sheet and replace the
 *                            cell value with the full address.
 * - `extractCity()`       — extract the city/state from an address string.
 */

/**
 * Geocodes an address string using the Google Maps Geocoding API.
 *
 * The geocoder bounding box is set from the `geocoderBound*` document properties
 * to bias results toward the deployment's service area.
 *
 * @param {string} address - The address to geocode.
 * @param {string} returnType - Controls what is returned:
 *   - `"lat"`               — latitude as a number
 *   - `"lng"`               — longitude as a number
 *   - `"formatted_address"` — canonical formatted address string
 *   - `"global_plus_code"`  — Open Location Code (e.g. `"849VCWC7+RW"`)
 *   - `"raw"`               — first 50 000 characters of the raw JSON response
 *   - `"object"`            — structured object with fields:
 *       `{status, street, street2?, city, state, postalCode, country, lat, lng, long,
 *        global_plus_code?, compound_plus_code?}`.
 *       `street` is `"Refer to lat/long coordinates"` when no street address is returned.
 *   - Any other value       — `"Error: Invalid Return Type"`
 *
 * On a geocoding error the function returns `"Error: <status>"` for string return types,
 * or `{status: "<status>"}` for `"object"`. Partial matches return
 * `"Error: partial match: <formatted_address>"` or `{status: "partial_match"}`.
 *
 * @returns {string|number|Object} The requested geocoding result.
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
            if (component.types.includes('postal_code')) addressObj.postalCode = component.short_name
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

/**
 * Determines whether a geocoding response should be treated as a partial match
 * that is too imprecise to use. A result is considered a partial match if the
 * first result's `partial_match` flag is set AND the geometry `location_type`
 * is `"APPROXIMATE"`.
 * @param {Object} geocodeResults - The raw response object from `Maps.newGeocoder().geocode()`.
 * @returns {boolean} `true` if the result is an approximate partial match.
 */
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

/**
 * Looks up a full address by its short name in the "Addresses" sheet.
 * The lookup is case-insensitive and trims leading/trailing whitespace.
 * @param {string} shortName - The short name to look up in the "Short Name" column.
 * @returns {string|undefined} The trimmed address value from the "Address" column,
 *   or `undefined` if the sheet does not exist or no matching row is found.
 */
function getAddressByShortName(shortName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = ss.getSheetByName('Addresses')
    if (sheet) {
      const dataRange = sheet.getDataRange()
      const data = getRangeValuesAsTable(dataRange)
      const searchTerm = shortName?.toString().toLowerCase().trim()
      if (searchTerm) {
        const foundRow = data.find((row) => row["Short Name"].toString().toLowerCase().trim() === searchTerm)
        if (foundRow) {
          const result = foundRow["Address"].trim()
          return result
        }
      }
    }
  } catch(e) {
    logError(e)
  }
}

/**
 * Replaces a cell's value with the full address looked up by its short name
 * in the "Addresses" sheet. On success, clears the cell's note and background.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell whose value is the short name to look up.
 * @returns {boolean} `true` if the address was found and the cell updated; `false` otherwise.
 */
function setAddressByShortName(range) {
  try {
    const shortName = range.getValue()
    const result = getAddressByShortName(shortName)
    if (result) {
      range.setValue(result)
      range.setNote("")
      range.setBackground(null)
      return true
    } else {
      return false
    }
  } catch(e) {
    return false
  }
}

/**
 * Geocodes a cell's raw address value via the Maps API and writes the canonical
 * result back to the cell. The address is first parsed by `parseAddress()` to
 * extract any plus code and street address components. The function then:
 * 1. Resolves a global plus code from a compound plus code or passes through an
 *    existing global plus code.
 * 2. Geocodes the street address component to a formatted address string.
 * 3. Assembles the result as `"<plus_code>; <formatted_address> (<parenText>)"`
 *    (omitting absent components), then writes it back to the cell.
 *
 * On any geocoding error, the cell note is set to the error message(s), the cell
 * background is set to `errorBackgroundColor`, and a toast is shown.
 * On success, the cell note and background are cleared.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The cell containing the raw address.
 * @returns {boolean} `true` if geocoding succeeded and the cell was updated; `false` on error.
 */
function setAddressByApi(range) {
  try {
    const app = SpreadsheetApp
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

/**
 * Returns the driving distance and/or duration between two locations using
 * the Google Maps Directions API.
 *
 * @param {string} origin - The starting address or coordinates.
 * @param {string} destination - The ending address or coordinates.
 * @param {string} returnType - Controls what is returned:
 *   - `"meters"`      — distance in metres (number)
 *   - `"kilometers"`  — distance in kilometres (number)
 *   - `"miles"`       — distance in miles (number)
 *   - `"seconds"`     — duration in seconds (number)
 *   - `"minutes"`     — duration in minutes (number)
 *   - `"hours"`       — duration in hours (number)
 *   - `"days"`        — duration in days (number)
 *   - `"milesAndDays"` — `{miles, days}` object
 *   - `"milesAndHours"` — `{miles, hours}` object
 *   - `"raw"`         — first 50 000 characters of the raw JSON response
 *   - Any other value — `"Error: Invalid Unit Type"`
 * @returns {number|string|{miles: number, days: number}|{miles: number, hours: number}}
 */
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
        return (durationInSeconds / 3600)
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

/**
 * Parses a raw address string into its constituent parts.
 *
 * Handles four address formats that RideSheet supports:
 * - **Parenthesised notes**: Text inside `(...)` is extracted as `parenText` and
 *   stripped from further processing. It is never sent to the geocoding API.
 * - **Global plus code** (e.g. `849VCWC7+RW`): Extracted as `globalPlusCode` and
 *   `geocodeAddress`. Any remaining text becomes `addressToFormat`.
 * - **Compound plus code** (e.g. `CWC7+RW Mountain View, California`): Extracted
 *   as `compoundPlusCode`. The compound code is resolved to a global plus code
 *   by the caller. Any remaining text becomes `addressToFormat`. If there is no
 *   remaining text, the compound code is used as `geocodeAddress` directly.
 * - **Plain street address**: Set as both `geocodeAddress` and `addressToFormat`.
 *
 * Result object fields (all optional except where noted):
 * - `geocodeAddress`   {string} — The string to pass to the Maps API for driving directions.
 * - `addressToFormat`  {string} — Street address to pass to the geocoding API for formatting.
 * - `parenText`        {string} — Content extracted from parentheses.
 * - `globalPlusCode`   {string} — Full 11-character Open Location Code (e.g. `849VCWC7+RW`).
 * - `compoundPlusCode` {string} — Short plus code with locality (e.g. `CWC7+RW Mountain View, CA`).
 *
 * @param {string} rawAddress - The raw address string to parse.
 * @returns {{geocodeAddress?: string, addressToFormat?: string, parenText?: string,
 *            globalPlusCode?: string, compoundPlusCode?: string}}
 */
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

/**
 * Builds a Google Maps driving directions URL for the given address.
 * The address is parsed by `parseAddress()` and the `geocodeAddress` field
 * is used as the destination, encoded for use in a URL query parameter.
 * @param {string} address - The destination address string.
 * @returns {string} A Google Maps directions URL (driving mode) to the destination.
 */
function createGoogleMapsDirectionsURL(address) {
  const baseURL  = "https://www.google.com/maps/dir/?api=1"
  const travelMode  = "&travelmode=driving"
  const destAddress = parseAddress(address).geocodeAddress
  const destination = "&destination=" + encodeURIComponent(destAddress)
  return baseURL + travelMode + destination
}

/**
 * Extracts the city and state from a formatted address string.
 *
 * Tries two regex patterns against the `geocodeAddress` portion of the parsed input:
 * 1. Standard US address format: `", City, State NNNNN, USA"`
 * 2. Plus-code address: `"XXXX+XX, City, State, USA"`
 *
 * If the address contains a plus code that doesn't match either pattern, the
 * plus code is geocoded (via reverse geocoding) to obtain a formatted address
 * and the function is called recursively on the result. Results are cached in
 * Apps Script's script cache for 72 hours to limit Maps API calls.
 *
 * @param {string} address - The address string to extract a city from.
 * @returns {string} The city and state string (e.g. `"Chicago, IL"`), or
 *   `"Unspecified area"` if the city cannot be determined.
 */
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