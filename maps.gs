/**
* Get latitude, longitude, or a formatted street address for a given 
* @param address Address as string Ex. "300 N LaSalles St, Chicago, IL"
* @param return_type Return type as string "lat", "long", or "formatted_address"
* @customfunction
*/
function getGeocode(address,returnType) {
  const bounds = getDocProps([
    {name: "geocoderBoundSwLatitude",  altValue: defaultGeocoderBoundSwLatitude},
    {name: "geocoderBoundSwLongitude", altValue: defaultGeocoderBoundSwLongitude},
    {name: "geocoderBoundNeLatitude",  altValue: defaultGeocoderBoundNeLatitude},
    {name: "geocoderBoundNeLongitude", altValue: defaultGeocoderBoundNeLongitude}
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
    return "Error: " + result["status"]
  } else if (result["results"][0]["partial_match"]) {
    return "Error: partial match: " + result["results"][0]["formatted_address"]
  }
  switch(returnType){
    case "lat":               return result["results"][0]["geometry"]["location"]["lat"]
    case "lng":               return result["results"][0]["geometry"]["location"]["lng"]
    case "formatted_address": return result["results"][0]["formatted_address"]
    default:                  return "Error: Invalid Return Type"
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
      default:
        return "Error: Invalid Unit Type"
    }
  }
}

function parseAddress(rawAddress) {
  result = {}
  parenText = rawAddress.match(/\([^)]*\)/)
  if (parenText) {
    result.parenText      = parenText[0]
    result.geocodeAddress = rawAddress.replace(parenText[0],"")
  } else {
    result.geocodeAddress = rawAddress
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