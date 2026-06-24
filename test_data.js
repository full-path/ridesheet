/**
 * @fileoverview End-to-end dummy data generation for RideSheet demos and testing.
 *
 * The main entry point is `testCreateDummyData()`, which orchestrates generation of the
 * full data set: POI addresses, home addresses, customers, vehicles, drivers, run templates,
 * runs, trips, route matrices, and trip assignments.
 *
 * Most generator functions accept a shared `params` configuration object (see
 * `testCreateDummyData` for the full structure) and support a `params.useCache` flag.
 * When `useCache` is `true`, the function reads existing data from the corresponding
 * spreadsheet sheet rather than regenerating it — useful for partial re-runs during
 * development.
 *
 * POI address lookup is delegated to test_data_maps.js; the strategy is selected by
 * `params.addressSource` (`"nominatim"`, `"osm"`, `"google"`, or `"geocode"`).
 */

/**
 * Orchestrates generation of a complete RideSheet demo data set and writes each
 * entity type to its corresponding spreadsheet sheet. Logs progress at each stage.
 *
 * Configuration is defined inline via the `params` object, including base location,
 * number of customers, trip date, solver settings, and address source strategy.
 * After generating trips, assembles a route matrix for all referenced addresses and
 * submits trips to the assignment solver via `getTripAssignments`.
 */
function testCreateDummyData() {
  // Fail fast if the Apps Script project timezone doesn't match localTimeZone.
  // JS Date methods (.getDay(), .setDate(), etc.) operate in the script timezone,
  // so a mismatch silently produces wrong weekday calculations and date boundaries.
  const localTZ = getDocProp("localTimeZone")
  const scriptTZ = Session.getScriptTimeZone()
  if (scriptTZ !== localTZ) {
    const msg =
      `Timezone Mismatch — Cannot Continue\n` +
      `The Apps Script project timezone ('${scriptTZ}') does not match the localTimeZone ` +
      `document property ('${localTZ}'). ` +
      `Fix it via: Extensions → Apps Script → Project Settings → Time zone. ` +
      `Then re-run testCreateDummyData().`
    log(msg)
    Logger.log(msg)
    safeGetUi()?.alert(msg)
    return
  }

  const baseLocation = "108 SW Frazer Ave, Pendleton, OR 97801"
  const baseFormattedAddress = getGeocode(baseLocation, "formatted_address")
  const baseLocationObj = getGeocode(baseLocation, "object")

  // Fail fast if baseLocation is in a different timezone than localTimeZone.
  // Skipped silently when the Time Zone API is unavailable (getTimezoneForLatLng returns null).
  const baseLocationTZ = getTimezoneForLatLng(baseLocationObj.lat, baseLocationObj.lng)
  if (baseLocationTZ && baseLocationTZ !== localTZ) {
    const msg =
      `Timezone Mismatch — Cannot Continue\n` +
      `The baseLocation timezone ('${baseLocationTZ}') does not match the localTimeZone ` +
      `document property ('${localTZ}'). ` +
      `Either update the localTimeZone property to '${baseLocationTZ}', or change baseLocation ` +
      `to an address in the '${localTZ}' timezone.`
    log(msg)
    Logger.log(msg)
    safeGetUi()?.alert(msg)
    return
  }

  const params = {
    baseFormattedAddress: baseFormattedAddress,
    startingLocation: {
      lat: baseLocationObj.lat,
      lng: baseLocationObj.lng
    },
    agencyDomain: "sampledomain.org",
    areaCode: "(503)",
    numCustomers: 10,
    addressRadius: 10,
    startDate: Utilities.parseDate("2026-06-10 00:00:00", localTZ, "yyyy-MM-dd HH:mm:ss"),
    tripDate: Utilities.parseDate("2026-06-15 00:00:00", localTZ, "yyyy-MM-dd HH:mm:ss"),
    goWindowDuration: 30,
    returnWindowDuration: 30,
    solverTimeLimitSeconds: 10,
    maxSlackVehicleMinutes: 600,
    defaultPickupService: 2,
    defaultDropoffService: 2,
    defaultPenalty: 300,
    maxTimeFunction: (estTripHours) => {
      return Math.ceil(Math.max((estTripHours * 60) + 15, (estTripHours * 60 * 1.5) + 10))
    },
    useCache: false,
    addressSource: "nominatim"  // "nominatim", "osm", "google", or "geocode"
  }

  Logger.log("Generating POI addresses...")
  const poiAddresses = generatePoiAddresses(params)
  Logger.log("POI addresses generated.")

  Logger.log("Generating home addresses...")
  const homeAddresses = generateRandomHomeAddresses(params)
  Logger.log("Home addresses generated.")

  Logger.log("Generating customers...")
  const customers = generateCustomers(params, homeAddresses, poiAddresses)
  Logger.log("Customers generated.")

  Logger.log("Generating vehicles...")
  const vehicles = generateVehicles(params)
  Logger.log("Vehicles generated.")

  Logger.log("Generating drivers...")
  const drivers = generateDrivers(params)
  Logger.log("Drivers generated.")

  Logger.log("Generating run templates...")
  const runTemplateRows = generateRunTemplateRows(params, drivers, vehicles)
  Logger.log("Run templates generated.")

  Logger.log("Generating runs...")
  const runs = generateRuns(params)
  Logger.log("Runs generated.")

  Logger.log("Generating trips...")
  const trips = generateTrips(params, runs, customers, poiAddresses)
  Logger.log("trips generated.")

  // Build address → purpose lookup from known address sets
  const addressPurposeMap = new Map()
  addressPurposeMap.set(baseFormattedAddress, "Depot")
  poiAddresses.forEach(a => addressPurposeMap.set(a["Address"], a["Default Trip Purpose"]))
  homeAddresses.forEach(a => addressPurposeMap.set(a["formattedAddress"], "Home"))

  // Collect only addresses referenced by actual trips, plus depot
  const tripAddressSet = new Set([baseFormattedAddress])
  trips.forEach(trip => {
    if (trip["PU Address"]) tripAddressSet.add(trip["PU Address"])
    if (trip["DO Address"]) tripAddressSet.add(trip["DO Address"])
  })

  const allAddresses = Array.from(tripAddressSet).map(addr => ({
    "Address": addr,
    "Default Trip Purpose": addressPurposeMap.get(addr) || "Home"
  }))

  Logger.log("Generating routes...")
  //params.useCache = false
  // params.useCache = true
  const routes = getRoutes(params, allAddresses)
  Logger.log("Routes generated.")

  Logger.log("Generating assignments...")
  const assignments = getTripAssignments(params, trips, runs, vehicles, routes)
  Logger.log("Assignment results received...")
}

/**
 * Resolves the IANA timezone identifier for a lat/lng coordinate using the
 * Google Maps Time Zone API. Returns `null` if the API key is absent or the
 * request fails, so callers can skip the check gracefully.
 * @param {number} lat - Latitude.
 * @param {number} lng - Longitude.
 * @returns {string|null} IANA timezone string (e.g. `"America/Los_Angeles"`), or `null`.
 */
function getTimezoneForLatLng(lat, lng) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GOOGLE_MAPS_API_KEY")
  if (!apiKey) return null
  const timestamp = Math.floor(Date.now() / 1000)
  const url =
    `https://maps.googleapis.com/maps/api/timezone/json` +
    `?location=${lat},${lng}&timestamp=${timestamp}&key=${apiKey}`
  try {
    const result = JSON.parse(UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getContentText())
    if (result.status === "OK") return result.timeZoneId
  } catch (e) { logError(e) }
  return null
}

/**
 * Moves trips and runs for the date specified in `params.tripDate` into the review
 * workflow by delegating to `moveTripsToReview` with date-matching filters.
 * @param {Object} params - Generation parameters; uses `params.tripDate` (Date).
 */
function moveTestRecordsToReview(params) {
  const tripFilter = function(row) {
    return row["Trip Date"] && row["Trip Date"].getTime() === params.tripDate.getTime()
  }
  const runFilter = function(row) {
    return row["Run Date"] && row["Run Date"].getTime() === params.tripDate.getTime()
  }
  moveTripsToReview(tripFilter, runFilter)
}

/**
 * Geocodes an address and returns the formatted address only if it resolves to a
 * navigable street location. Results are filtered to `premise`, `route`, or
 * `street_address` types and must include a street number and have no restricted
 * travel modes on their navigation points.
 * @param {string} address - Address string or `"lat,lng"` coordinate pair to geocode.
 * @returns {string|undefined} The formatted address if a navigable result is found,
 *   or `undefined` if no suitable result exists.
 */
function getNavigableFormattedAddress(address) {
  try {
    const allowedLocationTypes = ["premise", "route", "street_address"]
    const mapsResults = Maps.newGeocoder().geocode(address)
    if (mapsResults.status !== "OK") return

    const goodResult = mapsResults.results.find(result => {
      //return result?.geometry?.location_type === "ROOFTOP" &&
      //["premise","route"].filter(item => result.types.includes(item)).length &&
      return result.types.filter(item => allowedLocationTypes.includes(item)).length &&
        result.address_components.some(component => component.types.includes("street_number")) &&
        !result?.navigation_points?.some(navPoint =>
          navPoint.restricted_travel_modes && navPoint.restricted_travel_modes.length > 0
        )
    })
    if (goodResult) {
      return goodResult.formatted_address
    } else {
      Logger.log(`Bad result: ${address}: ${mapsResults.results[0].formatted_address}`)
    }
  } catch (e) { logError(e) }
}

/**
 * Generates `params.numCustomers` random home addresses within `params.addressRadius`
 * miles of `params.startingLocation`. Each candidate point is reverse-geocoded through
 * `getNavigableFormattedAddress`; points that don't resolve to a navigable street are
 * discarded and retried. Distance sampling is weighted so that most addresses fall
 * within the inner half of the radius.
 *
 * When `params.useCache` is `true`, returns addresses already in the Customers sheet
 * column I instead of generating new ones.
 * @param {Object} params - Generation parameters; uses `startingLocation`, `addressRadius`,
 *   `numCustomers`, and `useCache`.
 * @returns {Object[]} Array of address objects, each with a `formattedAddress` string property.
 */
function generateRandomHomeAddresses(params) {
  let newAddresses = []

  if (params.useCache) {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const customerSheet = ss.getSheetByName("Customers")
    newAddresses = JSON.parse(JSON.stringify(
      customerSheet.getRange("I2:I").getValues().map(row => {
        return { formattedAddress: row[0] }
      })
    ))
    return newAddresses
  }

  function sampleDistance(addressRadius) {
    const p = Math.random();
    if (p < 0.4) return Math.random() * (addressRadius / 4);
    else if (p < 0.9) return (addressRadius / 4) + Math.random() * (addressRadius / 4);
    else return (addressRadius / 2) + Math.random() * (addressRadius / 2);
  }

  do {
    // Random location within radius, weighted
    // let address = randomPointInRadius(params.startingLocation, params.addressRadius, sampleDistance)
    const address = randomPointInRadius(params.startingLocation, params.addressRadius, sampleDistance)
    //Logger.log(`${address.lat},${address.lng}`)
    const formattedAddress = getNavigableFormattedAddress(`${address.lat},${address.lng}`)
    if (formattedAddress) {
      address.formattedAddress = formattedAddress
      newAddresses.push(address)
      Logger.log(newAddresses.length)
    }
    Utilities.sleep(1000)
  } while (newAddresses.length < params.numCustomers)
  return newAddresses
}

/**
 * Creates one run template row per driver, pairing each driver with the vehicle at
 * the same index, set to run Monday–Saturday 8 AM–6 PM. Clears the Run Template sheet,
 * writes the new rows, then returns the raw row objects (not the re-read sheet data).
 *
 * When `params.useCache` is `true`, returns the existing Run Template sheet data instead.
 * @param {Object} params - Generation parameters; uses `useCache`.
 * @param {Object[]} drivers - Driver records as returned by `generateDrivers`.
 * @param {Object[]} vehicles - Vehicle records as returned by `generateVehicles`.
 * @returns {Object[]} Array of run template row objects written to the sheet.
 */
function generateRunTemplateRows(params, drivers, vehicles) {
  let newRunTemplates = []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const runTemplateSheet = ss.getSheetByName("Run Template")

  if (params.useCache) {
    newRunTemplates = getRangeValuesAsTable(runTemplateSheet.getDataRange())
    return newRunTemplates
  }

  const newRunTemplateRows = []
  drivers.forEach((driver, i) => {
    newRunTemplateRows.push({
      "Driver ID": driver["Driver ID"],
      "Vehicle ID": vehicles[i]["Vehicle ID"],
      "Days of Week": "Monday, Tuesday, Wednesday, Thursday, Friday, Saturday",
      "Scheduled Start Time": "8:00 AM",
      "Scheduled End Time": "6:00 PM"
    })
  })

  clearSheet(runTemplateSheet)
  createRows(runTemplateSheet, newRunTemplateRows)
  newRunTemplates = getRangeValuesAsTable(runTemplateSheet.getDataRange())
  return newRunTemplateRows
}

/**
 * Converts degrees to radians.
 * @param {number} deg
 * @returns {number}
 */
function toRadians(deg) {
  return deg * Math.PI / 180;
}

/**
 * Converts radians to degrees.
 * @param {number} rad
 * @returns {number}
 */
function toDegrees(rad) {
  return rad * 180 / Math.PI;
}

/**
 * Computes a latitude/longitude bounding box around a center point.
 * @param {{lat: number, lng: number}} center - Center coordinate.
 * @param {number} radiusMiles - Box half-extent in miles.
 * @returns {{south: number, west: number, north: number, east: number}}
 */
function computeBoundingBox(center, radiusMiles) {
  const milesPerDegLat = 69;
  const milesPerDegLng = 69 * Math.cos(toRadians(center.lat));
  const latDelta = radiusMiles / milesPerDegLat;
  const lngDelta = radiusMiles / milesPerDegLng;

  return {
    south: center.lat - latDelta,
    north: center.lat + latDelta,
    west: center.lng - lngDelta,
    east: center.lng + lngDelta
  };
}

/**
 * Returns a random point within `radiusMiles` of `center`, using a uniform bearing.
 * @param {{lat: number, lng: number}} center - Center coordinate.
 * @param {number} radiusMiles - Maximum distance from center in miles.
 * @param {function} [sampleRadiusFn] - Optional function `(radiusMiles) => distanceMiles`
 *   for non-uniform distance sampling; defaults to uniform random.
 * @returns {{lat: number, lng: number}}
 */
function randomPointInRadius(center, radiusMiles, sampleRadiusFn) {
  // Determine distance: use provided sampler or uniform distribution
  let dist;
  if (typeof sampleRadiusFn === 'function') {
    dist = sampleRadiusFn(radiusMiles);
  } else {
    dist = Math.random() * radiusMiles;
  }

  const bearing = Math.random() * 2 * Math.PI;
  const milesPerDegLat = 69;
  const milesPerDegLng = 69 * Math.cos(toRadians(center.lat));

  const deltaLat = (dist * Math.cos(bearing)) / milesPerDegLat;
  const deltaLng = (dist * Math.sin(bearing)) / milesPerDegLng;

  return {
    lat: center.lat + deltaLat,
    lng: center.lng + deltaLng
  };
}

/**
 * Shuffles an array in place using the Fisher–Yates algorithm.
 * @param {Array} arr - The array to shuffle.
 */
function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

/**
 * Generates a set of POI (point-of-interest) addresses for Medical, Work, and Other
 * trip purposes, then writes them to the Addresses sheet. The address lookup strategy
 * is selected by `params.addressSource`:
 * - `"google"` — Google Places API v1
 * - `"geocode"` — Google Maps geocoder fishing
 * - `"nominatim"` — OpenStreetMap Nominatim (default)
 * - anything else — Overpass API
 *
 * Deduplicates short names after collection: if two addresses produce the same short
 * name, a numeric suffix is appended to the later one.
 *
 * When `params.useCache` is `true`, returns addresses already in the Addresses sheet
 * columns B–C instead of generating new ones.
 * @param {Object} params - Generation parameters; uses `addressSource`, `startingLocation`,
 *   `baseFormattedAddress`, `addressRadius`, and `useCache`.
 * @returns {Object[]} Address objects with `Short Name`, `Address`, and `Default Trip Purpose`.
 */
function generatePoiAddresses(params) {
  let newAddresses = []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const addressSheet = ss.getSheetByName("Addresses")

  if (params.useCache) {
    newAddresses = addressSheet.getRange("B2:C").getValues().map(row => {
      return {
        "Address": row[0],
        "Default Trip Purpose": row[1]
      }
    })
    return newAddresses
  }

  if (params.addressSource === "google") {
    newAddresses.push(...getGooglePlacesAddresses(params.baseFormattedAddress, params.addressRadius,
      ["hospital", "doctor", "medical_clinic", "dentist"], "Medical", 5))
    newAddresses.push(...getGooglePlacesAddresses(params.baseFormattedAddress, params.addressRadius,
      ["local_government_office", "lawyer", "insurance_agency", "accounting"], "Work", 5))
    newAddresses.push(...getGooglePlacesAddresses(params.baseFormattedAddress, params.addressRadius,
      ["church", "library", "bank", "pharmacy", "movie_theater"], "Other", 5))
  } else if (params.addressSource === "geocode") {
    newAddresses.push(...getGeocodeFishedAddresses(params.startingLocation, params.addressRadius,
      ["hospital", "medical clinic", "pharmacy", "dentist"], "Medical", 5))
    Utilities.sleep(1000)
    newAddresses.push(...getGeocodeFishedAddresses(params.startingLocation, params.addressRadius,
      ["city hall", "library", "post office", "senior center"], "Work", 5))
    Utilities.sleep(1000)
    newAddresses.push(...getGeocodeFishedAddresses(params.startingLocation, params.addressRadius,
      ["church", "grocery store", "bank", "movie theater"], "Other", 5))
  } else if (params.addressSource === "nominatim") {
    newAddresses.push(...getNominatimAddresses(params.startingLocation, params.addressRadius,
      [{ key: "amenity", value: "hospital" }, { key: "amenity", value: "clinic" },
      { key: "amenity", value: "pharmacy" }, { key: "amenity", value: "dentist" },
      { key: "amenity", value: "doctors" }],
      "Medical", 5))
    Utilities.sleep(1000)
    newAddresses.push(...getNominatimAddresses(params.startingLocation, params.addressRadius,
      [{ key: "amenity", value: "townhall" }, { key: "amenity", value: "library" },
      { key: "amenity", value: "post_office" }, { key: "amenity", value: "social_facility" }],
      "Work", 5))
    Utilities.sleep(1000)
    newAddresses.push(...getNominatimAddresses(params.startingLocation, params.addressRadius,
      [{ key: "amenity", value: "place_of_worship" }, { key: "amenity", value: "bank" },
      { key: "amenity", value: "community_centre" }, { key: "shop", value: "supermarket" },
      { key: "amenity", value: "cinema" }],
      "Other", 5))
  } else {
    newAddresses.push(...getOsmAddresses(params.startingLocation, params.addressRadius,
      '["amenity"~"clinic|hospital|doctor|dentist"]["name"]["addr:housenumber"]["addr:street"]["addr:postcode"]', "Medical", 5))
    Utilities.sleep(1000);
    newAddresses.push(...getOsmAddresses(params.startingLocation, params.addressRadius,
      '["office"~"^(government|company|lawyer|insurance|accountant|charity|ngo|yes)$"]["name"]["addr:housenumber"]["addr:street"]["addr:postcode"]', "Work", 5))
    Utilities.sleep(1000);
    newAddresses.push(...getOsmAddresses(params.startingLocation, params.addressRadius,
      '["amenity"~"place_of_worship|community_centre|library|bank|pharmacy|cinema"]["name"]["addr:housenumber"]["addr:street"]["addr:postcode"]', "Other", 5))
  }

  if ((new Set(newAddresses.map(a => a["Short Name"]))).size !== newAddresses.length) {
    const seen = {};
    newAddresses.forEach(address => {
      if (!seen[address["Short Name"]]) {
        seen[address["Short Name"]] = 1
      } else {
        address["Short Name"] = `${address["Short Name"]}${seen[address["Short Name"]]}`
        seen[address["Short Name"]]++
      }
    })
  }

  clearSheet(addressSheet)
  createRows(addressSheet, newAddresses)

  return newAddresses
}

/**
 * Generates a fixed fleet of six vehicles (three Buses and three Vans, numbered 1–3)
 * with predefined capacity and accessibility attributes, then writes them to the
 * Vehicles sheet.
 *
 * When `params.useCache` is `true`, returns the existing Vehicles sheet data instead.
 * @param {Object} params - Generation parameters; uses `startDate`, `baseFormattedAddress`,
 *   and `useCache`.
 * @returns {Object[]} Vehicle records with keys matching Vehicles sheet columns.
 */
function generateVehicles(params) {
  let newVehicles = []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const vehicleSheet = ss.getSheetByName("Vehicles")

  if (params.useCache) {
    newVehicles = getRangeValuesAsTable(vehicleSheet.getDataRange())
    return newVehicles
  }

  const vehicleTypes = {
    Bus: {
      "Seating Capacity": 14,
      "Wheelchair Capacity": 4,
      "Scooter Capacity": 2,
      "Has Ramp": "HAS RAMP",
    },
    Van: {
      "Seating Capacity": 3,
      "Wheelchair Capacity": 1,
      "Scooter Capacity": 1,
      "Has Lift": "HAS LIFT"
    },
    // Sedan: {
    //   "Seating Capacity": 3,
    //   "Wheelchair Capacity": 0,
    //   "Scooter Capacity": 0,
    // }
  }

  Object.keys(vehicleTypes).forEach((vehicleType) => {
    ["1", "2", "3"].forEach((num) => {
      let thisRow = {
        "Vehicle ID": `${vehicleType}${num}`,
        "Vehicle Name": `${vehicleType} Number ${num}`,
        "Vehicle Start Date": params.startDate,
        "Garage Address": params.baseFormattedAddress
      }
      let thisCompleteRow = Object.assign(thisRow, vehicleTypes[vehicleType])
      newVehicles.push(thisCompleteRow)
    })
  })

  clearSheet(vehicleSheet)
  createRows(vehicleSheet, newVehicles)
  return newVehicles
}

/**
 * Generates a fixed roster of six drivers with fictional names and email addresses
 * under `params.agencyDomain`, then writes them to the Drivers sheet.
 *
 * When `params.useCache` is `true`, returns the existing Drivers sheet data instead.
 * @param {Object} params - Generation parameters; uses `agencyDomain`, `startDate`,
 *   and `useCache`.
 * @returns {Object[]} Driver records with keys matching Drivers sheet columns.
 */
function generateDrivers(params) {
  let newDrivers = []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const driverSheet = ss.getSheetByName("Drivers")

  if (params.useCache) {
    newDrivers = getRangeValuesAsTable(driverSheet.getDataRange())
    return newDrivers
  }

  newDrivers = [
    {
      "Driver ID": "JB",
      "Driver Name": "Jocelyn Love",
      "Driver Email": `jlove@${params.agencyDomain}`,
      "Driver Start Date": params.startDate
    },
    {
      "Driver ID": "AW",
      "Driver Name": "Amos Williams",
      "Driver Email": `awilliams@${params.agencyDomain}`,
      "Driver Start Date": params.startDate
    },
    {
      "Driver ID": "JV",
      "Driver Name": "Juanita Villarreal",
      "Driver Email": `jvillarreal@${params.agencyDomain}`,
      "Driver Start Date": params.startDate
    },
    {
      "Driver ID": "WC",
      "Driver Name": "Walter Chen",
      "Driver Email": `wchen@${params.agencyDomain}`,
      "Driver Start Date": params.startDate
    },
    {
      "Driver ID": "BN",
      "Driver Name": "Benny Newtrout",
      "Driver Email": `bnewtrout@${params.agencyDomain}`,
      "Driver Start Date": params.startDate
    },
    {
      "Driver ID": "DS",
      "Driver Name": "Diana Silver",
      "Driver Email": `dsilver@${params.agencyDomain}`,
      "Driver Start Date": params.startDate
    },
  ]

  clearSheet(driverSheet)
  createRows(driverSheet, newDrivers)
  newDrivers = getRangeValuesAsTable(driverSheet.getDataRange())
  return newDrivers
}

/**
 * Generates `params.numCustomers` customers with randomised names, phone numbers, and
 * a home address from `homeAddresses`. Assigns a random default DO address from
 * `poiAddresses` and a random service ID from the `lookupServiceIds` named range.
 * Writes results to the Customers sheet.
 *
 * When `params.useCache` is `true`, returns the existing Customers sheet data instead.
 * @param {Object} params - Generation parameters; uses `numCustomers`, `areaCode`, and `useCache`.
 * @param {Object[]} homeAddresses - Home address objects as returned by `generateRandomHomeAddresses`.
 * @param {Object[]} poiAddresses - POI address objects as returned by `generatePoiAddresses`.
 * @returns {Object[]} Customer records with keys matching Customers sheet columns.
 */
function generateCustomers(params, homeAddresses, poiAddresses) {
  let newCustomers = []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const customerSheet = ss.getSheetByName("Customers")

  if (params.useCache) {
    newCustomers = getRangeValuesAsTable(customerSheet.getDataRange())
    return newCustomers
  }

  let firstNames = ["Binky", "Zelda", "Mango", "Fuzz", "Noodle", "Pixel", "Gizmo", "Bubbles", "Quasar", "Nimbus",
    "John", "Mary", "William", "Elizabeth", "George", "Margaret", "Henry", "Dorothy", "Charles", "Mildred"];
  let lastNames = ["McFluffle", "Puddleton", "Snickerdoodle", "Fizzlebang", "Wobble", "Doodlebug", "Sprinkles", "Bubbleton", "Twinkles", "Jamboree",
    "Ramirez", "Johnson", "Brown", "Jones", "Nguyen", "Davis", "Wilson", "Moore", "Taylor", "Anderson"];
  const serviceIDs = ss.getRangeByName("lookupServiceIds").getValues().flat().filter(v => v)

  for (let i = 0; i < params.numCustomers; i++) {
    const customerID = i + 1;
    const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
    const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];

    newCustomers.push({
      "Customer ID": customerID,
      "Customer First Name": firstName,
      "Customer Last Name": lastName,
      "Customer Name and ID": `${lastName}, ${firstName} (${customerID})`,
      "Home Address": homeAddresses[i].formattedAddress,
      "Phone Number": `${params.areaCode} ${Math.floor(Math.random() * 900) + 100}-${Math.floor(Math.random() * 9000) + 1000}`,
      "Default PU Address": homeAddresses[i].formattedAddress,
      "Default DO Address": poiAddresses[Math.floor(Math.random() * poiAddresses.length)]["Address"],
      "Default Service ID": serviceIDs[Math.floor(Math.random() * serviceIDs.length)]
    });
  }

  clearSheet(customerSheet)
  createRows(customerSheet, newCustomers)

  return newCustomers;
}

/**
 * Generates one week of run records from the Run Template sheet starting on the first
 * Monday on or after `params.startDate`. Clears the Runs sheet, calls
 * `buildRunsFromTemplate`, then re-reads the sheet to return the written records.
 *
 * When `params.useCache` is `true`, returns the existing Runs sheet data instead.
 * @param {Object} params - Generation parameters; uses `startDate` and `useCache`.
 * @returns {Object[]} Run records with keys matching Runs sheet columns.
 */
function generateRuns(params) {
  let newRuns = []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const runsSheet = ss.getSheetByName("Runs")

  if (params.useCache) {
    newRuns = getRangeValuesAsTable(runsSheet.getDataRange())
    return newRuns
  }

  // set the start date to the first Monday on or after the param.startDate
  const startDate = new Date(params.startDate.getTime())
  startDate.setDate(startDate.getDate() + ((1 + 7 - params.startDate.getDay()) % 7))

  clearSheet(runsSheet)
  newRuns = buildRunsFromTemplate(startDate)
  newRuns = getRangeValuesAsTable(runsSheet.getDataRange())
  return newRuns
}

/**
 * Generates a go-trip and a return-trip for every customer on `params.tripDate`.
 * Go-trip pickup times are sampled uniformly between 8:30 AM and 1:00 PM; return
 * trips are scheduled after a random 30–60 minute stay at the destination.
 * Writes all trips to the Trips sheet.
 *
 * When `params.useCache` is `true`, returns the existing Trips sheet data instead.
 * @param {Object} params - Generation parameters; uses `tripDate`, `goWindowDuration`,
 *   `returnWindowDuration`, and `useCache`.
 * @param {Object[]} runs - Run records as returned by `generateRuns`.
 * @param {Object[]} customers - Customer records as returned by `generateCustomers`.
 * @param {Object[]} poiAddresses - POI address objects as returned by `generatePoiAddresses`.
 * @returns {Object[]} Trip records with keys matching Trips sheet columns.
 */
function generateTrips(params, runs, customers, poiAddresses) {
  let newTrips = []
  const earliestStartTime = 510 // 8:30 AM
  const latestStartTime = 780   // 1:00 PM
  const minStayDuration = 30
  const maxStayDuration = 60
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripsSheet = ss.getSheetByName("Trips")

  if (params.useCache) {
    newTrips = getRangeValuesAsTable(tripsSheet.getDataRange())
    return newTrips
  }

  // helper: add minutes to Date
  // function addMinutes(date, mins) {
  //   return new Date(date.getTime() + mins * 60000);
  // }

  // helper: round up to the nearest 15 minutes
  // function roundToNearest15(date) {
  //   const mins = date.getMinutes();
  //   const roundedMins = Math.ceil(mins / 15) * 15;
  //   const roundedDate = new Date(date);
  //   roundedDate.setMinutes(roundedMins, 0, 0);
  //   return roundedDate;
  // }

  // helper: sample from status distribution
  // function sampleStatus() {
  //   const p = Math.random();
  //   if (p < 0.90) return 'Completed';
  //   if (p < 0.96) return 'Early Cancel';
  //   if (p < 0.98) return 'Late Cancel';
  //   return 'No Show';
  // }

  // track per-date customer pool
  // let prevRunDate = null;
  // let pool = [];
  //const runDates = new Set(runs.map(run => run["Run Date"]))
  runDates = [params.tripDate]

  runDates.forEach(runDate => {
    customers.forEach((customer) => {
      const goTrip = {}
      doAddress = poiAddresses[Math.floor(Math.random() * poiAddresses.length)]

      goTrip["Customer ID"] = customer["Customer ID"]
      goTrip["Customer Name and ID"] = customer["Customer Name and ID"]
      goTrip["Trip Date"] = runDate
      goTrip["PU Address"] = customer["Default PU Address"]
      goTrip["DO Address"] = doAddress["Address"]
      let tripEstimate = getTripEstimate(goTrip["PU Address"], goTrip["DO Address"], "milesAndHours")
      goTrip["Est Hours"] = tripEstimate.hours
      goTrip["Est Miles"] = tripEstimate.miles
      goTrip["Trip Purpose"] = doAddress["_purpose"]
      goTrip["Service ID"] = customer["Default Service ID"]
      goTrip["Trip ID"] = Utilities.getUuid()
      const puTimeInMinutes = randBetween(earliestStartTime, latestStartTime)
      const apptTimeInMinutes = puTimeInMinutes + (goTrip["Est Hours"] * 60) + 10
      const roundedApptTimeInMinutes = Math.ceil(apptTimeInMinutes / 15) * 15
      goTrip["Appt Time"] = getTimeString(roundedApptTimeInMinutes)
      goTrip["DO Time"] = getTimeString(roundedApptTimeInMinutes - 10)
      goTrip["PU Time"] = getTimeString(roundedApptTimeInMinutes - 10 - (goTrip["Est Hours"] * 60))

      const returnTrip = {}
      returnTrip["Customer ID"] = customer["Customer ID"]
      returnTrip["Customer Name and ID"] = customer["Customer Name and ID"]
      returnTrip["Trip Date"] = runDate
      returnTrip["PU Address"] = goTrip["DO Address"]
      returnTrip["DO Address"] = goTrip["PU Address"]
      tripEstimate = getTripEstimate(goTrip["DO Address"], goTrip["PU Address"], "milesAndHours")
      returnTrip["Est Hours"] = tripEstimate.hours
      returnTrip["Est Miles"] = tripEstimate.miles
      returnTrip["Trip Purpose"] = goTrip["Trip Purpose"]
      returnTrip["Service ID"] = goTrip["Service ID"]
      returnTrip["Trip ID"] = Utilities.getUuid()
      const stayTime = Math.ceil(randBetween(minStayDuration, maxStayDuration) / 15) * 15
      returnTrip["PU Time"] = getTimeString(roundedApptTimeInMinutes + stayTime)
      returnTrip["DO Time"] = getTimeString(roundedApptTimeInMinutes + stayTime + tripEstimate.hours * 60)

      newTrips.push(goTrip)
      newTrips.push(returnTrip)
      // Logger.log(`${newTrips.length} trips`)
    })
  })

  clearSheet(tripsSheet)
  createRows(tripsSheet, newTrips)

  newTrips = getRangeValuesAsTable(tripsSheet.getDataRange())
  return newTrips
}

/**
 * Returns a random integer between `min` and `max`, inclusive.
 * @param {number} min - Lower bound (inclusive).
 * @param {number} max - Upper bound (inclusive).
 * @returns {number}
 */
function randBetween(min, max) {
  return Math.floor(min + Math.random() * (max - min + 1));
}

/**
 * Converts a place name into a short identifier. Stop words (`a`, `an`, `and`,
 * `for`, `in`, `of`, `on`, `the`) are removed, then the initial letter of each
 * remaining word is joined and uppercased. If only one word remains after filtering,
 * that word is returned as-is instead of a single-letter acronym.
 * @param {string} phrase - The place name to abbreviate (e.g. `"CHI Saint Anthony Hospital"`).
 * @returns {string} Short identifier (e.g. `"CSAH"`), or `''` if `phrase` is empty or not a string.
 */
const createAddressShortName = (phrase) => {
  if (!phrase || typeof phrase !== 'string') {
    return '';
  }

  // remove words that shouldn't be part of an acronym
  const stopWords = ['a', 'an', 'and', 'for', 'in', 'of', 'on', 'the']
  const stopWordsRegex = new RegExp(`\\b(${stopWords.join('|')})\\b`, 'gi')
  const phraseWithKeyWords = phrase.replace(stopWordsRegex, '')

  // Find all characters that are at the beginning of a word,
  // join them, and convert to uppercase.
  const matches = phraseWithKeyWords.match(/\b\w/g) || [];
  if (matches.length === 1) {
    return phraseWithKeyWords.trim()
  } else {
    return matches.join('').toUpperCase();
  }
};

/**
 * Returns a random integer between `min` and `max`, inclusive, using `Math.floor`.
 * @param {number} min - Lower bound (inclusive).
 * @param {number} max - Upper bound (inclusive).
 * @returns {number}
 */
function getRandomInteger(min, max) {
  min = Math.ceil(min)
  max = Math.floor(max)
  return Math.floor(Math.random() * (max - min + 1)) + min
}

/**
 * Builds a full route matrix for all `addresses` using the Google Maps Routes API
 * (`computeRouteMatrix`). Addresses are batched to stay within the API's 50-waypoint
 * limit (origins + destinations ≤ 50 per request). Results are written to the Routes
 * sheet and returned as an array of route records.
 *
 * Requires `GOOGLE_MAPS_API_KEY` in Script Properties. When `params.useCache` is `true`,
 * returns the existing Routes sheet data instead.
 * @param {Object} params - Generation parameters; uses `useCache`.
 * @param {Object[]} addresses - Address objects with an `"Address"` string property.
 *   Each address is cleaned via `parseAddress` before sending to the API.
 * @returns {Object[]|undefined} Route records with `Start Address`, `End Address`, `Miles`,
 *   `Minutes`, and `Default Trip Purpose`, or `undefined` if `GOOGLE_MAPS_API_KEY` is missing.
 */
function getRoutes(params, addresses) {
  const GOOGLE_MAPS_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');

  if (!GOOGLE_MAPS_API_KEY) {
    Logger.log('ERROR: GOOGLE_MAPS_API_KEY not set in Script Properties');
    return
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const routeSheet = ss.getSheetByName("Routes")
  let newRoutes = []
  if (params.useCache) {
    newRoutes = getRangeValuesAsTable(routeSheet.getDataRange())
    return newRoutes
  }
  const numAddresses = addresses.length
  if (numAddresses === 0) return []

  // Incoming addresses may not be ready for geocoding. Add a "clean" address to each address obj
  addresses.forEach(a => a.cleanAddress = parseAddress(a["Address"]).geocodeAddress)

  const MAX_WAYPOINTS_PER_REQUEST = 50
  const API_URL = 'https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix';

  // API limit: origins + destinations <= 50. Destinations are always all addresses,
  // so origins per batch = 50 - numAddresses.
  const batchSize = MAX_WAYPOINTS_PER_REQUEST - numAddresses

  if (batchSize < 1) {
    throw new Error(`The number of addresses (${numAddresses}) exceeds the Routes API limit. Maximum is ${MAX_WAYPOINTS_PER_REQUEST - 1}.`);
  }

  const addressesAsWaypoints = addresses.map(a => {
    return {
      waypoint: {
        address: a.cleanAddress
      }
    }
  })
  // Logger.log(JSON.stringify(addressesAsWaypoints,null,2))
  // return

  // Loop through the addresses in batches of the calculated size.
  for (let i = 0; i < numAddresses; i += batchSize) {
    const originWaypointBatch = addressesAsWaypoints.slice(i, i + batchSize)
    console.log(`Processing batch of ${originWaypointBatch.length} origins, starting from index ${i}...`);

    // Define origins and destinations
    var payload = {
      "origins": originWaypointBatch,
      "destinations": addressesAsWaypoints,
      "travelMode": "DRIVE",
      "routingPreference": "TRAFFIC_UNAWARE"
    }

    // Prepare request options
    var options = {
      'method': 'post',
      'contentType': 'application/json',
      'headers': {
        'X-Goog-Api-Key': GOOGLE_MAPS_API_KEY,
        'X-Goog-FieldMask': 'originIndex,destinationIndex,duration,distanceMeters,condition,status'
      },
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    }

    // Make the request
    const response = UrlFetchApp.fetch(API_URL, options)

    // Parse and log the response
    const result = JSON.parse(response.getContentText())
    // log(JSON.stringify(payload,null,2))
    // log(JSON.stringify(result,null,2))
    result.forEach(route => {
      const globalOriginIndex = i + route.originIndex
      const isSelf = globalOriginIndex === route.destinationIndex

      if (!isSelf && route.condition === "ROUTE_EXISTS") {
        const routeToSave = {}
        routeToSave["Start Address"] = addresses[globalOriginIndex].cleanAddress
        routeToSave["End Address"] = addresses[route.destinationIndex].cleanAddress
        routeToSave["Miles"] = route.distanceMeters ? route.distanceMeters / 1609.34 : 0
        routeToSave["Minutes"] = Math.ceil(parseInt(route.duration.slice(0, -1), 10) / 60)
        routeToSave["Default Trip Purpose"] = addresses[route.destinationIndex]["Default Trip Purpose"]
        newRoutes.push(routeToSave)
      }
    })
  }
  //Logger.log(JSON.stringify(newRoutes,null,2))

  clearSheet(routeSheet)
  createRows(routeSheet, newRoutes)

  return newRoutes
}

/**
 * Submits trips and runs for `params.tripDate` to the external trip-assignment solver
 * (URL from `TRIP_ASSIGNMENT_URL` in Script Properties). Constructs a time matrix from
 * `routes`, packages vehicles and trips into the solver's request format, and — when a
 * solution is found — writes `Sched PU Time`, `Sched DO Time`, `Vehicle ID`, and
 * `Driver ID` back to the Trips sheet.
 *
 * Trip time windows are derived from `Appt Time` (go-trips) or `PU Time` (return trips).
 * @param {Object} params - Generation parameters; uses `tripDate`, `goWindowDuration`,
 *   `returnWindowDuration`, `defaultPickupService`, `defaultDropoffService`,
 *   `solverTimeLimitSeconds`, `maxSlackVehicleMinutes`, `defaultPenalty`,
 *   `maxTimeFunction`, and `baseFormattedAddress`.
 * @param {Object[]} trips - Trip records as returned by `generateTrips`.
 * @param {Object[]} runs - Run records as returned by `generateRuns`.
 * @param {Object[]} vehicles - Vehicle records as returned by `generateVehicles`.
 * @param {Object[]} routes - Route records as returned by `getRoutes`.
 */
function getTripAssignments(params, trips, runs, vehicles, routes) {
  const TRIP_ASSIGNMENT_URL = PropertiesService.getScriptProperties().getProperty('TRIP_ASSIGNMENT_URL');

  if (!TRIP_ASSIGNMENT_URL) {
    Logger.log('ERROR: TRIP_ASSIGNMENT_URL not set in Script Properties');
    return
  }

  const runsThisDay = runs.filter(run => {
    return run["Run Date"] && run["Run Date"].getTime() === params.tripDate.getTime()
  })
  const runsToSend = runsThisDay.map(run => {
    return {
      id: `${run["Vehicle ID"]}-${run["Driver ID"]}-${run["Run ID"]}`,
      time_window: [
        getMinutesPastMidnight(run["Scheduled Start Time"]),
        getMinutesPastMidnight(run["Scheduled End Time"])
      ],
      seat_capacity: vehicles.find(v => v["Vehicle ID"] === run["Vehicle ID"])["Seating Capacity"],
      wc_capacity: vehicles.find(v => v["Vehicle ID"] === run["Vehicle ID"])["Wheelchair Capacity"]
    }
  })

  const tripsThisDay = trips.filter(trip => {
    return trip["Trip Date"] && trip["Trip Date"].getTime() === params.tripDate.getTime()
  })

  // Get addresses. The depot address is always the first address
  const addresses = [params.baseFormattedAddress]
  tripsThisDay.forEach(trip => {
    if (!addresses.includes(trip["PU Address"])) addresses.push(trip["PU Address"])
    if (!addresses.includes(trip["DO Address"])) addresses.push(trip["DO Address"])
  })

  const tripsOut = tripsThisDay.map(tripIn => {
    const tripOut = {
      id: tripIn["Trip ID"],
      pickup_base: addresses.indexOf(tripIn["PU Address"]),
      dropoff_base: addresses.indexOf(tripIn["DO Address"]),
      seats: parseInt(tripIn["Guests"] + 1, 10),
      max_ride: params.maxTimeFunction(tripIn["Est Hours"])
    }

    // Assuming here that a trip with an appt time is a "go" trip
    // and everything else is (or can be treated like) a return trip
    if (tripIn["Appt Time"]) {
      const windowEnd = getMinutesPastMidnight(tripIn["DO Time"])
      const windowStart = windowEnd - params.goWindowDuration
      tripOut.dropoff_tw = [windowStart, windowEnd]
    } else {
      const windowStart = getMinutesPastMidnight(tripIn["PU Time"])
      const windowEnd = windowStart + params.returnWindowDuration
      tripOut.pickup_tw = [windowStart, windowEnd]
    }
    return tripOut
  })

  //Logger.log(JSON.stringify(tripsOut,null,2))

  const time_matrix = []
  addresses.forEach(startAddress => {
    const cleanStartAddress = parseAddress(startAddress).geocodeAddress
    const thisRow = []
    addresses.forEach(endAddress => {
      const cleanEndAddress = parseAddress(endAddress).geocodeAddress
      if (startAddress === endAddress) {
        thisRow.push(0)
      } else {
        const thisRoute = routes.find(route => {
          return route["Start Address"] === cleanStartAddress && route["End Address"] === cleanEndAddress
        })
        if (!thisRoute) {
          Logger.log(`${cleanStartAddress} to ${cleanEndAddress}`)
        }
        thisRow.push(thisRoute["Minutes"])
      }
    })
    time_matrix.push(thisRow)
  })

  // Example base matrix (unique locations: depot=0, home=1, store=2)
  const payload = {
    base_time_matrix: time_matrix,
    depot_base_index: 0,
    vehicles: runsToSend,
    requests: tripsOut,
    same_place_travel_minutes: 0,
    default_pickup_service: params.defaultPickupService,
    default_dropoff_service: params.defaultDropoffService,
    solver_time_limit_sec: params.solverTimeLimitSeconds,
    max_slack_minutes: params.maxSlackVehicleMinutes,
    default_penalty: params.defaultPenalty
  }

  log(JSON.stringify(payload, null, 2))
  // return

  try {
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }

    const response = UrlFetchApp.fetch(TRIP_ASSIGNMENT_URL, options)
    Logger.log("Status: %s", response.getResponseCode())
    //Logger.log("Response: %s", response.getContentText())

    const solution = JSON.parse(response.getContentText())
    if (solution.solution_found) {
      const ss = SpreadsheetApp.getActiveSpreadsheet()
      const tripsSheet = ss.getSheetByName("Trips")
      const tripsRange = tripsSheet.getDataRange()
      const tripsUpdate = trips.map(trip => {
        const assignedTripIds = Object.keys(solution.request_assignments)
        const result = {}
        if (trip["Trip ID"] && assignedTripIds.includes(trip["Trip ID"])) {
          const assignment = solution.request_assignments[trip["Trip ID"]]
          result["Sched PU Time"] = getTimeString(assignment.pickup.arrival_minute)
          result["Sched DO Time"] = getTimeString(assignment.dropoff.arrival_minute)
          const runParts = assignment.vehicle_id.split("-")
          result["Vehicle ID"] = runParts[0]
          result["Driver ID"] = runParts[1]
        } else {
          result["Sched PU Time"] = ""
          result["Sched DO Time"] = ""
          result["Vehicle ID"] = ""
          result["Driver ID"] = ""
        }
        return result
      })
      setValuesByHeaderNames(tripsUpdate, tripsRange)
    }
    //Logger.log(JSON.stringify(solution.status, null, 2))
    //log(JSON.stringify(solution, null, 2))
  } catch (err) {
    Logger.log("Error: %s", err)
  }

}

/**
 * Returns the number of minutes elapsed since midnight for a given Date object.
 * @param {Date} dateObject - A Date whose hours and minutes are read in local time.
 * @returns {number} Minutes past midnight (0–1439).
 */
function getMinutesPastMidnight(dateObject) {
  const hours = dateObject.getHours();
  const minutes = dateObject.getMinutes();
  return (hours * 60) + minutes;
}

/**
 * Converts a minutes-past-midnight value to a 12-hour time string (e.g. `"2:05 PM"`).
 * @param {number} minutesPastMidnight - Minutes elapsed since midnight (0–1439).
 * @returns {string} Formatted time string in `"H:MM AM/PM"` format.
 */
function getTimeString(minutesPastMidnight) {
  const hours = Math.floor(minutesPastMidnight / 60)
  const minutes = Math.floor(minutesPastMidnight % 60)

  const formattedHours = String(hours > 12 ? hours - 12 : hours === 0 ? 12 : hours)
  const formattedMinutes = String(minutes).padStart(2, '0')
  const formattedPeriod = hours >= 12 ? "PM" : "AM"

  return `${formattedHours}:${formattedMinutes} ${formattedPeriod}`
}
