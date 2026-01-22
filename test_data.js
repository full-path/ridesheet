function testCreateDummyData() {
  //const baseLocation = "2124 SE Oak St Portland OR 97214"
  const baseLocation = "615 Cottonwood Ln, Condon, OR 97823"
  const baseFormattedAddress = getGeocode(baseLocation,"formatted_address")
  const baseLocationObj = getGeocode(baseLocation,"object")

  const params = {
    baseFormattedAddress: baseFormattedAddress,
    startingLocation: {
      lat: baseLocationObj.lat,
      lng: baseLocationObj.lng
    },
    agencyDomain: "thedomain.org",
    areaCode: "(541)",
    numCustomers: 10,
    addressRadius: 30,
    startDate: new Date("2025-10-01T00:00:00-07:00"),
    tripDate: new Date("2025-10-06T00:00:00-07:00"),
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
    useCache: false
  }

  // moveTestRecordsToReview(params)
  // return

  Logger.log("Generating POI addresses...")
  const poiAddresses = generatePoiAddresses(params)
  Logger.log("POI addresses generated.")

  Logger.log("Generating home addresses...")
  const homeAddresses = generateRandomHomeAddresses(params)
  Logger.log("Home addresses generated.")

  // Merge all addresses to send to the routes generator
  const allAddresses = [
    {"Address": baseFormattedAddress, "Default Trip Purpose": "Depot"},
    ...poiAddresses.map(a => {
      return {"Address": a["Address"], "Default Trip Purpose": a["Default Trip Purpose"]}
    }), 
    ...homeAddresses.map(a => {
      return {"Address": a["formattedAddress"], "Default Trip Purpose": "Home"}
    })
  ]
  // Logger.log(JSON.stringify(allAddresses,null,2))
  // return

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
  // params.useCache = false
  const trips = generateTrips(params, runs, customers, poiAddresses)
  Logger.log("Customers generated.")

  Logger.log("Generating routes...")
  // params.useCache = true
  const routes = getRoutes(params, allAddresses)
  Logger.log("Routes generated.")

  Logger.log("Generating assignments...")
  const assignments = getTripAssignments(params, trips, runs, vehicles, routes)
  Logger.log("Assignment results received...")
}

function moveTestRecordsToReview(params) {
  const tripFilter = function(row) { 
    return row["Trip Date"] && row["Trip Date"].getTime() === params.tripDate.getTime()
  }
  const runFilter = function(row) {
    return row["Run Date"] && row["Run Date"].getTime() === params.tripDate.getTime()
  }
  moveTripsToReview(tripFilter, runFilter)
}

function getNavigableFormattedAddress(address) {
  try {
    const allowedLocationTypes = ["premise","route","street_address"]
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
  } catch(e) { logError(e) }
}

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

function generateRunTemplateRows(params,drivers,vehicles) {
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
  createRows(runTemplateSheet,newRunTemplateRows)
  newRunTemplates = getRangeValuesAsTable(runTemplateSheet.getDataRange())
  return newRunTemplateRows
}

function toRadians(deg) {
  return deg * Math.PI / 180;
}

function toDegrees(rad) {
  return rad * 180 / Math.PI;
}

/**
 * Compute bounding box around a center point in miles
 * @param {{lat:number, lng:number}} center
 * @param {number} radiusMiles
 * @return {{south:number, west:number, north:number, east:number}}
 */
function computeBoundingBox(center, radiusMiles) {
  const milesPerDegLat = 69;
  const milesPerDegLng = 69 * Math.cos(toRadians(center.lat));
  const latDelta = radiusMiles / milesPerDegLat;
  const lngDelta = radiusMiles / milesPerDegLng;

  return {
    south: center.lat - latDelta,
    north: center.lat + latDelta,
    west:  center.lng - lngDelta,
    east:  center.lng + lngDelta
  };
}

/**
 * Generate a random point within a circle of given radius (miles) around a center
 * @param {{lat:number, lng:number}} center
 * @param {number} radiusMiles
 * @param {function} [sampleRadiusFn]  Optional function returning a random distance in miles
 * @return {{lat:number, lng:number}}
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
 * Shuffle an array in place (Fisher–Yates)
 */
function shuffleArray(arr) {
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
}

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

  newAddresses.push(...getOsmAddresses(params.startingLocation, params.addressRadius, 
      '["amenity"~"clinic|hospital|doctor|dentist"]["name"]["addr:housenumber"]["addr:street"]["addr:postcode"]', "Medical", 5))
  Utilities.sleep(1000);
  newAddresses.push(...getOsmAddresses(params.startingLocation, params.addressRadius, 
      '["office"~"^(government|company|lawyer|insurance|accountant|charity|ngo|yes)$"]["name"]["addr:housenumber"]["addr:street"]["addr:postcode"]', "Work", 5))
  Utilities.sleep(1000);
  newAddresses.push(...getOsmAddresses(params.startingLocation, params.addressRadius, 
      '["amenity"~"place_of_worship|community_centre|library|bank|pharmacy|cinema"]["name"]["addr:housenumber"]["addr:street"]["addr:postcode"]', "Other", 5))
  
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
  createRows(addressSheet,newAddresses)

  return newAddresses
}

function getOsmAddressesOld(startLocation, radiusMiles = 50, osmQuery, purpose, limit) {
  const bboxObj = computeBoundingBox(startLocation, radiusMiles);
  const bbox = `${bboxObj.south},${bboxObj.west},${bboxObj.north},${bboxObj.east}`;

  const endpoint = 'https://overpass-api.de/api/interpreter';
  // const endpoint = 'https://overpass.private.coffee/api/interpreter'

  const overpassQL = `
    [out:json][timeout:60];
    (
      node${osmQuery}(${bbox});
      way${osmQuery}(${bbox});
      relation${osmQuery}(${bbox});
    );
    out tags ${limit};
  `;

  // Logger.log(overpassQL)

  const response = UrlFetchApp.fetch(endpoint, {
    method: 'post',
    payload: { data: overpassQL },
    muteHttpExceptions: true
  });
  //Utilities.sleep(1000);

  let osmData = {}
  try {
    osmData = JSON.parse(response.getContentText());
  } catch(e) {
    log(response.getContentText())
    Logger.log(e.name + ': ' + e.message, e.stack)
    Logger.log(response.getContentText())
    return []
  }

  const osmElements = osmData.elements || [];
  osmElements.forEach((elem) => {
    const tags = elem.tags
    const googleMapsQuery = `${tags["addr:housenumber"]} ${tags["addr:street"]} ${tags["addr:city"]} ${tags["addr:postcode"]}`
    const formattedAddress = getGeocode(googleMapsQuery,"formatted_address")
    if (formattedAddress.startsWith("Error")) Logger.log(JSON.stringify(tags,null,2))
    elem.formattedAddress = formattedAddress
  })

  const shortNames = []
  const newAddresses = osmElements.map((elem) => {
    return {
      "Short Name": createAddressShortName(elem.tags.name),
      "Address": `${elem.formattedAddress} (${elem.tags.name})`,
      "Default Trip Purpose": purpose
    }
  })

  Logger.log(`Received ${newAddresses.length} ${purpose} addresses`)
  return newAddresses
}

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
    ["1","2","3"].forEach((num) => {
      let thisRow = {
        "Vehicle ID": `${vehicleType}${num}`,
        "Vehicle Name": `${vehicleType} Number ${num}`,
        "Vehicle Start Date": params.startDate,
        "Garage Address": params.baseFormattedAddress
      }
      let thisCompleteRow = Object.assign(thisRow,vehicleTypes[vehicleType])
      newVehicles.push(thisCompleteRow)
    })
  })

  clearSheet(vehicleSheet)
  createRows(vehicleSheet,newVehicles)
  return newVehicles
}

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
  createRows(driverSheet,newDrivers)
  newDrivers = getRangeValuesAsTable(driverSheet.getDataRange())
  return newDrivers
}

/**
 * Generate fake customers for dummy data
 * @param {{lat:number, lng:number}} startLocation - Starting location for customer generation
 * @param {number} numCustomers - Number of customers to generate
 * @param {{Medical:Object[], Work:Object[], Other:Object[]}} poiPool - Pool of points of interest
 * @return {Object[]} Array of customer records with keys matching spreadsheet columns
 */
function generateCustomers(params, homeAddresses, poiAddresses) {
  let newCustomers = []
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const customerSheet = ss.getSheetByName("Customers")

  if (params.useCache) {
    newCustomers = getRangeValuesAsTable(customerSheet.getDataRange())
    return newCustomers
  }

  let firstNames = ["Binky","Zelda","Mango","Fuzz","Noodle","Pixel","Gizmo","Bubbles","Quasar","Nimbus",
                      "John","Mary","William","Elizabeth","George","Margaret","Henry","Dorothy","Charles","Mildred"];
  let lastNames  = ["McFluffle","Puddleton","Snickerdoodle","Fizzlebang","Wobble","Doodlebug","Sprinkles","Bubbleton","Twinkles","Jamboree",
                      "Ramirez","Johnson","Brown","Jones","Nguyen","Davis","Wilson","Moore","Taylor","Anderson"];
  const serviceIDs = ss.getRangeByName("lookupServiceIds").getValues().flat().filter(v => v)

  for (let i = 0; i < params.numCustomers; i++) {
    const customerID = i + 1;
    const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
    const lastName  = lastNames[Math.floor(Math.random() * lastNames.length)];

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
 * Generate runs for dummy data
 * @param {string} startDateStr - Start date in MM/DD/YYYY format
 * @param {string[]} driverIDs - Array of driver IDs
 * @param {string[]} vehicleIDs - Array of vehicle IDs
 * @return {Object[]} Array of run records with keys matching spreadsheet columns
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
 * Generate trips for dummy data
 * @param {Object[]} runs - Array of run records
 * @param {Object[]} customers - Array of customer records
 * @param {{Medical:Object[], Work:Object[], Other:Object[]}} poiAddresses - Pool of points of interest
 * @return {Object[]} Array of trip records with keys matching spreadsheet columns
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
      returnTrip["DO Address"] =  goTrip["PU Address"]
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

// helper random between min and max inclusive
function randBetween(min, max) {
  return Math.floor(min + Math.random() * (max - min + 1));
}

/**
 * Creates an acronym using a regular expression.
 * @param {string} phrase The input string of words.
 * @returns {string} The resulting acronym in uppercase.
 */
const createAddressShortName = (phrase) => {
  if (!phrase || typeof phrase !== 'string') {
    return '';
  }

  // remove words that shouldn't be part of an acronym
  const stopWords =['a', 'an', 'and', 'for', 'in', 'of', 'on', 'the']
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

function testGetTimeObject() {
  minutesPastMidnight = 719
  const testRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lookups").getRange("G2")
  testRange.setValue(getTimeString(minutesPastMidnight))
  Logger.log(getTimeString(minutesPastMidnight))
  log(getTimeString(minutesPastMidnight))
}

function getRandomInteger(min, max) {
  min = Math.ceil(min)
  max = Math.floor(max)
  return Math.floor(Math.random() * (max - min + 1)) + min
}

/**
 * Calls the Google Maps Routes API computeRouteMatrix method
 * using UrlFetchApp in Google Apps Script.
 */
function getRoutes(params, addresses) {
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

  const MAX_ELEMENTS = 625
  const API_URL = 'https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix';
  
  // Calculate the optimal batch size for origins.
  const batchSize = Math.floor(MAX_ELEMENTS / numAddresses)

  if (batchSize < 1) {
    throw new Error(`The number of addresses (${numAddresses}) is too large to process. Maximum is ${MAX_ELEMENTS}.`);
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
        'X-Goog-Api-Key': MAPS_API_KEY,
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
        routeToSave["Minutes"] = Math.ceil(parseInt(route.duration.slice(0,-1), 10) / 60)
        routeToSave["Default Trip Purpose"] = addresses[route.destinationIndex]["Default Trip Purpose"]
        newRoutes.push(routeToSave)
      }
    })
  }

  //Logger.log(JSON.stringify(newRoutes,null,2))
  
  clearSheet(routeSheet)
  createRows(routeSheet, newRoutes)

  return newRoutes
  
  log(JSON.stringify(result, null, 2));
  return result;
}

function getTripAssignments(params, trips, runs, vehicles, routes) {
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
      seats: parseInt(tripIn["Guests"] + 1,10),
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
    Logger.log(JSON.stringify(solution.status, null, 2))
    log(JSON.stringify(solution, null, 2))
  } catch (err) {
    Logger.log("Error: %s", err)
  }

}

function getMinutesPastMidnight(dateObject) {
  const hours = dateObject.getHours();
  const minutes = dateObject.getMinutes();
  return (hours * 60) + minutes;
}

function getTimeString(minutesPastMidnight) {
  const hours = Math.floor(minutesPastMidnight / 60)
  const minutes = Math.floor(minutesPastMidnight % 60)

  const formattedHours = String(hours > 12 ? hours - 12 : hours === 0 ? 12 : hours)
  const formattedMinutes = String(minutes).padStart(2, '0')
  const formattedPeriod = hours >= 12 ? "PM" : "AM"

  return `${formattedHours}:${formattedMinutes} ${formattedPeriod}`
}