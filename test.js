function testCreateDummyData() {
  const params = {
    numCustomers: 20,
    serviceRadius: 50,
    startDate: "4/19/2025",
    startingLocation: {
      lat: 39.9749,
      lon: -98.1897
    }
  }
  createDummyData(params)
}

function createDummyData(params) {
  // Use existing from lookups: drivers, vehicles, services, trip purposes, trip results
  // create: customers, runs and trips
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const serviceIDs    = ss.getRangeByName("lookupServiceIds").getValues().flat().filter(v => v);
  const vehicleIDs    = ss.getRangeByName("lookupVehicleIds").getValues().flat().filter(v => v);
  const driverIDs     = ss.getRangeByName("lookupDriverIds").getValues().flat().filter(v => v);
  const tripPurposes  = ["Medical", "Work", "Other"];
  const tripResults   = ["Completed", "No Show", "Early Cancel", "Late Cancel"];

  const poiPool = generatePOIPool(params.startingLocation, params.serviceRadius)
  const customers = generateCustomers(params.startingLocation, params.numCustomers, poiPool)
  const runs = generateRuns(params.startDate, driverIDs, vehicleIDs)
  const trips = generateTrips(runs, customers, poiPool)

  // geocode addresses
  // Create a cache for geocoding results
  const geocodeCache = {};

  // Helper function to get geocoded address with caching
  function getCachedGeocode(address) {
    if (!address) return null;
    if (geocodeCache[address]) {
      return geocodeCache[address];
    }
    const result = getGeocode(address, "formatted_address");
    geocodeCache[address] = result;
    Utilities.sleep(500);
    return result;
  }

  // Geocode customer addresses
  customers.forEach(customer => {
    if (customer["Default PU Address"]) {
      customer["Default PU Address"] = getCachedGeocode(customer["Default PU Address"]);
    }
    if (customer["Default DO Address"]) {
      customer["Default DO Address"] = getCachedGeocode(customer["Default DO Address"]);
    }
  });

  // Geocode trip addresses
  trips.forEach(trip => {
    if (trip["PU Address"]) {
      trip["PU Address"] = getCachedGeocode(trip["PU Address"]);
    }
    if (trip["DO Address"]) {
      trip["DO Address"] = getCachedGeocode(trip["DO Address"]);
    }
  });

  // Get sheets and their headers
  const customerSheet = ss.getSheetByName("Customers");
  const tripSheet = ss.getSheetByName("Trips");
  const runSheet = ss.getSheetByName("Runs");

  const customerHeaders = getSheetHeaderNames(customerSheet);
  const tripHeaders = getSheetHeaderNames(tripSheet);
  const runHeaders = getSheetHeaderNames(runSheet);

  // Format data to match sheet headers
  const formattedCustomers = customers.map(cust => {
    const row = {};
    customerHeaders.forEach(header => {
      row[header] = cust[header] || null;
    });
    return row;
  });

  const formattedRuns = runs.map(run => {
    const row = {};
    runHeaders.forEach(header => {
      row[header] = run[header] || null;
    });
    return row;
  });

  const formattedTrips = trips.map(trip => {
    const row = {};
    tripHeaders.forEach(header => {
      row[header] = trip[header] || null;
    });
    return row;
  });

  // Write data to sheets
  createRows(customerSheet, formattedCustomers);
  createRows(runSheet, formattedRuns);
  createRows(tripSheet, formattedTrips);
}

function toRadians(deg) {
  return deg * Math.PI / 180;
}

function toDegrees(rad) {
  return rad * 180 / Math.PI;
}

/**
 * Compute bounding box around a center point in miles
 * @param {{lat:number, lon:number}} center
 * @param {number} radiusMiles
 * @return {{south:number, west:number, north:number, east:number}}
 */
function computeBoundingBox(center, radiusMiles) {
  const milesPerDegLat = 69;
  const milesPerDegLon = 69 * Math.cos(toRadians(center.lat));
  const latDelta = radiusMiles / milesPerDegLat;
  const lonDelta = radiusMiles / milesPerDegLon;

  return {
    south: center.lat - latDelta,
    north: center.lat + latDelta,
    west:  center.lon - lonDelta,
    east:  center.lon + lonDelta
  };
}

/**
 * Generate a random point within a circle of given radius (miles) around a center
 * @param {{lat:number, lon:number}} center
 * @param {number} radiusMiles
 * @param {function} [sampleRadiusFn]  Optional function returning a random distance in miles
 * @return {{lat:number, lon:number}}
 */
function randomPointInRadius(center, radiusMiles, sampleRadiusFn) {
  // Determine distance: use provided sampler or uniform distribution
  let dist;
  if (typeof sampleRadiusFn === 'function') {
    dist = sampleRadiusFn();
  } else {
    dist = Math.random() * radiusMiles;
  }

  const bearing = Math.random() * 2 * Math.PI;
  const milesPerDegLat = 69;
  const milesPerDegLon = 69 * Math.cos(toRadians(center.lat));

  const deltaLat = (dist * Math.cos(bearing)) / milesPerDegLat;
  const deltaLon = (dist * Math.sin(bearing)) / milesPerDegLon;

  return {
    lat: center.lat + deltaLat,
    lon: center.lon + deltaLon
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

/**
 * Generate a POI pool using OSM Overpass API
 * @param {{lat:number, lon:number}} startLocation
 * @param {number} radiusMiles
 * @return {{Medical:Object[], Work:Object[], Other:Object[]}}
 */
function generatePOIPool(startLocation, radiusMiles = 50) {
  const bboxObj = computeBoundingBox(startLocation, radiusMiles);
  const bbox = `${bboxObj.south},${bboxObj.west},${bboxObj.north},${bboxObj.east}`;

  const categories = {
    Medical: { query: 'node["amenity"~"clinic|hospital|doctors|dentist"]', limit: 10 },
    Work:    { query: 'node["office"]',                 limit: 20 },
    Other:   { query: 'node["amenity"~"community_centre|library|bank|pharmacy|cinema|post_office"]', limit: 20 }
  };

  const endpoint = 'https://overpass-api.de/api/interpreter';
  const results = { Medical: [], Work: [], Other: [] };

  for (let cat in categories) {
    const { query, limit } = categories[cat];
    const overpassQL = `
      [out:json][timeout:25];
      (
        ${query}(${bbox});
      );
      out center;
    `;

    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      payload: { data: overpassQL }
    });
    Utilities.sleep(1000);

    const data = JSON.parse(response.getContentText());
    const elems = data.elements || [];
    const pois = elems.map(e => ({
      name: e.tags.name || '',
      lat: e.lat,
      lon: e.lon,
      type: cat
    }));

    shuffleArray(pois);
    results[cat] = pois.slice(0, limit);
  }

  return results;
}

/**
 * Generate fake customers for dummy data
 * @param {{lat:number, lon:number}} startLocation - Starting location for customer generation
 * @param {number} numCustomers - Number of customers to generate
 * @param {{Medical:Object[], Work:Object[], Other:Object[]}} poiPool - Pool of points of interest
 * @return {Object[]} Array of customer records with keys matching spreadsheet columns
 */
function generateCustomers(startLocation, numCustomers, poiPool) {
  const customers = [];

  const firstNames = ["Binky","Zelda","Mango","Fuzz","Noodle","Pixel","Gizmo","Bubbles","Quasar","Nimbus",
                      "John","Mary","William","Elizabeth","George","Margaret","Henry","Dorothy","Charles","Mildred"];
  const lastNames  = ["McFluffle","Puddleton","Snickerdoodle","Fizzlebang","Wobble","Doodlebug","Sprinkles","Bubbleton","Twinkles","Jamboree",
                      "Smith","Johnson","Brown","Jones","Miller","Davis","Wilson","Moore","Taylor","Anderson"];

  function sampleDistance() {
    const p = Math.random();
    if (p < 0.4) return Math.random() * 5;
    else if (p < 0.9) return 5 + Math.random() * 5;
    else return 10 + Math.random() * 10;
  }

  for (let i = 0; i < numCustomers; i++) {
    const customerID = i + 1;
    const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
    const lastName  = lastNames[Math.floor(Math.random() * lastNames.length)];

    // Random home within radius, weighted
    const { lat: homeLat, lon: homeLon } = randomPointInRadius(startLocation, 20, sampleDistance);

    // 20% chance default drop-off
    let defaultDO = null;
    let defaultTripPurpose = null;
    let doLat = null;
    let doLon = null;
    if (Math.random() < 0.2) {
      const allPOIs = [];
      ['Medical','Work','Other'].forEach(cat => {
        poiPool[cat].forEach(o => allPOIs.push({...o, cat}));
      });
      const pick = allPOIs[Math.floor(Math.random() * allPOIs.length)];
      defaultDO = pick;
      doLat = pick.lat;
      doLon = pick.lon;
      defaultTripPurpose = pick.cat;
    }

    customers.push({
      "Customer ID": customerID,
      "Customer First Name": firstName,
      "Customer Last Name": lastName,
      "Customer Name and ID": `${lastName}, ${firstName} (${customerID})`,
      "Phone Number": '999-9999',
      "Default PU Address": `${homeLat}, ${homeLon}`,
      "Default DO Address": defaultDO ? `${defaultDO.lat}, ${defaultDO.lon}` : '',
      "Default Trip Purpose": defaultTripPurpose || '',
      homeLat,
      homeLon,
      doLat,
      doLon
    });
  }

  return customers;
}

/**
 * Generate runs for dummy data
 * @param {string} startDateStr - Start date in MM/DD/YYYY format
 * @param {string[]} driverIDs - Array of driver IDs
 * @param {string[]} vehicleIDs - Array of vehicle IDs
 * @return {Object[]} Array of run records with keys matching spreadsheet columns
 */
function generateRuns(startDateStr, driverIDs, vehicleIDs) {
  const runs = [];
  const startDate = new Date(startDateStr);
  for (let d = 0; d < 7; d++) {
    const date = new Date(startDate);
    date.setDate(startDate.getDate() + d);
    const dow = date.getDay(); // 0=Sun,6=Sat
    if (dow === 0) continue; // skip Sunday
    if (dow === 6 && Math.random < 0.15) continue; // fewer drivers on Saturday

    // Determine schedule window
    let startH, startM, endH, endM;
    if (dow >= 1 && dow <= 5) {
      if (Math.random() < 0.5) { startH = 8; startM = 30; endH = 16; endM = 0; }
      else                  { startH = 9; startM =  0; endH = 18; endM = 0; }
    } else { // Saturday
      startH = 9; startM = 0; endH = 17; endM = 0;
    }

    // Shuffle and pair drivers/vehicles
    const drs = driverIDs.slice();
    const vhs = vehicleIDs.slice();
    shuffleArray(drs);
    shuffleArray(vhs);
    const count = Math.min(drs.length, vhs.length);
    for (let i = 0; i < count; i++) {
      // Create Date objects for start/end
      const startTime = new Date(date.getFullYear(), date.getMonth(), date.getDate(), startH, startM);
      const endTime   = new Date(date.getFullYear(), date.getMonth(), date.getDate(), endH,   endM);
      runs.push({
        "Run Date": formatDate(date),
        "Driver ID": drs[i],
        "Vehicle ID": vhs[i],
        "Scheduled Start Time": startTime,
        "Scheduled End Time": endTime
      });
    }
  }
  return runs;
}

/**
 * Generate trips for dummy data
 * @param {Object[]} runs - Array of run records
 * @param {Object[]} customers - Array of customer records
 * @param {{Medical:Object[], Work:Object[], Other:Object[]}} poiPool - Pool of points of interest
 * @return {Object[]} Array of trip records with keys matching spreadsheet columns
 */
function generateTrips(runs, customers, poiPool) {
  const trips = [];

  // helper: haversine distance in miles
  function haversine(lat1, lon1, lat2, lon2) {
    const R = 3958.8; // Earth radius in miles
    const dLat = toRadians(lat2 - lat1);
    const dLon = toRadians(lon2 - lon1);
    const a = Math.sin(dLat/2)**2 + Math.cos(toRadians(lat1))*Math.cos(toRadians(lat2))*Math.sin(dLon/2)**2;
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
    return R * c;
  }

  // helper: add minutes to Date
  function addMinutes(date, mins) {
    return new Date(date.getTime() + mins * 60000);
  }

  // helper: round up to the nearest 15 minutes
  function roundToNearest15(date) {
    const mins = date.getMinutes();
    const roundedMins = Math.ceil(mins / 15) * 15;
    const roundedDate = new Date(date);
    roundedDate.setMinutes(roundedMins, 0, 0);
    return roundedDate;
  }

  // helper: sample from status distribution
  function sampleStatus() {
    const p = Math.random();
    if (p < 0.90) return 'Completed';
    if (p < 0.96) return 'Early Cancel';
    if (p < 0.98) return 'Late Cancel';
    return 'No Show';
  }

  // track per-date customer pool
  let prevRunDate = null;
  let pool = [];
  // iterate each run
  runs.forEach(run => {
    // copy pool for this day
    if (run["Run Date"] !== prevRunDate) {
      pool = customers.slice();
      prevRunDate = run["Run Date"];
    }
    let cursorTime = new Date(run["Scheduled Start Time"]);
    // ensure buffer
    if (addMinutes(cursorTime, 30) > run["Scheduled End Time"]) return;

    let firstJourney = true;
    // schedule journeys
    while (addMinutes(cursorTime, 30) <= run["Scheduled End Time"] && pool.length) {
      // 1. pickup time
      let pu1;
      if (firstJourney) {
        pu1 = addMinutes(cursorTime, randBetween(10, 60));
        firstJourney = false;
      } else {
        pu1 = addMinutes(cursorTime, randBetween(15, 45));
      }
      // if beyond end, break
      if (addMinutes(pu1, 0) > run["Scheduled End Time"]) break;

      // 2. choose customer
      const custIdx = Math.floor(Math.random() * pool.length);
      const cust = pool.splice(custIdx, 1)[0];

      // 3. purpose
      const purposes = ['Medical','Work','Other'];
      const purpose = purposes[Math.floor(Math.random()*purposes.length)];

      // 4. select do1
      let do1;
      if (cust.doLat && cust.doLon && cust["Default Trip Purpose"] === purpose) {
        do1 = { lat: cust.doLat, lon: cust.doLon };
      } else {
        const list = poiPool[purpose];
        do1 = list[Math.floor(Math.random() * list.length)];
      }

      // 5. compute dist & duration
      const dist1 = haversine(cust.homeLat, cust.homeLon, do1.lat, do1.lon);
      const speed = dist1 > 30 ? 50 : 30; // mph
      const dur1 = Math.round((dist1/speed)*60 + 5);

      // 6. sample status
      const status1 = sampleStatus();
      if (status1 !== 'Completed') {
        // record a single trip
        trips.push({ 
          "Trip Date": run["Run Date"],
          "Customer Name and ID": cust["Customer Name and ID"],
          "PU Address": `${cust.homeLat}, ${cust.homeLon}`,
          "DO Address": `${do1.lat}, ${do1.lon}`,
          "PU Time": pu1,
          "Appt Time": addMinutes(pu1, dur1),
          "Vehicle ID": run["Vehicle ID"],
          "Driver ID": run["Driver ID"],
          "Trip Purpose": purpose,
          "Trip Result": status1,
          distance: dist1,
          duration: dur1,
          "Est Hours": (dur1 / 60).toFixed(2),
          "Est Miles": dist1.toFixed(1)
        });
        // advance cursorTime based on cancel type
        if (status1 === 'Late Cancel') {
          cursorTime = addMinutes(pu1, 15);
        } else if (status1 === 'No Show') {
          cursorTime = new Date(pu1);
        }
        continue; // next journey
      }

      // 7. record trip1
      // const roundedDur1 = roundTo15(dur1);
      const doTime1 = addMinutes(pu1, dur1);
      const aptTime1 = roundToNearest15(addMinutes(doTime1,2));
      trips.push({ 
        "Trip Date": run["Run Date"],
        "Customer Name and ID": cust["Customer Name and ID"],
        "PU Address": `${cust.homeLat}, ${cust.homeLon}`,
        "DO Address": `${do1.lat}, ${do1.lon}`,
        "PU Time": pu1,
        "DO Time": doTime1,
        "Appt Time": aptTime1,
        "Vehicle ID": run["Vehicle ID"],
        "Driver ID": run["Driver ID"],
        "Trip Purpose": purpose,
        "Trip Result": 'Completed',
        distance: dist1,
        duration: dur1,
        "Est Hours": (dur1 / 60).toFixed(2),
        "Est Miles": dist1.toFixed(1)
      });

      // 8. optional intermediate stop
      let lastDropoff = doTime1;
      let lastLocation = do1;
      if (purpose === 'Other' && Math.random() < 0.25) {
        const others = poiPool.Other.filter(o => o.lat !== do1.lat || o.lon !== do1.lon);
        const poiB = others[Math.floor(Math.random()*others.length)];
        const dist2 = haversine(do1.lat, do1.lon, poiB.lat, poiB.lon);
        const speed2 = dist2 > 30 ? 50 : 30;
        const dur2 = Math.round((dist2/speed2)*60 + 5);
        // const roundedDur2 = roundTo15(dur2);
        const pu2 = new Date(doTime1);
        const doTime2 = addMinutes(pu2, dur2);
        const aptTime2 = roundToNearest15(addMinutes(doTime2, 2));
        trips.push({ 
          "Trip Date": run["Run Date"],
          "Customer Name and ID": cust["Customer Name and ID"],
          "PU Address": `${do1.lat}, ${do1.lon}`,
          "DO Address": `${poiB.lat}, ${poiB.lon}`,
          "PU Time": pu2,
          "DO Time": doTime2,
          "Appt Time": aptTime2,
          "Vehicle ID": run["Vehicle ID"],
          "Driver ID": run["Driver ID"],
          "Trip Purpose": 'Other',
          "Trip Result": 'Completed',
          distance: dist2,
          duration: dur2,
          "Est Hours": (dur2 / 60).toFixed(2),
          "Est Miles": dist2.toFixed(2)
        });
        lastDropoff = doTime2;
        lastLocation = poiB;
      }

      // 9. return leg
      const bufferMins = [30, 60, 90][Math.floor(Math.random() * 3)];
      const puReturn = addMinutes(lastDropoff, bufferMins);
      const distR = haversine(lastLocation.lat, lastLocation.lon, cust.homeLat, cust.homeLon);
      const speedR = distR > 30 ? 50 : 30;
      const durR = Math.round((distR/speedR)*60 + 5);
      //const roundedDurR = roundTo15(durR);
      const doReturn = addMinutes(puReturn, durR);
      trips.push({ 
        "Trip Date": run["Run Date"],
        "Customer Name and ID": cust["Customer Name and ID"],
        "PU Address": `${lastLocation.lat}, ${lastLocation.lon}`,
        "DO Address": `${cust.homeLat}, ${cust.homeLon}`,
        "PU Time": puReturn,
        "DO Time": doReturn,
        "Appt Time": null,
        "Vehicle ID": run["Vehicle ID"],
        "Driver ID": run["Driver ID"],
        "Trip Purpose": purpose,
        "Trip Result": 'Completed',
        distance: distR,
        duration: durR,
        "Est Hours": (durR / 60).toFixed(2),
        "Est Miles": distR.toFixed(2)
      });

      // 10. advance cursor
      cursorTime = doReturn;
    }
  });

  return trips;
}

// helper random between min and max inclusive
function randBetween(min, max) {
  return Math.floor(min + Math.random() * (max - min + 1));
}

