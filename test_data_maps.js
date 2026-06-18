/**
 * @fileoverview Geographic address-lookup strategies for RideSheet test data generation.
 *
 * Provides four interchangeable strategies for discovering real POI addresses near a
 * given location. All strategies return arrays of address objects with the same shape:
 * `{ "Short Name": string, "Address": string, "Default Trip Purpose": string }`.
 *
 * Strategies:
 * - **Geocode fishing** (`getGeocodeFishedAddresses`) — Google Maps geocoder, free-text terms.
 * - **Nominatim** (`getNominatimAddresses`) — OpenStreetMap Nominatim structured search.
 * - **Overpass API** (`getOsmAddresses`) — OpenStreetMap Overpass QL query with server fallbacks.
 * - **Google Places** (`getGooglePlacesAddresses`, `queryPlacesByType`) — Google Places API v1.
 *
 * All strategies are called from `generatePoiAddresses()` in test_data.js based on
 * the `params.addressSource` value.
 */

/**
 * Searches for nearby establishments using the Google Maps geocoder and free-text terms.
 * For each term, geocodes `"<term> <city> <state>"` and keeps results typed as
 * `point_of_interest` or `establishment` that include a street number.
 * @param {{lat: number, lng: number}} startLocation - Center point for the search.
 * @param {number} [radiusMiles=50] - Bounding box radius in miles used to bias results.
 * @param {string[]} searchTerms - Free-text search terms (e.g. `"hospital"`, `"city hall"`).
 * @param {string} purpose - `"Default Trip Purpose"` value assigned to every returned address.
 * @param {number} limit - Maximum number of addresses to return across all terms.
 * @returns {Object[]} Address objects with `Short Name`, `Address`, and `Default Trip Purpose`.
 */
function getGeocodeFishedAddresses(startLocation, radiusMiles = 50, searchTerms, purpose, limit) {
  const centerStr = `${startLocation.lat},${startLocation.lng}`
  const locationInfo = getGeocode(centerStr, "object")
  if (!locationInfo || locationInfo.status !== "OK") {
    Logger.log(`ERROR: Could not reverse geocode starting location for POI fishing`)
    return []
  }
  const locationContext = `${locationInfo.city} ${locationInfo.state}`

  const bbox = computeBoundingBox(startLocation, radiusMiles)
  const geocoder = Maps.newGeocoder().setBounds(bbox.south, bbox.west, bbox.north, bbox.east)

  const allAddresses = []
  const seenAddresses = new Set()

  for (const term of searchTerms) {
    if (allAddresses.length >= limit) break

    try {
      const mapsResults = geocoder.geocode(`${term} ${locationContext}`)
      if (mapsResults.status !== "OK") {
        Logger.log(`No geocode results for "${term} ${locationContext}": ${mapsResults.status}`)
        Utilities.sleep(500)
        continue
      }

      for (const result of mapsResults.results) {
        if (!result.types.some(t => t === "point_of_interest" || t === "establishment")) continue
        if (!result.address_components.some(c => c.types.includes("street_number"))) continue
        if (seenAddresses.has(result.formatted_address)) continue

        seenAddresses.add(result.formatted_address)

        const establishmentComp = result.address_components.find(c => c.types.includes("establishment"))
        const name = establishmentComp ? establishmentComp.long_name : term

        allAddresses.push({
          "Short Name": createAddressShortName(name),
          "Address": `${result.formatted_address} (${name})`,
          "Default Trip Purpose": purpose
        })

        if (allAddresses.length >= limit) break
      }
    } catch(e) {
      Logger.log(`Error geocoding "${term} ${locationContext}": ${e.message}`)
    }

    Utilities.sleep(500)
  }

  Logger.log(`Geocode fishing found ${allAddresses.length} ${purpose} addresses`)
  return allAddresses
}

/**
 * Searches for nearby named places using the OpenStreetMap Nominatim search API.
 * Issues one HTTP request per `{key, value}` OSM tag pair (e.g. `{key: "amenity", value: "hospital"}`),
 * respecting Nominatim's 1-request-per-second rate limit. Each result is then geocoded
 * through Google Maps to produce a canonical formatted address. Results without a name
 * or a street address in OSM data are skipped.
 * @param {{lat: number, lng: number}} startLocation - Center point for the search.
 * @param {number} [radiusMiles=50] - Viewbox radius in miles.
 * @param {Array<{key: string, value: string}>} osmTags - OSM tag filters to query in sequence.
 * @param {string} purpose - `"Default Trip Purpose"` value assigned to every returned address.
 * @param {number} limit - Maximum number of addresses to return across all tag queries.
 * @returns {Object[]} Address objects with `Short Name`, `Address`, and `Default Trip Purpose`.
 */
function getNominatimAddresses(startLocation, radiusMiles = 50, osmTags, purpose, limit) {
  const bbox = computeBoundingBox(startLocation, radiusMiles)
  // Nominatim viewbox format: west,north,east,south
  const viewbox = `${bbox.west},${bbox.north},${bbox.east},${bbox.south}`

  const allAddresses = []
  const seenDisplayNames = new Set()

  for (const {key, value} of osmTags) {
    if (allAddresses.length >= limit) break

    const url = `https://nominatim.openstreetmap.org/search?${key}=${encodeURIComponent(value)}&bounded=1&viewbox=${viewbox}&format=json&addressdetails=1&limit=${limit}`

    try {
      const response = UrlFetchApp.fetch(url, {
        headers: { 'User-Agent': 'RideSheet/1.0 (https://github.com/full-path/ridesheet-demo)' },
        muteHttpExceptions: true
      })

      const responseCode = response.getResponseCode()
      if (responseCode !== 200) {
        Logger.log(`Nominatim error for ${key}=${value}: HTTP ${responseCode}`)
        Utilities.sleep(1000)
        continue
      }

      const results = JSON.parse(response.getContentText())

      for (const result of results) {
        if (allAddresses.length >= limit) break
        if (seenDisplayNames.has(result.display_name)) continue
        seenDisplayNames.add(result.display_name)

        const addr = result.address || {}
        const name = result.name

        const houseNumber = addr.house_number || ''
        const road = addr.road || ''
        const city = addr.city || addr.town || addr.village || addr.hamlet || ''
        const postcode = addr.postcode || ''

        if (!name || !houseNumber || !road) {
          Logger.log(`Skipping "${result.display_name}": ${!name ? 'no name' : 'no street address'} in OSM data`)
          continue
        }

        const addressQuery = `${houseNumber} ${road} ${city} ${postcode}`.trim()
        const formattedAddress = getGeocode(addressQuery, "formatted_address")

        if (formattedAddress.startsWith("Error")) {
          Logger.log(`Could not geocode "${name}": ${formattedAddress}`)
          continue
        }

        allAddresses.push({
          "Short Name": createAddressShortName(name),
          "Address": `${formattedAddress} (${name})`,
          "Default Trip Purpose": purpose
        })
      }

    } catch(e) {
      Logger.log(`Error querying Nominatim for ${key}=${value}: ${e.message}`)
    }

    Utilities.sleep(1000) // Nominatim rate limit: 1 req/sec
  }

  Logger.log(`Nominatim found ${allAddresses.length} ${purpose} addresses`)
  return allAddresses
}

/**
 * Queries the OpenStreetMap Overpass API using an OverpassQL filter expression.
 * Tries four public Overpass servers in sequence, moving to the next on rate-limit
 * (429), service-unavailable (503), or other 5xx responses. Each matching OSM element
 * is geocoded through Google Maps to produce a canonical formatted address.
 * @param {{lat: number, lng: number}} startLocation - Center point for the bounding box.
 * @param {number} [radiusMiles=50] - Bounding box radius in miles.
 * @param {string} osmQuery - OverpassQL tag filter string, e.g.
 *   `'["amenity"~"clinic|hospital"]["name"]["addr:housenumber"]'`.
 * @param {string} purpose - `"Default Trip Purpose"` value assigned to every returned address.
 * @param {number} limit - Maximum number of results requested from Overpass (`out tags <limit>`).
 * @returns {Object[]} Address objects with `Short Name`, `Address`, and `Default Trip Purpose`,
 *   or an empty array if all servers fail or return no results.
 */
function getOsmAddresses(startLocation, radiusMiles = 50, osmQuery, purpose, limit) {
  const bboxObj = computeBoundingBox(startLocation, radiusMiles);
  const bbox = `${bboxObj.south},${bboxObj.west},${bboxObj.north},${bboxObj.east}`;

  // Multiple public Overpass API servers as fallbacks
  const endpoints = [
    'https://overpass-api.de/api/interpreter',
    'https://overpass.kumi.systems/api/interpreter',
    'https://overpass.private.coffee/api/interpreter',
    'https://overpass.nchc.org.tw/api/interpreter'
  ];

  const overpassQL = `
    [out:json][timeout:60];
    (
      node${osmQuery}(${bbox});
      way${osmQuery}(${bbox});
      relation${osmQuery}(${bbox});
    );
    out tags ${limit};
  `;

  // Try each endpoint until one succeeds
  for (let i = 0; i < endpoints.length; i++) {
    const endpoint = endpoints[i];
    Logger.log(`Trying ${purpose} query on server ${i + 1}/${endpoints.length}: ${endpoint}`);

    try {
      const response = UrlFetchApp.fetch(endpoint, {
        method: 'post',
        payload: { data: overpassQL },
        muteHttpExceptions: true,
        timeout: 30 // 30 second timeout
      });

      const responseCode = response.getResponseCode();

      // Check for rate limiting or server errors
      if (responseCode === 429 || responseCode === 503 || responseCode >= 500) {
        Logger.log(`Server ${endpoint} returned ${responseCode}, trying next server...`);
        continue;
      }

      let osmData = {};
      try {
        osmData = JSON.parse(response.getContentText());
      } catch(e) {
        Logger.log(`Parse error on ${endpoint}: ${e.message}`);
        continue;
      }

      const osmElements = osmData.elements || [];

      if (osmElements.length === 0) {
        Logger.log(`No results from ${endpoint}, but query succeeded`);
        return []; // Valid empty result
      }

      // Process addresses
      osmElements.forEach((elem) => {
        const tags = elem.tags;
        const googleMapsQuery = `${tags["addr:housenumber"]} ${tags["addr:street"]} ${tags["addr:city"]} ${tags["addr:postcode"]}`;
        const formattedAddress = getGeocode(googleMapsQuery, "formatted_address");
        if (formattedAddress.startsWith("Error")) {
          Logger.log(JSON.stringify(tags, null, 2));
        }
        elem.formattedAddress = formattedAddress;
      });

      const newAddresses = osmElements.map((elem) => {
        return {
          "Short Name": createAddressShortName(elem.tags.name),
          "Address": `${elem.formattedAddress} (${elem.tags.name})`,
          "Default Trip Purpose": purpose
        }
      });

      Logger.log(`✓ Successfully received ${newAddresses.length} ${purpose} addresses from ${endpoint}`);
      return newAddresses;

    } catch(e) {
      Logger.log(`Error with ${endpoint}: ${e.name} - ${e.message}`);
      if (i === endpoints.length - 1) {
        // Last server failed
        Logger.log(`All Overpass servers failed for ${purpose} query`);
        return [];
      }
      // Try next server
      Utilities.sleep(1000);
    }
  }

  Logger.log(`All ${endpoints.length} servers failed for ${purpose}`);
  return [];
}

/**
 * Searches for nearby places using the Google Places API v1 (`places:searchNearby`).
 * Requires `GOOGLE_MAPS_API_KEY` to be set in Script Properties. Issues one API call
 * per entry in `placeTypes` via `queryPlacesByType`, then trims the combined results
 * to `limitPerType * placeTypes.length`.
 * @param {string} startLocation - A geocodable address string used as the search center.
 * @param {number} [radiusMiles=50] - Search radius in miles.
 * @param {string[]} placeTypes - Google Place type identifiers (e.g. `"hospital"`, `"dentist"`).
 * @param {string} purpose - `"Default Trip Purpose"` value assigned to every returned address.
 * @param {number} limitPerType - Maximum results to request per place type.
 * @returns {Object[]} Address objects with `Short Name`, `Address`, and `Default Trip Purpose`,
 *   or an empty array if `GOOGLE_MAPS_API_KEY` is missing or the starting location cannot be geocoded.
 */
function getGooglePlacesAddresses(startLocation, radiusMiles = 50, placeTypes, purpose, limitPerType) {
  const GOOGLE_MAPS_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');

  if (!GOOGLE_MAPS_API_KEY) {
    Logger.log('ERROR: GOOGLE_MAPS_API_KEY not set in Script Properties');
    return [];
  }

  const geoResult = getGeocode(startLocation, "object");
  if (!geoResult || geoResult.status !== "OK") {
    Logger.log(`ERROR: Could not geocode starting location: ${startLocation}`);
    return [];
  }
  const { lat, lng } = geoResult;

  const radiusMeters = Math.round(radiusMiles * 1609.34);

  const allPlaces = [];

  // Query each place type
  for (const placeType of placeTypes) {
    const places = queryPlacesByType(lat, lng, radiusMeters, placeType, limitPerType, GOOGLE_MAPS_API_KEY);
    allPlaces.push(...places);
    Utilities.sleep(500); // Rate limiting between type queries
  }

  // Limit total results
  const limitedPlaces = allPlaces.slice(0, limitPerType * placeTypes.length);

  const newAddresses = limitedPlaces.map((place) => {
    const name = place.displayName?.text || "";
    return {
      "Short Name": createAddressShortName(name),
      "Address": `${place.formattedAddress} (${name})`,
      "Default Trip Purpose": purpose
    }
  });

  Logger.log(`Received ${newAddresses.length} ${purpose} addresses`);
  return newAddresses;
}

/**
 * Performs a single Google Places API v1 `searchNearby` request for one place type.
 * Helper for `getGooglePlacesAddresses`; not intended to be called directly.
 * Returns raw Places API place objects (`{ displayName, formattedAddress }`), not the
 * address-object shape used by the rest of this file.
 * @param {number} lat - Latitude of the search center.
 * @param {number} lng - Longitude of the search center.
 * @param {number} radiusMeters - Search radius in meters (capped at 50,000 by the API).
 * @param {string} placeType - A single Google Place type identifier (e.g. `"hospital"`).
 * @param {number} limit - Maximum number of results (capped at 20 by the API).
 * @param {string} apiKey - Google Maps API key with Places API enabled.
 * @returns {Object[]} Raw place objects from the Places API, or an empty array on error.
 */
function queryPlacesByType(lat, lng, radiusMeters, placeType, limit, apiKey) {
  const endpoint = 'https://places.googleapis.com/v1/places:searchNearby';

  const payload = {
    includedTypes: [placeType],
    maxResultCount: Math.min(limit, 20),
    locationRestriction: {
      circle: {
        center: { latitude: lat, longitude: lng },
        radius: radiusMeters
      }
    }
  };

  let data = {};
  try {
    const response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-Goog-Api-Key': apiKey,
        'X-Goog-FieldMask': 'places.displayName,places.formattedAddress'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      Logger.log(`API Error ${responseCode}: ${response.getContentText()}`);
      return [];
    }

    data = JSON.parse(response.getContentText());
  } catch(e) {
    Logger.log(`Error fetching/parsing response: ${e.name}: ${e.message}`);
    return [];
  }

  return (data.places || []).slice(0, limit);
}
