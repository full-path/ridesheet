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

function getGooglePlacesAddresses(startLocation, radiusMiles = 50, placeTypes, purpose, limitPerType) {
  const GOOGLE_MAPS_API_KEY = PropertiesService.getScriptProperties().getProperty('GOOGLE_MAPS_API_KEY');
  
  if (!GOOGLE_MAPS_API_KEY) {
    Logger.log('ERROR: GOOGLE_MAPS_API_KEY not set in Script Properties');
    return [];
  }

  // Get coordinates for the starting location
  const coords = getGeocode(startLocation, "coordinates");
  if (!coords || coords.startsWith("Error")) {
    Logger.log(`ERROR: Could not geocode starting location: ${startLocation}`);
    return [];
  }

  const radiusMeters = Math.round(radiusMiles * 1609.34); // Convert miles to meters
  const [lat, lng] = coords.split(',').map(c => parseFloat(c.trim()));
  
  const allPlaces = [];
  
  // Query each place type
  for (const placeType of placeTypes) {
    const places = queryPlacesByType(lat, lng, radiusMeters, placeType, limitPerType, GOOGLE_MAPS_API_KEY);
    allPlaces.push(...places);
    
    // Check if we've hit the overall limit
    if (allPlaces.length >= limitPerType * placeTypes.length) {
      break;
    }
    
    Utilities.sleep(500); // Rate limiting between type queries
  }

  // Limit total results
  const limitedPlaces = allPlaces.slice(0, limitPerType * placeTypes.length);

  const newAddresses = limitedPlaces.map((place) => {
    return {
      "Short Name": createAddressShortName(place.name),
      "Address": `${place.formatted_address || place.vicinity} (${place.name})`,
      "Default Trip Purpose": purpose
    }
  });

  Logger.log(`Received ${newAddresses.length} ${purpose} addresses`);
  return newAddresses;
}

function queryPlacesByType(lat, lng, radiusMeters, placeType, limit, apiKey) {
  const endpoint = 'https://maps.googleapis.com/maps/api/place/nearbysearch/json';
  
  let allResults = [];
  let nextPageToken = null;
  
  do {
    const params = {
      location: `${lat},${lng}`,
      radius: radiusMeters,
      type: placeType,
      key: apiKey
    };
    
    if (nextPageToken) {
      params.pagetoken = nextPageToken;
      Utilities.sleep(2000); // Required delay for next page token
    }
    
    const url = endpoint + '?' + Object.keys(params).map(key => 
      `${key}=${encodeURIComponent(params[key])}`
    ).join('&');
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true
    });
    
    let data = {};
    try {
      data = JSON.parse(response.getContentText());
    } catch(e) {
      Logger.log(`Error parsing response: ${e.name}: ${e.message}`);
      Logger.log(response.getContentText());
      break;
    }
    
    if (data.status !== 'OK' && data.status !== 'ZERO_RESULTS') {
      Logger.log(`API Error: ${data.status} - ${data.error_message || 'No error message'}`);
      break;
    }
    
    const results = data.results || [];
    allResults.push(...results);
    
    // Check if we have enough results or if there's a next page
    if (allResults.length >= limit || !data.next_page_token) {
      break;
    }
    
    nextPageToken = data.next_page_token;
    
  } while (nextPageToken && allResults.length < limit);
  
  return allResults.slice(0, limit);
}
