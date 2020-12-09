
function getRuns() {
  try {
    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasRuns) {
        let params = {}
        params.resource = "runs"
        params.version = endPoint.version
        params.apiKey = endPoint.apiKey
        let hmac = {}
        hmac.nonce = Utilities.getUuid()
        hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
        hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)
        
        const options = {method: 'GET', contentType: 'application/json'}
        let response = UrlFetchApp.fetch(endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac), options)
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
        } catch(e) {
          responseObject = {status: "LOCAL_ERROR:" + e.name}
        }
          
        let runArray = []
        if (responseObject.runs) {
          responseObject.runs.forEach(run => {
            runArray.push(["Source", endPoint.name, "Run Date", run.runDate])
            runArray.push(["Seats", run.ambulatorySpacePoints, "Wheelchair spaces", run.standardWheelchairSpacePoints])
            runArray.push(["Lift", run.hasLift ? "Yes" : "No", "Ramp", run.hasLift ? "Yes" : "No"])
            runArray.push(["Start location", run.startLocation, "End location", run.endLocation])
            run.stops.forEach(stop => {
              runArray.push([null, stop.time, stop.city, stop.riderChange])
            })
          })
        }
        log("Received from GET:",JSON.stringify(runArray))
      }
    })
  } catch(e) { logError(e) }
}

function getTrips() {
  try {
    endPoints = getDocProp("apiGetAccess")
    endPoints.forEach(endPoint => {
      if (endPoint.hasTrips) {
        let params = {}
        params.resource = "trips"
        params.version = endPoint.version
        params.apiKey = endPoint.apiKey
        let hmac = {}
        hmac.nonce = Utilities.getUuid()
        hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
        hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)
        
        const options = {method: 'GET', contentType: 'application/json'}
        let response = UrlFetchApp.fetch(endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac), options)
        let responseObject
        try {
          responseObject = JSON.parse(response.getContentText())
        } catch(e) {
          responseObject = {status: "LOCAL_ERROR:" + e.name}
        }
          
        let tripArray = []
        if (responseObject.trips) {
          responseObject.trips.forEach(trip => {
          })
        }
        log("Received from GET:",JSON.stringify(responseObject))
      }
    })
  } catch(e) { logError(e) }
}