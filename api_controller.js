function doGet(e) {
  try {
    const timeNow = new Date()
    const params = e.parameter
    const hmacParamKeys = ["nonce","timestamp","signature"]
    let baseParams = {}
    let validatedApiAccount
    if (params.apiKey) {
      const apiAccounts = getDocProp("apiGiveAccess")
      const statedApiAccount = apiAccounts[params.apiKey]
      Object.keys(params).forEach(key => {
        if (hmacParamKeys.indexOf(key) === -1) baseParams[key] = params[key]
      })
      let receivedTimestamp = new Date(params.timestamp)
      let timePassed = timeNow.getTime() - receivedTimestamp.getTime() // In milliseconds
      if (params.signature === generateHmacHexString(statedApiAccount.secret, params.nonce, params.timestamp, baseParams) && 
          timePassed < 300000) {
        validatedApiAccount = statedApiAccount
      }
    }
  
    let response = ContentService.createTextOutput()
    response.setMimeType(ContentService.MimeType.JSON)
    let content = {}
    if (validatedApiAccount) {
      if (params.version === "v1") {
        if (params.resource === "runs") {
          content.status = "OK"
          content.results = receiveRequestForRuns()
        } else if (params.resource === "tripRequests") {
          content.status = "OK"
          content.results = receiveRequestForTripRequestsReturnTripRequests(validatedApiAccount)
        } else {
          content.status = "INVALID_REQUEST"
        }
      } else {
        content.status = "INVALID_REQUEST"
      }
    } else {
      content.status = "UNAUTHORIZED"
    }
    response.setContent(JSON.stringify(content))
    return response
  } catch(e) { logError(e) }
}
  
function doPost(e) {
  try {
    const timeNow = new Date()
    const params = e.parameter
    const hmacParamKeys = ["nonce","timestamp","signature"]
    let baseParams = {}
    let validatedApiAccount
    if (params.apiKey) {
      const apiAccounts = getDocProp("apiGiveAccess")
      const statedApiAccount = apiAccounts[params.apiKey]
      Object.keys(params).forEach(key => {
        if (hmacParamKeys.indexOf(key) === -1) baseParams[key] = params[key]
      })
      let receivedTimestamp = new Date(params.timestamp)
      let timePassed = timeNow.getTime() - receivedTimestamp.getTime() // In milliseconds
      if (params.signature === generateHmacHexString(statedApiAccount.secret, params.nonce, params.timestamp, baseParams) && 
          timePassed < 300000) {
        validatedApiAccount = statedApiAccount
      }
    }

    let response = ContentService.createTextOutput()
    response.setMimeType(ContentService.MimeType.JSON)
    let content = {}
    if (validatedApiAccount) {
      let payload = e.postData.contents
      if (params.version === "v1") {
        if (params.resource === "tripRequestResponses") {
          const processedResponses = receiveTripRequestResponses(JSON.parse(payload))
          content = returnClientOrderConfirmations(processedResponses, validatedApiAccount)
        } else if (params.resource === "providerOrderConfirmations") {
          content = receiveProviderOrderConfirmationsReturnCustomerInformation(payload, validatedApiAccount)
        } else {
          content.status = "INVALID_REQUEST"
        }
      } else {
        content.status = "INVALID_REQUEST"
      }
    } else {
      content.status = "UNAUTHORIZED"
    }
    response.setContent(JSON.stringify(content))
    return response
  } catch(e) { logError(e) }
}
