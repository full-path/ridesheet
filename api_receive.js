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
          content.results = receiveRequestForTripRequests()
        } else if (params.resource === "tripRequestResponses") {
          content.status = "OK"
          content.results = receiveTripRequestResponses()
        } else if (params.resource === "clientOrderConfirmations") {
          content.status = "OK"
          content.results = receiveClientOrderConfirmations()
        } else if (params.resource === "providerOrderConfirmations") {
          content.status = "OK"
          content.results = receiveProviderOrderConfirmations()
        } else if (params.resource === "customerInfo") {
          content.status = "OK"
          content.results = receiveCustomerInfo()
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
  let output = ContentService.createTextOutput()
  output.setContent(JSON.stringify(e))
  output.setMimeType(ContentService.MimeType.JSON)
  return output
}
