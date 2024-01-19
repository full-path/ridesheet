function doGet(e) {
  try {
    const params = e.parameter
    const hmacHeader = params.authorization
    const pathHeader = params.endpointPath || ''

    // TODO: match error response with actual TDS spec
    if (!hmacHeader) {
      return createErrorResponse("MISSING_AUTHORIZATION_HEADER")
    }

    const [signature, senderId, receiverId, timestamp, nonce] = hmacHeader.split(':')

    if (!(signature && senderId && receiverId && timestamp && nonce)) {
      return createErrorResponse("INVALID_AUTHORIZATION_HEADER")
    }

    const apiAccounts = getDocProp("apiGiveAccess")
    const apiAccount = apiAccounts[senderId]

    if (!apiAccount) {
      return createErrorResponse("INVALID_API_KEY")
    }

    const isValid = validateHmacSignature(signature, senderId, receiverId, timestamp, nonce, 'GET', '', pathHeader)

    if (!isValid) {
      return createErrorResponse("UNAUTHORIZED")
    }

    let response = ContentService.createTextOutput()
    response.setMimeType(ContentService.MimeType.JSON)
    let content = {}

    if (params.resource === "runs") {
      content.status = "OK"
      content.results = receiveRequestForRuns()
    } else if (params.resource === "tripRequests") {
      content.status = "OK"
      content.results = receiveRequestForTripRequestsReturnTripRequests(apiAccount)
    } else {
      return createErrorResponse("INVALID_REQUEST");
    }

    response.setContent(JSON.stringify(content))
    return response
  } catch(e) { 
    logError(e) 
  }
}
  
function doPost(e) {
  try {
    const params = e.parameter
    const hmacHeader = params.authorization
    const pathHeader = params.endpointPath || ''

    // TODO: match error response with actual TDS spec
    if (!hmacHeader) {
      return createErrorResponse("MISSING_AUTHORIZATION_HEADER")
    }

    const [signature, senderId, receiverId, timestamp, nonce] = hmacHeader.split(':')

    if (!(signature && senderId && receiverId && timestamp && nonce)) {
      return createErrorResponse("INVALID_AUTHORIZATION_HEADER")
    }

    const apiAccounts = getDocProp("apiGiveAccess")
    const apiAccount = apiAccounts[senderId]

    if (!apiAccount) {
      return createErrorResponse("INVALID_API_KEY")
    }

    const payload = JSON.parse(e.postData.contents)

    const isValid = validateHmacSignature(signature, senderId, receiverId, timestamp, nonce, 'POST', payload, pathHeader)

    if (!isValid) {
      return createErrorResponse("UNAUTHORIZED")
    }

    let response = ContentService.createTextOutput()
    response.setMimeType(ContentService.MimeType.JSON)
    let content = {}
    if (params.endpointPath) {
      if (params.endpointPath === "/v1/TripRequest") {
        content = receiveTripRequest(payload)
      } else {
        return createErrorResponse("INVALID_REQUEST")
      }
    } else {
      // Old ridesheet "resource"-based API
      if (params.resource === "tripRequestResponses") {
        const processedResponses = receiveTripRequestResponses(payload)
        content = returnClientOrderConfirmations(processedResponses, apiAccount)
      } else if (params.resource === "providerOrderConfirmations") {
        content = receiveProviderOrderConfirmationsReturnCustomerInformation(payload, apiAccount)
      } else {
        return createErrorResponse("INVALID_REQUEST")
      }
    }
    response.setContent(JSON.stringify(content))
    return response
  } catch(e) { logError(e) }
}


