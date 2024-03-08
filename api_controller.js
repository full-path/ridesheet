function doGet(e) {
  try {
    const params = e.parameter
    const hmacHeader = params.authorization
    const pathHeader = params.endpointPath || ''

    // TODO: match error response with actual TDS spec
    if (!hmacHeader) {
      return createErrorResponse("403", "MISSING_AUTHORIZATION_HEADER")
    }

    const [signature, senderId, receiverId, timestamp, nonce] = hmacHeader.split(':')

    if (!(signature && senderId && receiverId && timestamp && nonce)) {
      return createErrorResponse("403", "INVALID_AUTHORIZATION_HEADER")
    }

    const apiAccounts = getDocProp("apiGiveAccess")
    const apiAccount = apiAccounts[senderId]

    if (!apiAccount) {
      return createErrorResponse("403", "INVALID_API_KEY")
    }

    const isValid = validateHmacSignature(signature, senderId, receiverId, timestamp, nonce, 'GET', '', pathHeader)

    if (!isValid) {
      return createErrorResponse("403", "UNAUTHORIZED")
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
      return createErrorResponse("400","INVALID_REQUEST");
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
      return createErrorResponse("403", "MISSING_AUTHORIZATION_HEADER")
    }

    const [signature, senderId, receiverId, timestamp, nonce] = hmacHeader.split(':')

    if (!(signature && senderId && receiverId && timestamp && nonce)) {
      return createErrorResponse("403","INVALID_AUTHORIZATION_HEADER")
    }

    const apiAccounts = getDocProp("apiGiveAccess")
    const apiAccount = apiAccounts[senderId]

    if (!apiAccount) {
      return createErrorResponse("403", "INVALID_API_KEY")
    }

    const body = e.postData.contents

    const isValid = validateHmacSignature(signature, senderId, receiverId, timestamp, nonce, 'POST', body, pathHeader)

    if (!isValid) {
      return createErrorResponse("403", "Unauthorized")
    }

    const payload = JSON.parse(body)

    let response = ContentService.createTextOutput()
    response.setMimeType(ContentService.MimeType.JSON)
    let content = {}
    if (params.endpointPath) {
      if (params.endpointPath === "/v1/TripRequest") {
        content = receiveTripRequest(payload, senderId)
      } else if (params.endpointPath === "/v1/TripRequestResponse") {
        content = receiveTripRequestResponse(payload, senderId)
      } else if (params.endpointPath === "/v1/ClientOrderConfirmation") {
        content = receiveClientOrderConfirmation(payload)
      } else if (params.endpointPath === "/v1/CustomerReferral") {
        content = receiveCustomerReferral(payload, senderId)
      } else if (params.endpointPath === "/v1/TripStatusChange") {
        content = receiveTripStatusChange(payload, senderId)
      } else if (params.endpointPath === "/v1/ProviderOrderConfirmation") {
        content = receiveProviderOrderConfirmation(payload, senderId)
      } else if (params.endpointPath === "/v1/CustomerReferralResponse") {
        content = receiveCustomerReferralResponse(payload, senderId)
      } else if (params.endpointPath === "/v1/TripTaskCompletion") {
        content = receiveTripTaskCompletion(payload)
      } else {
        return createErrorResponse("400", "Invalid request")
      }
    } else {
      // Old ridesheet "resource"-based API
      if (params.resource === "tripRequestResponses") {
        const processedResponses = receiveTripRequestResponses(payload)
        content = returnClientOrderConfirmations(processedResponses, apiAccount)
      } else if (params.resource === "providerOrderConfirmations") {
        content = receiveProviderOrderConfirmationsReturnCustomerInformation(payload, apiAccount)
      } else {
        return createErrorResponse("400", "Invalid request")
      }
    }
    response.setContent(JSON.stringify(content))
    return response
  } catch(e) { logError(e) }
}


