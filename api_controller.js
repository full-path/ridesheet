function doGet(e) {
  try {
    const params = e.parameter
    const hmacHeader = params.authorization
    const pathHeader = params.endpointPath

    // TODO: match error response with actual TDS spec
    if (!(hmacHeader && pathHeader)) {
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
    const pathHeader = params.endpointPath

    // TODO: match error response with actual TDS spec
    if (!(hmacHeader && pathHeader)) {
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

    const payload = e.postData.contents

    const isValid = validateHmacSignature(signature, senderId, receiverId, timestamp, nonce, 'POST', payload, pathHeader)

    if (!isValid) {
      return createErrorResponse("UNAUTHORIZED")
    }

    let response = ContentService.createTextOutput()
    response.setMimeType(ContentService.MimeType.JSON)
    let content = {}

    if (params.resource === "tripRequestResponses") {
      const processedResponses = receiveTripRequestResponses(JSON.parse(payload))
      content = returnClientOrderConfirmations(processedResponses, validatedApiAccount)
    } else if (params.resource === "providerOrderConfirmations") {
      content = receiveProviderOrderConfirmationsReturnCustomerInformation(payload, validatedApiAccount)
    } else {
      return createErrorResponse("INVALID_REQUEST")
    }
    response.setContent(JSON.stringify(content))
    return response
  } catch(e) { logError(e) }
}

// Function to validate HMAC signature
function validateHmacSignature(signature, senderId, receiverId, timestamp, nonce, method, body, urlEndpoint) {
  try {
    // Fetch the secret associated with the receiverId (apiKey)
    const apiAccounts = getDocProp("apiGiveAccess")
    const statedApiAccount = apiAccounts[receiverId]

    let receivedTimestamp = new Date(timestamp)
    let timeNow = new Date()
    let timePassed = timeNow.getTime() - receivedTimestamp.getTime()

    if (timePassed > 300000) return false

    // Generate the expected HMAC signature based on the retrieved secret and other parameters
    const expectedSignature = generateHmacHexString(
      statedApiAccount.secret,
      senderId,
      receiverId,
      timestamp,
      nonce,
      method,
      body,
      urlEndpoint
    )

    // Compare the expected signature with the received signature
    return signature === expectedSignature
  } catch (e) {
    logError(e)
    return false
  }
}

function createErrorResponse(status) {
  const response = ContentService.createTextOutput()
  response.setMimeType(ContentService.MimeType.JSON)
  const errorContent = { status: status }
  response.setContent(JSON.stringify(errorContent))
  return response
}
