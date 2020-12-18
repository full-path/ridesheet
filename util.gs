function log(d, ...args) {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let logSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log")
  logSheet.appendRow([new Date(), d].concat(args))
}

function logError(e) {
  log(e.name + ': ' + e.message, e.stack)
}

function clearLog() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let logSheet = ss.getSheetByName("Debug Log")
  logSheet.deleteRows(1,logSheet.getLastRow())
}

function logProperties() {
  let docProps = PropertiesService.getDocumentProperties()
  docProps.getKeys().forEach(prop => {
    log(prop,docProps.getProperty(prop))
  })
}

function makeBackup(destFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const file = DriveApp.getFileById(ss.getId())
    const newName = formatDate(null,null,"yyyy-MM-dd") + "_" + ss.getName()
    const destFolder = DriveApp.getFolderById(destFolderId)
    file.makeCopy(newName, destFolder)
    return true
  } catch(e) {
    return false
  }
}

function rotateBackups(folderId, retentionInDays) {
  const files = DriveApp.getFolderById(folderId).getFiles()
  let file
  let ageInDays
  for (const files = DriveApp.getFolderById(folderId).getFiles(); files.hasNext(); file = files.next()) {
    ageInDays = Math.ceil(((new Date()) - file.getDateCreated()) / (1000 * 60 * 60 * 24))
    if (ageInDays > retentionInDays) { file.setTrashed(true) }
  }
}

function isValidDate(date) {
  return date && Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date)
}

function formatDate(date, timeZone, dateFormat) {
  if (!timeZone) timeZone = getDocProp("localTimeZone")
  if (!dateFormat) dateFormat = "M/d/yyyy"
  if (!date) {
    return Utilities.formatDate(new Date(), timeZone, dateFormat)
  } else if (isValidDate(date)) {
    return Utilities.formatDate(date, timeZone, dateFormat)
  } else {
    const thisDate = new Date(date)
    if (isValidDate(thisDate)) return Utilities.formatDate(thisDate, timeZone, dateFormat)
  }
}

function dateAdd(date, days) {
  if (!date) date = new Date()
  let result = new Date(date)
  result.setDate(result.getDate() + days)
  return result
}

function timeAdd(date, milliseconds) {
  if (!date) date = new Date()
  return new Date(date.getTime() + milliseconds)
}

function dateOnly(dateTime) {
  try {
    let thisDateTime
    if (!dateTime) {
      thisDateTime = new Date()
    } else if (typeof dateTime === "string") {
      thisDateTime = new Date(dateTime)
    } else {
      thisDateTime = dateTime
    } 
    return new Date(thisDateTime.setHours(0,0,0,0))
  } catch(e) { logError(e) }
}

function dateToday() {
  try {
    return dateOnly()
  } catch(e) { logError(e) }
}

function parseDate(date, alternateValue) {
  dateVal = Date.parse(date.toString())
  return isNaN(dateVal) ? alternateValue : new Date(dateVal)
}

function timeOnlyAsMilliseconds(dateTime) {
  try {
    let thisDateTime
    if (!dateTime) {
      thisDateTime = new Date()
    } else if (typeof dateTime === "string") {
      thisDateTime = new Date(dateTime)
    } else {
      thisDateTime = dateTime
    }
    return thisDateTime.getHours() * 3600000 + thisDateTime.getMinutes() * 60000 + thisDateTime.getSeconds() * 1000 + thisDateTime.getMilliseconds()
  } catch(e) { logError(e) }
}

function combineDateAndTime(date, time) {
  try {
    return new Date(dateOnly(date).getTime() + timeOnlyAsMilliseconds(time))
  } catch(e) { logError(e) }
}

function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

function getType(value) {
  let objectClass = Object.prototype.toString.call(value)
  let classes = {
    "[object Array]":      "array",
    "[object BigInt]":     "bigint",
    "[object Boolean]":    "boolean",
    "[object Date]":       "date",
    "[object Map]":        "map",
    "[object Null]":       "null",
    "[object Number]":     "number",
    "[object Object]":     "object",
    "[object RegExp]":     "regexp",
    "[object Set]":        "set",
    "[object String]":     "string",
    "[object Symbol]":     "symbol",
    "[object Undefined]":  "undefined"
  }
  if (objectClass in classes) {
    return classes[objectClass]
  } else {
    return "string"
  }
}

function getCustomerNameAndId(first, last, id) {
  return `${last}, ${first} (${id})`
}

function convertNamedRangeToTriggerName(namedRange) {
  // remove numeric suffix
  return namedRange.getName().replace(/\d+$/g,'')
}

function urlQueryString(params) {
  try {
    const keys = Object.keys(params)
    result = keys.reduce((a, key, i) => {
      if (i === 0) {
        return key + "=" + params[key]
      } else {
        return a + "&" + key + "=" + params[key]
      }
    },"")
    return result
  } catch(e) { logError(e) }
}

function byteArrayToHexString(bytes) {
  try {
    return bytes.map(byte => {
      return ("0" + (byte & 0xFF).toString(16)).slice(-2)
    }).join('')
  } catch(e) { logError(e) }
}

function generateHmacHexString(secret, nonce, timestamp, params) {
  try {
    let orderedParams = {}
    Object.keys(params).sort().forEach(key => {
      orderedParams[key] = params[key]
    })
    const value = [nonce, timestamp, JSON.stringify(orderedParams)].join(':')
    const sigAsByteArray = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_256, value, secret)
    return byteArrayToHexString(sigAsByteArray)
  } catch(e) { logError(e) }
}

function getResource(endPoint, params) {
  try {
    params.version = endPoint.version
    params.apiKey = endPoint.apiKey
    let hmac = {}
    hmac.nonce = Utilities.getUuid()
    hmac.timestamp = JSON.parse(JSON.stringify(new Date()))
    hmac.signature = generateHmacHexString(endPoint.secret, hmac.nonce, hmac.timestamp, params)
    
    const options = {method: 'GET', contentType: 'application/json'}
    const response = UrlFetchApp.fetch(endPoint.url + "?" + urlQueryString(params) + "&" + urlQueryString(hmac), options)
    return response
  } catch(e) { logError(e) }
}