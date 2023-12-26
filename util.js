function log(d, ...args) {
  if (getDocProp("logLevel") === "normal") return
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
  logSheet.deleteRows(2,logSheet.getLastRow())
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

function formatDateFromTrip(field, dateFormat) {
  let val = field['@time']
  return formatDate(new Date(val), null, dateFormat)
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

function isInDay(value, date) {
  if (!value || !date) {
    return false
  } else {
    return (value.getTime() >= date.getTime() && value.getTime < (date.getTime() + 1))
  }
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

function replaceText(templateString, data) {
  try {
    //templateString = "This is {a} test {with} words in {braces]"
    let result = templateString
    const pattern = /{(.*?)}/g
    const innerMatches = [...templateString.matchAll(pattern)].map(match => match[1])
    innerMatches.forEach(field => {
      if (isValidDate(data[field])) {
        if (field.match(/\bdate\b/i)) {
          datum = formatDate(data[field])
        } else if (field.match(/\btime\b/i)) {
          datum = formatDate(data[field], null, "h:mm aa")
        } else {
          datum = formatDate(data[field], null, "h:mm aa M/d/yy")
        }
      } else {
        datum = data[field]
      }
      if (Object.keys(data).includes(field)) {
        result = result.replace("{" + field + "}", datum)
      } else {
        result = result.replace("{" + field + "}", field + " not specified")
      }
    })
    return result
  } catch(e) { logError(e) }
}

function getCustomerNameAndId(first, last, id) {
  return `${last}, ${first} (${id})`
}

function convertNamedRangeToTriggerName(namedRange) {
  // remove numeric suffix
  return namedRange.getName().replace(/\d+$/g,'')
}

function loadConfigFile(fileName) {
  try {
    const configFolderId = getDocProp("configFolderId")
    const configFolder = DriveApp.getFolderById(configFolderId)
    const files = configFolder.getFilesByName(fileName)
    if (files.hasNext()) {
      const file = files.next()
      const content = file.getBlob().getDataAsString()
      const json = JSON.parse(content)
      return json
    }
  } catch(e) { logError(e) }
}

function testLoadConfigFile() {
  log(JSON.stringify(loadConfigFile("columns.json.txt")))
}