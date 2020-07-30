function log(d, ...args) {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let logSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log")
  logSheet.appendRow([new Date(), d].concat(args))
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
  if (!timeZone) timeZone = getDocProp("localTimeZone",defaultLocalTimeZone)
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

function testFormatDate() {log(formatDate("2020-5-24"))}

function dateAdd(date, days) {
  if (!date) date = new Date()
  let result = new Date(date)
  result.setDate(result.getDate() + days)
  return result
}

function dateOnly(dateTime) {
  if (!dateOnly) dateTime = new Date()
  return new Date(dateTime.setHours(0,0,0,0))
}

function parseDate(date, alternateValue) {
  dateVal = Date.parse(date.toString())
  return isNaN(dateVal) ? alternateValue : new Date(dateVal)
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