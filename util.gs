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
  docProps = PropertiesService.getDocumentProperties()
  docProps.getKeys().forEach(prop => {
    log(prop,docProps.getProperty(prop))
  })
}

function makeBackup(destFolderId) {
  let startTime = new Date()
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const newName = formattedDate + ss.getName()
  const destFolder = DriveApp.getFolderById(backupFolderId)
  const file = DriveApp.getFileById(ss.getId())
  
  file.makeCopy(newName, destFolder);
  log("onEdit duration:",(new Date()) - startTime)
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

function iterateTemplate() {
  const doc = DocumentApp.openById("1SPjf_8oVA2BM6wTzSf3Ww521Slet1p6NOorKtDn_HOQ")
  const body = doc.getBody()
  const paragraphs = body.getParagraphs()
  paragraphs.forEach((paragraph) => {
    log(paragraph.getHeading())
  })
}

function isValidDate(date) {
  return date && Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date)
}

function formatDate(date, timeZone, dateFormat) {
  if (!date) date = new Date()
  if (!timeZone) timeZone = localTimeZone
  if (!dateFormat) dateFormat = "M/d/yy"
  if (!isValidDate(date)) date = Date.parse(date) 
  return Utilities.formatDate(date, timeZone, dateFormat)
}

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

function testParseDate() {
  log(dateOnly(dateAdd(new Date(), 1)))
  
}