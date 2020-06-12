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
  const formattedDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm:ss ")
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