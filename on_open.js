function onOpen(e) {
  const startTime = new Date()
  try {
    buildMenus()
  } catch(e) { logError(e) }
  try {
    buildDocumentPropertiesIfEmpty()
    buildDocumentPropertiesFromDefaults()
    purgeOldDocumentProperties()
  } catch(e) { logError(e) }
  try {
    buildNamedRanges()
  } catch(e) { logError(e) }
  checkTimezone()
  log("onOpen duration:",(new Date()) - startTime)
}

function checkTimezone() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const scriptTimeZone = Session.getScriptTimeZone();
  const ssTimeZone = ss.getSpreadsheetTimeZone();
  const propTimeZone = getDocProp("localTimeZone");

  // Spreadsheet timezone check
  if (ssTimeZone !== propTimeZone) {
    ss.setSpreadsheetTimeZone(propTimeZone);
    SpreadsheetApp.getActive().toast(
      `Spreadsheet timezone was out of sync. Updated to: ${propTimeZone}`,
      "Timezone Updated",
      5
    );
    log("Updated sheet timezone", `Spreadsheet timezone updated to '${propTimeZone}'`);
  }

  // Script timezone check (manual intervention required)
  if (scriptTimeZone !== propTimeZone) {
    ui.alert(
      "Script Timezone Mismatch",
      `The script timezone ('${scriptTimeZone}') does not match the expected timezone ('${propTimeZone}').\n\n` +
      `Please update it manually via the Apps Script Editor (Project Settings).\n` +
      `If you're unsure how to do this, please contact support.`,
      ui.ButtonSet.OK
    );
    log("Warning: Timezone mismatch", `Script timezone '${scriptTimeZone}' does not match '${propTimeZone}'. Manual update required.`);
  }
}