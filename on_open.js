/**
 * @fileoverview Spreadsheet onOpen trigger for RideSheet.
 *
 * Runs each time the spreadsheet is opened (or re-opened after a permission
 * change). Performs the following setup steps in order:
 * 1. Build the RideSheet menu (`buildMenus()`).
 * 2. Run first-open tasks if needed (`runFirstOpenTasks()`).
 * 3. Ensure document properties exist and are up to date
 *    (`buildDocumentPropertiesIfEmpty()`, `buildDocumentPropertiesFromDefaults()`,
 *    `purgeOldDocumentProperties()`).
 * 4. Ensure all named ranges exist (`buildNamedRanges()`).
 * 5. Verify that the spreadsheet and script timezones match the configured
 *    `localTimeZone` property (`checkTimezone()`).
 *
 * Each major step is wrapped in its own try/catch so a failure in one step
 * does not prevent later steps from running.
 */

/**
 * The Google Apps Script onOpen trigger entry point.
 * Logs total execution time on every open.
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e - The onOpen event object.
 */
function onOpen(e) {
  const startTime = new Date()
  try {
    buildMenus()
  } catch(e) { logError(e) }
  try {
    runFirstOpenTasks()
  } catch(e) { logError(e) }
  try {
    buildDocumentPropertiesIfEmpty()
    buildDocumentPropertiesFromDefaults()
    purgeOldDocumentProperties()
  } catch(e) { logError(e) }
  try {
    buildNamedRanges()
  } catch(e) { logError(e) }
  try {
    fixFrozenRows()
  } catch(e) { logError(e) }
  checkTimezone()
  log("onOpen duration:",(new Date()) - startTime)
}

/**
 * Checks that the spreadsheet timezone and script timezone both match the
 * `localTimeZone` document property.
 *
 * - **Spreadsheet timezone**: can be corrected programmatically via
 *   `Spreadsheet.setSpreadsheetTimeZone()`. If out of sync, it is fixed
 *   automatically and a toast notification is shown.
 * - **Script timezone**: is a project-level setting that cannot be changed
 *   by script code. If out of sync, the user is shown an alert with
 *   instructions to correct it manually in the Apps Script project settings.
 */
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