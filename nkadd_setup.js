/**
 * @fileoverview One-time setup for the NKADD regional coordination demo.
 *
 * The core RideSheet build routines (`repairSheets()`, `rebuildAllMetadata()`)
 * repair sheets that already exist — they do not create missing sheets or
 * columns. `setUpNkaddDemo()` fills that gap for the coordination workflow by:
 *
 * 1. Adding the `"Trip Statuses"` and `"Agencies"` lookup columns and seeding
 *    them from `nkaddTripStatuses` / `nkaddAgencies` (constants_local.js).
 * 2. Adding `"Status"`, `"Agency"`, and `"Declined By"` to the Trips sheet and
 *    `"Status"` / `"Agency"` to Trip Archive.
 * 3. Creating the `"Claimed Trips"` sheet with the header row defined in
 *    `localColumns` (constants_local.js).
 * 4. Rebuilding named ranges, developer metadata, and validation so the new
 *    columns behave like every other configured column.
 *
 * The function is idempotent: it only adds what is missing and never
 * overwrites existing header cells or lookup values. It is safe to re-run
 * after editing the column configuration.
 *
 * New columns are appended to the right end of each sheet's header row. Once
 * created they can be dragged anywhere — Google Sheets moves the named ranges
 * and developer metadata along with the column, so the workflow keeps working.
 */

/**
 * Creates and repairs everything the NKADD coordination workflow needs.
 * Exposed on the "NKADD Coordination" menu (see `build_local.js`).
 */
function setUpNkaddDemo() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const ui = safeGetUi()
    let changes = []

    // 1. Lookup columns and their seed values.
    const lookupsSheet = ss.getSheetByName("Lookups")
    if (!lookupsSheet) {
      if (ui) ui.alert("The Lookups sheet is missing. Run RideSheet's normal setup first.")
      return
    }
    changes = changes.concat(nkaddAddMissingColumns(lookupsSheet, ["Trip Statuses", "Agencies"]))
    changes = changes.concat(nkaddSeedLookupColumn(lookupsSheet, "Trip Statuses", nkaddTripStatuses))
    changes = changes.concat(nkaddSeedLookupColumn(lookupsSheet, "Agencies", nkaddAgencies))

    // 2. Coordination columns on the sheets that already exist.
    changes = changes.concat(nkaddAddMissingColumns(ss.getSheetByName("Trips"), ["Status", "Agency", "Declined By"]))
    changes = changes.concat(nkaddAddMissingColumns(ss.getSheetByName("Trip Archive"), ["Status", "Agency"]))

    // 3. The Claimed Trips sheet itself.
    changes = changes.concat(nkaddCreateClaimedSheet())

    // 4. Make the new columns behave like configured columns.
    buildNamedRanges()
    rebuildAllMetadata()
    repairSheets()

    const summary = changes.length
      ? changes.join("\n")
      : "Everything was already in place."
    log("NKADD demo setup", summary)
    if (ui) {
      ui.alert("NKADD Coordination Setup", `Setup complete.\n\n${summary}`, ui.ButtonSet.OK)
    }
  } catch(e) { logError(e) }
}

/**
 * Appends any of `columnNames` that are not already present to the right end of
 * the sheet's header row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to extend.
 * @param {string[]} columnNames - Header names that should exist on the sheet.
 * @returns {string[]} Human-readable descriptions of the columns that were added.
 */
function nkaddAddMissingColumns(sheet, columnNames) {
  if (!sheet) return []
  const changes = []
  columnNames.forEach(columnName => {
    // Re-read the headers each pass so the position accounts for columns added
    // earlier in this loop.
    const headerNames = getSheetHeaderNames(sheet, {forceRefresh: true})
    if (headerNames.includes(columnName)) return
    // `getSheetHeaderNames()` spans every column in the sheet and reports blank
    // headers as " ", so reuse the first empty column before widening the sheet.
    const blankPosition = headerNames.findIndex(headerName => !headerName.toString().trim()) + 1
    const newPosition = blankPosition || headerNames.length + 1
    if (sheet.getMaxColumns() < newPosition) {
      sheet.insertColumnAfter(sheet.getMaxColumns())
    }
    sheet.getRange(1, newPosition).setValue(columnName)
    changes.push(`Added column "${columnName}" to ${sheet.getName()}.`)
  })
  if (changes.length) getSheetHeaderNames(sheet, {forceRefresh: true})
  return changes
}

/**
 * Writes `values` into a Lookups column, but only when that column has no data
 * yet, so hand-edited lookup lists are never clobbered.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} lookupsSheet - The Lookups sheet.
 * @param {string} columnName - The lookup column header name.
 * @param {string[]} values - The values to seed the column with.
 * @returns {string[]} Human-readable descriptions of what was written.
 */
function nkaddSeedLookupColumn(lookupsSheet, columnName, values) {
  const headerNames = getSheetHeaderNames(lookupsSheet, {forceRefresh: true})
  const columnPosition = headerNames.indexOf(columnName) + 1
  if (!columnPosition) return []
  const lastRow = lookupsSheet.getMaxRows()
  const existing = lastRow > 1
    ? lookupsSheet.getRange(2, columnPosition, lastRow - 1, 1).getValues().flat().filter(value => !isBlankCell(value))
    : []
  if (existing.length) return []
  if (lookupsSheet.getMaxRows() < values.length + 1) {
    lookupsSheet.insertRowsAfter(lookupsSheet.getMaxRows(), values.length + 1 - lookupsSheet.getMaxRows())
  }
  lookupsSheet.getRange(2, columnPosition, values.length, 1).setValues(values.map(value => [value]))
  return [`Seeded ${pluralize(values.length, "value")} into Lookups!${columnName}.`]
}

/**
 * Creates the Claimed Trips sheet if it does not exist, using the column order
 * declared for it in `localColumns` (constants_local.js). If the sheet already
 * exists, any configured columns it is missing are appended instead.
 * @returns {string[]} Human-readable descriptions of what was created.
 */
function nkaddCreateClaimedSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const configuredColumns = Object.keys(getConfiguredColumns()[nkaddClaimedSheetName] || {})
  if (!configuredColumns.length) {
    logError(new Error(`No column configuration found for "${nkaddClaimedSheetName}".`))
    return []
  }
  let sheet = ss.getSheetByName(nkaddClaimedSheetName)
  if (sheet) {
    return nkaddAddMissingColumns(sheet, configuredColumns)
  }
  // Place the new sheet immediately after Trips so the two read as a pair.
  const tripsSheet = ss.getSheetByName("Trips")
  const insertIndex = tripsSheet ? tripsSheet.getIndex() : ss.getNumSheets()
  sheet = ss.insertSheet(nkaddClaimedSheetName, insertIndex)
  if (sheet.getMaxColumns() < configuredColumns.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), configuredColumns.length - sheet.getMaxColumns())
  } else if (sheet.getMaxColumns() > configuredColumns.length) {
    sheet.deleteColumns(configuredColumns.length + 1, sheet.getMaxColumns() - configuredColumns.length)
  }
  sheet.getRange(1, 1, 1, configuredColumns.length).setValues([configuredColumns]).setFontWeight("bold")
  sheet.setFrozenRows(1)
  return [`Created the "${nkaddClaimedSheetName}" sheet with ${pluralize(configuredColumns.length, "column")}.`]
}
