/**
 * @fileoverview Build and repair utilities for the RideSheet spreadsheet environment.
 *
 * Handles five areas of responsibility:
 *
 * 1. **Menu construction** — builds the RideSheet menu and Settings submenu
 *    (`buildMenus()`), with hooks for local additions via `buildLocalMenus()`.
 *
 * 2. **Named ranges** — creates and repairs named ranges that drive cell triggers
 *    and data validation lookups (`buildNamedRanges()`, `buildNamedRange()`).
 *    The active set of named ranges is the union of `defaultNamedRanges` and
 *    `localNamedRanges`, minus any listed in `localNamedRangesToRemove`.
 *
 * 3. **Document properties** — initializes properties from the Document Properties
 *    sheet (`buildDocumentPropertiesFromSheet()`), adds properties that are missing
 *    because of code updates (`buildDocumentPropertiesFromDefaults()`), and removes
 *    obsolete properties (`purgeOldDocumentProperties()`).
 *
 * 4. **Developer metadata** — stores column configuration (header names, number
 *    formats, data validation rules) as spreadsheet developer metadata so that
 *    column-level settings survive sheet edits and can be used to repair the sheet
 *    (`buildMetadata()`, `fixSheetNames()`, `fixHeaderNames()`,
 *    `fixNumberFormatting()`, `fixDataValidation()`).
 *
 * 5. **Installation wizard** — interactive first-run setup that collects folder
 *    locations, creates the manifest template document, and seeds document
 *    properties (`setupNewInstall()`, `runFirstOpenTasks()`).
 */

/**
 * Builds the RideSheet application menu and the Settings submenu.
 * The "Generate weekly runs from template" item is included only when
 * `createRunMode` is `"default"`. Calls `buildLocalMenus()` at the end so
 * org-specific forks can add their own menu items.
 */
function buildMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('RideSheet')
  menu.addItem('Add return trip', 'createReturnTrip')
  menu.addItem('Add stop', 'addStop')
  menu.addItem('Create manifests for day', 'createManifestsByRunForDate')
  menu.addItem('Create manifests for selected trips', 'createSelectedManifestsByRun')
  menu.addItem('Move past data to review', 'moveTripsToReview')
  menu.addItem('Add data to runs in review','addDataToRunsInReview')
  menu.addItem('Move reviewed data to archive', 'moveTripsToArchive')
  if (getDocProp("createRunMode") === "default") {
    menu.addItem('Generate weekly runs from template', 'buildRunsFromTemplate')
  }
  menu.addSeparator()
  let settingsMenu = ui.createMenu('Settings')
  settingsMenu.addItem('Refresh document properties sheet', 'presentProperties')
  settingsMenu.addItem('Repair sheets', 'repairSheets')
  settingsMenu.addItem('Rebuild metadata', 'rebuildAllMetadata')
  settingsMenu.addItem('Show metadata as column header notes', 'showColumnMetadata')
  settingsMenu.addItem('Clear metadata notes', 'clearHeaderNotes')
  settingsMenu.addItem('Set up new installation', 'setupNewInstall')
  menu.addSubMenu(settingsMenu)
  menu.addToUi()
  buildLocalMenus()
}

/**
 * Creates and repairs all configured named ranges in the spreadsheet.
 *
 * The effective set of named ranges is computed by starting with
 * `defaultNamedRanges`, removing any entries listed in
 * `localNamedRangesToRemove`, and then merging in `localNamedRanges`.
 *
 * For each already-existing named range:
 * - If it is in `localNamedRangesToRemove`, it is deleted.
 * - If it exists in the effective config but its bounds are out of date
 *   (wrong start row or does not extend far enough), it is rebuilt.
 *
 * Any named range in the effective config that does not yet exist is created.
 * Named ranges whose sheet is in `localSheetsToRemove` are skipped.
 */
function buildNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const existingNamedRanges = ss.getNamedRanges()
  const currentRangeNames = existingNamedRanges.map(nr => nr.getName())
  const configuredNamedRanges = (function() {
    const defaultNamedRangesMinusRemoved = Object.fromEntries(
      Object.entries(defaultNamedRanges).filter(([key]) => !localNamedRangesToRemove.includes(key))
    )
    return {...defaultNamedRangesMinusRemoved, ...localNamedRanges}
  })()
  const buildRangeNames = Object.keys(configuredNamedRanges)
  existingNamedRanges.forEach(namedRange => {
    try {
      if (localNamedRangesToRemove.includes(namedRange.getName())) {
        namedRange.remove()
      } else if (buildRangeNames.indexOf(namedRange.getName()) !== -1) {
        const namedRangeName = namedRange.getName()
        const namedRangeConfig = configuredNamedRanges[namedRangeName]
        const startRow = namedRangeConfig.headerName ? 2 : 1
        if (namedRange.getRange().getRow() !== startRow || namedRange.getRange().getLastRow() !== namedRange.getRange().getSheet().getMaxRows() + 1000) {
          buildNamedRange(ss, namedRangeName, namedRangeConfig)
        }
      }
    } catch(e) {
      logError(e)
    }
  })
  buildRangeNames.forEach(rangeName => {
    try {
      const newRangeConfig = configuredNamedRanges[rangeName]
      if (!localSheetsToRemove.includes(newRangeConfig.sheetName)) {
        if (currentRangeNames.indexOf(rangeName) === -1) {
          buildNamedRange(ss, rangeName, newRangeConfig)
        }
      }
    } catch(e) {
      logError(e)
    }
  })
}

/**
 * Creates or updates a single named range in the spreadsheet.
 *
 * The range bounds are determined by the properties of `rangeConfigObj`
 * (exactly one strategy must match):
 * - **`headerName`**: column whose header matches this name; range starts at
 *   row 2 (skipping the header row).
 * - **`column`**: explicit column letter (e.g. `"A"`); range starts at row 1.
 * - **`startHeaderName` + `endHeaderName`**: multi-column span. The range
 *   starts at row 1 if `headerOnly` or `allRows` is set, otherwise row 2.
 *   The range ends at row 1 if `headerOnly` is set.
 *
 * All ranges extend to `sheet.getMaxRows() + 1000` (unless `headerOnly`)
 * to accommodate future rows without requiring a rebuild.
 * Does nothing if the target sheet does not exist.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - The active spreadsheet.
 * @param {string} rangeName - The name to assign to the named range.
 * @param {Object} rangeConfigObj - Configuration for the range.
 * @param {string} rangeConfigObj.sheetName - The sheet the range lives on.
 * @param {string} [rangeConfigObj.headerName] - Header name identifying the column.
 * @param {string} [rangeConfigObj.column] - Explicit column letter.
 * @param {string} [rangeConfigObj.startHeaderName] - Start column header for a multi-column range.
 * @param {string} [rangeConfigObj.endHeaderName] - End column header for a multi-column range.
 * @param {boolean} [rangeConfigObj.headerOnly] - Limit the range to the header row only.
 * @param {boolean} [rangeConfigObj.allRows] - Start the range at row 1 (includes the header).
 */
function buildNamedRange(ss, rangeName, rangeConfigObj) {
  const sheet = ss.getSheetByName(rangeConfigObj.sheetName)
  if (sheet) {
    if (rangeConfigObj.headerName) {
      const headerNames = getSheetHeaderNames(sheet)
      const columnPosition = headerNames.indexOf(rangeConfigObj.headerName) + 1
      if (columnPosition) {
        const columnLetter = getColumnLettersFromPosition(columnPosition)
        const range = sheet.getRange(`${columnLetter}2:${columnLetter}${sheet.getMaxRows() + 1000}`)
        ss.setNamedRange(rangeName, range)
      }
    } else if (rangeConfigObj.column) {
      const range = sheet.getRange(`${rangeConfigObj.column}1:${rangeConfigObj.column}${sheet.getMaxRows() + 1000}`)
      ss.setNamedRange(rangeName, range)
    } else if (rangeConfigObj.startHeaderName && rangeConfigObj.endHeaderName) {
      const headerNames = getSheetHeaderNames(sheet)
      const startColumnPosition = headerNames.indexOf(rangeConfigObj.startHeaderName) + 1
      const endColumnPosition = headerNames.indexOf(rangeConfigObj.endHeaderName) + 1
      if (startColumnPosition && endColumnPosition) {
        const startColumnLetter = getColumnLettersFromPosition(startColumnPosition)
        const endColumnLetter = getColumnLettersFromPosition(endColumnPosition)
        const firstRow = rangeConfigObj.headerOnly || rangeConfigObj.allRows ? 1 : 2
        const lastRow = rangeConfigObj.headerOnly ? 1 : sheet.getMaxRows() + 1000
        const range = sheet.getRange(`${startColumnLetter}${firstRow}:${endColumnLetter}${lastRow}`)
        ss.setNamedRange(rangeName, range)
      }
    }
  }
}

/**
 * Runs tasks that should execute on first open or after a copy is made.
 *
 * Detects a new copy by checking whether any document properties exist. If
 * this is a new copy:
 * - Sets the `showNewInstallMenu` property to `TRUE` so the install menu
 *   persists across subsequent opens.
 * - Shows a "NEW INSTALL" menu entry with a shortcut to `setupNewInstall()`.
 * - Shows a welcome alert directing the user to the installation guide.
 *
 * If `showNewInstallMenu` is `TRUE` (but this is not the first open), shows
 * the new install menu without the welcome alert.
 */
function runFirstOpenTasks() {
  try {
    const ui = safeGetUi()
    if (ui) {
      if (isNewCopy() || getDocProp("showNewInstallMenu")) {
        const menu = ui.createMenu('⭐️NEW INSTALL⭐️')
        menu.addItem('Set up new installation', "setupNewInstall")
        menu.addToUi()
      }
      if (isNewCopy()) {
        const ss = SpreadsheetApp.getActiveSpreadsheet()
        const ui = safeGetUi()
        const propSheet = ss.getSheetByName("Document Properties")
        const propSheetDataRange = propSheet.getDataRange()
        const propSheetData = propSheetDataRange.getValues()
        updatePropertyRange(propSheetData, "showNewInstallMenu", "TRUE")
        propSheetDataRange.setValues(propSheetData)
        buildDocumentPropertiesFromSheet()
        const msg = `
          It looks like you have a fresh copy of RideSheet.\n
          If you would like to set up its environment,
          select "Set up new installation" from the "NEW INSTALL" menu,
          and then grant RideSheet's permission request by clicking
          "Select all" then scrolling down and clicking "Continue".\n
          To learn more about installing RideSheet, visit https://docs.ridesheet.org/technical-guide/installing-ridesheet/
        `
        ui.alert("Welcome to RideSheet!", msg, ui.ButtonSet.OK)
      }
    }
  } catch(e) { logError(e) }
}

/**
 * Returns `true` if no document properties have been set yet, which is the
 * case immediately after a spreadsheet copy is made before any setup has run.
 * @returns {boolean}
 */
function isNewCopy() {
  const propCount = PropertiesService.getDocumentProperties().getProperties()
  return Object.keys(propCount).length === 0
}

/**
 * Interactive installation wizard run from the "NEW INSTALL" menu.
 *
 * Guides the user through two prompts to collect Google Drive folder URLs:
 * 1. The folder where driver manifests will be saved.
 * 2. The folder where the manifest template document will be stored.
 *
 * For each folder, access is verified by creating and immediately trashing a
 * test file. If either step fails or the user cancels, setup is aborted.
 *
 * On success:
 * - Creates the manifest template document from the `manifest_template` HTML
 *   file, moves the page header and footer elements into place, and removes
 *   the placeholder body elements.
 * - Writes `driverManifestFolderId`, `driverManifestTemplateDocId`, and
 *   `showNewInstallMenu` to the Document Properties sheet and reloads
 *   document properties.
 */
function setupNewInstall() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const ui = safeGetUi()

    // Instructions
    const manifestsResponse = ui.prompt(
      'New Install Step 1: Set Folder Where Driver Manifests Will Be Saved',
      'RideSheet needs to know where driver manifests will be saved.\n\n' +
      'We recommend that the folder be named "RideSheet Driver Manifests" and\n' +
      'that it be located in the same folder as RideSheet itself.\n\n' +
      'Create the folder now in another browser window, double-click into it, ' +
      'copy its address from the browser address bar, and enter that address below.',
      ui.ButtonSet.OK_CANCEL
    )
    if (manifestsResponse.getSelectedButton() !== ui.Button.OK) {
        ss.toast("New installation cancelled.")
        return
      }
    const manifestsFolderId = extractFolderId(manifestsResponse.getResponseText())

    // Make sure we can truly put a file in the manifest folder
    try {
      const testDocId = createDoc("Test File", manifestsFolderId, "Just testing", "text/plain")
      Drive.Files.update({ trashed: true }, testDocId, null, { supportsAllDrives: true })
    } catch(e) {
      ui.alert("Error Testing Access to Driver Manifest Folder",
        'Check that the folder location is correct.\n\n' +
        'New installation cancelled.', ui.ButtonSet.OK)
      return
    }

    // Get Settings folder ID
    const settingsResponse = ui.prompt(
      'New Install Step 2: Set Folder Where The Driver Manifest Template Will Be Saved',
      'RideSheet needs to know where to save the driver manifest template.\n\n' +
      'We recommend that the folder be named "RideSheet Settings" and\n' +
      'that it be located in the same folder as RideSheet itself.\n\n' +
      'Create the folder now in another browser window, double-click into it, ' +
      'copy its address from the browser address bar, and enter that address below.',
      ui.ButtonSet.OK_CANCEL
    )
    if (settingsResponse.getSelectedButton() !== ui.Button.OK) {
      ss.toast("Setup cancelled.")
      return
    }
    const settingsFolderId = extractFolderId(settingsResponse.getResponseText())

    // Do the same testing with the settings folder
    try {
      const testDocId = createDoc("Test File", settingsFolderId, "Just testing", "text/plain")
      Drive.Files.update({ trashed: true }, testDocId, null, { supportsAllDrives: true })
    } catch(e) {
      ui.alert("Error Testing Access to Settings Folder",
        'Check that the folder location is correct.\n\n' +
        'New installation cancelled.', ui.ButtonSet.OK)
      return
    }

    // Create the driver manifest template via an import from HTML
    // Imports from HTML cannot set the page header or footer
    const templateSourceHtml = HtmlService.createHtmlOutputFromFile('manifest_template').getContent()
    const templateDocId = createDoc("RideSheet Manifest Template", settingsFolderId, templateSourceHtml, "text/html")

    // Open up the doc and put the page header and footer into place
    prepareTemplate(templateDocId)
    const doc = DocumentApp.openById(templateDocId)
    appendTemplateRange(doc.getNamedRanges("PAGE_HEADER")[0].getRange(), doc.addHeader())
    appendTemplateRange(doc.getNamedRanges("PAGE_FOOTER")[0].getRange(), doc.addFooter())

    // Now delete the body elements that held the page header and footer text
    // This text wouldn't break anything, but it would be confusing to the user
    const rangeNamesToRemove = ["OUTER_PAGE_HEADER","OUTER_PAGE_FOOTER"]
    rangeNamesToRemove.forEach(namedRangeName => {
      const namedRange = doc.getNamedRanges(namedRangeName)[0].getRange()
      const rangeElements = namedRange.getRangeElements()
      rangeElements.forEach(rangeElement => {
        const element = rangeElement.getElement()
        element.removeFromParent()
      })
    })

    const propSheet = ss.getSheetByName("Document Properties")
    const propSheetDataRange = propSheet.getDataRange()
    const propSheetData = propSheetDataRange.getValues()
    updatePropertyRange(propSheetData, "driverManifestFolderId",      manifestsFolderId)
    updatePropertyRange(propSheetData, "driverManifestTemplateDocId", templateDocId)
    updatePropertyRange(propSheetData, "showNewInstallMenu",          "FALSE")
    propSheetDataRange.setValues(propSheetData)
    buildDocumentPropertiesFromSheet()

    ui.alert("Installation Complete",
      "You can now generate driver manifests. Go to the settings folder you entered to view the manifest template and tailor it to your needs.\n\n" +
      "For more details about using RideSheet, visit https://docs.ridesheet.org.", ui.ButtonSet.OK
    )
  } catch(e) {
    safeGetUi()?.alert(e.name + ': ' + e.message)
    logError(e)
  }
}

/**
 * Seeds document properties from the Document Properties sheet if no
 * properties have been set yet (i.e. this is a new copy).
 * Returns `true` if properties were built, `undefined` otherwise.
 * @returns {boolean|undefined}
 */
function buildDocumentPropertiesIfEmpty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const propSheet = ss.getSheetByName("Document Properties")
  if (isNewCopy() && propSheet) {
    buildDocumentPropertiesFromSheet()
    return true
  }
}

/**
 * Reads the Document Properties sheet and writes all recognized property
 * values to script document properties via `setDocProps()`.
 *
 * Only rows whose property name appears in `defaultDocumentProperties` are
 * processed; unrecognized rows are silently ignored. Values are coerced to
 * the type declared in `defaultDocumentProperties` before being stored.
 * Calls `updatePropertiesSheet()` after writing to keep the sheet in sync.
 */
function buildDocumentPropertiesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const propSheet = ss.getSheetByName("Document Properties")
  let propsGrid = propSheet.getDataRange().getValues()
  propsGrid.shift() // Remove the header row from the array
  let defaultPropNames = Object.keys(defaultDocumentProperties)
  let newProps = []
  propsGrid.forEach(row => {
    if (defaultPropNames.indexOf(row[0]) !== -1) {
      let prop = {}
      prop.name = row[0]
      prop.value = coerceValue(row[1], defaultDocumentProperties[row[0]].type)
      prop.description = row[2]
      newProps.push(prop)
    }
  })
  setDocProps(newProps)
  updatePropertiesSheet()
}

/**
 * Adds any default document properties that are missing from the current
 * document properties, and syncs descriptions that are out of date.
 *
 * Called on every open, so new properties introduced by code updates are
 * automatically added without requiring a manual re-import from the sheet.
 *
 * A property is (re)written if:
 * - It is absent from document properties entirely, or
 * - Its stored type differs from the type declared in `defaultDocumentProperties`.
 *
 * A description-only entry is written if the stored description does not
 * match the default description. Only calls `setDocProps()` and
 * `updatePropertiesSheet()` when there is actually something to update.
 */
function buildDocumentPropertiesFromDefaults() {
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  let newProps = []
  Object.keys(defaultDocumentProperties).forEach(propName => {
    if (!docProps[propName] || defaultDocumentProperties[propName].type !== getPropParts(docProps[propName]).type) {
      let prop = {}
      prop.name = propName
      prop.value = coerceValue(defaultDocumentProperties[propName].value, defaultDocumentProperties[propName].type)
      if (defaultDocumentProperties[propName].description !== docProps[propName + propDescSuffix]) prop.description = defaultDocumentProperties[propName].description
      newProps.push(prop)
    } else if (!docProps[propName + propDescSuffix] || defaultDocumentProperties[propName].description !== getPropParts(docProps[propName + propDescSuffix]).value) {
      let prop = {}
      prop.name = propName + propDescSuffix
      prop.value = defaultDocumentProperties[propName].description
      newProps.push(prop)
    }
  })
  if (newProps.length) {
    setDocProps(newProps)
    updatePropertiesSheet()
  }
}

/**
 * Compares the actual column headers on each sheet against the configured
 * columns and returns a diagnostic report.
 *
 * @returns {Object.<string, {defaultPresent: string[], defaultMissing: string[], configPresent: string[], configMissing: string[], notTracked: string[]}>}
 *   An object keyed by sheet name, where each value has:
 *   - `defaultPresent` — configured default columns that exist on the sheet.
 *   - `defaultMissing` — configured default columns absent from the sheet.
 *   - `configPresent` — extra (local) columns that exist on the sheet.
 *   - `configMissing` — extra (local) columns absent from the sheet.
 *   - `notTracked` — columns present on the sheet but not in any config.
 */
function assessMetadata() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const configuredColumns = getConfiguredColumns()
    const configuredSheetsWithHeaders = getConfiguredSheetsWithHeaders()

    let results = {}
    configuredSheetsWithHeaders.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName)
      const sheetHeaderNames = getSheetHeaderNames(sheet)
      const configuredSheetHeaderNames = Object.keys(configuredColumns[sheetName] || {})

      let sheetResults = {}
      sheetResults["defaultPresent"] = configuredSheetHeaderNames.filter(x => sheetHeaderNames.includes(x))
      sheetResults["defaultMissing"] = configuredSheetHeaderNames.filter(x => !sheetHeaderNames.includes(x))
      const sheetHeaderNamesForConfig = sheetHeaderNames.filter(x => !sheetResults["defaultPresent"].includes(x))
      sheetResults["configPresent"] = extraSheetHeaderNames.filter(x => sheetHeaderNamesForConfig.includes(x))
      sheetResults["configMissing"] = extraSheetHeaderNames.filter(x => !sheetHeaderNamesForConfig.includes(x))
      sheetResults["notTracked"] = sheetHeaderNamesForConfig.filter(x => !sheetResults["configPresent"].includes(x))
      results[sheetName] = sheetResults
    })
    return results
  } catch(e) { logError(e) }
}

/**
 * Writes developer metadata to all configured sheets and their columns.
 *
 * For each configured sheet, adds two sheet-level metadata entries:
 * - `"sheetName"` — the canonical sheet name.
 * - `"hasHeader"` — `"true"` or `"false"` depending on whether the sheet
 *   appears in `sheetsWithHeaders`.
 *
 * For each column in a sheet with headers whose name appears in
 * `defaultColumns`, adds column-level metadata for `"headerName"` plus any
 * additional keys defined in the column settings (e.g. `"numberFormat"`,
 * `"dataValidation"`).
 *
 * This metadata is used by `fixSheetNames()`, `fixHeaderNames()`,
 * `fixNumberFormatting()`, and `fixDataValidation()` to repair the sheet.
 */
function buildMetadata() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const configuredColumns = getConfiguredColumns()
    const configuredSheets = getConfiguredSheets()
    const configuredSheetsWithHeaders = getConfiguredSheetsWithHeaders()
    configuredSheets.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName)
      if (sheet) {
        const hasHeader = configuredSheetsWithHeaders.includes(sheetName)
        sheet.addDeveloperMetadata("sheetName",sheetName,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
        sheet.addDeveloperMetadata("hasHeader",JSON.stringify(hasHeader),SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
      } else {
        log(`Attempted to add sheet metadata to sheet '${sheetName}', but sheet not found.`)
      }
    })
    configuredSheetsWithHeaders.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName)
      if (sheet) {
        let sheetHeaderNames = getSheetHeaderNames(sheet)
        let configuredColumnsThisSheet = Object.keys(configuredColumns[sheetName])
        sheetHeaderNames.forEach((columnName, i) => {
          if (configuredColumnsThisSheet.includes(columnName)) {
            let letter = getColumnLettersFromPosition(i + 1)
            let range = sheet.getRange(`${letter}:${letter}`)
            let columnSettings = configuredColumns[sheetName][columnName]
            if (columnSettings) {
              range.addDeveloperMetadata("headerName",JSON.stringify(columnName),SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
              Object.keys(columnSettings).forEach((key) => {
                range.addDeveloperMetadata(key, JSON.stringify(columnSettings[key]), SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
              })
            }
          }
        })
      } else {
        log(`Attempted to add column metadata to sheet '${sheetName}', but sheet not found.`)
      }
    })
  } catch(e) {
    logError(e)
  }
}

/**
 * Removes all developer metadata from the spreadsheet.
 * Typically called before `buildMetadata()` to perform a clean rebuild
 * via `rebuildAllMetadata()`.
 */
function clearMetadata() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let mds = ss.createDeveloperMetadataFinder().find()
    mds.forEach(md => {
      md.remove()
    })
  } catch(e) { logError(e) }
}

/**
 * Clears all developer metadata and rebuilds it from the current column
 * configuration. Exposed as a menu item in the Settings submenu.
 */
function rebuildAllMetadata() {
  try {
    clearMetadata()
    buildMetadata()
  } catch(e) { logError(e) }
}

/**
 * Repairs common sheet issues by restoring sheet names, number formatting,
 * and data validation rules from developer metadata.
 * Exposed as a menu item in the Settings submenu.
 */
function repairSheets() {
  fixSheetNames()
  fixNumberFormatting()
  fixDataValidation()
}

/**
 * Restores any sheet names that have drifted from their canonical values
 * stored in `"sheetName"` developer metadata. Logs each rename.
 */
function fixSheetNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let mds = ss.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET).
      withKey("sheetName").find()
    mds.forEach(md => {
      const sheet = md.getLocation().getSheet()
      if (sheet.getName() !== md.getValue()) {
        log(`Sheet Name '${sheet.getName()}' updated to '${md.getValue()}'`)
        sheet.setName(md.getValue())
      }
    })
  } catch(e) { logError(e) }
}

/**
 * Returns all column-type developer metadata entries within `scope` that
 * match the given key.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet|GoogleAppsScript.Spreadsheet.Sheet} scope
 *   The spreadsheet or sheet to search.
 * @param {string} key - The metadata key to look up (e.g. `"dataValidation"`, `"numberFormat"`).
 * @returns {GoogleAppsScript.Spreadsheet.DeveloperMetadata[]}
 */
function getColumnMetadata(scope, key) {
  let mds = scope.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .withKey(key)
    .find()
  return mds
}

/**
 * Applies column data-validation rules from developer metadata to every
 * column in a single row. Used when a new row is inserted to ensure it
 * immediately has the correct validation.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - A single-row range.
 */
function fixRowDataValidation(range) {
  let sheet = range.getSheet()
  let mds = getColumnMetadata(sheet, 'dataValidation')
  mds.forEach(md => {
    let col = md.getLocation().getColumn().getColumn()
    let row = range.getRow()
    let cell = sheet.getRange(row, col, 1, 1)
    const ruleAttributes = JSON.parse(md.getValue())
    let rule = getValidationRule(ruleAttributes)
    cell.setDataValidation(rule)
  })
}

/**
 * Builds a `DataValidationRule` from a JSON attribute object stored in
 * column developer metadata. Supports three criteria strategies:
 * - **`VALUE_IN_RANGE`** — looks up a named range by `ruleAttributes.namedRange`
 *   and trims it to actual sheet bounds before creating the rule.
 * - **`VALUE_IN_LIST`** — uses an explicit list of values.
 * - **`CHECKBOX`** — supports custom checked/unchecked values or plain checkbox.
 * - **Other** — passes `ruleAttributes.args` directly as criteria arguments.
 *
 * @param {Object} ruleAttributes - Parsed JSON from column developer metadata.
 * @param {string} ruleAttributes.criteriaType - A `DataValidationCriteria` key name.
 * @param {boolean} [ruleAttributes.allowInvalid] - Allow values that fail validation.
 * @param {string} [ruleAttributes.namedRange] - Named range for `VALUE_IN_RANGE`.
 * @param {boolean} [ruleAttributes.showDropdown] - Show dropdown for list/range criteria.
 * @param {Array} [ruleAttributes.values] - Values for `VALUE_IN_LIST`.
 * @param {*} [ruleAttributes.checkedValue] - Checked value for `CHECKBOX`.
 * @param {*} [ruleAttributes.uncheckedValue] - Unchecked value for `CHECKBOX`.
 * @param {Array} [ruleAttributes.args] - Arguments for other criteria types.
 * @param {string} [ruleAttributes.helpText] - Help text shown on validation failure.
 * @returns {GoogleAppsScript.Spreadsheet.DataValidation|undefined}
 */
function getValidationRule(ruleAttributes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const criteriaName = ruleAttributes.criteriaType
    const criteria = SpreadsheetApp.DataValidationCriteria[criteriaName]
    const allowInvalid = !!ruleAttributes.allowInvalid
    let args = []
    let builder
    if (criteriaName === "VALUE_IN_RANGE") {
      // Named ranges can extend past the actual number of rows, but
      // the ranges used for data validation cannot, so we're building a new range.
      const lookupRange = ss.getRangeByName(ruleAttributes.namedRange)
      const inBoundsRange = getInBoundsRange(lookupRange)
      const dropdown = ruleAttributes.showDropdown
      args = [inBoundsRange, dropdown]
      builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
    } else if (criteriaName === "VALUE_IN_LIST") {
      const dropdown = ruleAttributes.showDropdown
      const values = ruleAttributes.values
      args = [values, dropdown]
      builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
    } else if (criteriaName === "CHECKBOX") {
      if (Object.hasOwn(ruleAttributes,"checkedValue")) {
        if (Object.hasOwn(ruleAttributes,"uncheckedValue")) {
          builder = SpreadsheetApp.newDataValidation().requireCheckbox(ruleAttributes.checkedValue, ruleAttributes.uncheckedValue).setAllowInvalid(allowInvalid)
        } else {
          builder = SpreadsheetApp.newDataValidation().requireCheckbox(ruleAttributes.checkedValue).setAllowInvalid(allowInvalid)
        }
      } else {
        builder = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(allowInvalid)
      }
    } else {
      args = ruleAttributes.args || args
      builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
    }
    if (builder) {
      if (Object.hasOwn(ruleAttributes,"helpText")) {
        builder = builder.setHelpText(ruleAttributes.helpText)
      }
      let rule = builder.build()
      return rule
    }
  } catch(e) { logError(e) }
}

/**
 * Applies data-validation rules from developer metadata to all data rows
 * (skipping row 1) in the given sheet, or across the entire spreadsheet
 * if no sheet is specified.
 * @param {GoogleAppsScript.Spreadsheet.Sheet|string|null} [sheet=null] - A sheet
 *   object, a sheet name, or `null` to apply across all sheets.
 */
function fixDataValidation(sheet=null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let scope = ss
  if (sheet) {
    if (typeof sheet === "object") {
      scope = sheet
    } else {
      scope = ss.getSheetByName(sheet)
    }
  }
  let mds = getColumnMetadata(scope, 'dataValidation')
  mds.forEach(md => {
    let fullCol = md.getLocation().getColumn()
    let numRows = fullCol.getHeight()
    let col = fullCol.offset(1, 0, numRows - 1)
    const ruleAttributes = JSON.parse(md.getValue())
    let rule = getValidationRule(ruleAttributes)
    col.setDataValidation(rule)
  })
}

/**
 * Applies number formats from developer metadata to all data rows
 * (skipping row 1) in the given sheet, or across the entire spreadsheet
 * if no sheet is specified.
 * @param {GoogleAppsScript.Spreadsheet.Sheet|string|null} [sheet=null] - A sheet
 *   object, a sheet name, or `null` to apply across all sheets.
 */
function fixNumberFormatting(sheet=null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let scope = ss
  if (sheet) {
    if (typeof sheet === "object") {
      scope = sheet
    } else {
      scope = ss.getSheetByName(sheet)
    }
  }
  let mds = getColumnMetadata(scope, 'numberFormat')
  mds.forEach(md => {
    let fullCol = md.getLocation().getColumn()
    let numRows = fullCol.getHeight()
    let col = fullCol.offset(1, 0, numRows - 1)
    let format = JSON.parse(md.getValue())
    col.setNumberFormat(format)
  })
}

/**
 * Applies column number formats from developer metadata to every column
 * in a single row. Used when a new row is inserted.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - A single-row range.
 */
function fixRowNumberFormatting(range) {
    let sheet = range.getSheet()
    let mds = getColumnMetadata(sheet, 'numberFormat')
    mds.forEach(md => {
      let col = md.getLocation().getColumn().getColumn()
      let row = range.getRow()
      let cell = sheet.getRange(row, col, 1, 1)
      let format = md.getValue()
      cell.setNumberFormat(format)
    })
}

/**
 * Restores header row values that have drifted from their canonical values
 * stored in column developer metadata.
 *
 * For each column in `rangeIn`, looks up the `"headerFormula"` or
 * `"headerName"` metadata key. If the current cell value or formula does
 * not match the stored value, it is overwritten. All out-of-sync columns
 * are corrected in a single `setValues()` call.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} rangeIn - The header row range
 *   (or a portion of it) to inspect and repair.
 */
function fixHeaderNames(rangeIn) {
  try {
    const headerMetadata = rangeIn.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN).
      onIntersectingLocations().find()
    const metadataByColumn = headerMetadata.reduce((metadataAcc,metadataItem) => {
      const column = metadataItem.getLocation().getColumn()
      const columnPosition = column.getColumn()
      metadataAcc[columnPosition] = metadataAcc[columnPosition] || {}
      metadataAcc[columnPosition][metadataItem.getKey()] = JSON.parse(metadataItem.getValue())
      return metadataAcc
    },{})
    const rangeValues = rangeIn.getValues()
    const rangeFormulas = rangeIn.getFormulas()
    const rangeStartColumnPosition = rangeIn.getColumn()
    const currentHeaderFormulasOrValues = rangeFormulas[0].reduce((acc, formula, index) => {
      acc[index + rangeStartColumnPosition] = (formula || rangeValues[0][index])
      return acc
    },{})
    let columnsPositionsToFix = []
    let intendedHeaderNames = {}
    Object.keys(currentHeaderFormulasOrValues).forEach((columnPosition) => {
      let examineColumn = true
      if (Object.hasOwn(metadataByColumn[columnPosition] || {},"headerFormula")) {
        intendedHeaderNames[columnPosition] = metadataByColumn[columnPosition].headerFormula
      } else if (Object.hasOwn(metadataByColumn[columnPosition] || {},"headerName")) {
        intendedHeaderNames[columnPosition] = metadataByColumn[columnPosition].headerName
      } else {
        examineColumn = false
      }
      if (
        examineColumn &&
        currentHeaderFormulasOrValues[columnPosition] !== intendedHeaderNames[columnPosition]
      ) {
        columnsPositionsToFix.push(columnPosition)
      }
    })
    if (columnsPositionsToFix.length) {
      let firstColPos = Math.min(...columnsPositionsToFix)
      let lastColPos = Math.max(...columnsPositionsToFix)
      let newRange = rangeIn.getSheet().getRange(1, firstColPos, 1, lastColPos - firstColPos + 1)
      let values = [[]]
      for (let i = firstColPos; i <= lastColPos; i++) values[0].push(intendedHeaderNames[i])
      newRange.setValues(values)
    }
  } catch(e) { logError(e) }
}

/**
 * Runs `fixHeaderNames()` for all sheets that have developer metadata
 * indicating they have a header row (`hasHeader = true`).
 */
function fixAllHeaderNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const mds = ss.createDeveloperMetadataFinder()
      .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET)
      .withKey("hasHeader")
      .withValue(JSON.stringify(true))
      .find()
  mds.forEach((md) => {
    const sheet = md.getLocation().getSheet()
    const range = getFullRow(sheet.getRange("A1"))
    fixHeaderNames(range)
  })
}

/**
 * Logs all column-level `"headerName"` metadata entries to the script log.
 * Useful for debugging metadata configuration.
 * @private
 */
function logMetadata() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let mds = ss.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN).
      withKey("headerName").find()
    mds.forEach(md => {
      const sheet = md.getLocation().getSheet()
      log(md.getKey(), md.getValue(), md.getLocation().getColumn().getSheet().getName() + "!" + md.getLocation().getColumn().getA1Notation())
    })
  } catch(e) { logError(e) }
}

/**
 * Updates a property value in an in-memory grid of Document Properties sheet
 * data before it is written back to the sheet. Mutates `dataRange` in place.
 * @param {Array<Array<*>>} dataRange - The full sheet data as a 2D array
 *   (rows × columns), where column 0 is the property name and column 1 is
 *   the value.
 * @param {string} propName - The property name to find and update.
 * @param {*} newPropValue - The new value to set.
 */
function updatePropertyRange(dataRange, propName, newPropValue) {
  dataRange.forEach(row => {
    if (row[0] === propName) { row[1] = newPropValue }
  })
}

/**
 * Writes all column-level developer metadata as notes on the header row cells
 * of each sheet. Existing header notes are cleared first.
 * Exposed as a menu item in the Settings submenu. Useful for inspecting
 * column configuration without using the Apps Script editor.
 */
function showColumnMetadata() {
  try {
    clearHeaderNotes()
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let mds = ss.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN).find()
    let metadata = {}
    mds.forEach(md => {
      const range = md.getLocation().getColumn()
      const sheetName = range.getSheet().getName()
      const column = range.getColumn()
      if (!metadata.hasOwnProperty(sheetName)) metadata[sheetName] = {}
      if (metadata[sheetName].hasOwnProperty(column)) {
        metadata[sheetName][column] =  metadata[sheetName][column] + "\n" + md.getKey() + ": " + md.getValue()
      } else {
        metadata[sheetName][column] =  md.getKey() + ": " + md.getValue()
      }
    })
    Object.keys(metadata).forEach ((sheetName) => {
      const lastColumnNumber = Math.max(...Object.keys(metadata[sheetName]))
      let headerNotes = new Array(lastColumnNumber - 1)
      for (let i = 0; i < lastColumnNumber; i++) {
        if (metadata[sheetName].hasOwnProperty(i + 1)) {
          headerNotes[i] = metadata[sheetName][i + 1]
        } else {
          headerNotes[i] = ""
        }
      }
      const sheet = ss.getSheetByName(sheetName)
      const range = sheet.getRange(1,1,1,lastColumnNumber)
      range.setNotes([headerNotes])
    })
  } catch(e) { logError(e) }
}

/**
 * Clears the notes from the header row of every configured sheet that has
 * headers. Used to undo the notes added by `showColumnMetadata()`.
 * Exposed as a menu item in the Settings submenu.
 */
function clearHeaderNotes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    getConfiguredSheetsWithHeaders().forEach((sheetName) => {
      const sheet = ss.getSheetByName(sheetName)
      if (sheet) {
        let range = sheet.getRange(1, 1, 1, sheet.getLastColumn())
        let notes = []
        for (let i = 0; i < range.getNumColumns(); i++) { notes.push("") }
        range.setNotes([notes])
      }
    })
  } catch(e) { logError(e) }
}

/**
 * Returns the effective column configuration by merging `defaultColumns` with
 * `localColumns`, after removing any sheets listed in `localSheetsToRemove`
 * and any columns listed in `localColumnsToRemove`.
 * @returns {Object.<string, Object.<string, Object>>} An object keyed by sheet
 *   name, whose values are objects keyed by column header name.
 */
function getConfiguredColumns() {
  // Initial configuration based on defaultColumns, excluding removed sheets and columns
  const baseConfig = Object.keys(defaultColumns).reduce((sheetAcc, sheetName) => {
    if (!localSheetsToRemove.includes(sheetName)) {
      const columns = Object.keys(defaultColumns[sheetName]).reduce((columnAcc, columnName) => {
        if (!(localColumnsToRemove[sheetName] || []).includes(columnName)) {
          columnAcc[columnName] = defaultColumns[sheetName][columnName]
        }
        return columnAcc
      }, {})
      sheetAcc[sheetName] = columns
    }
    return sheetAcc
  }, {})

  // Add or update from localColumns
  return Object.keys(localColumns).reduce((sheetAcc, sheetName) => {
    sheetAcc[sheetName] = { ...(sheetAcc[sheetName] || {}), ...localColumns[sheetName] }
    return sheetAcc
  }, baseConfig)
}

/**
 * Returns the effective list of sheet names, combining `defaultSheets`
 * (minus any in `localSheetsToRemove`) with `localSheets`.
 * @returns {string[]}
 */
function getConfiguredSheets() {
  return [...defaultSheets.filter((sheetName) => !localSheetsToRemove.includes(sheetName)),...localSheets]
}

/**
 * Returns the effective list of sheet names that have header rows, combining
 * `sheetsWithHeaders` (minus any in `localSheetsToRemove`) with
 * `localSheetsWithHeaders`.
 * @returns {string[]}
 */
function getConfiguredSheetsWithHeaders() {
  return [...sheetsWithHeaders.filter((sheetName) => !localSheetsToRemove.includes(sheetName)),...localSheetsWithHeaders]
}

/**
 * Returns a copy of `range` trimmed to the actual row bounds of its sheet,
 * so that over-extended named ranges (which go beyond `getMaxRows()`) can
 * be used safely where in-bounds ranges are required (e.g. data validation).
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to trim.
 * @returns {GoogleAppsScript.Spreadsheet.Range}
 */
function getInBoundsRange(range) {
  const sheet = range.getSheet()
  const sheetLastRow = sheet.getMaxRows()
  const newRowCount = range.getLastRow() > sheetLastRow ? sheetLastRow - range.getRow() + 1 : range.getNumRows()
  return sheet.getRange(range.getRow(),range.getColumn(), newRowCount,range.getNumColumns())
}

/**
 * Extracts a Google Drive folder ID from a full Drive URL or returns the
 * input as-is if it appears to already be a bare folder ID.
 * Throws if `input` is empty.
 * @param {string} input - A Drive folder URL or bare folder ID.
 * @returns {string} The extracted folder ID.
 */
function extractFolderId(input) {
  if (!input) {
    throw new Error('No folder ID provided')
  }
  input = input.trim()
  // Try to extract from full URL
  // https://drive.google.com/drive/folders/FOLDER_ID
  const urlMatch = input.match(/\/folders\/([a-zA-Z0-9_-]+)/)
  if (urlMatch) {
    return urlMatch[1]
  }
  // Assume it's already just the ID
  return input
}