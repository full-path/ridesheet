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

function isNewCopy() {
  const propCount = PropertiesService.getDocumentProperties().getProperties()
  return Object.keys(propCount).length === 0
}

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
    templateDocId = createDoc("RideSheet Manifest Template", settingsFolderId, templateSourceHtml, "text/html")

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

function buildDocumentPropertiesIfEmpty() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const propSheet = ss.getSheetByName("Document Properties")
  if (isNewCopy() && propSheet) {
    buildDocumentPropertiesFromSheet()
    return true
  }
}

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

// If there are any default document properties that are missing from the actual
// document properties for this sheet, this adds them.
// This is useful for code updates that involve creating new document properties.
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

function clearMetadata() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let mds = ss.createDeveloperMetadataFinder().find()
    mds.forEach(md => {
      md.remove()
    })
  } catch(e) { logError(e) }
}

function rebuildAllMetadata() {
  try {
    clearMetadata()
    buildMetadata()
  } catch(e) { logError(e) }
}

function repairSheets() {
  fixSheetNames()
  fixNumberFormatting()
  fixDataValidation()
}

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

function getColumnMetadata(scope, key) {
  let mds = scope.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .withKey(key)
    .find()
  return mds
}

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

function updatePropertyRange(dataRange, propName, newPropValue) {
  dataRange.forEach(row => {
    if (row[0] === propName) { row[1] = newPropValue }
  })
}

// Take spreadsheet's column-type metadata and put in a note in the top
// cell of the associated column. Useful for testing configurations.
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

// Clears out the notes fields of the top row of sheets.
// Useful for clearing out the notes put in place by showColumnMetadata()
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

function getConfiguredSheets() {
  return [...defaultSheets.filter((sheetName) => !localSheetsToRemove.includes(sheetName)),...localSheets]
}

function getConfiguredSheetsWithHeaders() {
  return [...sheetsWithHeaders.filter((sheetName) => !localSheetsToRemove.includes(sheetName)),...localSheetsWithHeaders]
}

function getInBoundsRange(range) {
  const sheet = range.getSheet()
  const sheetLastRow = sheet.getMaxRows()
  const newRowCount = range.getLastRow() > sheetLastRow ? sheetLastRow - range.getRow() + 1 : range.getNumRows()
  return sheet.getRange(range.getRow(),range.getColumn(), newRowCount,range.getNumColumns())
}

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