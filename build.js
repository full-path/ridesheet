function buildMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('RideSheet')
  menu.addItem('Refresh driver calendars', 'updateDriverCalendars')
  menu.addItem('Add return trip', 'createReturnTrip')
  menu.addItem('Add stop', 'addStop')
  menu.addItem('Create manifests for day', 'createManifests')
  menu.addItem('Create manifests for selected items', 'createSelectedManifests')
  menu.addItem('Move past data to review', 'moveTripsToReview')
  menu.addItem('Add data to runs in review','addDataToRunsInReview')
  menu.addItem('Move reviewed data to archive', 'moveTripsToArchive')
  menu.addSeparator()
  let settingsMenu = ui.createMenu('Settings')
  settingsMenu.addItem('Refresh document properties sheet', 'presentProperties')
  settingsMenu.addItem('Scheduled calendar updates', 'presentCalendarTrigger')
  settingsMenu.addItem('Repair sheets', 'repairSheets')
  settingsMenu.addItem('Build Metadata', 'buildMetadata')
  menu.addSubMenu(settingsMenu)
  if (getDocProp("apiShowMenuItems")) {
    const menuApi = ui.createMenu('Ride Sharing')
    menuApi.addItem('Get trip requests', 'sendRequestForTripRequests')
    menuApi.addItem('Send responses to trip requests', 'sendTripRequestResponses')
    menuApi.addItem('Refresh outside runs', 'sendRequestForRuns')
    menuApi.addToUi()
  }
  menu.addToUi()
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
      } else if (buildRangeNames.indexOf(namedRange.getName()) !== -1 &&
          (namedRange.getRange().getRow() !== 1 || 
          namedRange.getRange().getLastRow() !== namedRange.getRange().getSheet().getMaxRows() + 1000)) {
        const name = namedRange.getName()
        let newRange = configuredNamedRanges[name]
        buildNamedRange(ss, name, newRange.sheetName, newRange.column, newRange.headerName)
      }
    } catch(e) {
      logError(e)
    }
  })
  buildRangeNames.forEach(rangeName => {
    try {
      const newRange = configuredNamedRanges[rangeName]
      if (!localSheetsToRemove.includes(newRange.sheetName)) {
        if (currentRangeNames.indexOf(rangeName) === -1) {
          buildNamedRange(ss, rangeName, newRange.sheetName, newRange.column, newRange.headerName)
        }
      }
    } catch(e) {
      logError(e)
    }
  })
}

function buildNamedRange(ss, name, sheetName, column, headerName) {
  const sheet = ss.getSheetByName(sheetName)
  if (sheet) {
    if (headerName) {
      const headerNames = getSheetHeaderNames(sheet)
      const columnPosition = headerNames.indexOf(headerName) + 1
      if (columnPosition > 0) column = sheet.getRange(1, columnPosition).getA1Notation().slice(0,-1)
    }
    if (column) {
      const range = sheet.getRange(column + "1:" + column + (sheet.getMaxRows() + 1000))
      ss.setNamedRange(name, range)
    }
  }
}

// Document properties don't pass on to copied sheets.
// This recreates the ones put into the properties sheet.
function buildDocumentPropertiesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const propSheet = ss.getSheetByName("Document Properties")
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  if (Object.keys(docProps).length === 0 && propSheet) {
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
              range.addDeveloperMetadata("headerName",columnName,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
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
    let ss = SpreadsheetApp.getActiveSpreadsheet()
    let criteriaName = ruleAttributes.criteriaType
    let criteria = SpreadsheetApp.DataValidationCriteria[criteriaName]
    let allowInvalid = !!ruleAttributes.allowInvalid
    let builder
    let args = []
    if (criteriaName === "VALUE_IN_RANGE") {
      let rng = ss.getRangeByName(ruleAttributes.namedRange)
      let a1 = rng.getA1Notation()
      let lookupsheet = rng.getSheet()
      let colLetter = a1.substring(0,1)
      let simplifiedRange = lookupsheet.getRange(colLetter + ':' + colLetter)
      let dropdown = ruleAttributes.showDropdown
      args = [simplifiedRange, dropdown]
      builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
    } else if (criteriaName === "VALUE_IN_LIST") {
      let dropdown = ruleAttributes.showDropdown
      let values = ruleAttributes.values
      args = [values, dropdown]
      builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
    } else if (criteriaName === "CHECKBOX") {
      if (ruleAttributes.hasOwnProperty("checkedValue")) {
        if (ruleAttributes.hasOwnProperty("uncheckedValue")) {
          builder = SpreadsheetApp.newDataValidation().requireCheckbox(ruleAttributes.checkedValue, ruleAttributes.uncheckedValue).setAllowInvalid(allowInvalid)
        } else {
          builder = SpreadsheetApp.newDataValidation().requireCheckbox(ruleAttributes.checkedValue).setAllowInvalid(allowInvalid)
        }
      } else {
        builder = SpreadsheetApp.newDataValidation().requireCheckbox().setAllowInvalid(allowInvalid)
      }
    } else if (criteriaName === "TEXT_IS_VALID_EMAIL") {
      builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
    } else if (criteriaName === "DATE_IS_VALID_DATE") {
      builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
    }
    if (builder !== undefined) {
      if (ruleAttributes.hasOwnProperty("helpText")) {
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

function fixHeaderNames(range) {
  try {
    const mds = range.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN).
      onIntersectingLocations().
      withKey("headerName").
      find()
    let columnsPositionsToFix = []
    let headerNames = {}
    mds.forEach(md => {
      const column = md.getLocation().getColumn()
      const columnIndex = column.getColumn()
      const actualHeaderValue = column.getValue()
      const intendedHeaderValue = md.getValue()

      headerNames[columnIndex] = intendedHeaderValue;
      if (actualHeaderValue !== intendedHeaderValue) {
        columnsPositionsToFix.push(columnIndex);
      }
    })
    if (columnsPositionsToFix.length) {
      let firstColPos = Math.min(...columnsPositionsToFix)
      let lastColPos = Math.max(...columnsPositionsToFix)
      let newRange = range.getSheet().getRange(1, firstColPos, 1, lastColPos - firstColPos + 1)
      let values = [[]]
      for (let i = firstColPos; i <= lastColPos; i++) values[0].push(headerNames[i])
      newRange.setValues(values)
    }
  } catch(e) { logError(e) }
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

// Sets up a new instance of RideSheet
function buildRideSheetInstall(destFolderId, sourceRideSheetFileId, namePrefix) {
  const manifestTemplateName = "Manifest Template"
  const reportFileName = "Monthly Reporting"

  const sourceRideSheetFile = DriveApp.getFileById(sourceRideSheetFileId)
  const sourceFolder = sourceRideSheetFile.getParents().next()
  const sourceTemplateFile = sourceFolder.getFoldersByName("Settings").next().
    getFilesByName(manifestTemplateName).next()
  const sourceReportFile = sourceFolder.getFoldersByName("Reports").next().
    getFilesByName(reportFileName).next()

  const destFolder = DriveApp.getFolderById(destFolderId)
  const newManifestFolder = destFolder.createFolder("Manifests")
  const newReportsFolder = destFolder.createFolder("Reports")
  const newSettingsFolder = destFolder.createFolder("Settings")
  const newRideSheetFile = sourceRideSheetFile.makeCopy(destFolder).setName(namePrefix + " RideSheet")
  const newTemplateFile = sourceTemplateFile.makeCopy(newSettingsFolder).setName(manifestTemplateName)
  const newReportFile = sourceReportFile.makeCopy(newReportsFolder).setName(reportFileName)

  const newRideSheet = SpreadsheetApp.open(newRideSheetFile)
  const propSheet = newRideSheet.getSheetByName("Document Properties")
  const propSheetDataRange = propSheet.getDataRange()
  const propSheetData = propSheetDataRange.getValues()
  updatePropertyRange(propSheetData,"driverManifestFolderId",newManifestFolder.getId())
  updatePropertyRange(propSheetData,"driverManifestTemplateDocId",newTemplateFile.getId())
  updatePropertyRange(propSheetData,"configFolderId",newSettingsFolder.getId())
  propSheetDataRange.setValues(propSheetData)
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
    getConfiguredSheetsWithHeaders.forEach((sheetName) => {
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