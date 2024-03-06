function buildMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('RideSheet')
  menu.addItem('Refresh driver calendars', 'updateDriverCalendars')
  menu.addItem('Add return trip', 'createReturnTrip')
  menu.addItem('Add stop', 'addStop')
  menu.addItem('Create manifests for day', 'createManifests')
  menu.addItem('Create manifests for selected items', 'createSelectedManifests')
  menu.addItem('Move past data to review', 'moveTripsToReview')
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
    menuApi.addItem('Get trip requests (Deprecated)', 'sendRequestForTripRequests')
    menuApi.addItem('Send trip requests', 'sendTripRequests')
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
        const firstRow = rangeConfigObj.headerOnly ? 1 : 2
        const lastRow = rangeConfigObj.headerOnly ? 1 : sheet.getMaxRows() + 1000
        const range = sheet.getRange(`${startColumnLetter}${firstRow}:${endColumnLetter}${lastRow}`)
        ss.setNamedRange(rangeName, range)
      }
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
    if (builder) {
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
