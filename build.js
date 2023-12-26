// Make sure that all named ranges go to the correct last row, and add any missing named ranges

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
    menuApi.addItem('Get trip requests', 'sendRequestForTripRequests')
    menuApi.addItem('Send responses to trip requests', 'sendTripRequestResponses')
    menuApi.addItem('Refresh outside runs', 'sendRequestForRuns')
    menuApi.addToUi()
  }
  menu.addToUi()
}

function buildNamedRanges() {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const namedRanges = ss.getNamedRanges()
    const currentRangeNames = namedRanges.map(nr => nr.getName())
    const configuredNamedRanges = {...defaultNamedRanges, ...localNamedRanges}
    const buildRangeNames = Object.keys(configuredNamedRanges)
    namedRanges.forEach(namedRange => {
      try {
        if (buildRangeNames.indexOf(namedRange.getName()) !== -1 && 
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
        if (currentRangeNames.indexOf(rangeName) === -1) {
          let newRange = configuredNamedRanges[rangeName]
          buildNamedRange(ss, rangeName, newRange.sheetName, newRange.column, newRange.headerName)
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
      log(name,range.getSheet().getName() + "!" + column + "1:" + column + (sheet.getMaxRows() + 1000))
      ss.setNamedRange(name, range)
    }
  } else {
    log(`Attempted to build named Range to '${name}' to sheet '${sheetName}', but sheet not found.`)
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
    const extraHeaderNames = getDocProp("extraHeaderNames")
    let results = {}
    sheetsWithHeaders.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName)
      const sheetHeaderNames = getSheetHeaderNames(sheet)
      const defaultSheetHeaderNames = Object.keys(defaultColumns[sheetName] || {})
      const extraSheetHeaderNames = extraHeaderNames[sheetName]

      let sheetResults = {}
      sheetResults["defaultPresent"] = defaultSheetHeaderNames.filter(x => sheetHeaderNames.includes(x))
      sheetResults["defaultMissing"] = defaultSheetHeaderNames.filter(x => !sheetHeaderNames.includes(x))
      const sheetHeaderNamesForConfig = sheetHeaderNames.filter(x => !sheetResults["defaultPresent"].includes(x))
      sheetResults["configPresent"] = extraSheetHeaderNames.filter(x => sheetHeaderNamesForConfig.includes(x))
      sheetResults["configMissing"] = extraSheetHeaderNames.filter(x => !sheetHeaderNamesForConfig.includes(x))
      sheetResults["notTracked"] = sheetHeaderNamesForConfig.filter(x => !sheetResults["configPresent"].includes(x))
      results[sheetName] = sheetResults
    })
  } catch(e) { logError(e) }
}

function buildMetadata() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let sheetMetadata = ss.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET).
      withKey("sheetName").find()
    let labeledSheets = sheetMetadata.map(md => md.getValue())
    defaultSheets.forEach(sheetName => {
      if (!labeledSheets.includes(sheetName)) {
        let sheet = ss.getSheetByName(sheetName)
        if (sheet) {
          sheet.addDeveloperMetadata("sheetName",sheetName,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
        } else {
          log(`Attempted to add metadata to 'defaultSheets' sheet '${sheetName}', but sheet not found.`)
        }
      }
    })
    sheetsWithHeaders.forEach(sheetName => {
      let sheet = ss.getSheetByName(sheetName)
      if (sheet) {
        let extraHeaderNames = getDocProp("extraHeaderNames")
        let sheetHeaderNames = getSheetHeaderNames(sheet)
        let registeredColumns = [...Object.keys(defaultColumns[sheetName] || {}),...extraHeaderNames[sheetName]]
        sheetHeaderNames.forEach((columnName, i) => {
          if (registeredColumns.includes(columnName)) {
            let letter = getColumnLettersFromPosition(i + 1)
            let range = sheet.getRange(`${letter}:${letter}`)
            let colMetadata = range.createDeveloperMetadataFinder()
              .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
              .find()
            if (colMetadata.length < 1) {
              let colSettings = defaultColumns[sheetName][columnName]
              range.addDeveloperMetadata("headerName",columnName,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
              if (colSettings && colSettings["numberFormat"]) {
                range.addDeveloperMetadata("numberFormat", colSettings["numberFormat"], SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
              }
              if (colSettings && colSettings["dataValidation"]) {
                let validationRules = JSON.stringify(colSettings["dataValidation"])
                range.addDeveloperMetadata("dataValidation", validationRules, SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
              }
            }
          }
        })
      } else {
        log(`Attempted to add metadata to 'sheetsWithHeaders' sheet '${sheetName}', but sheet not found.`)
      }
    })
  } catch(e) {
    logError(e)
  }
}

function rebuildAllMetadata() {
  try {
    clearMetadata()
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const extraHeaderNames = getDocProp("extraHeaderNames")
    defaultSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName)
      if (sheet) {
        sheet.addDeveloperMetadata("sheetName",sheetName,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
      } else {
        log(`Attempted to add metadata to 'defaultSheets' sheet '${sheetName}', but sheet not found.`)
      }
    })
    sheetsWithHeaders.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName)
      if (sheet) {
        const sheetHeaderNames = getSheetHeaderNames(sheet)
        const correctHeaderNames = [...Object.keys(defaultColumns[sheetName] || {}),...extraHeaderNames[sheetName]]
        sheetHeaderNames.forEach((shn, i) => {
          if (correctHeaderNames.includes(shn)) {
            let letter = getColumnLettersFromPosition(i + 1)
            let range = sheet.getRange(`${letter}:${letter}`)
            range.addDeveloperMetadata("headerName",shn,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
            let colSettings = defaultColumns[sheetName][shn]
            if (colSettings && colSettings["numberFormat"]) {
              range.addDeveloperMetadata("numberFormat", colSettings["numberFormat"], SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
            }
            if (colSettings && colSettings["dataValidation"]) {
              let validationRules = JSON.stringify(colSettings["dataValidation"])
              range.addDeveloperMetadata("dataValidation", validationRules, SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
            }
          }
        })
      } else {
        log(`Attempted to add metadata to 'sheetsWithHeaders' sheet '${sheetName}', but sheet not found.`)
      }
    })
  } catch(e) { logError(e) }
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
      log(criteria, JSON.stringify(args))
      //builder = SpreadsheetApp.newDataValidation().withCriteria(criteria, args).setAllowInvalid(allowInvalid)
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
    let format = md.getValue()
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
    const sheet = range.getSheet()
    const mds = range.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN).
      onIntersectingLocations().
      withKey("headerName").
      find()
    let columnsPositionsToFix = []
    let headerNames = {}
    mds.forEach(md => {
      headerNames[md.getLocation().getColumn().getColumn()] = md.getValue()
      if (md.getLocation().getColumn().getValue() !== md.getValue()) {
        columnsPositionsToFix.push(md.getLocation().getColumn().getColumn())
      }
    })
    if (columnsPositionsToFix.length) {
      let firstColPos = Math.min(...columnsPositionsToFix)
      let lastColPos = Math.max(...columnsPositionsToFix)
      let range = sheet.getRange(1, firstColPos, 1, lastColPos - firstColPos + 1)
      let values = [[]]
      for (let i = firstColPos; i <= lastColPos; i++) values[0].push(headerNames[i])
      range.setValues(values)
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
      //Logger.log(Object.keys(metadata[sheetName]).map((key) ))
      const lastColumnNumber = Math.max(...Object.keys(metadata[sheetName]))
      //Logger.log(sheetName + ": " + lastColumnNumber)
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
    sheetsWithHeaders.forEach((sheetName) => {
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