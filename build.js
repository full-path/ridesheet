// Make sure that all named ranges go to the correct last row, and add any missing named ranges

function buildMenus() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('RideSheet')
  menu.addItem('Refresh driver calendars', 'updateDriverCalendars')
  menu.addItem('Refresh outside runs', 'sendRequestForRuns')
  menu.addItem('Add return trip', 'createReturnTrip')
  menu.addItem('Add stop', 'addStop')
  menu.addItem('Create manifests for day', 'createManifests')
  menu.addItem('Create manifests for selected items', 'createSelectedManifests')
  menu.addItem('Move past data to review', 'moveTripsToReview')
  menu.addItem('Move reviewed data to archive', 'moveTripsToArchive')
  menu.addSeparator()
  let settingsMenu = ui.createMenu('Settings')
  settingsMenu.addItem('Application properties', 'presentProperties')
  settingsMenu.addItem('Scheduled calendar updates', 'presentCalendarTrigger')
  menu.addSubMenu(settingsMenu)
  menu.addToUi()
}

function buildNamedRanges() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const namedRanges = ss.getNamedRanges()
    const currentRangeNames = namedRanges.map(nr => nr.getName())
    const buildRangeNames = Object.keys(defaultNamedRanges)
    namedRanges.forEach(namedRange => {
      if (buildRangeNames.indexOf(namedRange.getName()) !== -1 && 
          (namedRange.getRange().getRow() !== 1 || 
          namedRange.getRange().getLastRow() !== namedRange.getRange().getSheet().getMaxRows() + 1000)) {
        const name = namedRange.getName()
        let newRange = defaultNamedRanges[name]
        //namedRange.remove()
        buildNamedRange(ss, name, newRange.sheetName, newRange.column, newRange.headerName)
      }
    })
    buildRangeNames.forEach(rangeName => {
      if (currentRangeNames.indexOf(rangeName) === -1) {
        let newRange = defaultNamedRanges[rangeName]
        buildNamedRange(ss, rangeName,newRange.sheetName, newRange.column, newRange.headerName)
      }
    })
  } catch(e) { logError(e) }
}

function buildNamedRange(ss, name, sheetName, column, headerName) {
  const sheet = ss.getSheetByName(sheetName)
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
}

// Document properties don't pass on to copied sheets. This recreates the ones put into the properties sheet.
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
  }
}

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
  setDocProps(newProps)
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
    log(JSON.stringify(results))
  } catch(e) { logError(e) }
}

function buildMetadata() {
  try {
    clearMetadata()
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const extraHeaderNames = getDocProp("extraHeaderNames")
    defaultSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName)
      sheet.addDeveloperMetadata("sheetName",sheetName,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    })
    sheetsWithHeaders.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName)
      const sheetHeaderNames = getSheetHeaderNames(sheet)
      const correctHeaderNames = [...Object.keys(defaultColumns[sheetName] || {}),...extraHeaderNames[sheetName]]
      sheetHeaderNames.forEach((shn, i) => {
        if (correctHeaderNames.includes(shn)) {
          let letter = getColumnLettersFromPosition(i + 1)
          let range = sheet.getRange(`${letter}:${letter}`)
          range.addDeveloperMetadata("headerName",shn,SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
        }
      })
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

function fixSheetNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    let mds = ss.createDeveloperMetadataFinder().
      withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET).
      withKey("sheetName").find()
    mds.forEach(md => {
      const sheet = md.getLocation().getSheet()
      //log(md.getKey(), md.getValue(), md.getLocation().getSheet().getName())
      if (sheet.getName() !== md.getValue()) {
        log(`Sheet Name '${sheet.getName()}' updated to '${md.getValue()}'`)
        sheet.setName(md.getValue())
      }
    })
  } catch(e) { logError(e) }
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

function findMetadata() {
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