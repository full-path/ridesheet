// Make sure that all named ranges go to the correct last row, and add any missing named ranges

function buildMenus() {
  let ui = SpreadsheetApp.getUi()
  let menu = ui.createMenu('RideSheet')
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
  settingsMenu.addItem('View or reload properties', 'presentProperties')
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
  } catch(e) {
    log("buildNamedRanges", e.name + ': ' + e.message)
  }
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

function buildOnChangeTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers()
  const allTriggersLength = allTriggers.length
  let exists = false
  for (var i = 0; i < allTriggersLength; i++) {
    if (allTriggers[i].getEventType() == ScriptApp.EventType.ON_CHANGE &&
        allTriggers[i].getHandlerFunction() == "onChange") { 
      exists = true
      break
    }
  }
  if (!exists) {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    ScriptApp.newTrigger("onChange").forSpreadsheet(ss).onChange().create();
  }
}

function addAndPopulateTripIdColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const tripSheet = ss.getSheetByName("Trips")
  if (getSheetHeaderNames(tripSheet).indexOf("Trip ID") === -1) {
    const lastRowPosition = tripSheet.getLastRow()
    let lastColumnPosition = tripSheet.getLastColumn()
    tripSheet.insertColumnAfter(lastColumnPosition)
    lastColumnPosition++
    tripSheet.getRange(1, lastColumnPosition).setValue("Trip ID")
    getSheetHeaderNames(tripSheet, {forceRefresh: true})
    const dataRange = tripSheet.getRange(2, 1, lastRowPosition - 1, lastColumnPosition)
    let values = getRangeValuesAsTable(dataRange)
    values.forEach((row) => {
      if (row["Trip Date"] && row["Customer ID"]) {
        row["Trip ID"] = Utilities.getUuid()
      }
    })
    setValuesByHeaderNames(values, dataRange)
  }
}