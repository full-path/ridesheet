var cachedHeaderNames = {}

/**
 * Test whether a range is fully inside or matches another range. 
 * If the inner range is not fully inside the outer range, returns false.
 * @param {range} innerRange The inner range
 * @param {range} outerRange The range that the inner must be inside of or match exactly
 * @return {boolean}
 */
function isInRange(innerRange, outerRange) {
  return (
    innerRange.getSheet().getName() == outerRange.getSheet().getName() &&
    innerRange.getRow()             >= outerRange.getRow() &&
    innerRange.getLastRow()         <= outerRange.getLastRow() &&
    innerRange.getColumn()          >= outerRange.getColumn() &&
    innerRange.getLastColumn()      <= outerRange.getLastColumn()
  )
}

/**
 * Test whether two ranges overlap. See:
 * https://stackoverflow.com/questions/306316/determine-if-two-rectangles-overlap-each-other
 * @param {range} firstRange The first range
 * @param {range} secondRange The second range
 * @return {boolean}
 */
function rangesOverlap(firstRange, secondRange) {
    return (
      firstRange.getSheet().getName() === secondRange.getSheet().getName() &&
      firstRange.getRow()              <= secondRange.getLastRow()         &&
      firstRange.getLastRow()          >= secondRange.getRow()             &&
      firstRange.getColumn()           <= secondRange.getLastColumn()      &&
      firstRange.getLastColumn()       >= secondRange.getColumn()
    )
}

/**
 * Given a range, return the entire row that corresponds with the row
 * of the upper left corner of the passed in range. Useful with managing events.
 * @param {range} range The source range 
 * @return {range}
 */
function getFullRow(range) {
  const rowPosition = range.getRow()
  return range.getSheet().getRange("A" + rowPosition + ":" + rowPosition)
}

/**
 * Given a range, return the full width of all the rows that correspond with the passed in range.
 * @param {range} range The source range 
 * @return {range}
 */
function getFullRows(range) {
  return range.getSheet().getRange("A" + range.getRow() + ":" + range.getLastRow())
}

/**
 * Given sheet object and a filter function,
 * return the first found (0-based) range that is a single row where the
 * filter criteria are met
 * match the value in the headerNames parameter.
 * @return {number}
 */
function findFirstRowByHeaderNames(sheet, filter) {
  const data = getRangeValuesAsTable(sheet.getDataRange())  
  const matchingRows = data.filter(row => filter(row))
  if (matchingRows.length > 0) {
    return matchingRows[0]
  }
}

function moveRows(sourceSheet, destSheet, filter) {
  const sourceData = getRangeValuesAsTable(sourceSheet.getDataRange())
  const rowsToMove = sourceData.filter(row => filter(row))
  const lastRowPosition = sourceSheet.getLastRow()
  rowsToMove.forEach(row => appendDataRow(sourceSheet, destSheet, row))
  if (sourceSheet.getMaxRows() === lastRowPosition) { sourceSheet.insertRowAfter(lastRowPosition) }
  const rowsToDelete = rowsToMove.map(row => row.rowPosition).sort((a,b)=>b-a)
  rowsToDelete.forEach(rowPosition => sourceSheet.deleteRow(rowPosition))
}

// Take an incoming map of values and append them to the sheet, matching column names to map key names
// If columns present in the source data are missing in the destination sheet,
// Those columns will be added to the destination sheet, right after the column they're after in the
// source sheet.
function appendDataRow(sourceSheet, destSheet, dataMap) {
  const sourceColumnNames = Object.keys(dataMap)
  const destColumnNamesOriginalState = getSheetHeaderNames(destSheet)
  let destColumnNamesCurrentState = destColumnNamesOriginalState
  let missingDestColumns = []
  sourceColumnNames.forEach((sourceColumnName, i) => {
    if (destColumnNamesOriginalState.indexOf(sourceColumnName) === -1 && sourceColumnName !== "rowPosition") {
      let colPosition = 1
      if (i === 0) {
        destSheet.insertColumns(colPosition)
      } else {
        let positionOfColumnToInsertAfter = destColumnNamesCurrentState.indexOf(sourceColumnNames[i-1]) + 1
        if (positionOfColumnToInsertAfter === 0) {
          destSheet.insertColumns(colPosition)
        } else {
          colPosition = positionOfColumnToInsertAfter + 1
          destSheet.insertColumnAfter(positionOfColumnToInsertAfter)
        }
      }
      let sourceRange = sourceSheet.getRange(2, getSheetHeaderNames(sourceSheet).indexOf(sourceColumnName) + 1)
      let destHeaderRange = destSheet.getRange(1, colPosition)
      let destDataRange = destSheet.getRange(2, colPosition, destSheet.getMaxRows()-1)
      
      destHeaderRange.setValue(sourceColumnName)
      sourceRange.copyFormatToRange(destSheet, colPosition, colPosition, 2, destSheet.getMaxRows())
      let rule = sourceRange.getDataValidation()
      if (rule == null) {
        destDataRange.clearDataValidations()
      } else {
        destDataRange.setDataValidation(rule)
      }
      destColumnNamesCurrentState = getSheetHeaderNames(destSheet, {forceRefresh: true})
    }
  })
  const dataArray = destColumnNamesCurrentState.map(colName => dataMap[colName])
  destSheet.appendRow(dataArray)
}

// Takes a range and returns an array of objects, each object containing key/value pairs. 
// If the range includes row 1 of the spreadsheet, that top row will be used as the keys. 
// Otherwise row 1 will be collected separately and used as the source for keys.
function getRangeValuesAsTable(range) {
  let topRowPosition = range.getRow()
  let data = range.getValues()
  let rangeHeaderNames
  if (topRowPosition === 1) {
    if (data.length > 1) {
      rangeHeaderNames = data.shift()
      topRowPosition = 2
    } else {
      return []
    }
  } else {
    rangeHeaderNames = getRangeHeaderNames(range)
  }
  let result = data.map((row, index) => {
    let rowMap = {}
    rowMap.rowPosition = index + topRowPosition
    rangeHeaderNames.forEach((headerName, i) => rowMap[headerName] = row[i])
    return rowMap
  })
  return result            
}

/**
 * Given a desired column name and a range, 
 * return the display value of the first row of the column whose header row value
 * matches the headerName. Returns null if column cannot be found.
 * @param {string} headerName The name of the header
 * @param {range} range The range 
 * @return {object}
 */
function getDisplayValueByHeaderName(headerName, range) {
  const columnIndex = getRangeHeaderNames(range).indexOf(headerName)
  if (columnIndex == -1) {
    return null
  } else {
    return range.getDisplayValues()[0][columnIndex]
  }
}

/**
 * Given a desired column name and a range, 
 * return the value of the first row of the column whose header row value
 * matches the headerName. Returns null if column cannot be found.
 * @param {string} headerName The name of the header
 * @param {range} range The range 
 * @return {object}
 */
function getValueByHeaderName(headerName, range) {
  let columnIndex = getRangeHeaderNames(range).indexOf(headerName)
  if (columnIndex > -1) {
    return range.getValues()[0][columnIndex]
  } else {
    return null
  }
}

function setValuesByHeaderNames(newValues, range) {
  try {
    const rangeHeaderNames = getRangeHeaderNames(range)
    const rangeIncludesHeaderRow = (range.getRow() === 1)
    const rowOffset = (rangeIncludesHeaderRow ? 1 : 0)
    let rangeValues = range.getValues()
    //log("newValues: " + newValues.length, "rangeValues: " + rangeValues.length)
    rangeValues.forEach((sheetRow, sheetRowIndex) => {
      if (rangeIncludesHeaderRow && sheetRowIndex === 0) {
        // skip header row
      } else {
        rangeHeaderNames.forEach((rangeHeaderName, rangeHeaderIndex) => {
          if (Object.keys(newValues[sheetRowIndex - rowOffset]).indexOf(rangeHeaderName) > -1) {
            sheetRow[rangeHeaderIndex] = newValues[sheetRowIndex - rowOffset][rangeHeaderName]
          }
        })
      }
    })
    return range.setValues(rangeValues)
  } catch(e) {
    logError(e)
  }
}

function appendValuesByHeaderNames(values, sheet) {
  const sheetHeaderColumnNames = getSheetHeaderNames(sheet)
  values.forEach(row => {
    const rowArray = sheetHeaderColumnNames.map(colName => row[colName])
    sheet.appendRow(rowArray)
  })
}

// Cache header info in a global variable so it only needs to be collected once per sheet per onEdit call.
function getSheetHeaderNames(sheet, {forceRefresh = false} = {}) {
  const sheetName = sheet.getName()
  if (!cachedHeaderNames[sheetName] || forceRefresh) {
    const headerNames = sheet.getRange("A1:1").getValues()[0]
    cachedHeaderNames[sheetName] = headerNames.map(headerName => !headerName ? " " : headerName)
  }
  return cachedHeaderNames[sheetName]
}
    
// Get header information for column range only, rather than the entire sheet
// Uses getSheetHeaderNames for caching purposes.
function getRangeHeaderNames(range, {forceRefresh = false} = {}) {
  const sheetHeaderNames = getSheetHeaderNames(range.getSheet(), {forceRefresh: forceRefresh})
  const rangeStartColumnIndex = range.getColumn() - 1
  return sheetHeaderNames.slice(rangeStartColumnIndex, rangeStartColumnIndex + range.getWidth())
}

function getMaxValueInRange(range) {
  let values = range.getValues().flat().filter(Number.isFinite)
  return values.reduce((a, b) => Math.max(a, b))
}

function testSetDisplayValueByHeaderName() {
  let startTime = new Date()
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips")
  let range = sheet.getRange("A7:M7")
  setValuesByHeaderNames({"Trip Date": "5/23/2020", "Customer Name":"Johnson, Howard (1)"}, range)
}

function testGetDisplayValueByHeaderName() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips")
  let range = sheet.getRange("A7:M7")

  let displayValue = getDisplayValueByHeaderName("Trip Date",range)
  log(displayValue.toString(), Array.isArray(displayValue))
  displayValue = getValueByHeaderName("Trip Date",range)
  log(displayValue.toString(), Array.isArray(displayValue))
}