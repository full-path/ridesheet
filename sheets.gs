var headerInformation = {}
const sheetsWithHeaders = ["Trips","Runs","Trip Review","Run Review","Recurring Trips","Customers","Trip Archive","Run Archive","Drivers","Services"]

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
 * Given sheet object and a filter function,
 * return the first found (0-based) range that is a single row where the
 * filter criteria are met
 * match the value in the headerNames parameter.
 * @return {number}
 */
function findFirstRowByHeaderNames(sheet, filter) {
  const data = getDataRangeAsTable(sheet.getDataRange().getValues())  
  const matchingRows = data.filter(row => filter(row))
  if (matchingRows.length > 0) {
    return matchingRows[0]
  }
}

function moveRows(sourceSheet, destSheet, filter) {
  const sourceData = getDataRangeAsTable(sourceSheet.getDataRange().getValues())
  const rowsToMove = sourceData.filter(row => filter(row))
  rowsToMove.forEach(row => appendDataRow(sourceSheet, destSheet, row))
  const rowsToDelete = rowsToMove.map(row => row.rowPosition).sort((a,b)=>b-a)
  rowsToDelete.forEach(rowPosition => sourceSheet.deleteRow(rowPosition))
}

/**
 * Given a range, return the entire row that corresponds with the row
 * of the upper left corner of the passed in range. Useful with managing events.
 * @param {range} range The source range 
 * @return {range}
 */
function getFullRow(range) {
  const rowNum = range.getRow()
  return range.getSheet().getRange("A" + rowNum + ":" + rowNum)
}

// Take an incoming map of values and append them to the sheet, matching column names to map key names
function appendDataRow(sourceSheet, destSheet, dataMap) {
  const sourceColumnNames = Object.keys(dataMap)
  const originalDestColumnInfo = getHeaderInformation(destSheet.getName())
  let currentDestColumnInfo = originalDestColumnInfo
  let missingDestColumns = []
  sourceColumnNames.forEach((sourceColumnName, i) => {
    if (Object.keys(originalDestColumnInfo).indexOf(sourceColumnName) === -1 && sourceColumnName !== "rowPosition") {
      let colPosition = 1
      if (i > 0) {
        let positionOfColumnToInsertAfter = Object.keys(currentDestColumnInfo).indexOf(sourceColumnNames[i-1]) + 1
        if (positionOfColumnToInsertAfter > 0) colPosition = positionOfColumnToInsertAfter + 1
      }
      destSheet.insertColumns(colPosition)
      let sourceRange = sourceSheet.getRange(2, getHeaderInformation(sourceSheet.getName())[sourceColumnName] + 1)
      let destHeaderRange = destSheet.getRange(1, colPosition)
      let destDataRange = destSheet.getRange(2, colPosition, destSheet.getMaxRows()-1)
      
      destHeaderRange.setValue(sourceColumnName)
      sourceRange.copyFormatToRange(destSheet, colPosition, colPosition + 1, 2, destSheet.getMaxRows())
      let rule = sourceRange.getDataValidation()
      if (rule == null) {
        destDataRange.clearDataValidations()
      } else {
        destDataRange.setDataValidation(rule)
      }
      currentDestColumnInfo = getHeaderInformation(destSheet.getName(), true)
    }
  })
  const dataArray = Object.keys(currentDestColumnInfo).map(colName => dataMap[colName])
  destSheet.appendRow(dataArray)
}

/**
 * Given a desired column name and a range, 
 * return the first found (0-based) column numbers whose header row value
 * match the value in the headerNames parameter.
 * @param {string} headerName The header name
 * @param {range} range The range 
 * @return {number}
 */
function getColNumberByHeaderName(headerName, range) {
  return getHeaderInformation(range.getSheet().getName())[headerName]
}

/**
 * Given an array of desired column names and a range, 
 * return an array of first-found (0-based) column numbers whose header row values
 * match each value in the headerNames array.
 * @param {array} headerNames The names of the headers
 * @param {range} range The range 
 * @return {array}
 */
function getColNumbersByHeaderNames(headerNames, range) {
  let headerValues = getHeaderInformation(range.getSheet().getName())
  return headerNames.map(i => headerValues[i])
}

// Takes a data range with a first row header and turns it into an array of data objects that function as maps.
function getDataRangeAsTable(dataRange) {
  let result = []
  const headerNames = dataRange.shift()
  dataRange.forEach((row, index) => {
    let rowMap = {}
    rowMap.rowPosition = index + 2
    headerNames.forEach((headerName, i) => rowMap[headerName] = row[i])
    result.push(rowMap)
  })
  return result            
}

/**
 * Given an array of desired column names and a range, 
 * return the display values of the first row of the columns whose header row values
 * matches the headerNames array. Returns null if column cannot be found.
 * @param {array} headerNames The names of the headers
 * @param {range} range The range 
 * @return {object}
 */
function getDisplayValuesByHeaderNames(headerNames, range) {
  let columnNumbers = getColNumbersByHeaderNames(headerNames, range)
  result = {}
  displayValues = range.getDisplayValues()[0]
  columnNumbers.forEach((colNumber, i) => {
    if (colNumber == -1) {
      result[headerNames[i]] = null
    } else {
      result[headerNames[i]] = displayValues[colNumber]
    }
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
  let colNumber = getColNumberByHeaderName(headerName, range)
  if (colNumber == -1) {
    return null
  } else {
    return range.getDisplayValues()[0][colNumber]
  }
}

/**
 * Given an array of desired column names and a range, 
 * return the display values of the first row of the columns whose header row values
 * matches the headerNames array. Returns null if column cannot be found.
 * @param {array} headerNames The names of the headers
 * @param {range} range The range 
 * @return {object}
 */
function getValuesByHeaderNames(headerNames, range) {
  let columnNumbers = getColNumbersByHeaderNames(headerNames, range)
  result = {}
  values = range.getValues()[0]
  columnNumbers.forEach((colNumber, i) => {
    if (colNumber == -1) {
      result[headerNames[i]] = null
    } else {
      result[headerNames[i]] = values[colNumber]
    }
  })
  return result
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
  let colNumber = getColNumberByHeaderName(headerName, range)
  if (colNumber > -1) {
    return range.getValues()[0][colNumber]
  } else {
    return null
  }
}

function setValuesByHeaderNames(values, range) {
  const headerNames = Object.keys(values)
  const columnNumbers = getColNumbersByHeaderNames(headerNames, range)
  let rangeValues = range.getValues()
  columnNumbers.forEach((colNumber, i) => {
    if (colNumber > -1) {
      rangeValues[0][colNumber] = values[headerNames[i]]
    }
  })
  return range.setValues(rangeValues)
}

// Cache header info in a global variable so it only needs to be collected once per sheet per onEdit call.
function getHeaderInformation(sheetName, forceRefresh) {
  if (!headerInformation[sheetName] || forceRefresh) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
    if (sheet) {
      const headerValues = sheet.getRange("A1:1").getDisplayValues()[0]
      let headerHash = {}
      headerValues.forEach((colName, i) => {
        if (colName) headerHash[colName] = i
      })
      headerInformation[sheetName] = headerHash
    }
    //log("Collected header information")
  }
  return headerInformation[sheetName]
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

  let columnNumber = getColNumbersByHeaderNames(["Trip Date", "CustomerName"],range)
  log(columnNumber.toString())
  columnNumber = getColNumberByHeaderName("Trip Date",range)
  log(columnNumber.toString())
  
  let displayValues = getDisplayValuesByHeaderNames(["Trip Date", "Customer Name"],range)
  log(displayValues.toString(), Array.isArray(displayValues))
  displayValues = getValuesByHeaderNames(["Trip Date", "Customer Name"],range)
  log(displayValues.toString(), Array.isArray(displayValues))
  let displayValue = getDisplayValueByHeaderName("Trip Date",range)
  log(displayValue.toString(), Array.isArray(displayValue))
  displayValue = getValueByHeaderName("Trip Date",range)
  log(displayValue.toString(), Array.isArray(displayValue))
}

function getMaxValueInRange(range) {
  let values = range.getValues().flat().filter(Number.isFinite)
  return values.reduce((a, b) => Math.max(a, b))
}