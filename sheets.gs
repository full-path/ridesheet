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
 * Given a hash of header names and data values to filter by and a sheet object, 
 * return the first found (0-based) range that is a single row where all
 * filter criteria are met
 * match the value in the headerNames parameter. Returns -1 if column cannot be found.
 * @param {string} headerName The header name
 * @param {range} range The range 
 * @return {number}
 */
function findFirstRowByHeaderNames(values, sheet) {
  const range = sheet.getDataRange()
  const data = range.getDisplayValues()
  const headerNames = Object.keys(values)
  const columnNumbers = getColNumbersByHeaderNames(headerNames, range)
  let isEqual
  let result
  
  // Row loop
  for (r = 0, l = data.length; r < l; r++) {
    isEqual = true
    columnLoop:
    for (i = 0, n = columnNumbers.length; i < n; i++) {
      if (data[r][columnNumbers[i]] !== values[headerNames[i]]) {
        isEqual = false
        break columnLoop
      }
    }
    if (isEqual) { 
      result = sheet.getRange(r + 1, 1, 1, data[0].length) 
      break
    }
  }
  return result
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

/**
 * Given a desired column name and a range, 
 * return the first found (0-based) column numbers whose header row value
 * match the value in the headerNames parameter. Returns -1 if column cannot be found.
 * @param {string} headerName The header name
 * @param {range} range The range 
 * @return {number}
 */
function getColNumberByHeaderName(headerName, range) {
  return JSON.parse(PropertiesService.getDocumentProperties().getProperty(sheetHeaderPropertyName(range.getSheet().getName())))[headerName]
  //let headerValues = range.getSheet().getRange(headerRow, range.getColumn(), 1, range.getWidth()).getDisplayValues()[0]
  //return headerValues.indexOf(headerName)
}

/**
 * Given an array of desired column names and a range, 
 * return an array of first-found (0-based) column numbers whose header row values
 * match each value in the headerNames array. Returns -1 for each column that cannot be found.
 * @param {array} headerNames The names of the headers
 * @param {range} range The range 
 * @return {array}
 */
function getColNumbersByHeaderNames(headerNames, range) {
  //let headerValues = range.getSheet().getRange(headerRow, range.getColumn(), 1, range.getWidth()).getDisplayValues()[0]
  let headerValues = JSON.parse(PropertiesService.getDocumentProperties().getProperty(sheetHeaderPropertyName(range.getSheet().getName())))
  return headerNames.map(i => headerValues[i])
}

// Takes a data range with a first row header and turns it into an array of data objects that function as maps.
function getDataRangeAsTable(range) {
  let result = []
  const headerNames = range.shift()
  range.forEach(row => {
    let rowMap = {}
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

function storeHeaderInformation(e) {
  const docProperties = PropertiesService.getDocumentProperties()
  let allProperties = []
  sheetsWithHeaders.forEach(sheetName => {
    let sheet = e.source.getSheetByName(sheetName)
    if (sheet) {
      let headerRange = sheet.getRange("A1:1")
      let headerValues = headerRange.getDisplayValues()[0]
      let headerHash = {}
      headerValues.forEach((colName, i) => {
        if (colName) { headerHash[colName] = i }
      })
      let propHash = {}
      propHash.name = sheetHeaderPropertyName(sheetName)
      propHash.value = headerHash
      allProperties.push(propHash)
    }
  })
  setDocProps(allProperties)
}

function sheetHeaderPropertyName(sheet) {
  return `sheetHeaders["${sheet}"]_`
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