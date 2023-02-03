var cachedHeaderNames = {}

/**
 * Test whether a range is fully inside or matches another range. 
 * If the inner range is not fully inside the outer range, returns false.
 * @param {range} innerRange The inner range
 * @param {range} outerRange The range that the inner must be inside of or match exactly
 * @return {boolean}
 */
function isInRange(innerRange, outerRange) {
  try {
    return (
      innerRange.getSheet().getName() == outerRange.getSheet().getName() &&
      innerRange.getRow()             >= outerRange.getRow() &&
      innerRange.getLastRow()         <= outerRange.getLastRow() &&
      innerRange.getColumn()          >= outerRange.getColumn() &&
      innerRange.getLastColumn()      <= outerRange.getLastColumn()
    )
  } catch(e) { logError(e) }
}

/**
 * Test whether two ranges overlap. See:
 * https://stackoverflow.com/questions/306316/determine-if-two-rectangles-overlap-each-other
 * @param {range} firstRange The first range
 * @param {range} secondRange The second range
 * @return {boolean}
 */
function rangesOverlap(firstRange, secondRange) {
  try {
    return (
      firstRange.getSheet().getName() === secondRange.getSheet().getName() &&
      firstRange.getRow()              <= secondRange.getLastRow()         &&
      firstRange.getLastRow()          >= secondRange.getRow()             &&
      firstRange.getColumn()           <= secondRange.getLastColumn()      &&
      firstRange.getLastColumn()       >= secondRange.getColumn()
    )
  } catch(e) { logError(e) }
}

/**
 * Given a range, return the entire row that corresponds with the row
 * of the upper left corner of the passed in range. Useful with managing events.
 * @param {range} range The source range 
 * @return {range}
 */
function getFullRow(range) {
  try {
    const rowPosition = range.getRow()
    return range.getSheet().getRange("A" + rowPosition + ":" + rowPosition)
  } catch(e) { logError(e) }
}

/**
 * Given a range, return the full width of all the rows that correspond with the passed in range.
 * @param {range} range The source range 
 * @return {range}
 */
function getFullRows(range) {
  try {
    return range.getSheet().getRange("A" + range.getRow() + ":" + range.getLastRow())
  } catch(e) { logError(e) }
}

/**
 * Given sheet object and a filter function,
 * return the first found (0-based) range that is a single row where the
 * filter criteria are met
 * match the value in the headerNames parameter.
 * @return {number}
 */
function findFirstRowByHeaderNames(sheet, filter) {
  try {
    const data = getRangeValuesAsTable(sheet.getDataRange())
    const matchingRows = data.filter(row => filter(row))
    if (matchingRows.length > 0) {
      return matchingRows[0]
    }
  } catch(e) { logError(e) }
}

function moveRows(sourceSheet, destSheet, filter) {
  try {
    const sourceData = getRangeValuesAsTable(sourceSheet.getDataRange(), {includeFormulaValues: false})
    const rowsToMove = sourceData.filter(row => filter(row))
    const rowsMovedSuccessfully = createRows(destSheet, rowsToMove, false)
    if (rowsMovedSuccessfully) {
      safelyDeleteRows(sourceSheet, rowsToMove)
    }
  } catch(e) { logError(e) }
}

function moveRow(sourceRange, destSheet, {extraFields = {}} = {}) {
  try {
    const sourceSheet = sourceRange.getSheet()
    const sourceData = getRangeValuesAsTable(sourceRange, {includeFormulaValues: false})[0]
    Object.keys(extraFields).forEach(key => sourceData[key] = extraFields[key])
    if (createRow(destSheet, sourceData, true)) {
      safelyDeleteRow(sourceSheet, sourceData)  
    }
  } catch(e) { logError(e) }
}

function createRows(destSheet, data, createNewColumns=true) {
  if (createNewColumns) {
    let firstRow = data[0]
    createColumns(destSheet, firstRow)
  }
  let columnNames = getSheetHeaderNames(destSheet, {forceRefresh: createNewColumns})
  let success = data.every(row => createRow(destSheet, row))
  return success
}

const defaultColumnFilter = colHeader => {
  const colsToSkip = ["Action", "Go", "Share", "Shared", "Earliest PU Time", "Latest PU Time"]
  if (colHeader.trim() == '') {
    return false
  }
  if (colHeader.startsWith('_')) {
    return false
  }
  if (colsToSkip.includes(colHeader)) {
    return false
  }
  return true
}

function createColumns(sheet, dataRow, columnFilter=defaultColumnFilter, colOffset=0) {
  let columnNames = getSheetHeaderNames(sheet)
  let dataCols = Object.keys(dataRow).filter(colHeader => columnFilter(colHeader))
  dataCols.forEach((col) => {
    if (columnNames.indexOf(col) === -1) {
      let lastCol = sheet.getLastColumn() - colOffset
      if (lastCol < 1) lastCol = sheet.getLastColumn()
      sheet.insertColumns(lastCol)
      let headerRange = sheet.getRange(1, lastCol)
      headerRange.setValue(col)
    }
  })
}

function createRow(destSheet, data, createNewColumns=false) {
  try {
    if (createNewColumns) {
      createColumns(destSheet, data)
    }
    let columnNames = getSheetHeaderNames(destSheet, {forceRefresh: createNewColumns})
    let dataArray = columnNames.map(colName => data[colName] ? data[colName] : null)
    destSheet.appendRow(dataArray)
    let newRowIndex = destSheet.getLastRow()
    let newRow = destSheet.getRange(newRowIndex + ':' + newRowIndex)
    fixRowNumberFormatting(newRow)
    fixRowDataValidation(newRow)
    return true
  } catch(e) {
    logError(e)
    return false
  }
}

function testRowFormat() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName('Trip Review')
  let newRowIndex = sheet.getLastRow()
  let newRow = sheet.getRange(newRowIndex + ':' + newRowIndex)
  let a = newRow.getA1Notation()
  fixRowNumberFormatting(newRow)
  fixRowDataValidation(newRow)
}

function safelyDeleteRows(sheet, data) {
  if (data.length < 1) { return }
  let ss = SpreadsheetApp.getActive()
  let sheetId = sheet.getSheetId()
  let lastRowPosition = sheet.getLastRow()
  if (sheet.getMaxRows() === lastRowPosition) {
    sheet.insertRowAfter(lastRowPosition)
  }
  let rowsToDelete = data.map(row => {
    let offset = row._rowIndex + 1
    return { deleteDimension: { range: { sheetId, startIndex: offset, endIndex: offset + 1, dimension: "ROWS"}}}
    }).reverse()
  Sheets.Spreadsheets.batchUpdate({requests: rowsToDelete}, ss.getId());
}

function safelyDeleteRow(sheet, row) {
  const lastRowPosition = sheet.getLastRow()
  if (sheet.getMaxRows() === lastRowPosition) {
    sheet.insertRowAfter(lastRowPosition)
  }
  sheet.deleteRow(row._rowPosition)
}

// Takes a range and returns an array of objects, each object containing key/value pairs. 
// If the range includes row 1 of the spreadsheet, that top row will be used as the keys. 
// Otherwise row 1 will be collected separately and used as the source for keys.
function getRangeValuesAsTable(range, {headerRowPosition = 1, includeFormulaValues = true} = {}) {
  try {
    let topDataRowPosition = range.getRow()
    let values = range.getValues()
    let formulas
    if (!includeFormulaValues) formulas = range.getFormulas()
    let rangeHeaderNames
    if (topDataRowPosition <= headerRowPosition) {
      if (values.length > (headerRowPosition + 1 - topDataRowPosition)) {
        rangeHeaderNames = values[headerRowPosition - topDataRowPosition]
        values.splice(0, headerRowPosition + 1 - topDataRowPosition)
        if (!includeFormulaValues) formulas.splice(0, headerRowPosition + 1 - topDataRowPosition)
        topDataRowPosition = headerRowPosition + 1
      } else {
        return []
      }
    } else if (topDataRowPosition > headerRowPosition) {
      rangeHeaderNames = getRangeHeaderNames(range, {headerRowPosition: headerRowPosition})
    }
    let result = values.map((row, rowIndex) => {
      let rowObject = {}
      rowObject._rowPosition = rowIndex + topDataRowPosition
      rowObject._rowIndex = rowIndex
      rangeHeaderNames.forEach((headerName, columnIndex) => {
        if (includeFormulaValues || (!includeFormulaValues && !formulas[rowIndex][columnIndex])) {
          rowObject[headerName] = row[columnIndex]
        }
      })
      return rowObject
    })
    return result
  } catch(e) { logError(e) }
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
  try {
    const columnIndex = getRangeHeaderNames(range).indexOf(headerName)
    if (columnIndex == -1) {
      return null
    } else {
      return range.getDisplayValues()[0][columnIndex]
    }
  } catch(e) { logError(e) }
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
  try {
    let columnIndex = getRangeHeaderNames(range).indexOf(headerName)
    if (columnIndex > -1) {
      return range.getValues()[0][columnIndex]
    } else {
      return null
    }
  } catch(e) { logError(e) }
}

function setValuesByHeaderNames(newValues, range, {headerRowPosition = 1, overwriteAll = false} = {}) {
  try {
    const sheetHeaderNames = getSheetHeaderNames(range.getSheet(), {headerRowPosition: headerRowPosition})
    const rangeHeaderNames = getRangeHeaderNames(range, {headerRowPosition: headerRowPosition})
    const topRangeRowPosition = range.getRow()
    const topDataRowPosition = (topRangeRowPosition > headerRowPosition) ? topRangeRowPosition : headerRowPosition + 1
    const initialNumRows = range.getLastRow() - topDataRowPosition + 1
    const newValuesToApply = (initialNumRows === newValues.length) ? newValues : newValues.slice(topDataRowPosition - topRangeRowPosition)
    if (initialNumRows !== newValuesToApply.length) {
      throw new Error("Values array length does not match the number of range rows")
    }

    let narrowedRange
    let narrowedRangeHeaderNames
    let narrowedRangeValues
    let narrowedNewValuesToApply
    if (overwriteAll) {
      narrowedRange = range.getSheet().getRange(topDataRowPosition,range.getColumn(), initialNumRows, range.getNumColumns())
      narrowedRangeHeaderNames = rangeHeaderNames
      narrowedRangeValues = Array(initialNumRows).fill(null).map(row => Array(range.getNumColumns()).fill(null))
      narrowedNewValuesToApply = newValuesToApply
    } else {
      // Gather a list of the indexes of all the rows with data.
      const indexesOfRowsWithData = newValuesToApply.map((r, i) => Object.keys(r).length === 0 ? -1 : i).filter(r => r > -1)
      // If there's no actual data, quit now
      if (indexesOfRowsWithData.length === 0) return range

      // ROWS
      // Find the smallest series of rows that will update the columns that need to be updated in one update action
      const startDataRowIndex = Math.min(...indexesOfRowsWithData)
      const endDataRowIndex = Math.max(...indexesOfRowsWithData) + 1
      const firstRowPosition = topDataRowPosition + startDataRowIndex
      const numRows = endDataRowIndex - startDataRowIndex

      // COLUMNS
      // Get the full list of header names to be updated across all rows
      let headerNamesInNewValues = []
      newValuesToApply.forEach(row => {
        Object.keys(row).forEach(headerName => {
          if (!headerNamesInNewValues.includes(headerName)) headerNamesInNewValues.push(headerName)
        })
      })
      // Find the smallest series of columns that will update all the columns that need to be updated in one update action
      const headerNamePositions = headerNamesInNewValues.filter(
        headerName => rangeHeaderNames.includes(headerName)
        ).map(headerName => sheetHeaderNames.indexOf(headerName) + 1)
      // If none of the header names are in the range passed in, quit now
      if (headerNamePositions.length === 0) return range
      const firstColumnPosition = Math.min(...headerNamePositions)
      const numColumns = Math.max(...headerNamePositions) - firstColumnPosition + 1

      // PREP RANGE AND DATA
      // Create the narrowed range, based on narrowed row and column data
      narrowedRange = range.getSheet().getRange(firstRowPosition, firstColumnPosition, numRows, numColumns)
      narrowedRangeHeaderNames = getRangeHeaderNames(narrowedRange)
      narrowedRangeValues = narrowedRange.getValues()
      narrowedNewValuesToApply = newValuesToApply.slice(startDataRowIndex, endDataRowIndex)
    }

    // Update the array of arrays with the new values
    narrowedRangeValues.forEach((sheetRow, sheetRowIndex) => {
      narrowedRangeHeaderNames.forEach((rangeHeaderName, rangeHeaderIndex) => {
        if (Object.keys(narrowedNewValuesToApply[sheetRowIndex]).indexOf(rangeHeaderName) > -1) {
          sheetRow[rangeHeaderIndex] = narrowedNewValuesToApply[sheetRowIndex][rangeHeaderName]
        }
      })
    })
    // Do the actual update
    narrowedRange.setValues(narrowedRangeValues)
    // Return the original range, for chaining
    return range
  } catch(e) { logError(e) }
}

function appendValuesByHeaderNames(values, sheet) {
  try {
    const sheetHeaderColumnNames = getSheetHeaderNames(sheet)
    values.forEach(row => {
      const rowArray = sheetHeaderColumnNames.map(colName => row[colName])
      appendRowWithFormatting(sheet, rowArray)
    })
  } catch(e) { logError(e) }
}

// Cache header info in a global variable so it only needs to be collected once per sheet per onEdit call.
function getSheetHeaderNames(sheet, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetName = sheet.getName()
    if (!cachedHeaderNames[sheetName] || forceRefresh) {
      const headerNames = sheet.getRange("A" + headerRowPosition + ":" + headerRowPosition).getValues()[0]
      cachedHeaderNames[sheetName] = headerNames.map(headerName => !headerName ? " " : headerName)
    }
    return cachedHeaderNames[sheetName]
  } catch(e) { logError(e) }
}
    
// Get header information for column range only, rather than the entire sheet
// Uses getSheetHeaderNames for caching purposes.
function getRangeHeaderNames(range, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetHeaderNames = getSheetHeaderNames(range.getSheet(), {forceRefresh: forceRefresh, headerRowPosition: headerRowPosition})
    const rangeStartColumnIndex = range.getColumn() - 1
    return sheetHeaderNames.slice(rangeStartColumnIndex, rangeStartColumnIndex + range.getWidth())
  } catch(e) { logError(e) }
}

function getMaxValueInRange(range) {
  try {
    let values = range.getValues().flat().filter(Number.isFinite)
    return values.reduce((a, b) => Math.max(a, b))
  } catch(e) { logError(e) }
}

function applyFormats(formatGroups, sheet) {
  try {
    Object.keys(formatGroups).forEach(groupName => {
      const ranges = formatGroups[groupName].ranges
      if (ranges.length) formatGroups[groupName].formats(sheet.getRangeList(ranges))
    })
  } catch(e) { logError(e) }
}

function getColumnLettersFromPosition(colPosition) {
  try {
    const letterSeriesStart = "A".charCodeAt()
    const letterCount = "Z".charCodeAt() - letterSeriesStart + 1
    let columnLetters = []
    let remainder = colPosition - 1
    while (remainder >= 0) {
      columnLetters.unshift(String.fromCharCode((remainder % letterCount) + letterSeriesStart))
      remainder = Math.floor(remainder / letterCount) - 1
    }
    return columnLetters.join("")
  } catch(e) { logError(e) }
}