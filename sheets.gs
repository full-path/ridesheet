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
function getRangeValuesAsTable(range, {headerRowPosition = 1} = {}) {
  let topRowPosition = range.getRow()
  let data = range.getValues()
  let rangeHeaderNames
  if (topRowPosition <= headerRowPosition) {
    if (data.length > (headerRowPosition + 1 - topRowPosition)) {
      rangeHeaderNames = data[headerRowPosition - topRowPosition]
      data.splice(0, headerRowPosition + 1 - topRowPosition)
      topRowPosition = headerRowPosition + 1
    } else {
      return []
    }
  } else if (topRowPosition > headerRowPosition) {
    rangeHeaderNames = getRangeHeaderNames(range, {headerRowPosition: headerRowPosition})
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

function setValuesByHeaderNames(newValues, range, {headerRowPosition = 1} = {}) {
  try {
    const sheetHeaderNames = getSheetHeaderNames(range.getSheet(), {headerRowPosition: headerRowPosition})
    const rangeHeaderNames = getRangeHeaderNames(range, {headerRowPosition: headerRowPosition})
    const topRangeRowPosition = range.getRow()
    const topDataRowPosition = (topRangeRowPosition > headerRowPosition) ? topRangeRowPosition : headerRowPosition + 1
    const numRows = range.getLastRow() - topDataRowPosition + 1
    const newValuesToApply = newValues.slice(topDataRowPosition - topRangeRowPosition)

    // Get the full list of header names for columns to be updated
    let headerNamesInNewValues = []
    newValuesToApply.forEach(row => {
      Object.keys(row).forEach(headerName => {
        if (headerNamesInNewValues.indexOf(headerName) === -1) headerNamesInNewValues.push(headerName)
      })
    })
    
    // Find the smallest range that will update the columns that need to be updated in one update action
    const headerNamePositions = headerNamesInNewValues.filter(
      headerName => rangeHeaderNames.indexOf(headerName) > -1 
      ).map(headerName => sheetHeaderNames.indexOf(headerName) + 1)
    const firstColumnPosition = Math.min(...headerNamePositions)
    const numColumns = Math.max(...headerNamePositions) - firstColumnPosition + 1
    const narrowedRange = range.getSheet().getRange(topDataRowPosition, firstColumnPosition, numRows, numColumns)
    const narrowedRangeHeaderNames = getRangeHeaderNames(narrowedRange)
    let narrowedRangeValues = narrowedRange.getValues()

    if (numRows !== newValuesToApply.length) {
      throw new Error("Values array length does not match the number of range rows")
    }
    
    // Update the array of arrays with the new values
    narrowedRangeValues.forEach((sheetRow, sheetRowIndex) => {
      narrowedRangeHeaderNames.forEach((rangeHeaderName, rangeHeaderIndex) => {
        if (Object.keys(newValuesToApply[sheetRowIndex]).indexOf(rangeHeaderName) > -1) {
          sheetRow[rangeHeaderIndex] = newValuesToApply[sheetRowIndex][rangeHeaderName]
        }
      })
    })

    // Do the actual update
    return narrowedRange.setValues(narrowedRangeValues)
  } catch(e) { logError(e) }
}

function appendValuesByHeaderNames(values, sheet) {
  const sheetHeaderColumnNames = getSheetHeaderNames(sheet)
  values.forEach(row => {
    const rowArray = sheetHeaderColumnNames.map(colName => row[colName])
    sheet.appendRow(rowArray)
  })
}

// Cache header info in a global variable so it only needs to be collected once per sheet per onEdit call.
function getSheetHeaderNames(sheet, {forceRefresh = false, headerRowPosition = 1} = {}) {
  const sheetName = sheet.getName()
  if (!cachedHeaderNames[sheetName] || forceRefresh) {
    const headerNames = sheet.getRange("A" + headerRowPosition + ":" + headerRowPosition).getValues()[0]
    cachedHeaderNames[sheetName] = headerNames.map(headerName => !headerName ? " " : headerName)
  }
  return cachedHeaderNames[sheetName]
}
    
// Get header information for column range only, rather than the entire sheet
// Uses getSheetHeaderNames for caching purposes.
function getRangeHeaderNames(range, {forceRefresh = false, headerRowPosition = 1} = {}) {
  const sheetHeaderNames = getSheetHeaderNames(range.getSheet(), {forceRefresh: forceRefresh, headerRowPosition: headerRowPosition})
  const rangeStartColumnIndex = range.getColumn() - 1
  return sheetHeaderNames.slice(rangeStartColumnIndex, rangeStartColumnIndex + range.getWidth())
}

function getMaxValueInRange(range) {
  let values = range.getValues().flat().filter(Number.isFinite)
  return values.reduce((a, b) => Math.max(a, b))
}

function applyFormats(formatGroups, sheet) {
  try {
    Object.keys(formatGroups).forEach(groupName => {
      const ranges = formatGroups[groupName].ranges
      if (ranges.length) formatGroups[groupName].formats(sheet.getRangeList(ranges))
    })
  } catch(e) { logError(e) }
}