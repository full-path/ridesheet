var cachedHeaderNames = {}
var cachedHeaderFormulas = {}

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

function createRows(destSheet, data, timestampColName, overwrite=false) {
  try {
    const timestamp = new Date()
    let destColumnNames = getSheetHeaderNames(destSheet)
    let sourceColumnNames = Object.keys(data[0])
    let missingDestColumns = sourceColumnNames.reduce((a, c) => {
      if (!destColumnNames.includes(c) && c.slice(0,1) !== "_") a.push(c)
      return a
    }, [])
    if (missingDestColumns.length) {
      SpreadsheetApp.getUi().alert(
        `Sheet "${destSheet.getSheetName()}" is missing the column${missingDestColumns.length === 1 ? "" : "s"} ${missingDestColumns.map((e) => '"' + e + '"').join(", ")}.
        Rows will not be moved to the "${destSheet.getSheetName()}" sheet.`)
      return false
    }
    let values = data.map(row => {
      return destColumnNames.map(colName => {
        if (timestampColName && colName === timestampColName) {
          return timestamp
        } else {
          return isBlankCell(row[colName]) ? null: row[colName]
        }
      })
    })
    let firstRow = overwrite ? 2 : destSheet.getLastRow() + 1
    let newRows = destSheet.getRange(firstRow, 1, values.length, values[0].length)
    newRows.setValues(values)
    applySheetFormatsAndValidation(destSheet, firstRow)
    return true
  } catch(e) {
      logError(e)
      return false
  }
}

// ex. of format for getConfiguredColumns()
// {"Trips": {
//     "Trip Date": {
//       numberFormat: "M/d/yyyy",
//       dataValidation: {
//         criteriaType: "DATE_IS_VALID_DATE",
//         helpText: "Value must be a valid date.",
//       },
//     },
//     "Customer Name and ID": {
//       dataValidation: {
//         criteriaType: "VALUE_IN_RANGE",
//         namedRange: "lookupCustomerNames",
//         showDropdown: true,
//         allowInvalid: false,
//         helpText: "Value must be a valid customer name and ID.",
//       },
//     },
// }}

/**
 * Applies formatting and validation rules to a sheet based on a set of default column rules.
 * @param {Sheet} sheet - The sheet to apply the rules to.
 * @param {number} [startRow=2] - The starting row of the range to apply the rules to.
 */
function applySheetFormatsAndValidation(sheet, startRow=2) {
  let sheetName = sheet.getName()
  let rules = getConfiguredColumns()[sheetName]
  let configuredHeaderNames = Object.keys(rules)

  // Get the headers of the sheet
  let headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn())
  let sheetHeaders = headerRange.getValues()[0]

  // Get the range of rows beginning with startRow and ending at the last row in the sheet
  // Set formatting on that range to ensure text is normal weight (not bold)
  let dataRange = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, sheet.getLastColumn())
  dataRange.setFontWeight('normal')

  // Loop through configuredHeaderNames and apply formatting and validation rules as appropriate
  for (let i = 0; i < configuredHeaderNames.length; i++) {
    let headerName = configuredHeaderNames[i]
    let rule = rules[headerName]
    if (rule.numberFormat || rule.dataValidation) {
      let index = sheetHeaders.indexOf(headerName)
      if (index >= 0) {
        let columnRange = sheet.getRange(startRow, index + 1, sheet.getLastRow() - startRow + 1)
        if (rule.numberFormat) {
          columnRange.setNumberFormat(rule.numberFormat)
        }
        if (rule.dataValidation) {
          let ruleAttributes = rule.dataValidation
          let validationRule = getValidationRule(ruleAttributes)
          columnRange.setDataValidation(validationRule)
        }
      }
    }
  }
}

// This function no longer handles formatting. If using, it is recommended
// to call applySheetFormatsAndValidation after creating new row(s)
function createRow(destSheet, data) {
  try {
    let columnNames = getSheetHeaderNames(destSheet)
    let dataArray = columnNames.map(colName => data[colName] ? data[colName] : null)
    destSheet.appendRow(dataArray)
    //let newRowIndex = destSheet.getLastRow()
    // These row based formatting errors are broken; leaving them here as a reminder
    // fixRowNumberFormatting(newRow)
    // fixRowDataValidation(newRow)
    return true
  } catch(e) {
    logError(e)
    return false
  }
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
  Sheets.Spreadsheets.batchUpdate({requests: rowsToDelete}, ss.getId())
}

function safelyDeleteRow(sheet, row) {
  const lastRowPosition = sheet.getLastRow()
  if (sheet.getMaxRows() === lastRowPosition) {
    sheet.insertRowAfter(lastRowPosition)
  }
  sheet.deleteRow(row._rowPosition)
}

const defaultColumnFilter = colHeader => {
  const colsToSkip = ["Action", "Go", "Earliest PU Time", "Latest PU Time"]
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

function testRowFormat() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName('Trip Review')
  let newRowIndex = sheet.getLastRow()
  let newRow = sheet.getRange(newRowIndex + ':' + newRowIndex)
  let a = newRow.getA1Notation()
  fixRowNumberFormatting(newRow)
  fixRowDataValidation(newRow)
}

// Takes a range and returns an array of objects, each object containing key/value pairs.
// If the range includes row 1 of the spreadsheet, that top row will be used as the keys.
// Otherwise row 1 will be collected separately and used as the source for keys.
// If includeFormulaValues = false, then formulas are searched for in the actual source cell
// and in the corresponding header cell. The latter check is premised on the idea that any
// array formulas will be embedded in the header cell
function getRangeValuesAsTable(range, {headerRowPosition = 1, includeFormulaValues = true} = {}) {
  try {
    let topDataRowPosition = range.getRow()
    let values = range.getValues()
    let formulas
    let rangeHeaderNames
    let rangeHeaderFormulas
    if (!includeFormulaValues) formulas = range.getFormulas()
    if (topDataRowPosition <= headerRowPosition) {
      if (values.length > (headerRowPosition + 1 - topDataRowPosition)) {
        // If the header row is already in the selected range, then collect the header names
        // and remove them from the values and formulas arrays
        rangeHeaderNames = values[headerRowPosition - topDataRowPosition]
        values.splice(0, headerRowPosition + 1 - topDataRowPosition)
        if (!includeFormulaValues) {
          rangeHeaderFormulas = formulas[headerRowPosition - topDataRowPosition]
          formulas.splice(0, headerRowPosition + 1 - topDataRowPosition)
        }
        topDataRowPosition = headerRowPosition + 1
      } else {
        return []
      }
    } else if (topDataRowPosition > headerRowPosition) {
      rangeHeaderNames = getRangeHeaderNames(range, {headerRowPosition: headerRowPosition})
      rangeHeaderFormulas = getRangeHeaderFormulas(range, {headerRowPosition: headerRowPosition})
    }
    let result = values.map((row, rowIndex) => {
      let rowObject = {}
      rowObject._rowPosition = rowIndex + topDataRowPosition
      rowObject._rowIndex = rowIndex
      rangeHeaderNames.forEach((headerName, columnIndex) => {
        if (
          includeFormulaValues ||
          (
            !includeFormulaValues &&
            !formulas[rowIndex][columnIndex] &&
            !rangeHeaderFormulas[columnIndex] &&
            headerName[0] !== "|"
          )
        ) {
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
    let narrowedRangeFormulas
    let narrowedHeaderFormulas
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
      narrowedRangeHeaderNames = getRangeHeaderNames(narrowedRange, {headerRowPosition: headerRowPosition})
      narrowedRangeValues = narrowedRange.getValues()
      narrowedRangeFormulas = narrowedRange.getFormulas()

      // Remove values derived from in-cell formulas or array formulas placed in the header row.
      // Otherwise, they'll get put in as literal values that will break the array formula.
      // There's no way to easily discern when a two-dimensional array formula is being used for
      // columns not directly under the source formula, so as a workaround, the function will
      // also check to see if the "|" (pipe) character is the first character of a header value.
      narrowedHeaderFormulas = getRangeHeaderFormulas(narrowedRange, {headerRowPosition: headerRowPosition})
      if (narrowedHeaderFormulas.some((formula) => formula !== "")) {
        const narrowedRangeValuesWithoutFormulaValues = narrowedRangeValues.map((row, rowIndex) => {
          return row.map((value, columnIndex) => {
            if (
              narrowedRangeFormulas[rowIndex][columnIndex]
            ) {
              return narrowedRangeFormulas[rowIndex][columnIndex]
            } else if (
              narrowedHeaderFormulas[columnIndex] ||
              narrowedRangeHeaderNames[columnIndex][0] === "|"
            ) {
              return ""
            } else {
              return value
            }
          })
        })
        narrowedRangeValues = narrowedRangeValuesWithoutFormulaValues
      }
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

// Cache header formula info in a global variable so it only needs to be collected once per sheet per onEdit call.
function getSheetHeaderFormulas(sheet, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetName = sheet.getName()
    if (!cachedHeaderFormulas[sheetName] || forceRefresh) {
      const headerFormulas = sheet.getRange("A" + headerRowPosition + ":" + headerRowPosition).getFormulas()[0]
      cachedHeaderFormulas[sheetName] = headerFormulas
    }
    return cachedHeaderFormulas[sheetName]
  } catch(e) { logError(e) }
}

// Get header formula information for column range only, rather than the entire sheet
// Uses getSheetHeaderFormulas for caching purposes.
function getRangeHeaderFormulas(range, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetHeaderFormulas = getSheetHeaderFormulas(range.getSheet(), {forceRefresh: forceRefresh, headerRowPosition: headerRowPosition})
    const rangeStartColumnIndex = range.getColumn() - 1
    return sheetHeaderFormulas.slice(rangeStartColumnIndex, rangeStartColumnIndex + range.getWidth())
  } catch(e) { logError(e) }
}

function getMaxValueInRange(range) {
  try {
    let values = range.getValues().flat().filter(Number.isFinite)
    if (!values.length) return null
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

// Null and zero-length strings are considered blank values.
// Numeric "0" and boolean "false" are not blank.
function isBlankCell(value) {
  return (value === "" || value === null)
}

function clearSheet(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const numRows = lastRow - 2;
    if (numRows > 0) {
      sheet.deleteRows(3, numRows);
    }
    const dataRange = sheet.getRange(2, 1, 1, sheet.getLastColumn());
    dataRange.clearContent();
  }
}
