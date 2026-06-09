/**
 * @fileoverview Core spreadsheet helper library for RideSheet.
 *
 * Provides utilities for reading and writing data to Google Sheets using
 * column header names as keys, with a shared cache layer that avoids redundant
 * API calls within a single script execution. Key capabilities:
 *
 * - `getRangeValuesAsTable()` — converts a range to an array of row objects.
 * - `setValuesByHeaderNames()` — writes row objects back to a range with
 *   minimal API calls via row/column narrowing.
 * - `getSheetHeaderNames()` / `getRangeHeaderNames()` — cached header lookup.
 * - `createRows()` / `createRow()` — append or overwrite rows with header alignment.
 * - `safelyDeleteRows()` / `safelyDeleteRow()` — delete rows without row-shift bugs.
 * - `applySheetFormatsAndValidation()` — applies column number formats and data
 *   validation rules defined in `getConfiguredColumns()` (constants.js).
 */

/**
 * Module-level cache for sheet header names, keyed by sheet name.
 * Populated by `getSheetHeaderNames()` and used by all header-name lookup functions.
 * @type {Object.<string, string[]>}
 */
var cachedHeaderNames = {}
/**
 * Module-level cache for sheet header formulas, keyed by sheet name.
 * Populated alongside `cachedHeaderNames` by `getSheetHeaderNames()`.
 * Non-formula cells are stored as empty strings `""`.
 * @type {Object.<string, string[]>}
 */
var cachedHeaderFormulas = {}

/**
 * Test whether a range is fully inside or matches another range.
 * If the inner range is not fully inside the outer range, returns false.
 * @param {GoogleAppsScript.Spreadsheet.Range} innerRange The inner range
 * @param {GoogleAppsScript.Spreadsheet.Range} outerRange The range that the inner must be inside of or match exactly
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
 * @param {GoogleAppsScript.Spreadsheet.Range} firstRange The first range
 * @param {GoogleAppsScript.Spreadsheet.Range} secondRange The second range
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
 * @param {GoogleAppsScript.Spreadsheet.Range} range The source range
 * @return {GoogleAppsScript.Spreadsheet.Range}
 */
function getFullRow(range) {
  try {
    const rowPosition = range.getRow()
    return range.getSheet().getRange("A" + rowPosition + ":" + rowPosition)
  } catch(e) { logError(e) }
}

/**
 * Given a range, return the full width of all the rows that correspond with the passed in range.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The source range
 * @return {GoogleAppsScript.Spreadsheet.Range}
 */
function getFullRows(range) {
  try {
    return range.getSheet().getRange("A" + range.getRow() + ":" + range.getLastRow())
  } catch(e) { logError(e) }
}

/**
 * Searches all data rows in a sheet and returns the first row object that
 * satisfies the provided filter function. Returns the matching row as an object
 * in the same format as `getRangeValuesAsTable()`, or `undefined` if no row matches.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to search.
 * @param {function(Object): boolean} filter - A function that receives a row
 *   object and returns `true` for the desired row.
 * @returns {Object|undefined} The first matching row object, or `undefined` if none found.
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

/**
 * Writes multiple data rows to a destination sheet, aligning values to column
 * headers by name, and applies sheet formatting and data validation afterward.
 *
 * If any key in the source data does not match a destination column header (and
 * does not start with `"_"`), an alert is shown and no rows are written.
 * When `timestampColName` is provided, that column is set to the current date/time
 * for every row regardless of the source data value.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destSheet - The sheet to write rows into.
 * @param {Object[]} data - Array of row objects. Keys must match destination column headers.
 *   Keys starting with `"_"` (private metadata fields) are ignored.
 * @param {string|null} timestampColName - Column name to fill with the current timestamp,
 *   or `null` to skip timestamp injection.
 * @param {boolean} [overwrite=false] - When `true`, data is written starting at row 2,
 *   overwriting existing content. When `false`, rows are appended after the last used row.
 * @returns {boolean} `true` if rows were written successfully, `false` otherwise.
 */
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
 * Applies formatting and validation rules to a sheet based on column definitions
 * from `getConfiguredColumns()` (constants.js). Sets font weight to normal on all
 * data rows, then applies number formats and data validation rules per column.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to apply the rules to.
 * @param {number} [startRow=2] - The 1-based row from which formatting is applied downward.
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

/**
 * Appends a single data row to a sheet by mapping object keys to column headers.
 * Uses `appendRow()`, so formatting and data validation are not applied automatically.
 * Call `applySheetFormatsAndValidation()` separately afterward if needed.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destSheet - The sheet to append to.
 * @param {Object} data - A row object whose keys match column headers.
 * @returns {boolean} `true` if the row was appended successfully, `false` otherwise.
 */
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

/**
 * Deletes multiple rows from a sheet in a single batch using the Advanced Sheets
 * API v4 (`Sheets.Spreadsheets.batchUpdate`), processing deletions in reverse
 * order to avoid row-index shifting. Before deleting, ensures the sheet retains
 * at least one row beyond the last data row so the sheet is never left empty.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to delete rows from.
 * @param {Object[]} data - Array of row objects as returned by `getRangeValuesAsTable()`.
 *   Each object must have a `_rowIndex` property (0-based index within the queried range).
 */
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

/**
 * Deletes a single row from a sheet using `sheet.deleteRow()`. Before deleting,
 * ensures the sheet retains at least one row beyond the last data row so the
 * sheet is never left empty.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to delete the row from.
 * @param {Object} row - A row object as returned by `getRangeValuesAsTable()`.
 *   Must have a `_rowPosition` property (1-based spreadsheet row number).
 */
function safelyDeleteRow(sheet, row) {
  const lastRowPosition = sheet.getLastRow()
  if (sheet.getMaxRows() === lastRowPosition) {
    sheet.insertRowAfter(lastRowPosition)
  }
  sheet.deleteRow(row._rowPosition)
}

/**
 * Default filter function used by `createColumns()` to select which keys from
 * a data row should become sheet columns. Excludes blank header names, headers
 * starting with `"_"` (private fields), and the reserved columns `"Action"`,
 * `"Go"`, `"Earliest PU Time"`, and `"Latest PU Time"`.
 * @param {string} colHeader - A column header name to test.
 * @returns {boolean} `true` if the column should be created; `false` to skip it.
 */
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

/**
 * Adds any missing columns to a sheet based on the keys present in a sample data row.
 * Only keys accepted by the `columnFilter` function and not already present in the
 * sheet's header row are added. New column headers are inserted just before the last
 * `colOffset` columns from the right edge of the sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to add columns to.
 * @param {Object} dataRow - A representative data row; its keys indicate desired column names.
 * @param {function(string): boolean} [columnFilter=defaultColumnFilter] - Filter function
 *   controlling which keys become new columns.
 * @param {number} [colOffset=0] - Number of columns from the right end to leave in place;
 *   new columns are inserted before the last `colOffset` columns.
 */
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

/**
 * Development debug helper. The `fixRowNumberFormatting` and `fixRowDataValidation`
 * helpers it references are no longer implemented.
 * @private
 */
function testRowFormat() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName('Trip Review')
  let newRowIndex = sheet.getLastRow()
  let newRow = sheet.getRange(newRowIndex + ':' + newRowIndex)
  let a = newRow.getA1Notation()
  fixRowNumberFormatting(newRow)
  fixRowDataValidation(newRow)
}

/**
 * Converts a sheet range into an array of row objects keyed by column header name.
 *
 * If the range includes the header row, the header row is consumed from the front
 * and used as keys; otherwise the header row is fetched separately via
 * `getRangeHeaderNames()`. Returns `[]` if the range contains only the header row.
 *
 * Each returned row object includes two private metadata fields added by this function:
 * - `_rowPosition` {number} — The 1-based spreadsheet row number of the row.
 * - `_rowIndex` {number}   — The 0-based index of the row within the returned array.
 *
 * When `includeFormulaValues` is `false`, cells whose values are derived from an
 * in-cell formula or from an array formula anchored in the header row are excluded
 * from the row object. Additionally, columns whose header name begins with `"|"` (pipe)
 * are always excluded when `includeFormulaValues` is `false`. These exclusions prevent
 * formula-derived values from being treated as editable data.
 *
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to convert.
 * @param {Object} [options]
 * @param {number} [options.headerRowPosition=1] - The 1-based row number of the header row.
 * @param {boolean} [options.includeFormulaValues=true] - When `false`, cells computed by
 *   formulas, array-formula headers, or `|`-prefixed headers are omitted from row objects.
 * @returns {Object[]} Array of row objects mapping header names to cell values, each with
 *   `_rowPosition` and `_rowIndex` metadata fields. Returns `[]` if the range is empty
 *   or contains only the header row.
 */
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
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range
 * @return {*}
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
 * @param {GoogleAppsScript.Spreadsheet.Range} range The range
 * @return {*}
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

/**
 * Writes an array of row objects back to a range, using column header names as keys.
 *
 * Only the keys present in each row object are updated; all other cell values in the
 * range are preserved. When `overwriteAll` is `false` (the default), the write is
 * narrowed to the minimal rectangular sub-range covering only the rows and columns
 * that actually have data, reducing the number of API calls required. Empty row
 * objects (`{}`) are skipped during this narrowing.
 *
 * Cells whose values come from in-cell formulas are preserved (the formula string is
 * written back). Cells under array-formula headers or under `|`-prefixed header names
 * are blanked rather than overwritten to avoid corrupting array formula output.
 *
 * @param {Object[]} newValues - Array of row objects mapping header names to new values.
 *   Empty objects `{}` are treated as "no change" rows and skipped when narrowing.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The target range to update.
 * @param {Object} [options]
 * @param {number} [options.headerRowPosition=1] - The 1-based row number of the header row.
 * @param {boolean} [options.overwriteAll=false] - When `true`, writes all values across
 *   the full range without narrowing.
 * @returns {GoogleAppsScript.Spreadsheet.Range} The original `range` (enables chaining).
 */
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

/**
 * Appends multiple rows to a sheet by mapping object keys to column headers,
 * applying row formatting to each appended row via `appendRowWithFormatting()`.
 * @param {Object[]} values - Array of row objects whose keys match column headers.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to append rows to.
 */
function appendValuesByHeaderNames(values, sheet) {
  try {
    const sheetHeaderColumnNames = getSheetHeaderNames(sheet)
    values.forEach(row => {
      const rowArray = sheetHeaderColumnNames.map(colName => row[colName])
      appendRowWithFormatting(sheet, rowArray)
    })
  } catch(e) { logError(e) }
}

/**
 * Returns the header names for an entire sheet as an array of strings, one per column.
 * Results are cached in the module-level `cachedHeaderNames` object (keyed by sheet name)
 * so the header row is read only once per script execution unless `forceRefresh` is `true`.
 * Blank header cells are stored as `" "` (a single space) to preserve column alignment.
 * Also populates `cachedHeaderFormulas` with the header row's formula strings.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to read headers from.
 * @param {Object} [options]
 * @param {boolean} [options.forceRefresh=false] - When `true`, bypasses the cache and
 *   re-reads headers from the spreadsheet.
 * @param {number} [options.headerRowPosition=1] - The 1-based row number of the header row.
 * @returns {string[]} Array of header name strings, one per column.
 */
function getSheetHeaderNames(sheet, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetName = sheet.getName()
    if (!cachedHeaderNames[sheetName] || forceRefresh) {
      const headerRange = sheet.getRange("A" + headerRowPosition + ":" + headerRowPosition)
      const headerNames = headerRange.getValues()[0]
      const headerFormulas = headerRange.getFormulas()[0]
      cachedHeaderNames[sheetName] = headerNames.map(headerName => !headerName ? " " : headerName)
      cachedHeaderFormulas[sheetName] = headerFormulas
    }
    return cachedHeaderNames[sheetName]
  } catch(e) { logError(e) }
}

/**
 * Convenience wrapper around `setValuesByHeaderNames()` that updates a single
 * sheet row identified by its row number.
 * @param {Object} newValues - A row object mapping header names to new values.
 * @param {number} rowNumber - The 1-based row number to update.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet containing the row.
 * @param {Object} [options]
 * @param {number} [options.headerRowPosition=1] - The 1-based row number of the header row.
 * @param {boolean} [options.overwriteAll=false] - Passed through to `setValuesByHeaderNames()`.
 * @returns {GoogleAppsScript.Spreadsheet.Range} The updated range.
 */
function setValuesForRow(newValues, rowNumber, sheet, {headerRowPosition = 1, overwriteAll = false} = {}) {
  try {
    const range = sheet.getRange(rowNumber + ":" + rowNumber)
    return setValuesByHeaderNames([newValues], range, {headerRowPosition: headerRowPosition, overwriteAll: overwriteAll})
  } catch(e) { logError(e) }
}

/**
 * Returns header names for only the columns spanned by `range`, sliced from
 * the full sheet header array returned by `getSheetHeaderNames()`.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range whose column span determines
 *   which header names are returned.
 * @param {Object} [options]
 * @param {boolean} [options.forceRefresh=false] - Passed through to `getSheetHeaderNames()`.
 * @param {number} [options.headerRowPosition=1] - The 1-based row number of the header row.
 * @returns {string[]} Array of header name strings for the columns covered by `range`.
 */
function getRangeHeaderNames(range, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetHeaderNames = getSheetHeaderNames(range.getSheet(), {forceRefresh: forceRefresh, headerRowPosition: headerRowPosition})
    const rangeStartColumnIndex = range.getColumn() - 1
    return sheetHeaderNames.slice(rangeStartColumnIndex, rangeStartColumnIndex + range.getWidth())
  } catch(e) { logError(e) }
}

/**
 * Returns the header-row formula strings for an entire sheet as an array, one per column.
 * Results are cached alongside header names by `getSheetHeaderNames()`, so no extra
 * API calls are made. Non-formula cells are represented as empty strings `""`.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to read header formulas from.
 * @param {Object} [options]
 * @param {boolean} [options.forceRefresh=false] - Passed through to `getSheetHeaderNames()`.
 * @param {number} [options.headerRowPosition=1] - The 1-based row number of the header row.
 * @returns {string[]} Array of formula strings, one per column. Empty string for non-formula cells.
 */
function getSheetHeaderFormulas(sheet, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetName = sheet.getName()
    getSheetHeaderNames(sheet, {forceRefresh: forceRefresh, headerRowPosition: headerRowPosition})
    return cachedHeaderFormulas[sheetName]
  } catch(e) { logError(e) }
}

/**
 * Returns header-row formula strings for only the columns spanned by `range`,
 * sliced from the full sheet header formula array returned by `getSheetHeaderFormulas()`.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range whose column span determines
 *   which header formulas are returned.
 * @param {Object} [options]
 * @param {boolean} [options.forceRefresh=false] - Passed through to `getSheetHeaderFormulas()`.
 * @param {number} [options.headerRowPosition=1] - The 1-based row number of the header row.
 * @returns {string[]} Array of formula strings for the columns in `range`.
 */
function getRangeHeaderFormulas(range, {forceRefresh = false, headerRowPosition = 1} = {}) {
  try {
    const sheetHeaderFormulas = getSheetHeaderFormulas(range.getSheet(), {forceRefresh: forceRefresh, headerRowPosition: headerRowPosition})
    const rangeStartColumnIndex = range.getColumn() - 1
    return sheetHeaderFormulas.slice(rangeStartColumnIndex, rangeStartColumnIndex + range.getWidth())
  } catch(e) { logError(e) }
}

/**
 * Returns the maximum finite numeric value found anywhere in the given range,
 * or `null` if the range contains no finite numbers.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The range to search.
 * @returns {number|null} The maximum finite number, or `null`.
 */
function getMaxValueInRange(range) {
  try {
    let values = range.getValues().flat().filter(Number.isFinite)
    if (!values.length) return null
    return values.reduce((a, b) => Math.max(a, b))
  } catch(e) { logError(e) }
}

/**
 * Applies a collection of named format groups to a sheet. Each group provides
 * an array of A1-notation range strings and a formatting function to run against them.
 * @param {Object.<string, {ranges: string[], formats: function(GoogleAppsScript.Spreadsheet.RangeList): void}>} formatGroups
 *   Map of group name → `{ranges, formats}`. `ranges` is an array of A1-notation strings;
 *   `formats` receives the corresponding `RangeList` and applies formatting.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to apply formats to.
 */
function applyFormats(formatGroups, sheet) {
  try {
    Object.keys(formatGroups).forEach(groupName => {
      const ranges = formatGroups[groupName].ranges
      if (ranges.length) formatGroups[groupName].formats(sheet.getRangeList(ranges))
    })
  } catch(e) { logError(e) }
}

/**
 * Converts a 1-based column position to its A1-notation column letter string.
 * Examples: `1` → `"A"`, `26` → `"Z"`, `27` → `"AA"`, `702` → `"ZZ"`.
 * @param {number} colPosition - The 1-based column position.
 * @returns {string} The column letter string in A1 notation.
 */
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

/**
 * Returns `true` if `value` is a zero-length string `""` or `null`.
 * Numeric `0` and boolean `false` are not considered blank.
 * @param {*} value - The cell value to test.
 * @returns {boolean}
 */
function isBlankCell(value) {
  return (value === "" || value === null)
}

/**
 * Clears all data rows from a sheet, leaving only the header row (row 1).
 * Rows 3 and beyond are deleted; row 2's content is cleared. Does nothing if
 * the sheet has only the header row.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to clear.
 */
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

/**
 * Clears out hand-entered values that block "array spill" formulas located in headers.
 *
 * This function is intended to be called from an onEdit trigger. It inspects the
 * header row for spill-formula headers that evaluate to "#REF!" and match the
 * expected array-literal pattern (e.g. formulas starting with `={"`). For any
 * such header, it calculates the spill range underneath the header and clears
 * any blocking values in that range so the array spill formula can recalculate.
 * It then refreshes cached header names and notifies the user via a toast.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The edit event object describing
 *     the user edit that triggered this handler, including the edited range.
 * @returns {void}
 */
function clearSpillBlockages(e) {
  try {
    const sheet = e.range.getSheet()
    const headerValues = getSheetHeaderNames(sheet)
    const headerFormulas = getSheetHeaderFormulas(sheet)
    let blockagesCleared = 0

    headerValues.forEach((headerValue, i) => {
      if (headerValue === "#REF!" && headerFormulas[i]?.startsWith(`={"`)) {
        const spillColumnCount = getSpillColumnCount(headerFormulas[i])
        const spillRowCount = sheet.getLastRow() - 1
        if (spillColumnCount > 0 && spillRowCount > 0) {
          const rangeToClear = sheet.getRange(2,i+1,spillRowCount,spillColumnCount)
          rangeToClear.clearContent()
          blockagesCleared++
        }
      }
    })

    if (blockagesCleared) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Data blocking ${blockagesCleared === 1 ? "a" : blockagesCleared} calculated column${blockagesCleared === 1 ? "" : "s"} has been cleared.`
      )
      // Refresh the header name and formula caches for any calls after this that rely on it.
      getSheetHeaderNames(sheet,{forceRefresh: true})
    }
  } catch(e) { logError(e) }
}

/**
 * Parses a Google Sheets array-literal "spill" formula to determine how many
 * columns the first row of the spilled array will occupy.
 *
 * The function expects a formula string that contains an array literal,
 * typically of the form:
 *   ={"col1","col2","col3"; ... }
 * It inspects only the first row of the array and counts top-level commas
 * that are not inside quotes, parentheses, or nested array literals.
 *
 * @param {string} formula The full formula string containing an array
 *   literal whose first row's width (number of columns) should be computed.
 * @return {number} The number of columns in the first row of the spilled
 *   array; returns 0 if no valid array literal start or row separator is found.
 */
function getSpillColumnCount(formula) {
  try {
    if (typeof formula !== 'string' || formula.length === 0) return -1
    const start = formula.indexOf('{')
    if (start === -1) return 0
    let endOfRowPos = findTopLevelSemicolon(formula, start)
    if (endOfRowPos === -1) endOfRowPos = formula.lastIndexOf('}')
    if (endOfRowPos === -1 || endOfRowPos < start) return 0

    const firstRow = formula.substring(start + 1, endOfRowPos)
    let commas = 0
    let inQuote = false
    let parenDepth = 0
    let braceDepth = 0

    for (let i = 0; i < firstRow.length; i++) {
      const char = firstRow[i]
      const nextChar = firstRow[i + 1]

      if (inQuote) {
        // Check for escaped quote (double quote)
        if (char === '"' && nextChar === '"') {
          i++; // skip next quote
        } else if (char === '"') {
          inQuote = false
        }
      } else {
        if (char === '"') {
          inQuote = true
        } else if (char === '(') {
          parenDepth++
        } else if (char === ')') {
          parenDepth--
        } else if (char === '{') {
          braceDepth++
        } else if (char === '}') {
          braceDepth--
        } else if (char === ',' && parenDepth === 0 && braceDepth === 0) {
          commas++
        }
      }
    }
    return commas + 1
  } catch(e) {
    logError(e)
    return -1
  }
}

/**
 * Finds the first semicolon at the top level of nesting in a formula string.
 * Semicolons that appear inside quotes, parentheses, or braces are ignored.
 *
 * @param {string} formula The formula text to search.
 * @param {number} startPos The index from which to start scanning (typically the index of '{').
 * @return {number} The index of the first top-level semicolon, or -1 if none is found.
 */
function findTopLevelSemicolon(formula, startPos) {
  try {
    if (typeof formula !== 'string' || formula.length === 0) return -1
    let inQuote = false
    let parenDepth = 0
    let braceDepth = 0

    for (let i = startPos + 1; i < formula.length; i++) {
      const char = formula[i]
      const nextChar = formula[i + 1]

      if (inQuote) {
        if (char === '"' && nextChar === '"') {
          i++
        } else if (char === '"') {
          inQuote = false
        }
      } else {
        if (char === '"') {
          inQuote = true
        } else if (char === '(') {
          parenDepth++
        } else if (char === ')') {
          parenDepth--
        } else if (char === '{') {
          braceDepth++
        } else if (char === '}') {
          braceDepth--
        } else if (char === ';' && parenDepth === 0 && braceDepth === 0) {
          return i
        }
      }
    }
    return -1
  } catch(e) {
    logError(e)
    return -1
  }
}
