/**
 * @fileoverview General-purpose utility functions for RideSheet.
 *
 * Covers:
 * - Debug logging to the "Debug Log" sheet
 * - Drive backup creation and rotation
 * - Date/time helpers (formatting, arithmetic, parsing, extraction)
 * - String utilities (regex escaping, template substitution, pluralization)
 * - Type detection
 * - Miscellaneous spreadsheet helpers
 *
 * These functions have no dependencies on other RideSheet files except
 * `getDocProp()` from properties.js (used for time zone and log level).
 */

/**
 * Appends a row to the "Debug Log" sheet with a timestamp and the given values.
 * Does nothing when the `logLevel` document property is set to `"normal"`;
 * only logs when set to `"verbose"`.
 * @param {*} d - Primary value to log.
 * @param {...*} args - Additional values appended as extra columns in the log row.
 */
function log(d, ...args) {
  if (getDocProp("logLevel") === "normal") return
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let logSheet = ss.getSheetByName("Debug Log") || ss.insertSheet("Debug Log")
  logSheet.appendRow([new Date(), d].concat(args))
}

/**
 * Logs a caught error's name, message, and stack trace to the Debug Log sheet.
 * Intended to be called from catch blocks throughout the codebase.
 * @param {Error} e - The caught error object.
 */
function logError(e) {
  log(e.name + ': ' + e.message, e.stack)
}

/**
 * Clears all data rows from the "Debug Log" sheet, leaving the header row intact.
 */
function clearLog() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let logSheet = ss.getSheetByName("Debug Log")
  logSheet.deleteRows(2,logSheet.getLastRow())
}

/**
 * Logs all raw document property key/value pairs to the Debug Log sheet.
 * Useful for inspecting serialized property values during development.
 */
function logProperties() {
  let docProps = PropertiesService.getDocumentProperties()
  docProps.getKeys().forEach(prop => {
    log(prop,docProps.getProperty(prop))
  })
}

/**
 * Creates a dated copy of the active spreadsheet in the specified Drive folder.
 * The copy is named `"YYYY-MM-DD_<spreadsheet name>"`.
 * @deprecated This function is not currently called anywhere in the codebase.
 *   It is preserved as a utility for custom or future use.
 * @param {string} destFolderId - The ID of the destination Google Drive folder.
 * @returns {boolean} `true` if the backup was created successfully, `false` otherwise.
 */
function makeBackup(destFolderId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const file = DriveApp.getFileById(ss.getId())
    const newName = formatDate(null,null,"yyyy-MM-dd") + "_" + ss.getName()
    const destFolder = DriveApp.getFolderById(destFolderId)
    file.makeCopy(newName, destFolder)
    return true
  } catch(e) {
    return false
  }
}

/**
 * Moves files in a Drive folder to the trash if they are older than the retention period.
 * Intended to be used alongside `makeBackup()` to limit backup storage usage.
 * @deprecated This function is not currently called anywhere in the codebase.
 *   It is preserved as a utility for custom or future use.
 * @param {string} folderId - The ID of the Google Drive folder to scan.
 * @param {number} retentionInDays - Files older than this many days will be trashed.
 */
function rotateBackups(folderId, retentionInDays) {
  let file
  let ageInDays
  for (const files = DriveApp.getFolderById(folderId).getFiles(); files.hasNext(); file = files.next()) {
    ageInDays = Math.ceil(((new Date()) - file.getDateCreated()) / (1000 * 60 * 60 * 24))
    if (ageInDays > retentionInDays) { file.setTrashed(true) }
  }
}

/**
 * Returns true if the given value is a valid, non-NaN JavaScript Date object.
 * @param {*} date - The value to check.
 * @returns {boolean}
 */
function isValidDate(date) {
  return date && Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date)
}

/**
 * Formats a date value as a string using the given time zone and format pattern.
 * Falls back to the `localTimeZone` document property and `"M/d/yyyy"` when
 * `timeZone` or `dateFormat` are omitted. Returns the formatted string for today
 * when `date` is omitted or falsy.
 * @param {Date|string|null} date - The date to format, or `null`/`undefined` for today.
 * @param {string|null} timeZone - An IANA time zone string (e.g. `"America/Los_Angeles"`),
 *   or `null` to use the `localTimeZone` document property.
 * @param {string|null} dateFormat - A `Utilities.formatDate()` pattern string,
 *   or `null` to use `"M/d/yyyy"`.
 * @returns {string|undefined} The formatted date string, or `undefined` if
 *   `date` is provided but cannot be parsed.
 */
function formatDate(date, timeZone, dateFormat) {
  if (!timeZone) timeZone = getDocProp("localTimeZone")
  if (!dateFormat) dateFormat = "M/d/yyyy"
  if (!date) {
    return Utilities.formatDate(new Date(), timeZone, dateFormat)
  } else if (isValidDate(date)) {
    return Utilities.formatDate(date, timeZone, dateFormat)
  } else {
    const thisDate = new Date(date)
    if (isValidDate(thisDate)) return Utilities.formatDate(thisDate, timeZone, dateFormat)
  }
}

/**
 * Returns a new Date offset from the given date by the specified number of days.
 * @param {Date|null} date - The starting date, or `null`/`undefined` to use today.
 * @param {number} days - Number of days to add (use a negative value to subtract).
 * @returns {Date}
 */
function dateAdd(date, days) {
  if (!date) date = new Date()
  let result = new Date(date)
  result.setDate(result.getDate() + days)
  return result
}

/**
 * Returns a new Date offset from the given date by the specified number of milliseconds.
 * @param {Date|null} date - The starting date, or `null`/`undefined` to use now.
 * @param {number} milliseconds - Number of milliseconds to add (use a negative value to subtract).
 * @returns {Date}
 */
function timeAdd(date, milliseconds) {
  if (!date) date = new Date()
  return new Date(date.getTime() + milliseconds)
}

/**
 * Returns a Date with the time portion zeroed out (midnight local time).
 * @param {Date|string|null} dateTime - The source date/time, or `null`/`undefined` for today.
 * @returns {Date}
 */
function dateOnly(dateTime) {
  try {
    let thisDateTime
    if (!dateTime) {
      thisDateTime = new Date()
    } else if (typeof dateTime === "string") {
      thisDateTime = new Date(dateTime)
    } else {
      thisDateTime = dateTime
    } 
    return new Date(thisDateTime.setHours(0,0,0,0))
  } catch(e) { logError(e) }
}

/**
 * Returns today's date with the time portion zeroed out (midnight local time).
 * @returns {Date}
 */
function dateToday() {
  try {
    return dateOnly()
  } catch(e) { logError(e) }
}

/**
 * Parses a value into a Date. Returns `alternateValue` if the input cannot be parsed.
 * @param {*} date - The value to parse (will be coerced to string before parsing).
 * @param {*} alternateValue - Value to return when parsing fails.
 * @returns {Date|*} The parsed Date, or `alternateValue` if parsing failed.
 */
function parseDate(date, alternateValue) {
  const dateVal = Date.parse(date.toString())
  return isNaN(dateVal) ? alternateValue : new Date(dateVal)
}

/**
 * Extracts the time-of-day portion of a date/time value as a number of milliseconds
 * since midnight. Useful for comparing times across different dates.
 * @param {Date|string|null} dateTime - The source date/time, or `null`/`undefined` for now.
 * @returns {number} Milliseconds elapsed since midnight.
 */
function timeOnlyAsMilliseconds(dateTime) {
  try {
    let thisDateTime
    if (!dateTime) {
      thisDateTime = new Date()
    } else if (typeof dateTime === "string") {
      thisDateTime = new Date(dateTime)
    } else {
      thisDateTime = dateTime
    }
    return thisDateTime.getHours() * 3600000 + thisDateTime.getMinutes() * 60000 + thisDateTime.getSeconds() * 1000 + thisDateTime.getMilliseconds()
  } catch(e) { logError(e) }
}

/**
 * Combines the date portion of one value with the time portion of another into a single Date.
 * @param {Date|string} date - Supplies the year, month, and day.
 * @param {Date|string} time - Supplies the hours, minutes, seconds, and milliseconds.
 * @returns {Date}
 */
function combineDateAndTime(date, time) {
  try {
    return new Date(dateOnly(date).getTime() + timeOnlyAsMilliseconds(time))
  } catch(e) { logError(e) }
}

/**
 * Escapes all special regular expression characters in a string so it can be
 * used safely as a literal pattern inside a `RegExp` constructor.
 * @param {string} string - The string to escape.
 * @returns {string} The escaped string.
 */
function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

/**
 * Returns a lowercase string identifying the JavaScript type of the given value.
 * More precise than `typeof` — distinguishes `"array"`, `"date"`, `"null"`,
 * `"map"`, `"set"`, etc. Unknown object classes fall back to `"string"`.
 * @param {*} value - The value to inspect.
 * @returns {string} One of: `"array"`, `"bigint"`, `"boolean"`, `"date"`, `"map"`,
 *   `"null"`, `"number"`, `"object"`, `"regexp"`, `"set"`, `"string"`,
 *   `"symbol"`, or `"undefined"`.
 */
function getType(value) {
  let objectClass = Object.prototype.toString.call(value)
  let classes = {
    "[object Array]":      "array",
    "[object BigInt]":     "bigint",
    "[object Boolean]":    "boolean",
    "[object Date]":       "date",
    "[object Map]":        "map",
    "[object Null]":       "null",
    "[object Number]":     "number",
    "[object Object]":     "object",
    "[object RegExp]":     "regexp",
    "[object Set]":        "set",
    "[object String]":     "string",
    "[object Symbol]":     "symbol",
    "[object Undefined]":  "undefined"
  }
  if (objectClass in classes) {
    return classes[objectClass]
  } else {
    return "string"
  }
}

/**
 * Replaces placeholder tokens in a template string with values from a data object.
 * Used in manifest generation (`manifests.js`) to evaluate whether a template
 * element will have content before committing the in-document substitution.
 *
 * Supports two token syntaxes:
 * - **Conditional block**: `{?field}...content...{field}` — The entire block is removed
 *   if `field` is absent, null, undefined, or an empty string. If the field has a value,
 *   only the opening `{?field}` marker is removed, leaving the content in place.
 * - **Regular substitution**: `{field}` — Replaced with the field's value from `data`.
 *   Date values are auto-formatted: fields whose name contains "date" use `"M/d/yyyy"`;
 *   fields containing "time" use `"h:mm aa"`; other date fields use `"h:mm aa M/d/yy"`.
 *   If the field name is not present in `data`, it is replaced with `"<fieldName> not specified"`.
 *
 * @param {string} templateString - The template string containing `{field}` tokens.
 * @param {Object} data - Key/value pairs used to fill in the template tokens.
 * @returns {string} The template string with all tokens replaced.
 */
function replaceText(templateString, data) {
  try {
    let result = templateString

    // First pass: Process conditional fields {?field}...{field}
    const conditionalPattern = /\{\?([^}]+)\}(.*?)\{\1\}/g
    const conditionalMatches = [...result.matchAll(conditionalPattern)]

    conditionalMatches.forEach(match => {
      const fullMatch = match[0]
      const fieldName = match[1]
      if (Object.keys(data).includes(fieldName)) {
        const hasValue = data[fieldName] !== null &&
                         data[fieldName] !== undefined &&
                         data[fieldName] !== ''
        if (hasValue) {
          // Field has value - remove just the conditional marker {?field}
          result = result.replace('{?' + fieldName + '}', '')
        } else {
          // Field is empty - remove the entire conditional block
          result = result.replace(fullMatch, '')
        }
      }
    })

    // Second pass: Process regular fields {field}
    const pattern = /{(.*?)}/g
    const innerMatches = [...result.matchAll(pattern)].map(match => match[1])
    innerMatches.forEach(fieldName => {
      let datum
      if (isValidDate(data[fieldName])) {
        if (fieldName.match(/\bdate\b/i)) {
          datum = formatDate(data[fieldName])
        } else if (fieldName.match(/\btime\b/i)) {
          datum = formatDate(data[fieldName], null, "h:mm aa")
        } else {
          datum = formatDate(data[fieldName], null, "h:mm aa M/d/yy")
        }
      } else {
        datum = data[fieldName] || ''
      }
      if (Object.keys(data).includes(fieldName)) {
        result = result.replace("{" + fieldName + "}", datum)
      } else {
        result = result.replace("{" + fieldName + "}", fieldName + " not specified")
      }
    })
    return result
  } catch(e) { logError(e) }
}

/**
 * Formats a customer's name and ID into the canonical display string used
 * throughout the spreadsheet (e.g. in the "Customer Name and ID" column).
 * @param {string} first - Customer's first name.
 * @param {string} last - Customer's last name.
 * @param {string|number} id - Customer's ID.
 * @returns {string} Formatted as `"Last, First (ID)"`.
 */
function getCustomerNameAndId(first, last, id) {
  return `${last}, ${first} (${id})`
}

/**
 * Converts a named range object to its corresponding trigger key by stripping
 * any trailing numeric suffix. For example, `"codeFormatAddress3"` becomes
 * `"codeFormatAddress"`, which is then looked up in the `rangeTriggers` map.
 * @param {GoogleAppsScript.Spreadsheet.NamedRange} namedRange - The named range to convert.
 * @returns {string} The trigger key name with any trailing digits removed.
 */
function convertNamedRangeToTriggerName(namedRange) {
  // remove numeric suffix
  return namedRange.getName().replace(/\d+$/g,'')
}

/**
 * Returns a count combined with the correctly pluralized form of a word.
 * Handles the common case automatically by appending `"s"`, or accepts an
 * explicit plural form. Treats `1` and `"1.0"` (etc.) as singular.
 * @param {number} count - The quantity to display.
 * @param {string} singular - The singular form of the word.
 * @param {string} [plural] - The plural form of the word. Defaults to `singular + "s"`.
 * @returns {string} E.g. `"1 trip"` or `"3 trips"`.
 */
function pluralize(count, singular, plural){
  try {
    let word = ""
    if (count == 1 || /^1(\.0+)?$/.test((count || "").toString())) {
      word = singular
    } else {
      word = plural || `${singular}s`
    }
    return `${count || 0} ${word}`
  } catch (e) {
    logError(e)
  }
}

/**
 * Returns the `SpreadsheetApp.Ui` instance, or `null` if the UI is not available
 * (e.g. when the script is running in a non-interactive context such as a
 * time-based trigger).
 * @returns {GoogleAppsScript.Base.Ui|null}
 */
function safeGetUi() {
  try {
    return SpreadsheetApp.getUi()
  } catch (e) {
    return null
  }
}