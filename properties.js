/**
 * @fileoverview Document property management for RideSheet.
 *
 * RideSheet stores per-installation configuration as Google Apps Script
 * document properties (key/value string pairs). This file provides a typed
 * property layer on top of that raw storage:
 *
 * - Properties are serialized with an embedded type tag (e.g. `{{number   }}42`)
 *   so that arrays, booleans, dates, etc. survive the string-only storage format.
 * - A matching description for each property is stored under the key
 *   `propName + "__description__"` and surfaced in the "Document Properties" sheet.
 * - Properties whose names end with `"_"` are private and hidden from users.
 * - `getDocProp()` caches retrieved values in `cachedDocProps` for the lifetime
 *   of the current script execution to minimise calls to `PropertiesService`.
 * - Default values are defined in `defaultDocumentProperties` (constants.js)
 *   and are returned by `getDocProp()` when a property has not been explicitly set.
 */

/** @type {string} Suffix appended to a property name to store its description. */
const propDescSuffix = "__description__"
/**
 * In-memory cache for deserialized document property values.
 * Populated lazily by `getDocProp()` and `getDocProps()`.
 * @type {Object.<string, *>}
 */
const cachedDocProps = {}
/**
 * Reserved for future use; not currently set to `true` anywhere in the codebase.
 * @type {boolean}
 */
let allDocPropsCached = false

/**
 * Returns an array of all document properties as plain objects, suitable
 * for display in the "Document Properties" sheet.
 *
 * Each returned object contains:
 * - `name` {string}         - The property key.
 * - `value` {string}        - The raw (unserialized) value string.
 * - `description` {string}  - The associated description, or `""` if none
 *   exists. Description-only entries (keys ending with `propDescSuffix`)
 *   are included as rows but without a `description` field of their own.
 *
 * @param {boolean} [showPrivateProperties=false] - When `true`, includes
 *   properties whose names end with `"_"`, which are hidden from users by default.
 * @returns {Array<{name: string, value: string, description?: string}>}
 */
function getProperties(showPrivateProperties) {
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  let docPropKeys = Object.keys(docProps).sort()
  let filteredDocPropKeys = showPrivateProperties ? docPropKeys : docPropKeys.filter(key => !key.endsWith("_"))
  let propsArray = []
  filteredDocPropKeys.forEach(propName => {
    let thisRow = {name: propName, value: getPropParts(docProps[propName]).value}
    if (propName.indexOf(propDescSuffix) === -1) {
      if (docPropKeys.indexOf(propName + propDescSuffix) === -1) {
        thisRow.description = ""
      } else {
        thisRow.description = getPropParts(docProps[propName + propDescSuffix]).value
      }
    }
    propsArray.push(thisRow)
  })
  return propsArray
}

/**
 * Reads a JSON string from the currently selected spreadsheet cell and writes
 * its contents as document properties. The JSON must be an array of objects
 * with `name`, `value`, and optionally `description` fields, matching the
 * format accepted by `setDocProps()`.
 * Intended as a developer utility for bulk-importing properties.
 */
function loadPropertiesFromJSON() {
  const range = SpreadsheetApp.getActiveRange()
  const props = JSON.parse(range.getValue())
  setDocProps(props)
}

/**
 * Rebuilds the "Document Properties" sheet with current property names, values,
 * and descriptions. Creates the sheet if it does not already exist.
 * Applies header formatting and auto-resizes all three columns.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The updated "Document Properties" sheet.
 */
function updatePropertiesSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let propSheet = ss.getSheetByName("Document Properties") || ss.insertSheet("Document Properties")
  propSheet.getDataRange().clear()
  const headerValues = ["Property Name","Property Value","Property Description"]
  let header = propSheet.getRange(1, 1, 1, 3)
  
  header.setValues([["Property Name","Property Value","Property Description"]])
  header.setBackground(headerBackgroundColor).setFontWeight("bold")
  propSheet.setFrozenRows(1)
  propSheet.setFrozenColumns(1)
  const props = getProperties().map(row => [row.name, row.value, row.description])
  if (props.length > 0) {
    const propRange = propSheet.getRange(2,1,props.length,3)
    propRange.setValues(props)
    propSheet.autoResizeColumns(1,3)
  }
  return propSheet
}

/**
 * Rebuilds the "Document Properties" sheet and navigates the active user to it.
 * Called from the Settings menu item "Refresh document properties sheet".
 */
function presentProperties() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(updatePropertiesSheet())
}

/**
 * Handles edits to the "Document Properties" sheet, updating the corresponding
 * document property when the user changes a value in column 2 (Property Value).
 * Reverts the cell to its previous value if:
 * - The edit is not in column 2 or not below the header row.
 * - The property name is not found in the stored properties.
 * - The new value is blank.
 * - The value cannot be coerced to the property's declared type.
 * Triggered via `initialSheetTriggers` in on_edit.js.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The onEdit event object.
 */
function updateProperties(e) {
  const row = e.range.getRow()
  const column = e.range.getColumn()
  if (row > 1 && column === 2) {
    const sheet = e.range.getSheet()
    const propName = sheet.getRange(row,1).getValue()
    const propValue = e.value
    const docProps = PropertiesService.getDocumentProperties()
    if (propName && docProps.getKeys().indexOf(propName) !== -1) {
      if (propValue) {
        const propType = getPropParts(docProps.getProperty(propName)).type
        try {
          setDocProp(propName, coerceValue(propValue, propType))
          e.source.toast(`Property "${propName}" updated to "${e.value}".`,"Success")
        } catch(error) {
          e.source.toast(`Property "${propName}" could not be updated: "${error.message}".`,"Update Error",-1)
          e.range.setValue(e.oldValue)
        }
      }
    }
  } else {
    e.range.setValue(e.oldValue)
  }
}

/**
 * Removes any stored document properties (and their description entries) that
 * are no longer defined in `defaultDocumentProperties`. Useful after a code
 * update that removes or renames properties. Refreshes the "Document Properties"
 * sheet if any properties are removed.
 */
function purgeOldDocumentProperties() {
  const docProps = PropertiesService.getDocumentProperties()
  const docPropKeys = Object.keys(docProps.getProperties())
  const defaultDocPropKeys = Object.keys(defaultDocumentProperties)
  const oldDocPropKeys = docPropKeys.filter(docPropKey => {
    if (docPropKey.indexOf(propDescSuffix) === -1) {
      return !defaultDocPropKeys.includes(docPropKey)
    } else {
      return !defaultDocPropKeys.includes(docPropKey.slice(0,-propDescSuffix.length))
    }
  })
  if (oldDocPropKeys.length) {
    oldDocPropKeys.forEach(oldDocPropKey => docProps.deleteProperty(oldDocPropKey))
    updatePropertiesSheet()
  }
}

/**
 * Adds a single document property from its definition in `defaultDocumentProperties`,
 * setting it to its default value and description. Shows a toast and logs a
 * message if the property name is not found in the defaults.
 * @param {string} propName - The property name to add.
 * @returns {*} The default value that was set, or `undefined` if the property was not found.
 */
function addDocProp(propName) {
  if (defaultDocumentProperties[propName] && defaultDocumentProperties[propName].value) {
    setDocProp(propName, defaultDocumentProperties[propName].value, defaultDocumentProperties[propName].description)
    return defaultDocumentProperties[propName].value
  } else {
    const msg = "Property " + propName + " not found"
    SpreadsheetApp.getActiveSpreadsheet().toast(msg)
    log(msg)
  }
}

/**
 * Serializes and saves a single document property with type information.
 * Optionally stores an associated description under `propName + "__description__"`.
 * @param {string} propName - The property name.
 * @param {*} value - The value to store. Serialized automatically with a type tag.
 * @param {string} [description] - Optional human-readable description stored alongside the value.
 */
function setDocProp(propName, value, description) {
  const type = getType(value)
  let props = {}
  props[propName] = serializeProp(value, type)
  if (description) props[propName + propDescSuffix] = description
  PropertiesService.getDocumentProperties().setProperties(props)
}

/**
 * Serializes and saves multiple document properties in a single call.
 * Each entry's description, if provided, is stored under `name + "__description__"`.
 * @param {Array<{name: string, value: *, description?: string}>} props - Array of property objects to save.
 */
function setDocProps(props) {
  let docProps = {}
  props.forEach(prop => {
    docProps[prop.name] = serializeProp(prop.value)
    if (prop.description) docProps[prop.name + propDescSuffix] = prop.description
  })
  PropertiesService.getDocumentProperties().setProperties(docProps)
}

/**
 * Retrieves and deserializes a single document property by name.
 * Results are cached in `cachedDocProps` for the lifetime of the script
 * execution to avoid repeated calls to `PropertiesService`.
 * Falls back to the default value from `defaultDocumentProperties` if the
 * property has not been explicitly set.
 * @param {string} propName - The property name.
 * @returns {*} The deserialized value, or `null` if not found in either
 *   stored properties or the defaults.
 */
function getDocProp(propName) {
  try {
    if (cachedDocProps[propName]) {
      return cachedDocProps[propName]
    } else {
      const prop = PropertiesService.getDocumentProperties().getProperty(propName)
      if (prop) {
        let result = deserializeProp(prop)
        cachedDocProps[propName] = result
        return result
      } else if (defaultDocumentProperties.hasOwnProperty(propName) &&
          defaultDocumentProperties[propName].hasOwnProperty("value")) {
        let result = defaultDocumentProperties[propName].value
        cachedDocProps[propName] = result
        return result
      } else {
        return null
      }
    }
  } catch(e) { logError(e) }
}

/**
 * Retrieves and deserializes multiple document properties in a single call.
 * Results are cached individually in `cachedDocProps`. Each entry in `props`
 * can be either a property name string or an object with a `name` field.
 * @param {Array<string|{name: string}>} props - Property names or name-objects to retrieve.
 * @returns {Object.<string, *>} An object mapping each property name to its deserialized value.
 */
function getDocProps(props) {
  try {
    const docProps = PropertiesService.getDocumentProperties().getProperties()
    let result = {}
    props.forEach(prop => {
      let propName
      if (getType(prop) === "object") {
        propName = prop.name
      } else {
        propName = prop
      }
      if (cachedDocProps[propName]) {
        result[propName] = cachedDocProps[propName]
      } else if (docProps.hasOwnProperty(propName)) {
        let thisResult = deserializeProp(docProps[propName])
        cachedDocProps[propName] = thisResult
        result[propName] = thisResult
      } else if (defaultDocumentProperties.hasOwnProperty(propName) &&
          defaultDocumentProperties[propName].hasOwnProperty("value")) {
        let result = defaultDocumentProperties[propName].value
        cachedDocProps[propName] = result
        return result
      } else {
        return null
      }
    })
    return result
  } catch(e) { logError(e) }
}

/**
 * Serializes a JavaScript value to a string with an embedded 13-character type
 * tag prefix, so the value and its type can be faithfully recovered from
 * Apps Script's string-only property storage.
 *
 * Format: `{{type     }}<value>` where `type` is padded to 9 characters.
 * Examples:
 * - `"hello"` → `"{{string   }}hello"`
 * - `42`       → `"{{number   }}42"`
 * - `[1, 2]`   → `"{{array    }}[1, 2]"` (JSON-encoded)
 *
 * @param {*} value - The value to serialize.
 * @returns {string} The serialized string with type tag prefix.
 */
function serializeProp(value) {
  const type = getType(value)
  
  if      (type === "array")     { return '{{array    }}' + JSON.stringify(value, null, 2) }
  else if (type === "bigint")    { return '{{bigint   }}' + value }
  else if (type === "boolean")   { return '{{boolean  }}' + JSON.stringify(value) }
  else if (type === "date")      { return '{{date     }}' + JSON.stringify(value) }
  else if (type === "map")       { return '{{map      }}' + JSON.stringify(Array.from(value.entries()), null, 2) }
  else if (type === "null")      { return '{{null     }}' }
  else if (type === "number")    { return '{{number   }}' + value}
  else if (type === "object")    { return '{{object   }}' + JSON.stringify(value, null, 2) }
  else if (type === "set")       { return '{{set      }}' + JSON.stringify(Array.from(value.keys()), null, 2) }
  else if (type === "string")    { return '{{string   }}' + value }
  else if (type === "undefined") { return '{{undefined}}' }
  else                           { return '{{string   }}' + value }
}

/**
 * Parses a serialized property string (as produced by `serializeProp()`) and
 * returns the value coerced back to its original JavaScript type.
 * @param {string} prop - A serialized property string with a `{{type}}` prefix.
 * @returns {*} The deserialized value.
 */
function deserializeProp(prop) {
  const parts = getPropParts(prop)
  return coerceValue(parts.value, parts.type)
}

/**
 * Splits a serialized property string into its type tag and raw value string.
 * Recognizes the fixed-width `{{type     }}` prefix format used by `serializeProp()`.
 * If the string does not begin with a valid `{{...}}` prefix, the entire string
 * is returned as a value with type `"string"` for backwards compatibility with
 * any un-tagged legacy values.
 * @param {string} prop - A serialized property string.
 * @returns {{value: string, type: string}} The raw value string and type name.
 */
function getPropParts(prop) {
  const frontMatter = prop.slice(0,13)
  if (frontMatter.slice(0,2) === '{{' && frontMatter.slice(-2) === '}}') {
    const value = prop.slice(13)
    const type = frontMatter.slice(2,11).trim()
    return {value: value, type: type}
  } else {
    return {value: prop, type: 'string'}
  }
}

/**
 * Converts a value to the specified JavaScript type. Used when deserializing
 * stored document properties and when processing user input from the
 * "Document Properties" sheet.
 *
 * Coercion rules:
 * - `"array"` / `"object"` — parsed from JSON
 * - `"map"` / `"set"` — reconstructed from JSON-encoded entries/keys
 * - `"boolean"` — `"false"`, `"no"`, `"0"`, and falsy values → `false`; anything else → `true`
 * - `"number"` — converted with `Number()`; throws if result is not finite
 * - `"date"` — parsed from a JSON-encoded date string
 * - `"null"` → `null`, `"undefined"` → `undefined`
 * - `"string"` or unknown type — returned as-is
 * - If `type` is falsy or already matches the value's current type, returns value unchanged.
 *
 * @param {*} value - The value to coerce (typically a string from storage or user input).
 * @param {string} type - The target type name (e.g. `"number"`, `"array"`, `"boolean"`).
 * @returns {*} The coerced value.
 * @throws {Error} If `type` is `"number"` and the value cannot be converted to a finite number.
 */
function coerceValue(value, type) {
  if      (!type || type === getType(value)) { return value }
  else if (type === "array")     { return JSON.parse(value) }
  else if (type === "bigint")    { return BigInt(value) }
  else if (type === "boolean")   {
    if (value.toLowerCase() === "false" || value.toLowerCase() === "no" || value === "0" || !value) {
      return false
    } else {
      return true
    }
  }
  else if (type === "date")      { return new Date(JSON.parse(value)) } 
  else if (type === "map")       { return new Map(JSON.parse(value)) }
  else if (type === "null")      { return null }
  else if (type === "number")    { 
    const result = Number(value) 
    if (isFinite(result)) {
      return result
    } else {
      throw new Error("Invalid Number")
    }
  }
  else if (type === "object")    { return JSON.parse(value) }
  else if (type === "set")       { return new Set(JSON.parse(value))}
  else if (type === "string")    { return value }
  else if (type === "undefined") { return undefined }
  else                           { return value }
}

/**
 * Deletes a document property and its associated description entry
 * (the key `propName + "__description__"`) if one exists.
 * @param {string} propName - The property name to delete.
 */
function deleteDocProp(propName) {
  const docProps = PropertiesService.getDocumentProperties()
  docProps.deleteProperty(propName)
  docProps.deleteProperty(propName + propDescSuffix)
}

/**
 * Deletes all document properties, including all description entries.
 * Intended as a developer utility for fully resetting a spreadsheet's
 * property state.
 */
function deleteAllDocProps() {
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  Object.keys(docProps).forEach(propName => {
    deleteDocProp(propName)
  })
}

/**
 * Deletes any stored document properties not present in `defaultDocumentProperties`,
 * including their associated description entries. Unlike `purgeOldDocumentProperties()`,
 * this does not refresh the "Document Properties" sheet afterward.
 * Intended as a developer/migration utility.
 */
function deleteDeprecatedProps() {
  try {
    const defaultPropNames = Object.keys(defaultDocumentProperties)
    const defaultPropDescriptions = defaultPropNames.map(propName => propName + propDescSuffix)
    const currentPropNames = Object.keys(PropertiesService.getDocumentProperties().getProperties())
    currentPropNames.forEach(propName => {
      if (defaultPropNames.indexOf(propName) === -1 && 
          defaultPropDescriptions.indexOf(propName) === -1) deleteDocProp(propName)
    })
  } catch(e) {
    logError(e)    
  }
}