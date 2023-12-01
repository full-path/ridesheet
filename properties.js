const propDescSuffix = "__description__"
var cachedDocProps = {}
var allDocPropsCached = false

function getProperties(showPrivateProperties) {
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  let docPropKeys = Object.keys(docProps).sort()
  let filteredDocPropKeys = showPrivateProperties ? docPropKeys : docPropKeys.filter(key => !key.endsWith("_"))
  propsArray = []
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

function loadPropertiesFromJSON() {
  const range = SpreadsheetApp.getActiveRange()
  const props = JSON.parse(range.getValue())
  setDocProps(props)
}

function updatePropertiesSheet() {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let propSheet = ss.getSheetByName("Document Properties") || ss.insertSheet("Document Properties")
  propSheet.getDataRange().clear()
  headerValues = ["Property Name","Property Value","Property Description"]
  let header = propSheet.getRange(1, 1, 1, 3)
  
  header.setValues([["Property Name","Property Value","Property Description"]])
  header.setBackground(headerBackgroundColor).setFontWeight("bold")
  propSheet.setFrozenRows(1)
  propSheet.setFrozenColumns(1)
  props = getProperties().map(row => [row.name, row.value, row.description])
  if (props.length > 0) {
    propRange = propSheet.getRange(2,1,props.length,3)
    propRange.setValues(props)
    propSheet.autoResizeColumns(1,3)
  }
  return propSheet
}

function presentProperties() {
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(updatePropertiesSheet())
}

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

function addDocProp(propName) {
  if (defaultDocumentProperties[propName] && defaultDocumentProperties[propName].value) {
    setDocProp(propName, defaultDocumentProperties[propName].value, defaultDocumentProperties[propName].description)
    return defaultDocumentProperties[propName].value
  } else {
    msg = "Property " + propName + " not found"
    SpreadsheetApp.getActiveSpreadsheet().toast(msg)
    log(msg)
  }
}

function setDocProp(propName, value, description) {
  const type = getType(value)
  let props = {}
  props[propName] = serializeProp(value, type)
  if (description) props[propName + propDescSuffix] = description
  PropertiesService.getDocumentProperties().setProperties(props)
}

function setDocProps(props) {
  let docProps = {}
  props.forEach(prop => {
    docProps[prop.name] = serializeProp(prop.value)
    if (prop.description) docProps[prop.name + propDescSuffix] = prop.description
  })
  PropertiesService.getDocumentProperties().setProperties(docProps)
}

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

function serializeProp(value) {
  const type = getType(value)
  
  if      (type === "array")     { return '{{array    }}' + JSON.stringify(value) } 
  else if (type === "bigint")    { return '{{bigint   }}' + value }
  else if (type === "boolean")   { return '{{boolean  }}' + JSON.stringify(value) } 
  else if (type === "date")      { return '{{date     }}' + JSON.stringify(value) } 
  else if (type === "map")       { return '{{map      }}' + JSON.stringify(Array.from(value.entries())) }
  else if (type === "null")      { return '{{null     }}' }
  else if (type === "number")    { return '{{number   }}' + value}
  else if (type === "object")    { return '{{object   }}' + JSON.stringify(value) } 
  else if (type === "set")       { return '{{set      }}' + JSON.stringify(Array.from(value.keys())) } 
  else if (type === "string")    { return '{{string   }}' + value }
  else if (type === "undefined") { return '{{undefined}}' }  
  else                           { return '{{string   }}' + value }
}

function deserializeProp(prop) {
  const parts = getPropParts(prop)
  return coerceValue(parts.value, parts.type)
}

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

function deleteDocProp(propName) {
  const docProps = PropertiesService.getDocumentProperties()
  docProps.deleteProperty(propName)
  docProps.deleteProperty(propName + propDescSuffix)
}

function deleteAllDocProps() {
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  Object.keys(docProps).forEach(propName => {
    deleteDocProp(propName)
  })
}

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

function testTypes() {
//  deleteDocProp("tripReviewRequiredFields")
  repairProps()
//  log(PropertiesService.getDocumentProperties().getProperty("tripReviewRequiredFields"))
//  setDocProp("testArray",[1,2,3])
//  setDocProp("testBigInt",BigInt(123))
//  setDocProp("testBool", true)
//  setDocProp("testBoolFalse", false)
//  setDocProp("testDate",new Date())
//  setDocProp("testMap",new Map([[1,"yes"],[2,"no"]]))
//  setDocProp("testNull",null)
//  setDocProp("testNumber",3.1415)
//  setDocProp("testObject",{1:2,3:4,5:"six","seven":8})
//  setDocProp("testSet",new Set([1,2,3]))
//  setDocProp("testString","Test!")
//  setDocProp("testSet",new Set([1,2,3]))
//  setDocProp("testUndefined",undefined)
}

function cleanUpTestTypes() {
//  deleteDocProp("testArray")
//  deleteDocProp("testBigInt")
//  deleteDocProp("testBool")
//  deleteDocProp("testBoolFalse")
//  deleteDocProp("testDate")
//  deleteDocProp("testMap")
//  deleteDocProp("testNull")
//  deleteDocProp("testNumber")
//  deleteDocProp("testObject")
//  deleteDocProp("testSet")
//  deleteDocProp("testString")
//  deleteDocProp("testSet")
//  deleteDocProp("testUndefined")
}