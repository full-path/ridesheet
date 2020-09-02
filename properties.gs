const propDescSuffix = "__description__"

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

// Document properties don't pass on to copied sheets. This recreates the ones put into the properties sheet.
function loadPropertiesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const propSheet = ss.getSheetByName("Document Properties")
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  if (Object.keys(docProps).length === 0 && propSheet) {
    let propsGrid = propSheet.getDataRange().getValues().shift()
    let propsObjects = propsGrid.map(row => {
      let prop = {}
      prop.name = row[0]
      prop.value = row[1]
      prop.description = row[2]
      return prop
    })
    setDocProps(propsObjects)
  }
}

function presentProperties() {
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
  ss.setActiveSheet(propSheet)
}

function updateProperties(e) {
  const row = e.range.getRow()
  if (row > 1) {
    const sheet = e.range.getSheet()
    const propName = sheet.getRange(row,1).getValue()
    const docProps = PropertiesService.getDocumentProperties()
    if (propName && docProps.getKeys().indexOf(propName) !== -1) {
      if (e.value && e.range.getColumn() === 2) {
        const propType = getPropParts(PropertiesService.getDocumentProperties().getProperty(propName)).type
        try {
          setDocProp(propName, coerceValue(e.value, propType))
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

function addDocProp(propName) {
  if (defaultProps[propName] && defaultProps[propName].value) {
    setDocProp(propName, defaultProps[propName].value, defaultProps[propName].description)
    return defaultProps[propName].value
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

function getDocProp(propName, altValue) {
  const prop = PropertiesService.getDocumentProperties().getProperty(propName)
  if (prop) {
    return deserializeProp(prop)
  } else {
    return addDocProp(propName)
  }
}

function getDocProps(props) {
  const docProps = PropertiesService.getDocumentProperties().getProperties()
  let result = {}
  props.forEach(prop => {
    if (getType(prop) === "object") {
      if (prop.name in docProps) {
        result[prop.name] = deserializeProp(docProps[prop.name])
      } else {
        result[prop.name] = addDocProp(prop.name)
      }
    } else {
      result[prop] = deserializeProp(docProps[prop])
    }
  })
  return result
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
  if      (!type)                { return value}
  else if (type === "array")     { return JSON.parse(value) }
  else if (type === "bigint")    { return BigInt(value) }
  else if (type === "boolean")   { return new Boolean(JSON.parse(value)) }
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
  PropertiesService.getDocumentProperties().deleteProperty(propName)
  PropertiesService.getDocumentProperties().deleteProperty(propName + propDescSuffix)
}

function deleteAllDocProps() {
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  Object.keys(docProps).forEach(propName => {
    deleteDocProp(propName)
  })
}

function repairProps() {
  let docProps = PropertiesService.getDocumentProperties().getProperties()
  let newProps = []
  Object.keys(defaultProps).forEach(propName => {
    if (!docProps[propName] || getType(defaultProps[propName].value) !== getPropParts(docProps[propName]).type) {
      let prop = {}
      prop.name = propName
      prop.value = defaultProps[propName].value
      if (defaultProps[propName].description !== docProps[propName + propDescSuffix]) prop.description = defaultProps[propName].description
      newProps.push(prop)
    } else if (!docProps[propName + propDescSuffix] || defaultProps[propName].description !== getPropParts(docProps[propName + propDescSuffix]).value) {
      let prop = {}
      prop.name = propName + propDescSuffix
      prop.value = defaultProps[propName].description
      newProps.push(prop)
    }
  })
  setDocProps(newProps)
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
