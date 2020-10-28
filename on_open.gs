function onOpen(e) {
  const startTime = new Date()
  try {
    buildMenus()
  } catch(e) {
    log("buildMenus", e.name + ': ' + e.message)
  }
  try {
    buildDocumentPropertiesFromSheet()
    buildDocumentPropertiesFromDefaults()
  } catch(e) {
    log("buildDocumentProperties", e.name + ': ' + e.message)
  }
  try {
    buildNamedRanges()
  } catch(e) {
    log("buildNamedRanges", e.name + ': ' + e.message)
  }
  log("onOpen duration:",(new Date()) - startTime)
}

