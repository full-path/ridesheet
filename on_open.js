function onOpen(e) {
  const startTime = new Date()
  try {
    buildMenus()
  } catch(e) { logError(e) }
  try {
    buildDocumentPropertiesFromSheet()
    buildDocumentPropertiesFromDefaults()
  } catch(e) { logError(e) }
  try {
    buildNamedRanges()
  } catch(e) { logError(e) }
  log("onOpen duration:",(new Date()) - startTime)
}