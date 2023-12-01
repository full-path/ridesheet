function onOpen(e) {
  const startTime = new Date()
  try {
    buildMenus()
  } catch(e) { logError(e) }
  try {
    buildDocumentPropertiesFromSheet()
    buildDocumentPropertiesFromDefaults()
    purgeOldDocumentProperties()
  } catch(e) { logError(e) }
  try {
    buildNamedRanges()
    buildMetadata()
  } catch(e) { logError(e) }
  log("onOpen duration:",(new Date()) - startTime)
}