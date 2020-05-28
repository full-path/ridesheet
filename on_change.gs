function onChange(e) {
  const startTime = new Date()
  if (e.changeType == "INSERT_COLUMN" || e.changeType == "REMOVE_COLUMN") {
    storeHeaderInformation(e)
  }
  log("onChange duration:",(new Date()) - startTime)
}