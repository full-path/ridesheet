function onChange(e) {
  if (e.changeType == "INSERT_COLUMN" || e.changeType == "REMOVE_COLUMN") {
    const startTime = new Date()
    storeHeaderInformation(e)
    log("onChange duration:",(new Date()) - startTime)
  }
}