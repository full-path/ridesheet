function doGet(e) {
  const params = JSON.stringify(e)
  //log(e.queryString)
  //log(JSON.stringify(e.parameters))
  log(JSON.stringify(e))
  let output = ContentService.createTextOutput()
  output.setMimeType(ContentService.MimeType.JSON)
  output.setContent(JSON.stringify(params))
  return output
}

function doPost(e) {
  log(JSON.stringify(e))
  let output = ContentService.createTextOutput()
  output.setContent(JSON.stringify(e))
  output.setMimeType(ContentService.MimeType.JSON)
  return output
}

"https://script.google.com/a/fullpath.io/macros/s/AKfycbz3o_gnLxEExcwvRC6qjLqNFqTmWYf6EKlfHnacVw/exec"