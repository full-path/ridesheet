
function testSendGet() {
  const url = "https://script.google.com/macros/s/AKfycbzf4vpUph1sLBmGXIn2Ei9qBLc4VgRKcyuG6EgASMeUL293u67d/exec"
  const query = "?pish=posh"
  const options = {
    method: 'GET',
    //followRedirects: true,
    //muteHttpExceptions: true,
    contentType: 'application/json',
  }
  let response = UrlFetchApp.fetch(url + query, options)
  //log(response)
  log("Received from GET:",JSON.stringify(JSON.parse(response.getContentText())))
}

function testSendPost() {
  const url = "https://script.google.com/macros/s/AKfycbzf4vpUph1sLBmGXIn2Ei9qBLc4VgRKcyuG6EgASMeUL293u67d/exec"
  const query = "?pish=posh&what=hey&what=boom"
  const options = {
    method: 'POST',
    //followRedirects: true,
    //muteHttpExceptions: true,
    contentType: 'application/json',
  }
  let response = UrlFetchApp.fetch(url + query, options)
  //log(response)
  log("Received from POST:",JSON.stringify(JSON.parse(response.getContentText())))
  //log(JSON.stringify(response.getAllHeaders()))
}
