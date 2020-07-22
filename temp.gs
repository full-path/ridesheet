function testGeoCode() {
  log(getGeocode("63826 Rock House, CV 97641","raw"))
}

function getBackgroundColor() {
  cell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getCurrentCell() 
  value = cell.getBackgroundObject().asRgbColor().asHexString()
  log(value)
}

function setStuff() {
  setDocProp("backupFolderId","19cDx90nCYrwLyoJLHrxyO7TwXQ3KmZ-e")
  setDocProp("monthlyBackupFolderId","1UWlgQbR9JTf_17dSwz71oekYA81H-MIj")
  setDocProp("weeklyBackupFolderId","10ah2KEwSCm_XsXV5lv-kHUKqLO_FmRju")
  setDocProp("nightlyBackupFolderId","19R7u9mmSa_eIS_qxj5WCaZi22P5QMyW0")
  setDocProp("nightlyFileRetentionInDays",90)
  setDocProp("weeklyFileRetentionInDays",730)
}

function testDateThing() {
  //log(formatDate(null,null,"yyyy-MM-dd"))
  makeBackup(getDocProp("weeklyBackupFolderId"))
}