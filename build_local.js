function buildLocalMenus() {
    if (getDocProp("apiShowMenuItems")) {
        const ui = SpreadsheetApp.getUi()
        const menuApi = ui.createMenu('Ride Sharing')
        menuApi.addItem('Get trip requests (Deprecated)', 'sendRequestForTripRequests')
        menuApi.addItem('Send trip requests', 'sendTripRequests')
        menuApi.addItem('Send responses to trip requests', 'sendTripRequestResponses')
        menuApi.addItem('Refresh outside runs', 'sendRequestForRuns')
        menuApi.addToUi()
    }
}
