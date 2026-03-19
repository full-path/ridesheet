function addSortingMenuItems() {
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu('Trip Sorting')
  const config = getSortConfig()
  for (let i = 0; i < Math.min(config.length, 9); i++) {
    menu.addItem(config[i].name, `sortingMenuItem${i}`)
  }
  menu.addToUi()
}

function sortTrips(index) {
  const config = getSortConfig()

  const sortingByColumnName = config[index].sort

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = ss.getActiveSheet()
  const dataRange = sheet.getDataRange()
  const numRows = dataRange.getNumRows()
  const numCols = dataRange.getNumColumns()
  if (numRows <= 1) return
  const rangeToSort = sheet.getRange(2, 1, numRows - 1, numCols)
  const sheetHeaderNames = getSheetHeaderNames(sheet)

  let missingHeaders = []
  const sortingByColumnNumber = sortingByColumnName.map(item => {
    const columnNumber = sheetHeaderNames.indexOf(item.column) + 1
    if (!columnNumber) missingHeaders.push(item.column)
    return {column: columnNumber, ascending: item.ascending}
  })
  if (missingHeaders.length) {
    ss.toast(`Sorting failed. Columns missing: ${missingHeaders.join(", ")}`)
    return
  }

  rangeToSort.sort(sortingByColumnNumber)
}

function sortingMenuItem0() { sortTrips(0) }
function sortingMenuItem1() { sortTrips(1) }
function sortingMenuItem2() { sortTrips(2) }
function sortingMenuItem3() { sortTrips(3) }
function sortingMenuItem4() { sortTrips(4) }
function sortingMenuItem5() { sortTrips(5) }
function sortingMenuItem6() { sortTrips(6) }
function sortingMenuItem7() { sortTrips(7) }
function sortingMenuItem8() { sortTrips(8) }
function sortingMenuItem9() { sortTrips(9) }

function getSortConfig() {
  return [
    {
      name: "Date, PU Time",
      sort: [
        {
          column: "Trip Date",
          ascending: false
        },
        {
          column: "PU Time",
          ascending: true
        },
      ]
    },
    {
      name: "Date, Vehicle, PU Time",
      sort: [
        {
          column: "Trip Date",
          ascending: false
        },
        {
          column: "Vehicle ID",
          ascending: true
        },
        {
          column: "PU Time",
          ascending: true
        },
      ]
    },
    {
      name: "Customer, Date, PU Time",
      sort: [
        {
          column: "Customer Name and ID",
          ascending: true
        },
        {
          column: "Trip Date",
          ascending: false
        },
        {
          column: "PU Time",
          ascending: true
        },
      ]
    }
  ]
}