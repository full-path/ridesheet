const localNamedRanges = {
  "queryAddresses": {
    "sheetName":"Addresses",
    "headerName":"Short Name"
  },
  "lookupAddressShortNames": {
    "sheetName":"Lookups",
    "headerName":"Address Short Names"
  },
  "localCodeExpandAddress1": {
    "sheetName":"Trips",
    "headerName":"|PU|"
  },
  "localCodeExpandAddress2": {
    "sheetName":"Trips",
    "headerName":"|DO|"
  }
}
const localNamedRangesToRemove = []

const localSheetsToRemove = []
const localSheets = []
const localSheetsWithHeaders = []

const localColumnsToRemove = {}
const localColumns = {
  "Trips": {
    "|PU|": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupAddressShortNames",
        showDropdown: false,
        allowInvalid: false
      },
    },
    "|DO|": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupAddressShortNames",
        showDropdown: false,
        allowInvalid: false
      }
    }
  },
  "Lookups": {
    "Address Short Names": {
      headerFormula: `={"Address Short Names";QUERY(queryAddresses,"SELECT Col1 WHERE Col1 IS NOT NULL ORDER BY Col1",0)}`
    }
  }
}
