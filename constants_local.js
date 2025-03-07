const localNamedRanges = {
  "lookupRiderTypes": {
    "sheetName":"Lookups",
    "headerName":"Rider Types"
  },
  "lookupMobilityFactors": {
    "sheetName":"Lookups",
    "headerName":"Mobility Factors"
  },
  "lookupFareTypes": {
    "sheetName":"Lookups",
    "headerName":"Fare Types"
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
  },
  "queryAddresses": {
    "sheetName":"Addresses",
    "headerName":"Short Name"
  }
}

const localNamedRangesToRemove = [
  "queryServiceEndDate",
  "queryServiceId"
]

const localSheetsToRemove = ["Services"]
const localSheets = []
const localSheetsWithHeaders = []

const localColumnsToRemove = {}
const localColumns = {
  "Lookups": {
    "Rider Types": {},
    "Mobility Factors": {},
    "Fare Types": {},
    "Address Short Names": {
      headerFormula: `={"Address Short Names";QUERY(queryAddresses,"SELECT Col1 WHERE Col1 IS NOT NULL ORDER BY Col1",0)}`,
    }
  },
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
      },
    }
  }
}
