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
  "formulaTripReviewCoreData": {
    "sheetName":"Trip Review",
    "startHeaderName":"Trip Date",
    "endHeaderName":"Run ID",
  },
  "formulaRunsSheet": {
    "sheetName":"Runs",
    "startHeaderName":"Run Date",
    "endHeaderName":"Scheduled End Time",
    "allRows": true
  },
  "formulaRunReviewRunDate": {
    "sheetName":"Run Review",
    "headerName":"Run Date"
  },
  "formulaRunReviewDriverId": {
    "sheetName":"Run Review",
    "headerName":"Driver ID"
  },
  "formulaRunReviewVehicleId": {
    "sheetName":"Run Review",
    "headerName":"Vehicle ID"
  },
  "formulaRunReviewRunId": {
    "sheetName":"Run Review",
    "headerName":"Run ID"
  },
  "formulaRunReviewFareRevenue": {
    "sheetName":"Run Review",
    "headerName":"Fare Revenue"
  },
  "formulaRunReviewDonationRevenue": {
    "sheetName":"Run Review",
    "headerName":"Donation Revenue"
  },
  "formulaRunReviewTicketRevenue": {
    "sheetName":"Run Review",
    "headerName":"Ticket Revenue"
  },
  "formulaRunReviewOdoStart": {
    "sheetName":"Run Review",
    "headerName":"Odometer Start"
  },
  "formulaRunReviewOdoEnd": {
    "sheetName":"Run Review",
    "headerName":"Odometer End"
  },
  "formulaRunsRunDate": {
    "sheetName":"Runs",
    "headerName":"Run Date"
  },
  "formulaRunsDriverId": {
    "sheetName":"Runs",
    "headerName":"Driver ID"
  },
  "formulaRunsVehicleId": {
    "sheetName":"Runs",
    "headerName":"Vehicle ID"
  },
  "formulaRunsRunId": {
    "sheetName":"Runs",
    "headerName":"Run ID"
  },
  "formulaTripsCoreData": {
    "sheetName":"Trips",
    "startHeaderName":"Trip Date",
    "endHeaderName":"Run ID",
  },
  "formulaTripsCoreHeaders": {
    "sheetName":"Trips",
    "startHeaderName":"Trip Date",
    "endHeaderName":"Run ID",
    "headerOnly": true
  }
}
const localNamedRangesToRemove = [
  "codeCheckSourceOnShare",
  "codeVerifySourceOnEdit",
  "queryServiceEndDate",
  "queryServiceId"
]

const localSheetsToRemove = ["Services","Sent Trips","Outside Trips","Outside Runs"]
const localSheets = ["Addresses"]
const localSheetsWithHeaders = ["Addresses"]

const localColumnsToRemove = {}
const localColumns = {
  Customers: {
    "Default Rider Type": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRiderTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Default Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRiderTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Default Mobility Factors": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupMobilityFactors",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid mobility factor.",
      },
    },
    "Default Fare": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupFareTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid fare type.",
      },
    },
    "Default Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip purpose.",
      },
    },
  },
  Trips: {
    "Will Call": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "WILL CALL",
        allowInvalid: false,
      },
    },
    "SDR": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: false,
      },
    },
    "Rider Type": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRiderTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Mobility Factors": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupMobilityFactors",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid mobility factor.",
      },
    },
    "Fare": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupFareTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid fare type.",
      },
    },
    "Day of Week": {
      headerFormula: `={"Day of Week";arrayformula(if(ISBLANK(A2:A), "", (TEXT(A2:A, "DDDD"))))}`
    }
  },
  "Runs": {
    "|First PU Time|": {
      headerFormula: `={"|First PU Time|","|Last DO Time|","|Trip Count|";MAP(formulaRunsRunDate, formulaRunsDriverId, formulaRunsVehicleId, formulaRunsRunId, LAMBDA(RunDate,RunDriverId,RunVehicleID,RunRunId, QUERY_RUN_TRIP_TIMES(RunDate,RunDriverId,RunVehicleID,RunRunId, formulaTripsCoreData, formulaTripsCoreHeaders)))}`,
      numberFormat: 'h":"mm am/pm',
    },
    "|Last DO Time|": {
      headerFormula: "",
      numberFormat: 'h":"mm am/pm',
    },
    "|Trip Count|": {
      headerFormula: "",
      numberFormat: '0',
    },
  },
  "Trip Review": {
    "Will Call": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "WILL CALL",
        allowInvalid: true,
      },
    },
    "SDR": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: true,
      },
    },
    "Rider Type": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRiderTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Mobility Factors": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupMobilityFactors",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid mobility factor.",
      },
    },
    "Fare": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupFareTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid fare type.",
      },
    },
  },
  "Run Review": {
    "Tickets Used": {
      numberFormat: "0"
    },
    "Fare Revenue": {
      numberFormat: "$0.00"
    },
    "Donation Revenue": {
      numberFormat: "$0.00"
    },
    "Ticket Revenue": {
      numberFormat: "$0.00"
    },
    "Total Revenue": {
      headerFormula: `={"Total Revenue";MAP(formulaRunReviewFareRevenue,formulaRunReviewDonationRevenue,formulaRunReviewTicketRevenue,LAMBDA(source1,source2,source3,IF(COUNTA(source1,source2,source3) < 3,"",SUM(source1,source2,source3))))}`
    },
    "Starting Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Ending Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Vehicle Garage Address": {},
    "First PU Address": {},
    "Last DO Address": {}
  },
  "Trip Archive": {
    "Will Call": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "WILL CALL",
        allowInvalid: false,
      },
    },
    "SDR": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: false,
      },
    },
    "Rider Type": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRiderTypes",
        showDropdown: false,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: false,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Mobility Factors": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupMobilityFactors",
        showDropdown: false,
        allowInvalid: false,
        helpText: "Value must be a valid mobility factor.",
      },
    },
    "Fare": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupFareTypes",
        showDropdown: false,
        allowInvalid: false,
        helpText: "Value must be a valid fare type.",
      },
    },
  },
  "Run Archive": {
    "Tickets Used": {
      numberFormat: "0"
    },
    "Fare Revenue": {
      numberFormat: "$0.00"
    },
    "Donation Revenue": {
      numberFormat: "$0.00"
    },
    "Ticket Revenue": {
      numberFormat: "$0.00"
    },
    "Starting Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Ending Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Vehicle Garage Address": {},
    "First PU Address": {},
    "Last DO Address": {}
  },
  "Addresses": {
    "Short Name": {},
    "Address": {},
  },
  "Lookups": {
    "Rider Types": {},
    "Mobility Factors": {},
    "Fare Types": {},
  }

}