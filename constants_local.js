const localNamedRanges = {
  "formulaTripsTripDate": {
    "sheetName":"Trips",
    "headerName":"Trip Date"
  },
  "formulaTripsPuTime": {
    "sheetName":"Trips",
    "headerName":"PU Time"
  },
  "formulaTripsDoTime": {
    "sheetName":"Trips",
    "headerName":"DO Time"
  },
  "formulaTripsTripDriverId": {
    "sheetName":"Trips",
    "headerName":"Driver ID"
  },
  "formulaTripsTripVehicleId": {
    "sheetName":"Trips",
    "headerName":"Vehicle ID"
  },
  "formulaTripsTripRunId": {
    "sheetName":"Trips",
    "headerName":"Run ID"
  },
  "formulaTripsCoreHeaders": {
    "sheetName":"Trips",
    "startHeaderName":"Trip Date",
    "endHeaderName":"Run ID",
    "headerOnly": true
  },
  "formulaTripsCoreData": {
    "sheetName":"Trips",
    "startHeaderName":"Trip Date",
    "endHeaderName":"Run ID",
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
  "formulaRunsCoreHeaders": {
    "sheetName":"Runs",
    "startHeaderName":"Run Date",
    "endHeaderName":"Scheduled End Time",
    "headerOnly": true
  },
  "formulaRunsCoreData": {
    "sheetName":"Runs",
    "startHeaderName":"Run Date",
    "endHeaderName":"Scheduled End Time",
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
  "formulaRunReviewStartingDeadheadMiles": {
    "sheetName":"Run Review",
    "headerName":"Starting Deadhead Miles"
  },
  "formulaRunReviewEndingDeadheadMiles": {
    "sheetName":"Run Review",
    "headerName":"Ending Deadhead Miles"
  },
  "formulaRunReviewTotalVehicleMiles": {
    "sheetName":"Run Review",
    "headerName":"Total Vehicle Miles"
  },
  "formulaRunReviewTotalDeadheadMiles": {
    "sheetName":"Run Review",
    "headerName":"Total Deadhead Miles"
  },
  "codeFormatAddress9": {
    "sheetName":"Addresses",
    "headerName":"Address"
  },
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
    "|Run OK?|": {
      headerFormula: `={"|Run OK?|";MAP(formulaTripsTripDate, formulaTripsPuTime, formulaTripsDoTime, formulaTripsTripDriverId, formulaTripsTripVehicleId, formulaTripsTripRunId, LAMBDA(TripDate,TripPuTime,TripDoTime,TripDriverId,TripVehicleID,TripRunId, QUERY_RUN_MATCH_COUNT(TripDate,TripPuTime,TripDoTime,TripDriverId,TripVehicleID,TripRunId,formulaRunsCoreData,formulaRunsCoreHeaders)))}`
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
    "Total Trip Fares": {
      headerFormula: `={"Total Trip Fares";MAP(formulaRunReviewRunDate,formulaRunReviewDriverId,formulaRunReviewVehicleId,formulaRunReviewRunId,LAMBDA(tripDate,driverId,vehicleId,runId,QUERY_TRIP_FARE_SUM(tripDate,driverId,vehicleID,runId,'Trip Review'!A2:S,'Trip Review'!A1:S1)))}`
    },
    "Total Revenue": {
      headerFormula: `={"Total Revenue";MAP(formulaRunReviewFareRevenue,formulaRunReviewDonationRevenue,formulaRunReviewTicketRevenue,LAMBDA(source1,source2,source3,IF(COUNTA(source1,source2,source3) < 3,"",SUM(source1,source2,source3))))}`
    },
    "Total Vehicle Miles": {
      headerFormula: `={"Total Vehicle Miles";MAP(formulaRunReviewOdoStart,formulaRunReviewOdoEnd,LAMBDA(startOdo,endOdo,IF(COUNTA(startOdo,endOdo) < 2,"",endOdo-startOdo)))}`
    },
    "Total Deadhead Miles": {
      headerFormula: `={"Total Vehicle Miles";MAP(formulaRunReviewOdoStart,formulaRunReviewOdoEnd,LAMBDA(startOdo,endOdo,IF(COUNTA(startOdo,endOdo) < 2,"",endOdo-startOdo)))}`
    },
    "Revenue Miles": {
      headerFormula: `={"Revenue Miles";MAP(formulaRunReviewTotalVehicleMiles,formulaRunReviewTotalDeadheadMiles,LAMBDA(vehicleMiles,deadheadMiles,IF(COUNTBLANK(vehicleMiles,deadheadMiles) > 0,"",vehicleMiles-deadheadMiles)))}`
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