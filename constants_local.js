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
  "lookupDispatchIds": {
    "sheetName":"Lookups",
    "headerName":"Dispatch IDs"
  },
  "lookupFares": {
    "sheetName":"Lookups",
    "headerName":"Fares"
  },
  "lookupAddressShortNames": {
    "sheetName":"Lookups",
    "headerName":"Address Short Names"
  },
  "localCodeExpandAddress1": {
    "sheetName":"Trips",
    "headerName":"|PU|"
  },
  "codeFillHoursAndMiles001": {
    "sheetName":"Dispatch",
    "headerName":"PU Address"
  },
  "codeFillHoursAndMiles002": {
    "sheetName":"Dispatch",
    "headerName":"DO Address"
  },
  "codeFormatAddress001": {
    "sheetName":"Dispatch",
    "headerName":"PU Address"
  },
  "codeFormatAddress002": {
    "sheetName":"Dispatch",
    "headerName":"DO Address"
  },
  "codeFillRequestCells001": {
    "sheetName":"Dispatch",
    "headerName":"Customer Name and ID"
  },
  "codeTripActionButton001": {
    "sheetName":"Dispatch",
    "headerName":"|Go|"
  },
  "codeUpdateTripTimes001": {
    "sheetName":"Dispatch",
    "headerName":"PU Time"
  },
  "codeUpdateTripTimes002": {
    "sheetName":"Dispatch",
    "headerName":"DO Time"
  },
  "codeUpdateTripTimes003": {
    "sheetName":"Dispatch",
    "headerName":"Appt Time"
  },
  "codeUpdateTripVehicle001": {
    "sheetName":"Dispatch",
    "headerName":"Driver ID"
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
  "lookupDispatchIds": {
    "sheetName":"Lookups",
    "headerName":"Dispatch IDs"
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
  "localCodeExpandAddress3": {
    "sheetName":"Dispatch",
    "headerName":"|PU|"
  },
  "localCodeExpandAddress4": {
    "sheetName":"Dispatch",
    "headerName":"|DO|"
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
  "formulaRunArchiveFareRevenue": {
    "sheetName":"Run Archive",
    "headerName":"Fare Revenue"
  },
  "formulaRunArchiveDonationRevenue": {
    "sheetName":"Run Archive",
    "headerName":"Donation Revenue"
  },
  "formulaRunArchiveTicketRevenue": {
    "sheetName":"Run Archive",
    "headerName":"Ticket Revenue"
  },
  "queryAddresses": {
    "sheetName":"Addresses",
    "headerName":"Short Name"
  },
  "formulaDispatchTripDate": {
    "sheetName":"Dispatch",
    "headerName":"Trip Date"
  },
  "formulaDispatchPuTime": {
    "sheetName":"Dispatch",
    "headerName":"PU Time"
  },
  "formulaDispatchDoTime": {
    "sheetName":"Dispatch",
    "headerName":"DO Time"
  },
  "formulaDispatchDriverId": {
    "sheetName":"Dispatch",
    "headerName":"Driver ID"
  },
  "formulaDispatchVehicleId": {
    "sheetName":"Dispatch",
    "headerName":"Vehicle ID"
  },
  "formulaDispatchRunId": {
    "sheetName":"Dispatch",
    "headerName":"Run ID"
  },
}

const localNamedRangesToRemove = [
  "queryServiceEndDate",
  "queryServiceId"
]

const localSheetsToRemove = ["Services"]
const localSheets = [
  "Dispatch"
]
const localSheetsWithHeaders = [
  "Dispatch"
]

const localColumnsToRemove = {}
const localColumns = {
  "Customers": {
    "Other Phone #": {},
    "Birthdate": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Default Rider Type": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRiderTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Default Fare": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupFareTypes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid fare type.",
      }
    }
  },
  "Lookups": {
    "Rider Types": {},
    "Mobility Factors": {},
    "Fare Types": {},
    "Dispatch IDs": {},
    "Address Short Names": {
      headerFormula: `={"Address Short Names";QUERY(queryAddresses,"SELECT Col1 WHERE Col1 IS NOT NULL ORDER BY Col1",0)}`,
    }
  },
  "Trips": {
    "Day of Week": {
      headerFormula: `={"Day of Week";arrayformula(if(ISBLANK(A2:A), "", (TEXT(A2:A, "DDDD"))))}`
    },
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
    },
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
        checkedValue: "SDR",
        allowInvalid: false,
      },
    },
    "One Way": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "ONE WAY",
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
      }
    },
    "Dispatch ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDispatchIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid dispatch ID.",
      }
    }
  },
  "Lookups": {
    "Rider Types": {},
    "Mobility Factors": {},
    "Fare Types": {},
    "Dispatch IDs": {},
    "Address Short Names": {
      headerFormula: `={"Address Short Names";QUERY(queryAddresses,"SELECT Col1 WHERE Col1 IS NOT NULL ORDER BY Col1",0)}`,
    }
  },
  "Dispatch": {
    "Trip Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Customer Name and ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupCustomerNames",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid customer name and ID.",
      },
    },
    "|Action|": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "Add return trip",
          "Add stop",
        ],
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid action.",
      },
    },
    "|Go|": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        allowInvalid: true,
      }
    },
    "Trip Result": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripResults",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip result.",
      },
    },
    "Earliest PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Latest PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "DO Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Appt Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "PU Address": {},
    "DO Address": {},
    "Driver ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDriverIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid driver ID.",
      },
    },
    "Vehicle ID": {
      numberFormat: '@',
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "|Run OK?|": {
      headerFormula: `={"|Run OK?|";MAP(formulaDispatchTripDate, formulaDispatchPuTime, formulaDispatchDoTime, formulaDispatchDriverId, formulaDispatchVehicleId, formulaDispatchRunId, LAMBDA(TripDate,TripPuTime,TripDoTime,TripDriverId,TripVehicleID,TripRunId, QUERY_RUN_MATCH_COUNT(TripDate,TripPuTime,TripDoTime,TripDriverId,TripVehicleID,TripRunId,formulaRunsSheet)))}`
    },
    "Run ID": {},
    "Service ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupServiceIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid service ID.",
      },
    },
    "Guests": {
      dataValidation: {
        criteriaType: "NUMBER_GREATER_THAN_OR_EQUAL_TO",
        args: [0],
        allowInvalid: false,
        helpText: "Value must be the number of guests (0 or more).",
      },
    },
    "Mobility Factors": {},
    "Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip purpose.",
      },
    },
    "Notes": {},
    "Est Hours": {
      numberFormat: "0.00"
    },
    "Est Miles": {},
    "Trip ID": {},
    "Customer ID": {},
    "Day of Week": {
      headerFormula: `={"Day of Week";arrayformula(if(ISBLANK(A2:A), "", (TEXT(A2:A, "DDDD"))))}`
    },
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
    },
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
        checkedValue: "SDR",
        allowInvalid: false,
      },
    },
    "One Way": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "ONE WAY",
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
      }
    },
    "Dispatch ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDispatchIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid dispatch ID.",
      }
    }
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
    "Prepaid Revenue": {
      numberFormat: "$0.00"
    },
    "Total Completed Trip Fares": {
      headerFormula: `={"Total Completed Trip Fares";MAP(formulaRunReviewRunDate,formulaRunReviewDriverId,formulaRunReviewVehicleId,formulaRunReviewRunId,LAMBDA(tripDate,driverId,vehicleId,runId,QUERY_TRIP_FARE_SUM(tripDate,driverId,vehicleID,runId,formulaTripReviewCoreData,formulaTripReviewCoreHeaders)))}`
    },
    "Total Revenue": {
      headerFormula: `={"Total Revenue";MAP(formulaRunReviewFareRevenue,formulaRunReviewDonationRevenue,formulaRunReviewTicketRevenue,LAMBDA(source1,source2,source3,IF(COUNTBLANK(source1,source2,source3)>0,"",SUM(source1,source2,source3))))}`
    },
  },
  "Trip Review": {
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
        checkedValue: "SDR",
        allowInvalid: false,
      },
    },
    "One Way": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "ONE WAY",
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
      }
    },
    "Dispatch ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDispatchIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid dispatch ID.",
      }
    }
  },
  "Trip Archive": {
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
        checkedValue: "SDR",
        allowInvalid: true,
      },
    },
    "One Way": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "ONE WAY",
        allowInvalid: false,
      },
    },
    "Rider Type": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRiderTypes",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid rider type.",
      },
    },
    "Mobility Factors": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupMobilityFactors",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid mobility factor.",
      },
    },
    "Fare": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupFareTypes",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid fare type.",
      }
    },
    "Dispatch ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDispatchIds",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid dispatch ID.",
      }
    },
    "Trip Sequence": {
      headerFormula: `={"Trip Sequence";MAP(formulaTripArchiveTripDate, formulaTripArchiveCustomerId, formulaTripArchivePuTime, LAMBDA(tripDate,customerId, PuTime, IF(ISBLANK(tripDate),"",COUNTIFS(formulaTripArchiveTripDate, tripDate, formulaTripArchiveCustomerId, customerId, formulaTripArchivePuTime, "<" & PuTime)+1)))}`
    },
    "Trip Reverse Sequence": {
      headerFormula: `={"Trip Reverse Sequence";MAP(formulaTripArchiveTripDate, formulaTripArchiveCustomerId, formulaTripArchivePuTime, LAMBDA(tripDate,customerId, PuTime, IF(ISBLANK(tripDate),"",COUNTIFS(formulaTripArchiveTripDate, tripDate, formulaTripArchiveCustomerId, customerId, formulaTripArchivePuTime, ">" & PuTime)+1)))}`
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
    "Prepaid Revenue": {
      numberFormat: "$0.00"
    },
    "Total Revenue": {
      headerFormula: `={"Total Revenue";MAP(formulaRunArchiveFareRevenue,formulaRunArchiveDonationRevenue,formulaRunArchiveTicketRevenue,LAMBDA(source1,source2,source3,IF(COUNTA(source1,source2,source3) < 3,"",SUM(source1,source2,source3))))}`
    }
  }
}
