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
  "Customers": {
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
