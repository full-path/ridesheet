const localNamedRangesToRemove = [
  "codeFillHoursAndMiles1",
  "codeFillHoursAndMiles2",
  "codeFillHoursAndMiles3",
  "codeFillHoursAndMiles4",
  "codeFillRequestCells1",
  "codeFormatAddress1",
  "codeFormatAddress2",
  "codeFormatAddress3",
  "codeFormatAddress4",
  "codeFormatAddress5",
  "codeFormatAddress6",
  "codeFormatAddress7",
  "codeScanForDuplicates1",
  "codeSetCustomerKey1",
  "codeSetCustomerKey2",
  "codeSetCustomerKey3",
  "codeTripActionButton1",
  "codeUpdateTripVehicle",
  "codeUpdateTripTimes1",
  "codeUpdateTripTimes2",
  "codeUpdateTripTimes3",
  "codeCheckSourceOnShare",
  "codeVerifySourceOnEdit",
  "lookupCustomerNames",
  "lookupDriverIds",
  "lookupVehicleIds",
  "lookupServiceIds",
  "lookupTripResults",
  "queryCustomerNameAndId",
  "queryCustomerId",
  "queryCustomerEndDate",
  "queryServiceId",
  "queryServiceEndDate",
  "queryDriverId",
  "queryDriverEndDate",
  "queryVehicleID",
  "queryVehicleEndDate"
]

const localNamedRanges = {
 "localCodeReferralActionButton1": {
    "sheetName":"Call Log",
    "headerName":"Go"
  },
 "localCodeFillReferralCells1": {
    "sheetName":"Call Log",
    "headerName":"Call Date"
  },
  "lookupStaffIds": {
    "sheetName": "Lookups",
    "headerName": "Staff ID"
  },
  "lookupCounties": {
    "sheetName": "Lookups",
    "headerName": "Counties"
  },
  "lookupAgencies": {
    "sheetName": "Lookups",
    "headerName": "Agencies"
  },
  "lookupTDSAgencies": {
    "sheetName": "Lookups",
    "headerName": "TDS Agencies"
  },
  "lookupCities": {
    "sheetName": "Lookups",
    "headerName": "Cities"
  },
  "lookupGaps": {
    "sheetName": "Lookups",
    "headerName": "Gaps"
  },
  "lookupRaces": {
    "sheetName": "Lookups",
    "headerName": "Races"
  },
  "lookupEthnicities": {
    "sheetName": "Lookups",
    "headerName": "Ethnicities"
  },
  "lookupGenders": {
    "sheetName": "Lookups",
    "headerName": "Genders"
  }
}

const localSheetsToRemove = ["Trips", "Trip Review", "Outside Trips", "Customers", "Sent Trips", "Run Review", "Run Archive", "Trip Archive", "Runs", "Outside Runs", "Vehicles", "Drivers", "Services"]
const localSheets = ["Call Log", "TDS Referrals"]
const localSheetsWithHeaders = ["Call Log", "TDS Referrals", "Lookups"]
const localColumnsToRemove = {
  "Lookups": [
    "Customer Names and IDs",
    "Driver IDs",
    "Vehicle IDs",
    "Service IDs",
    "Trip Results"
  ]
}

const localColumns = {
  "Call Log": {
    "Call Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Staff ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupStaffIds",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid staff ID.",
      },
    },
    "Customer First Name": {},
    "Customer Last Name": {},
    "Phone Number": {},
    "Email": {
      dataValidation: {
        criteriaType: "TEXT_IS_VALID_EMAIL",
        allowInvalid: false,
        helpText: "Value must be a valid email address.",
      }
    },
    "Home Address": {},
    "County": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupCounties",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid county.",
      },
    },
    "City": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupCities",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid city.",
      },
    },
    "Over 60?": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "Yes",
          "No",
          "Unknown"
        ],
        showDropdown: true,
        allowInvalid: false,
        helpText: "",
      }
    },
    "Language": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "English",
          "Spanish",
          "Other"
        ],
        showDropdown: true,
        allowInvalid: true,
        helpText: "",
      }
    },
    "Other Language": {},
    "Veteran?": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "Yes",
          "No",
          "Unknown"
        ],
        showDropdown: true,
        allowInvalid: false,
        helpText: "",
      }
    },
    "Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid trip purpose.",
      },
    },
    "Follow-up Needed": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: true,
      }
    },
    "Call Notes": {},
    "Non-TDS Referral 1": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupAgencies",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid agency.",
      },
    },
    "Non-TDS Referral 2": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupAgencies",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid agency.",
      },
    },
    "Referral/Other Notes": {},
    "Gaps": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupGaps",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid gap.",
      },
    },
    "Make TDS Referral To": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTDSAgencies",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid TDS agency.",
      }
    },
    "Go": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: true,
      }
    },
    "TDS Referrals": {},
    "Call ID": {}
  },
  "TDS Referrals": {
    "Call ID": {},
    "Customer ID": {},
    "Agency": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTDSAgencies",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid TDS agency.",
      }
    },
    "Referral Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Customer First Name": {},
    "Customer Nickname": {},
    "Customer Middle Name": {},
    "Customer Last Name": {},
    "Gender": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupGenders",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid gender.",
      },
    },
    "Home Phone Number": {},
    "Mobile Phone Number": {},
    "Email": {
      dataValidation: {
        criteriaType: "TEXT_IS_VALID_EMAIL",
        allowInvalid: false,
        helpText: "Value must be a valid email address.",
      }
    },
    "Home Address": {},
    "Date of Birth": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Low Income?": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "Yes",
          "No",
          "Unknown"
        ],
        showDropdown: true,
        allowInvalid: false,
        helpText: "",
      }
    },
    "Disability?": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "Yes",
          "No",
          "Unknown"
        ],
        showDropdown: true,
        allowInvalid: false,
        helpText: "",
      }
    },
    "Veteran?": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "Yes",
          "No",
          "Unknown"
        ],
        showDropdown: true,
        allowInvalid: false,
        helpText: "",
      }
    },
    "Language": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: [
          "English",
          "Spanish",
          "Other"
        ],
        showDropdown: true,
        allowInvalid: true,
        helpText: "",
      }
    },
    "Race": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupRace",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid race.",
      },
    },
    "Ethnicity": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupEthnicities",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid ethnicity.",
      },
    },
    "Mailing Address": {},
    "Funding Type?": {},
    "Funding Entity": {},
    "Billing Information": {},
    "Caregiver Contact Information": {},
    "Emergency Phone Number": {},
    "Emergency Contact Name": {},
    "Emergency Contact Relationship": {},
    "Comments About Care Required": {},
    "Referral Notes": {},
    "Referral Sent Timestamp": {},
    "Referral Response Timestamp": {},
    "Referral Response": {},
    "Response Notes": {},
    "Referral ID": {}
  },
  "Lookups": {
    "Trip Purposes": {},
    "Staff ID": {},
    "Counties": {},
    "Agencies": {},
    "TDS Agencies": {},
    "Cities": {},
    "Gaps": {},
    "Races": {},
    "Ethnicities": {}
  }
}
