const debugLogging                       = false
const allowPropDescriptionEdits          = false

const errorBackgroundColor               = "#f4cccc"
const defaultBackgroundColor             = "#ffffff"
const headerBackgroundColor              = "#fff2cc"
const highlightBackgroundColor           = "#ffff00"

// Config for the state of Oregon
const defaultLocalTimeZone               = "America/Los_Angeles"
const defaultGeocoderBoundSwLatitude     = 41.997013
const defaultGeocoderBoundSwLongitude    = -124.560974
const defaultGeocoderBoundNeLatitude     = 46.299097
const defaultGeocoderBoundNeLongitude    = -116.463363

const defaultDwellTimeInMinutes          = 10
const defaultTripPaddingPerHourInMinutes = 5

const defaultSheets = [
  "Customers",
  "Trips",
  "Runs",
  "Trip Review",
  "Run Review",
  "Trip Archive",
  "Run Archive",
  "Vehicles",
  "Drivers",
  "Services",
  "Lookups",
  "Document Properties",
  "Debug Log",
  "Addresses"
]

const sheetsWithHeaders = [
  "Customers",
  "Trips",
  "Runs",
  "Trip Review",
  "Run Review",
  "Trip Archive",
  "Run Archive",
  "Vehicles",
  "Drivers",
  "Services",
  "Lookups"
]

const defaultDocumentProperties = {
  lastCustomerID_: {
    type: "number",
    value: 0,
    description: "The value of the last set customer ID."
  },
  driverManifestFolderId: {
    type: "string",
    value: "Enter ID here",
    description: "The ID of the folder where newly created trip manifests will be saved."
  },
  driverManifestTemplateDocId: {
    type: "string",
    value: "Enter ID here",
    description: "The document ID of the Google Doc you'll be using as your manifest template."
  },
  geocoderBoundNeLatitude: {
    type: "number",
    value: 46.299097,
    description: "The north latitude of the box where Google Maps gives extra preference when geocoding addresses."
  },
  geocoderBoundNeLongitude: {
    type: "number",
    value: -116.463363,
    description: "The east longitude of the box where Google Maps gives extra preference when geocoding addresses."
  },
  geocoderBoundSwLatitude: {
    type: "number",
    value: 41.997013,
    description: "The south latitude of the box where Google Maps gives extra preference when geocoding addresses."
  },
  geocoderBoundSwLongitude: {
    type: "number",
    value: -124.560974,
    description: "The west longitude of the box where Google Maps gives extra preference when geocoding addresses."
  },
  localTimeZone:  {
    type: "string",
    value: "America/Los_Angeles",
    description: "The local time zone. Use one of the TZ database names found here: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones"
  },
  dwellTimeInMinutes: {
    type: "number",
    value: 10,
    description: "The length of time in minutes added to the journey time to account for the time it takes to pick up and drop off a rider"
  },
  tripPaddingPerHourInMinutes: {
    type: "number",
    value: 5,
    description: "The length of time in minutes added to each hour of estimated travel time to account for weather, traffic, or other possible delays"
  },
  dropOffToAppointmentTimeInMinutes: {
    type: "number",
    value: 10,
    description: "The length of time in minutes between the drop off time and the appointment time"
  },
  tripReviewCompletedTripResults: {
    type: "array",
    value: ["Completed"],
    description: "The values of trip results where other required fields must be filled in."
  },
  tripReviewRequiredFields: {
    type: "array",
    value: ["Trip Result", "Actual PU Time", "Actual DO Time"],
    description: "The names of trip columns that must have data in them in order to be archived."
  },
  runUserReviewRequiredFields: {
    type: "array",
    value: [],
    description: "The names of run columns that must have data in them in order to for RideSheet to calculate deadhead or other run information."
  },
  runFullReviewRequiredFields: {
    type: "array",
    value: [],
    description: "The names of run columns that must have data in them in order to be archived."
  },
  providerName: {
    type: "string",
    value: "Enter provider name here",
    description: "Name of the agency using this RideSheet document"
  },
  logLevel: {
    type: "string",
    value: "normal",
    description: "Set logging level to normal or verbose"
  },
  defaultStayDuration: {
    type: "number",
    value: 60,
    description: "When creating a next leg or return trip, this is the length of time in minutes to set as the duration between rider dropoff or appt time and the pickup time of the next trip. Set to -1 (negative one) to keep the pickup time for the new trip blank"
  }
}

const defaultColumns = {
  "Customers": {
    "Customer Name and ID": {},
    "Customer ID": {},
    "Customer First Name": {},
    "Customer Last Name": {},
    "Date of Birth": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Phone Number": {},
    "Email": {
      dataValidation: {
        criteriaType: "TEXT_IS_VALID_EMAIL",
        allowInvalid: false,
        helpText: "Value must be a valid email address.",
      },
    },
    "Mailing Address": {},
    "Home Address": {},
    "Default PU Address": {},
    "Default DO Address": {},
    "Default Service ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupServiceIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid service ID.",
      },
    },
    "Default Service Level": {},
    "Customer Manifest Notes": {},
    "Customer Private Notes": {},
    "Customer Start Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Customer End Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
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
    "Default Mobility Factors": {}
  },
  "Trips": {
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
    "Customer ID": {}
  },
  "Runs": {
    "Run Date": {
      numberFormat: "MM/dd/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
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
    "Run ID": {},
    "Scheduled Start Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
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
    "Scheduled End Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    }
  },
  "Trip Review": {
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
    "Trip Result": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripResults",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip result.",
      },
    },
    "Actual PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual DO Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Start Odo": {},
    "End Odo": {},
    "PU Time": {
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
    "Driver ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDriverIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid driver ID.",
      },
    },
    "Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip type.",
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
    "Run ID": {},
    "PU Address": {},
    "DO Address": {},
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
    "Notes": {},
    "Est Hours": {
      numberFormat: "0.00"
    },
    "Est Miles": {},
    "Trip ID": {},
    "Customer ID": {},
    "Review TS": {
      numberFormat: "m/d/yyyy h:mm:ss"
    },
  },
  "Run Review": {
    "Run Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
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
    "Run ID": {},
    "Scheduled Start Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Scheduled End Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual Start Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual End Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Odometer Start": {},
    "Odometer End": {},
    "Break Time in Minutes": {
      numberFormat: "0",
      dataValidation: {
        criteriaType: "NUMBER_BETWEEN",
        args: [0,500],
        helpText: "Value must be the duration of all breaks in the run, up to 500 minutes"
      },
    },
    "Review TS": {
      numberFormat: "m/d/yyyy h:mm:ss"
    },
    "Total Vehicle Miles": {
      headerFormula: `={"Total Vehicle Miles";MAP(formulaRunReviewOdoStart,formulaRunReviewOdoEnd,LAMBDA(startOdo,endOdo,IF(COUNTBLANK(startOdo,endOdo)>0,"",endOdo-startOdo)))}`
    },
    "Total Deadhead Miles": {
      headerFormula: `={"Total Deadhead Miles";MAP(formulaRunReviewStartingDeadheadMiles,formulaRunReviewEndingDeadheadMiles,LAMBDA(start,end,IF(COUNTBLANK(start,end)>0,"",start+end)))}`
    },
    "Revenue Miles": {
      headerFormula: `={"Revenue Miles";MAP(formulaRunReviewTotalVehicleMiles,formulaRunReviewTotalDeadheadMiles,LAMBDA(vehicleMiles,deadheadMiles,IF(COUNTBLANK(vehicleMiles,deadheadMiles) > 0,"",vehicleMiles-deadheadMiles)))}`
    },
    "Total Vehicle Hours": {
      headerFormula: `={"Total Vehicle Hours";MAP(formulaRunReviewTimeStart,formulaRunReviewTimeEnd,LAMBDA(start,end,IF(COUNTBLANK(start,end)>0,"",end-start)))}`,
      numberFormat: "[h]:mm"
    },
    "Total Deadhead Hours": {
      headerFormula: `={"Total Non-Revenue Hours";MAP(formulaRunReviewStartingDeadheadHours,formulaRunReviewEndingDeadheadHours,formulaRunReviewBreakTime,LAMBDA(start,end,middle,IF(COUNTBLANK(start,end,middle)>0,"",start+end+(middle/1440))))}`,
      numberFormat: "[h]:mm"
    },
    "Revenue Hours": {
      headerFormula: `={"Revenue Hours";MAP(formulaRunReviewTotalVehicleHours,formulaRunReviewTotalNonRevenueHours,LAMBDA(vehicleHours,deadheadHours,IF(COUNTBLANK(vehicleHours,deadheadHours)>0,"",vehicleHours-deadheadHours)))}`,
      numberFormat: "[h]:mm"
    },
    "Starting Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Ending Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Starting Deadhead Hours": {
      numberFormat: "[h]:mm"
    },
    "Ending Deadhead Hours": {
      numberFormat: "[h]:mm"
    },
    "Vehicle Garage Address": {},
    "First PU Address": {},
    "Last DO Address": {}
  },
  "Trip Archive": {
    "Trip Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    },
    "Customer Name and ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupCustomerNames",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid customer name and ID.",
      },
    },
    "Trip Result": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripResults",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid trip result.",
      },
    },
    "Actual PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual DO Time": {
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
    "Driver ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDriverIds",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid driver ID.",
      },
    },
    "Vehicle ID": {
      numberFormat: '@',
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "Run ID": {},
    "PU Address": {},
    "DO Address": {},
    "Service ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupServiceIds",
        showDropdown: true,
        allowInvalid: true,
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
        showDropdown: false,
        allowInvalid: false,
        helpText: "Value must be a valid trip type.",
      },
    },
    "Notes": {},
    "Est Hours": {
      numberFormat: "0.00"
    },
    "Est Miles": {},
    "Trip ID": {},
    "Calendar ID": {},
    "Customer ID": {},
    "Review TS": {
      numberFormat: "m/d/yyyy h:mm:ss"
    },
    "Archive TS": {
      numberFormat: "m/d/yyyy h:mm:ss"
    },
  },
  "Run Archive": {
    "Run Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    },
    "Driver ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupDriverIds",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid driver ID.",
      },
    },
    "Vehicle ID": {
      numberFormat: '@',
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "Run ID": {},
    "Scheduled Start Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Scheduled End Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual Start Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual End Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Odometer Start": {},
    "Odometer End": {},
    "Break Time in Minutes": {
      numberFormat: "0",
      dataValidation: {
        criteriaType: "NUMBER_BETWEEN",
        args: [0,500],
        helpText: "Value must be the duration of all breaks in the run, up to 500 minutes"
      },
    },
    "Review TS": {
      numberFormat: "m/d/yyyy h:mm:ss"
    },
    "Archive TS": {
      numberFormat: "m/d/yyyy h:mm:ss"
    },
    "Total Vehicle Miles": {
      headerFormula: `={"Total Vehicle Miles";MAP(formulaRunArchiveOdoStart,formulaRunArchiveOdoEnd,LAMBDA(startOdo,endOdo,IF(COUNTBLANK(startOdo,endOdo)>0,"",endOdo-startOdo)))}`
    },
    "Total Deadhead Miles": {
      headerFormula: `={"Total Deadhead Miles";MAP(formulaRunArchiveStartingDeadheadMiles,formulaRunArchiveEndingDeadheadMiles,LAMBDA(start,end,IF(COUNTBLANK(start,end)>0,"",start+end)))}`
    },
    "Revenue Miles": {
      headerFormula: `={"Revenue Miles";MAP(formulaRunArchiveTotalVehicleMiles,formulaRunArchiveTotalDeadheadMiles,LAMBDA(vehicleMiles,deadheadMiles,IF(COUNTBLANK(vehicleMiles,deadheadMiles)>0,"",vehicleMiles-deadheadMiles)))}`
    },
    "Total Vehicle Hours": {
      headerFormula: `={"Total Vehicle Hours";MAP(formulaRunArchiveTimeStart,formulaRunArchiveTimeEnd,LAMBDA(start,end,IF(COUNTBLANK(start,end)>0,"",end-start)))}`,
      numberFormat: "[h]:mm"
    },
    "Total Non-Revenue Hours": {
      headerFormula: `={"Total Non-Revenue Hours";MAP(formulaRunArchiveStartingDeadheadHours,formulaRunArchiveEndingDeadheadHours,formulaRunArchiveBreakTime,LAMBDA(start,end,middle,IF(COUNTBLANK(start,end,middle)>0,"",start+end+(middle/1440))))}`,
      numberFormat: "[h]:mm"
    },
    "Revenue Hours": {
      headerFormula: `={"Revenue Hours";MAP(formulaRunArchiveTotalVehicleHours,formulaRunArchiveTotalNonRevenueHours,LAMBDA(vehicleHours,deadheadHours,IF(COUNTBLANK(vehicleHours,deadheadHours)>0,"",vehicleHours-deadheadHours)))}`,
      numberFormat: "[h]:mm"
    },
    "Starting Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Ending Deadhead Miles": {
      numberFormat: "0.0"
    },
    "Starting Deadhead Hours": {
      numberFormat: "[h]:mm"
    },
    "Ending Deadhead Hours": {
      numberFormat: "[h]:mm"
    },
    "Vehicle Garage Address": {},
    "First PU Address": {},
    "Last DO Address": {}
  },
  "Vehicles": {
    "Vehicle ID": {
      numberFormat: '@',
    },
    "Vehicle Name": {},
    "Garage Address": {},
    "Seating Capacity": {},
    "Wheelchair Capacity": {},
    "Scooter Capacity": {},
    "Has Ramp": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "HAS RAMP",
        allowInvalid: true,
      },
    },
    "Has Lift": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "HAS LIFT",
        allowInvalid: true,
      },
    },
    "Vehicle Start Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    },
    "Vehicle End Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    }
  },
  "Drivers": {
    "Driver ID": {},
    "Driver Name": {},
    "Driver Email": {
      dataValidation: {
        criteriaType: "TEXT_IS_VALID_EMAIL",
        allowInvalid: false,
        helpText: "Value must be a valid email address.",
      }
    },
    "Default Vehicle ID": {
      numberFormat: '@',
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "Driver Calendar ID": {},
    "Driver Start Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    },
    "Driver End Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    }
  },
  "Services": {
    "Service ID": {},
    "Service Name": {},
    "Service Funder": {},
    "Service Start Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    },
    "Service End Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      }
    }
  },
  "Addresses": {
    "Short Name": {},
    "Address": {},
  },
  "Lookups": {
    "Customer Names and IDs": {
      headerFormula: `={"Customer Names and IDs";QUERY({queryCustomerNameAndId,queryCustomerEndDate},"SELECT Col1 WHERE Col1 IS NOT NULL AND Col2 IS NULL ORDER BY Col1",0)}`
    },
    "Driver IDs": {
      headerFormula: `={"Driver IDs";QUERY({queryDriverId,queryDriverEndDate},"SELECT Col1 WHERE Col1 IS NOT NULL AND Col2 IS NULL ORDER BY Col1",0)}`
    },
    "Vehicle IDs": {
      headerFormula: `={"Vehicle IDs";QUERY({queryVehicleID,queryVehicleEndDate},"SELECT Col1 WHERE Col1 IS NOT NULL AND Col2 IS NULL ORDER BY Col1",0)}`
    },
    "Service IDs": {
      headerFormula: `={"Service IDs";QUERY({queryServiceId,queryServiceEndDate},"SELECT Col1 WHERE Col1 IS NOT NULL AND Col2 IS NULL ORDER BY Col1",0)}`},
    "Trip Purposes": {},
    "Trip Results": {}
  }
}

const defaultNamedRanges = {
  "codeFillHoursAndMiles1": {
    "sheetName":"Trips",
    "headerName":"PU Address"
  },
  "codeFillHoursAndMiles2": {
    "sheetName":"Trips",
    "headerName":"DO Address"
  },
  "codeFillHoursAndMiles3": {
    "sheetName":"Trip Review",
    "headerName":"PU Address"
  },
  "codeFillHoursAndMiles4": {
    "sheetName":"Trip Review",
    "headerName":"DO Address"
  },
  "codeFillRequestCells1": {
    "sheetName":"Trips",
    "headerName":"Customer Name and ID"
  },
  "codeFormatAddress1": {
    "sheetName":"Trips",
    "headerName":"PU Address"
  },
  "codeFormatAddress2": {
    "sheetName":"Trips",
    "headerName":"DO Address"
  },
  "codeFormatAddress3": {
    "sheetName":"Customers",
    "headerName":"Home Address"
  },
  "codeFormatAddress4": {
    "sheetName":"Customers",
    "headerName":"Default PU Address"
  },
  "codeFormatAddress5": {
    "sheetName":"Customers",
    "headerName":"Default DO Address"
  },
  "codeFormatAddress6": {
    "sheetName":"Vehicles",
    "headerName":"Garage Address"
  },
  "codeFormatAddress7": {
    "sheetName":"Trip Review",
    "headerName":"PU Address"
  },
  "codeFormatAddress8": {
    "sheetName":"Trip Review",
    "headerName":"DO Address"
  },
  "codeFormatAddress9": {
    "sheetName":"Addresses",
    "headerName":"Address"
  },
  "codeScanForDuplicates1": {
    "sheetName":"Customers",
    "headerName":"Customer ID"
  },
  "codeSetCustomerKey1": {
    "sheetName":"Customers",
    "headerName":"Customer First Name"
  },
  "codeSetCustomerKey2": {
    "sheetName":"Customers",
    "headerName":"Customer Last Name"
  },
  "codeSetCustomerKey3": {
    "sheetName":"Customers",
    "headerName":"Customer ID"
  },
  "codeTripActionButton1": {
    "sheetName":"Trips",
    "headerName":"Go"
  },
  "codeUpdateTripVehicle": {
    "sheetName":"Trips",
    "headerName":"Driver ID"
  },
  "codeUpdateTripTimes1": {
    "sheetName":"Trips",
    "headerName":"PU Time"
  },
  "codeUpdateTripTimes2": {
    "sheetName":"Trips",
    "headerName":"DO Time"
  },
  "codeUpdateTripTimes3":{
    "sheetName":"Trips",
    "headerName":"Appt Time"
  },
  "lookupCustomerNames": {
    "sheetName":"Lookups",
    "headerName":"Customer Names and IDs"
  },
  "lookupDriverIds": {
    "sheetName":"Lookups",
    "headerName":"Driver IDs"
  },
  "lookupVehicleIds": {
    "sheetName":"Lookups",
    "headerName":"Vehicle IDs"
  },
  "lookupServiceIds": {
    "sheetName":"Lookups",
    "headerName":"Service IDs"
  },
  "lookupTripPurposes": {
    "sheetName":"Lookups",
    "headerName":"Trip Purposes"
  },
  "lookupTripResults": {
    "sheetName":"Lookups",
    "headerName":"Trip Results"
  },
  "queryCustomerNameAndId": {
    "sheetName":"Customers",
    "headerName":"Customer Name and ID"
  },
  "queryCustomerId": {
    "sheetName":"Customers",
    "headerName":"Customer ID"
  },
  "queryCustomerEndDate": {
    "sheetName":"Customers",
    "headerName":"Customer End Date"
  },
  "queryServiceId": {
    "sheetName":"Services",
    "headerName":"Service ID"
  },
  "queryServiceEndDate": {
    "sheetName":"Services",
    "headerName":"Service End Date"
  },
  "queryDriverId": {
    "sheetName":"Drivers",
    "headerName":"Driver ID"
  },
  "queryDriverEndDate": {
    "sheetName":"Drivers",
    "headerName":"Driver End Date"
  },
  "queryVehicleID": {
    "sheetName":"Vehicles",
    "headerName":"Vehicle ID"
  },
  "queryVehicleEndDate": {
    "sheetName":"Vehicles",
    "headerName":"Vehicle End Date"
  },
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
  "formulaRunReviewRunDate": {
    "sheetName":"Run Review",
    "headerName":"Run Date"
  },
  "formulaRunsVehicleId": {
    "sheetName":"Runs",
    "headerName":"Vehicle ID"
  },
  "formulaRunsRunId": {
    "sheetName":"Runs",
    "headerName":"Run ID"
  },
  "formulaTripReviewCoreHeaders": {
    "sheetName":"Trip Review",
    "startHeaderName":"Trip Date",
    "endHeaderName":"Run ID",
    "headerOnly": true
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
  "formulaRunReviewBreakTime": {
    "sheetName":"Run Review",
    "headerName":"Break Time in Minutes"
  },
  "formulaRunReviewTimeStart": {
    "sheetName":"Run Review",
    "headerName":"Actual Start Time"
  },
  "formulaRunReviewTimeEnd": {
    "sheetName":"Run Review",
    "headerName":"Actual End Time"
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
  "formulaRunReviewTotalDeadheadMiles": {
    "sheetName":"Run Review",
    "headerName":"Total Deadhead Miles"
  },
  "formulaRunReviewTotalVehicleMiles": {
    "sheetName":"Run Review",
    "headerName":"Total Vehicle Miles"
  },
  "formulaRunReviewStartingDeadheadHours": {
    "sheetName":"Run Review",
    "headerName":"Starting Deadhead Hours"
  },
  "formulaRunReviewEndingDeadheadHours": {
    "sheetName":"Run Review",
    "headerName":"Ending Deadhead Hours"
  },
  "formulaRunReviewTotalNonRevenueHours": {
    "sheetName":"Run Review",
    "headerName":"Total Non-Revenue Hours"
  },
  "formulaRunReviewTotalVehicleHours": {
    "sheetName":"Run Review",
    "headerName":"Total Vehicle Hours"
  },
  "formulaRunArchiveRunDate": {
    "sheetName":"Run Archive",
    "headerName":"Run Date"
  },
  "formulaRunArchiveBreakTime": {
    "sheetName":"Run Archive",
    "headerName":"Break Time in Minutes"
  },
  "formulaRunArchiveOdoStart": {
    "sheetName":"Run Archive",
    "headerName":"Odometer Start"
  },
  "formulaRunArchiveOdoEnd": {
    "sheetName":"Run Archive",
    "headerName":"Odometer End"
  },
  "formulaRunArchiveTimeStart": {
    "sheetName":"Run Archive",
    "headerName":"Actual Start Time"
  },
  "formulaRunArchiveTimeEnd": {
    "sheetName":"Run Archive",
    "headerName":"Actual End Time"
  },
  "formulaRunArchiveStartingDeadheadMiles": {
    "sheetName":"Run Archive",
    "headerName":"Starting Deadhead Miles"
  },
  "formulaRunArchiveEndingDeadheadMiles": {
    "sheetName":"Run Archive",
    "headerName":"Ending Deadhead Miles"
  },
  "formulaRunArchiveTotalDeadheadMiles": {
    "sheetName":"Run Archive",
    "headerName":"Total Deadhead Miles"
  },
  "formulaRunArchiveTotalVehicleMiles": {
    "sheetName":"Run Archive",
    "headerName":"Total Vehicle Miles"
  },
  "formulaRunArchiveStartingDeadheadHours": {
    "sheetName":"Run Archive",
    "headerName":"Starting Deadhead Hours"
  },
  "formulaRunArchiveEndingDeadheadHours": {
    "sheetName":"Run Archive",
    "headerName":"Ending Deadhead Hours"
  },
  "formulaRunArchiveTotalNonRevenueHours": {
    "sheetName":"Run Archive",
    "headerName":"Total Non-Revenue Hours"
  },
  "formulaRunArchiveTotalVehicleHours": {
    "sheetName":"Run Archive",
    "headerName":"Total Vehicle Hours"
  },
  "formulaTripArchiveTripDate": {
    "sheetName":"Trip Archive",
    "headerName":"Trip Date"
  },
  "formulaTripArchiveCustomerId": {
    "sheetName":"Trip Archive",
    "headerName":"Customer ID"
  },
  "formulaTripArchivePuTime": {
    "sheetName":"Trip Archive",
    "headerName":"PU Time"
  }
}
