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
  "Sent Trips",
  "Trip Review",
  "Run Review",
  "Trip Archive",
  "Run Archive",
  "Vehicles",
  "Drivers",
  "Services",
  "Outside Trips",
  "Outside Runs",
  "Lookups",
  "Document Properties",
  "Debug Log"
]

const sheetsWithHeaders = [
  "Customers",
  "Trips",
  "Runs",
  "Sent Trips",
  "Trip Review",
  "Run Review",
  "Trip Archive",
  "Run Archive",
  "Vehicles",
  "Drivers",
  "Services",
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
  monthlyBackupFolderId: {
    type: "string",
    value: "Enter ID here",
    description: "The ID of the folder where monthly backups will be saved"
  },
  monthlyFileRetentionInDays: {
    type: "number",
    value: 365,
    description: "How many days monthly backups should be held onto before they're automatically deleted"
  },
  nightlyBackupFolderId: {
    type: "string",
    value: "Enter ID here",
    description: "The ID of the folder where nightly backups will be saved"
  },
  nightlyFileRetentionInDays: {
    type: "number",
    value: 90,
    description: "How many days weekly backups should be held onto before they're automatically deleted"
  },
  weeklyBackupFolderId: {
    type: "string",
    value: "Enter ID here",
    description: "The ID of the folder where nightly backups will be saved"
  },
  weeklyFileRetentionInDays: {
    type: "number",
    value: 180,
    description: "How many days weekly backups should be held onto before they're automatically deleted"
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
  runReviewRequiredFields: {
    type: "array",
    value: [],
    description: "The names of run columns that must have data in them in order to be archived."
  },
  calendarIdForUnassignedTrips: {
    type: "string",
    value: "Enter ID here",
    description: "The ID of the Google Calendar where trips without drivers are shown"
  },
  tripCalendarEntryTitleTemplate: {
    type: "string",
    value: "{Customer Name and ID}",
    description: "The template for calendar entries made for scheduled trips. Field names should be entered in braces like so: {Driver ID}"
  },
  providerName: {
    type: "string",
    value: "Enter provider name here",
    description: "Name of the agency using this RideSheet document"
  },
  notificationEmail: {
    type: "string",
    value: "",
    description: "Email address to use for notifications"
  },
  apiGetAccess: {
    type: "array",
    value: [
      {
        name: "Example agency name",
        url: "https://example.com",
        apiVersion: "v1",
        apiKey: "Enter key here",
        secret: "Enter secret here",
        hasRuns: true,
        hasTrips: true
      }
    ],
    description: "API information needed to connect to other agencies"    
  },
  apiGiveAccess: {
    type: "object",
    value: {
      Enter_agency_key_here: {
        name: "Example agency name with API access to data in this sheet",
          secret: "Enter secret here"
      }  
    },
    description: "API information needed to allow agencies to connect to this sheet. We recommend using https://www.uuidgenerator.net/ to generate API keys."    
  },
  apiShowMenuItems: {
    type: "boolean",
    value: false,
    description: "Show menu items for manually triggering API calls?"
  },
  logLevel: {
    type: "string",
    value: "normal",
    description: "Set logging level to normal or verbose"
  },
  extraHeaderNames: {
    type: "object",
    value: {
      Customers: [],
      Trips: [],
      Runs: [],
      "Sent Trips": [],
      "Outside Trips": [],
      "Outside Runs": [],
      "Trip Review": [],
      "Run Review": [],
      "Trip Archive": [],
      "Run Archive": [],
      Vehicles: [],
      Drivers: [],
      Services: []
    },
    description: "Show menu items for manually triggering API calls?"
  },
  configFolderId: {
    type: "string",
    value: "Enter ID here",
    description: "The ID of the folder where configuration files will be located."
  },
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
    "Default Mobility Factors": {},
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
    }
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
    "Action": {},
    "Go": {},
    "Share": {},
    "Declined By": {},
    "Trip Result": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripResults",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip result.",
      },
    },
    "Source": {},
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
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "Service ID": {      
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupServiceIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid service ID.",
      },
    },
    "Guests": {},
    "Mobility Factors": {},
    "Notes": {},
    "Est Hours": {
      numberFormat: "[h]:mm:ss"      
    },
    "Est Miles": {},
    "Driver Calendar ID": {},
    "Trip Event ID": {},
    "Trip ID": {},
    "Customer ID": {},
    "Shared": {}
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
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "Scheduled Start Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "First PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Last DO Time": {
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
    }
  },
  "Sent Trips": {
    "Claimed By": {},
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
    "Declined By": {},
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
    "Guests": {},
    "Mobility Factors": {},
    "Notes": {},
    "Est Hours": {
      numberFormat: "[h]:mm:ss"
    },
    "Est Miles": {},
    "Trip ID": {},
    "Customer ID": {}
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
    "Action": {
      dataValidation: {
        criteriaType: "VALUE_IN_LIST",
        values: ["Add return trip","Add stop"],
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid action.",
      },
    },
    "Go": {
        dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: false,
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
    "Share": {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: false,
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
    "Source": {},
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
    "Vehicle ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
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
    "Guests": {},
    "Mobility Factors": {},
    "Notes": {},
    "Est Hours": {
      numberFormat: "[h]:mm:ss"
    },
    "Est Miles": {},
    "Manifest ID": {},
    "Calendar ID": {},
    "Customer ID": {}
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
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "Scheduled Start Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "First PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Last DO Time": {
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
    "Odometer End": {}
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
    "Source": {},
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
    "Vehicle ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
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
    "Guests": {},
    "Mobility Factors": {},
    "Notes": {},
    "Est Hours": {
      numberFormat: "[h]:mm:ss"
    },
    "Est Miles": {},
    "Manifest ID": {},
    "Calendar ID": {},
    "Customer ID": {}
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
        allowInvalid: false,
        helpText: "Value must be a valid driver ID.",
      },
    },
    "Vehicle ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupVehicleIds",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid vehicle ID.",
      },
    },
    "First PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Last DO Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
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
    "Odometer End": {}
  },
  "Vehicles": {
    "Vehicle ID": {},
    "Vehicle Name": {},
    "Garage Address": {},
    "Seating Capacity": {},
    "Wheelchair Capacity": {},
    "Scooter Capacity": {},
    "Has Ramp": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: false,
      },
    },
    "Has Lift": {
      dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: false,
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
    "headerName":"Default DO Address"
  },
  "codeFormatAddress5": {
    "sheetName":"Vehicles",
    "headerName":"Garage Address"
  },
  "codeFormatAddress6": {
    "sheetName":"Trip Review",
    "headerName":"PU Address"
  },
  "codeFormatAddress7": {
    "sheetName":"Trip Review",
    "headerName":"DO Address"
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
  "codeCheckSourceOnShare": {
    "sheetName":"Trips",
    "headerName": "Share"
  },
  "codeVerifySourceOnEdit": {
    "sheetName":"Trips",
    "headerName":"Source"
  },
  "lookupCustomerNames": {
    "sheetName":"Lookups",
    "column":"A"
  },
  "lookupDriverIds": {
    "sheetName":"Lookups",
    "column":"B"
  },
  "lookupVehicleIds": {
    "sheetName":"Lookups",
    "column":"C"
  },
  "lookupServiceIds": {
    "sheetName":"Lookups",
    "column":"D"
  },
  "lookupTripPurposes": {
    "sheetName":"Lookups",
    "column":"E"
  },
  "lookupTripResults": {
    "sheetName":"Lookups",
    "column":"F"
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
}