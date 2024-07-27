const localNamedRanges = {
    "codeCheckSourceOnShare": {
        "sheetName":"Trips",
        "headerName": "Share"
    },
    "codeVerifySourceOnEdit": {
        "sheetName":"Trips",
        "headerName":"Source"
    },
}
const localNamedRangesToRemove = []

const localSheetsToRemove = []
const localSheets = [
    "Sent Trips",
    "Outside Trips",
    "Outside Runs"
]
const localSheetsWithHeaders = [
    "Sent Trips",
    "Outside Trips"
]

const localColumnsToRemove = {}
const localColumns = {
    "Trips": {
        "Share": {
            dataValidation: {
            criteriaType: "CHECKBOX",
            checkedValue: "TRUE",
            allowInvalid: true,
            }
        },
        "Source": {},
        "Shared": {},
        "Declined By": {}
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
            numberFormat: "0.00"
        },
        "Est Miles": {},
        "Trip ID": {},
        "Customer ID": {}
    },
    "Outside Trips": {
        "Decline": {dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: true,
        }},
        "Claim": {dataValidation: {
        criteriaType: "CHECKBOX",
        checkedValue: "TRUE",
        allowInvalid: true,
        }},
        "Trip Date": {
        numberFormat: "M/d/yyyy",
        dataValidation: {
            criteriaType: "DATE_IS_VALID_DATE",
            helpText: "Value must be a valid date.",
        },
        },
        "Earliest PU Time": {
        numberFormat: 'h":"mm am/pm',
        dataValidation: {
            criteriaType: "DATE_IS_VALID_DATE",
            helpText: "Value must be a valid time.",
        },
        },
        "Requested PU Time": {
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
        "Requested DO Time": {
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
        numberFormat: "0.00"
        },
        "Est Miles": {},
        "Trip ID": {},
        "Customer Info": {},
        "Extra Fields": {},
        "Pending": {},
    },
    "Trip Review": {
        "Source": {},
        "Shared": {},
        "Declined By": {}
    },
    "Trip Archive": {
        "Source": {}
    }
}
