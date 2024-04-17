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
  "lookupTripPurposes",
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
  }
}

const localSheetsToRemove = ["Trips", "Trip Review", "Outside Trips", "Customers", "Sent Trips", "Run Review", "Run Archive", "Trip Archive", "Runs", "Outside Runs", "Vehicles", "Drivers", "Services"]
const localSheets = ["Call Log", "TDS Referrals"]
const localSheetsWithHeaders = ["Call Log", "TDS Referrals"]

const localColumnsToRemove = {}
const localColumns = {
  "Call Log": {
    "Call Date": {},
    "Staff ID": {},
    "Customer First Name": {},
    "Customer Last Name": {},
    "Phone Number": {},
    "Email": {},
    "Home Address": {},
    "County": {},
    "City": {},
    "Over 60?": {},
    "Language": {},
    "Other Language": {},
    "Veteran?": {},
    "Trip Purpose": {},
    "Follow-up Needed": {},
    "Call Notes": {},
    "Non-TDS Referral 1": {},
    "Non-TDS Referral 2": {},
    "Referral/Other Notes": {},
    "Gaps": {},
    "Make TDS Referral To": {},
    "Go": {},
    "TDS Referrals": {},
    "Call ID": {}
  },
  "TDS Referrals": {
    "Call ID": {},
    "Customer ID": {},
    "Agency": {},
    "Referral Date": {},
    "Customer First Name": {},
    "Customer Nickname": {},
    "Customer Middle Name": {},
    "Customer Last Name": {},
    "Gender": {},
    "Home Phone Number": {},
    "Mobile Phone Number": {},
    "Email": {},
    "Home Address": {},
    "Date of Birth": {},
    "Low Income?": {},
    "Disability?": {},
    "Veteran?": {},
    "Language": {},
    "Race": {},
    "Ethnicity": {},
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
  }
}
