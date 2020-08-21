const allowPropDescriptionEdits          = false

const errorBackgroundColor               = "#f4cccc"
const defaultBackgroundColor             = "#ffffff"
const headerBackgroundColor              = "#fff2cc"

// Config for the state of Oregon
const defaultLocalTimeZone               = "America/Los_Angeles"
const defaultGeocoderBoundSwLatitude     = 41.997013
const defaultGeocoderBoundSwLongitude    = -124.560974
const defaultGeocoderBoundNeLatitude     = 46.299097
const defaultGeocoderBoundNeLongitude    = -116.463363

const defaultDwellTimeInMinutes          = 10
const defaultTripPaddingPerHourInMinutes = 5

const defaultProps = {
  lastCustomerID_: {
    value: 0,
    description: "The value of the last set customer ID."
  },
  driverManifestFolderId: {
    value: "Enter ID here",
    description: "The ID of the folder where newly created trip manifests will be saved."
  },
  driverManifestTemplateDocId: {
    value: "Enter ID here",
    description: "The document ID of the Google Doc you'll be using as your manifest template."
  },
  geocoderBoundNeLatitude: {
    value: 46.299097,
    description: "The north latitude of the box where Google Maps gives extra preference when geocoding addresses."
  }, 
  geocoderBoundNeLongitude: {
    value: -116.463363,
    description: "The east longitude of the box where Google Maps gives extra preference when geocoding addresses."
  },
  geocoderBoundSwLatitude: {
    value: 41.997013,
    description: "The south latitude of the box where Google Maps gives extra preference when geocoding addresses."
  },
  geocoderBoundSwLongitude: {
    value: -124.560974,
    description: "The west longitude of the box where Google Maps gives extra preference when geocoding addresses."
  },
  localTimeZone:  {
    value: "America/Los_Angeles",
    description: "The local time zone. Use one of the TZ database names found here: https://en.wikipedia.org/wiki/List_of_tz_database_time_zones"
  },
  monthlyBackupFolderId: {
    value: "Enter ID here",
    description: "The ID of the folder where monthly backups will be saved"
  },
  monthlyFileRetentionInDays: {
    value: 365,
    description: "How many days monthly backups should be held onto before they're automatically deleted"
  },
  nightlyBackupFolderId: {
    value: "Enter ID here",
    description: "The ID of the folder where nightly backups will be saved"
  },
  nightlyFileRetentionInDays: {
    value: 90,
    description: "How many days weekly backups should be held onto before they're automatically deleted"
  },
  weeklyBackupFolderId: {
    value: "Enter ID here",
    description: "The ID of the folder where nightly backups will be saved"
  },
  weeklyFileRetentionInDays: {
    value: 180,
    description: "How many days weekly backups should be held onto before they're automatically deleted"
  },
  dwellTimeInMinutes: {
    value: 10,
    description: "The length of time in minutes added to the journey time to account for the time it takes to pick up and drop off a rider"
  },
  tripPaddingPerHourInMinutes: {
    value: 5,
    description: "The length of time in minutes added to each hour of estimated travel time to account for weather, traffic, or other possible delays"
  }
}