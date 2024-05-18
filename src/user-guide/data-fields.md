# Data Fields

## Data Types
### Address Fields
RideSheet uses different types of address fields to manage locations effectively:

#### Street Addresses
Street addresses are used to specify physical locations. They should be entered in a standard format (e.g., "1600 Pennsylvania Avenue NW, Washington, DC 20500, USA").

#### Plus Codes
Plus Codes are short codes derived from latitude and longitude coordinates. They provide precise locations, especially in rural or less accessible areas. An example of a Plus Code format is "VXX7+3FV, Washington, DC, USA".

#### Address Descriptions
Address descriptions are optional and used to provide additional information about a location, such as specific instructions for drivers. They are enclosed in parentheses and do not affect geolocation.

#### Combining Address Elements
RideSheet supports combining address elements. For example:
- Street Address: "1600 Pennsylvania Avenue NW, Washington, DC 20500, USA"
- Plus Code: "VXX7+3FV"
- Combined: "VXX7+3FV; 1600 Pennsylvania Avenue NW, Washington, DC 20500, USA"

#### Common Addresses
Common addresses can be stored with short names for easy reference. When a short name is entered, RideSheet replaces it with the corresponding full address.

#### Validation
Upon entering a new address, RideSheet verifies its accuracy using Google Maps. If successful, it formats the address according to Google's standards.


### Dates and Times
[Blank]

## Fields by Sheet
### Trips
| Field Name                | Description                  |
|---------------------------|------------------------------|
| Trip Date                 | description                  |
| Customer Name and ID      | description                  |
| Action                    | description                  |
| Go                        | description                  |
| PU Time                   | description                  |
| DO Time                   | description                  |
| Appt Time                 | description                  |
| PU Address                | description                  |
| DO Address                | description                  |
| Driver ID                 | description                  |
| Vehicle ID                | description                  |
| Service ID                | description                  |
| Guests                    | description                  |
| Mobility Factors          | description                  |
| Notes                     | description                  |
| Est Hours                 | description                  |
| Est Miles                 | description                  |
| Trip ID                   | description                  |

### Customers
[Blank]

### Runs
[Blank]

### Trip Review
Trip Review has most of the same fields as [Trips](#trips).

The following fields are unique to Trip Review:

| Field Name                | Description                  |
|---------------------------|------------------------------|
| Trip Result                 | description                  |
| Actual PU Time      | description                  |
| Actual DO Time                   | description                  |
| Start ODO                      | description                  |
| End ODO                  | description                  |

### Trip Archive
Trip Archive has all the fields from [Trip Review](#trip-review). Trip Archive is used to keep track of all past trips, and is used in reporting. Trips should not be edited or removed once in the archive.

### Run Review
[Blank]

### Run Archive
[Blank]

### Addresses
[Blank]

### Vehicles
[Blank]

### Drivers
[Blank]

### Services
[Blank]

### Lookups
[Blank]

### Document Properties
[Blank]