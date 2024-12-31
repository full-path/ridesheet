# Importing Data

When migrating to RideSheet, you'll need to import your existing trip and customer data. Since every organization stores their data differently, there's no one-size-fits-all import process. However, here are some general tips for importing data:

## Tips for Importing Data

### From a Database
- Export your data to CSV files
- Import CSV data into another Google Sheets document first, then copy rows or columns over into your new RideSheet install

### From Another Spreadsheet
- Create a new sheet in your RideSheet workbook for temporary import
- Copy and paste your data into this temporary sheet
- Once formatted correctly, copy the transformed data into the appropriate RideSheet sheets

!!! warning "Large Imports"
    When copying many rows at once, be aware that RideSheet will attempt to geocode all addresses immediately. This can cause timeouts if importing too much data at once. Consider importing rows which contain address data in smaller batches.

### General Considerations
- Start by importing a small test set of data to verify your import process
- Validate that required fields are present and formatted correctly
- Check that dates and times are in the expected format
- Consider what amount of historical data is necessary to import, and consider handling it in batches.

If you need help planning your data migration, consider reaching out to [RideSheet support](../user-guide/getting-help.md) for guidance specific to your situation.
