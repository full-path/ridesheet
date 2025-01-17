# Customization

Because RideSheet is a spreadsheet-based application, it is highly customizable to a wide variety of agency needs. Some customizations are possible without any specialized technical knowledge, although familiarity with spreadsheets is advisable.

## Customizing RideSheet's Appearance

You should be able to customize RideSheet's appearance without affecting its functionality. Note that in certain places, the existing coloring is meaningful (fields which should not be edited are blue, `|Run OK?|` in trips uses conditional formatting to indicate whether a run is valid).

However, editing font-size, row colors, or adjusting column widths should have no negative effects on the application.

Users who are comfortable with conditional formatting in spreadsheets can use it if desired. 

!!! tip "Date Formats and Other Field Settings"
    Some fields in RideSheet, such as date fields, have their number format set in the code. Operations like "repair sheets" will reset date fields to "m/d/yyyy" format. If you customize the format of these fields, be aware that running certain RideSheet operations may revert them back to their default settings.


## Customizing Driver Manifests

Driver manifests are generated using a template, built from a Google Doc. That template should be availabe in Google Drive in your `Settings` folder. The Manifest Template includes editing instructions. To be safe, make a copy of your current template before making any edits. 

## Adding Custom Fields

Your agency may want to collect data that is not present in RideSheet, such as fares or customer birthdays.

!!! warning "Test new columns carefully"
    If you add a new column to RideSheet, ensure it doesn't have any undesirable downstream effects. In particular, check that reports are still working correctly, and that you are still able to move trips and runs to review. 

If you add a new column to `Trips`, you will need to add the same column to `Trip Review` and `Trip Archive`. Similarly, if you add a new column to `Runs`, you will need to add it to `Run Review` and `Run Archive`.

You will have to add validation and formatting rules yourself, and RideSheet will not manage these fields apart from moving them to review and archive. 

If you would like to add a new field and have it managed by RideSheet and incorporated into other data, such as `Reports`, you should speak with the technician who installed RideSheet for your agency.