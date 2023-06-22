# RideSheet

RideSheet is an Open Source ride scheduling application for small, demand responsive services. It is built on Google Sheets, and uses Apps Script to provide additional functionality for autofilling addresses as well as estimated distance and mileage, moving and verifying data as part of a review-based workflow, and managing information on customers, trips, vehicles, drivers, and more. It also uses the Transactional Data Specification (TDS) to allow sharing trip requests between agencies.

## Getting Started

RideSheet requires a copy of the spreadsheets in Google Sheets set up with the correct sheet and header names. The easiest way to do this is to copy an existing RideSheet. [Use the public sample to get started](https://docs.google.com/spreadsheets/d/1U_rmR08qW63hEK_5IWblzVXK4ZqQElaD1ymAQNGpNiU/edit#gid=1387872535). Simply make a copy of the sheet in your own local directory, and then make sure you have the latest version of the script installed.

*Note that you may need to reload the sheet (simply refresh the browser) in order to see the RideSheet options in your new copy. You will know that the RideSheet scripts are loaded if you see "API" and "RideSheet" in the main options menu.* 
<img width="602" alt="Screen Shot 2023-06-08 at 12 59 25 PM" src="https://github.com/full-path/ridesheet/assets/9342771/8fd65c9f-fd71-4794-a50d-c05ecb8bbb48">

To get the latest version of the code, you can use [Clasp](https://developers.google.com/apps-script/guides/clasp) to manage Apps Script. You can find the scriptID by opening Extensions > Apps Script and then selecting Project Settings in the lefthand menu.

```
git clone git@github.com:full-path/ridesheet.git
clasp setting scriptId YOUR_SCRIPT_ID
clasp push
```

Once you have the latest script installed, you will want to run a few basic upkeep actions to make sure everything is running smoothly.

### Ensure Permissions are Enabled

To give Apps Script permission to run, select any option for either the API or RideSheet menu. Google will open a pop-up asking you to authorise the app. If you have made a copy of an existing sheet, it's possible that Google will decide the code is potentially unsafe. You will have to click a small link at the bottom of the pop-up that says *Advanced Options (unsafe)* and then give permission to continue.

### Update Document Properties

In the main menu, under RideSheet > Settings, run `Application Properties`. This will update the `Document Properties` spreadsheet with any changes from the latest code.

Open the `Document Properties` spreadsheet and fill in the values in the second column. *TO-DO: Add description of each of the properties*

### Sheet Hygiene

To ensure sheets are cleaned up and incorporate the latest changes, select RideSheet > Settings > Build Metadata. After that runs, select RideSheet > Settings > Repair Sheets.

### Additional tasks

- Add actual data to Customers, Vehicles, Drivers and Services
- In the Apps Script editor, under Triggers, set up any cron jobs
- Create a test trip and run through review and archive to make sure everything is working
- Set up reporting (*TO-DO: Create Documentation*)

## Common Issues & Debugging

- In the `Document Properties` sheet, set `logLevel` to `verbose`  
  *This ensures that the most detailed possible errors show up in the `Debug Log` sheet*
  
- Ensure all checkbox columns are set so that an empty checkbox is null, rather than `False`
  *Under Data > Data Validation in the main menu, you can see all rules for the current sheet. For any checkbox ranges, the option `use custom cell values` should be selected, and unchecked should be empty*

  - Look for any red error markers on cells; these may indicate a missing or incorrect validation rule 

