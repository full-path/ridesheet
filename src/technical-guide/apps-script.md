# Apps Script

Google Apps Script is a JavaScript-based platform that allows you to automate tasks and add custom functionality to Google Workspace applications like Google Sheets. RideSheet is built on top of Apps Script, which handles calculations, validation, formatting, and automation.

## Accessing the Code

To view and edit the Apps Script code for your RideSheet installation:

1. Open your RideSheet spreadsheet
2. Click "Extensions" in the top menu
3. Select "Apps Script" from the dropdown
4. The Apps Script editor will open in a new tab

This will show you all the code files (ending in `.gs`) that power RideSheet's functionality.

## Learning Apps Script

If you're new to Apps Script, here are some helpful resources:

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Apps Script Fundamentals](https://developers.google.com/apps-script/guides/sheets)
- [Spreadsheet Service Reference](https://developers.google.com/apps-script/reference/spreadsheet)

## Running Functions

You can run individual functions directly from the Apps Script editor:

1. Open any `.gs` file
2. Select the function you want to run from the dropdown menu at the top
3. Click the "Run" button (play icon)
4. Grant any necessary permissions when prompted

This is useful for testing changes or running utility functions manually.

## Code Organization

RideSheet's code is organized into several `.gs` files, each handling different aspects of functionality:

### Core Files

`build.gs`
: Handles sheet formatting, repairs, metadata management, named ranges, and menu creation.

`constants.gs`
: Defines global constants used throughout the codebase, including sheet names, column names, and expected formats.

`maps.gs`
: Contains all functionality related to Google Maps integration, including address geocoding and distance calculations.

`properties.gs`
: Manages Document Properties, which store configuration settings for your RideSheet installation.

`on_edit.gs`
: Contains trigger functions that run automatically when users edit cells, handling validation and updates.

`on_open.gs`
: Defines functions that run when the spreadsheet is opened, setting up menus and performing some basic clean-up.

`reports.gs`
: Manages the generation of driver manifests and handles template application.

`review.gs`
: Contains logic for moving completed trips and runs to review/archive sheets.

`runs.gs`
: Handles calculations and updates for run-related fields.

`trips.gs`
: Manages trip-related functionality including return trip creation and validation.

### Utilities

`sheets.gs`
: Provides helper functions for common spreadsheet operations.

`util.gs`
: Contains general utility functions for logging, date handling, and other common tasks.

`setup.gs`
: Provides functionality for migrating data between RideSheet instances.

### Local Customization

RideSheet supports local customization through two optional files:

`build_local.gs`
: Add custom functionality and menu items

`constants_local.gs`
: Override default constants and add new ones

For information about setting up local development, using GitHub, or updating your code, see the [Updating the Code](updating-the-code.md) page.

!!! note "Developer Documentation"
    This page provides an overview of RideSheet's code organization. For detailed function documentation, please refer to the inline comments in the code files or visit the [GitHub repository](https://github.com/full-path/ridesheet).
