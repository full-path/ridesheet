# Troubleshooting

This guide covers common issues you may encounter while using RideSheet and how to resolve them.

## Debug Log

The `Debug Log` sheet records errors and events that occur in RideSheet. By default, it only logs errors, but you can get more detailed information by changing the value of `logLevel` in the `Document Properties` sheet from "normal" to "verbose".

This will cause RideSheet to log most events and actions, which can be helpful for troubleshooting.

## Viewing Developer Metadata 

RideSheet uses developer metadata to track important information about columns and sheets. If you're experiencing issues, you can view this metadata by:

1. Opening the Apps Script editor
2. Opening `build.gs`
3. Running the `showColumnMetadata` function

This will display all metadata in the Debug Log sheet for review.

## Getting Help

### Apps Script Resources

If you need help with the code that powers RideSheet, these resources may be helpful:

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Stack Overflow - google-apps-script tag](https://stackoverflow.com/questions/tagged/google-apps-script)

### Opening an Issue

If you've found a bug or have a feature request, please [open an issue](https://github.com/full-path/ridesheet/issues/new) on GitHub. When opening an issue:

1. Clearly describe what you were doing when the problem occurred
2. Include any relevant error messages from the Debug Log
3. Describe what you expected to happen vs what actually happened
4. Include screenshots if relevant

## Common Issues and Solutions

### Checkbox Columns

Make sure all checkbox columns are configured correctly:

- Empty checkboxes should be set to `null` rather than `False`
- This setting can be found in the column's data validation settings
- Incorrect settings can cause validation errors

### Validation Errors

Red markers in cells indicate validation issues:

- These can be easy to miss, so check cells carefully
- Hover over the red marker to see the specific validation error
- Fix the data to match the required format or values

### Sheet Problems

If you're experiencing general issues with sheets or data, try these steps:

1. From the RideSheet menu, select "Settings"
2. Run "Rebuild metadata" to recreate all column tracking information
3. Run "Repair sheets" to fix formatting and validation issues

If problems persist after trying these solutions, consider reaching out to [RideSheet support](../user-guide/getting-help.md) for additional help.
