# RideSheet

[![RideSheet Documentation](https://img.shields.io/badge/View-Documentation-blue)](https://docs.ridesheet.org/)

RideSheet is an Open Source ride scheduling application designed for small, demand responsive transportation services. Built on Google Sheets with Apps Script extensions, it provides essential functionality for transportation providers:

- Customer, trip, vehicle and driver data management
- Reporting integration
- Address autofill and geocoding
- Distance and travel time estimation
- Multi-step workflow with data validation

## Key Benefits

- Free and open source
- No complex software to install - runs in Google Sheets
- Easy to customize for your needs
- Secure - your data stays in your Google account
- Works on any device with a web browser

## Documentation

[Read the complete guide](https://docs.ridesheet.org/) for detailed installation instructions, usage tutorials, and customization options.

## Version Policy

New version releases will be tagged on the `Main` branch. 

* MAJOR (1.x.x → 2.0.0, etc.) – Breaking changes, to be accompanied by a “Migration Guide”.
* MINOR (1.0.x → 1.1.0, etc.) – New features, no breaking changes.
* PATCH (1.0.0 → 1.0.1, etc.) – Bug fixes, no breaking changes.

Forks and custom versions may use:

* 1.2.0-custom-feature – Fork with extra functionality.
* 1.3.0-orgname – Organization-specific version.

Check the CHANGELOG for updates and the Migration Guide for major changes.

## Forking & Extending

RideSheet is designed to be customized for specific organizations without modifying the core files. Three `_local` files are the intended extension points for forks:

### `constants_local.js`
Controls which sheets, columns, and named ranges are active. Use this to:
- Add or override named ranges (`localNamedRanges`)
- Remove default named ranges (`localNamedRangesToRemove`)
- Add extra sheets (`localSheets`, `localSheetsWithHeaders`)
- Remove default sheets (`localSheetsToRemove`)
- Add or override column definitions (`localColumns`)
- Remove default columns (`localColumnsToRemove`)

### `on_edit_local.js`
Adds org-specific onEdit behavior. Use this to:
- Add sheet-level triggers that fire before or after core cell triggers (`initialLocalSheetTriggers`, `finalLocalSheetTriggers`)
- Add cell-level triggers keyed to named ranges prefixed with `"localCode"` (`rangeTriggersLocal`)

### `build_local.js`
Adds org-specific menu items. Implement `buildLocalMenus()` to add entries to the UI — it is called at the end of `buildMenus()` after the core RideSheet menu is built.

### Recommended approach
Keep all org-specific logic in the `_local` files and avoid editing the core files (`constants.js`, `on_edit.js`, `build.js`, etc.). This makes it easier to pull in upstream updates and bugfixes without merge conflicts.