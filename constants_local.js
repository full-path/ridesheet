/**
 * @fileoverview Local overrides and extensions for RideSheet constants.
 *
 * This file is the intended customization point for org-specific forks.
 * Modify these variables to add, remove, or replace sheets, columns, and
 * named ranges without touching the core `constants.js` file.
 *
 * - `localNamedRanges` — named ranges to add or override (merged over `defaultNamedRanges`).
 * - `localNamedRangesToRemove` — named range keys to delete from the active config.
 * - `localSheets` — additional sheet names beyond `defaultSheets`.
 * - `localSheetsToRemove` — sheet names to exclude from the active config.
 * - `localSheetsWithHeaders` — additional sheets that have header rows.
 * - `localColumns` — column definitions to add or override per sheet.
 * - `localColumnsToRemove` — column names to exclude per sheet.
 *
 * The effective configuration is computed at runtime in `build.js` by merging
 * these locals with the defaults. See `getConfiguredColumns()`,
 * `getConfiguredSheets()`, and `buildNamedRanges()` for details.
 */

const localNamedRanges = {}
const localNamedRangesToRemove = []

const localSheetsToRemove = []
const localSheets = []
const localSheetsWithHeaders = []

const localColumnsToRemove = {}
const localColumns = {}
