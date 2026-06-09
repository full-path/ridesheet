/**
 * @fileoverview Local menu extensions for RideSheet.
 *
 * This file is the intended customization point for org-specific forks.
 * Implement `buildLocalMenus()` to add org-specific items to the RideSheet
 * menu. It is called at the end of `buildMenus()` in `build.js`, after the
 * core menu has been constructed, so `SpreadsheetApp.getUi()` and the
 * existing menu are already set up.
 *
 * Example:
 * ```js
 * function buildLocalMenus() {
 *   const ui = SpreadsheetApp.getUi()
 *   ui.createMenu('My Org').addItem('Custom action', 'myCustomFunction').addToUi()
 * }
 * ```
 */

function buildLocalMenus() {}
