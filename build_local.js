/**
 * @fileoverview Local menu extensions for RideSheet.
 *
 * This file is the intended customization point for org-specific forks.
 * Implement `buildLocalMenus()` to add org-specific items to the RideSheet
 * menu. It is called at the end of `buildMenus()` in `build.js`, after the
 * core menu has been constructed, so `SpreadsheetApp.getUi()` and the
 * existing menu are already set up.
 *
 * This instance adds an "NKADD Coordination" menu for the regional trip
 * coordination demo (see constants_local.js for the workflow overview).
 */

/**
 * Adds the NKADD Coordination menu:
 * - **Partner agency trip board** — the modal dialog partner agencies use to
 *   read, accept, decline, and copy trips (`nkadd_partner_ui.js`).
 * - **Offer selected trips to an agency** — NKADD's side of the workflow;
 *   sets Agency and moves Status to "Offered" (`nkadd_workflow.js`).
 * - **Set up coordination sheets** — creates the Status/Agency columns, the
 *   Claimed Trips sheet, and the lookups they depend on (`nkadd_setup.js`).
 * - **Load ConOps 5.1 demo trips** — the split round trip from the ConOps,
 *   offered to OCPT and Proudworks (`nkadd_demo_data.js`).
 */
function buildLocalMenus() {
  const ui = SpreadsheetApp.getUi()
  ui.createMenu("NKADD Coordination")
    .addItem("Partner agency trip board…", "showPartnerTripBoard")
    .addItem("Offer selected trips to an agency…", "nkaddOfferSelectedTrips")
    .addSeparator()
    .addItem("Set up coordination sheets", "setUpNkaddDemo")
    .addItem("Load ConOps 5.1 demo trips…", "nkaddLoadDemoTrips")
    .addToUi()
}
