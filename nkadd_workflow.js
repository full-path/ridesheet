/**
 * @fileoverview Trip coordination workflow for the NKADD regional demo.
 *
 * Moves a trip between the three sheets that make up the coordination
 * lifecycle, driven by the `"Status"` column defined in constants_local.js:
 *
 * - **Trips** — trips that still need a way to be fulfilled. NKADD offers them
 *   to an agency here.
 * - **Claimed Trips** — an agency has accepted the trip and owns scheduling it.
 *   The trip waits here until the agency reports results.
 * - **Trip Archive** — results reported; the trip is done and available for
 *   reporting.
 *
 * Status transitions that move a row:
 *
 * | Sheet         | Status set to | Effect                                                      |
 * |---------------|---------------|-------------------------------------------------------------|
 * | Trips         | `Claimed`     | Moves to Claimed Trips, stamps `Claim TS`. Requires `Agency`. |
 * | Trips         | `Closed`      | Moves straight to Trip Archive. Requires `Trip Result`.       |
 * | Claimed Trips | `Returned`    | Moves back to Trips as `Needs Agency`; agency added to `Declined By`. |
 * | Claimed Trips | `Closed`      | Moves to Trip Archive. Requires `Trip Result`.                |
 *
 * Entering a `"Trip Result"` on Claimed Trips sets `Status` to `Closed` and
 * archives the trip in one step, so an agency reporting results never has to
 * touch the Status column.
 *
 * Every transition is also reachable programmatically — `nkaddClaimTrip()`,
 * `nkaddDeclineTrip()`, `nkaddReturnTrip()`, and `nkaddReportResults()` are
 * what the partner-agency dialog in `nkadd_partner_ui.js` calls.
 */

// ---------------------------------------------------------------------------
// onEdit triggers (wired up in on_edit_local.js)
// ---------------------------------------------------------------------------

/**
 * Handles a user edit to a `"Status"` cell on either the Trips or the
 * Claimed Trips sheet, performing the sheet-to-sheet move the new status
 * calls for. Statuses that do not move a row are left alone.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The edited Status cell.
 */
function nkaddStatusEditTrigger(range) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const sheet = range.getSheet()
    const sheetName = sheet.getName()
    const tripRow = getFullRow(range)
    const trip = getRangeValuesAsTable(tripRow, {includeFormulaValues: false})[0]
    const status = range.getValue()
    // A multi-column paste can fire this after a sibling trigger already moved
    // the row, leaving us pointed at a blank or shifted row.
    if (!isCompleteTrip(trip)) return

    if (sheetName === "Trips") {
      if (status === nkaddStatusClaimed) {
        if (!trip["Agency"]) {
          range.setValue(nkaddStatusOffered)
          ss.toast(`Set an Agency before marking a trip "${nkaddStatusClaimed}".`, "Agency required", 8)
          return
        }
        nkaddMoveToClaimed(tripRow, trip["Agency"])
        ss.toast(`Trip moved to "${nkaddClaimedSheetName}" for ${trip["Agency"]}.`, "Trip claimed", 5)
      } else if (status === nkaddStatusClosed) {
        if (!trip["Trip Result"]) {
          range.setValue(trip["Agency"] ? nkaddStatusOffered : nkaddStatusNeedsAgency)
          ss.toast(`Enter a Trip Result before marking a trip "${nkaddStatusClosed}".`, "Trip Result required", 8)
          return
        }
        nkaddMoveToArchive(tripRow)
        ss.toast("Trip moved to Trip Archive.", "Trip closed", 5)
      }
    } else if (sheetName === nkaddClaimedSheetName) {
      if (status === nkaddStatusReturned) {
        nkaddMoveToTrips(tripRow, trip["Agency"])
        ss.toast(`Trip returned to the Trips sheet as "${nkaddStatusNeedsAgency}".`, "Trip returned", 5)
      } else if (status === nkaddStatusClosed) {
        if (!trip["Trip Result"]) {
          range.setValue(nkaddStatusClaimed)
          ss.toast(`Enter a Trip Result before marking a trip "${nkaddStatusClosed}".`, "Trip Result required", 8)
          return
        }
        nkaddMoveToArchive(tripRow)
        ss.toast("Trip moved to Trip Archive.", "Trip closed", 5)
      }
    }
  } catch(e) { logError(e) }
}

/**
 * Handles a user edit to a `"Trip Result"` cell on the Claimed Trips sheet.
 * Reporting a result is what closes a trip, so this sets `Status` to
 * `"Closed"` and archives the row. Clearing the cell does nothing.
 * @param {GoogleAppsScript.Spreadsheet.Range} range - The edited Trip Result cell.
 */
function nkaddTripResultEditTrigger(range) {
  try {
    if (range.getSheet().getName() !== nkaddClaimedSheetName) return
    if (isBlankCell(range.getValue())) return
    const tripRow = getFullRow(range)
    // See the note in nkaddStatusEditTrigger about multi-column pastes.
    if (!isCompleteTrip(getRangeValuesAsTable(tripRow, {includeFormulaValues: false})[0])) return
    nkaddMoveToArchive(tripRow)
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Results recorded. Trip moved to Trip Archive.", "Trip closed", 5
    )
  } catch(e) { logError(e) }
}

// ---------------------------------------------------------------------------
// Sheet-to-sheet moves
// ---------------------------------------------------------------------------

/**
 * Moves a trip row from Trips to the Claimed Trips sheet, setting `Status` to
 * `"Claimed"`, recording the claiming agency, and stamping `"Claim TS"`.
 * @param {GoogleAppsScript.Spreadsheet.Range} tripRow - Full-row range on Trips.
 * @param {string} agency - The agency claiming the trip.
 * @param {string} [agencyNotes] - Optional note from the claiming agency.
 */
function nkaddMoveToClaimed(tripRow, agency, agencyNotes) {
  const destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nkaddClaimedSheetName)
  let extraFields = {
    "Status": nkaddStatusClaimed,
    "Agency": agency,
    "Claim TS": new Date(),
  }
  if (agencyNotes) extraFields["Agency Notes"] = agencyNotes
  nkaddMoveRow(tripRow, destSheet, extraFields)
}

/**
 * Moves a trip row from Claimed Trips back to the Trips sheet as
 * `"Needs Agency"`. The returning agency is appended to `"Declined By"` so it
 * is not offered the same trip again, and `Agency` / `Claim TS` are cleared.
 * @param {GoogleAppsScript.Spreadsheet.Range} tripRow - Full-row range on Claimed Trips.
 * @param {string} agency - The agency giving the trip back.
 */
function nkaddMoveToTrips(tripRow, agency) {
  const trip = getRangeValuesAsTable(tripRow, {includeFormulaValues: false})[0]
  const destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips")
  nkaddMoveRow(tripRow, destSheet, {
    "Status": nkaddStatusNeedsAgency,
    "Agency": null,
    "Claim TS": null,
    "Agency Notes": null,
    "Declined By": nkaddAddToDeclinedBy(trip["Declined By"], agency),
  })
}

/**
 * Moves a trip row to Trip Archive with `Status` set to `"Closed"` and an
 * `"Archive TS"` stamp. Works from either Trips or Claimed Trips; columns the
 * archive does not have (such as `"Claim TS"`) are dropped by `createRow()`.
 * @param {GoogleAppsScript.Spreadsheet.Range} tripRow - Full-row range to archive.
 */
function nkaddMoveToArchive(tripRow) {
  const destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trip Archive")
  nkaddMoveRow(tripRow, destSheet, {
    "Status": nkaddStatusClosed,
    "Archive TS": new Date(),
  })
}

/**
 * Thin wrapper around `moveRow()` that reapplies the destination sheet's
 * number formats and data validation to the newly appended row. `createRow()`
 * uses `appendRow()`, which does not carry formatting across.
 * @param {GoogleAppsScript.Spreadsheet.Range} sourceRow - Full-row range to move.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destSheet - The destination sheet.
 * @param {Object} extraFields - Field values to merge in before writing.
 */
function nkaddMoveRow(sourceRow, destSheet, extraFields) {
  moveRow(sourceRow, destSheet, {extraFields: extraFields})
  applySheetFormatsAndValidation(destSheet, destSheet.getLastRow())
}

// ---------------------------------------------------------------------------
// Programmatic entry points, used by the partner-agency dialog
// ---------------------------------------------------------------------------

/**
 * Claims a trip on behalf of an agency and moves it to Claimed Trips.
 * @param {string} tripId - The `"Trip ID"` of the trip to claim.
 * @param {string} agency - The claiming agency.
 * @param {string} [agencyNotes] - Optional note from the claiming agency.
 * @returns {{ok: boolean, message: string}} Result suitable for display in the dialog.
 */
function nkaddClaimTrip(tripId, agency, agencyNotes) {
  try {
    const tripsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips")
    const tripRow = nkaddFindRowByTripId(tripsSheet, tripId)
    if (!tripRow) {
      return {ok: false, message: "That trip is no longer on the Trips sheet. Another agency may have claimed it."}
    }
    nkaddMoveToClaimed(tripRow, agency, agencyNotes)
    log("NKADD trip claimed", `${tripId} claimed by ${agency}`)
    return {ok: true, message: `Claimed. The trip is now on the ${nkaddClaimedSheetName} sheet.`}
  } catch(e) {
    logError(e)
    return {ok: false, message: "Something went wrong claiming that trip. Check the Debug Log."}
  }
}

/**
 * Declines a trip on behalf of an agency. The trip stays on the Trips sheet
 * with `Status` set to `"Declined"`, `Agency` cleared, and the declining
 * agency appended to `"Declined By"` so NKADD can offer it elsewhere.
 * @param {string} tripId - The `"Trip ID"` of the trip being declined.
 * @param {string} agency - The declining agency.
 * @param {string} [reason] - Optional reason, appended to the trip's Notes.
 * @returns {{ok: boolean, message: string}} Result suitable for display in the dialog.
 */
function nkaddDeclineTrip(tripId, agency, reason) {
  try {
    const tripsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Trips")
    const tripRow = nkaddFindRowByTripId(tripsSheet, tripId)
    if (!tripRow) {
      return {ok: false, message: "That trip is no longer on the Trips sheet. Another agency may have claimed it."}
    }
    const trip = getRangeValuesAsTable(tripRow, {includeFormulaValues: false})[0]
    let valuesToChange = {
      "Status": nkaddStatusDeclined,
      "Agency": null,
      "Declined By": nkaddAddToDeclinedBy(trip["Declined By"], agency),
    }
    if (reason) {
      valuesToChange["Notes"] = [trip["Notes"], `${agency} declined: ${reason}`].filter(Boolean).join(" | ")
    }
    setValuesByHeaderNames([valuesToChange], tripRow)
    log("NKADD trip declined", `${tripId} declined by ${agency}`)
    return {ok: true, message: "Declined. NKADD can now offer this trip to another agency."}
  } catch(e) {
    logError(e)
    return {ok: false, message: "Something went wrong declining that trip. Check the Debug Log."}
  }
}

/**
 * Gives a claimed trip back to the pool, moving it from Claimed Trips to Trips.
 * @param {string} tripId - The `"Trip ID"` of the trip to return.
 * @param {string} agency - The agency giving the trip back.
 * @returns {{ok: boolean, message: string}} Result suitable for display in the dialog.
 */
function nkaddReturnTrip(tripId, agency) {
  try {
    const claimedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nkaddClaimedSheetName)
    const tripRow = nkaddFindRowByTripId(claimedSheet, tripId)
    if (!tripRow) {
      return {ok: false, message: `That trip is no longer on the ${nkaddClaimedSheetName} sheet.`}
    }
    nkaddMoveToTrips(tripRow, agency)
    log("NKADD trip returned", `${tripId} returned by ${agency}`)
    return {ok: true, message: "Returned. NKADD will offer this trip to another agency."}
  } catch(e) {
    logError(e)
    return {ok: false, message: "Something went wrong returning that trip. Check the Debug Log."}
  }
}

/**
 * Records the outcome of a claimed trip and archives it.
 * @param {string} tripId - The `"Trip ID"` of the trip being reported on.
 * @param {Object} results - Values to write before archiving, keyed by column
 *   name (e.g. `"Trip Result"`, `"Actual PU Time"`, `"Actual DO Time"`).
 * @returns {{ok: boolean, message: string}} Result suitable for display in the dialog.
 */
function nkaddReportResults(tripId, results) {
  try {
    if (!results || !results["Trip Result"]) {
      return {ok: false, message: "Choose a trip result first."}
    }
    const claimedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nkaddClaimedSheetName)
    const tripRow = nkaddFindRowByTripId(claimedSheet, tripId)
    if (!tripRow) {
      return {ok: false, message: `That trip is no longer on the ${nkaddClaimedSheetName} sheet.`}
    }
    setValuesByHeaderNames([results], tripRow)
    nkaddMoveToArchive(tripRow)
    log("NKADD trip results reported", `${tripId}: ${results["Trip Result"]}`)
    return {ok: true, message: "Results recorded. The trip moved to Trip Archive."}
  } catch(e) {
    logError(e)
    return {ok: false, message: "Something went wrong recording those results. Check the Debug Log."}
  }
}

/**
 * Offers every selected trip on the Trips sheet to an agency, setting
 * `"Agency"` and moving `"Status"` to `"Offered"`. This is NKADD's side of the
 * workflow (ConOps 5.1, step 3) and is exposed on the NKADD Coordination menu.
 */
function nkaddOfferSelectedTrips() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const ui = safeGetUi()
    const sheet = ss.getActiveSheet()
    if (sheet.getName() !== "Trips") {
      ss.toast("Select the trips to offer on the Trips sheet first.", "Nothing offered", 8)
      return
    }
    const agencies = nkaddGetAgencies()
    if (!agencies.length) {
      ss.toast("No agencies found. Add them to the Agencies column on the Lookups sheet.", "Nothing offered", 8)
      return
    }
    const response = ui.prompt(
      "Offer trips to an agency",
      `Which agency should these trips be offered to?\n\n${agencies.join("\n")}`,
      ui.ButtonSet.OK_CANCEL
    )
    if (response.getSelectedButton() !== ui.Button.OK) return
    const agency = agencies.find(name => name.toLowerCase() === response.getResponseText().trim().toLowerCase())
    if (!agency) {
      ui.alert(`"${response.getResponseText()}" is not a participating agency.`)
      return
    }

    const selectedRows = getFullRows(sheet.getActiveRange())
    const trips = getRangeValuesAsTable(selectedRows, {includeFormulaValues: false})
      .filter(trip => isCompleteTrip(trip))
    if (!trips.length) {
      ss.toast("No complete trips were selected.", "Nothing offered", 8)
      return
    }
    trips.forEach(trip => {
      setValuesForRow({"Status": nkaddStatusOffered, "Agency": agency}, trip._rowPosition, sheet)
    })
    ss.toast(`${pluralize(trips.length, "trip")} offered to ${agency}.`, "Trips offered", 5)
  } catch(e) { logError(e) }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Finds the full-row range for a trip on a sheet by its `"Trip ID"`.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet to search.
 * @param {string} tripId - The Trip ID to look for.
 * @returns {GoogleAppsScript.Spreadsheet.Range|null} The full-row range, or `null` if not found.
 */
function nkaddFindRowByTripId(sheet, tripId) {
  if (!sheet || !tripId) return null
  const trips = getRangeValuesAsTable(sheet.getDataRange(), {includeFormulaValues: false})
  const trip = trips.find(row => row["Trip ID"] === tripId)
  if (!trip) return null
  return getFullRow(sheet.getRange(trip._rowPosition, 1))
}

/**
 * Appends an agency to a comma-separated `"Declined By"` value, ignoring
 * agencies that are already listed.
 * @param {string} existing - The current `"Declined By"` cell value.
 * @param {string} agency - The agency to add.
 * @returns {string} The updated comma-separated list.
 */
function nkaddAddToDeclinedBy(existing, agency) {
  const agencies = (existing || "").toString().split(",").map(name => name.trim()).filter(Boolean)
  if (agency && !agencies.includes(agency)) agencies.push(agency)
  return agencies.join(", ")
}

/**
 * Returns the participating agency list from the Lookups sheet, falling back to
 * `nkaddAgencies` (constants_local.js) if the lookup column is missing or empty.
 * @returns {string[]}
 */
function nkaddGetAgencies() {
  return nkaddGetLookupValues("lookupAgencies", nkaddAgencies)
}

/**
 * Returns the trip status list from the Lookups sheet, falling back to
 * `nkaddTripStatuses` (constants_local.js).
 * @returns {string[]}
 */
function nkaddGetTripStatuses() {
  return nkaddGetLookupValues("lookupTripStatuses", nkaddTripStatuses)
}

/**
 * Returns the trip result list from the Lookups sheet, falling back to an
 * empty list so callers can decide what to show.
 * @returns {string[]}
 */
function nkaddGetTripResults() {
  return nkaddGetLookupValues("lookupTripResults", [])
}

/**
 * Reads the non-blank values of a single-column named range.
 * @param {string} namedRange - The named range to read.
 * @param {string[]} fallback - Values to return when the range is missing or empty.
 * @returns {string[]}
 */
function nkaddGetLookupValues(namedRange, fallback) {
  try {
    const range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(namedRange)
    if (!range) return fallback
    const values = getInBoundsRange(range).getValues().flat()
      .filter(value => !isBlankCell(value))
      .map(value => value.toString())
    return values.length ? values : fallback
  } catch(e) {
    logError(e)
    return fallback
  }
}
