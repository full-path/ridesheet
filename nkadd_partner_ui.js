/**
 * @fileoverview Server side of the partner-agency trip board dialog.
 *
 * The trip board is the interface a partner agency (Proudworks, in the ConOps
 * 5.1 demo) uses instead of reading the Trips sheet directly. It is opened with
 * `SpreadsheetApp.getUi().showModalDialog()` and rendered by
 * `nkadd_trip_board.html`. (The view and its server code cannot share a base
 * name: Apps Script file names must be unique across file types.) It gives an
 * agency dispatcher:
 *
 * 1. A readable card per trip — no horizontal scrolling — with Google Maps
 *    links for the pickup, the drop-off, and the pickup-to-drop-off route,
 *    each opening in a new browser tab.
 * 2. Accept and Decline buttons per trip, which move the trip between the
 *    Trips and Claimed Trips sheets via `nkadd_workflow.js`.
 * 3. A copy panel that renders claimed trips in a configurable format (TSV,
 *    CSV, labeled block, or a custom `{Field}` template) for pasting into the
 *    agency's own scheduling system.
 *
 * All formatting of the copy text happens in the browser from the `fields` map
 * on each trip, so changing format options re-renders instantly without a
 * server round trip.
 *
 * **Settings storage.** Dialog preferences live in *user* properties, not
 * document properties: `purgeOldDocumentProperties()` deletes any document
 * property that is not declared in `defaultDocumentProperties` (constants.js),
 * which would wipe them on the next open. Per-user storage also matches the
 * semantics — the selected agency and preferred paste format belong to the
 * person at the keyboard, not to the shared spreadsheet.
 */

/** @type {string} User-properties key holding the dialog's saved preferences. */
const nkaddUiSettingsKey = "nkaddPartnerUiSettings"

/**
 * Fields offered in the copy panel, in the order they are listed there.
 * Any other non-formula column found on the source sheet is appended after
 * these, so adding a column to the Trips sheet makes it exportable without a
 * code change.
 * @type {string[]}
 */
const nkaddPreferredExportFields = [
  "Trip Date",
  "PU Time",
  "Earliest PU Time",
  "Latest PU Time",
  "DO Time",
  "Appt Time",
  "Customer Name and ID",
  "PU Address",
  "DO Address",
  "Trip Purpose",
  "Mobility Factors",
  "Guests",
  "Notes",
  "Est Hours",
  "Est Miles",
  "Service ID",
  "Agency",
  "Status",
  "Customer ID",
  "Trip ID",
  "Maps Link",
]

/** @type {Object} Preferences applied the first time a user opens the dialog. */
const nkaddDefaultUiSettings = {
  agency: "Proudworks",
  format: "labeled",
  includeHeader: true,
  fields: [
    "Trip Date",
    "PU Time",
    "Appt Time",
    "Customer Name and ID",
    "PU Address",
    "DO Address",
    "Trip Purpose",
    "Mobility Factors",
    "Guests",
    "Notes",
  ],
  template: "{Trip Date} {PU Time} — {Customer Name and ID}\n  Pick up: {PU Address}\n  Drop off: {DO Address}\n  Notes: {Notes}",
}

/**
 * Opens the partner-agency trip board as a modal dialog.
 * Exposed on the "NKADD Coordination" menu (see `build_local.js`).
 */
function showPartnerTripBoard() {
  try {
    const html = HtmlService.createHtmlOutputFromFile("nkadd_trip_board")
      .setWidth(1000)
      .setHeight(680)
    SpreadsheetApp.getUi().showModalDialog(html, "Partner Agency Trip Board")
  } catch(e) { logError(e) }
}

/**
 * Collects everything the dialog needs to render, for one agency's point of view.
 *
 * Trips are sorted soonest-first and split into three lists:
 * - `offered` — on the Trips sheet, `Status` is `"Offered"`, `Agency` is this agency.
 * - `open` — on the Trips sheet, available to anyone: `Status` is `"Needs Agency"`
 *   or `"Declined"`, no agency assigned, and this agency has not already passed.
 * - `claimed` — on the Claimed Trips sheet and held by this agency.
 *
 * @param {string} [agency] - The agency to build the board for. Defaults to the
 *   agency saved in the viewer's settings.
 * @returns {Object} Board data. `trips` in each list are plain JSON-safe objects.
 */
function nkaddGetBoardData(agency) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const settings = nkaddGetUiSettings()
    const agencies = nkaddGetAgencies()
    const selectedAgency = agency || settings.agency || agencies[0]

    const tripsSheet = ss.getSheetByName("Trips")
    const claimedSheet = ss.getSheetByName(nkaddClaimedSheetName)
    const setupHint = 'Run "NKADD Coordination → Set up coordination sheets" first.'
    if (!tripsSheet || !claimedSheet) {
      return {error: `The "${nkaddClaimedSheetName}" sheet is missing. ${setupHint}`}
    }
    const tripsHeaders = getSheetHeaderNames(tripsSheet, {forceRefresh: true})
    const missingColumns = ["Status", "Agency"].filter(name => !tripsHeaders.includes(name))
    if (missingColumns.length) {
      return {
        error: `The Trips sheet is missing the ${missingColumns.join(" and ")} ` +
          `column${missingColumns.length === 1 ? "" : "s"}. ${setupHint}`
      }
    }

    const tripRows = getRangeValuesAsTable(tripsSheet.getDataRange(), {includeFormulaValues: false})
      .filter(row => isCompleteTrip(row))
    const claimedRows = getRangeValuesAsTable(claimedSheet.getDataRange(), {includeFormulaValues: false})
      .filter(row => isCompleteTrip(row))

    const offered = tripRows.filter(row =>
      row["Status"] === nkaddStatusOffered && row["Agency"] === selectedAgency
    )
    const open = tripRows.filter(row =>
      !row["Agency"] &&
      [nkaddStatusNeedsAgency, nkaddStatusDeclined].includes(row["Status"]) &&
      !nkaddHasDeclined(row["Declined By"], selectedAgency)
    )
    const claimed = claimedRows.filter(row =>
      row["Agency"] === selectedAgency && !row["Trip Result"]
    )

    const exportFields = nkaddGetExportFields([tripsSheet, claimedSheet])
    return {
      agency: selectedAgency,
      agencies: agencies,
      tripResults: nkaddGetTripResults(),
      exportFields: exportFields,
      settings: settings,
      offered: nkaddBuildTripCards(offered, exportFields),
      open: nkaddBuildTripCards(open, exportFields),
      claimed: nkaddBuildTripCards(claimed, exportFields),
    }
  } catch(e) {
    logError(e)
    return {error: "Could not load the trip board. Check the Debug Log for details."}
  }
}

/**
 * Converts trip row objects into the JSON-safe shape the dialog renders.
 * Dates and times become preformatted strings so the browser never has to
 * reason about the spreadsheet's time zone.
 * @param {Object[]} rows - Trip row objects from `getRangeValuesAsTable()`.
 * @param {string[]} exportFields - Field names the copy panel can offer.
 * @returns {Object[]} Trip cards, sorted soonest first.
 */
function nkaddBuildTripCards(rows, exportFields) {
  const timeZone = getDocProp("localTimeZone") || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()
  return rows
    .map(row => {
      const puAddress = (row["PU Address"] || "").toString()
      const doAddress = (row["DO Address"] || "").toString()
      let fields = {}
      exportFields.forEach(fieldName => {
        fields[fieldName] = nkaddFormatFieldValue(fieldName, row[fieldName], timeZone)
      })
      fields["Maps Link"] = nkaddDirectionsUrl(puAddress, doAddress)
      return {
        id: (row["Trip ID"] || "").toString(),
        sortKey: nkaddTripSortKey(row),
        date: nkaddFormatFieldValue("Trip Date", row["Trip Date"], timeZone),
        dayOfWeek: isValidDate(row["Trip Date"]) ? formatDate(row["Trip Date"], timeZone, "EEE") : "",
        customer: (row["Customer Name and ID"] || "").toString(),
        status: (row["Status"] || "").toString(),
        agency: (row["Agency"] || "").toString(),
        declinedBy: (row["Declined By"] || "").toString(),
        puTime: nkaddFormatFieldValue("PU Time", row["PU Time"], timeZone),
        doTime: nkaddFormatFieldValue("DO Time", row["DO Time"], timeZone),
        apptTime: nkaddFormatFieldValue("Appt Time", row["Appt Time"], timeZone),
        earliestPuTime: nkaddFormatFieldValue("Earliest PU Time", row["Earliest PU Time"], timeZone),
        latestPuTime: nkaddFormatFieldValue("Latest PU Time", row["Latest PU Time"], timeZone),
        puAddress: puAddress,
        doAddress: doAddress,
        puCity: nkaddSafeCity(puAddress),
        doCity: nkaddSafeCity(doAddress),
        puMapUrl: nkaddPlaceUrl(puAddress),
        doMapUrl: nkaddPlaceUrl(doAddress),
        directionsUrl: fields["Maps Link"],
        purpose: (row["Trip Purpose"] || "").toString(),
        mobility: (row["Mobility Factors"] || "").toString(),
        guests: row["Guests"] ? row["Guests"].toString() : "",
        serviceId: (row["Service ID"] || "").toString(),
        notes: (row["Notes"] || "").toString(),
        agencyNotes: (row["Agency Notes"] || "").toString(),
        estHours: row["Est Hours"] ? Number(row["Est Hours"]).toFixed(2) : "",
        estMiles: row["Est Miles"] ? Number(row["Est Miles"]).toFixed(1) : "",
        fields: fields,
      }
    })
    .sort((a, b) => a.sortKey - b.sortKey)
}

/**
 * Builds a sortable numeric key from a trip's date and pickup time, so trips
 * read chronologically. Trips with no date sort last.
 * @param {Object} row - A trip row object.
 * @returns {number} Milliseconds since the epoch, or `Number.MAX_SAFE_INTEGER`
 *   when undated. (A finite value, because `Infinity` is not valid JSON and
 *   would not survive the trip to the dialog.)
 */
function nkaddTripSortKey(row) {
  if (!isValidDate(row["Trip Date"])) return Number.MAX_SAFE_INTEGER
  const time = isValidDate(row["PU Time"]) ? timeOnlyAsMilliseconds(row["PU Time"]) : 0
  return dateOnly(row["Trip Date"]).getTime() + time
}

/**
 * Formats a single cell value for display and for copying, choosing a date,
 * time, or date-and-time pattern based on the column name.
 * @param {string} fieldName - The column header name.
 * @param {*} value - The raw cell value.
 * @param {string} timeZone - IANA time zone for date formatting.
 * @returns {string} A display-ready string; empty when the cell is blank.
 */
function nkaddFormatFieldValue(fieldName, value, timeZone) {
  if (value === undefined || isBlankCell(value)) return ""
  if (isValidDate(value)) {
    if (/\bdate\b/i.test(fieldName)) return formatDate(value, timeZone, "M/d/yyyy")
    if (/\btime\b/i.test(fieldName)) return formatDate(value, timeZone, "h:mm a")
    return formatDate(value, timeZone, "M/d/yyyy h:mm a")
  }
  return value.toString()
}

/**
 * Returns the ordered list of field names the copy panel can offer: the
 * preferred fields that actually exist on the given sheets, followed by any
 * other real (non-formula, non-blank) columns those sheets have.
 * @param {GoogleAppsScript.Spreadsheet.Sheet[]} sheets - Sheets to draw column names from.
 * @returns {string[]}
 */
function nkaddGetExportFields(sheets) {
  let available = new Set(["Maps Link"])
  sheets.forEach(sheet => {
    if (!sheet) return
    getSheetHeaderNames(sheet, {forceRefresh: true}).forEach(headerName => {
      const name = headerName.toString()
      if (name.trim() && name[0] !== "|" && name[0] !== "_") available.add(name)
    })
  })
  const preferred = nkaddPreferredExportFields.filter(name => available.has(name))
  const remaining = [...available].filter(name => !preferred.includes(name)).sort()
  return preferred.concat(remaining)
}

/**
 * Builds a Google Maps directions URL from a pickup address to a drop-off
 * address, opening in driving mode. Falls back to a plain place lookup when
 * only one of the two addresses is known.
 * @param {string} origin - The pickup address.
 * @param {string} destination - The drop-off address.
 * @returns {string} A maps URL, or an empty string when neither address is set.
 */
function nkaddDirectionsUrl(origin, destination) {
  const from = origin ? parseAddress(origin).geocodeAddress : ""
  const to = destination ? parseAddress(destination).geocodeAddress : ""
  if (!from && !to) return ""
  if (!from || !to) return nkaddPlaceUrl(from || to)
  return "https://www.google.com/maps/dir/?api=1&travelmode=driving" +
    "&origin=" + encodeURIComponent(from) +
    "&destination=" + encodeURIComponent(to)
}

/**
 * Builds a Google Maps URL that shows a single address.
 * @param {string} address - The address to show.
 * @returns {string} A maps URL, or an empty string when `address` is blank.
 */
function nkaddPlaceUrl(address) {
  if (!address) return ""
  const query = parseAddress(address).geocodeAddress
  if (!query) return ""
  return "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(query)
}

/**
 * Returns the city and state for an address, or an empty string when it cannot
 * be determined. Wraps `extractCity()` so a geocoding failure never stops the
 * whole board from rendering.
 * @param {string} address - The address to summarize.
 * @returns {string}
 */
function nkaddSafeCity(address) {
  if (!address) return ""
  try {
    const city = extractCity(address)
    return city === "Unspecified area" ? "" : city
  } catch(e) {
    return ""
  }
}

/**
 * Returns `true` when `agency` appears in a comma-separated `"Declined By"` value.
 * @param {string} declinedBy - The `"Declined By"` cell value.
 * @param {string} agency - The agency to look for.
 * @returns {boolean}
 */
function nkaddHasDeclined(declinedBy, agency) {
  return (declinedBy || "").toString()
    .split(",")
    .map(name => name.trim())
    .includes(agency)
}

// ---------------------------------------------------------------------------
// Actions invoked from the dialog via google.script.run
// ---------------------------------------------------------------------------

/**
 * Claims a trip for an agency, then returns fresh board data so the dialog can
 * re-render in a single round trip.
 * @param {string} tripId - The Trip ID to claim.
 * @param {string} agency - The claiming agency.
 * @param {string} [notes] - Optional note from the claiming agency.
 * @returns {Object} `{ok, message, board}`.
 */
function nkaddUiClaim(tripId, agency, notes) {
  const result = nkaddClaimTrip(tripId, agency, notes)
  return {...result, board: nkaddGetBoardData(agency)}
}

/**
 * Declines a trip on an agency's behalf and returns fresh board data.
 * @param {string} tripId - The Trip ID being declined.
 * @param {string} agency - The declining agency.
 * @param {string} [reason] - Optional reason, appended to the trip's Notes.
 * @returns {Object} `{ok, message, board}`.
 */
function nkaddUiDecline(tripId, agency, reason) {
  const result = nkaddDeclineTrip(tripId, agency, reason)
  return {...result, board: nkaddGetBoardData(agency)}
}

/**
 * Gives a claimed trip back to the pool and returns fresh board data.
 * @param {string} tripId - The Trip ID being returned.
 * @param {string} agency - The agency returning the trip.
 * @returns {Object} `{ok, message, board}`.
 */
function nkaddUiReturn(tripId, agency) {
  const result = nkaddReturnTrip(tripId, agency)
  return {...result, board: nkaddGetBoardData(agency)}
}

/**
 * Records a trip result for a claimed trip, archiving it, and returns fresh
 * board data.
 * @param {string} tripId - The Trip ID being reported on.
 * @param {string} agency - The reporting agency.
 * @param {string} tripResult - The trip result value.
 * @returns {Object} `{ok, message, board}`.
 */
function nkaddUiReportResults(tripId, agency, tripResult) {
  const result = nkaddReportResults(tripId, {"Trip Result": tripResult})
  return {...result, board: nkaddGetBoardData(agency)}
}

// ---------------------------------------------------------------------------
// Dialog preferences
// ---------------------------------------------------------------------------

/**
 * Reads the viewer's saved dialog preferences, merged over the defaults.
 * @returns {Object} The effective settings object.
 */
function nkaddGetUiSettings() {
  try {
    const stored = PropertiesService.getUserProperties().getProperty(nkaddUiSettingsKey)
    if (!stored) return {...nkaddDefaultUiSettings}
    return {...nkaddDefaultUiSettings, ...JSON.parse(stored)}
  } catch(e) {
    logError(e)
    return {...nkaddDefaultUiSettings}
  }
}

/**
 * Saves the viewer's dialog preferences (selected agency, copy format, chosen
 * fields, custom template).
 * @param {Object} settings - The settings object to persist.
 * @returns {boolean} `true` when the settings were saved.
 */
function nkaddSaveUiSettings(settings) {
  try {
    const merged = {...nkaddGetUiSettings(), ...settings}
    PropertiesService.getUserProperties().setProperty(nkaddUiSettingsKey, JSON.stringify(merged))
    return true
  } catch(e) {
    logError(e)
    return false
  }
}
