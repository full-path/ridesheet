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
 *
 * ---
 *
 * **NKADD regional coordination demo.**
 *
 * This instance is owned by the NKADD ADRC and shared with partner agencies
 * (Proudworks and others). It supports the "Split Round Trip (Model A)"
 * scenario in section 5.1 of the Regional Mobility Initiative ConOps: NKADD
 * logs a trip request, offers the leg to a partner agency, the agency claims
 * or declines it, and the agency reports results back when the leg is run.
 *
 * The workflow adds two columns to the Trips sheet — `"Status"` and
 * `"Agency"` — plus a `"Claimed Trips"` sheet that holds trips between the
 * moment an agency claims them and the moment results are reported.
 *
 * A split round trip is two rows, not one: RideSheet's existing "Add return
 * trip" action creates the return leg, and each leg then carries its own
 * Status and Agency. That is what makes Model A work — in the 5.1 scenario the
 * outbound leg goes to OCPT and the return leg to Proudworks, and each agency
 * only ever sees and claims its own leg.
 *
 * Trip lifecycle across the three sheets:
 *
 * ```
 *   Trips              Claimed Trips           Trip Archive
 *   -----              -------------           ------------
 *   Draft
 *   Needs Agency  ──►  Claimed            ──►  Closed
 *   Offered            Scheduled
 *   Declined           Rider Notified
 *   Unfilled           Returned  ──┐
 *        ▲                         │
 *        └─────────────────────────┘
 * ```
 *
 * See `nkadd_workflow.js` for the sheet-to-sheet moves and
 * `nkadd_partner_ui.js` / `nkadd_trip_board.html` for the partner-agency
 * trip board dialog.
 */

// ---------------------------------------------------------------------------
// NKADD demo configuration
//
// These constants live in this file (rather than a separate one) because
// `localColumns` and `localNamedRanges` below reference them at load time, and
// Apps Script does not guarantee the evaluation order of project files.
// ---------------------------------------------------------------------------

/** @type {string} Sheet that holds trips an agency has claimed but not yet reported results for. */
const nkaddClaimedSheetName = "Claimed Trips"

/**
 * Every value the `"Status"` column can hold, in workflow order.
 *
 * Status tracks the **coordination state** of a trip — who owes the next
 * action. It is deliberately distinct from `"Trip Result"`, which records the
 * **service outcome** (Completed / No-Show / Cancelled) once the leg has run.
 *
 * Values in the pre-claim group live on the Trips sheet, values in the
 * post-claim group live on the Claimed Trips sheet, and `"Closed"` moves the
 * trip to Trip Archive.
 *
 * @type {string[]}
 */
const nkaddTripStatuses = [
  // --- On the Trips sheet: the trip still needs a way to be fulfilled ---
  "Draft",          // Logged by the ADRC call center, not yet vetted or ready to offer.
  "Needs Agency",   // Vetted and ready. In the pool; no agency assigned yet.
  "Offered",        // Offered to the agency named in "Agency". Awaiting their response.
  "Declined",       // The offered agency passed. Back in the pool; see "Declined By".
  "Unfilled",       // Booking window closed with no claim. Falls back to single-provider service.

  // --- On the Claimed Trips sheet: an agency owns the trip ---
  "Claimed",        // Agency accepted and is responsible for scheduling it.
  "Scheduled",      // Agency has placed it on a run in their own system.
  "Rider Notified", // NKADD has confirmed pickup time, location, and provider with the rider.
  "Returned",       // Agency is giving the trip back. Moves to Trips as "Needs Agency".

  // --- Terminal: moves to Trip Archive ---
  "Closed",         // Results reported. Moves to Trip Archive for reporting.
]

/** @type {string} Status that sends a trip from Trips to Claimed Trips. */
const nkaddStatusClaimed = "Claimed"

/** @type {string} Status that sends a trip from Claimed Trips back to Trips. */
const nkaddStatusReturned = "Returned"

/** @type {string} Status that sends a trip from Claimed Trips to Trip Archive. */
const nkaddStatusClosed = "Closed"

/** @type {string} Status a returned or declined trip lands on when it re-enters the pool. */
const nkaddStatusNeedsAgency = "Needs Agency"

/** @type {string} Status set when a trip is offered to a specific agency. */
const nkaddStatusOffered = "Offered"

/** @type {string} Status set when the offered agency declines. */
const nkaddStatusDeclined = "Declined"

/**
 * Agencies that trips can be offered to, drawn from the ConOps participating
 * provider list. Written to the Lookups sheet by `setUpNkaddDemo()`; edit the
 * Lookups sheet directly to change the live list without redeploying code.
 * @type {string[]}
 */
const nkaddAgencies = [
  "Proudworks",
  "OCPT",
  "TANK",
  "Scheben Care Center",
  "Meals on Wheels",
]

const localNamedRanges = {
  "lookupTripStatuses": {
    "sheetName": "Lookups",
    "headerName": "Trip Statuses"
  },
  "lookupAgencies": {
    "sheetName": "Lookups",
    "headerName": "Agencies"
  },
  // Cell triggers must be prefixed "localCode" so `callLocalCellTriggers()`
  // picks them up; the numeric suffix is stripped to find the handler key.
  "localCodeNkaddStatus1": {
    "sheetName": "Trips",
    "headerName": "Status"
  },
  "localCodeNkaddStatus2": {
    "sheetName": nkaddClaimedSheetName,
    "headerName": "Status"
  },
  "localCodeNkaddTripResult1": {
    "sheetName": nkaddClaimedSheetName,
    "headerName": "Trip Result"
  },
}
const localNamedRangesToRemove = []

const localSheetsToRemove = []
const localSheets = [nkaddClaimedSheetName]
const localSheetsWithHeaders = [nkaddClaimedSheetName]

const localColumnsToRemove = {}

/**
 * Column definitions for the status/agency columns and the Claimed Trips sheet.
 *
 * `getConfiguredColumns()` merges these over `defaultColumns`, so the entries
 * for "Trips", "Trip Archive", and "Lookups" *add* columns rather than
 * replacing the sheet's existing definitions.
 */
const localColumns = {
  "Trips": {
    "Status": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripStatuses",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip status.",
      },
    },
    "Agency": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupAgencies",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a participating agency.",
      },
    },
    // Running list of agencies that have passed on this trip, so the same
    // agency is not offered it twice and the history is visible to NKADD.
    "Declined By": {},
  },
  [nkaddClaimedSheetName]: {
    "Trip Date": {
      numberFormat: "M/d/yyyy",
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid date.",
      },
    },
    "Status": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripStatuses",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip status.",
      },
    },
    "Agency": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupAgencies",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a participating agency.",
      },
    },
    "Customer Name and ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupCustomerNames",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid customer name and ID.",
      },
    },
    "Trip Result": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripResults",
        showDropdown: true,
        allowInvalid: false,
        helpText: "Value must be a valid trip result. Entering one closes the trip and archives it.",
      },
    },
    "Earliest PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Latest PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "DO Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Appt Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual PU Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "Actual DO Time": {
      numberFormat: 'h":"mm am/pm',
      dataValidation: {
        criteriaType: "DATE_IS_VALID_DATE",
        helpText: "Value must be a valid time.",
      },
    },
    "PU Address": {},
    "DO Address": {},
    "Service ID": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupServiceIds",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid service ID.",
      },
    },
    "Guests": {
      numberFormat: "0",
      dataValidation: {
        criteriaType: "NUMBER_GREATER_THAN_OR_EQUAL_TO",
        args: [0],
        helpText: "Value must be a number of 0 or greater.",
      },
    },
    "Mobility Factors": {},
    "Trip Purpose": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripPurposes",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid trip purpose.",
      },
    },
    "Notes": {},
    // Free-text field the claiming agency can use for their own run/dispatch notes.
    "Agency Notes": {},
    // Carried across so decline history survives a claim-then-return round trip.
    "Declined By": {},
    "Est Hours": {
      numberFormat: "0.00",
    },
    "Est Miles": {},
    "Trip ID": {},
    "Customer ID": {},
    "Claim TS": {
      numberFormat: "M/d/yyyy h:mm am/pm",
    },
  },
  "Trip Archive": {
    "Status": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupTripStatuses",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a valid trip status.",
      },
    },
    "Agency": {
      dataValidation: {
        criteriaType: "VALUE_IN_RANGE",
        namedRange: "lookupAgencies",
        showDropdown: true,
        allowInvalid: true,
        helpText: "Value must be a participating agency.",
      },
    },
  },
  "Lookups": {
    "Trip Statuses": {},
    "Agencies": {},
  },
}
