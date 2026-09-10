/**
 * @fileoverview Sample data for walking through ConOps scenario 5.1.
 *
 * `nkaddLoadDemoTrips()` creates the split round trip described in section 5.1
 * of the Regional Mobility Initiative ConOps — "Betty," age 74, travelling from
 * Grant County to a cardiology follow-up at St. Elizabeth Edgewood and home
 * again — as **two trip rows**, which is the point of the exercise:
 *
 * - The **outbound leg** is offered to OCPT, whose vehicle is already heading
 *   north that morning (ConOps step 4).
 * - The **return leg** is offered to Proudworks, whose vehicle will be in the
 *   area that afternoon returning clients south (step 5).
 *
 * Each leg carries its own `Status` and `Agency`, so each agency sees and
 * claims only its own leg on the partner trip board. That is what makes Model
 * A work without any change to RideSheet's trip schema.
 *
 * Two notes on the addresses used here. The ConOps draft places Betty "near
 * Warsaw" but names Grant County three times, and Warsaw is in Gallatin
 * County — this uses a Grant County address, since the county is what drives
 * the coordination decision. It also has the vehicle heading "toward Boone
 * County" while naming a destination in Kenton County; the real St. Elizabeth
 * Edgewood address is used. Both are worth resolving in the ConOps itself.
 */

/** @type {Object} The ConOps 5.1 rider. Created on the Customers sheet if absent. */
const nkaddDemoCustomer = {
  firstName: "Betty",
  lastName: "Caldwell",
  dateOfBirth: new Date(1952, 2, 14),
  phone: "859-555-0142",
  homeAddress: "412 N Main St, Williamstown, KY 41097",
  notes: "Uses a walker. Does not require a wheelchair-accessible vehicle.",
}

/** @type {string} Where Betty's cardiology follow-up is. */
const nkaddDemoDestination = "1 Medical Village Dr, Edgewood, KY 41017"

/**
 * Creates the ConOps 5.1 split round trip on the Trips sheet, offered to OCPT
 * (outbound) and Proudworks (return).
 *
 * The trip date is placed three days out, matching the regional 72-hour
 * advance booking standard, and nudged to the following Monday if that lands
 * on a weekend. Re-running adds another pair of trips; the rider record is
 * reused rather than duplicated.
 *
 * Exposed on the "NKADD Coordination" menu (see `build_local.js`).
 */
function nkaddLoadDemoTrips() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const ui = safeGetUi()
    const tripsSheet = ss.getSheetByName("Trips")
    if (!tripsSheet) return

    const tripsHeaders = getSheetHeaderNames(tripsSheet, {forceRefresh: true})
    if (!tripsHeaders.includes("Status") || !tripsHeaders.includes("Agency")) {
      if (ui) {
        ui.alert(
          "Setup needed",
          'Run "NKADD Coordination → Set up coordination sheets" before loading demo trips.',
          ui.ButtonSet.OK
        )
      }
      return
    }

    const tripDate = nkaddDemoTripDate()
    if (ui) {
      const response = ui.alert(
        "Load ConOps 5.1 demo trips",
        `This adds two trips to the Trips sheet for ${formatDate(tripDate)}:\n\n` +
        `  • Outbound, Williamstown → St. Elizabeth Edgewood, offered to OCPT\n` +
        `  • Return, St. Elizabeth Edgewood → Williamstown, offered to Proudworks\n\n` +
        `It also adds ${nkaddDemoCustomer.firstName} ${nkaddDemoCustomer.lastName} ` +
        `to the Customers sheet if she is not already there.\n\nContinue?`,
        ui.ButtonSet.YES_NO
      )
      if (response !== ui.Button.YES) return
    }

    const customer = nkaddFindOrCreateDemoCustomer()
    if (!customer) {
      if (ui) ui.alert("Could not create the demo rider. Check the Debug Log.")
      return
    }

    // Let the Customers sheet's lookup formula pick up a newly added rider
    // before the Trips sheet validates her name against it.
    SpreadsheetApp.flush()

    const purpose = nkaddMatchLookupValue(nkaddGetLookupValues("lookupTripPurposes", []), "Medical")
    const shared = {
      "Trip Date": tripDate,
      "Customer Name and ID": customer["Customer Name and ID"],
      "Customer ID": customer["Customer ID"],
      "Mobility Factors": "Walker",
      "Guests": 0,
      "Status": nkaddStatusOffered,
    }
    if (purpose) shared["Trip Purpose"] = purpose

    const outbound = {
      ...shared,
      "Trip ID": Utilities.getUuid(),
      "PU Time": nkaddTimeValue(8, 15),
      "DO Time": nkaddTimeValue(9, 40),
      "Appt Time": nkaddTimeValue(10, 0),
      "PU Address": nkaddDemoCustomer.homeAddress,
      "DO Address": nkaddDemoDestination,
      "Agency": "OCPT",
      "Notes": "Outbound leg of a split round trip (ConOps 5.1). Curb-to-curb; rider uses a walker. " +
        "Coordination candidate: origin in a southern county, destination in a northern county.",
    }

    const returnLeg = {
      ...shared,
      "Trip ID": Utilities.getUuid(),
      "Earliest PU Time": nkaddTimeValue(13, 15),
      "PU Time": nkaddTimeValue(13, 30),
      "Latest PU Time": nkaddTimeValue(14, 0),
      "DO Time": nkaddTimeValue(14, 55),
      "PU Address": nkaddDemoDestination,
      "DO Address": nkaddDemoCustomer.homeAddress,
      "Agency": "Proudworks",
      "Notes": "Return leg of a split round trip (ConOps 5.1). Pickup point at the destination must be " +
        "pre-communicated to the rider. Appointment may run long; see the pickup window.",
    }

    nkaddPrepareRowsForWrite(tripsSheet, 2)
    if (!createRows(tripsSheet, [outbound, returnLeg])) return
    log("NKADD demo trips loaded", `Split round trip for ${formatDate(tripDate)}`)
    ss.toast(
      `Outbound offered to OCPT, return offered to Proudworks for ${formatDate(tripDate)}.`,
      "Demo trips added", 8
    )
  } catch(e) { logError(e) }
}

/**
 * Returns the demo rider's row from the Customers sheet, creating it first if
 * she is not there. The row is created without a Customer ID and then handed to
 * the core `setCustomerKeyOnEdit()` so the ID and the "Customer Name and ID"
 * key are generated exactly the way a hand-entered rider's would be.
 * @returns {Object|undefined} The customer row object, or `undefined` on failure.
 */
function nkaddFindOrCreateDemoCustomer() {
  try {
    const customersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers")
    if (!customersSheet) return
    const matches = function(row) {
      return row["Customer First Name"] === nkaddDemoCustomer.firstName &&
        row["Customer Last Name"] === nkaddDemoCustomer.lastName
    }
    const existing = findFirstRowByHeaderNames(customersSheet, matches)
    if (existing) return existing

    nkaddPrepareRowsForWrite(customersSheet, 1)
    const created = createRows(customersSheet, [{
      "Customer First Name": nkaddDemoCustomer.firstName,
      "Customer Last Name": nkaddDemoCustomer.lastName,
      "Date of Birth": nkaddDemoCustomer.dateOfBirth,
      "Phone Number": nkaddDemoCustomer.phone,
      "Home Address": nkaddDemoCustomer.homeAddress,
      "Customer Manifest Notes": nkaddDemoCustomer.notes,
    }])
    if (!created) return

    setCustomerKeyOnEdit(customersSheet.getRange(customersSheet.getLastRow(), 1))
    return findFirstRowByHeaderNames(customersSheet, matches)
  } catch(e) { logError(e) }
}

/**
 * Makes room for `count` rows at the end of a sheet and strips data validation
 * from them.
 *
 * `fixDataValidation()` applies each column's rule down the whole column, so
 * the blank rows below the data are already strict — and `setValues()` throws
 * rather than writing a value a strict rule rejects. The most likely casualty
 * is "Customer Name and ID", whose lookup list is a QUERY formula that may not
 * have recalculated yet. `createRows()` reapplies validation to the rows it
 * writes, so clearing first costs nothing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet about to be appended to.
 * @param {number} count - How many rows will be appended.
 */
function nkaddPrepareRowsForWrite(sheet, count) {
  const firstRow = sheet.getLastRow() + 1
  const shortfall = firstRow + count - 1 - sheet.getMaxRows()
  if (shortfall > 0) sheet.insertRowsAfter(sheet.getMaxRows(), shortfall)
  sheet.getRange(firstRow, 1, count, sheet.getMaxColumns()).clearDataValidations()
}

/**
 * Returns the demo trip date: three days out, per the regional 72-hour advance
 * booking standard, moved to Monday when that falls on a weekend.
 * @returns {Date}
 */
function nkaddDemoTripDate() {
  let tripDate = dateAdd(dateToday(), 3)
  const dayOfWeek = tripDate.getDay()
  if (dayOfWeek === 6) tripDate = dateAdd(tripDate, 2)
  if (dayOfWeek === 0) tripDate = dateAdd(tripDate, 1)
  return tripDate
}

/**
 * Builds a time-only value the way Google Sheets stores one — a Date on the
 * spreadsheet epoch (30 December 1899) — so it sorts and compares alongside
 * times a user typed in by hand.
 * @param {number} hours - Hours in 24-hour form.
 * @param {number} minutes - Minutes past the hour.
 * @returns {Date}
 */
function nkaddTimeValue(hours, minutes) {
  return new Date(1899, 11, 30, hours, minutes)
}

/**
 * Finds a value in a lookup list, ignoring case and surrounding whitespace.
 * Used so the demo only fills a validated dropdown column when the instance
 * actually offers a matching option.
 * @param {string[]} values - The lookup list to search.
 * @param {string} wanted - The value to look for.
 * @returns {string|undefined} The list's own spelling of the value, if present.
 */
function nkaddMatchLookupValue(values, wanted) {
  return values.find(value => value.toString().trim().toLowerCase() === wanted.toLowerCase())
}
