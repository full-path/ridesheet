/**
 * @fileoverview Driver manifest creation for RideSheet.
 *
 * Generates per-run Google Docs (and optionally PDFs) from a template document
 * stored in Google Drive. The pipeline is:
 *
 * 1. **Entry point** — `createManifestsByRunForDate()` (all past-due trips for a
 *    date) or `createSelectedManifestsByRun()` (selected rows only).
 * 2. **`getManifestData(filterFn)`** — reads the Trips, Drivers, Vehicles, and
 *    Customers sheets, applies the filter, joins lookup data via `mergeAttributes()`,
 *    then builds two event rows per trip (one `PICKUP`, one `DROP OFF`), sorted by
 *    time, and returns `{trips, events}`.
 * 3. **`groupManifestDataByRun(manifestData)`** — groups trips and events into
 *    per-run manifest groups keyed by date + driver + vehicle + run ID.
 * 4. **`createManifests(templateDocId, groupedData, fileNameFn)`** — iterates
 *    manifest groups, calling `createManifest()` for each, and optionally
 *    `createPdfFromDocFile()` and trashing the source Doc.
 * 5. **`createManifest(group, templateDoc, fileName, folderId)`** — creates a new
 *    Google Doc by copying template structure (page settings, header, footer,
 *    named ranges for `HEADER`, `PICKUP`, `DROP OFF`, `FOOTER` sections) and
 *    running template substitution via `replaceElementText()`.
 *
 * **Template substitution** uses two field syntaxes:
 * - `{FieldName}` — replaced with the event's value for that field. Date/time
 *   fields are auto-formatted. Address fields get a Google Maps link appended
 *   when the `addManifestAddressLinks` property is enabled.
 * - `{?FieldName}...{FieldName}` — conditional block: the content between the
 *   markers is kept only if the field has a non-blank value; the `{?FieldName}`
 *   opening marker is stripped, leaving the content visible.
 *
 * **Template preparation** (`prepareTemplate()`) scans the template doc for
 * `[BEGIN SectionName]` / `[END SectionName]` marker paragraphs on first use
 * (or after the template is modified) and registers the content between them
 * as Google Docs named ranges so they can be efficiently copied at manifest
 * creation time. Results are cached via the `manifestTemplateLastUpdated_`
 * private document property.
 */

/**
 * Entry point for creating driver manifests for all runs on a chosen date.
 *
 * Prompts the user for a date (defaulting to the date of the active cell if
 * in the Trips or Runs sheet, otherwise tomorrow). Calls `getManifestData()`,
 * `groupManifestDataByRun()`, and `createManifests()` in sequence. Shows a
 * toast with the count of manifests created on completion.
 */
function createManifestsByRunForDate() {
  try {
    const templateDocId = getDocProp("driverManifestTemplateDocId")
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeSheet = ss.getActiveSheet()
    const ui = safeGetUi()
    let defaultDate
    let date
    let runDate
    if (activeSheet.getName() == "Trips") {
      runDate = getValueByHeaderName("Trip Date", getFullRows(activeSheet.getActiveCell()))
    } else if (activeSheet.getName() == "Runs") {
      runDate = getValueByHeaderName("Run Date", getFullRows(activeSheet.getActiveCell()))
    } else {
      // tomorrow, at midnight
      runDate = dateOnly(dateAdd(new Date(), 1))
    }

    if (isValidDate(runDate)) {
      defaultDate = runDate
    } else {
      // tomorrow, at midnight
      defaultDate = dateOnly(dateAdd(new Date(), 1))
    }

    if (ui) {
      let promptResult = ui.prompt("Create Manifests",
          "Enter date for manifests. Leave blank for " + formatDate(defaultDate, null, null),
          ui.ButtonSet.OK_CANCEL)
      const startTime = new Date()
      if (promptResult.getSelectedButton() !== ui.Button.OK) {
        ss.toast("Action cancelled as requested.")
        return
      }
      if (promptResult.getResponseText() == "") {
        date = defaultDate
      } else {
        date = parseDate(promptResult.getResponseText(),"Invalid Date")
      }
      if (!isValidDate(date)) {
        ss.toast("Invalid date, action cancelled.")
        return
      }
    } else {
      date = defaultDate
    }

    const dateFilter = createDateFilterForManifestData(date)
    const manifestData = getManifestData(dateFilter)
    const groupedManifestData = groupManifestDataByRun(manifestData)
    const manifestCount = createManifests(templateDocId, groupedManifestData, getManifestFileNameByRun)
    ss.toast(manifestCount + " created.","Manifest creation complete.")
  } catch(e) { logError(e) }
}

/**
 * Entry point for creating manifests for a user-selected set of trip rows.
 *
 * Reads all selected rows from the active sheet, validates that they contain
 * the required trip fields, and creates one manifest per unique run (date +
 * driver + vehicle combination). Shows a toast if no valid trips are selected.
 */
function createSelectedManifestsByRun() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const activeSheet = ss.getActiveSheet()
    const rangeList = activeSheet.getActiveRangeList().getRanges()
    const templateDocId = getDocProp("driverManifestTemplateDocId")
    let selectedRows = []
    rangeList.forEach(range => {selectedRows.push(...getRangeValuesAsTable(getFullRows(range)))})

    if (!selectedRows.length ||
        !Object.hasOwn(selectedRows[0], "Trip Date") ||
        !Object.hasOwn(selectedRows[0], "Driver ID") ||
        !Object.hasOwn(selectedRows[0], "Vehicle ID")) {
      ss.toast("No trips selected, no manifests created.")
      return
    }

    let runList = []
    selectedRows.forEach(row => {
      let thisRun = {}
      thisRun["Trip Date"] = row["Trip Date"]
      thisRun["Driver ID"] = row["Driver ID"]
      thisRun["Vehicle ID"] = row["Vehicle ID"]
      runList.push(thisRun)
    })

    const runFilter = createRunFilterForManifestData(runList)
    const manifestData = getManifestData(runFilter, activeSheet.getName())
    const groupedManifestData = groupManifestDataByRun(manifestData)
    const manifestCount = createManifests(templateDocId, groupedManifestData, getManifestFileNameByRun)
    ss.toast(manifestCount + " created.","Manifest creation complete.")
  } catch(e) { logError(e) }
}

/**
 * Iterates over a set of grouped manifest data and creates one manifest document
 * per group. Optionally exports each to PDF and/or trashes the source Doc,
 * based on the `createManifestPdf` and `keepManifestDoc` document properties.
 * @param {string} templateDocId - The Drive file ID of the manifest template Doc.
 * @param {Object[]} groupedManifestData - Array of run-grouped manifest objects,
 *   as returned by `groupManifestDataByRun()`.
 * @param {function(Object): string} fileNameFunction - Called with each manifest
 *   group to produce the file name for the created Doc (and PDF).
 * @returns {number} The number of manifest documents created.
 */
function createManifests(templateDocId, groupedManifestData, fileNameFunction) {
  try {
    const manifestFolderId = getDocProp("driverManifestFolderId")
    const templateDoc = DocumentApp.openById(templateDocId)
    prepareTemplate(templateDocId)

    let manifestCount = 0
    groupedManifestData.forEach(manifestGroup => {
      const manifestFileName = fileNameFunction(manifestGroup)
      const manifestDocId = createManifest(manifestGroup, templateDoc, manifestFileName, manifestFolderId)
      if (getDocProp("createManifestPdf")) {
        createPdfFromDocFile(manifestDocId, manifestFileName, manifestFolderId)
      }
      if (getDocProp("sendManifestToDriver")) {
        emailManifestToDriver(manifestGroup, manifestDocId, manifestFileName)
      }
      if (!getDocProp("keepManifestDoc")) {
        Drive.Files.update({ trashed: true }, manifestDocId, null, { supportsAllDrives: true })
      }
      manifestCount++
    })
    return manifestCount
  } catch(e) {
    logError(e)
  }
}

/**
 * Creates a single manifest Google Doc by copying structure from the template.
 *
 * A blank placeholder document is created first (using a `"{{Temporary Text}}"`
 * sentinel so Drive has something to import), then the template's page settings,
 * header, footer, and named-range sections are appended with template field
 * substitution applied via `replaceElementText()`. The temporary first paragraph
 * is removed before saving.
 *
 * Named ranges used from the template: `HEADER`, `PICKUP`, `DROP OFF`, `FOOTER`.
 * Header and footer template substitution uses the first and last event in the
 * group respectively.
 *
 * @param {Object} manifestGroup - A single run group from `groupManifestDataByRun()`.
 *   Must contain `Events` (array of event row objects) and run-level fields.
 * @param {GoogleAppsScript.Document.Document} templateDoc - The opened template Doc.
 * @param {string} manifestFileName - The file name for the new Doc.
 * @param {string} folderId - The Drive folder ID to save the Doc into.
 * @returns {string} The Drive file ID of the newly created manifest Doc.
 */
function createManifest(manifestGroup, templateDoc, manifestFileName, folderId) {
  const templateBody    = templateDoc.getBody()
  const tempText        = "{{Temporary Text}}"
  const manifestDocId   = createDoc(manifestFileName, folderId, tempText, "text/plain")
  const manifestDoc     = DocumentApp.openById(manifestDocId)
  const manifestBody    = manifestDoc.getBody()

  // Update page settings
  manifestBody.setMarginTop(templateBody.getMarginTop())
  manifestBody.setMarginRight(templateBody.getMarginRight())
  manifestBody.setMarginBottom(templateBody.getMarginBottom())
  manifestBody.setMarginLeft(templateBody.getMarginLeft())
  manifestBody.setPageHeight(templateBody.getPageHeight())
  manifestBody.setMarginLeft(templateBody.getMarginLeft())

  // Update page header and page footer
  const templateHeader = templateDoc.getHeader()
  if (templateHeader) {
    const manifestHeader = manifestDoc.addHeader()
    for (let i = 0, c = templateHeader.getNumChildren(); i < c; i++) {
      appendElement(manifestHeader, templateHeader.getChild(i).copy())
    }
    replaceElementText(manifestHeader, manifestGroup["Events"][0])
  }
  const templateFooter = templateDoc.getFooter()
  if (templateFooter) {
    const manifestFooter = manifestDoc.addFooter()
    for (let i = 0, c = templateFooter.getNumChildren(); i < c; i++) {
      appendElement(manifestFooter, templateFooter.getChild(i).copy())
    }
    replaceElementText(manifestFooter, manifestGroup["Events"][manifestGroup["Events"].length - 1])
  }

  // Add the document header elements
  appendTemplateRange(templateDoc.getNamedRanges("HEADER")[0]?.getRange(), manifestBody, manifestGroup["Events"][0])

  // Add all the PU and DO elements. Use the section name of each event to decide whether to add a PU or DO range.
  manifestGroup["Events"].forEach((event, i) => {
    appendTemplateRange(templateDoc.getNamedRanges(event["Section Name"])[0]?.getRange(), manifestBody, event)
  })

  // Add the footer elements
  appendTemplateRange(templateDoc.getNamedRanges("FOOTER")[0]?.getRange(), manifestBody, manifestGroup["Events"][manifestGroup["Events"].length - 1])

  // Remove the tempText needed to create the file
  manifestBody.removeChild(manifestBody.getChild(0))
  manifestDoc.saveAndClose()
  return manifestDocId
}

/**
 * Exports a Google Doc to PDF using the Drive v3 export API and saves the
 * resulting PDF file to the specified folder.
 * @param {string} manifestDocId - The Drive file ID of the source Google Doc.
 * @param {string} manifestFileName - Base name for the PDF file (`.pdf` is appended).
 * @param {string} manifestFolderId - The Drive folder ID to save the PDF into.
 * @returns {string} The Drive file ID of the created PDF file.
 */
function createPdfFromDocFile(manifestDocId, manifestFileName, manifestFolderId) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + manifestDocId + '/export?mimeType=application/pdf'
  const options = {
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    }
  }
  const pdfFileName = manifestFileName + ".pdf"
  const pdfBlob = UrlFetchApp.fetch(url, options).getBlob().setName(pdfFileName)

  const createdPdfFile = Drive.Files.create(
    {
      name: pdfFileName,
      mimeType: 'application/pdf',
      parents: [manifestFolderId]
    },
    pdfBlob,
    {
      supportsAllDrives: true
    })
  return createdPdfFile.id
}

/**
 * Sends a driver manifest to the driver's email address as a PDF attachment.
 * Exports the manifest Google Doc as PDF via the Drive export API.
 * Does nothing if the manifest group has no driver email address.
 * @param {Object} manifestGroup - A run group from `groupManifestDataByRun()`.
 * @param {string} manifestDocId - The Drive file ID of the manifest Google Doc.
 * @param {string} manifestFileName - The base file name used for the attachment.
 */
function emailManifestToDriver(manifestGroup, manifestDocId, manifestFileName) {
  try {
    const driverEmail = manifestGroup["Driver Email"]
    if (!driverEmail) return

    const subject = applyEmailTemplate(manifestEmailSubject, manifestGroup)
    const body = applyEmailTemplate(manifestEmailBody, manifestGroup)

    const url = 'https://www.googleapis.com/drive/v3/files/' + manifestDocId + '/export?mimeType=application/pdf'
    const attachmentBlob = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() } })
      .getBlob()
      .setName(manifestFileName + ".pdf")

    MailApp.sendEmail({ to: driverEmail, subject: subject, body: body, attachments: [attachmentBlob] })
  } catch(e) {
    logError(e)
  }
}

/**
 * Replaces template placeholders in an email subject or body string.
 * Supported placeholders:
 * - `{Driver Name}` — replaced with the driver's name from the manifest group.
 * - `{Trip Date}` — replaced with the formatted trip date.
 * - `{h:mm am/pm}` — replaced with the current time in h:mm AM/PM format.
 * @param {string} template - The template string containing placeholders.
 * @param {Object} manifestGroup - A run group from `groupManifestDataByRun()`.
 * @returns {string} The template with all placeholders substituted.
 */
function applyEmailTemplate(template, manifestGroup) {
  return template
    .replace(/\{Driver Name\}/g, manifestGroup["Driver Name"] || "")
    .replace(/\{Trip Date\}/g, formatDate(manifestGroup["Trip Date"]) || "")
    .replace(/\{h:mm am\/pm\}/g, formatDate(new Date(), null, "h:mm aa"))
}

/**
 * Creates a new Google Doc in the specified Drive folder by uploading a blob
 * of initial content. Used to create the blank placeholder Doc that is then
 * populated by `createManifest()`.
 * @param {string} fileName - The name for the new Google Doc.
 * @param {string} folderId - The Drive folder ID to place the file in.
 * @param {string} content - Initial file content (used as the import blob).
 * @param {string} contentType - MIME type of `content` (e.g. `"text/plain"`).
 * @returns {string} The Drive file ID of the newly created Google Doc.
 * @throws Re-throws any Drive API error to allow callers to handle it.
 */
function createDoc(fileName, folderId, content, contentType) {
  try {
    const blob = Utilities.newBlob(content, contentType)
    const file = Drive.Files.create(
      {
        name: fileName,
        mimeType: 'application/vnd.google-apps.document',
        parents: [folderId]
      },
      blob,
      {
        supportsAllDrives: true
      }
    )
    return file.id
  } catch(e) {
    logError(e)
    // Re-throw to allow callers to handle it
    throw e
  }
}

/**
 * Copies elements from a named range in the template document to a section
 * of the manifest document, applying template field substitution to each element.
 *
 * Before appending, `replaceText()` is called on the element's text to evaluate
 * conditional and regular field markers. Elements are only appended if they will
 * have text content after substitution, or if they contain no field markers at
 * all (e.g. blank lines used for spacing).
 *
 * If `range` is `null` or `undefined` (the named range does not exist in the
 * template), the function returns immediately without error.
 *
 * @param {GoogleAppsScript.Document.Range|undefined} range - The named range from
 *   the template document to copy elements from.
 * @param {GoogleAppsScript.Document.Body|GoogleAppsScript.Document.HeaderSection|GoogleAppsScript.Document.FooterSection} docSection
 *   The destination section of the manifest document to append elements into.
 * @param {Object} data - The event row object used for template field substitution.
 */
function appendTemplateRange(range, docSection, data) {
  if (!range) return
  const rangeElements = range.getRangeElements()
  rangeElements.forEach(rangeElement => {
    const templateElement = rangeElement.getElement()
    const newElement = templateElement.copy()
    if (data) {
      const tempText = replaceText(templateElement.getText(), data)
      // Append the element if it will ultimately have text or
      // if the element has no fields to populate (e.g., it's just a blank line)
      if (tempText.trim() || elementFieldCount(templateElement) === 0) {
        appendElement(docSection, newElement)
        replaceElementText(newElement, data)
      }
    } else {
      appendElement(docSection, newElement)
    }
  })
}

/**
 * Performs in-place template field substitution on a document element.
 *
 * Two passes are made:
 * 1. **Conditional fields** `{?FieldName}...{FieldName}` — if the field has a
 *    non-blank value, the opening `{?FieldName}` marker is stripped and the
 *    content is kept; if the field is blank, the entire block is removed.
 * 2. **Regular fields** `{FieldName}` — replaced with the field value from
 *    `data`. Date/time values are auto-formatted: fields with "date" in the
 *    name use `M/d/yyyy`; fields with "time" use `hh:mm aa`; others use
 *    `hh:mm aa M/d/yy`. Address fields additionally receive a Google Maps
 *    driving-directions hyperlink when `addManifestAddressLinks` is enabled.
 *
 * @param {GoogleAppsScript.Document.Element} element - The document element to
 *   perform substitution on (must support `getText()` and `replaceText()`).
 * @param {Object} data - The event row object whose keys are field names and
 *   values are the substitution values.
 */
function replaceElementText(element, data) {
  let elementText = element.getText()

  // First pass: Process conditional fields {?field}...{field}
  const conditionalPattern = /\{\?([^}]+)\}(.*?)\{\1\}/g
  const conditionalMatches = [...elementText.matchAll(conditionalPattern)]

  conditionalMatches.forEach(match => {
    const fullMatch = match[0]
    const fieldName = match[1]
    if (Object.keys(data).indexOf(fieldName) != -1) {
      const hasValue = data[fieldName] !== null &&
                       data[fieldName] !== undefined &&
                       data[fieldName] !== ''
      if (hasValue) {
        // Field has value - remove just the conditional marker {?field}
        element.replaceText(escapeRegex('{?' + fieldName + '}'), '')
      } else {
        // Field is empty - remove the entire conditional block
        element.replaceText(escapeRegex(fullMatch), '')
      }
    }
  })

  // Second pass: Process regular fields {field}
  elementText = element.getText()
  const pattern = /{(.*?)}/g
  const innerMatches = [...elementText.matchAll(pattern)].map(match => match[1])
  let datum
  innerMatches.forEach(fieldName => {
    if (isValidDate(data[fieldName])) {
      if (fieldName.match(/\bdate\b/i)) {
        datum = formatDate(data[fieldName])
      } else if (fieldName.match(/\btime\b/i)) {
        datum = formatDate(data[fieldName], null, "hh:mm aa")
      } else {
        datum = formatDate(data[fieldName], null, "hh:mm aa M/d/yy")
      }
    } else {
      datum = data[fieldName]
    }
    if (Object.keys(data).indexOf(fieldName) != -1) {
      element.replaceText("{" + fieldName + "}", datum)
      if (fieldName.match(/\baddress\b/i) && datum && getDocProp("addManifestAddressLinks")) {
        const url = createGoogleMapsDirectionsURL(datum)
        const text = element.asText()
        let addressRange = text.findText(escapeRegex(datum))
        if (addressRange) {
          do {
            text.setLinkUrl(addressRange.getStartOffset(), addressRange.getEndOffsetInclusive(), url)
            addressRange = text.findText(escapeRegex(datum), addressRange)
          } while (addressRange)
        }
      } 
    }
  })
}

/**
 * Counts the number of template field placeholders (`{FieldName}`) in a
 * document element's text. Used by `appendTemplateRange()` to determine
 * whether an element with no substituted text should still be appended
 * (e.g. blank spacer lines that have no field markers).
 * @param {GoogleAppsScript.Document.Element} element - The element to inspect.
 * @returns {number} The count of `{...}` field markers found in the element text.
 */
function elementFieldCount(element) {
  const elementText = element.getText()
  const pattern = /{(.*?)}/g
  const matches = [...elementText.matchAll(pattern)]
  return matches.length
}

/**
 * Reads and joins trip, driver, vehicle, and customer data to produce the
 * input for manifest creation.
 *
 * For each trip passing `filterFunction`:
 * - Driver, vehicle, and customer attributes are merged in from their respective
 *   sheets via `mergeAttributes()` (non-conflicting keys are added directly to
 *   the trip row; all attributes are also available under
 *   `_<KeyName>-attributes`).
 * - `"Manifest Creation Time"` and `"Manifest Creation Date"` are set to `new Date()`.
 * - Two event rows are derived: one with `Section Name = "PICKUP"` and one with
 *   `Section Name = "DROP OFF"`. Each gets an `"Event Time"` and a `"Sort Field"`
 *   (formatted as `"HH:mm PU"` or `"HH:mm DO"`).
 * - Events are sorted by `"Sort Field"` (chronological with DO after PU at the
 *   same time).
 *
 * @param {function(Object): boolean} filterFunction - Predicate applied to each
 *   trip row to select which trips to include.
 * @param {string} [tripsSheetName="Trips"] - Name of the sheet to read trips from
 *   (e.g. `"Trip Review"` when creating manifests from reviewed data).
 * @returns {{trips: Object[], events: Object[]}} The filtered/joined trip rows
 *   and the sorted pickup+drop-off event rows.
 */
function getManifestData(filterFunction, tripsSheetName = "Trips") {
  // Get all the raw data
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const drivers = getRangeValuesAsTable(ss.getSheetByName("Drivers").getDataRange())
  const vehicles = getRangeValuesAsTable(ss.getSheetByName("Vehicles").getDataRange())
  const customers = getRangeValuesAsTable(ss.getSheetByName("Customers").getDataRange())
  const trips = getRangeValuesAsTable(ss.getSheetByName(tripsSheetName).getDataRange())
  let manifestTrips = trips.filter(filterFunction)
  
  // Pull in the lookup table data
  manifestTrips.forEach(tripRow => {
    mergeAttributes(tripRow, drivers,   "Driver ID"  )
    mergeAttributes(tripRow, vehicles,  "Vehicle ID" )
    mergeAttributes(tripRow, customers, "Customer ID")
    tripRow["Manifest Creation Time"] = new Date()
    tripRow["Manifest Creation Date"] = new Date()
  })
  
  // For events, create two rows for each trip -- one for PU, and one for DO
  let pickups = manifestTrips.map(tripRow => {
    let newRow = Object.assign({},tripRow)
    newRow["Section Name"] = "PICKUP"
    newRow["Event Name"]   = "pickup"
    newRow["Event Time"]   = tripRow["PU Time"]
    newRow["Sort Field"]   = formatDate(new Date(tripRow["PU Time"]), null, "HH:mm") + " PU"
    return newRow
  })
  let dropOffs = manifestTrips.map(tripRow => {
    let newRow = Object.assign({},tripRow)
    newRow["Section Name"] = "DROP OFF"
    newRow["Event Name"]   = "drop off"
    newRow["Event Time"]   = tripRow["DO Time"]
    newRow["Sort Field"]   = formatDate(new Date(tripRow["DO Time"]), null, "HH:mm") + " DO"
    return newRow
  })
  let manifestEvents = pickups.concat(dropOffs).sort((a,b) => {
    if (a["Sort Field"] < b["Sort Field"]) { return -1 }
    if (a["Sort Field"] > b["Sort Field"]) { return  1 }
    return 0
  })

  const result = {
    "trips": manifestTrips,
    "events": manifestEvents
  }
  return result
}

/**
 * Groups manifest trips and events into per-run manifest group objects.
 *
 * Each group represents a unique combination of Trip Date + Driver ID +
 * Vehicle ID + Run ID. The group object contains:
 * - `"Driver ID"`, `"Vehicle ID"`, `"Run ID"`, `"Driver Name"`, `"Driver Email"`,
 *   `"Vehicle Name"`, `"Trip Date"` — from the first trip in the group.
 * - `"Trips"` — array of trip row objects belonging to the run.
 * - `"Events"` — array of event row objects (pickups + drop-offs) belonging
 *   to the run, in the order they appear in `manifestData.events` (already
 *   sorted chronologically by `getManifestData()`).
 *
 * @param {{trips: Object[], events: Object[]}} manifestData - As returned by
 *   `getManifestData()`.
 * @returns {Object[]} Array of manifest group objects, one per unique run.
 */
function groupManifestDataByRun(manifestData) {
  // Group the trips into runs -- A run is a collection of trips on the same day
  // with the same driver, vehicle, and run id
  // TODO add date and run ID to code
  let manifestGroups = []
  manifestData.trips.forEach(trip => {
    let runIndex = manifestGroups.findIndex(r =>
      r["Trip Date"].getTime() == trip["Trip Date"].getTime() &&
      r["Driver ID"] == trip["Driver ID"] &&
      r["Vehicle ID"] == trip["Vehicle ID"] &&
      r["Run ID"] == trip["Run ID"]
    )
    if (runIndex == -1) {
      let newRun = {}
      newRun["Driver ID"]    = trip["Driver ID"]
      newRun["Vehicle ID"]   = trip["Vehicle ID"]
      newRun["Run ID"]       = trip["Run ID"]
      newRun["Driver Name"]  = trip["Driver Name"]
      newRun["Driver Email"] = trip["Driver Email"]
      newRun["Vehicle Name"] = trip["Vehicle Name"]
      newRun["Trip Date"]    = trip["Trip Date"]
      newRun["Trips"]        = [trip]
      newRun["Events"]       = []
      manifestGroups.push(newRun)
    } else {
      manifestGroups[runIndex]["Trips"].push(trip)
    }
  })
  // Group the manifest events into the same runs
  manifestData.events.forEach(event => {
    let matchedRun = manifestGroups.find(run =>
      run["Trip Date"].getTime() == event["Trip Date"].getTime() &&
      run["Driver ID"] == event["Driver ID"] &&
      run["Vehicle ID"] == event["Vehicle ID"] &&
      run["Run ID"] == event["Run ID"]
    )
    matchedRun["Events"].push(event)
  })
  return manifestGroups
}

/**
 * Legacy manifest creation function. Incomplete — `manifestSections` is computed
 * but never used, and the function has no return value or write-back logic.
 * @deprecated Use `createManifestsByRunForDate()` or `createSelectedManifestsByRun()` instead.
 * @private
 */
function createDriverManifest(manifestDate, driverId) {
  const driverManifestFolderId = getDocProp("driverManifestFolderId")
  const driverManifestTemplateDocId = getDocProp("driverManifestTemplateDocId")
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const filter = function(row) { return row["Driver ID"] == driverId }
  const driverRow = findFirstRowByHeaderNames(ss.getSheetByName("Drivers"), filter)
  const driverName = driverRow["Driver Name"]
  const manifestFileName = formatDate(manifestDate, null, "yyyy-MM-dd") + " Manifest for " + driverName
  const manifestFolder = DriveApp.getFolderById(driverManifestFolderId)
  const manifestFile = DriveApp.getFileById(driverManifestTemplateDocId).makeCopy(manifestFolder).setName(manifestFileName)
  const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  const manifestDoc = DocumentApp.openById(manifestFile.getId())
  deleteAllNamedRanges(templateDoc)
  const manifestSections = templateDoc.getBody().getParent()
}

/**
 * Returns `true` if the given document element type supports `getText()` and
 * `editAsText()`. Used by `prepareTemplate()` to safely skip element types
 * that do not have text content (e.g. `INLINE_IMAGE`, `PAGE_BREAK`).
 * @param {GoogleAppsScript.Document.Element} element - The document element to test.
 * @returns {boolean}
 */
// Element types without text: COMMENT_SECTION, DOCUMENT, EQUATION_FUNCTION_ARGUMENT_SEPARATOR,
// EQUATION_SYMBOL, FOOTNOTE, HORIZONTAL_RULE, INLINE_DRAWING, INLINE_IMAGE, PAGE_BREAK, UNSUPPORTED
function elementHasText(element) {
  const elementTypesWithText = [DocumentApp.ElementType.BODY_SECTION,
                          DocumentApp.ElementType.EQUATION,
                          DocumentApp.ElementType.EQUATION_FUNCTION,
                          DocumentApp.ElementType.FOOTER_SECTION,
                          DocumentApp.ElementType.FOOTNOTE_SECTION,
                          DocumentApp.ElementType.HEADER_SECTION,
                          DocumentApp.ElementType.LIST_ITEM,
                          DocumentApp.ElementType.PARAGRAPH,
                          DocumentApp.ElementType.TABLE,
                          DocumentApp.ElementType.TABLE_CELL,
                          DocumentApp.ElementType.TABLE_OF_CONTENTS,
                          DocumentApp.ElementType.TABLE_ROW,
                          DocumentApp.ElementType.TEXT]
  return (elementTypesWithText.indexOf(element.getType()) > -1)
}

/**
 * Removes all named ranges from a Google Doc. Called on the template before
 * re-scanning it for `[BEGIN]`/`[END]` markers in `prepareTemplate()`.
 * @param {GoogleAppsScript.Document.Document} doc - The document to clear named ranges from.
 */
function deleteAllNamedRanges(doc) {
  doc.getNamedRanges().forEach(range => {
    range.remove()
  })
}

/**
 * Scans the template document for `[BEGIN SectionName]` / `[END SectionName]`
 * marker paragraphs and registers the content between each pair as a Google
 * Docs named range (e.g. `"HEADER"`, `"PICKUP"`, `"DROP OFF"`, `"FOOTER"`).
 * An outer named range (`"OUTER_SectionName"`) spanning the markers themselves
 * is also registered but is not used during manifest creation.
 *
 * Preparation is skipped if the template's Drive `modifiedTime` has not changed
 * since the last run (cached in the `manifestTemplateLastUpdated_` private
 * document property). After scanning, the template is saved and the timestamp
 * is updated.
 *
 * @param {string} driverManifestTemplateDocId - The Drive file ID of the template Doc.
 */
function prepareTemplate(driverManifestTemplateDocId) {
  const lastUpdated = getFileLastUpdated(driverManifestTemplateDocId)

  if (lastUpdated > getDocProp("manifestTemplateLastUpdated_")) {
    const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
    deleteAllNamedRanges(templateDoc)
    const body = templateDoc.getBody()
    for (let i = 0, c = body.getNumChildren(); i < c; i++) {
      let element = body.getChild(i)
      if (elementHasText(element)) {
        const match = element.getText().match(/^\s*\[BEGIN (?<sectionName>.+?)\]\s*$/)
        if (match) {
          // We're at the beginning of a section. 
          // We want to:
          // - Jump to the next element
          // - Set up a named range with a name matching the section name
          // - Begin a loop where we:
          //   - Check each element to see if it's the closing element
          //   - If it's not the closing element, add that element to a named range
          //   - If it's the closing element, complete the building of the named range and and exit the loop
          const outerRangeBuilder = templateDoc.newRange()
          const innerRangeBuilder = templateDoc.newRange()
          outerRangeBuilder.addElement(element)
          if (i < c) { i++ }
          const sectionName = match.groups["sectionName"]
          const regex = new RegExp(`^\\s*\\[END ${sectionName}\\]\\s*$`)
          for (let stayInLoop = true; stayInLoop && i < c ; i++) {
            element = body.getChild(i)
            if (elementHasText(element) && element.getText().match(regex)) {
              outerRangeBuilder.addElement(element)
              templateDoc.addNamedRange(`OUTER_${sectionName}`, outerRangeBuilder.build())
              if (innerRangeBuilder.getRangeElements().length > 0) {
                templateDoc.addNamedRange(sectionName, innerRangeBuilder.build())
              }
              stayInLoop = false
            } else {
              outerRangeBuilder.addElement(element)
              innerRangeBuilder.addElement(element)
            }
          }
        }
      }
    }
    templateDoc.saveAndClose()
    setDocProp("manifestTemplateLastUpdated_", getFileLastUpdated(driverManifestTemplateDocId))
  }
}

/**
 * Development utility for testing named-range copying between documents.
 * Creates a test manifest Doc from the template and appends each named range's
 * elements. Not used in production manifest creation.
 * @private
 */
function copyNamedRanges(source, destination) {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const driverManifestFolderId = getDocProp("driverManifestFolderId")
  const driverManifestTemplateDocId = getDocProp("driverManifestTemplateDocId")
  const manifestFileName = formatDate(new Date(), null, "yyyy-MM-dd") + " Test Manifest"
  const manifestFolder = DriveApp.getFolderById(driverManifestFolderId)
  const manifestFile = DriveApp.getFileById(driverManifestTemplateDocId).makeCopy(manifestFolder).setName(manifestFileName)
  const templateDoc = DocumentApp.openById(driverManifestTemplateDocId)
  const manifestDoc = DocumentApp.openById(manifestFile.getId())
  const manifestBody = manifestDoc.getBody()

  templateDoc.getNamedRanges().forEach(range => {
    //range = templateDoc.getNamedRanges()[0]
    range.getRange().getRangeElements().forEach(element => {
      //element = range.getRange().getRangeElements()[0]
      appendElement(manifestBody, element.getElement().copy())
    })
  })
}

/**
 * Appends a document element to a body, header, or footer section using the
 * appropriate `append*` method for the element's type. Supports: `PARAGRAPH`,
 * `TABLE`, `LIST_ITEM`, `HORIZONTAL_RULE`, `INLINE_IMAGE`, `PAGE_BREAK`.
 * Elements of unrecognised types are silently skipped.
 * @param {GoogleAppsScript.Document.Body|GoogleAppsScript.Document.HeaderSection|GoogleAppsScript.Document.FooterSection} body
 *   The destination section to append to.
 * @param {GoogleAppsScript.Document.Element} element - The element to append.
 */
function appendElement(body, element) {
  let type = element.getType()
  if (type == DocumentApp.ElementType.PARAGRAPH) {
    body.appendParagraph(element)
  } else if (type == DocumentApp.ElementType.TABLE) {
    body.appendTable(element)
  } else if (type == DocumentApp.ElementType.LIST_ITEM) {
    body.appendListItem(element)
  } else if (type == DocumentApp.ElementType.HORIZONTAL_RULE) {
    body.appendHorizontalRule(element)
  } else if (type == DocumentApp.ElementType.INLINE_IMAGE) {
    body.appendImage(element)
  } else if (type == DocumentApp.ElementType.PAGE_BREAK) {
    body.appendPageBreak(element)
  } 
}

/**
 * Merges attributes from a matching row in a secondary lookup table into a
 * primary row object. The match is made by comparing `primaryRow[primaryKeyName]`
 * against `secondaryRow[secondaryKeyName]`.
 *
 * On a match:
 * - All keys from the secondary row that are not already present in the primary
 *   row are copied directly into the primary row.
 * - The full secondary row is stored under `primaryRow["_" + primaryKeyName + "-attributes"]`
 *   for downstream access without key collision risk.
 *
 * Does nothing if the primary row's key field is blank or no match is found.
 *
 * @param {Object} primaryRow - The row object to merge attributes into (mutated in place).
 * @param {Object[]} secondaryTable - Array of row objects to search for a match.
 * @param {string} primaryKeyName - The field name in `primaryRow` to match on.
 * @param {string} [secondaryKeyName=primaryKeyName] - The field name in the secondary
 *   row to match against. Defaults to `primaryKeyName` if omitted.
 */
function mergeAttributes(primaryRow, secondaryTable, primaryKeyName, secondaryKeyName) {
  secondaryKeyName = secondaryKeyName || primaryKeyName
  if (primaryRow[primaryKeyName]) {
    let matchingSecondaryRow = secondaryTable.find(secondaryRow => primaryRow[primaryKeyName] == secondaryRow[secondaryKeyName])
    if (matchingSecondaryRow) {
      primaryRow["_" + primaryKeyName + "-attributes"] = matchingSecondaryRow
      Object.keys(matchingSecondaryRow).forEach(key => {
        if (!primaryRow.hasOwnProperty(key)) {
          primaryRow[key] = matchingSecondaryRow[key]
        }
      })
    }
  }
}

/**
 * Generates a manifest file name for a run group in the format:
 * `"YYYY-MM-DD manifest for <Driver Name> on <Vehicle Name>"`.
 * @param {Object} manifestGroup - A manifest group object as returned by
 *   `groupManifestDataByRun()`, with `"Trip Date"`, `"Driver Name"`, and
 *   `"Vehicle Name"` fields.
 * @returns {string} The formatted file name string.
 */
function getManifestFileNameByRun(manifestGroup) {
  const manifestFileName = `${formatDate(manifestGroup["Trip Date"], null, "yyyy-MM-dd")} manifest for ${manifestGroup["Driver Name"]} on ${manifestGroup["Vehicle Name"]}`
  return manifestFileName
}

/**
 * Returns a filter function that selects trips whose `"Trip Date"` matches
 * the given date exactly (by valueOf comparison).
 * @param {Date} date - The date to filter by.
 * @returns {function(Object): boolean}
 */
function createDateFilterForManifestData(date) {
  return function(trip) {
    return new Date(trip["Trip Date"]).valueOf() === date.valueOf()
  }
}

/**
 * Returns a filter function that selects trips matching any run in the provided
 * list by date + driver + vehicle (Run ID is not checked here, since selected
 * rows may not have a Run ID assigned yet).
 * @param {Array<{"Trip Date": Date, "Driver ID": string, "Vehicle ID": string}>} runs
 *   Array of run-identifying objects (typically derived from selected trip rows).
 * @returns {function(Object): boolean}
 */
function createRunFilterForManifestData(runs) {
  return function (trip) {
    return runs.some(run => {
      return trip["Trip Date"] instanceof Date &&
              trip["Trip Date"].getTime() === run["Trip Date"].getTime() &&
              trip["Driver ID"] === run["Driver ID"] &&
              trip["Vehicle ID"] === run["Vehicle ID"]
    })
  }
}

/**
 * Returns the last-modified timestamp of a Drive file as a millisecond epoch value.
 * Used by `prepareTemplate()` to detect whether the template has changed since
 * it was last scanned.
 * @param {string} fileId - The Drive file ID to check.
 * @returns {number} Milliseconds since Unix epoch of the file's `modifiedTime`.
 */
function getFileLastUpdated(fileId) {
  const fileMetadata = Drive.Files.get(
    fileId,
    {
      fields: 'modifiedTime',
      supportsAllDrives: true
    }
  )
  return new Date(fileMetadata.modifiedTime).getTime()
}