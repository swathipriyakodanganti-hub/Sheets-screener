// ============================================================
//  CAMPAIGN PORTAL — Google Apps Script Backend
//  Pearl Hire / BTL Recruitment
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
// Replace this with the URL of YOUR Registry Sheet
// (the one GAS will use to store all campaign links)
const REGISTRY_URL = "https://docs.google.com/spreadsheets/d/1_kSQNnquFNcZ0UN3sp80t65iPlFHZBSMfpn_PvY_9TU/edit?gid=0#gid=0";

// Column indices in the Registry Sheet (0-based)
const COL = {
  ID:       0,  // A — auto-generated ID
  NAME:     1,  // B — campaign label (e.g. "Graphic + Marketing Sep 2026")
  URL:      2,  // C — Google Sheet URL
  ADDED_ON: 3,  // D — date added
  ROW_COUNT:4,  // E — cached row count
};

// Column headers expected in campaign sheets (form responses)
// Adjust these to match your exact Google Form column names
const FORM_COLS = {
  TIMESTAMP:  "Timestamp",
  NAME:       "Full Name",
  PHONE:      "Phone number",
  EMAIL:      "Email",
  LOCATION:   "Your current location?",
  GRAD_YEAR:  "Graduation year",
  ROLE:       "Applying for?",
  RESUME:     "Resume",
  PORTFOLIO:  "Please attach your portfolio if you are applying",
};


// ── ROUTER ──────────────────────────────────────────────────
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    const params = e.parameter || {};
    const action = params.action || "";

    let result;
    switch (action) {
      case "addCampaign":    result = addCampaign(params.name, params.url);             break;
      case "listCampaigns":  result = listCampaigns();                                  break;
      case "deleteCampaign": result = deleteCampaign(params.id);                        break;
      case "getRows":        result = getRows(params.url);                              break;
      case "shortlist":      result = shortlistCandidate(params.url, params.rowIndex);  break;
      case "getShortlist":   result = getShortlist(params.url);                        break;
      case "removeShortlist":result = removeShortlist(params.url, params.rowIndex);    break;
      default:               result = { error: "Unknown action: " + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── REGISTRY HELPERS ────────────────────────────────────────
function getRegistrySheet() {
  const ss = SpreadsheetApp.openByUrl(REGISTRY_URL);
  let sheet = ss.getSheetByName("Campaigns");
  if (!sheet) {
    sheet = ss.insertSheet("Campaigns");
    sheet.appendRow(["ID", "Campaign Name", "Sheet URL", "Added On", "Row Count"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function generateId() {
  return "c_" + Date.now().toString(36);
}


// ── ACTION: addCampaign ─────────────────────────────────────
// Saves a new campaign URL + label to the registry
function addCampaign(name, url) {
  if (!name || !url) throw new Error("name and url are required");

  const sheet = getRegistrySheet();
  const id    = generateId();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy");

  // Try to get row count from the target sheet
  let rowCount = 0;
  try {
    const target = SpreadsheetApp.openByUrl(url);
    rowCount = Math.max(0, target.getSheets()[0].getLastRow() - 1); // minus header
  } catch (_) {}

  sheet.appendRow([id, name, url, today, rowCount]);
  return { success: true, id, name, url, addedOn: today, rowCount };
}


// ── ACTION: listCampaigns ───────────────────────────────────
// Returns all saved campaigns from the registry
function listCampaigns() {
  const sheet = getRegistrySheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return { campaigns: [] };

  const campaigns = data.slice(1).map(row => ({
    id:       row[COL.ID],
    name:     row[COL.NAME],
    url:      row[COL.URL],
    addedOn:  row[COL.ADDED_ON],
    rowCount: row[COL.ROW_COUNT],
  }));

  return { campaigns };
}


// ── ACTION: deleteCampaign ──────────────────────────────────
// Removes a campaign from the registry by ID
function deleteCampaign(id) {
  if (!id) throw new Error("id is required");

  const sheet = getRegistrySheet();
  const data  = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.ID] === id) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw new Error("Campaign not found: " + id);
}


// ── ACTION: getRows ─────────────────────────────────────────
// Reads candidate rows from a campaign sheet
function getRows(url) {
  if (!url) throw new Error("url is required");

  const ss     = SpreadsheetApp.openByUrl(url);
  const sheet  = ss.getSheets()[0];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { rows: [] };

  const headers = values[0];

  // Map header name → column index
  const idx = {};
  Object.entries(FORM_COLS).forEach(([key, label]) => {
    const i = headers.findIndex(h => h.toString().trim().toLowerCase().includes(label.toLowerCase()));
    idx[key] = i >= 0 ? i : -1;
  });

  // Check which rows are already shortlisted
  const slSheet    = getOrCreateShortlistSheet(ss);
  const slData     = slSheet.getDataRange().getValues().slice(1);
  const shortlisted = new Set(slData.map(r => String(r[0]))); // row index stored in col A

  const rows = values.slice(1).map((row, i) => ({
    rowIndex:   i + 2,  // 1-based sheet row (row 1 = header, row 2 = first data row)
    timestamp:  idx.TIMESTAMP  >= 0 ? row[idx.TIMESTAMP]  : "",
    name:       idx.NAME       >= 0 ? row[idx.NAME]       : "",
    phone:      idx.PHONE      >= 0 ? row[idx.PHONE]      : "",
    email:      idx.EMAIL      >= 0 ? row[idx.EMAIL]      : "",
    location:   idx.LOCATION   >= 0 ? row[idx.LOCATION]   : "",
    gradYear:   idx.GRAD_YEAR  >= 0 ? row[idx.GRAD_YEAR]  : "",
    role:       idx.ROLE       >= 0 ? row[idx.ROLE]       : "",
    resumeUrl:  idx.RESUME     >= 0 ? row[idx.RESUME]     : "",
    portfolio:  idx.PORTFOLIO  >= 0 ? row[idx.PORTFOLIO]  : "",
    shortlisted: shortlisted.has(String(i + 2)),
  }));

  return { rows, sheetTitle: sheet.getName() };
}


// ── SHORTLIST HELPERS ───────────────────────────────────────
function getOrCreateShortlistSheet(ss) {
  let sheet = ss.getSheetByName("Shortlisted");
  if (!sheet) {
    sheet = ss.insertSheet("Shortlisted");
    sheet.appendRow([
      "Row Index", "Name", "Phone", "Email",
      "Location", "Grad Year", "Position",
      "Applied On", "Shortlisted On"
    ]);
    sheet.setFrozenRows(1);

    // Light green header
    sheet.getRange(1, 1, 1, 9)
         .setBackground("#e6f4ea")
         .setFontWeight("bold");
  }
  return sheet;
}


// ── ACTION: shortlist ───────────────────────────────────────
// Adds a candidate to the Shortlisted tab of their campaign sheet
function shortlistCandidate(url, rowIndex) {
  if (!url || !rowIndex) throw new Error("url and rowIndex are required");

  const ss       = SpreadsheetApp.openByUrl(url);
  const mainSheet = ss.getSheets()[0];
  const slSheet  = getOrCreateShortlistSheet(ss);

  // Check not already shortlisted
  const existing = slSheet.getDataRange().getValues().slice(1);
  if (existing.some(r => String(r[0]) === String(rowIndex))) {
    return { success: true, alreadyExists: true };
  }

  // Read the candidate's row from the main sheet
  const values  = mainSheet.getDataRange().getValues();
  const headers = values[0];
  const row     = values[rowIndex - 1]; // rowIndex is 1-based

  const idx = {};
  Object.entries(FORM_COLS).forEach(([key, label]) => {
    const i = headers.findIndex(h => h.toString().trim().toLowerCase().includes(label.toLowerCase()));
    idx[key] = i >= 0 ? i : -1;
  });

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd MMM yyyy");

  // Format applied date from timestamp
  let appliedOn = "";
  if (idx.TIMESTAMP >= 0 && row[idx.TIMESTAMP]) {
    try {
      appliedOn = Utilities.formatDate(new Date(row[idx.TIMESTAMP]), Session.getScriptTimeZone(), "dd MMM yyyy");
    } catch (_) {
      appliedOn = String(row[idx.TIMESTAMP]);
    }
  }

  slSheet.appendRow([
    rowIndex,
    idx.NAME      >= 0 ? row[idx.NAME]      : "",
    idx.PHONE     >= 0 ? row[idx.PHONE]     : "",
    idx.EMAIL     >= 0 ? row[idx.EMAIL]     : "",
    idx.LOCATION  >= 0 ? row[idx.LOCATION]  : "",
    idx.GRAD_YEAR >= 0 ? row[idx.GRAD_YEAR] : "",
    idx.ROLE      >= 0 ? row[idx.ROLE]      : "",
    appliedOn,
    today,
  ]);

  return { success: true, shortlistedOn: today };
}


// ── ACTION: getShortlist ────────────────────────────────────
// Returns all shortlisted candidates for a campaign
function getShortlist(url) {
  if (!url) throw new Error("url is required");

  const ss      = SpreadsheetApp.openByUrl(url);
  const slSheet = getOrCreateShortlistSheet(ss);
  const data    = slSheet.getDataRange().getValues();
  if (data.length <= 1) return { shortlist: [] };

  const shortlist = data.slice(1).map(row => ({
    rowIndex:      row[0],
    name:          row[1],
    phone:         row[2],
    email:         row[3],
    location:      row[4],
    gradYear:      row[5],
    role:          row[6],
    appliedOn:     row[7],
    shortlistedOn: row[8],
  }));

  return { shortlist };
}


// ── ACTION: removeShortlist ─────────────────────────────────
// Removes a candidate from the Shortlisted tab
function removeShortlist(url, rowIndex) {
  if (!url || !rowIndex) throw new Error("url and rowIndex are required");

  const ss      = SpreadsheetApp.openByUrl(url);
  const slSheet = getOrCreateShortlistSheet(ss);
  const data    = slSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(rowIndex)) {
      slSheet.deleteRow(i + 1);
      return { success: true };
    }
  }

  return { success: false, message: "Row not found in shortlist" };
}