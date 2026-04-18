// ═══════════════════════════════════════════════════════════════════
// HELIX INDUSTRIES — Apps Script  (paste this into Extensions > Apps Script)
// ═══════════════════════════════════════════════════════════════════

var SHEET_ID = "1bnYCSCMFb6FeU1pPaE66smxYu4xGRUO_Pq2egxIpskM";

var DISPATCH_SHEETS = [
  "Dispatch - Mukesh",
  "Dispatch - Subhash",
  "Dispatch - Anandaram",
  "Dispatch - Guard"
];

var MUKESH_VEHICLES = [
  "RJ 19 RF 8056",
  "RJ 07 RA 9480",
  "RJ 19 GJ 9279",
  "RJ 19 GG 2398"
];

var DISPATCH_HEADERS = [
  "Date","Time","Entered By","Party Name",
  "Bill/Challan Number","Vehicle No.",
  "Machine","Product","Thickness","Color","Material Quantity"
];


// ── doGet: connectivity test ──────────────────────────────────────────────────
function doGet(e) {
  return json_({status:"ready"});
}


// ── doPost: main entry point ──────────────────────────────────────────────────
function doPost(e) {
  try {
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var data = JSON.parse(e.postData.contents);

    if (data.testOnly) return json_({status:"ok", test:true});

    // writeAll: atomic multi-sheet write sent by the app
    if (data.action === "writeAll") {
      // Server-side duplicate guard: skip if this rowId was already written
      if (data.rowId && alreadyWritten_(ss, data.rowId)) {
        return json_({status:"ok", skipped:true, rowId:data.rowId});
      }
      var summaryRows = 0;
      (data.sheets || []).forEach(function(s) {
        if (!s.rows || s.rows.length === 0) return;
        appendRows_(ss, s.sheet, s.headers, s.rows);
        if (s.rebuildSummaryAfter) summaryRows = rebuildAll_(ss);
      });
      if (data.rowId) markWritten_(ss, data.rowId);
      return json_({status:"ok", rowId:data.rowId, summaryRows:summaryRows});
    }

    return json_({status:"error", message:"Unknown action"});

  } catch(err) {
    return json_({status:"error", message:String(err)});
  }
}


// ── appendRows_: append rows to a sheet, create with header if new ────────────
// RULE: never touches a sheet unless rows.length > 0
function appendRows_(ss, name, headers, rows) {
  if (!rows || rows.length === 0) return;
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    styleHeader_(sh, headers.length);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    styleHeader_(sh, headers.length);
  }
  rows.forEach(function(r) { sh.appendRow(r); });
}


// ── rebuildAll_: rebuild Outward, Combined, Daily Summary from Dispatch sheets ─
// RULE: never creates a sheet unless there is data to write into it
function rebuildAll_(ss) {
  var tz = Session.getScriptTimeZone();

  // Read all rows from every Dispatch sheet
  var allRows = [];
  DISPATCH_SHEETS.forEach(function(sName) {
    var sh = ss.getSheetByName(sName);
    if (!sh || sh.getLastRow() < 2) return;
    var nc   = sh.getLastColumn();
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, nc).getValues();
    rows.forEach(function(row) {
      // Normalise date (col 0) to dd/MM/yyyy string
      var d = row[0];
      row[0] = (d instanceof Date && !isNaN(d))
               ? Utilities.formatDate(d, tz, "dd/MM/yyyy")
               : String(d).trim();
      row[2] = String(row[2]).trim(); // Entered By
      // Trim to 11 cols — ignore any extra columns users may have added
      allRows.push(row.slice(0, DISPATCH_HEADERS.length));
    });
  });

  if (allRows.length === 0) return 0; // no data — do nothing, create nothing

  // Sort by date ascending
  allRows.sort(function(a, b) { return parseDate_(a[0]) - parseDate_(b[0]); });

  // Outward = all rows
  writeSheet_(ss, "Outward", DISPATCH_HEADERS, allRows);

  // Combined = filtered rows
  var combinedRows = allRows.filter(function(row) {
    var who = row[2], veh = String(row[5]).trim();
    if (who === "Guard")  return false;
    if (who === "Mukesh") return MUKESH_VEHICLES.indexOf(veh) === -1;
    return true;
  });
  if (combinedRows.length > 0) {
    writeSheet_(ss, "Combined", DISPATCH_HEADERS, combinedRows);
  }

  // Daily Summary = aggregated from Combined
  if (combinedRows.length === 0) return 0;

  var totals = {};
  combinedRows.forEach(function(row) {
    var d  = row[0];
    // Use dashes so Sheets does NOT auto-convert "25-02-2026" to a Date object
    var ds = (d instanceof Date && !isNaN(d))
             ? Utilities.formatDate(d, tz, "dd-MM-yyyy")
             : String(d).replace(/\//g, "-");
    var mac = String(row[6] || "").trim();
    var pro = String(row[7] || "").trim();
    var thk = String(row[8] || "").trim();
    var col = String(row[9] || "").trim();
    var qty = Number(row[10]) || 0;
    if (!mac || !ds) return;
    var key = [ds, mac, pro, thk, col].join("|");
    totals[key] = (totals[key] || 0) + qty;
  });

  var SH = ["Date","Machine","Product","Thickness","Color","Total Dispatched"];
  var summaryRows = Object.keys(totals).map(function(k) {
    var p = k.split("|");
    return [p[0], p[1], p[2], p[3], p[4], totals[k]];
  }).sort(function(a, b) {
    var d = parseDashDate_(a[0]) - parseDashDate_(b[0]);
    return d !== 0 ? d : String(a[1]).localeCompare(String(b[1]));
  });

  writeSheet_(ss, "Daily Summary", SH, summaryRows);
  return summaryRows.length;
}


// ── writeSheet_: clear and rewrite a sheet completely ─────────────────────────
// RULE: only creates the sheet if rows.length > 0 — never creates blank tabs
function writeSheet_(ss, name, headers, rows) {
  if (!rows || rows.length === 0) return; // hard guard — never create blank tabs

  var nc = headers.length;
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clearContents();

  // Pad/trim every row to exactly nc columns
  var safe = rows.map(function(row) {
    var r = row.slice(0, nc);
    while (r.length < nc) r.push("");
    return r;
  });

  var data = [headers].concat(safe);
  sh.getRange(1, 1, data.length, nc).setValues(data);
  styleHeader_(sh, nc);
}


// ── styleHeader_: format the header row ──────────────────────────────────────
function styleHeader_(sh, nc) {
  var h = sh.getRange(1, 1, 1, nc);
  h.setFontWeight("bold").setBackground("#2c3b48").setFontColor("#f7b810");
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, nc, 140);
}


// ── Deduplication: track written rowIds in a hidden sheet ─────────────────────
var ID_SHEET = "WrittenIDs";
var MAX_IDS  = 5000;

function alreadyWritten_(ss, rowId) {
  var sh = ss.getSheetByName(ID_SHEET);
  if (!sh || sh.getLastRow() === 0) return false;
  var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues()
              .map(function(r) { return String(r[0]); });
  return ids.indexOf(String(rowId)) !== -1;
}

function markWritten_(ss, rowId) {
  var sh = ss.getSheetByName(ID_SHEET);
  if (!sh) { sh = ss.insertSheet(ID_SHEET); sh.hideSheet(); }
  sh.appendRow([String(rowId), new Date()]);
  if (sh.getLastRow() > MAX_IDS) {
    sh.deleteRows(1, sh.getLastRow() - MAX_IDS);
  }
}


// ── onEdit: fires on manual cell edits ────────────────────────────────────────
function onEdit(e) {
  try {
    var name    = e.range.getSheet().getName();
    var watched = DISPATCH_SHEETS.concat(["Outward","Combined"]);
    if (watched.indexOf(name) === -1) return;
    rebuildAll_(SpreadsheetApp.getActiveSpreadsheet());
  } catch(err) { /* ignore errors from manual runs */ }
}


// ── manualRebuild: run from Apps Script editor to force full rebuild ──────────
function manualRebuild() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var count = rebuildAll_(ss);
  SpreadsheetApp.getUi().alert("✓ Rebuild complete — Daily Summary: " + count + " rows.");
}


// ── Date parsing helpers ──────────────────────────────────────────────────────
function parseDate_(s) {      // "25/02/2026"
  var p = String(s).split("/");
  return p.length === 3 ? new Date(p[2], p[1]-1, p[0]).getTime() : 0;
}
function parseDashDate_(s) {  // "25-02-2026"
  var p = String(s).split("-");
  return p.length === 3 ? new Date(p[2], p[1]-1, p[0]).getTime() : 0;
}


// ── JSON response helper ──────────────────────────────────────────────────────
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}