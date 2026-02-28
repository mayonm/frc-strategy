// ============================================================
// SA_DATA_MASTER → TEAM SHEETS
//
// Reads every row from "SA_DATA_MASTER" and copies it into
// the matching team-number tab, pasted 5 rows below the
// last row that contains ANY data.
//
// HOW TO USE:
//   1. Paste into Extensions → Apps Script (new file).
//   2. Set TEAM_NUMBER_COLUMN below to whatever column your
//      team numbers live in (A, B, C … or 1, 2, 3 …).
//   3. Run pushSADataToTeamSheets() — or use the menu.
//
// RUN IT AGAIN ANYTIME: it checks what's already been copied
// using a hidden tracking column so it never double-pastes.
// ============================================================


// ── !! CHANGE THIS TO MATCH YOUR DATA !! ─────────────────────────────────────

const SA_TEAM_COLUMN = "A";   // Column letter (or number) where team numbers live
                               // e.g. "A", "B", "C", "D" … or 1, 2, 3 …

// ─────────────────────────────────────────────────────────────────────────────


const SA_SOURCE_SHEET   = "SA_DATA_MASTER";  // Source tab name
const SA_ROWS_GAP       = 5;                 // Blank rows to leave above pasted block
const SA_TRACKING_COL   = "SA_PASTED";       // Header name written in tracking column
                                              // (added to SA_DATA_MASTER automatically)


// ── MAIN ─────────────────────────────────────────────────────────────────────

function pushSADataToTeamSheets() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SA_SOURCE_SHEET);

  if (!src) {
    SpreadsheetApp.getUi().alert(
      `❌ Sheet "${SA_SOURCE_SHEET}" not found.\n\nCreate that tab and try again.`
    );
    return;
  }

  const lastCol = src.getLastColumn();
  const lastRow = src.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("SA_DATA_MASTER has no data rows yet.");
    return;
  }

  // ── Resolve team column index (1-based) ───────────────────────────────────
  const teamColIndex = resolveColumn(SA_TEAM_COLUMN);
  if (!teamColIndex) {
    SpreadsheetApp.getUi().alert(
      `❌ Invalid SA_TEAM_COLUMN value: "${SA_TEAM_COLUMN}".\n\nSet it to a column letter (A, B, C…) or number (1, 2, 3…).`
    );
    return;
  }

  // ── Find or create tracking column ────────────────────────────────────────
  // We add a column at the far right of SA_DATA_MASTER to mark rows already copied.
  // This prevents duplicates if you run the script multiple times.
  const trackingColIndex = getOrCreateTrackingColumn(src, lastCol);

  // ── Read all data ─────────────────────────────────────────────────────────
  const allData    = src.getRange(1, 1, lastRow, trackingColIndex).getValues();
  const headerRow  = allData[0];  // Row 1 = headers (we copy these too if present)

  let pushed = 0, skipped = 0, newSheets = 0;
  const log = [];

  for (let r = 1; r < allData.length; r++) {   // r=0 is header row
    const row        = allData[r];
    const trackVal   = row[trackingColIndex - 1];

    // Skip already-pasted rows
    if (trackVal === SA_TRACKING_COL || trackVal === true || trackVal === "TRUE") {
      skipped++;
      continue;
    }

    // Read team number
    const rawTeam = row[teamColIndex - 1];
    if (rawTeam === "" || rawTeam === null || rawTeam === undefined) continue;

    const teamNum = Number(rawTeam);
    if (isNaN(teamNum) || teamNum <= 0) {
      log.push(`⚠️  Row ${r + 1}: "${rawTeam}" is not a valid team number — skipped`);
      continue;
    }

    const teamStr = teamNum.toString();

    // Get or create team sheet
    let teamSheet = ss.getSheetByName(teamStr);
    if (!teamSheet) {
      teamSheet = ss.insertSheet(teamStr);
      newSheets++;
      log.push(`✨ Created new sheet: ${teamStr}`);
    }

    // Find the true last row with content (search from bottom up)
    const pasteRow = findLastContentRow(teamSheet) + SA_ROWS_GAP + 1;

    // Build the data slice to paste (exclude tracking column)
    const dataToPaste = [row.slice(0, trackingColIndex - 1)];

    // Paste the row
    teamSheet
      .getRange(pasteRow, 1, 1, dataToPaste[0].length)
      .setValues(dataToPaste);

    // Mark row as pasted in tracking column
    src.getRange(r + 1, trackingColIndex).setValue(SA_TRACKING_COL);

    pushed++;
    log.push(`✅ Team ${teamStr} → row ${r + 1} pasted at line ${pasteRow}`);
  }

  SpreadsheetApp.flush();

  // ── Summary ───────────────────────────────────────────────────────────────
  const summary = [
    `Done!`,
    `  • ${pushed} row(s) pasted into team sheets`,
    `  • ${skipped} row(s) already pasted (skipped)`,
    `  • ${newSheets} new sheet(s) created`,
    "",
    pushed === 0 && skipped > 0
      ? "ℹ️  All rows were already pasted. Add new rows to SA_DATA_MASTER and run again."
      : "",
    log.length ? "\nDetails (last 20):\n" + log.slice(-20).join("\n") : "",
  ].join("\n").trim();

  Logger.log(summary);
  SpreadsheetApp.getUi().alert(summary);
}


// ── FORCE RE-PASTE (ignores tracking, repastes everything) ───────────────────
// Useful if you want to wipe and redo. Run this manually if needed.

function pushSADataToTeamSheets_FORCE() {
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SA_SOURCE_SHEET);
  if (!src) return;

  // Clear all tracking marks first
  const lastRow          = src.getLastRow();
  const trackingColIndex = getOrCreateTrackingColumn(src, src.getLastColumn());
  if (lastRow > 1) {
    src.getRange(2, trackingColIndex, lastRow - 1, 1).clearContent();
  }

  // Now run normally
  pushSADataToTeamSheets();
}


// ── UTILITIES ─────────────────────────────────────────────────────────────────

/**
 * Converts a column reference to a 1-based column index.
 * Accepts: "A", "B", "AA", 1, 2, 3 …
 */
function resolveColumn(ref) {
  if (!ref && ref !== 0) return null;
  if (typeof ref === "number") return ref > 0 ? ref : null;

  const s = ref.toString().trim().toUpperCase();
  if (/^\d+$/.test(s)) return parseInt(s, 10);   // numeric string like "3"

  // Column letter(s) → index
  let result = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i) - 64;            // A=1, B=2 …
    if (code < 1 || code > 26) return null;       // invalid character
    result = result * 26 + code;
  }
  return result;
}


/**
 * Searches a sheet from the BOTTOM UP to find the last row
 * that contains any non-empty cell. Returns 0 if the sheet is blank.
 *
 * This avoids the common bug where getLastRow() gets fooled by
 * formatting/blank rows with residual formatting.
 */
function findLastContentRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return 0;

  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const data    = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  for (let r = data.length - 1; r >= 0; r--) {
    const rowHasContent = data[r].some(cell => cell !== "" && cell !== null && cell !== undefined);
    if (rowHasContent) return r + 1;  // convert 0-based to 1-based
  }
  return 0;  // sheet is entirely empty
}


/**
 * Finds the tracking column in SA_DATA_MASTER (searching row 1 for SA_TRACKING_COL).
 * If it doesn't exist yet, creates it one column past the current last column.
 * Returns the 1-based column index of the tracking column.
 */
function getOrCreateTrackingColumn(sheet, currentLastCol) {
  const headerRow = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];

  // Check if tracking column already exists
  for (let i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === SA_TRACKING_COL) return i + 1;
  }

  // Doesn't exist — create it one column to the right
  const newCol = currentLastCol + 1;
  const cell   = sheet.getRange(1, newCol);
  cell.setValue(SA_TRACKING_COL);
  cell.setFontColor("#AAAAAA");
  cell.setFontStyle("italic");
  cell.setNote("Auto-managed by pushSADataToTeamSheets. Do not delete.");
  return newCol;
}


// ── MENU ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 Scouting Tools")
    .addItem("Push SA_DATA_MASTER → Team Sheets", "pushSADataToTeamSheets")
    .addItem("Force Re-paste All (ignore tracking)", "pushSADataToTeamSheets_FORCE")
    .addSeparator()
    .addItem("Push masterdata → Team Sheets", "pushMasterdataToTeamSheets")
    .addSeparator()
    .addItem("Refresh Statbotics Stats", "refreshStatboticsStats")
    .addItem("Build Graphs Sheet", "createGraphsSheet")
    .addToUi();
}