
// ============================================================
// GRAPH CONFIG
// Options: "EVENT_EPA" "PCT_ERROR" "QUAL_RANK" "WIN_RATE"
//          "DISTRICT_PTS" "FINAL_PLACE"
// ============================================================
const GRAPH_X     = "DISTRICT_PTS";
const GRAPH_Y     = "PCT_ERROR";
const GRAPH_TITLE = "DISTRICT_PTS VS PCT_ERROR";

const COLUMN_MAP = {
  "PCT_ERROR":    { col: 7,    label: "Avg % Error",    source: "sheet" },
  "QUAL_RANK":    { col: 4,    label: "Qual Rank",       source: "sheet" },
  "WIN_RATE":     { col: 5,    label: "Prelim Win Rate", source: "sheet" },
  "DISTRICT_PTS": { col: 11,   label: "District Points", source: "sheet" },
  "FINAL_PLACE":  { col: 3,    label: "Final Place",     source: "sheet" },
  "EVENT_EPA":    { col: null, label: "Event EPA",       source: "api"   }
};


// ============================================================
// FUNCTION 1 — Run ONCE to build all sheets from scratch.
// Uses fetchAll for parallel API requests — much faster.
// ============================================================
function createMasterFromTeamList_PLAINTEXT() {

  const ss        = SpreadsheetApp.getActive();
  const teamSheet = ss.getSheetByName("Teams");
  if (!teamSheet) throw new Error("No sheet named 'Teams' found.");

  const teams = teamSheet
    .getRange("A:A")
    .getValues()
    .flat()
    .filter(n => n)
    .map(n => Number(n));

  teams.forEach(team => {

    let sheet = ss.getSheetByName(team.toString());
    if (!sheet) sheet = ss.insertSheet(team.toString());

    sheet.clear();

    sheet.getRange("A1").setValue(team).setFontWeight("bold");
    sheet.getRange(1, 2, 1, 10).setValues([[
      "Rank (District/World)", "EPA", "Auto EPA", "Endgame EPA",
      "EPA Percentile", "Auto EPA Percentile", "Endgame EPA Percentile",
      "", "", "Win/Loss Ratio"
    ]]).setFontWeight("bold");

    sheet.getRange("A2").setValue("Current").setFontWeight("bold");
    sheet.getRange("A3").setValue("Last Year").setFontWeight("bold");
    sheet.getRange("A4").setValue("Past 3 Years").setFontWeight("bold");

    sheet.getRange(6, 1, 1, 11).setValues([[
      "", "district", "final place", "qual rank",
      "Prelim Record", "Elim Record",
      "avg percent error --predicted vs actual score (statbotics)",
      "Captain", "Pick 1", "Pick 2", "district points"
    ]]).setFontWeight("bold");

    writeStatRows(sheet, team);
    writeEventRows(sheet, team);
  });

  SpreadsheetApp.flush();
}


// ============================================================
// FUNCTION 2 — Refresh values only. Manual notes stay safe.
// ============================================================
function refreshStatboticsStats() {

  const ss = SpreadsheetApp.getActive();

  ss.getSheets().forEach(sheet => {
    const team = parseInt(sheet.getName());
    if (isNaN(team)) return;

    writeStatRows(sheet, team);
    writeEventRows(sheet, team);
  });

  SpreadsheetApp.flush();
}


// ============================================================
// FUNCTION 3 — Standalone graph builder.
// ============================================================
function createGraphsSheet() {

  const ss   = SpreadsheetApp.getActive();
  const xDef = COLUMN_MAP[GRAPH_X];
  const yDef = COLUMN_MAP[GRAPH_Y];

  if (!xDef || !yDef) throw new Error("Invalid GRAPH_X or GRAPH_Y.");

  const existing = ss.getSheetByName(GRAPH_TITLE);
  if (existing) ss.deleteSheet(existing);
  const graphSheet = ss.insertSheet(GRAPH_TITLE);

  graphSheet.getRange(1, 1, 1, 4)
    .setValues([["Team", "Event", xDef.label, yDef.label]])
    .setFontWeight("bold");

  let dataRow = 2;

  ss.getSheets().forEach(sheet => {
    const team = parseInt(sheet.getName());
    if (isNaN(team)) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 7) return;

    const numEventRows = lastRow - 6;
    if (numEventRows <= 0) return;

    const eventData = sheet.getRange(7, 1, numEventRows, 11).getValues();

    eventData.forEach(row => {
      const eventName = row[0];
      if (!eventName || eventName === "") return;

      let xVal = null;
      if (xDef.source === "api") {
        xVal = fetchEventEpaForRow(team, eventName);
      } else {
        xVal = resolveValue(GRAPH_X, row[xDef.col - 1]);
      }

      let yVal = null;
      if (yDef.source === "api") {
        yVal = fetchEventEpaForRow(team, eventName);
      } else {
        yVal = resolveValue(GRAPH_Y, row[yDef.col - 1]);
      }

      if (xVal === null || yVal === null) return;

      graphSheet.getRange(dataRow, 1, 1, 4).setValues([[
        team, eventName, xVal, yVal
      ]]);
      dataRow++;
    });
  });

  const totalPoints = dataRow - 2;

  if (totalPoints < 2) {
    graphSheet.getRange("A2").setValue("Not enough data — run createMasterFromTeamList_PLAINTEXT or refreshStatboticsStats first.");
    return;
  }

  graphSheet.autoResizeColumns(1, 4);

  const chart = graphSheet.newChart()
    .setChartType(Charts.ChartType.SCATTER)
    .addRange(graphSheet.getRange(1, 3, dataRow - 1, 2))
    .setOption("title", GRAPH_TITLE)
    .setOption("hAxis.title", xDef.label)
    .setOption("vAxis.title", yDef.label)
    .setOption("hAxis.minValue", 0)
    .setOption("vAxis.minValue", 0)
    .setOption("legend.position", "none")
    .setOption("pointSize", 6)
    .setOption("width", 800)
    .setOption("height", 500)
    .setPosition(2, 6, 0, 0)
    .build();

  graphSheet.insertChart(chart);
  SpreadsheetApp.flush();
  Logger.log("Graph built: " + GRAPH_TITLE + " (" + totalPoints + " points)");
}


// ============================================================
// Writes stat summary rows 2, 3, 4 using parallel fetchAll.
// ============================================================
function writeStatRows(sheet, team) {

  // Fire all 4 year requests in parallel
  const urls = [CURRENT_YEAR, CURRENT_YEAR-1, EVENT_YEAR, EVENT_YEAR-1, EVENT_YEAR-2].map(year =>
    ({ url: "https://api.statbotics.io/v3/team_year/" + team + "/" + year,
       muteHttpExceptions: true })
  );

  const responses = UrlFetchApp.fetchAll(urls);
  const dataArr   = responses.map(r =>
    r.getResponseCode() === 200 ? JSON.parse(r.getContentText()) : null
  );

  // Row 2: try current year, fall back to last year
  const currentData = dataArr[0] || dataArr[1];
  const lastData    = dataArr[2];
  const y2Data      = dataArr[3];
  const y3Data      = dataArr[4];

  function extract(d) {
    if (!d) return ["", "", "", "", "", "", "", "", ""];
    const distRank  = d.epa?.ranks?.district?.rank  ?? "";
    const worldRank = d.epa?.ranks?.total?.rank     ?? "";
    const epa       = d.epa?.breakdown?.total_points    ?? "";
    const autoEpa   = d.epa?.breakdown?.auto_points     ?? "";
    const endEpa    = d.epa?.breakdown?.endgame_points  ?? "";
    const pct       = d.epa?.ranks?.total?.percentile   ?? "";
    // Use W/L format with slash — avoids Google Sheets date auto-parsing
    const wl        = d.record ? d.record.wins + "/" + d.record.losses : "";
    return [distRank + " / " + worldRank, epa, autoEpa, endEpa, pct, pct, pct, "", wl];
  }

  const valid = [lastData, y2Data, y3Data].filter(d => d);

  function avgField(fn) {
    if (!valid.length) return "";
    return valid.reduce((s, d) => s + fn(d), 0) / valid.length;
  }

  const aggRow = valid.length ? [
    "",
    avgField(d => d.epa?.breakdown?.total_points   ?? 0),
    avgField(d => d.epa?.breakdown?.auto_points    ?? 0),
    avgField(d => d.epa?.breakdown?.endgame_points ?? 0),
    "", "", "", "", ""
  ] : ["", "", "", "", "", "", "", "", ""];

  sheet.getRange(2, 2, 1, 9).setValues([extract(currentData)]);
  sheet.getRange(3, 2, 1, 9).setValues([extract(lastData)]);
  sheet.getRange(4, 2, 1, 9).setValues([aggRow]);
}


// ============================================================
// Writes event rows starting at row 7 using parallel fetchAll.
//
// RECORD FORMAT FIX: Google Sheets auto-parses "7-4-0" as a
// date. Using "W-L-T" suffix format prevents this entirely.
// e.g. "7W 4L 0T" cannot be misread as a date.
// ============================================================
function writeEventRows(sheet, team) {

  // Step 1: get TBA event list
  const tbaRes = UrlFetchApp.fetch(
    "https://www.thebluealliance.com/api/v3/team/frc" + team + "/events/" + EVENT_YEAR,
    { headers: { "X-TBA-Auth-Key": TBA_KEY }, muteHttpExceptions: true }
  );

  if (tbaRes.getResponseCode() !== 200) return;

  const events = JSON.parse(tbaRes.getContentText());
  if (!events || events.length === 0) return;

  events.sort((a, b) => new Date(a.start_date) - new Date(b.start_date));

  // Step 2: fire all per-event Statbotics + TBA alliance + match requests in parallel
  const sbUrls       = events.map(e => ({
    url: "https://api.statbotics.io/v3/team_event/" + team + "/" + e.key,
    muteHttpExceptions: true
  }));

  const matchUrls    = events.map(e => ({
    url: "https://api.statbotics.io/v3/matches?event=" + e.key + "&limit=100",
    muteHttpExceptions: true
  }));

  const allianceUrls = events.map(e => ({
    url: "https://www.thebluealliance.com/api/v3/event/" + e.key + "/alliances",
    headers: { "X-TBA-Auth-Key": TBA_KEY },
    muteHttpExceptions: true
  }));

  const sbResponses       = UrlFetchApp.fetchAll(sbUrls);
  const matchResponses    = UrlFetchApp.fetchAll(matchUrls);
  const allianceResponses = UrlFetchApp.fetchAll(allianceUrls);

  const rows     = [];
  const boldCols = [];

  events.forEach((event, idx) => {

    const eventName    = event.name || "";
    const districtCode = event.district ? event.district.abbreviation.toUpperCase() : "";

    let finalPlace   = "";
    let qualRank     = "";
    let prelimRecord = "";
    let elimRecord   = "";
    let captain      = "";
    let pick1        = "";
    let pick2        = "";
    let districtPts  = "";

    // --- Statbotics team_event ---
    try {
      const sbRes = sbResponses[idx];
      if (sbRes.getResponseCode() === 200) {
        const sb = JSON.parse(sbRes.getContentText());

        if (sb.record?.qual) {
          qualRank = sb.record.qual.rank !== undefined ? sb.record.qual.rank : "";
          // FIX: use "W-L-T" suffix so Sheets can't misread as a date
          prelimRecord = sb.record.qual.wins + "W-" + sb.record.qual.losses + "L-" + sb.record.qual.ties + "T";
        }

        if (sb.record?.elim) {
          elimRecord = sb.record.elim.wins + "W-" + sb.record.elim.losses + "L-" + sb.record.elim.ties + "T";
        }

        if (sb.district_points !== undefined && sb.district_points !== null) {
          districtPts = sb.district_points;
        }
      }
    } catch (e) {
      Logger.log("Statbotics error - team " + team + " event " + event.key + ": " + e);
    }

    // --- Avg percent error from match predictions ---
    let avgPctError = "";
    try {
      const mRes = matchResponses[idx];
      if (mRes.getResponseCode() === 200) {
        const matches = JSON.parse(mRes.getContentText());
        const teamNum = Number(team);
        const errors  = [];

        matches.forEach(match => {
          if (match.elim !== false) return;
          if (!match.result || !match.pred) return;

          const redTeams  = match.alliances?.red?.team_keys  || [];
          const blueTeams = match.alliances?.blue?.team_keys || [];

          let predicted = null;
          let actual    = null;

          if (redTeams.includes(teamNum)) {
            predicted = match.pred.red_score;
            actual    = match.result.red_score;
          } else if (blueTeams.includes(teamNum)) {
            predicted = match.pred.blue_score;
            actual    = match.result.blue_score;
          }

          if (predicted != null && actual != null && actual > 0) {
            errors.push(Math.abs(predicted - actual) / actual * 100);
          }
        });

        if (errors.length) {
          avgPctError = Math.round(errors.reduce((s, e) => s + e, 0) / errors.length * 100) / 100;
        }
      }
    } catch (e) {
      Logger.log("Match error - team " + team + " event " + event.key + ": " + e);
    }

    // --- TBA alliances ---
    let boldCol = null;
    try {
      const aRes = allianceResponses[idx];
      if (aRes.getResponseCode() === 200) {
        const alliances = JSON.parse(aRes.getContentText());
        const teamKey   = "frc" + team;

        alliances.forEach(alliance => {
          const picks = alliance.picks || [];
          if (!picks.includes(teamKey)) return;

          captain = picks[0] ? picks[0].replace("frc", "") : "";
          pick1   = picks[1] ? picks[1].replace("frc", "") : "";
          pick2   = picks[2] ? picks[2].replace("frc", "") : "";

          const idx2 = picks.indexOf(teamKey);
          if      (idx2 === 0) boldCol = 8;
          else if (idx2 === 1) boldCol = 9;
          else if (idx2 === 2) boldCol = 10;

          if (alliance.status?.status === "won")   finalPlace = "Winner";
          else if (alliance.status?.level)          finalPlace = alliance.status.level;
        });
      }
    } catch (e) {
      Logger.log("TBA alliance error - team " + team + " event " + event.key + ": " + e);
    }

    boldCols.push(boldCol);

    rows.push([
      eventName, districtCode, finalPlace, qualRank,
      prelimRecord, elimRecord, avgPctError,
      captain, pick1, pick2, districtPts
    ]);
  });

  if (rows.length > 0) {
    sheet.getRange(7, 1, rows.length + 5, 11).clearContent();
    sheet.getRange(7, 1, rows.length, 11).setValues(rows);

    boldCols.forEach((col, i) => {
      if (col !== null) sheet.getRange(7 + i, col).setFontWeight("bold");
    });
  }
}


// ============================================================
// fetchAll-based event EPA lookup for graph builder
// ============================================================
const eventKeyCache = {};

function fetchEventEpaForRow(team, eventName) {
  const cacheKey = team + "_" + EVENT_YEAR;

  if (!eventKeyCache[cacheKey]) {
    eventKeyCache[cacheKey] = {};
    try {
      const res = UrlFetchApp.fetch(
        "https://www.thebluealliance.com/api/v3/team/frc" + team + "/events/" + EVENT_YEAR,
        { headers: { "X-TBA-Auth-Key": TBA_KEY }, muteHttpExceptions: true }
      );
      if (res.getResponseCode() === 200) {
        JSON.parse(res.getContentText()).forEach(e => {
          eventKeyCache[cacheKey][e.name] = e.key;
        });
      }
    } catch (e) {
      Logger.log("TBA event list error for team " + team + ": " + e);
    }
  }

  const eventKey = eventKeyCache[cacheKey][eventName];
  if (!eventKey) return null;

  try {
    const res = UrlFetchApp.fetch(
      "https://api.statbotics.io/v3/team_event/" + team + "/" + eventKey,
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() === 200) {
      const sb = JSON.parse(res.getContentText());
      if (sb.epa?.breakdown?.total_points !== undefined) {
        return sb.epa.breakdown.total_points;
      }
    }
  } catch (e) {
    Logger.log("Statbotics EPA error for " + team + "/" + eventKey + ": " + e);
  }

  return null;
}


function resolveValue(key, raw) {
  if (raw === "" || raw === null || raw === undefined) return null;

  if (key === "WIN_RATE") {
    // Handles both old "W-L-T" format and new "7W-4L-0T" format
    const nums = String(raw).match(/\d+/g);
    if (!nums || nums.length < 2) return null;
    const wins  = Number(nums[0]);
    const total = Number(nums[0]) + Number(nums[1]) + (Number(nums[2]) || 0);
    if (isNaN(wins) || isNaN(total) || total === 0) return null;
    return Math.round(wins / total * 1000) / 1000;
  }

  const num = Number(raw);
  if (isNaN(num)) return null;
  return num;
}