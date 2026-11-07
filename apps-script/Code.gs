/**
 * ============================================================================
 * CODEFORCES DAILY CHALLENGE TRACKER — FUNCTION DIRECTORY & ARCHITECTURE GUIDE
 * ============================================================================
 *
 * [TRIGGER & ENTRY POINT FUNCTIONS]
 * 1. syncAllExistingRowsNow()
 *    - Purpose: Loops through every participant on the "Dashboard" tab and
 *      refreshes their live Codeforces rating, yesterday's valid solves, and
 *      today's valid solves.
 *    - Trigger: Bound to the "On open" spreadsheet trigger (and can be run
 *      manually from the Apps Script toolbar).
 *
 * 2. onSheetEdit(e)
 *    - Purpose: Detects when a user types or pastes a Codeforces handle/URL
 *      into Column A of the "Dashboard" tab and immediately populates Columns
 *      B, C, D, and E for that specific row.
 *    - Trigger: Bound to the "On edit" installable spreadsheet trigger.
 *
 * 3. doGet(e)
 *    - Purpose: Handles all incoming HTTP GET requests to the Web App URL.
 *      • Default (?format=html): Renders the HackerRank/CodeChef-styled HTML
 *        leaderboard table in the browser.
 *      • Raw HTML (?format=rawhtml): Returns raw HTML without Google's top
 *        warning banner (used by custom wrappers/Cloudflare Worker).
 *      • Discord Bot API (?format=text&cmd=status|penalty|adduser): Returns
 *        clean Discord-formatted text responses for slash commands.
 *    - Trigger: Executed automatically whenever the Web App URL is visited.
 *
 * 4. runDailyCodeforcesCheck()
 *    - Purpose: Performs the daily midnight settlement. Evaluates everyone's
 *      valid solves from the completed 24-hour day, calculates missed-problem
 *      penalties (10 BDT per missed problem), updates "Dashboard", appends an
 *      audit log row to "History", and triggers a Discord alert if any penalty
 *      was charged.
 *    - Trigger: Bound to the "Time-driven -> Day timer -> Midnight to 1am" trigger.
 *
 * ----------------------------------------------------------------------------
 * [CORE HELPER & API FUNCTIONS]
 * 5. populateUserRow(dashSheet, row, handle)
 *    - Purpose: Core row-updater called by syncAllExistingRowsNow(), onSheetEdit(),
 *      and doGet(). Fetches a single user's rating and solve counts, calculates
 *      their dynamic minimum problem rating floor, and writes Columns B–E.
 *
 * 6. getUserRatings(handles)
 *    - Purpose: Calls the Codeforces `user.info` API to batch-fetch current
 *      ratings for a list of handles. Includes a 3-attempt auto-retry loop
 *      (with a 3-second delay) to survive temporary API lag.
 *
 * 7. getBothDaysSolvedCounts(handle, yestStartSec, todayStartSec, nowSec, minProblemRating)
 *    - Purpose: Calls the Codeforces `user.status` API (last 1,000 submissions)
 *      with a 3-attempt auto-retry loop. Filters for unique "OK" verdicts that
 *      were never solved before yesterday/today and meet `minProblemRating`.
 *
 * 8. sendDiscordNotification(message)
 *    - Purpose: Sends a POST payload to `WEBHOOK_URL` when a penalty is charged
 *      at midnight. Safely exits if the webhook placeholder hasn't been replaced.
 *
 * ----------------------------------------------------------------------------
 * [UTILITY & STYLING FUNCTIONS]
 * 9. extractHandle(rawInput)
 *    - Purpose: Sanitizes Column A inputs so users can enter either a raw handle
 *      ("tourist") or a full profile link ("https://codeforces.com/profile/tourist").
 *
 * 10. getCfColor(rating)
 *     - Purpose: Returns the official Codeforces rank hex color (Gray, Green,
 *       Cyan, Blue, Violet, Orange, Red) for rendering handles in the Web App.
 * ============================================================================
 */


const WEBHOOK_URL = "YOUR_DISCORD_WEBHOOK_URL"; // Paste Discord Webhook URL here when ready
const DAILY_GOAL = 5;
const PENALTY_PER_MISS = 10; // 10 BDT per missed problem
const COUNT_UNRATED_CONTEST_PROBLEMS = true;

// Helper: Extracts handle whether user enters "tourist" or "https://codeforces.com/profile/tourist"
function extractHandle(rawInput) {
  if (!rawInput) return "";
  return String(rawInput).trim().replace(/\/+$/, "").split("/").pop().trim();
}

// Helper: Official Codeforces Rating Color
function getCfColor(rating) {
  if (rating >= 2400) return "#ff0000"; // Grandmaster (Red)
  if (rating >= 2100) return "#ff8c00"; // Master (Orange)
  if (rating >= 1900) return "#aa00aa"; // Candidate Master (Violet)
  if (rating >= 1600) return "#0000ff"; // Expert (Blue)
  if (rating >= 1400) return "#03a89e"; // Specialist (Cyan)
  if (rating >= 1200) return "#008000"; // Pupil (Green)
  return "#808080";                     // Newbie (Gray)
}

// 1. RUN THIS MANUALLY IN APPS SCRIPT TO FILL ALL CURRENT ROWS RIGHT NOW
function syncAllExistingRowsNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("Dashboard");
  const data = dashSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const handle = extractHandle(data[i][0]);
    if (handle) {
      populateUserRow(dashSheet, i + 1, handle);
      Utilities.sleep(1500);
    }
  }
}

// 2. AUTO-TRIGGER WHEN YOU TYPE A HANDLE IN COLUMN A
function onSheetEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Dashboard") return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (col !== 1 || row <= 1) return;

  const handle = extractHandle(e.range.getValue());
  if (!handle) return;

  populateUserRow(sheet, row, handle);
}

// Core helper: Fills Col B (Rating), Col C (Solved Yesterday), Col D (Solved Today), Col E (Penalty)
function populateUserRow(dashSheet, row, handle) {
  const ratings = getUserRatings([handle]);
  const liveRating = ratings[handle.toLowerCase()];

  if (liveRating === undefined) {
    dashSheet.getRange(row, 2).setValue("❌ Invalid Handle");
    return;
  }

  const existingRating = Number(dashSheet.getRange(row, 2).getValue());
  const evalRating = existingRating > 0 ? existingRating : liveRating;
  const minProblemRating = Math.max(800, evalRating - 100 - (evalRating % 100));

  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const yesterdayMidnight = new Date(todayMidnight.getTime() - 24 * 60 * 60 * 1000);

  const yestStartSec = Math.floor(yesterdayMidnight.getTime() / 1000);
  const todayStartSec = Math.floor(todayMidnight.getTime() / 1000);
  const nowSec = Math.floor(now.getTime() / 1000);

  const stats = getBothDaysSolvedCounts(handle, yestStartSec, todayStartSec, nowSec, minProblemRating);
  const solvedYesterday = stats ? stats.solvedYesterday : 0;
  const solvedToday = stats ? stats.solvedToday : 0;

  const currentPenaltyCell = dashSheet.getRange(row, 5).getValue();
  const penaltyToSet = (currentPenaltyCell === "" || currentPenaltyCell === null || isNaN(Number(currentPenaltyCell)))
    ? 0
    : Number(currentPenaltyCell);

  dashSheet.getRange(row, 2, 1, 4).setValues([[evalRating, solvedYesterday, solvedToday, penaltyToSet]]);
  SpreadsheetApp.flush();
}

// 3. DISCORD SLASH COMMANDS & CLASSIC CP LEADERBOARD WEB APP
function doGet(e) {
  const cmd = (e && e.parameter && e.parameter.cmd) ? e.parameter.cmd : "status";
  const format = (e && e.parameter && e.parameter.format) ? e.parameter.format : "html";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("Dashboard");
  const data = dashSheet.getDataRange().getValues();

  // Command: /adduser
  if (cmd === "adduser") {
    const rawUser = (e.parameter.user || "").trim();
    const handle = extractHandle(rawUser);
    if (!handle) return ContentService.createTextOutput("❌ Please provide a valid Codeforces handle or profile URL.");

    for (let i = 1; i < data.length; i++) {
      if (extractHandle(data[i][0]).toLowerCase() === handle.toLowerCase()) {
        return ContentService.createTextOutput(`⚠️ **${handle}** is already on the Dashboard!`);
      }
    }

    const ratings = getUserRatings([handle]);
    const actualRating = ratings[handle.toLowerCase()];
    if (actualRating === undefined) {
      return ContentService.createTextOutput(`❌ Could not find Codeforces user **${handle}**.`);
    }

    const newRowIndex = dashSheet.getLastRow() + 1;
    dashSheet.getRange(newRowIndex, 1).setValue(handle);
    populateUserRow(dashSheet, newRowIndex, handle);

    const minRating = Math.max(800, actualRating - 100 - (actualRating % 100));
    return ContentService.createTextOutput(
      `✅ Added **${handle}** to the Challenge!\n• **Rating:** ${actualRating} *(Min Problem Rating: ${minRating}+)*`
    );
  }

  // Command: /penalty (Text mode for Discord)
  if (cmd === "penalty" && format === "text") {
    if (data.length <= 1) return ContentService.createTextOutput("No participants added yet.");
    let lines = ["**Current Penalty Standings**\n"];
    for (let i = 1; i < data.length; i++) {
      const handle = extractHandle(data[i][0]);
      if (!handle) continue;
      const solvedToday = Number(data[i][3]) || 0;
      const penalty = Number(data[i][4]) || 0;
      lines.push(`• **${handle}**: **${penalty} BDT** *(Solved Today: ${solvedToday}/${DAILY_GOAL})*`);
    }
    return ContentService.createTextOutput(lines.join("\n"));
  }

  if (data.length <= 1) {
    return ContentService.createTextOutput("No participants added yet.");
  }

  let rowsData = [];
  let textLines = [`**Today's Live Challenge Status (Since 12:00 AM)**\n`];

  for (let i = 1; i < data.length; i++) {
    const handle = extractHandle(data[i][0]);
    if (!handle) continue;

    populateUserRow(dashSheet, i + 1, handle);
    const updatedRow = dashSheet.getRange(i + 1, 2, 1, 4).getValues()[0];
    const evalRating = Number(updatedRow[0]) || 800;
    const solvedYesterday = Number(updatedRow[1]) || 0;
    const solvedToday = Number(updatedRow[2]) || 0;
    const currentPenalty = Number(updatedRow[3]) || 0;
    const minProblemRating = Math.max(800, evalRating - 100 - (evalRating % 100));
    const remaining = Math.max(0, DAILY_GOAL - solvedToday);

    rowsData.push({
      rank: rowsData.length + 1,
      handle,
      evalRating,
      cfColor: getCfColor(evalRating),
      minProblemRating,
      solvedYesterday,
      solvedToday,
      currentPenalty,
      remaining
    });

    const statusBadge = remaining === 0
      ? "✅ **SAFE**"
      : `⚠️ **Needs ${remaining} more** *(+${remaining * PENALTY_PER_MISS} BDT at midnight)*`;

    textLines.push(
      `• **${handle}** *(Min: ${minProblemRating}+)*: **${solvedToday}/${DAILY_GOAL}** — ${statusBadge} | *Owed: ${currentPenalty} BDT*`
    );
  }

  if (format === "text") {
    return ContentService.createTextOutput(textLines.join("\n"));
  }

  let html = `
  <!DOCTYPE html>
  <html>
  <head>
    <meta charset="utf-8">
    <title>Daily Challenge Standings</title>
    <style>
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Oxygen, Ubuntu, sans-serif;
        background-color: #f3f7f7;
        color: #1e272e;
      }
      .topbar {
        background-color: #1d2432;
        color: #ffffff;
        padding: 14px 32px;
        border-bottom: 3px solid #2ec866;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      .topbar-title {
        font-size: 16px;
        font-weight: 700;
        letter-spacing: 0.3px;
        display: flex;
        align-items: center;
        gap: 10px;
      }
      .brand-dot {
        display: inline-block;
        width: 10px;
        height: 10px;
        background-color: #2ec866;
      }
      .topbar-link {
        color: #d1d8e0;
        text-decoration: none;
        font-size: 13px;
        border: 1px solid #485460;
        padding: 6px 12px;
        border-radius: 3px;
        background: #273042;
      }
      .topbar-link:hover {
        background: #323e54;
        color: #fff;
      }
      .container {
        max-width: 980px;
        margin: 32px auto;
        padding: 0 20px;
      }
      .section-header {
        display: flex;
        justify-content: space-between;
        align-items: baseline;
        margin-bottom: 12px;
      }
      .section-header h1 {
        font-size: 20px;
        font-weight: 600;
        color: #39424e;
        margin: 0;
      }
      .section-header span {
        font-size: 13px;
        color: #738f93;
      }
      .table-wrap {
        background: #ffffff;
        border: 1px solid #d8e2e7;
        border-radius: 4px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.04);
        overflow-x: auto;
      }
      table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
      }
      th {
        background-color: #f8fafa;
        color: #576871;
        text-align: left;
        padding: 12px 16px;
        font-weight: 600;
        font-size: 12px;
        text-transform: uppercase;
        border-bottom: 2px solid #e1e8ed;
        white-space: nowrap;
      }
      td {
        padding: 14px 16px;
        border-bottom: 1px solid #eaeff2;
        vertical-align: middle;
        white-space: nowrap;
      }
      tr:nth-child(even) td {
        background-color: #fcfdfd;
      }
      tr:hover td {
        background-color: #f4f9f7;
      }
      .mono {
        font-family: "JetBrains Mono", Consolas, Monaco, "Courier New", monospace;
        font-size: 13.5px;
      }
      .cf-handle {
        font-weight: 700;
        text-decoration: none;
      }
      .cf-handle:hover {
        text-decoration: underline;
      }
      .sub-rating {
        color: #738f93;
        font-size: 12px;
        margin-left: 4px;
      }
      .status-ac {
        display: inline-block;
        white-space: nowrap;
        background-color: #e6f9ed;
        color: #1b8a44;
        border: 1px solid #b7ebc9;
        padding: 3px 10px;
        border-radius: 3px;
        font-size: 12px;
        font-weight: 600;
      }
      .status-pending {
        display: inline-block;
        white-space: nowrap;
        background-color: #fff8e6;
        color: #b76e00;
        border: 1px solid #ffe299;
        padding: 3px 10px;
        border-radius: 3px;
        font-size: 12px;
        font-weight: 600;
      }
      .penalty-zero {
        color: #738f93;
      }
      .penalty-owed {
        color: #d93025;
        font-weight: 700;
      }
    </style>
  </head>
  <body>
    <div class="topbar">
      <div class="topbar-title">
        <span class="brand-dot"></span>
        CODEFORCES DAILY PRACTICE TRACKER
      </div>
      <a class="topbar-link" href="${ss.getUrl()}" target="_blank">View Sheet Ledger</a>
    </div>

    <div class="container">
      <div class="section-header">
        <h1>Standings</h1>
        <span>Target: ${DAILY_GOAL} problems/day &bull; Penalty: ${PENALTY_PER_MISS} BDT/missed</span>
      </div>

      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th style="width: 50px;">#</th>
              <th>Coder</th>
              <th>Rating (Min Req)</th>
              <th>Yesterday</th>
              <th>Today</th>
              <th>Status</th>
              <th>Penalty Balance</th>
            </tr>
          </thead>
          <tbody>`;

  for (const r of rowsData) {
    const statusHtml = r.remaining === 0
      ? `<span class="status-ac">COMPLETED (5/5)</span>`
      : `<span class="status-pending">${r.remaining} remaining (+${r.remaining * PENALTY_PER_MISS} BDT if missed)</span>`;

    const penaltyClass = r.currentPenalty > 0 ? "penalty-owed" : "penalty-zero";

    html += `
            <tr>
              <td class="mono" style="color:#738f93;">${r.rank}</td>
              <td>
                <a class="cf-handle" style="color: ${r.cfColor};" href="https://codeforces.com/profile/${r.handle}" target="_blank">${r.handle}</a>
              </td>
              <td class="mono">
                <strong style="color: ${r.cfColor};">${r.evalRating}</strong>
                <span class="sub-rating">[min ${r.minProblemRating}]</span>
              </td>
              <td class="mono">${r.solvedYesterday} / ${DAILY_GOAL}</td>
              <td class="mono"><strong>${r.solvedToday} / ${DAILY_GOAL}</strong></td>
              <td>${statusHtml}</td>
              <td class="mono ${penaltyClass}">${r.currentPenalty} BDT</td>
            </tr>`;
  }

  html += `
          </tbody>
        </table>
      </div>
    </div>
  </body>
  </html>`;

  if (format === "rawhtml") {
    return ContentService.createTextOutput(html).setMimeType(ContentService.MimeType.TEXT);
  }

  return HtmlService.createHtmlOutput(html)
    .setTitle("Daily Challenge Standings")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 4. MIDNIGHT AUTOMATIC SETTLEMENT (Runs once between 12:00 AM - 1:00 AM daily)
function runDailyCodeforcesCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("Dashboard");
  const histSheet = ss.getSheetByName("History");

  const data = dashSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  const yesterdayMidnight = new Date(todayMidnight.getTime() - 24 * 60 * 60 * 1000);

  const yestStartSec = Math.floor(yesterdayMidnight.getTime() / 1000);
  const todayStartSec = Math.floor(todayMidnight.getTime() / 1000);
  const nowSec = Math.floor(now.getTime() / 1000);
  const dateLabel = Utilities.formatDate(yesterdayMidnight, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const handles = [];
  for (let i = 1; i < data.length; i++) {
    const h = extractHandle(data[i][0]);
    if (h) handles.push(h);
  }

  const latestRatings = getUserRatings(handles);
  Utilities.sleep(2100);

  let anyPenaltyCharged = false;
  let penaltyLines = [`🚨 **Codeforces Penalty Alert — ${dateLabel}**\n`];

  for (let i = 1; i < data.length; i++) {
    const handle = extractHandle(data[i][0]);
    if (!handle) continue;

    const latestRating = latestRatings[handle.toLowerCase()] || 800;
    const evalRating = Number(data[i][1]) || latestRating;
    const minProblemRating = Math.max(800, evalRating - 100 - (evalRating % 100));
    const prevPenalty = Number(data[i][4]) || 0;

    const stats = getBothDaysSolvedCounts(handle, yestStartSec, todayStartSec, nowSec, minProblemRating);
    if (!stats) {
      Logger.log(`Codeforces API unreachable after 3 attempts for ${handle}.`);
      continue;
    }

    const solvedYesterday = stats.solvedYesterday;
    const solvedToday = stats.solvedToday;

    const missed = Math.max(0, DAILY_GOAL - solvedYesterday);
    const penaltyAdded = missed * PENALTY_PER_MISS;
    const newPenalty = prevPenalty + penaltyAdded;

    // Update Dashboard: Col B (new rating for today), Col C (Solved Yesterday), Col D (Solved Today), Col E (New Penalty)
    dashSheet.getRange(i + 1, 2, 1, 4).setValues([[latestRating, solvedYesterday, solvedToday, newPenalty]]);

    if (histSheet) {
      histSheet.appendRow([dateLabel, handle, evalRating, minProblemRating, solvedYesterday, penaltyAdded, newPenalty]);
    }

    if (penaltyAdded > 0) {
      anyPenaltyCharged = true;
      penaltyLines.push(
        `💸 **${handle}** *(Min: ${minProblemRating}+)*: Solved **${solvedYesterday}/${DAILY_GOAL}** — **+${penaltyAdded} BDT penalty charged!** *(Total Owed: ${newPenalty} BDT)*`
      );
    } else {
      penaltyLines.push(
        `✅ **${handle}** *(Min: ${minProblemRating}+)*: Solved **${solvedYesterday}/${DAILY_GOAL}** — Safe *(Total Owed: ${newPenalty} BDT)*`
      );
    }

    Utilities.sleep(2100);
  }

  SpreadsheetApp.flush();

  if (anyPenaltyCharged) {
    penaltyLines.push(`\n🔗 [View Full Ledger in Google Sheets](${ss.getUrl()})`);
    sendDiscordNotification(penaltyLines.join("\n"));
  }
}

// Fetches user ratings with 3-attempt auto-retry
function getUserRatings(handles) {
  const ratingsMap = {};
  if (handles.length === 0) return ratingsMap;
  const url = `https://codeforces.com/api/user.info?handles=${encodeURIComponent(handles.join(";"))}`;

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        const json = JSON.parse(response.getContentText());
        if (json.status === "OK") {
          for (const user of json.result) {
            ratingsMap[user.handle.toLowerCase()] = user.rating || 800;
          }
          return ratingsMap;
        }
      }
    } catch (e) {
      Logger.log(`user.info attempt ${attempt} failed: ${e}`);
    }
    Utilities.sleep(3000); // Wait 3 seconds before retrying
  }
  return ratingsMap;
}

// Counts both Yesterday's valid solves AND Today's valid solves with 3-attempt auto-retry
function getBothDaysSolvedCounts(handle, yestStartSec, todayStartSec, nowSec, minProblemRating) {
  const url = `https://codeforces.com/api/user.status?handle=${encodeURIComponent(handle)}&from=1&count=1000`;

  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        const json = JSON.parse(response.getContentText());
        if (json.status === "OK") {
          const solvedBeforeYesterday = new Set();
          const solvedYesterday = new Set();
          const solvedToday = new Set();
          const submissions = json.result;

          for (let i = submissions.length - 1; i >= 0; i--) {
            const sub = submissions[i];
            if (sub.verdict !== "OK" || !sub.problem) continue;

            const problemId = `${sub.problem.contestId || ""}-${sub.problem.index || ""}-${sub.problem.name}`;
            const time = sub.creationTimeSeconds;
            const probRating = sub.problem.rating;

            const meetsRating = probRating !== undefined
              ? probRating >= minProblemRating
              : COUNT_UNRATED_CONTEST_PROBLEMS;

            if (time < yestStartSec) {
              solvedBeforeYesterday.add(problemId);
            } else if (time >= yestStartSec && time < todayStartSec) {
              if (!solvedBeforeYesterday.has(problemId) && meetsRating) {
                solvedYesterday.add(problemId);
              }
              solvedBeforeYesterday.add(problemId);
            } else if (time >= todayStartSec && time <= nowSec) {
              if (!solvedBeforeYesterday.has(problemId) && meetsRating) {
                solvedToday.add(problemId);
              }
            }
          }
          return {
            solvedYesterday: solvedYesterday.size,
            solvedToday: solvedToday.size
          };
        }
      }
    } catch (e) {
      Logger.log(`user.status attempt ${attempt} failed for ${handle}: ${e}`);
    }
    Utilities.sleep(3000); // Wait 3 seconds before retrying
  }
  return null;
}

function sendDiscordNotification(message) {
  if (!WEBHOOK_URL || !WEBHOOK_URL.startsWith("https://discord.com/api/webhooks/")) return;
  try {
    UrlFetchApp.fetch(WEBHOOK_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ content: message }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log(`Webhook error: ${e}`);
  }
}
