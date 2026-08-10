// @ts-ignore
/**
 * Life Management System — Full Version
 * Bugs fixed, stubs implemented, new features added.
 * Includes Hierarchical Goals, Pomodoro Timer with Time Tracking, and Unique Matrices.
 *
 * Required scopes (auto-detected by Apps Script):
 *   - https://www.googleapis.com/auth/spreadsheets
 *   - https://www.googleapis.com/auth/gmail.send
 *   - https://www.googleapis.com/auth/script.scriptapp
 *   - https://www.googleapis.com/auth/userinfo.email
 */

// ======================== CONSTANTS ========================

const DAILY_LOG_SHEET_NAME              = "Daily Log";
const MASTER_PROJECT_TRACKER_SHEET_NAME = "Master Project Tracker";
const DASHBOARD_SHEET_NAME              = "Dashboard";
const LISTS_SHEET_NAME                  = "Lists";
const OVERDUE_LOG_SHEET_NAME            = "Overdue Log";
const GOALS_SHEET_NAME                  = "Goals & Habits";
const CHART_DATA_SHEET_NAME             = "Chart Data";
const BRAIN_DUMP_SHEET_NAME             = "Brain Dump";
const ARCHIVE_SHEET_NAME                = "Archive";

// Lock Properties Key
const LOCK_PROPERTY = PropertiesService.getScriptProperties();
const LOCK_KEY = "SCRIPT_LOCK_KEY_LMS";
const LOCK_TIMEOUT_MS = 30000; // 30 seconds

// Colors (status-based)
const COLOR_TODO       = "#cfe2f3"; // soft blue
const COLOR_INPROGRESS = "#fce5cd"; // soft orange
const COLOR_DONE       = "#d9ead3"; // soft green
const COLOR_BLOCKED    = "#f4cccc"; // soft red
const COLOR_WAITING    = "#efefef"; // grey
const COLOR_OVERDUE    = "#ea9999"; // deep red for overdue rows

// Motivational quotes pool
const QUOTES = [
  "The secret of getting ahead is getting started. – Mark Twain",
  "It always seems impossible until it's done. – Nelson Mandela",
  "Focus on being productive instead of busy. – Tim Ferriss",
  "Your future is created by what you do today, not tomorrow.",
  "Small steps every day lead to big results.",
  "Done is better than perfect. – Sheryl Sandberg",
  "Work hard in silence; let success make the noise.",
  "Discipline is choosing between what you want now and what you want most.",
  "The key is not to prioritize what's on your schedule, but to schedule your priorities. – Stephen Covey",
  "Energy and persistence conquer all things. – Benjamin Franklin"
];

// GTD / Eisenhower coaching tips
const FRAMEWORK_TIPS = [
  "GTD: Capture everything in your Brain Dump sheet, then clarify each item into a project or next action.",
  "Eisenhower: Tasks in the 'Urgent & Important' quadrant should be done immediately — not delegated.",
  "GTD: A 'next action' should be a physical, visible activity. Not 'work on project' but 'email John re: proposal'.",
  "Eisenhower: 'Important but Not Urgent' tasks (Q2) are where real growth happens. Schedule them intentionally.",
  "GTD: Weekly review is the heartbeat of the system — review every open loop once a week.",
  "Time blocking: Assign your P1 tasks to your highest-energy hours of the day.",
  "The 2-minute rule: If a task takes less than 2 minutes, do it now instead of logging it.",
  "Batch similar tasks together to reduce context-switching overhead."
];

// Coaching Advice Text
const COACHING_ADVICE_TEXT = `Yes — I remember. You opened by asking me to tell you honestly about your strengths, flaws, and struggles; you’ve since shared a lot: depression, ADHD-like problems with follow-through, low confidence, dozens of ideas, and the pressure of needing income and to launch projects. I’m not going to sugarcoat anything or flatter you — I’ll be direct, practical, and kind.

Your biggest strength is your endless creative spark. Your mind generates ideas faster than anyone can execute them. But that spark is also your trap: without structure, you drift from project to project, building half-finished monuments.

To overcome this, use this Life Management System religiously:
1. Capture everything in the Brain Dump immediately. Clear your mind.
2. Structure your day using the "Daily Log". Commit to ONLY 3 'Do Now' tasks per day.
3. Use the Pomodoro Timer to keep your focus locked on one single physical action at a time.
4. Keep your Goals high and clear, and let the hierarchy show your slow, steady progress.

Remember, momentum beats perfection every single time. Finish today's work today.`;

// ======================== LOCK MECHANISM ========================

function isScriptRunning() {
  const lockValue = LOCK_PROPERTY.getProperty(LOCK_KEY);
  if (!lockValue) return false;
  
  const lockTime = parseInt(lockValue, 10);
  const now = new Date().getTime();
  
  if (now - lockTime > LOCK_TIMEOUT_MS) {
    Logger.log("[LOCK] Stale lock detected. Releasing.");
    LOCK_PROPERTY.deleteProperty(LOCK_KEY);
    return false;
  }
  return true;
}

function acquireLock() {
  const now = new Date().getTime();
  if (isScriptRunning()) {
    Logger.log("[LOCK] Script is already running. Aborting.");
    return false;
  }
  LOCK_PROPERTY.setProperty(LOCK_KEY, now.toString());
  Logger.log("[LOCK] Lock acquired.");
  return true;
}

function releaseLock() {
  LOCK_PROPERTY.deleteProperty(LOCK_KEY);
  Logger.log("[LOCK] Lock released.");
}

// ======================== HELPER FUNCTIONS ========================

/**
 * Returns a 0-based map of column header to index.
 */
function getColumnMap(sheet) {
  if (!sheet) return {};
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { if (h) map[String(h).trim()] = i; });
  return map;
}

/**
 * Returns data rows (row 2 onward) as a 2D array.
 */
function getSheetDataSafe(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
}

/**
 * Ensures a sheet exists; creates it with optional headers if missing.
 */
function ensureSheetExists(ss, name, headers) {
  headers = headers || [];
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    Logger.log("[SETUP] Creating sheet: " + name);
    sheet = ss.insertSheet(name);
    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Ensures a column with the given header exists on the sheet.
 * Returns the 0-based column index.
 */
function getOrCreateColumn(sheet, headerName, defaultValue) {
  const cols = getColumnMap(sheet);
  if (cols[headerName] !== undefined) return cols[headerName];
  const newIdx = sheet.getLastColumn(); // 0-based new index
  sheet.getRange(1, newIdx + 1).setValue(headerName);
  if (defaultValue !== undefined && defaultValue !== null) {
    const lastRow = sheet.getMaxRows();
    if (lastRow > 1) {
      const range = sheet.getRange(2, newIdx + 1, lastRow - 1, 1);
      if (defaultValue === "checkbox") {
        range.insertCheckboxes();
      } else {
        range.setValue(defaultValue);
      }
    }
  }
  Logger.log("[COLUMN] Created column \"" + headerName + "\" at index " + newIdx);
  return newIdx;
}

/**
 * Converts a 0-based column index to a column letter (A, B, ..., Z, AA, ...).
 */
function colToLetter(idx) {
  let letter = "";
  while (idx >= 0) {
    letter = String.fromCharCode(65 + (idx % 26)) + letter;
    idx = Math.floor(idx / 26) - 1;
  }
  return letter;
}

function findLastDataRow(sheet) {
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i].some(function(cell) { return cell !== ""; })) return i + 1;
  }
  return 0;
}

function parseDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === 'string' && value.trim() !== '') {
    const d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function parseCheckboxValue(val) {
  if (val === true || String(val).toUpperCase() === "TRUE" || val === "yes") return true;
  return false;
}

function validateAndGetValue(value, columnName, dropdownOptions) {
  if (!value) return "";
  const str = String(value).trim();
  return dropdownOptions[columnName]?.has(str) ? str : "";
}

function getDropdownOptions(listsSheet) {
  if (!listsSheet) return {};
  const lr = listsSheet.getLastRow();
  const lc = listsSheet.getLastColumn();
  if (lr <= 1 || lc === 0) return {};
  const headers = listsSheet.getRange(1, 1, 1, lc).getValues()[0];
  const data = listsSheet.getRange(2, 1, lr - 1, lc).getValues();
  const options = {};
  headers.forEach((h, i) => {
    const key = String(h).trim();
    if (key) options[key] = new Set(data.map(r => String(r[i] || "").trim()).filter(Boolean));
  });
  return options;
}

// ======================== MENU ========================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Task Manager")
    .addItem("Initial Setup", "initialSetup")
    .addSeparator()
    .addItem("Run Full Sync", "syncAllSheets")
    .addItem("Repair Data & Sync", "repairAndSync")
    .addSeparator()
    .addItem("Open Pomodoro Timer 🍅", "showPomodoroTimer")
    .addItem("Mark Habit Done Today", "markHabitDoneToday")
    .addItem("Archive Completed Tasks", "archiveCompletedTasks")
    .addSeparator()
    .addItem("Send Daily Summary Email", "sendDailyNotifications")
    .addItem("Send Weekly Review Email", "sendWeeklyReview")
    .addSeparator()
    .addItem("Show Color Keys", "showColorKeys")
    .addItem("Show Coach's Advice", "showCoachingAdvice")
    .addToUi();
  Logger.log("Menu loaded.");
}

// ======================== INITIAL SETUP ========================

function initialSetup() {
  const ui = SpreadsheetApp.getUi();
  if (!acquireLock()) { ui.alert("Error", "Another process is running.", ui.ButtonSet.OK); return; }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Time Logged (mins) added to Daily Log and Master Tracker
    ensureSheetExists(ss, DAILY_LOG_SHEET_NAME,
      ["Task", "Project", "Status", "Due Date", "Priority", "Life Area", "Goal Name", "Date Added", "Importance", "Fun", "Pomodoros 🍅", "Time Logged (mins)"]);
    ensureSheetExists(ss, MASTER_PROJECT_TRACKER_SHEET_NAME,
      ["Task", "Project", "Status", "Due Date", "Priority", "Life Area", "Goal Name", "Date Added", "Term", "Importance", "Fun", "Pomodoros 🍅", "Time Logged (mins)"]);
    ensureSheetExists(ss, DASHBOARD_SHEET_NAME, []);
    ensureSheetExists(ss, LISTS_SHEET_NAME, ["Status", "Priority", "Goal Name", "Life Area", "Project"]);
    ensureSheetExists(ss, OVERDUE_LOG_SHEET_NAME, ["Task", "Project", "Original Due Date", "Days Overdue", "Date Logged"]);

    // Goal Hierarchy and Habit tracking columns
    const goalsSheet = ensureSheetExists(ss, GOALS_SHEET_NAME,
      ["Goal Level", "Goal Name", "Parent Goal", "Linked Life Area", "Progress", "Date Added", "Habit Name", "Streak 🔥", "Last Completed", "Times This Week"]);
    getOrCreateColumn(goalsSheet, "Goal Level");
    getOrCreateColumn(goalsSheet, "Goal Name");
    getOrCreateColumn(goalsSheet, "Parent Goal");
    getOrCreateColumn(goalsSheet, "Progress");
    getOrCreateColumn(goalsSheet, "Habit Name");
    getOrCreateColumn(goalsSheet, "Streak 🔥");
    getOrCreateColumn(goalsSheet, "Last Completed");
    getOrCreateColumn(goalsSheet, "Times This Week");

    ensureSheetExists(ss, BRAIN_DUMP_SHEET_NAME, ["Idea", "Link to Task/Goal"]);

    const archiveSheet = ensureSheetExists(ss, ARCHIVE_SHEET_NAME,
      ["Task", "Project", "Status", "Due Date", "Priority", "Life Area", "Goal Name", "Date Added", "Archived On"]);
    try { archiveSheet.hideSheet(); } catch (e) {}

    ui.alert("Starting Setup", "Configuring sheets, triggers, and data validation…", ui.ButtonSet.OK);

    ensureStandardLists();
    phase1Setup();
    setupImportanceAndFunCheckboxes();
    formatDateColumns(ss);
    addDropdownsToMasterProjectTracker();
    addProgressBarsToGoals(ss);
    addQuickFiltersToGoals(ss);
    setupTriggers();

    // Setup Date Picker Calendars on both sheets as requested
    setupDueDateDatePicker(ss.getSheetByName(DAILY_LOG_SHEET_NAME));
    setupDueDateDatePicker(ss.getSheetByName(MASTER_PROJECT_TRACKER_SHEET_NAME));

    syncAllSheets();

    ui.alert("Setup Complete", "All sheets configured. Quadrants and goals update live.", ui.ButtonSet.OK);
  } catch (e) {
    Logger.log("FATAL initialSetup: " + e.message + "\n" + e.stack);
    ui.alert("Error", "Setup failed: " + e.message, ui.ButtonSet.OK);
  } finally {
    releaseLock();
  }
}

// ======================== DATE PICKER BUILDER ========================

function setupDueDateDatePicker(sheet) {
  if (!sheet) return;
  const cols = getColumnMap(sheet);
  const dueIdx = cols["Due Date"];
  if (dueIdx === undefined) return;
  const lastRow = sheet.getMaxRows();
  if (lastRow > 1) {
    const range = sheet.getRange(2, dueIdx + 1, lastRow - 1, 1);
    const rule = SpreadsheetApp.newDataValidation()
      .requireDate()
      .setAllowInvalid(true)
      .setHelpText("Double-click to open calendar date picker.")
      .build();
    range.setDataValidation(rule);
  }
  Logger.log("[DATE PICKER] Set up Calendar validation on " + sheet.getName());
}

// ======================== REPAIR ========================

function repairAndSync() {
  const ui = SpreadsheetApp.getUi();
  if (!acquireLock()) { ui.alert("Error", "Sync already running.", ui.ButtonSet.OK); return; }

  try {
    ui.alert("Repair started. This may take a few seconds.", ui.ButtonSet.OK);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    [DAILY_LOG_SHEET_NAME, MASTER_PROJECT_TRACKER_SHEET_NAME].forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      const cols = getColumnMap(sheet);

      // Ensure Date Added / Added Date check is resolved
      const dateAddedCol = cols["Added Date"] !== undefined ? "Added Date" : "Date Added";

      ["Due Date", dateAddedCol].forEach(function(colName) {
        const idx = cols[colName];
        if (idx === undefined) return;
        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return;
        const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
        const values = range.getValues();
        let changed = 0;
        const newVals = values.map(function(row) {
          const val = row[0];
          if (val instanceof Date) return [val];
          if (typeof val === "string" && val.trim() !== "") {
            const d = new Date(val);
            if (!isNaN(d.getTime())) { changed++; return [d]; }
          }
          return [val];
        });
        if (changed > 0) { range.setValues(newVals); Logger.log("Fixed " + changed + " dates in " + sheetName + ":" + colName); }
      });
    });

    syncAllSheets();
    ui.alert("Repair complete. Check View > Logs for details.", ui.ButtonSet.OK);
  } catch (e) {
    Logger.log("Repair error: " + e.message);
    ui.alert("Error", e.message, ui.ButtonSet.OK);
  } finally {
    releaseLock();
  }
}

// ======================== SETUP HELPERS ========================

function ensureStandardLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listsSheet = ensureSheetExists(ss, LISTS_SHEET_NAME, ["Status", "Priority", "Goal Name", "Life Area", "Project"]);
  const headers = listsSheet.getRange(1, 1, 1, listsSheet.getLastColumn()).getValues()[0].map(String);

  // Manual Eisenhower Priority Dropdowns as requested
  const defaults = {
    "Status":    ["To Do", "In Progress", "Waiting", "Blocked", "✅ Done"],
    "Priority":  ["🚨 Do Now", "📅 Schedule", "👋 Delegate", "🗑️ Eliminate"],
    "Life Area": ["Health", "Career", "Finance", "Relationships", "Personal Growth", "Fun & Recreation"]
  };

  Object.keys(defaults).forEach(function(col) {
    const values = defaults[col];
    const colIdx = headers.indexOf(col);
    if (colIdx < 0) return;
    const lastRow = listsSheet.getLastRow();
    const existing = lastRow > 1
      ? listsSheet.getRange(2, colIdx + 1, lastRow - 1, 1).getValues().map(function(r) { return String(r[0]); }).filter(Boolean)
      : [];
    values.forEach(function(v) {
      if (!existing.includes(v)) {
        const row = new Array(listsSheet.getLastColumn()).fill("");
        row[colIdx] = v;
        listsSheet.appendRow(row);
      }
    });
  });
  Logger.log("Standard lists ensured.");
}

function phase1Setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ensureSheetExists(ss, MASTER_PROJECT_TRACKER_SHEET_NAME);
  const brainDumpSheet = ensureSheetExists(ss, BRAIN_DUMP_SHEET_NAME);

  getOrCreateColumn(masterSheet, "Term");
  const masterColumns = getColumnMap(masterSheet);
  const termIdx = masterColumns["Term"];
  if (termIdx !== undefined) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Short (<3 months)", "Medium (3-12 months)", "Long (>12 months)"], true).build();
    masterSheet.getRange(2, termIdx + 1, masterSheet.getMaxRows() - 1, 1).setDataValidation(rule);
  }

  const brainColumns = getColumnMap(brainDumpSheet);
  if (brainColumns["Link to Task/Goal"] === undefined) {
    const lastCol = brainDumpSheet.getLastColumn();
    brainDumpSheet.getRange(1, lastCol + 1).setValue("Link to Task/Goal");
    const masterGid = masterSheet.getSheetId();
    brainDumpSheet.getRange(2, lastCol + 1, brainDumpSheet.getMaxRows() - 1, 1)
      .setFormula("=IF(A2<>\"\", HYPERLINK(\"#gid=" + masterGid + "&range=A:A\", \"View Master\"), \"\")");
    const existingFilter = brainDumpSheet.getFilter();
    if (existingFilter) existingFilter.remove();
    brainDumpSheet.getRange(1, 1, Math.max(brainDumpSheet.getLastRow(), 1), Math.max(brainDumpSheet.getLastColumn(), 1)).createFilter();
  }

  // Set data validation for Goal Level in Goals sheet
  const goalsSheet = ss.getSheetByName(GOALS_SHEET_NAME);
  if (goalsSheet) {
    const glIdx = getOrCreateColumn(goalsSheet, "Goal Level");
    if (glIdx !== undefined) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Weekly Goal", "Monthly Goal", "Overall Objective"], true).build();
      if (goalsSheet.getMaxRows() > 1) {
        goalsSheet.getRange(2, glIdx + 1, goalsSheet.getMaxRows() - 1, 1).setDataValidation(rule);
      }
    }
  }

  Logger.log("Phase 1 done.");
}

function setupImportanceAndFunCheckboxes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [MASTER_PROJECT_TRACKER_SHEET_NAME, DAILY_LOG_SHEET_NAME].forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    ["Importance", "Fun"].forEach(function(colName) {
      const idx = getOrCreateColumn(sheet, colName);
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // Strict convert to real interactive checkboxes
        const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
        const vals = range.getValues();
        const cleanVals = vals.map(function(r) { return [parseCheckboxValue(r[0])]; });
        range.setValues(cleanVals);
        range.insertCheckboxes();
      }
    });
  });
  Logger.log("Importance & Fun checkboxes set and checked values normalized.");
}

function formatDateColumns(ss) {
  const targets = [
    { name: DAILY_LOG_SHEET_NAME,               cols: ["Due Date"] },
    { name: MASTER_PROJECT_TRACKER_SHEET_NAME,   cols: ["Due Date"] },
    { name: OVERDUE_LOG_SHEET_NAME,              cols: ["Original Due Date", "Date Logged"] }
  ];
  targets.forEach(function(t) {
    const sheet = ss.getSheetByName(t.name);
    if (!sheet) return;
    const colMap = getColumnMap(sheet);

    // Add Added Date check
    const addedColName = colMap["Added Date"] !== undefined ? "Added Date" : "Date Added";
    const fullCols = t.cols.concat([addedColName]);

    fullCols.forEach(function(c) {
      const idx = colMap[c];
      if (idx !== undefined) sheet.getRange(2, idx + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-MM-dd");
    });
  });
  Logger.log("Date columns formatted.");
}

function addDropdownsToMasterProjectTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName(MASTER_PROJECT_TRACKER_SHEET_NAME);
  const dailySheet = ss.getSheetByName(DAILY_LOG_SHEET_NAME);
  const listsSheet  = ss.getSheetByName(LISTS_SHEET_NAME);
  if (!listsSheet) return;

  const listHeaders = getColumnMap(listsSheet);
  const listData    = getSheetDataSafe(listsSheet);

  ["Life Area", "Project", "Status", "Goal Name", "Priority"].forEach(function(header) {
    const lIdx = listHeaders[header];
    if (lIdx === undefined) return;
    const values = listData.map(function(r) { return r[lIdx]; }).filter(Boolean);
    const unique = values.filter(function(v, i, a) { return a.indexOf(v) === i; }).map(String);
    if (unique.length > 0) {
      const rule = SpreadsheetApp.newDataValidation().requireValueInList(unique, true).build();
      if (masterSheet) {
        const mCols = getColumnMap(masterSheet);
        const mIdx = mCols[header];
        if (mIdx !== undefined) {
          masterSheet.getRange(2, mIdx + 1, masterSheet.getMaxRows() - 1, 1).setDataValidation(rule);
        }
      }
      if (dailySheet) {
        const dCols = getColumnMap(dailySheet);
        const dIdx = dCols[header];
        if (dIdx !== undefined) {
          dailySheet.getRange(2, dIdx + 1, dailySheet.getMaxRows() - 1, 1).setDataValidation(rule);
        }
      }
    }
  });
  Logger.log("Dropdowns added.");
}

function addProgressBarsToGoals(ss) {
  const sheet = ss.getSheetByName(GOALS_SHEET_NAME);
  if (!sheet) return;
  const cols = getColumnMap(sheet);
  if (cols["Progress"] !== undefined) {
    const lastRow = findLastDataRow(sheet);
    if (lastRow > 1) sheet.getRange(2, cols["Progress"] + 1, lastRow - 1, 1).setNumberFormat("0%");
  }
}

function addQuickFiltersToGoals(ss) {
  const sheet = ss.getSheetByName(GOALS_SHEET_NAME);
  if (!sheet || sheet.getLastColumn() === 0) return;
  const f = sheet.getFilter();
  if (f) f.remove();
  sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), sheet.getLastColumn()).createFilter();
}

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (["syncAllSheets", "sendDailyNotifications", "sendWeeklyReview", "onEdit"].indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("syncAllSheets").timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger("sendDailyNotifications").timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger("sendWeeklyReview").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ScriptApp.newTrigger("onEdit").forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();

  Logger.log("Triggers created.");
}

// ======================== QUADRANTS — FUN / IMPORTANCE ========================

function createFunImportanceQuadrants(dashboardSheet, dailyLogSheet) {
  Logger.log("[QUADRANT] Building Fun/Importance quadrants...");
  if (!dailyLogSheet) return;

  const dCols = getColumnMap(dailyLogSheet);
  const taskCol    = dCols["Task"];
  const projectCol = dCols["Project"];
  const impCol     = dCols["Importance"];
  const funCol     = dCols["Fun"];
  const statusCol  = dCols["Status"];
  const dueCol     = dCols["Due Date"];

  if ([taskCol, impCol, funCol, statusCol, dueCol].some(function(v) { return v === undefined; })) {
    Logger.log("[QUADRANT] Missing required columns. Found: " + Object.keys(dCols).join(", "));
    return;
  }

  const T  = colToLetter(taskCol);
  const Pr = projectCol !== undefined ? colToLetter(projectCol) : T;
  const I  = colToLetter(impCol);
  const F  = colToLetter(funCol);
  const S  = colToLetter(statusCol);
  const D  = colToLetter(dueCol);
  const sn = DAILY_LOG_SHEET_NAME;

  // Wrapped in UNIQUE and ARRAY_CONSTRAIN to avoid duplicate display & expansion errors
  function makeFilter(impCond, funCond) {
    return "=IFERROR(ARRAY_CONSTRAIN(UNIQUE(FILTER({'" + sn + "'!" + T + "2:" + T + ",'" + sn + "'!" + Pr + "2:" + Pr + "}," +
      "('" + sn + "'!" + I + "2:" + I + impCond + ")*" +
      "('" + sn + "'!" + F + "2:" + F + funCond + ")*" +
      "('" + sn + "'!" + S + "2:" + S + "<>\"✅ Done\")" +
      ")), 11, 2),\"\")";
  }

  const sRow = 12;
  const sCol = 2;
  const qH   = 11;
  const qW   = 2;
  const gap  = 2;

  dashboardSheet.getRange(10, 1, sRow + qH * 2 + gap + 6, sCol + qW * 2 + 3).clearContent().clearFormat();

  // Section title
  dashboardSheet.getRange(sRow - 2, sCol).setValue("📊 Fun / Importance Matrix")
    .setFontWeight("bold").setFontSize(14);

  // Column headers
  dashboardSheet.getRange(sRow - 1, sCol).setValue("NOT FUN 😐")
    .setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange(sRow - 1, sCol + qW).setValue("FUN 🎉")
    .setFontWeight("bold").setHorizontalAlignment("center");

  // Row headers
  dashboardSheet.getRange(sRow, sCol - 1, qH, 1).merge()
    .setValue("IMPORTANT ⭐").setFontWeight("bold").setFontColor("#c0392b")
    .setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
  dashboardSheet.getRange(sRow + qH + gap, sCol - 1, qH, 1).merge()
    .setValue("NOT IMPORTANT").setFontWeight("bold").setFontColor("#7f8c8d")
    .setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");

  // Q1: Important + Not Fun → Do First
  dashboardSheet.getRange(sRow, sCol, qH, qW).setBackground("#fce4ec");
  dashboardSheet.getRange(sRow, sCol).setFormula(makeFilter("=TRUE", "=FALSE"));
  // Q2: Important + Fun → Do Now & Enjoy
  dashboardSheet.getRange(sRow, sCol + qW, qH, qW).setBackground("#e8f5e9");
  dashboardSheet.getRange(sRow, sCol + qW).setFormula(makeFilter("=TRUE", "=TRUE"));
  // Q3: Not Important + Not Fun → Avoid / Delegate
  dashboardSheet.getRange(sRow + qH + gap, sCol, qH, qW).setBackground("#fafafa");
  dashboardSheet.getRange(sRow + qH + gap, sCol).setFormula(makeFilter("=FALSE", "=FALSE"));
  // Q4: Not Important + Fun → Do Last / Breaks
  dashboardSheet.getRange(sRow + qH + gap, sCol + qW, qH, qW).setBackground("#fff8e1");
  dashboardSheet.getRange(sRow + qH + gap, sCol + qW).setFormula(makeFilter("=FALSE", "=TRUE"));

  Logger.log("[QUADRANT] Fun/Importance quadrants created.");
}

// ======================== QUADRANTS — EISENHOWER MATRIX ========================

function createEisenhowerMatrix(dashboardSheet, dailyLogSheet) {
  Logger.log("[EISENHOWER] Building manual Eisenhower matrix...");
  if (!dailyLogSheet) return;

  const dCols = getColumnMap(dailyLogSheet);
  const taskCol    = dCols["Task"];
  const projectCol = dCols["Project"];
  const priorityCol = dCols["Priority"];
  const statusCol  = dCols["Status"];

  if ([taskCol, priorityCol, statusCol].some(function(v) { return v === undefined; })) {
    Logger.log("[EISENHOWER] Missing columns.");
    return;
  }

  const T  = colToLetter(taskCol);
  const Pr = projectCol !== undefined ? colToLetter(projectCol) : T;
  const P  = colToLetter(priorityCol);
  const S  = colToLetter(statusCol);
  const sn = DAILY_LOG_SHEET_NAME;

  // Reads explicitly from the Priority column now (Manual assignment with UNIQUE prevention)
  function makeEisen(priorityStr) {
    return "=IFERROR(ARRAY_CONSTRAIN(UNIQUE(FILTER({'" + sn + "'!" + T + "2:" + T + ",'" + sn + "'!" + Pr + "2:" + Pr + "}," +
      "('" + sn + "'!" + P + "2:" + P + "=\"" + priorityStr + "\")*" +
      "('" + sn + "'!" + S + "2:" + S + "<>\"✅ Done\")" +
      ")), 11, 2),\"\")";
  }

  const sRow = 42;
  const sCol = 2;
  const qH   = 11;
  const qW   = 2;
  const gap  = 2;

  dashboardSheet.getRange(sRow - 3, sCol - 1, qH * 2 + gap + 7, qW * 2 + 3).clearContent().clearFormat();

  // Section title
  dashboardSheet.getRange(sRow - 2, sCol).setValue("🗂️ Eisenhower Matrix (Priority-Based)")
    .setFontWeight("bold").setFontSize(14);

  // Column headers
  dashboardSheet.getRange(sRow - 1, sCol).setValue("NOT URGENT ⏳")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#1565c0");
  dashboardSheet.getRange(sRow - 1, sCol + qW).setValue("URGENT 🔥")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#c62828");

  // Row labels
  dashboardSheet.getRange(sRow, sCol - 1, qH, 1).merge()
    .setValue("IMPORTANT ⭐").setFontWeight("bold").setFontColor("#c0392b")
    .setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");
  dashboardSheet.getRange(sRow + qH + gap, sCol - 1, qH, 1).merge()
    .setValue("NOT IMPORTANT").setFontWeight("bold").setFontColor("#7f8c8d")
    .setTextRotation(90).setVerticalAlignment("middle").setHorizontalAlignment("center");

  // Q2: Schedule (blue)
  dashboardSheet.getRange(sRow, sCol, qH, qW).setBackground("#e3f2fd");
  dashboardSheet.getRange(sRow, sCol).setFormula(makeEisen("📅 Schedule"));
  dashboardSheet.getRange(sRow - 1, sCol).setValue("📅 Schedule")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#1565c0");

  // Q1: Do Now (red)
  dashboardSheet.getRange(sRow, sCol + qW, qH, qW).setBackground("#ffebee");
  dashboardSheet.getRange(sRow, sCol + qW).setFormula(makeEisen("🚨 Do Now"));
  dashboardSheet.getRange(sRow - 1, sCol + qW).setValue("🚨 Do Now")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#c62828");

  // Q4: Eliminate (grey)
  dashboardSheet.getRange(sRow + qH + gap, sCol, qH, qW).setBackground("#f5f5f5");
  dashboardSheet.getRange(sRow + qH + gap, sCol).setFormula(makeEisen("🗑️ Eliminate"));
  dashboardSheet.getRange(sRow + qH + gap - 1, sCol).setValue("🗑️ Eliminate")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#757575");

  // Q3: Delegate (orange)
  dashboardSheet.getRange(sRow + qH + gap, sCol + qW, qH, qW).setBackground("#fff3e0");
  dashboardSheet.getRange(sRow + qH + gap, sCol + qW).setFormula(makeEisen("👋 Delegate"));
  dashboardSheet.getRange(sRow + qH + gap - 1, sCol + qW).setValue("👋 Delegate")
    .setFontWeight("bold").setHorizontalAlignment("center").setFontColor("#e65100");

  Logger.log("[EISENHOWER] Matrix created.");
}

// ======================== DASHBOARD — TODAY'S TASKS ========================

function createTodaysTasks(dashboardSheet, dailyLogSheet) {
  Logger.log("[TODAY] Building Today's Tasks...");
  if (!dailyLogSheet) return;

  const dCols = getColumnMap(dailyLogSheet);
  const taskCol    = dCols["Task"];
  const projectCol = dCols["Project"];
  const dueCol     = dCols["Due Date"];
  const statusCol  = dCols["Status"];

  if ([taskCol, dueCol, statusCol].some(function(v) { return v === undefined; })) {
    Logger.log("[TODAY] Missing required columns.");
    return;
  }

  const T  = colToLetter(taskCol);
  const Pr = projectCol !== undefined ? colToLetter(projectCol) : T;
  const D  = colToLetter(dueCol);
  const S  = colToLetter(statusCol);
  const sn = DAILY_LOG_SHEET_NAME;

  dashboardSheet.getRange(1, 1, 9, 6).clearContent().clearFormat();

  dashboardSheet.getRange(1, 2).setValue("📅 Today's Tasks")
    .setFontWeight("bold").setFontSize(14).setFontColor("#1a1a2e");

  const formula = "=IFERROR(ARRAY_CONSTRAIN(UNIQUE(FILTER({'" + sn + "'!" + T + "2:" + T + ", '" + sn + "'!" + Pr + "2:" + Pr + "}, " +
    "'" + sn + "'!" + D + "2:" + D + " = TODAY(), " +
    "'" + sn + "'!" + S + "2:" + S + " <> \"✅ Done\")), 8, 2), \"🎉 No tasks due today!\")";

  dashboardSheet.getRange(3, 2, 8, 2).setBackground("#fffde7");
  dashboardSheet.getRange(3, 2).setFormula(formula);
}

// ======================== ONEDIT ========================

function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    Logger.log("[ONEDIT] " + sheetName + " R" + row + "C" + col + " = " + e.value);

    // Automatically add today's date in 'Date Added' / 'Added Date' column of edited row for new tasks
    if (sheetName === DAILY_LOG_SHEET_NAME || sheetName === MASTER_PROJECT_TRACKER_SHEET_NAME) {
      const cols = getColumnMap(sheet);
      const addedColName = cols["Added Date"] !== undefined ? "Added Date" : "Date Added";

      if (cols["Task"] !== undefined && col === cols["Task"] + 1) {
        const taskVal = sheet.getRange(row, col).getValue();
        const dateAddedCol = cols[addedColName];
        if (taskVal && dateAddedCol !== undefined) {
          const dateAddedCell = sheet.getRange(row, dateAddedCol + 1);
          if (!dateAddedCell.getValue()) {
            dateAddedCell.setValue(new Date()).setNumberFormat("yyyy-MM-dd");
          }
        }
      }
    }

    // Update goal progress when Master status changes
    if (sheetName === MASTER_PROJECT_TRACKER_SHEET_NAME) {
      const cols = getColumnMap(sheet);
      if (cols["Status"] !== undefined && col === cols["Status"] + 1) {
        const goalName = sheet.getRange(row, cols["Goal Name"] + 1).getValue();
        if (goalName) calculateGoalHierarchyProgress(e.source);
      }
    }

    // Rebuild dashboard on Daily Log edits
    if (sheetName === DAILY_LOG_SHEET_NAME) {
      const cols = getColumnMap(sheet);
      const watchCols = ["Importance", "Fun", "Due Date", "Status", "Priority"]
        .map(function(c) { return cols[c]; })
        .filter(function(v) { return v !== undefined; });
      if (watchCols.indexOf(col - 1) >= 0) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
        if (dashboard) {
          updateDashboard(ss, sheet, dashboard, cols);
        }
      }
    }

    // Quadrant checkbox on Dashboard — mark task done
    if (sheetName === DASHBOARD_SHEET_NAME && range.isChecked()) {
      handleQuadrantCheckbox(range);
    }
  } catch (err) {
    Logger.log("onEdit error: " + err.message);
  }
}

function handleQuadrantCheckbox(checkedRange) {
  const dash = checkedRange.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailyLog = ss.getSheetByName(DAILY_LOG_SHEET_NAME);
  if (!dailyLog) return;

  const taskName = dash.getRange(checkedRange.getRow(), checkedRange.getColumn() + 1).getValue();
  if (!taskName) return;

  const data = getSheetDataSafe(dailyLog);
  const cols = getColumnMap(dailyLog);

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][cols["Task"]] || "").trim() === String(taskName).trim()) {
      dailyLog.getRange(i + 2, cols["Status"] + 1).setValue("✅ Done");
      checkedRange.uncheck();
      SpreadsheetApp.flush();
      calculateGoalHierarchyProgress(ss);
      Logger.log("Marked \"" + taskName + "\" as done.");
      break;
    }
  }
}

// ======================== HIERARCHICAL GOAL TRACKING ========================

function calculateGoalHierarchyProgress(ss) {
  Logger.log("[GOALS] Calculating Hierarchy Progress (Existing Goals Only)...");
  const master = ss.getSheetByName(MASTER_PROJECT_TRACKER_SHEET_NAME);
  const goals  = ss.getSheetByName(GOALS_SHEET_NAME);
  if (!master || !goals) return;

  const mCols = getColumnMap(master);
  const gCols = getColumnMap(goals);
  if (gCols["Goal Name"] === undefined || gCols["Progress"] === undefined) return;

  const mData = getSheetDataSafe(master);
  const gData = getSheetDataSafe(goals);

  // Map out tasks grouped by Goal Name from Master Project Tracker
  const taskGoalMap = {};
  mData.forEach(function(row) {
    const goalName = String(row[mCols["Goal Name"]] || "").trim();
    if (!goalName) return;
    if (!taskGoalMap[goalName]) {
      taskGoalMap[goalName] = { total: 0, done: 0 };
    }
    taskGoalMap[goalName].total++;
    const status = String(row[mCols["Status"]] || "").trim();
    if (status === "✅ Done" || status === "Done") taskGoalMap[goalName].done++;
  });

  const goalProgressTracker = {};

  // Step 1: Calculate direct task-based progress for any existing goal in Goals & Habits
  // if there are associated tasks in the Master Project Tracker.
  // If there are no tasks, initialize with its current Progress value.
  gData.forEach(function(row) {
    const goalName = String(row[gCols["Goal Name"]] || "").trim();
    if (!goalName) return;

    if (taskGoalMap[goalName]) {
      const entry = taskGoalMap[goalName];
      goalProgressTracker[goalName] = entry.total > 0 ? entry.done / entry.total : 0;
    } else {
      goalProgressTracker[goalName] = Number(row[gCols["Progress"]]) || 0;
    }
  });

  // Step 2: Build parent-to-children relationship mappings of existing goals
  const parentToChildren = {};
  gData.forEach(function(row) {
    const goalName = String(row[gCols["Goal Name"]] || "").trim();
    const parentGoal = String(row[gCols["Parent Goal"]] || "").trim();
    if (goalName && parentGoal) {
      if (!parentToChildren[parentGoal]) parentToChildren[parentGoal] = [];
      parentToChildren[parentGoal].push(goalName);
    }
  });

  // Step 3: Cascade progress upwards for parent goals that do NOT have direct tasks in Master Tracker
  // (If a parent goal has direct tasks associated with it, we let the direct task-based calculation take precedence)
  for (let iter = 0; iter < 3; iter++) {
    gData.forEach(function(row) {
      const goalName = String(row[gCols["Goal Name"]] || "").trim();
      if (!goalName) return;

      // Only cascade if this goal does NOT have direct tasks in the Master Project Tracker
      if (!taskGoalMap[goalName]) {
        const children = parentToChildren[goalName] || [];
        if (children.length > 0) {
          let sum = 0;
          let validCount = 0;
          children.forEach(function(childName) {
            if (goalProgressTracker[childName] !== undefined) {
              sum += goalProgressTracker[childName];
              validCount++;
            }
          });
          if (validCount > 0) {
            goalProgressTracker[goalName] = sum / validCount;
          }
        }
      }
    });
  }

  // Step 4: Write calculated progress back to the Goals & Habits sheet
  gData.forEach(function(row, idx) {
    const goalName = String(row[gCols["Goal Name"]] || "").trim();
    if (goalName && goalProgressTracker[goalName] !== undefined) {
      goals.getRange(idx + 2, gCols["Progress"] + 1).setValue(goalProgressTracker[goalName]);
    }
  });

  Logger.log("[GOALS] Goals progress recalculation complete.");
}

// ======================== HABIT TRACKER ========================

function markHabitDoneToday() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const goals = ss.getSheetByName(GOALS_SHEET_NAME);
  if (!goals) { ui.alert("Goals & Habits sheet not found."); return; }

  const cols = getColumnMap(goals);
  const habitCol = cols["Habit Name"];
  if (habitCol === undefined) {
    ui.alert("Error", "No 'Habit Name' column found. Please run 'Initial Setup' first.", ui.ButtonSet.OK);
    return;
  }

  const data = getSheetDataSafe(goals);
  const habits = data
    .map(function(row, i) { return { name: String(row[habitCol] || "").trim(), rowIdx: i }; })
    .filter(function(h) { return h.name !== ""; });

  if (habits.length === 0) {
    ui.alert("No habits found. Add habit names in the 'Habit Name' column of Goals & Habits.");
    return;
  }

  const habitList = habits.map(function(h, i) { return (i + 1) + ". " + h.name; }).join("\n");
  const response = ui.prompt(
    "Mark Habit Done Today",
    "Which habit did you complete today?\nEnter the number:\n\n" + habitList,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;
  const choice = parseInt(response.getResponseText().trim(), 10);
  if (isNaN(choice) || choice < 1 || choice > habits.length) {
    ui.alert("Invalid choice."); return;
  }

  const habit = habits[choice - 1];
  const sheetRow = habit.rowIdx + 2;
  const today = new Date(); today.setHours(0, 0, 0, 0);

  const streakCol    = cols["Streak 🔥"];
  const lastCompCol  = cols["Last Completed"];
  const timesWeekCol = cols["Times This Week"];

  if (streakCol === undefined || lastCompCol === undefined || timesWeekCol === undefined) {
    ui.alert("Error", "Habit tracking columns missing. Run Initial Setup first.", ui.ButtonSet.OK);
    return;
  }

  const rowData = goals.getRange(sheetRow, 1, 1, goals.getLastColumn()).getValues()[0];
  let streak    = Number(rowData[streakCol])    || 0;
  let lastComp  = rowData[lastCompCol];
  let timesWeek = Number(rowData[timesWeekCol]) || 0;

  const lastDate = lastComp instanceof Date
    ? new Date(lastComp)
    : (lastComp ? new Date(lastComp) : null);
  if (lastDate) lastDate.setHours(0, 0, 0, 0);

  // Already logged today?
  if (lastDate && lastDate.getTime() === today.getTime()) {
    ui.alert("✅ You already marked \"" + habit.name + "\" as done today!");
    return;
  }

  // Streak logic
  if (lastDate) {
    const dayDiff = Math.round((today - lastDate) / 86400000);
    streak = dayDiff === 1 ? streak + 1 : 1;
  } else {
    streak = 1;
  }

  // Times-this-week logic: reset if a new week has started since last completion
  if (lastDate) {
    const dayOfWeek = today.getDay(); // 0=Sun
    const lastDayOfWeek = lastDate.getDay();
    const msPerWeek = 7 * 86400000;
    const weeksSinceLast = Math.floor((today - lastDate) / msPerWeek);
    if (weeksSinceLast >= 1 || (dayOfWeek <= lastDayOfWeek && today > lastDate)) {
      timesWeek = 1;
    } else {
      timesWeek += 1;
    }
  } else {
    timesWeek = 1;
  }

  goals.getRange(sheetRow, streakCol    + 1).setValue(streak);
  goals.getRange(sheetRow, lastCompCol  + 1).setValue(today);
  goals.getRange(sheetRow, timesWeekCol + 1).setValue(timesWeek);

  ui.alert(
    "🔥 Habit \"" + habit.name + "\" marked done!\n" +
    "Streak: " + streak + " day" + (streak !== 1 ? "s" : "") + "\n" +
    "This week: " + timesWeek + " time" + (timesWeek !== 1 ? "s" : "")
  );
  Logger.log("[HABIT] " + habit.name + ": streak=" + streak + ", timesWeek=" + timesWeek);
}

// ======================== FULL SYNC ========================

function syncAllSheets() {
  Logger.log("=== FULL SYNC START ===");
  if (!acquireLock()) { Logger.log("Sync skipped: lock held."); return; }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const dailyLog   = ensureSheetExists(ss, DAILY_LOG_SHEET_NAME);
    const master     = ensureSheetExists(ss, MASTER_PROJECT_TRACKER_SHEET_NAME);
    const dashboard  = ensureSheetExists(ss, DASHBOARD_SHEET_NAME);
    const lists      = ensureSheetExists(ss, LISTS_SHEET_NAME);
    const overdueLog = ensureSheetExists(ss, OVERDUE_LOG_SHEET_NAME);
    const goals      = ensureSheetExists(ss, GOALS_SHEET_NAME);
    let chartData    = ss.getSheetByName(CHART_DATA_SHEET_NAME);
    if (!chartData) { chartData = ss.insertSheet(CHART_DATA_SHEET_NAME); chartData.hideSheet(); }

    const dCols = getColumnMap(dailyLog);
    const mCols = getColumnMap(master);
    const oCols = getColumnMap(overdueLog);

    // Apply Calendar Date Pickers on sync as well
    setupDueDateDatePicker(dailyLog);
    setupDueDateDatePicker(master);

    syncDailyLogToMaster(dailyLog, master, dCols, mCols);
    syncMasterToDailyLog(dailyLog, master, dCols, mCols, ss.getSpreadsheetTimeZone());
    logOverdueTasks(dailyLog, overdueLog, dCols, oCols);
    calculateGoalHierarchyProgress(ss);
    enhanceDailyLogColors(dailyLog, dCols);
    enhanceMasterSheetColors(master, mCols);
    archiveCompletedTasksInternal(ss, dailyLog, master, dCols, mCols);
    reorderDailyLogByPriority(dailyLog, dCols);
    updateDashboard(ss, dailyLog, dashboard, dCols);

    Logger.log("=== FULL SYNC END ===");
  } catch (e) {
    Logger.log("FATAL syncAllSheets: " + e.message + "\n" + e.stack);
  } finally {
    releaseLock();
  }
}

// ======================== SYNC: DAILY LOG TO MASTER ========================

function syncDailyLogToMaster(dailyLog, master, dCols, mCols) {
  Logger.log("[D->M] Starting sync from Daily Log to Master.");
  const dData = getSheetDataSafe(dailyLog);
  const mData = getSheetDataSafe(master);

  // Multi-criteria compound matching key (Task + Project + Goal Name) so dates are free to synchronize!
  const getCompoundSyncKey = (name, proj, goal) => {
    return `${String(name || "").trim()}|${String(proj || "").trim()}|${String(goal || "").trim()}`;
  };

  const masterMap = new Map();
  mData.forEach(function(row, i) {
    const name = String(row[mCols["Task"]] || "").trim();
    if (name) {
      const key = getCompoundSyncKey(name, row[mCols["Project"]], row[mCols["Goal Name"]]);
      masterMap.set(key, i + 2); // row index
    }
  });

  let updates = 0;
  let additions = 0;
  const masterLastCol = master.getLastColumn();

  // Handle support for "Added Date" or "Date Added" column dynamically
  const dAddedColName = dCols["Added Date"] !== undefined ? "Added Date" : "Date Added";
  const mAddedColName = mCols["Added Date"] !== undefined ? "Added Date" : "Date Added";

  dData.forEach(function(dRow, i) {
    const taskName = String(dRow[dCols["Task"]] || "").trim();
    if (!taskName) return;

    function getD(col) { return dCols[col] !== undefined ? dRow[dCols[col]] : ""; }

    const status     = getD("Status")         || "To Do";
    const dueDate    = getD("Due Date");
    const priority   = getD("Priority");
    const lifeArea   = getD("Life Area");
    const goalName   = getD("Goal Name");
    const project    = getD("Project");
    const importance = parseCheckboxValue(getD("Importance"));
    const fun        = parseCheckboxValue(getD("Fun"));
    const pomodoros  = getD("Pomodoros 🍅");
    const timeLogged = getD("Time Logged (mins)");

    const compoundKey = getCompoundSyncKey(taskName, project, goalName);

    if (masterMap.has(compoundKey)) {
      const mRowNum = masterMap.get(compoundKey);
      const mRow = master.getRange(mRowNum, 1, 1, masterLastCol).getValues()[0];

      const fieldsToSync = [
        { col: "Status",                 val: status },
        { col: "Due Date",               val: dueDate },
        { col: "Priority",               val: priority },
        { col: "Importance",             val: importance },
        { col: "Fun",                    val: fun },
        { col: "Pomodoros 🍅",           val: pomodoros },
        { col: "Time Logged (mins)",     val: timeLogged }
      ];

      fieldsToSync.forEach(function(f) {
        const mIdx = mCols[f.col];
        if (mIdx === undefined) return;

        let shouldUpdate = false;
        if (f.col === "Due Date") {
          const dDateObj = parseDate(f.val);
          const mDateObj = parseDate(mRow[mIdx]);
          if (dDateObj && mDateObj) {
            if (dDateObj.getTime() !== mDateObj.getTime()) shouldUpdate = true;
          } else if (!!dDateObj !== !!mDateObj) {
            shouldUpdate = true;
          }
        } else if (JSON.stringify(mRow[mIdx]) !== JSON.stringify(f.val)) {
          shouldUpdate = true;
        }

        if (shouldUpdate) {
          Logger.log(`[D->M] Updating Task '${taskName}' with Key '${compoundKey}', Col '${f.col}' to '${f.val}'`);
          const targetCell = master.getRange(mRowNum, mIdx + 1);
          targetCell.setValue(f.val);
          if (f.col === "Importance" || f.col === "Fun") {
             targetCell.insertCheckboxes();
          }
          updates++;
        }
      });
    } else {
      Logger.log(`[D->M] Task '${taskName}' with key '${compoundKey}' not found in Master. Creating new row.`);
      const newRow = new Array(masterLastCol).fill("");
      function setM(col, val) { if (mCols[col] !== undefined) newRow[mCols[col]] = val; }

      setM("Task",               taskName);
      setM("Status",             status);
      setM("Due Date",           dueDate);
      setM("Priority",           priority);
      setM("Life Area",          lifeArea);
      setM("Goal Name",          goalName);
      setM("Project",            project);

      setM("Importance",         importance);
      setM("Fun",                fun);
      setM("Pomodoros 🍅",       pomodoros);
      setM("Time Logged (mins)", timeLogged);

      const cleanAddedDate = getD(dAddedColName) || new Date();
      if (mCols[mAddedColName] !== undefined) newRow[mCols[mAddedColName]] = cleanAddedDate;

      master.appendRow(newRow);

      const appendedRowIdx = master.getLastRow();
      if (mCols["Importance"] !== undefined) master.getRange(appendedRowIdx, mCols["Importance"] + 1).insertCheckboxes();
      if (mCols["Fun"] !== undefined) master.getRange(appendedRowIdx, mCols["Fun"] + 1).insertCheckboxes();

      additions++;
      masterMap.set(compoundKey, appendedRowIdx); // Prevent duplicates inside the same batch
    }
  });
  Logger.log(`[D->M] Done. Updates: ${updates}, Additions: ${additions}`);
}

// ======================== SYNC: MASTER TO DAILY LOG ========================

function syncMasterToDailyLog(dailyLog, master, dCols, mCols, tz) {
  Logger.log("[M->D] Starting sync from Master to Daily Log.");
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const sevenDaysOut = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

  const mData = getSheetDataSafe(master);
  const dData = getSheetDataSafe(dailyLog);

  // Multi-criteria compound matching key (Task + Project + Goal Name) so dates are free to synchronize!
  const getCompoundSyncKey = (name, proj, goal) => {
    return `${String(name || "").trim()}|${String(proj || "").trim()}|${String(goal || "").trim()}`;
  };

  const dailyMap = new Map();
  dData.forEach(function(row, i) {
    const name = String(row[dCols["Task"]] || "").trim();
    if (name) {
      const key = getCompoundSyncKey(name, row[dCols["Project"]], row[dCols["Goal Name"]]);
      dailyMap.set(key, i + 2); // row index
    }
  });

  let adds = 0;
  let updates = 0;
  const dLastCol = dailyLog.getLastColumn();

  // Dynamic added date support
  const dAddedColName = dCols["Added Date"] !== undefined ? "Added Date" : "Date Added";
  const mAddedColName = mCols["Added Date"] !== undefined ? "Added Date" : "Date Added";

  mData.forEach(function(mRow) {
    const taskName = String(mRow[mCols["Task"]] || "").trim();
    if (!taskName) return;

    const status = String(mRow[mCols["Status"]] || "").trim();
    if (status === "✅ Done") return;

    const dueDateRaw = mCols["Due Date"] !== undefined ? mRow[mCols["Due Date"]] : null;
    if (!dueDateRaw) return;

    const dueDate = new Date(dueDateRaw); dueDate.setHours(0, 0, 0, 0);
    const isOverdue = dueDate < today;
    const isDueSoon = dueDate >= today && dueDate <= sevenDaysOut;

    const compoundKey = getCompoundSyncKey(taskName, mRow[mCols["Project"]], mRow[mCols["Goal Name"]]);

    if (dailyMap.has(compoundKey)) {
      // Two-way synchronization: If it already exists in Daily Log, keep everything updated (especially Due Date changes!)
      const dRowNum = dailyMap.get(compoundKey);
      const dRow = dailyLog.getRange(dRowNum, 1, 1, dLastCol).getValues()[0];

      const fieldsToSync = [
        { col: "Status",                 val: status },
        { col: "Due Date",               val: dueDateRaw },
        { col: "Priority",               val: mRow[mCols["Priority"]] },
        { col: "Importance",             val: parseCheckboxValue(mRow[mCols["Importance"]]) },
        { col: "Fun",                    val: parseCheckboxValue(mRow[mCols["Fun"]]) },
        { col: "Pomodoros 🍅",           val: mRow[mCols["Pomodoros 🍅"]] },
        { col: "Time Logged (mins)",     val: mRow[mCols["Time Logged (mins)"]] }
      ];

      fieldsToSync.forEach(function(f) {
        const dIdx = dCols[f.col];
        if (dIdx === undefined) return;

        let shouldUpdate = false;
        if (f.col === "Due Date") {
          const dDateObj = parseDate(dRow[dIdx]);
          const mDateObj = parseDate(f.val);
          if (dDateObj && mDateObj) {
            if (dDateObj.getTime() !== mDateObj.getTime()) shouldUpdate = true;
          } else if (!!dDateObj !== !!mDateObj) {
            shouldUpdate = true;
          }
        } else if (JSON.stringify(dRow[dIdx]) !== JSON.stringify(f.val)) {
          shouldUpdate = true;
        }

        if (shouldUpdate) {
          Logger.log(`[M->D] Syncing update for '${taskName}' with key '${compoundKey}', Col '${f.col}' to '${f.val}'`);
          const targetCell = dailyLog.getRange(dRowNum, dIdx + 1);
          targetCell.setValue(f.val);
          if (f.col === "Importance" || f.col === "Fun") {
            targetCell.insertCheckboxes();
          }
          updates++;
        }
      });
    } else if (isOverdue || isDueSoon) {
      // If it doesn't exist and matches timeframe criteria, append it to the Daily Log
      const newRow = new Array(dLastCol).fill("");
      function setD(col, val) { if (dCols[col] !== undefined) newRow[dCols[col]] = val; }
      function getM(col) { return mCols[col] !== undefined ? mRow[mCols[col]] : ""; }

      setD("Task",               taskName);
      setD("Status",             "To Do");
      setD("Due Date",           dueDateRaw);
      setD("Priority",           getM("Priority"));
      setD("Life Area",          getM("Life Area"));
      setD("Goal Name",          getM("Goal Name"));
      setD("Project",            getM("Project"));

      const impVal = parseCheckboxValue(getM("Importance"));
      const funVal = parseCheckboxValue(getM("Fun"));

      setD("Importance",         impVal);
      setD("Fun",                funVal);
      setD("Pomodoros 🍅",       getM("Pomodoros 🍅"));
      setD("Time Logged (mins)", getM("Time Logged (mins)"));

      const cleanAddedDate = getM(mAddedColName) || new Date();
      if (dCols[dAddedColName] !== undefined) newRow[dCols[dAddedColName]] = cleanAddedDate;

      dailyLog.appendRow(newRow);

      const appendedRowIdx = dailyLog.getLastRow();
      if (dCols["Importance"] !== undefined) dailyLog.getRange(appendedRowIdx, dCols["Importance"] + 1).insertCheckboxes();
      if (dCols["Fun"] !== undefined) dailyLog.getRange(appendedRowIdx, dCols["Fun"] + 1).insertCheckboxes();

      adds++;
      dailyMap.set(compoundKey, appendedRowIdx); // Prevent duplicates inside the same batch
      Logger.log("[M->D] Added: " + taskName + " with key: " + compoundKey);
    }
  });
  Logger.log("[M->D] Done. Adds: " + adds + ", Updates: " + updates);
}

// ======================== OVERDUE LOGIC ========================

function logOverdueTasks(dailyLog, overdueLog, dCols, oCols) {
  Logger.log("[OVERDUE] Processing...");
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const data = getSheetDataSafe(dailyLog);
  if (!data.length) return;

  const oData = getSheetDataSafe(overdueLog);
  const loggedTasks = new Set(oData.map(function(r) { return String(r[oCols["Task"]] || "").trim(); }));

  let logged = 0;
  data.forEach(function(row, i) {
    const taskName   = String(row[dCols["Task"]]   || "").trim();
    const status     = String(row[dCols["Status"]] || "").trim();
    const dueDateRaw = dCols["Due Date"] !== undefined ? row[dCols["Due Date"]] : null;
    if (!taskName || !dueDateRaw || status === "✅ Done") return;

    const dueDate = new Date(dueDateRaw); dueDate.setHours(0, 0, 0, 0);
    if (dueDate >= today) return;

    // Highlight overdue row
    dailyLog.getRange(i + 2, 1, 1, dailyLog.getLastColumn()).setBackground(COLOR_OVERDUE);

    // Log to Overdue Log (once per task)
    if (!loggedTasks.has(taskName)) {
      const daysOverdue = Math.round((today - dueDate) / 86400000);
      const newRow = new Array(overdueLog.getLastColumn()).fill("");
      function set(col, val) { if (oCols[col] !== undefined) newRow[oCols[col]] = val; }
      set("Task",              taskName);
      set("Project",           dCols["Project"] !== undefined ? row[dCols["Project"]] : "");
      set("Original Due Date", dueDateRaw);
      set("Days Overdue",      daysOverdue);
      set("Date Logged",       new Date());
      overdueLog.appendRow(newRow);
      loggedTasks.add(taskName);
      logged++;
    }
  });
  Logger.log("[OVERDUE] " + logged + " tasks newly logged.");
}

// ======================== COLOR CODING ========================

function enhanceDailyLogColors(sheet, cols) {
  if (!sheet || cols["Status"] === undefined) return;
  Logger.log("[COLOR] Daily Log...");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const dueIdx = cols["Due Date"];

  data.forEach(function(row, i) {
    const status  = String(row[cols["Status"]] || "").trim();
    const dueRaw  = dueIdx !== undefined ? row[dueIdx] : null;
    const dueDate = dueRaw instanceof Date ? new Date(dueRaw) : (dueRaw ? new Date(dueRaw) : null);
    if (dueDate) dueDate.setHours(0, 0, 0, 0);
    const isOverdue = dueDate && dueDate < today && status !== "✅ Done";
    const range = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn());

    if (isOverdue) {
      range.setBackground(COLOR_OVERDUE);
    } else {
      switch (status) {
        case "To Do":       range.setBackground(COLOR_TODO);       break;
        case "In Progress": range.setBackground(COLOR_INPROGRESS); break;
        case "✅ Done":       range.setBackground(COLOR_DONE);       break;
        case "Blocked":     range.setBackground(COLOR_BLOCKED);    break;
        case "Waiting":     range.setBackground(COLOR_WAITING);    break;
        default:            range.setBackground(null);
      }
    }
  });
  Logger.log("[COLOR] Daily Log done.");
}

function enhanceMasterSheetColors(sheet, cols) {
  if (!sheet || cols["Status"] === undefined) return;
  Logger.log("[COLOR] Master...");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, cols["Status"] + 1, lastRow - 1, 1).getValues();
  data.forEach(function(row, i) {
    const cell = sheet.getRange(i + 2, cols["Status"] + 1);
    switch (String(row[0] || "").trim()) {
      case "To Do":       cell.setBackground(COLOR_TODO);       break;
      case "In Progress": cell.setBackground(COLOR_INPROGRESS); break;
      case "✅ Done":       cell.setBackground(COLOR_DONE);       break;
      case "Blocked":     cell.setBackground(COLOR_BLOCKED);    break;
      case "Waiting":     cell.setBackground(COLOR_WAITING);    break;
      default:            cell.setBackground(null);
    }
  });
  Logger.log("[COLOR] Master done.");
}

// ======================== ARCHIVE ========================

function archiveCompletedTasksInternal(ss, dailyLog, master, dCols, mCols) {
  Logger.log("[ARCHIVE] Starting...");
  const archiveSheet = ensureSheetExists(ss, ARCHIVE_SHEET_NAME,
    ["Task", "Project", "Status", "Due Date", "Priority", "Life Area", "Goal Name", "Date Added", "Archived On"]);
  const archiveCols = getColumnMap(archiveSheet);

  const dData = getSheetDataSafe(dailyLog);
  const rowsToDelete = [];

  dData.forEach(function(row, i) {
    const status   = String(row[dCols["Status"]] || "").trim();
    const taskName = String(row[dCols["Task"]]   || "").trim();
    if (status !== "✅ Done" || !taskName) return;

    function getD(col) { return dCols[col] !== undefined ? row[dCols[col]] : ""; }
    const newRow = new Array(archiveSheet.getLastColumn()).fill("");
    function setA(col, val) { if (archiveCols[col] !== undefined) newRow[archiveCols[col]] = val; }
    setA("Task",        getD("Task"));
    setA("Project",     getD("Project"));
    setA("Status",      getD("Status"));
    setA("Due Date",    getD("Due Date"));
    setA("Priority",    getD("Priority"));
    setA("Life Area",   getD("Life Area"));
    setA("Goal Name",   getD("Goal Name"));

    const dAddedColName = dCols["Added Date"] !== undefined ? "Added Date" : "Date Added";
    setA("Date Added",  getD(dAddedColName));
    setA("Archived On", new Date());
    archiveSheet.appendRow(newRow);
    rowsToDelete.push(i + 2);
    Logger.log("[ARCHIVE] " + taskName);
  });

  rowsToDelete.reverse().forEach(function(rowNum) { dailyLog.deleteRow(rowNum); });
  Logger.log("[ARCHIVE] " + rowsToDelete.length + " tasks archived.");
}

function archiveCompletedTasks() {
  const ui = SpreadsheetApp.getUi();
  if (!acquireLock()) { ui.alert("Error", "Another process is running.", ui.ButtonSet.OK); return; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dailyLog = ss.getSheetByName(DAILY_LOG_SHEET_NAME);
    const master   = ss.getSheetByName(MASTER_PROJECT_TRACKER_SHEET_NAME);
    if (!dailyLog || !master) { ui.alert("Required sheets missing."); return; }
    archiveCompletedTasksInternal(ss, dailyLog, master, getColumnMap(dailyLog), getColumnMap(master));
    ui.alert("✅ Archiving complete.");
  } catch (e) {
    Logger.log("Archive error: " + e.message);
    ui.alert("Error", e.message, ui.ButtonSet.OK);
  } finally {
    releaseLock();
  }
}

// ======================== REORDER ========================

function reorderDailyLogByPriority(sheet, cols) {
  if (!sheet || cols["Priority"] === undefined) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;
  const priorityOrder = { "🚨 Do Now": 1, "📅 Schedule": 2, "👋 Delegate": 3, "🗑️ Eliminate": 4 };
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  data.sort(function(a, b) {
    const pA = priorityOrder[String(a[cols["Priority"]] || "")] || 99;
    const pB = priorityOrder[String(b[cols["Priority"]] || "")] || 99;
    return pA - pB;
  });
  sheet.getRange(2, 1, data.length, sheet.getLastColumn()).setValues(data);
  Logger.log("[REORDER] Sorted by priority.");
}

// ======================== DASHBOARD UPDATE ========================

function updateDashboard(ss, dailyLog, dashboard, dCols) {
  createTodaysTasks(dashboard, dailyLog);
  createFunImportanceQuadrants(dashboard, dailyLog);
  createEisenhowerMatrix(dashboard, dailyLog);
  Logger.log("[DASHBOARD] Updated.");
}

// ======================== DAILY EMAIL ========================

function sendDailyNotifications() {
  Logger.log("[EMAIL] Sending daily summary...");
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dailyLog   = ss.getSheetByName(DAILY_LOG_SHEET_NAME);
    const overdueLog = ss.getSheetByName(OVERDUE_LOG_SHEET_NAME);
    if (!dailyLog) { Logger.log("Daily Log not found."); return; }

    const dCols = getColumnMap(dailyLog);
    const oCols = overdueLog ? getColumnMap(overdueLog) : {};
    const today = new Date(); today.setHours(0, 0, 0, 0);
    const dData = getSheetDataSafe(dailyLog);
    const oData = overdueLog ? getSheetDataSafe(overdueLog) : [];

    const priorityOrder = { "🚨 Do Now": 1, "📅 Schedule": 2, "👋 Delegate": 3, "🗑️ Eliminate": 4 };

    const todayTasks = dData.filter(function(row) {
      const status  = String(row[dCols["Status"]] || "").trim();
      if (status === "✅ Done") return false;
      const dueRaw  = dCols["Due Date"] !== undefined ? row[dCols["Due Date"]] : null;
      if (!dueRaw) return false;
      const d = new Date(dueRaw); d.setHours(0, 0, 0, 0);
      return d.getTime() === today.getTime();
    }).sort(function(a, b) {
      return (priorityOrder[String(a[dCols["Priority"]] || "")] || 99) -
             (priorityOrder[String(b[dCols["Priority"]] || "")] || 99);
    });

    const top3Overdue = oData.slice(0, 3);

    const allDueToday = dData.filter(function(row) {
      const dueRaw = dCols["Due Date"] !== undefined ? row[dCols["Due Date"]] : null;
      if (!dueRaw) return false;
      const d = new Date(dueRaw); d.setHours(0, 0, 0, 0);
      return d.getTime() === today.getTime();
    });
    const doneDueToday = allDueToday.filter(function(r) {
      return String(r[dCols["Status"]] || "").trim() === "✅ Done";
    }).length;
    const score = allDueToday.length > 0 ? Math.round((doneDueToday / allDueToday.length) * 100) : 0;

    const quote = QUOTES[Math.floor(Math.random() * QUOTES.length)];
    const dateStr = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "EEEE, MMMM d, yyyy");

    const todayTasksHtml = todayTasks.length > 0
      ? todayTasks.map(function(r) {
          return "<li><b>" + r[dCols["Task"]] + "</b> [" + (r[dCols["Priority"]] || "–") + "]</li>";
        }).join("")
      : "<li><i>No tasks due today.</i></li>";

    const overdueHtml = top3Overdue.length > 0
      ? top3Overdue.map(function(r) {
          return "<li style=\"color:#c0392b\"><b>" + r[oCols["Task"]] + "</b> — " + r[oCols["Days Overdue"]] + " days overdue</li>";
        }).join("")
      : "<li><i>No overdue tasks. Keep it up! 🎉</i></li>";

    const scoreColor = score >= 80 ? "#27ae60" : score >= 50 ? "#f39c12" : "#e74c3c";

    const html =
      "<div style=\"font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#f9f9f9;border-radius:10px;overflow:hidden;\">" +
        "<div style=\"background:#1a1a2e;padding:24px;color:white;\">" +
          "<h1 style=\"margin:0;font-size:22px;\">📋 Daily Summary</h1>" +
          "<p style=\"margin:4px 0 0;opacity:0.7;\">" + dateStr + "</p>" +
        "</div>" +
        "<div style=\"padding:20px;\">" +
          "<div style=\"background:#fff;border-left:4px solid #3498db;padding:12px 16px;border-radius:6px;margin-bottom:16px;\">" +
            "<p style=\"margin:0;font-style:italic;color:#555;\">“" + quote + "”</p>" +
          "</div>" +
          "<h3 style=\"color:#1a1a2e;\">📊 Productivity Score</h3>" +
          "<div style=\"background:#fff;padding:12px;border-radius:6px;text-align:center;\">" +
            "<span style=\"font-size:48px;font-weight:bold;color:" + scoreColor + ";\">" + score + "%</span>" +
            "<p style=\"margin:4px 0 0;color:#777;\">" + doneDueToday + " of " + allDueToday.length + " tasks completed today</p>" +
          "</div>" +
          "<h3 style=\"color:#1a1a2e;\">🔴 Top Overdue Tasks</h3>" +
          "<ul style=\"background:#fff;padding:12px 12px 12px 28px;border-radius:6px;\">" + overdueHtml + "</ul>" +
          "<h3 style=\"color:#1a1a2e;\">📌 Today's Tasks (by priority)</h3>" +
          "<ul style=\"background:#fff;padding:12px 12px 12px 28px;border-radius:6px;\">" + todayTasksHtml + "</ul>" +
        "</div>" +
        "<div style=\"background:#eee;padding:12px;text-align:center;font-size:12px;color:#888;\">Sent by Life Management System · Google Sheets</div>" +
      "</div>";

    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "📋 Daily Summary — " + dateStr,
      htmlBody: html
    });
    Logger.log("[EMAIL] Daily summary sent.");
  } catch (e) {
    Logger.log("[EMAIL] Error: " + e.message);
  }
}

// ======================== WEEKLY REVIEW EMAIL ========================

function sendWeeklyReview() {
  Logger.log("[WEEKLY] Sending weekly review...");
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const master = ss.getSheetByName(MASTER_PROJECT_TRACKER_SHEET_NAME);
    const goals  = ss.getSheetByName(GOALS_SHEET_NAME);
    if (!master) { Logger.log("Master Tracker not found."); return; }

    const mCols = getColumnMap(master);
    const mData = getSheetDataSafe(master);
    const now   = new Date();

    // Compute week boundaries (Mon-Mon)
    const dayOfWeek = now.getDay(); // 0=Sun
    const daysToMon = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    const startOfThisWeek = new Date(now);
    startOfThisWeek.setDate(now.getDate() - daysToMon);
    startOfThisWeek.setHours(0, 0, 0, 0);

    const endOfThisWeek = new Date(startOfThisWeek);
    endOfThisWeek.setDate(startOfThisWeek.getDate() + 7);

    const startOfLastWeek = new Date(startOfThisWeek);
    startOfLastWeek.setDate(startOfThisWeek.getDate() - 7);
    const endOfLastWeek = new Date(startOfThisWeek);
    endOfLastWeek.setMilliseconds(-1);

    function isDoneInRange(row, from, to) {
      if (String(row[mCols["Status"]] || "").trim() !== "✅ Done") return false;
      const dueRaw = mCols["Due Date"] !== undefined ? row[mCols["Due Date"]] : null;
      if (!dueRaw) return false;
      const d = new Date(dueRaw); d.setHours(0, 0, 0, 0);
      return d >= from && d <= to;
    }

    const doneThisWeek = mData.filter(function(r) { return isDoneInRange(r, startOfThisWeek, now); });
    const doneLastWeek = mData.filter(function(r) { return isDoneInRange(r, startOfLastWeek, endOfLastWeek); });

    // Show upcoming tasks for the coming week
    const upcomingTasks = mData.filter(function(r) {
      const status = String(r[mCols["Status"]] || "").trim();
      if (status === "✅ Done") return false;
      const dueRaw = mCols["Due Date"] !== undefined ? r[mCols["Due Date"]] : null;
      if (!dueRaw) return false;
      const d = new Date(dueRaw); d.setHours(0, 0, 0, 0);
      return d >= now && d <= endOfThisWeek;
    });

    const delta      = doneThisWeek.length - doneLastWeek.length;
    const deltaStr   = delta >= 0 ? "+" + delta : "" + delta;
    const deltaColor = delta >= 0 ? "#27ae60" : "#e74c3c";
    const dateStr    = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MMMM d, yyyy");

    // Progression of Existing Goals (Weekly, Monthly, Objective)
    let goalsProgressionHtml = "<li><i>No existing goals logged yet.</i></li>";
    if (goals) {
      const gCols = getColumnMap(goals);
      const gData = getSheetDataSafe(goals);
      if (gData.length > 0) {
        goalsProgressionHtml = gData.map(function(row) {
          const gName  = row[gCols["Goal Name"]] || "Unnamed Goal";
          const gLevel = row[gCols["Goal Level"]] || "Goal";
          const gProg  = Number(row[gCols["Progress"]]) || 0;
          const progPct = Math.round(gProg * 100) + "%";
          return `<li><b>${gName}</b> (${gLevel}): <span style="font-weight:bold;color:#2980b9;">${progPct}</span></li>`;
        }).join("");
      }
    }

    const thisWeekList = doneThisWeek.length > 0
      ? doneThisWeek.map(function(r) {
          return "<li>" + r[mCols["Task"]] + " [" + (r[mCols["Goal Name"]] || "–") + "]</li>";
        }).join("")
      : "<li><i>None completed this week yet.</i></li>";

    const upcomingList = upcomingTasks.length > 0
      ? upcomingTasks.map(function(r) {
          const dueRaw = r[mCols["Due Date"]];
          const dStr = dueRaw instanceof Date ? Utilities.formatDate(dueRaw, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : String(dueRaw || "");
          return "<li>" + r[mCols["Task"]] + " (Due: " + dStr + ")</li>";
        }).join("")
      : "<li><i>No upcoming tasks scheduled for this week.</i></li>";

    const html =
      "<div style=\"font-family:Arial,sans-serif;max-width:600px;margin:auto;background:#f9f9f9;border-radius:10px;overflow:hidden;\">" +
        "<div style=\"background:#1a1a2e;padding:24px;color:white;\">" +
          "<h1 style=\"margin:0;font-size:22px;\">📆 Weekly Review</h1>" +
          "<p style=\"margin:4px 0 0;opacity:0.7;\">Week of " + dateStr + "</p>" +
        "</div>" +
        "<div style=\"padding:20px;\">" +
          "<h3 style=\"color:#1a1a2e;\">🏆 Completion Comparison</h3>" +
          "<table style=\"width:100%;border-collapse:collapse;margin-bottom:16px;\">" +
            "<tr>" +
              "<td style=\"background:#fff;padding:16px;border-radius:8px;text-align:center;width:33%;\">" +
                "<div style=\"font-size:36px;font-weight:bold;color:#3498db;\">" + doneLastWeek.length + "</div>" +
                "<div style=\"color:#888;\">Last Week</div>" +
              "</td>" +
              "<td style=\"width:2%;\"></td>" +
              "<td style=\"background:#fff;padding:16px;border-radius:8px;text-align:center;width:33%;\">" +
                "<div style=\"font-size:36px;font-weight:bold;color:#1a1a2e;\">" + doneThisWeek.length + "</div>" +
                "<div style=\"color:#888;\">This Week</div>" +
              "</td>" +
              "<td style=\"width:2%;\"></td>" +
              "<td style=\"background:#fff;padding:16px;border-radius:8px;text-align:center;width:30%;\">" +
                "<div style=\"font-size:36px;font-weight:bold;color:" + deltaColor + ";\">" + deltaStr + "</div>" +
                "<div style=\"color:#888;\">Change</div>" +
              "</td>" +
            "</tr>" +
          "</table>" +
          "<h3 style=\"color:#1a1a2e;\">🎯 Goals Progression Status</h3>" +
          "<ul style=\"background:#fff;padding:12px 12px 12px 28px;border-radius:6px;\">" + goalsProgressionHtml + "</ul>" +
          "<h3 style=\"color:#1a1a2e;\">✅ Completed This Week (" + doneThisWeek.length + ")</h3>" +
          "<ul style=\"background:#fff;padding:12px 12px 12px 28px;border-radius:6px;\">" + thisWeekList + "</ul>" +
          "<h3 style=\"color:#1a1a2e;\">🔮 Upcoming Tasks for the Week</h3>" +
          "<ul style=\"background:#fff;padding:12px 12px 12px 28px;border-radius:6px;\">" + upcomingList + "</ul>" +
        "</div>" +
        "<div style=\"background:#eee;padding:12px;text-align:center;font-size:12px;color:#888;\">Sent by Life Management System · Google Sheets</div>" +
      "</div>";

    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "📆 Weekly Review — " + dateStr,
      htmlBody: html
    });
    Logger.log("[WEEKLY] Review sent.");
  } catch (e) {
    Logger.log("[WEEKLY] Error: " + e.message);
  }
}

// ======================== COACHING ADVICE ========================

function showCoachingAdvice() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const dailyLog   = ss.getSheetByName(DAILY_LOG_SHEET_NAME);
  const overdueLog = ss.getSheetByName(OVERDUE_LOG_SHEET_NAME);
  const dCols = dailyLog   ? getColumnMap(dailyLog)   : {};
  const dData = dailyLog   ? getSheetDataSafe(dailyLog)   : [];
  const oData = overdueLog ? getSheetDataSafe(overdueLog) : [];

  const totalActive  = dData.filter(function(r) { return String(r[dCols["Status"]] || "").trim() !== "✅ Done"; }).length;
  const totalOverdue = oData.length;
  const doneToday    = dData.filter(function(r) { return String(r[dCols["Status"]] || "").trim() === "✅ Done"; }).length;
  const p1Count      = dData.filter(function(r) {
    return String(r[dCols["Priority"]] || "").trim() === "🚨 Do Now" &&
           String(r[dCols["Status"]] || "").trim()   !== "✅ Done";
  }).length;

  let personalTip;
  if (totalOverdue > 5) {
    personalTip = "⚠️ You have <b>" + totalOverdue + " overdue tasks</b>. Consider a 30-minute \"rescue session\" to complete, reschedule, or delete them.";
  } else if (p1Count > 3) {
    personalTip = "🚨 You have <b>" + p1Count + " Critical 'Do Now' tasks</b> open. Focus on one at a time — multitasking on critical work reduces quality.";
  } else if (doneToday > 0) {
    personalTip = "🌟 Great work! You've completed <b>" + doneToday + " task" + (doneToday !== 1 ? "s" : "") + " today</b>. Momentum is everything!";
  } else {
    personalTip = "💡 You have <b>" + totalActive + " active tasks</b>. Pick your single most important task and work on it for 25 minutes (one Pomodoro) right now.";
  }

  const frameworkTip = FRAMEWORK_TIPS[Math.floor(Math.random() * FRAMEWORK_TIPS.length)];
  const quote = QUOTES[Math.floor(Math.random() * QUOTES.length)];

  const html =
    "<style>" +
      "body{font-family:Arial,sans-serif;color:#1a1a2e;padding:8px;}" +
      ".card{background:#f9f9f9;border-radius:8px;padding:14px;margin-bottom:12px;border-left:4px solid #3498db;}" +
      ".card.green{border-color:#27ae60;}" +
      ".card.purple{border-color:#8e44ad;}" +
      "h3{margin:0 0 6px;font-size:14px;}" +
      "p{margin:0;line-height:1.6;font-size:13px;}" +
    "</style>" +
    "<div class=\"card\"><h3>📊 Personal Insight</h3><p>" + personalTip + "</p></div>" +
    "<div class=\"card green\"><h3>📚 Framework Tip</h3><p>" + frameworkTip + "</p></div>" +
    "<div class=\"card purple\"><h3>💬 Motivation</h3><p><i>“" + quote + "”</i></p></div>" +
    "<div style=\"margin-top:15px; font-size:12px; color:#7f8c8d;\">" +
      "<h4>🧠 A Message from your Coach:</h4>" +
      "<p style=\"font-style:italic;\">" + COACHING_ADVICE_TEXT.replace(/\n/g, "<br>") + "</p>" +
    "</div>";

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(520).setHeight(500),
    "🧠 Coach's Advice"
  );
}

// ======================== COLOR KEYS ========================

function showColorKeys() {
  const html =
    "<style>" +
      "body{font-family:Arial,sans-serif;padding:8px;color:#333;}" +
      "table{border-collapse:collapse;width:100%;margin-bottom:8px;}" +
      "td{padding:7px 10px;font-size:13px;}" +
      ".sw{width:22px;height:22px;border-radius:4px;display:inline-block;vertical-align:middle;border:1px solid #ccc;}" +
      "h3{margin:12px 0 6px;font-size:14px;color:#1a1a2e;}" +
    "</style>" +
    "<h3>🎨 Status Colors</h3>" +
    "<table>" +
      "<tr><td><span class=\"sw\" style=\"background:" + COLOR_TODO       + "\"></span></td><td>To Do</td></tr>" +
      "<tr><td><span class=\"sw\" style=\"background:" + COLOR_INPROGRESS + "\"></span></td><td>In Progress</td></tr>" +
      "<tr><td><span class=\"sw\" style=\"background:" + COLOR_DONE       + "\"></span></td><td>✅ Done</td></tr>" +
      "<tr><td><span class=\"sw\" style=\"background:" + COLOR_BLOCKED    + "\"></span></td><td>Blocked</td></tr>" +
      "<tr><td><span class=\"sw\" style=\"background:" + COLOR_WAITING    + "\"></span></td><td>Waiting</td></tr>" +
      "<tr><td><span class=\"sw\" style=\"background:" + COLOR_OVERDUE    + "\"></span></td><td>Overdue (row highlight)</td></tr>" +
    "</table>" +
    "<h3>📊 Fun / Importance Matrix</h3>" +
    "<table>" +
      "<tr><td style=\"background:#e8f5e9;padding:6px;border-radius:4px;\">Important + Fun → <b>Do Now & Enjoy</b></td></tr>" +
      "<tr><td style=\"background:#fce4ec;padding:6px;border-radius:4px;\">Important + Not Fun → <b>Do First</b></td></tr>" +
      "<tr><td style=\"background:#fff8e1;padding:6px;border-radius:4px;\">Not Important + Fun → <b>Do Last / Breaks</b></td></tr>" +
      "<tr><td style=\"background:#fafafa;padding:6px;border-radius:4px;border:1px solid #ddd;\">Not Important + Not Fun → <b>Avoid / Delegate</b></td></tr>" +
    "</table>" +
    "<h3>🗂️ Eisenhower Matrix (Priority-Based)</h3>" +
    "<table>" +
      "<tr><td style=\"background:#ffebee;padding:6px;border-radius:4px;\">🚨 Do Now</td></tr>" +
      "<tr><td style=\"background:#e3f2fd;padding:6px;border-radius:4px;\">📅 Schedule</td></tr>" +
      "<tr><td style=\"background:#fff3e0;padding:6px;border-radius:4px;\">👋 Delegate</td></tr>" +
      "<tr><td style=\"background:#f5f5f5;padding:6px;border-radius:4px;border:1px solid #ddd;\">🗑️ Eliminate</td></tr>" +
    "</table>";

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(530),
    "🎨 Color Keys & Matrix Legend"
  );
}

// ======================== POMODORO TIMER ========================

function getActiveTasksForPomodoro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dailyLog = ss.getSheetByName(DAILY_LOG_SHEET_NAME);
  if (!dailyLog) return [];
  const dCols = getColumnMap(dailyLog);
  const data = getSheetDataSafe(dailyLog);
  const tasks = [];
  data.forEach(function(row) {
    const status = String(row[dCols["Status"]] || "").trim();
    if (status !== "✅ Done" && row[dCols["Task"]]) {
      tasks.push(String(row[dCols["Task"]]).trim());
    }
  });
  return tasks;
}

function logPomodoroAndTime(taskName, elapsedSeconds) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const elapsedMinutes = Math.round(elapsedSeconds / 60);
  if (elapsedMinutes < 1) {
    Logger.log(`[POMODORO] Elapsed time too short (${elapsedSeconds}s). Not logging.`);
    return;
  }
  
  function increment(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const cols = getColumnMap(sheet);
    if (cols["Task"] === undefined) return;
    
    const data = getSheetDataSafe(sheet);
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][cols["Task"]] || "").trim() === taskName) {
        
        // Append Tomato if Pomodoro column exists
        if (cols["Pomodoros 🍅"] !== undefined) {
          let currentT = data[i][cols["Pomodoros 🍅"]] || "";
          sheet.getRange(i + 2, cols["Pomodoros 🍅"] + 1).setValue(String(currentT) + "🍅");
        }
        
        // Add elapsed minutes to Time Logged
        if (cols["Time Logged (mins)"] !== undefined) {
          let currentM = Number(data[i][cols["Time Logged (mins)"]]) || 0;
          sheet.getRange(i + 2, cols["Time Logged (mins)"] + 1).setValue(currentM + elapsedMinutes);
        }

        break;
      }
    }
  }

  increment(DAILY_LOG_SHEET_NAME);
  increment(MASTER_PROJECT_TRACKER_SHEET_NAME);
  Logger.log(`[POMODORO] Logged tomato and ${elapsedMinutes} mins for: ${taskName}`);
}

function showPomodoroTimer() {
  const html = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background-color: #fdfbf7; color: #333; text-align: center; }
          h2 { color: #e74c3c; margin-top: 0; }
          select { width: 100%; padding: 10px; margin: 15px 0; border: 1px solid #ccc; border-radius: 5px; font-size: 14px; }
          .timer { font-size: 60px; font-weight: bold; margin: 20px 0; color: #2c3e50; font-variant-numeric: tabular-nums; }
          .controls { display: flex; gap: 10px; justify-content: center; margin-bottom: 20px; flex-wrap: wrap; }
          button { padding: 10px 15px; font-size: 14px; border: none; border-radius: 5px; cursor: pointer; color: white; font-weight: bold; transition: opacity 0.2s; }
          button:hover { opacity: 0.8; }
          #startBtn { background-color: #27ae60; }
          #pauseBtn { background-color: #f39c12; }
          #stopBtn { background-color: #8e44ad; }
          #resetBtn { background-color: #e74c3c; }
          .presets { display: flex; gap: 5px; justify-content: center; margin-bottom: 20px; }
          .presetBtn { background-color: #ecf0f1; color: #333; padding: 5px 10px; font-size: 12px; border: 1px solid #ddd; }
          #status { margin-top: 20px; color: #7f8c8d; font-size: 14px; min-height: 20px; }
        </style>
      </head>
      <body>
        <h2>🍅 Pomodoro Timer</h2>

        <div style="text-align: left; font-size: 12px; font-weight: bold; color: #7f8c8d; margin-bottom: 5px;">SELECT TASK</div>
        <select id="taskSelect">
          <option value="">Loading tasks...</option>
        </select>

        <div class="presets">
          <button class="presetBtn" onclick="setTime(15)">15m</button>
          <button class="presetBtn" onclick="setTime(25)">25m</button>
          <button class="presetBtn" onclick="setTime(50)">50m</button>
        </div>

        <div class="timer" id="display">25:00</div>

        <div class="controls">
          <button id="startBtn" onclick="startTimer()">Start</button>
          <button id="pauseBtn" onclick="pauseTimer()">Pause</button>
          <button id="stopBtn" onclick="stopAndLogTimer()">Stop & Log</button>
          <button id="resetBtn" onclick="resetTimer()">Reset</button>
        </div>

        <div id="status">Ready to focus.</div>

        <script>
          let timeLeft = 25 * 60;
          let defaultTime = 25 * 60;
          let elapsedSeconds = 0;
          let timerInterval = null;
          let isRunning = false;
          let sound = new Audio('https://actions.google.com/sounds/v1/alarms/digital_watch_alarm_long.ogg');

          // Load tasks from Apps Script
          google.script.run.withSuccessHandler(function(tasks) {
            const select = document.getElementById('taskSelect');
            select.innerHTML = '<option value="">-- Choose a task --</option>';
            tasks.forEach(function(t) {
              const opt = document.createElement('option');
              opt.value = t;
              opt.textContent = t;
              select.appendChild(opt);
            });
          }).getActiveTasksForPomodoro();

          function updateDisplay() {
            const m = Math.floor(timeLeft / 60).toString().padStart(2, '0');
            const s = (timeLeft % 60).toString().padStart(2, '0');
            document.getElementById('display').innerText = m + ':' + s;
          }

          function setTime(minutes) {
            if (isRunning) pauseTimer();
            defaultTime = minutes * 60;
            timeLeft = defaultTime;
            elapsedSeconds = 0;
            updateDisplay();
            document.getElementById('status').innerText = "Timer set to " + minutes + " minutes.";
          }

          function startTimer() {
            if (isRunning) return;
            const task = document.getElementById('taskSelect').value;
            if (!task) {
              document.getElementById('status').innerText = "⚠️ Please select a task first!";
              return;
            }
            isRunning = true;
            document.getElementById('status').innerText = "Focusing on: " + task;
            timerInterval = setInterval(function() {
              timeLeft--;
              elapsedSeconds++;
              updateDisplay();
              if (timeLeft <= 0) {
                clearInterval(timerInterval);
                isRunning = false;
                try { sound.play(); } catch(e) {}
                document.getElementById('status').innerText = "🎉 Focus session complete! Logging...";
                logTimeOnScript(task, elapsedSeconds);
              }
            }, 1000);
          }

          function pauseTimer() {
            if (!isRunning) return;
            clearInterval(timerInterval);
            isRunning = false;
            document.getElementById('status').innerText = "Paused.";
          }

          function stopAndLogTimer() {
            if (isRunning) {
              clearInterval(timerInterval);
              isRunning = false;
            }
            const task = document.getElementById('taskSelect').value;
            if (!task) {
              document.getElementById('status').innerText = "⚠️ Please select a task to log.";
              return;
            }
            document.getElementById('status').innerText = "Logging focus time...";
            logTimeOnScript(task, elapsedSeconds);
          }

          function resetTimer() {
            if (isRunning) clearInterval(timerInterval);
            isRunning = false;
            timeLeft = defaultTime;
            elapsedSeconds = 0;
            updateDisplay();
            document.getElementById('status').innerText = "Timer reset.";
          }

          function logTimeOnScript(task, seconds) {
            google.script.run.withSuccessHandler(function() {
              const mins = Math.round(seconds / 60);
              if (mins >= 1) {
                document.getElementById('status').innerText = "✅ Successfully logged " + mins + " min(s) to '" + task + "'!";
              } else {
                document.getElementById('status').innerText = "Logged session ended (under 1 min, not added to sheet).";
              }
              elapsedSeconds = 0;
            }).withFailureHandler(function(err) {
              document.getElementById('status').innerText = "❌ Error logging time: " + err.message;
            }).logPomodoroAndTime(task, seconds);
          }
        </script>
      </body>
    </html>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(400).setHeight(480),
    "🍅 Pomodoro Timer"
  );
}
