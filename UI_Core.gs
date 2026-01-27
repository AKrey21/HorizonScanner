/*********************************
 * UI_Core.gs — FutureScans UI Backend (Core)
 * - Web App entry (doGet)
 * - Sidebar entry (optional)
 * - include() helper for HTML templating
 * - Web-app-safe spreadsheet getter (openById)
 *********************************/

const SPREADSHEET_ID = "1TZ6QknTE3LLQtn34UrzOYRHOakZG_QlhyGtruooNGW4";
const THEME_RULES_START_ROW = 5; // ThemeRules data starts at row 5

// ✅ Web App entry — MUST use template evaluate() or your <? ?> includes will print as text
function doGet() {
  const t = HtmlService.createTemplateFromFile("index");
  return t.evaluate()
    .setTitle("FutureScans")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function forceDriveAuth_() {
  DriveApp.getRootFolder().getName();
}

// Optional sidebar (if you still use it)
function showRawArticlesSidebar() {
  const t = HtmlService.createTemplateFromFile("index");
  SpreadsheetApp.getUi().showSidebar(
    t.evaluate().setTitle("FutureScans")
  );
}

// ✅ include helper for index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Web-app-safe
function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/** Safe sheet name getters (avoid ReferenceError if constants live elsewhere) */
function getControlSheetName_() {
  try {
    if (typeof CONTROL_SHEET !== "undefined" && CONTROL_SHEET) return CONTROL_SHEET;
  } catch (e) {}
  return "ThemeRules";
}

function getRawArticlesSheetName_() {
  try {
    if (typeof RAW_SHEET !== "undefined" && RAW_SHEET) return RAW_SHEET;
  } catch (e) {}
  return "Raw Articles";
}

function getFeedsSheetName_() {
  try {
    if (typeof FEEDS_SHEET !== "undefined" && FEEDS_SHEET) return FEEDS_SHEET;
  } catch (e) {}
  return "RSS Feeds";
}
