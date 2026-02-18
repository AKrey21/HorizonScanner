/*********************************
 * UI.gs — FutureScans UI Backend
 * - Web App entry (doGet)
 * - Sidebar entry (optional)
 * - include() helper for HTML templating
 * - Web-app-safe spreadsheet getter (openById)
 * - RSS feeds UI APIs
 * - Raw articles UI APIs
 *********************************/

/** ====== Core (previously UI_Core.gs) ====== */

const SPREADSHEET_ID = "1TZ6QknTE3LLQtn34UrzOYRHOakZG_QlhyGtruooNGW4";
const THEME_RULES_START_ROW = 5; // ThemeRules data starts at row 5

function onOpen() {
  try {
    if (typeof ingest_ensureDailyRawIngestTrigger_ === "function") {
      ingest_ensureDailyRawIngestTrigger_();
    }
    ui_addHorizonScannerMenu_();
  } catch (err) {
    console.log("Failed to ensure daily raw ingest trigger:", err);
  }
}

function ui_addHorizonScannerMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("HorizonScanner")
      .addItem("Run daily raw ingest now", "ui_runDailyRawIngestNow_")
      .addItem("Install daily raw ingest trigger", "ui_installDailyRawIngestTrigger_")
      .addItem("Check raw ingest trigger status", "ui_checkDailyRawIngestTrigger_")
      .addToUi();
  } catch (err) {
    console.log("Failed to create HorizonScanner menu:", err);
  }
}

function ui_installDailyRawIngestTrigger_() {
  if (typeof ingest_setupDailyRawIngestTrigger_ !== "function") {
    SpreadsheetApp.getUi().alert("Ingest setup function not available.");
    return;
  }
  var res = ingest_setupDailyRawIngestTrigger_();
  SpreadsheetApp.getUi().alert(
    "Daily raw ingest trigger installed.\nHandler: " + res.handler + "\nHour: " + res.atHour
  );
}

function ui_checkDailyRawIngestTrigger_() {
  var handlerName = "ingest_dailyRawArticles_";
  var triggers = ScriptApp.getProjectTriggers();
  var exists = triggers.some(function (t) {
    return t.getHandlerFunction && t.getHandlerFunction() === handlerName;
  });
  var lastRun = "";
  var lastStatus = "";
  var lastMessage = "";
  var lastRuntime = "";
  var lastStoppedEarly = "";
  var lastErrors = "";
  try {
    var props = PropertiesService.getScriptProperties();
    lastRun = props.getProperty("RAW_INGEST_LAST_RUN") || "";
    lastStatus = props.getProperty("RAW_INGEST_LAST_STATUS") || "";
    lastMessage = props.getProperty("RAW_INGEST_LAST_MESSAGE") || "";
    lastRuntime = props.getProperty("RAW_INGEST_LAST_RUNTIME_MS") || "";
    lastStoppedEarly = props.getProperty("RAW_INGEST_LAST_STOPPED_EARLY") || "";
    lastErrors = props.getProperty("RAW_INGEST_LAST_ERRORS") || "";
  } catch (err) {
    console.log("Failed to read RAW_INGEST_LAST_RUN:", err);
  }
  var status = exists
    ? "Daily raw ingest trigger is installed."
    : "Daily raw ingest trigger is not installed.";
  if (lastRun) {
    status += "\nLast run: " + lastRun;
  }
  if (lastStatus) {
    status += "\nLast status: " + lastStatus;
  }
  if (lastRuntime) {
    status += "\nLast runtime (ms): " + lastRuntime;
  }
  if (lastStoppedEarly === "true") {
    status += "\nStopped early to avoid time limit.";
  }
  if (lastMessage) {
    status += "\nLast message: " + lastMessage;
  }
  if (lastErrors) {
    status += "\nLast errors: " + lastErrors;
  }
  SpreadsheetApp.getUi().alert(
    status
  );
}

function ui_runDailyRawIngestNow_() {
  if (typeof ingest_dailyRawArticles_ !== "function") {
    SpreadsheetApp.getUi().alert("Daily ingest function not available.");
    return;
  }
  var res;
  try {
    res = ingest_dailyRawArticles_();
  } catch (err) {
    SpreadsheetApp.getUi().alert("Daily ingest failed to run: " + (err && err.message ? err.message : err));
    return;
  }
  var summary = res && res.message ? res.message : "Daily ingest completed.";
  if (res && res.errors && res.errors.length) {
    summary += "\nErrors: " + res.errors.slice(0, 3).join(" | ");
  }
  SpreadsheetApp.getUi().alert(summary);
}

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

/* =========================================================================
 * UI: Approved picks sidebar (LLM cache -> Weekly Picks fallback)
 * ========================================================================= */

function ui_getApprovedPicks() {
  return rpt_getApprovedPicks_();
}

function rpt_getApprovedPicks_() {
  const llmPicks = rpt_getApprovedPicksFromLlm_();
  if (llmPicks.length) return llmPicks;
  return rpt_getWeeklyPicks_();
}

function rpt_getApprovedPicksFromLlm_() {
  if (typeof raw_readLlmRankSheet_ !== "function") return [];
  const cache = raw_readLlmRankSheet_();
  const results = Array.isArray(cache?.results) ? cache.results : [];
  if (!results.length) return [];

  const items = results
    .filter(item => rpt_isLlmRecommended_(item?.llm || {}))
    .map(item => ({
      title: String(item?.title || item?.llm?.title || "").trim(),
      url: String(item?.url || item?.llm?.url || "").trim()
    }))
    .filter(item => item.url);

  return items;
}

function rpt_getWeeklyPicks_() {
  const ss = getSpreadsheet_();
  const sheetName = (typeof PICKS_SHEET !== "undefined" && PICKS_SHEET) ? PICKS_SHEET : "Weekly Picks";
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 2) return [];

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const hmap = (typeof headerMap_ === "function")
    ? headerMap_(headers)
    : rpt_headerMap_(headers);

  const colTitle = hmap["title"] || 1;
  const colLink = hmap["link"] || hmap["url"] || 2;
  const colDate = hmap["date"] || 0;
  const colScore = hmap["score"] || 0;

  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rows = values.map(row => ({
    title: String(row[colTitle - 1] || "").trim(),
    url: String(row[colLink - 1] || "").trim(),
    date: colDate ? row[colDate - 1] : "",
    score: colScore ? Number(row[colScore - 1]) || 0 : null
  })).filter(item => item.title && item.url);

  if (!rows.length) return [];

  const hasScore = colScore > 0;
  rows.sort((a, b) => {
    if (hasScore && b.score !== a.score) return b.score - a.score;
    return rpt_safeDateMs_(b.date) - rpt_safeDateMs_(a.date);
  });

  return rows.slice(0, 12).map(item => ({ title: item.title, url: item.url }));
}

function rpt_headerMap_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const key = String(h || "").trim().toLowerCase();
    if (key) map[key] = i + 1;
  });
  return map;
}

function rpt_safeDateMs_(value) {
  const t = new Date(value).getTime();
  return isNaN(t) ? 0 : t;
}

function rpt_isLlmRecommended_(llm) {
  const raw = String((llm && (llm.publish_recommendation || llm.recommendation)) || "").trim().toLowerCase();
  const yesTokens = ["yes", "y", "publish", "feature", "recommend", "recommended", "include", "maybe"];
  const noTokens = ["no", "n", "skip", "exclude", "not"];
  if (yesTokens.some(token => raw === token || raw.indexOf(token) !== -1)) return true;
  if (noTokens.some(token => raw === token || raw.indexOf(token) !== -1)) return false;
  return false;
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

/** ====== Theme Rules (UI) ====== */

function ui_getThemeRules() {
  try {
    const sheetName = (typeof repo_getControlSheetName_ === "function")
      ? repo_getControlSheetName_()
      : getControlSheetName_();
    const sh = (typeof repo_getSheetOrThrow_ === "function")
      ? repo_getSheetOrThrow_(sheetName)
      : getSpreadsheet_().getSheetByName(sheetName);

    if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);

    const startRow =
      (typeof THEME_RULES_START_ROW !== "undefined" && Number(THEME_RULES_START_ROW) >= 1)
        ? Number(THEME_RULES_START_ROW)
        : 5;

    const lastRow = sh.getLastRow();
    if (lastRow < startRow) {
      return { ok: true, rows: [], totalRows: 0, sheet: sheetName };
    }

    const values = sh.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
    const rows = [];

    values.forEach((row, idx) => {
      const theme = String(row[0] || "").trim();
      const poi = String(row[1] || "").trim();
      const keywords = String(row[2] || "").trim();
      const activeRaw = row[3];
      const active =
        activeRaw === true || activeRaw === 1 || String(activeRaw).trim().toLowerCase() === "true";

      const hasContent = theme || poi || keywords || active;
      if (!hasContent) return;

      rows.push({
        rowIndex: startRow + idx,
        theme,
        poi,
        keywords,
        active
      });
    });

    return { ok: true, rows, totalRows: lastRow - startRow + 1, sheet: sheetName };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function ui_getThemeRuleDropdowns() {
  try {
    const sheetName = (typeof repo_getControlSheetName_ === "function")
      ? repo_getControlSheetName_()
      : getControlSheetName_();
    const sh = (typeof repo_getSheetOrThrow_ === "function")
      ? repo_getSheetOrThrow_(sheetName)
      : getSpreadsheet_().getSheetByName(sheetName);

    if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);

    const startRow =
      (typeof THEME_RULES_START_ROW !== "undefined" && Number(THEME_RULES_START_ROW) >= 1)
        ? Number(THEME_RULES_START_ROW)
        : 5;

    const lastRow = sh.getLastRow();
    if (lastRow < startRow) {
      return { ok: true, themes: [], pois: [] };
    }

    const values = sh.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
    const themes = new Set();
    const pois = new Set();

    values.forEach(row => {
      const theme = String(row[0] || "").trim();
      const poi = String(row[1] || "").trim();
      if (theme) themes.add(theme);
      if (poi) pois.add(poi);
    });

    return {
      ok: true,
      themes: Array.from(themes).sort((a, b) => a.localeCompare(b)),
      pois: Array.from(pois).sort((a, b) => a.localeCompare(b))
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function ui_saveThemeRules(payload) {
  try {
    const edits = payload?.edits || [];
    if (!Array.isArray(edits) || !edits.length) {
      return { ok: true, saved: 0 };
    }

    const sheetName = (typeof repo_getControlSheetName_ === "function")
      ? repo_getControlSheetName_()
      : getControlSheetName_();
    const sh = (typeof repo_getSheetOrThrow_ === "function")
      ? repo_getSheetOrThrow_(sheetName)
      : getSpreadsheet_().getSheetByName(sheetName);

    if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);

    let saved = 0;
    edits.forEach(edit => {
      const rowIndex = Number(edit?.rowIndex);
      if (!rowIndex || rowIndex < 1) return;

      const theme = String(edit?.theme || "").trim();
      const poi = String(edit?.poi || "").trim();
      const keywords = String(edit?.keywords || "").trim();
      const active = edit?.active === true;

      sh.getRange(rowIndex, 1, 1, 4).setValues([[theme, poi, keywords, active]]);
      sh.getRange(rowIndex, 4, 1, 1).insertCheckboxes();
      saved += 1;
    });

    return { ok: true, saved };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function ui_addThemeRule() {
  try {
    const sheetName = (typeof repo_getControlSheetName_ === "function")
      ? repo_getControlSheetName_()
      : getControlSheetName_();
    const sh = (typeof repo_getSheetOrThrow_ === "function")
      ? repo_getSheetOrThrow_(sheetName)
      : getSpreadsheet_().getSheetByName(sheetName);

    if (!sh) throw new Error(`Sheet not found: "${sheetName}"`);

    const startRow =
      (typeof THEME_RULES_START_ROW !== "undefined" && Number(THEME_RULES_START_ROW) >= 1)
        ? Number(THEME_RULES_START_ROW)
        : 5;

    sh.insertRowBefore(startRow);
    sh.getRange(startRow, 1, 1, 4).setValues([["", "", "", false]]);
    sh.getRange(startRow, 4, 1, 1).insertCheckboxes();

    return {
      ok: true,
      row: {
        rowIndex: startRow,
        theme: "",
        poi: "",
        keywords: "",
        active: false
      }
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

/** ====== RSS Feeds (simple UI, previously UI_Feeds.gs) ====== */

// Model used for AI suggestions is controlled via Script Properties (AI_PROVIDER / AI_MODEL)

/** Return feeds for UI list */
function ui_getFeeds_v1() {
  const sh = feeds_ensureSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  const out = [];

  for (let i = 0; i < values.length; i++) {
    const r = values[i];
    const url = String(r[0] || "").trim();
    if (!url) continue;

    out.push({
      row: i + 2,
      url,
      active: r[1] === true,
      sourceName: String(r[2] || "").trim(),
      notes: String(r[3] || "").trim()
    });
  }
  return out;
}

/** Add a feed (deduped by normalized URL) */
function ui_addFeed_v1(feedUrl, sourceName, active, notes) {
  const sh = feeds_ensureSheet_();
  const url = String(feedUrl || "").trim();
  if (!url) throw new Error("Feed URL is required.");

  const normalized = feeds_normalizeLink_(url);
  if (!normalized) throw new Error("Invalid feed URL.");

  const existing = feeds_getExistingUrlSet_(sh);
  if (existing.has(normalized)) {
    return { ok: true, message: "Already exists (skipped).", url: normalized };
  }

  const row = sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 4).setValues([[
    url,
    active === true,
    String(sourceName || "").trim(),
    String(notes || "").trim()
  ]]);
  sh.getRange(row, 2, 1, 1).insertCheckboxes();

  return { ok: true, message: "Added.", url: normalized };
}

/** Toggle Active checkbox for a row */
function ui_setFeedActive_v1(row, isActive) {
  const sh = feeds_ensureSheet_();
  const r = Number(row);
  if (!r || r < 2) throw new Error("Invalid row.");
  sh.getRange(r, 2).setValue(isActive === true);
  return { ok: true };
}

/** Delete a feed row */
function ui_deleteFeed_v1(row) {
  const sh = feeds_ensureSheet_();
  const r = Number(row);
  if (!r || r < 2) throw new Error("Invalid row.");
  sh.deleteRow(r);
  return { ok: true };
}

/** ====== AI Settings ====== */

function ui_getAiSettings_v1() {
  const props = PropertiesService.getScriptProperties();
  const provider = String(props.getProperty("AI_PROVIDER") || "gemini").toLowerCase();
  const model = String(props.getProperty("AI_MODEL") || "");
  return { provider: provider === "openai" ? "openai" : "gemini", model };
}

function ui_setAiSettings_v1(provider, model) {
  const props = PropertiesService.getScriptProperties();
  const nextProvider = String(provider || "").toLowerCase() === "openai" ? "openai" : "gemini";
  const nextModel = String(model || "").trim();
  props.setProperty("AI_PROVIDER", nextProvider);
  if (nextModel) props.setProperty("AI_MODEL", nextModel);
  else props.deleteProperty("AI_MODEL");
  return { ok: true };
}

/**
 * Gemini-assisted suggestion:
 * - paste homepage URL (or any text prompt)
 * - returns list of feed URLs (strings)
 */
function ui_geminiSuggestFeeds_v1(homepageOrPrompt) {
  const q = String(homepageOrPrompt || "").trim();
  if (!q) throw new Error("Please paste a homepage URL (or a short prompt).");

  // First try deterministic discovery if available
  let deterministic = [];
  if (typeof discoverFeedsFromHomepage_ === "function") {
    try { deterministic = discoverFeedsFromHomepage_(q) || []; } catch (e) {}
  }
  if (deterministic && deterministic.length) {
    return Array.from(new Set(deterministic)).slice(0, 10);
  }

  // Require AIService helper
  if (typeof aiGenerateText_ !== "function") {
    throw new Error("Missing aiGenerateText_(). Add/keep AIService.gs.");
  }

  const prompt =
`You are helping a user find RSS/Atom feed URLs for a news site.

Input: ${q}

Return ONLY JSON (no markdown, no commentary) in this exact shape:
{"feeds":[ "https://example.com/rss.xml", "..."]}

Rules:
- feeds must be valid absolute URLs (start with https://)
- include RSS or Atom feeds if possible
- if uncertain, include a Google News RSS query for the site domain:
  https://news.google.com/rss/search?q=site:DOMAIN&hl=en-SG&gl=SG&ceid=SG:en
- return at most 10 feeds.`;

  const text = aiGenerateText_(prompt, {
    temperature: 0.2,
    maxOutputTokens: 512,
    responseMimeType: "application/json"
  });

  const parsed = feeds_safeParseJsonObject_(text);
  const feeds = (parsed && Array.isArray(parsed.feeds)) ? parsed.feeds : [];

  // Clean + normalize + unique
  const cleaned = [];
  const seen = new Set();
  feeds.forEach(u => {
    const raw = String(u || "").trim();
    const n = feeds_normalizeLink_(raw);
    if (!n) return;
    if (seen.has(n)) return;
    seen.add(n);
    cleaned.push(raw);
  });

  return cleaned.slice(0, 10);
}

/** ====== Internals (prefixed to avoid collisions) ====== */

function feeds_ensureSheet_() {
  const ss = getSpreadsheet_(); // web-app safe
  const sheetName = getFeedsSheetName_();

  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);

  // Ensure headers exist
  const header = sh.getRange(1, 1, 1, 4).getValues()[0];
  const want = ["Feed URL", "Active?", "Source Name", "Notes"];
  const mismatch = want.some((v, i) => String(header[i] || "").trim() !== v);

  if (mismatch) {
    sh.getRange(1, 1, 1, 4).setValues([want]).setFontWeight("bold");
  }

  // Ensure checkboxes in col B
  const maxRows = Math.max(sh.getMaxRows(), 200);
  sh.getRange(2, 2, maxRows - 1, 1).insertCheckboxes();

  return sh;
}

function feeds_getExistingUrlSet_(sh) {
  const lastRow = sh.getLastRow();
  const set = new Set();
  if (lastRow < 2) return set;

  const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  vals.forEach(r => {
    const u = String(r[0] || "").trim();
    if (!u) return;
    const n = feeds_normalizeLink_(u);
    if (n) set.add(n);
  });
  return set;
}

/**
 * Normalizer wrapper:
 * - If your project already has normalizeLink_(), we reuse it
 * - Else do a minimal safe normalization
 */
function feeds_normalizeLink_(url) {
  const s = String(url || "").trim();
  if (!s) return null;

  if (typeof normalizeLink_ === "function") {
    try { return normalizeLink_(s); } catch (e) {}
  }

  // minimal fallback
  try {
    const u = new URL(s);
    u.hash = "";
    return u.toString();
  } catch (e) {
    return null;
  }
}

function feeds_safeParseJsonObject_(text) {
  if (!text) return null;

  // direct parse
  try {
    const obj = JSON.parse(text);
    if (obj && typeof obj === "object") return obj;
  } catch (e) {}

  // extract first {...}
  const m = String(text).match(/\{[\s\S]*\}/);
  if (!m) return null;

  try {
    const obj = JSON.parse(m[0]);
    if (obj && typeof obj === "object") return obj;
  } catch (e) {}

  return null;
}

/** ====== Raw Articles (previously UI_RawArticles) ====== */

/**
 * Existing endpoint (kept) — paged + query search.
 * FIXES: paage typo -> page
 */
function ui_getRawArticles_v4(opts) {
  try {
    opts = opts || {};
    const page = Math.max(1, Number(opts.page || 1));
    const pageSize = Math.max(1, Math.min(500, Number(opts.pageSize || 50)));
    const query = String(opts.query || "").trim().toLowerCase();

    const ss = getSpreadsheet_();
    const sheetName = getRawArticlesSheetName_();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, message: `Sheet not found: "${sheetName}"` };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 1 || lastCol < 1) {
      return { ok: true, sheet: sheetName, page, pageSize, totalRows: 0, filteredRows: 0, headers: [], rows: [] };
    }

    // strings only
    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
      .map(s => String(s || "").trim());

    if (lastRow < 2) {
      return { ok: true, sheet: sheetName, page, pageSize, totalRows: 0, filteredRows: 0, headers, rows: [] };
    }

    // strings only
    const dataAll = sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues()
      .map(r => r.map(c => String(c || "")));

    // remove fully empty rows
    const data = dataAll.filter(r => r.some(c => c.trim() !== ""));
    const totalRows = data.length;

    let filtered = data;
    if (query) {
      filtered = data.filter(r => r.join(" | ").toLowerCase().includes(query));
    }

    const filteredRows = filtered.length;
    const maxPage = Math.max(1, Math.ceil(filteredRows / pageSize));
    const safePage = Math.min(page, maxPage); // ✅ FIXED
    const start = (safePage - 1) * pageSize;

    return {
      ok: true,
      sheet: sheetName,
      page: safePage,
      pageSize,
      totalRows,
      filteredRows,
      headers,
      rows: filtered.slice(start, start + pageSize)
    };

  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

/**
 * NEW endpoint for revamped UI:
 * - returns rows as objects (title/link/date/source/theme/poi/keywords)
 * - returns facets (themes/pois/sources/keywords) with counts
 * - supports your UI to do multi-filter + sort locally
 */
function ui_maybePruneRawArticlesByRetention_() {
  try {
    var days = (typeof INGEST_PRUNE_DAYS !== "undefined" && INGEST_PRUNE_DAYS) ? Number(INGEST_PRUNE_DAYS) : 14;
    if (!isFinite(days) || days <= 0) days = 14;
    days = Math.max(1, Math.round(days));

    var props = PropertiesService.getScriptProperties();
    var key = "RAW_ARTICLES_RETENTION_LAST_PRUNE_MS";
    var now = Date.now();
    var lastMs = Number(props.getProperty(key) || 0);
    var minIntervalMs = 60 * 60 * 1000; // at most once per hour

    if (lastMs && (now - lastMs) < minIntervalMs) {
      return { ok: true, skipped: true, maxAgeDays: days };
    }

    var stats = repo_pruneRawArticlesByAge_(days) || { removed: 0, kept: 0 };
    stats.maxAgeDays = days;
    props.setProperty(key, String(now));
    return { ok: true, skipped: false, prune: stats, maxAgeDays: days };
  } catch (err) {
    return { ok: false, message: err && err.message ? err.message : String(err) };
  }
}

function ui_getRawArticles_bootstrap_v1() {
  try {
    ui_maybePruneRawArticlesByRetention_();

    const ss = getSpreadsheet_();
    const sheetName = getRawArticlesSheetName_();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, message: `Sheet not found: "${sheetName}"` };

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();

    if (lastRow < 2 || lastCol < 1) {
      return {
        ok: true,
        sheet: sheetName,
        meta: { total: 0 },
        facets: { themes: [], pois: [], sources: [], keywords: [] },
        rows: []
      };
    }

    const range = sh.getRange(1, 1, lastRow, lastCol);
    const display = range.getDisplayValues().map(r => r.map(c => String(c || "").trim()));

    const headers = display[0];
    const data = display.slice(1).filter(r => r.some(c => c.trim() !== ""));

    const norm = (s) =>
      String(s || "").trim().toLowerCase()
        .replace(/\s+/g, " ")
        .replace(/[_\-]+/g, " ")
        .replace(/[^\w\s]/g, "");

    const headerMap = {};
    headers.forEach((h, i) => { const k = norm(h); if (k) headerMap[k] = i; });

    const findCol = (aliases) => {
      for (const a of aliases) {
        const k = norm(a);
        if (k in headerMap) return headerMap[k];
      }
      return -1;
    };

    const idx = {
      title:    findCol(["title"]),
      link:     findCol(["link", "url"]),
      date:     findCol(["date", "published", "published date"]),
      source:   findCol(["source"]),
      theme:    findCol(["theme"]),
      poi:      findCol(["point of interest", "poi", "point_of_interest"]),
      keywords: findCol(["matching keywords", "keywords", "matching_keyword"]),
      llmScore: findCol(["llm score", "llm_score"]),
      llmRecommendation: findCol(["llm recommendation", "llm_recommendation", "publish recommendation", "publish_recommendation"]),
      llmSummary: findCol(["llm summary", "llm_summary"]),
      llmReasons: findCol(["llm reasons", "llm_reasons", "score reasons", "score_reasons"])
    };

    const llmEntriesRaw = (typeof raw_readLlmRankSheet_ === "function") ? raw_readLlmRankSheet_() : [];
    const llmEntries = Array.isArray(llmEntriesRaw)
      ? llmEntriesRaw
      : (Array.isArray(llmEntriesRaw?.results) ? llmEntriesRaw.results : []);
    const llmByLink = new Map();
    const llmByTitle = new Map();

    llmEntries.forEach((item) => {
      const linkNorm = feeds_normalizeLink_(item?.url || item?.link);
      if (linkNorm) llmByLink.set(linkNorm, item?.llm || {});
      const titleKey = String(item?.title || "").trim().toLowerCase();
      if (titleKey) llmByTitle.set(titleKey, item?.llm || {});
    });

    const inc = (obj, key) => { if (!key) return; obj[key] = (obj[key] || 0) + 1; };

    const themeCounts = {};
    const poiCounts = {};
    const sourceCounts = {};
    const keywordCounts = {};

    const splitKeywords = (s) => String(s || "")
      .split(/[,;|]/g)
      .map(x => String(x || "").trim().toLowerCase())
      .filter(Boolean);

    const parseDateMs = (s) => {
      // your sheet sometimes wraps lines (e.g. "Tue, 20 Jan...\n2026 ... GMT")
      const cleaned = String(s || "").replace(/\s+/g, " ").trim();
      const t = Date.parse(cleaned);
      return isNaN(t) ? 0 : t;
    };

    const rows = data.map((r, i) => {
      const title = (idx.title >= 0) ? r[idx.title] : "";
      const link  = (idx.link  >= 0) ? r[idx.link]  : "";
      const dateS = (idx.date  >= 0) ? r[idx.date]  : "";
      const source= (idx.source>= 0) ? r[idx.source]: "";
      const theme = (idx.theme >= 0) ? r[idx.theme] : "";
      const poi   = (idx.poi   >= 0) ? r[idx.poi]   : "";
      const kwS   = (idx.keywords>=0) ? r[idx.keywords] : "";

      const llmScoreCell = (idx.llmScore >= 0) ? r[idx.llmScore] : "";
      const llmRecommendationCell = (idx.llmRecommendation >= 0) ? r[idx.llmRecommendation] : "";
      const llmSummaryCell = (idx.llmSummary >= 0) ? r[idx.llmSummary] : "";
      const llmReasonsCell = (idx.llmReasons >= 0) ? r[idx.llmReasons] : "";
      const llmScoreSheet = Number(llmScoreCell);
      const llmScoreSheetSafe = Number.isFinite(llmScoreSheet) ? llmScoreSheet : null;
      const llmRecommendationSheet = String(llmRecommendationCell || "").trim();
      const llmSummarySheet = String(llmSummaryCell || "").trim();
      const llmReasonsSheet = String(llmReasonsCell || "")
        .split(/[|\n•;]/g)
        .map((x) => String(x || "").trim())
        .filter(Boolean);

      const linkNorm = feeds_normalizeLink_(link);
      let llm = linkNorm ? llmByLink.get(linkNorm) : null;
      if (!llm && title) llm = llmByTitle.get(String(title).trim().toLowerCase());
      llm = llm || {};
      const llmScoreRaw = Number(llm.final_score);
      const llmScoreCache = Number.isFinite(llmScoreRaw) ? llmScoreRaw : null;
      const llmRecommendationCache = String(llm.publish_recommendation || llm.recommendation || "").trim();
      const llmSummaryCache = String(llm.summary || "").trim();
      const llmReasonsCache = Array.isArray(llm.score_reasons)
        ? llm.score_reasons.map((x) => String(x || "").trim()).filter(Boolean)
        : [];

      const llmScore = (llmScoreSheetSafe != null) ? llmScoreSheetSafe : llmScoreCache;
      const llmRecommendation = llmRecommendationSheet || llmRecommendationCache;
      const llmSummary = llmSummarySheet || llmSummaryCache;
      const llmReasons = llmReasonsSheet.length ? llmReasonsSheet : llmReasonsCache;
      const llmRecommended = (typeof raw_isLlmRecommended_ === "function")
        ? raw_isLlmRecommended_({ publish_recommendation: llmRecommendation, recommendation: llmRecommendation })
        : false;

      const kwList = splitKeywords(kwS);

      inc(themeCounts, theme);
      inc(poiCounts, poi);
      inc(sourceCounts, source);
      kwList.forEach(k => inc(keywordCounts, k));

      return {
        id: i + 2,
        title,
        link,
        dateDisplay: dateS,
        dateMs: parseDateMs(dateS),
        source,
        theme,
        poi,
        keywords: kwS,
        keywordsList: kwList,
        llmScore,
        llmRecommendation,
        llmRecommended,
        llmSummary,
        llmReasons,
        searchText: (title + " " + link + " " + source + " " + theme + " " + poi + " " + kwS + " " + llmRecommendation + " " + llmSummary + " " + llmReasons.join(" ")).toLowerCase()
      };
    });

    const toFacet = (countsObj) => Object.keys(countsObj)
      .map(k => ({ value: k, count: countsObj[k] }))
      .sort((a, b) => (b.count - a.count) || a.value.localeCompare(b.value));

    return {
      ok: true,
      sheet: sheetName,
      meta: { total: rows.length },
      facets: {
        themes: toFacet(themeCounts),
        pois: toFacet(poiCounts),
        sources: toFacet(sourceCounts),
        keywords: toFacet(keywordCounts)
      },
      rows
    };

  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

/** ====== RSS Feeds (table UI, previously UI_ThemeRules) ====== */

/*********************************
 * UI_RSSFeeds.gs — RSS Feeds UI API (collision-safe)
 * Sheet starts at row RSS_FEEDS_START_ROW, columns A:E by default:
 *   A: Source
 *   B: Feed URL
 *   C: Tags / Theme (optional)
 *   D: Notes (optional)
 *   E: Active (checkbox)
 *********************************/

// Fallbacks (won’t error if you haven’t defined constants elsewhere)
// Renamed to reduce collision risk across files
const RSS_UI_FEEDS_SHEET_FALLBACK = "RSS Feeds";
const RSS_UI_FEEDS_START_ROW_FALLBACK = 2;
const RSS_UI_FEEDS_COLS_FALLBACK = 5;

function rss_getRssFeedsSheetName_() {
  // If you already have a constant elsewhere, we’ll use it.
  if (typeof RSS_FEEDS_SHEET === "string" && RSS_FEEDS_SHEET.trim()) return RSS_FEEDS_SHEET.trim();
  if (typeof SC_RSS_SHEET === "string" && SC_RSS_SHEET.trim()) return SC_RSS_SHEET.trim();
  return RSS_UI_FEEDS_SHEET_FALLBACK;
}

function rss_getRssFeedsStartRow_() {
  if (typeof RSS_FEEDS_START_ROW === "number" && RSS_FEEDS_START_ROW > 0) return RSS_FEEDS_START_ROW;
  return RSS_UI_FEEDS_START_ROW_FALLBACK;
}

function rss_getRssFeedsCols_() {
  if (typeof RSS_FEEDS_COLS === "number" && RSS_FEEDS_COLS > 0) return RSS_FEEDS_COLS;
  return RSS_UI_FEEDS_COLS_FALLBACK;
}

function rss_getFeedsHeaderRow_(startRow) {
  const row = Number(startRow || 1);
  return row > 1 ? row - 1 : 1;
}

function rss_normalizeHeader_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^\w\s]/g, "")
    .replace(/\s+/g, " ");
}

function rss_buildFeedColumnMap_(header) {
  const map = {};
  const labels = {
    source: ["source", "source name"],
    url: ["feed url", "url", "rss url"],
    tags: ["tags", "tag", "theme", "themes"],
    notes: ["notes", "note"],
    active: ["active", "active", "enabled", "status"]
  };

  header.forEach((cell, idx) => {
    const key = rss_normalizeHeader_(cell);
    if (!key) return;
    Object.keys(labels).forEach(field => {
      if (map[field]) return;
      if (labels[field].includes(key)) {
        map[field] = idx + 1;
      }
    });
  });

  return {
    map,
    matches: Object.values(map).filter(Boolean).length
  };
}

function rss_guessFeedColumnMap_(rowValues, totalCols) {
  const cols = Math.max(4, Number(totalCols) || 0);
  const values = rowValues || [];
  const looksLikeUrl = (val) => /^https?:\/\//i.test(String(val || "").trim());
  const looksLikeBool = (val) => {
    if (val === true || val === false) return true;
    const s = String(val || "").trim().toLowerCase();
    return s === "true" || s === "false";
  };

  if (looksLikeUrl(values[0]) && looksLikeBool(values[1])) {
    return {
      url: 1,
      active: 2,
      source: cols >= 3 ? 3 : undefined,
      notes: cols >= 4 ? 4 : undefined,
      tags: cols >= 5 ? 5 : undefined
    };
  }

  return null;
}

function rss_getFeedColumnMap_(sh, startRow) {
  const headerRow = rss_getFeedsHeaderRow_(startRow);
  const lastCol = Math.max(5, sh.getLastColumn());
  const header = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0] || [];
  const headerMap = rss_buildFeedColumnMap_(header);

  let best = headerMap;
  let bestHeader = header;

  if (Number(startRow) !== headerRow) {
    const altHeader = sh.getRange(startRow, 1, 1, lastCol).getValues()[0] || [];
    const altMap = rss_buildFeedColumnMap_(altHeader);
    if (altMap.matches > best.matches) {
      best = altMap;
      bestHeader = altHeader;
    }
  }

  if (!best.matches) {
    const sampleRow = sh.getRange(startRow, 1, 1, lastCol).getValues()[0] || [];
    const guessed = rss_guessFeedColumnMap_(sampleRow, lastCol);
    if (guessed) return guessed;
    return {
      source: 1,
      url: 2,
      tags: 3,
      notes: 4,
      active: 5
    };
  }

  if (!best.map.active && bestHeader.length >= 5) {
    best.map.active = 5;
  }

  return best.map;
}

function rss_getFeedMaxCol_(map) {
  const cols = Object.values(map).filter(Boolean);
  return cols.length ? Math.max(...cols) : 5;
}

// Renamed to avoid collisions with other helperFunctions.gs etc.
function rss_clampInt_(n, min, max) {
  n = Number(n);
  if (!Number.isFinite(n)) n = min;
  n = Math.floor(n);
  return Math.max(min, Math.min(max, n));
}

// Renamed to avoid collisions with other helperFunctions.gs etc.
function rss_rowMatchesQuery_(vals, qLower) {
  if (!qLower) return true;
  for (let i = 0; i < vals.length; i++) {
    const s = String(vals[i] ?? "").toLowerCase();
    if (s.includes(qLower)) return true;
  }
  return false;
}

// Safe spreadsheet getter: prefer your shared getSpreadsheet_(), fallback to active
function rss_getSpreadsheetSafe_() {
  if (typeof getSpreadsheet_ === "function") return getSpreadsheet_();
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * GET RSS Feeds (paged + search)
 * params: { page: 1.., pageSize: 10/25/50/100, q: "search" }
 */
function ui_getRssFeeds(params) {
  try {
    params = params || {};
    const q = String(params.q || "").trim();
    const qLower = q.toLowerCase();

    const pageSize = rss_clampInt_(params.pageSize || 50, 1, 500);
    const page = rss_clampInt_(params.page || 1, 1, 999999);

    const ss = rss_getSpreadsheetSafe_();
    const sheetName = rss_getRssFeedsSheetName_();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, message: `Sheet not found: "${sheetName}"` };

    const startRow = rss_getRssFeedsStartRow_();
    const lastRow = sh.getLastRow();
    if (lastRow < startRow) {
      return { ok: true, sheet: sheetName, page, pageSize, totalRows: 0, totalMatches: 0, rows: [] };
    }

    const columnMap = rss_getFeedColumnMap_(sh, startRow);
    const cols = rss_getFeedMaxCol_(columnMap) || rss_getRssFeedsCols_();
    const height = lastRow - startRow + 1;

    const values = sh.getRange(startRow, 1, height, cols).getValues();

    // Build row objects + filter blanks
    const allRows = values
      .map((r, idx) => {
        const activeRaw = columnMap.active ? r[columnMap.active - 1] : false;
        const active =
          activeRaw === true ||
          activeRaw === 1 ||
          String(activeRaw).trim().toLowerCase() === "true";

        return {
          rowIndex: startRow + idx,
          source: columnMap.source ? String(r[columnMap.source - 1] || "").trim() : "",
          url: columnMap.url ? String(r[columnMap.url - 1] || "").trim() : "",
          tags: columnMap.tags ? String(r[columnMap.tags - 1] || "").trim() : "",
          notes: columnMap.notes ? String(r[columnMap.notes - 1] || "").trim() : "",
          active
        };
      })
      .filter(r => r.source || r.url || r.tags || r.notes);

    // Search filter
    const filtered = allRows.filter(r =>
      rss_rowMatchesQuery_([r.source, r.url, r.tags, r.notes, r.active], qLower)
    );

    const totalMatches = filtered.length;

    // Paging
    const startIdx = (page - 1) * pageSize;
    const pageRows = filtered.slice(startIdx, startIdx + pageSize);

    return {
      ok: true,
      sheet: sheetName,
      page,
      pageSize,
      totalRows: allRows.length,
      totalMatches,
      rows: pageRows
    };

  } catch (err) {
    return { ok: false, message: err?.message || String(err), stack: err?.stack || "" };
  }
}

/**
 * SAVE edits
 * payload: { edits: [{rowIndex, source, url, tags, notes, active}] }
 */
function ui_saveRssFeeds(payload) {
  try {
    payload = payload || {};
    const edits = Array.isArray(payload.edits) ? payload.edits : [];
    if (!edits.length) return { ok: true, saved: 0 };

    const ss = rss_getSpreadsheetSafe_();
    const sheetName = rss_getRssFeedsSheetName_();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, message: `Sheet not found: "${sheetName}"` };

    const startRow = rss_getRssFeedsStartRow_();
    const columnMap = rss_getFeedColumnMap_(sh, startRow);

    edits.forEach(e => {
      const rowIndex = Number(e.rowIndex);
      if (!rowIndex || rowIndex < startRow) return;

      if (columnMap.source) sh.getRange(rowIndex, columnMap.source).setValue(String(e.source || "").trim());
      if (columnMap.url) sh.getRange(rowIndex, columnMap.url).setValue(String(e.url || "").trim());
      if (columnMap.tags) sh.getRange(rowIndex, columnMap.tags).setValue(String(e.tags || "").trim());
      if (columnMap.notes) sh.getRange(rowIndex, columnMap.notes).setValue(String(e.notes || "").trim());
      if (columnMap.active) {
        const cell = sh.getRange(rowIndex, columnMap.active);
        cell.setValue(e.active === true);
        cell.insertCheckboxes();
      }
    });

    return { ok: true, saved: edits.length };
  } catch (err) {
    return { ok: false, message: err?.message || String(err), stack: err?.stack || "" };
  }
}

/**
 * ADD new RSS feed row (copies formatting + data validation from template row)
 */
function ui_addRssFeed() {
  try {
    const ss = rss_getSpreadsheetSafe_();
    const sheetName = rss_getRssFeedsSheetName_();
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return { ok: false, message: `Sheet not found: "${sheetName}"` };

    const startRow = rss_getRssFeedsStartRow_();
    const columnMap = rss_getFeedColumnMap_(sh, startRow);
    const maxCol = rss_getFeedMaxCol_(columnMap) || rss_getRssFeedsCols_();
    const lastRow = Math.max(sh.getLastRow(), startRow - 1);

    sh.insertRowAfter(lastRow);
    const newRow = lastRow + 1;

    const template = sh.getRange(startRow, 1, 1, maxCol);
    const dest = sh.getRange(newRow, 1, 1, maxCol);

    template.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    template.copyTo(dest, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

    dest.clearContent();
    if (columnMap.active) {
      const cell = sh.getRange(newRow, columnMap.active);
      cell.setValue(true);
      cell.insertCheckboxes();
    }

    return {
      ok: true,
      row: { rowIndex: newRow, source: "", url: "", tags: "", notes: "", active: true }
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err), stack: err?.stack || "" };
  }
}

/**
 * CREATE RSS feeds from a homepage URL
 * payload: { url: "https://site.com", sourceName?: "BBC" }
 */
function ui_createRssFeedFromUrl(payload) {
  try {
    payload = payload || {};
    const url = String(payload.url || "").trim();
    if (!url) return { ok: false, message: "Missing URL." };

    const providedSource = String(payload.sourceName || "").trim();

    const ss = rss_getSpreadsheetSafe_();
    const sheetName = rss_getRssFeedsSheetName_();
    const feedsSh = ss.getSheetByName(sheetName);
    if (!feedsSh) return { ok: false, message: `Sheet not found: "${sheetName}"` };

    const existing = loadExistingFeedUrls_(feedsSh);

    let html = "";
    let warning = "";
    try {
      html = fetchHtml_(url);
    } catch (e) {
      warning = `Homepage fetch failed (${e.message || e}). Using Google News fallback.`;
    }

    let finalFeeds = [];
    if (html) {
      const candidates = extractFeedCandidatesFromHtml_(url, html);

      let geminiSuggestions = [];
      try {
        geminiSuggestions = geminiSuggestFeeds_(url, html, candidates);
      } catch (e) {
        geminiSuggestions = [];
      }

      const combined = dedupeList_([].concat(candidates, geminiSuggestions));

      if (typeof VERIFY_FEED_FETCH !== "undefined" && VERIFY_FEED_FETCH) {
        finalFeeds = combined
          .filter(u => looksLikeFeedUrl_(u))
          .filter(u => verifyFeedUrl_(u));
      } else {
        finalFeeds = combined.filter(u => looksLikeFeedUrl_(u));
      }
    }

    if (!finalFeeds.length) {
      finalFeeds = [buildGoogleNewsSiteRss_(url)];
    }

    const sourceName = providedSource || deriveSourceName_(url);
    const added = appendFeeds_(feedsSh, finalFeeds, existing, sourceName, "");

    return {
      ok: true,
      addedCount: added.length,
      added,
      feeds: finalFeeds,
      sourceName,
      warning
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err), stack: err?.stack || "" };
  }
}
