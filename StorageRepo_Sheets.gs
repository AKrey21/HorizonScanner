/*********************************
 * StorageRepo_Sheets.gs — All Sheets IO lives here (SAFE)
 * - Avoids ReferenceError if constants not defined
 * - Avoids auto-creating sheets on read paths (prevents web-app insert failures)
 * - Fixes RSS Feeds column mapping to A:Source B:URL C:Tags D:Notes E:Active
 *********************************/

// --- Safe getters (use your UI_Core getters if present) ---
function repo_getRawSheetName_() {
  if (typeof getRawArticlesSheetName_ === "function") return getRawArticlesSheetName_();
  try { if (typeof RAW_SHEET !== "undefined" && RAW_SHEET) return RAW_SHEET; } catch (e) {}
  return "Raw Articles";
}

function repo_getFeedsSheetName_() {
  if (typeof getFeedsSheetName_ === "function") return getFeedsSheetName_();
  try { if (typeof FEEDS_SHEET !== "undefined" && FEEDS_SHEET) return FEEDS_SHEET; } catch (e) {}
  return "RSS Feeds";
}

function repo_getControlSheetName_() {
  if (typeof getControlSheetName_ === "function") return getControlSheetName_();
  try { if (typeof CONTROL_SHEET !== "undefined" && CONTROL_SHEET) return CONTROL_SHEET; } catch (e) {}
  return "ThemeRules";
}

// --- Sheet access helpers ---
function repo_getSheetOrThrow_(name) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(name);
  if (!sh) {
    const available = ss.getSheets().map(s => s.getName()).join(", ");
    throw new Error(`Sheet not found: "${name}". Available: ${available}`);
  }
  return sh;
}

function repo_getOrCreateSheet_(name) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// --- RAW ARTICLES ---
function repo_resetRawArticles_() {
  const sh = repo_getOrCreateSheet_(repo_getRawSheetName_());
  sh.clearContents();

  sh.appendRow([
    "Title", "Link", "Date", "Source",
    "Theme", "Point of Interest",
    "Matching Keywords"
  ]);

  return sh;
}

function repo_appendRawArticles_(rows) {
  if (!rows || !rows.length) return 0;
  const sh = repo_getOrCreateSheet_(repo_getRawSheetName_());
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  return rows.length;
}

// --- RSS FEEDS (A:Source B:URL C:Tags D:Notes E:Active) ---
function repo_getActiveFeeds_() {
  const sh = repo_getSheetOrThrow_(repo_getFeedsSheetName_());
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const columnMap = repo_getFeedsColumnMap_(sh);
  const maxCol = repo_getFeedsMaxCol_(columnMap);
  const values = sh.getRange(2, 1, lastRow - 1, maxCol).getValues();

  return values
    .map(r => {
      const source = columnMap.source ? String(r[columnMap.source - 1] || "").trim() : "";
      const url    = columnMap.url ? String(r[columnMap.url - 1] || "").trim() : "";
      const tags   = columnMap.tags ? String(r[columnMap.tags - 1] || "").trim() : "";
      const notes  = columnMap.notes ? String(r[columnMap.notes - 1] || "").trim() : "";
      const activeRaw = columnMap.active ? r[columnMap.active - 1] : false;
      const active =
        activeRaw === true || activeRaw === 1 || String(activeRaw).trim().toLowerCase() === "true";

      return { url, active, sourceName: source, tags, notes };
    })
    .filter(x => x.active && x.url);
}

function repo_normalizeFeedHeader_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^\w\s]/g, "")
    .replace(/\s+/g, " ");
}

function repo_getFeedsColumnMap_(sh) {
  const lastCol = Math.max(4, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const map = {};
  const labels = {
    source: ["source", "source name"],
    url: ["feed url", "url", "rss url"],
    tags: ["tags", "tag", "theme", "themes"],
    notes: ["notes", "note"],
    active: ["active", "enabled", "status"]
  };

  header.forEach((cell, idx) => {
    const key = repo_normalizeFeedHeader_(cell);
    if (!key) return;
    Object.keys(labels).forEach(field => {
      if (map[field]) return;
      if (labels[field].includes(key)) {
        map[field] = idx + 1;
      }
    });
  });

  if (!Object.keys(map).length) {
    return { source: 1, url: 2, tags: 3, notes: 4, active: 5 };
  }

  return map;
}

function repo_getFeedsMaxCol_(map) {
  const cols = Object.values(map).filter(Boolean);
  return cols.length ? Math.max(...cols) : 5;
}

// --- THEME RULES (A:Theme B:POI C:Keywords D:Active) ---
function repo_getActiveThemeRules_() {
  const sh = repo_getSheetOrThrow_(repo_getControlSheetName_());
  const lastRow = sh.getLastRow();

  const startRow =
    (typeof THEME_RULES_START_ROW !== "undefined" && Number(THEME_RULES_START_ROW) >= 1)
      ? Number(THEME_RULES_START_ROW)
      : 5;

  if (lastRow < startRow) return [];

  const rows = sh.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();

  return rows
    .filter(r => {
      const theme = String(r[0] || "").trim();
      const keys  = String(r[2] || "").trim();
      const activeRaw = r[3];
      const isOn =
        activeRaw === true || activeRaw === 1 || String(activeRaw).trim().toLowerCase() === "true";

      // POI can be optional; don’t require it
      return isOn && theme && keys;
    })
    .map(r => {
      const theme = String(r[0] || "").trim();
      const poi   = String(r[1] || "").trim();

      const keywords = String(r[2] || "")
        .split(",")
        .map(s => s.trim().toLowerCase())
        .filter(Boolean);

      // Build regexes only if helper exists
      const keywordRegexes = (typeof ingest_buildKeywordRegex_ === "function")
        ? keywords
            .map(k => ({ k, re: ingest_buildKeywordRegex_(k) }))
            .filter(x => x.re)
        : [];

      return { theme, poi, keywords, keywordRegexes };
    });
}
