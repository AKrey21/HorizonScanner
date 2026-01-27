/*********************************
 * UI_Feeds.gs — RSS Feeds UI Backend (for RSSFeeds.html)
 * Sheet: "RSS Feeds" (A:D)
 * A: Feed URL
 * B: Active? (checkbox)
 * C: Source Name
 * D: Notes
 *********************************/

// Model used for Gemini suggestions
const GEMINI_MODEL_ID = "gemini-2.5-flash";

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

  // Require AIService gemini helper
  if (typeof geminiGenerateText_ !== "function") {
    throw new Error("Missing geminiGenerateText_(). Add/keep AIService.gs.");
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

  const text = geminiGenerateText_(prompt, {
    model: GEMINI_MODEL_ID,
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
