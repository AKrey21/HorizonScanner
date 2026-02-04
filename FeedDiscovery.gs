/*********************************
 * FEED DISCOVERY (AI-assisted)
 *
 * UX goal:
 * - Non-tech user pastes a homepage URL in "Feed Discovery"
 * - Clicks a big button (Drawing) assigned to DISCOVER_RSS_BUTTON()
 * - Script discovers RSS/Atom feeds (or Google News fallback) and appends to "RSS Feeds"
 *
 * Sheets:
 * - "Feed Discovery" (input)
 *   A: Homepage URL
 *   B: Source Name (optional)
 *   C: Notes (optional)
 *   D: Status (output)
 *   E: Found Feeds (output)
 *
 * - "RSS Feeds" (output)
 *   A: Feed URL
 *   B: Active? (checkbox)
 *   C: Source Name
 *   D: Notes
 *********************************/

const FEED_DISCOVERY_SHEET = "Feed Discovery";

// AI config is controlled via Script Properties (AI_PROVIDER / AI_MODEL)
// Behavior
const MAX_HTML_CHARS_SENT_TO_AI = 12000; // keep prompts bounded
const VERIFY_FEED_FETCH = true;              // verify candidates are real RSS/Atom
const AUTO_ACTIVATE_NEW_FEEDS = true;        // new feeds Active? default
const CLEAR_DISCOVERY_OUTPUT_BEFORE_RUN = false; // if true, clears D:E before processing

/**************
 * MENU + BUTTON ENTRYPOINTS
 **************/
/**
 * Assign your Drawing button to this function name (Excel-like UX).
 * In Sheets: click drawing → ⋮ → Assign script → DISCOVER_RSS_BUTTON
 */
function DISCOVER_RSS_BUTTON() {
  discoverRssFeedsFromHomepages();
}

/**************
 * ONE-CLICK SETUP (optional but recommended)
 * - Creates/repairs the "Feed Discovery" sheet
 * - Adds headers + instructional text
 * - Adds some formatting for readability
 *
 * NOTE: Google Sheets does NOT allow Apps Script to create/assign a Drawing button directly.
 * You still insert a Drawing manually and assign script name: DISCOVER_RSS_BUTTON
 **************/
function setupFeedDiscoverySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FEED_DISCOVERY_SHEET);
  if (!sh) sh = ss.insertSheet(FEED_DISCOVERY_SHEET);

  // Layout
  sh.clear();

  // Title + instructions
  sh.getRange("A1").setValue("Feed Discovery (Paste homepage URLs → click button)");
  sh.getRange("A2").setValue("How to use:");
  sh.getRange("A3").setValue("1) Paste a news site homepage URL in column A (e.g. https://www.forbes.com/)");
  sh.getRange("A4").setValue("2) Click the button: “Discover RSS Feeds” (assign script: DISCOVER_RSS_BUTTON)");
  sh.getRange("A5").setValue("3) Status appears in column D. Found feeds appear in column E and are appended to 'RSS Feeds'.");

  // Headers at row 7
  sh.getRange("A7").setValue("Homepage URL");
  sh.getRange("B7").setValue("Source Name (optional)");
  sh.getRange("C7").setValue("Notes (optional)");
  sh.getRange("D7").setValue("Status");
  sh.getRange("E7").setValue("Found Feeds");

  // Formatting
  sh.getRange("A1:E1").merge();
  sh.getRange("A1").setFontSize(14).setFontWeight("bold");
  sh.getRange("A2:A5").setFontSize(10);
  sh.getRange("A7:E7").setFontWeight("bold");
  sh.setFrozenRows(7);

  sh.setColumnWidth(1, 420); // Homepage URL
  sh.setColumnWidth(2, 220); // Source
  sh.setColumnWidth(3, 260); // Notes
  sh.setColumnWidth(4, 260); // Status
  sh.setColumnWidth(5, 520); // Found feeds

  sh.getRange("A8:E200").setWrap(true).setVerticalAlignment("top");
  sh.getRange("D8:E200").setBackground("#f7f7f7");

  SpreadsheetApp.getUi().alert(
    "Feed Discovery sheet is set up.\n\nNext: Insert → Drawing, create a button, then Assign script = DISCOVER_RSS_BUTTON"
  );
}

/**************
 * MAIN
 **************/

function discoverRssFeedsFromHomepages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(FEED_DISCOVERY_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${FEED_DISCOVERY_SHEET}`);

  const feedsSh = ss.getSheetByName(FEEDS_SHEET);
  if (!feedsSh) throw new Error(`Sheet not found: ${FEEDS_SHEET}`);

  // Support both layouts:
  // - If user ran setupFeedDiscoverySheet(): headers on row 7 and data starts row 8
  // - Otherwise: assume headers on row 1 and data starts row 2
  const headerRow = detectHeaderRow_(sh);
  const startRow = headerRow + 1;

  const lastRow = sh.getLastRow();
  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert(`No homepage URLs found. Add URLs in "${FEED_DISCOVERY_SHEET}" column A.`);
    return;
  }

  if (CLEAR_DISCOVERY_OUTPUT_BEFORE_RUN) {
    sh.getRange(startRow, 4, Math.max(0, lastRow - startRow + 1), 2).clearContent(); // D:E
  }

  // Load existing feed URLs to avoid duplicates
  const existing = loadExistingFeedUrls_(feedsSh);

  // Read input rows A:E
  const values = sh.getRange(startRow, 1, lastRow - startRow + 1, 5).getValues();

  // Lock to prevent accidental double-runs
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert("Another run is in progress. Try again in a moment.");
    return;
  }

  try {
    for (let i = 0; i < values.length; i++) {
      const rowNum = startRow + i;
      const homepage = (values[i][0] || "").toString().trim();
      const sourceName = (values[i][1] || "").toString().trim();
      const notes = (values[i][2] || "").toString().trim();

      if (!homepage) {
        sh.getRange(rowNum, 4).setValue("(skipped) empty homepage");
        sh.getRange(rowNum, 5).setValue("");
        continue;
      }

      // UX: show progress in sheet
      sh.getRange(rowNum, 4).setValue("Fetching homepage…");
      sh.getRange(rowNum, 5).setValue("");
      SpreadsheetApp.flush();

      let html;
      try {
        html = fetchHtml_(homepage);
      } catch (e) {
        sh.getRange(rowNum, 4).setValue(`ERROR — homepage fetch failed: ${e.message || e}`);
        continue;
      }

      // Deterministic discovery first
      const candidates = extractFeedCandidatesFromHtml_(homepage, html);

      sh.getRange(rowNum, 4).setValue("Asking Gemini…");
      SpreadsheetApp.flush();

      let geminiSuggestions = [];
      try {
        geminiSuggestions = geminiSuggestFeeds_(homepage, html, candidates);
      } catch (e) {
        // If Gemini fails, still proceed with extracted candidates + fallback
        sh.getRange(rowNum, 4).setValue(`Gemini failed, using extracted/fallback…`);
        geminiSuggestions = [];
      }

      // Combine & de-dupe (preserve order)
      const combined = dedupeList_([].concat(candidates, geminiSuggestions));

      // Verify candidates are real RSS/Atom feeds
      sh.getRange(rowNum, 4).setValue(VERIFY_FEED_FETCH ? "Verifying feeds…" : "Finalizing…");
      SpreadsheetApp.flush();

      let finalFeeds = combined;

      if (VERIFY_FEED_FETCH) {
        finalFeeds = combined
          .filter(u => looksLikeFeedUrl_(u))
          .filter(u => verifyFeedUrl_(u));
      } else {
        finalFeeds = combined.filter(u => looksLikeFeedUrl_(u));
      }

      // If none, fallback to Google News site:domain RSS
      if (!finalFeeds.length) {
        finalFeeds = [buildGoogleNewsSiteRss_(homepage)];
      }

      // Write to RSS Feeds sheet
      const added = appendFeeds_(
        feedsSh,
        finalFeeds,
        existing,
        sourceName || deriveSourceName_(homepage),
        notes
      );

      const statusMsg = added.length
        ? `OK — added ${added.length} feed(s)`
        : `OK — no new feeds (already existed)`;

      sh.getRange(rowNum, 4).setValue(statusMsg);
      sh.getRange(rowNum, 5).setValue(finalFeeds.join("\n"));

      // Update in-memory existing set
      added.forEach(u => existing.add(normalizeUrlKey_(u)));
    }

    SpreadsheetApp.getUi().alert("Feed discovery complete. Check 'Feed Discovery' and 'RSS Feeds'.");
  } finally {
    lock.releaseLock();
  }
}

/***********************
 * Layout detection
 ***********************/
function detectHeaderRow_(sh) {
  // If setupFeedDiscoverySheet() was used, headers are on row 7 with "Homepage URL"
  const v7 = (sh.getRange("A7").getValue() || "").toString().trim().toLowerCase();
  if (v7 === "homepage url") return 7;

  // Otherwise assume row 1 headers
  return 1;
}

/***********************
 * Core helpers
 ***********************/

function fetchHtml_(url) {
  const res = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0",
      "Accept": "text/html,application/xhtml+xml"
    }
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error(`HTTP ${code}`);
  return res.getContentText() || "";
}

function extractFeedCandidatesFromHtml_(baseUrl, html) {
  if (!html) return [];

  const out = [];

  // <link rel="alternate" type="application/rss+xml" href="...">
  const linkTagRe = /<link\b[^>]*>/gi;
  const tags = html.match(linkTagRe) || [];
  tags.forEach(tag => {
    const type = getAttr_(tag, "type");
    const rel = getAttr_(tag, "rel");
    const href = getAttr_(tag, "href");
    if (!href) return;

    const t = (type || "").toLowerCase();
    const r = (rel || "").toLowerCase();

    if (r.includes("alternate") && (t.includes("rss") || t.includes("atom") || t.includes("xml"))) {
      out.push(toAbsoluteUrl_(baseUrl, href));
    }
  });

  // Any href containing rss/feed/atom/xml
  const aHrefRe = /href\s*=\s*["']([^"']+)["']/gi;
  let m;
  while ((m = aHrefRe.exec(html)) !== null) {
    const href = (m[1] || "").trim();
    if (!href) continue;

    const lower = href.toLowerCase();
    if (
      lower.includes("rss") ||
      lower.includes("atom") ||
      lower.includes("/feed") ||
      lower.endsWith(".xml")
    ) {
      out.push(toAbsoluteUrl_(baseUrl, href));
    }
  }

  return dedupeList_(out)
    .filter(u => looksLikeFeedUrl_(u))
    .slice(0, 30);
}

function looksLikeFeedUrl_(url) {
  if (!url) return false;
  const u = url.toLowerCase();
  if (u.endsWith(".png") || u.endsWith(".jpg") || u.endsWith(".css") || u.endsWith(".js")) return false;
  return u.includes("rss") || u.includes("atom") || u.includes("/feed") || u.endsWith(".xml");
}

function verifyFeedUrl_(url) {
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "User-Agent": "Mozilla/5.0", "Accept": "application/xml,text/xml,application/rss+xml,*/*" }
    });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) return false;

    const txt = (res.getContentText() || "").trim().slice(0, 800).toLowerCase();
    return txt.includes("<rss") || txt.includes("<feed") || txt.includes("<rdf:rdf") || txt.includes("<channel");
  } catch (e) {
    return false;
  }
}

function appendFeeds_(feedsSh, feedUrls, existingSet, sourceName, notes) {
  const added = [];
  const rows = [];

  feedUrls.forEach(u => {
    const key = normalizeUrlKey_(u);
    if (existingSet.has(key)) return;

    rows.push([
      u,
      AUTO_ACTIVATE_NEW_FEEDS,
      sourceName,
      notes ? `${notes} (auto-discovered)` : "auto-discovered"
    ]);
    added.push(u);
  });

  if (rows.length) {
    feedsSh.getRange(feedsSh.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }

  return added;
}

function loadExistingFeedUrls_(feedsSh) {
  const lastRow = feedsSh.getLastRow();
  if (lastRow < 2) return new Set();

  const vals = feedsSh.getRange(2, 1, lastRow - 1, 1).getValues(); // col A
  const set = new Set();
  vals.forEach(r => {
    const u = (r[0] || "").toString().trim();
    if (u) set.add(normalizeUrlKey_(u));
  });
  return set;
}

function normalizeUrlKey_(url) {
  return String(url || "")
    .trim()
    .replace(/^http:\/\//i, "https://")
    .replace(/\/+$/, "");
}

function dedupeList_(arr) {
  const seen = new Set();
  const out = [];
  (arr || []).forEach(x => {
    const k = normalizeUrlKey_(x);
    if (!k || seen.has(k)) return;
    seen.add(k);
    out.push(x);
  });
  return out;
}

function toAbsoluteUrl_(baseUrl, maybeRelative) {
  try {
    return new URL(maybeRelative, baseUrl).toString();
  } catch (e) {
    return maybeRelative;
  }
}

function getAttr_(tag, attrName) {
  const re = new RegExp(attrName + "\\s*=\\s*[\"']([^\"']+)[\"']", "i");
  const m = tag.match(re);
  return m ? m[1] : "";
}

function deriveSourceName_(homepage) {
  try {
    const u = new URL(homepage);
    return u.hostname.replace(/^www\./i, "");
  } catch (e) {
    return "Unknown";
  }
}

function buildGoogleNewsSiteRss_(homepage) {
  let host = "";
  try { host = new URL(homepage).hostname.replace(/^www\./i, ""); } catch (e) {}
  const q = host ? `site:${host}` : homepage;
  return `https://news.google.com/rss/search?q=${encodeURIComponent(q)}&hl=en-SG&gl=SG&ceid=SG:en`;
}

/***********************
 * AI call
 ***********************/

function geminiSuggestFeeds_(homepage, html, extractedCandidates) {
  const snippet = (html || "").slice(0, MAX_HTML_CHARS_SENT_TO_AI);

  const prompt =
`You are helping a non-technical user set up RSS feeds for a news site.

Homepage: ${homepage}

We already extracted these candidate URLs from the HTML (may include duplicates or non-feeds):
${(extractedCandidates || []).slice(0, 30).join("\n")}

Homepage HTML snippet (may help you find RSS/Atom <link> tags):
${snippet}

Task:
Return ONLY valid RSS or Atom feed URLs for this site, as a JSON array of strings.
- Prefer official/native feeds from this site.
- Include at most 10 feed URLs.
- If you cannot find any native RSS/Atom feeds, return ONE Google News RSS fallback of the form:
  https://news.google.com/rss/search?q=site:DOMAIN&hl=en-SG&gl=SG&ceid=SG:en
- Do not include explanations. JSON only.`;

  const text = aiGenerateText_(prompt, {
    temperature: 0.2,
    maxOutputTokens: 800,
    responseMimeType: "application/json"
  }).trim();
  if (!text) return [];

  // Expect JSON array of strings
  let arr;
  try {
    arr = JSON.parse(text);
  } catch (e) {
    const cleaned = text.replace(/```json|```/gi, "").trim();
    arr = JSON.parse(cleaned);
  }

  if (!Array.isArray(arr)) return [];

  return arr
    .map(s => (s || "").toString().trim())
    .filter(Boolean)
    .slice(0, 20);
}
