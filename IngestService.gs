/*********************************
 * IngestService.gs — Rules-only RSS ingest (SELF-CONTAINED, NO TDZ)
 * - Web-app safe (openById fallback)
 * - Includes repo_* functions + keyword regex builder
 *********************************/

/* ========= SAFE CONFIG (NO SHADOWING) ========= */
var ING_CFG = (function () {
  // If other files define these globals, we reuse them; otherwise fall back.
  var controlSheet = (typeof CONTROL_SHEET !== "undefined") ? CONTROL_SHEET : "ThemeRules";
  var feedsSheet   = (typeof FEEDS_SHEET   !== "undefined") ? FEEDS_SHEET   : "RSS Feeds";
  var rawSheet     = (typeof RAW_SHEET     !== "undefined") ? RAW_SHEET     : "Raw Articles";

  var rulesStartRow = (typeof THEME_RULES_START_ROW !== "undefined") ? THEME_RULES_START_ROW : 5;

  // If your project defines INGEST_ALL_ARTICLES elsewhere, reuse it.
  var ingestAll = (typeof INGEST_ALL_ARTICLES !== "undefined") ? INGEST_ALL_ARTICLES : false;

  // If your UI.gs defines SPREADSHEET_ID, reuse it.
  var spreadsheetId = (typeof SPREADSHEET_ID !== "undefined") ? SPREADSHEET_ID : "";

  return {
    CONTROL_SHEET: controlSheet,
    FEEDS_SHEET: feedsSheet,
    RAW_SHEET: rawSheet,
    THEME_RULES_START_ROW: rulesStartRow,
    INGEST_ALL_ARTICLES: ingestAll,
    SPREADSHEET_ID: spreadsheetId
  };
})();

/* ========= UI ENDPOINTS ========= */

// Always return OBJECTS from UI endpoints
function ui_importRSS_7days()  { return ui_importRSS_7days_v2(); }
function ui_importRSS_30days() { return ui_importRSS_30days_v2(); }
function ui_importRSS_all()    { return ui_importRSS_all_v2(); }

// Object endpoints (what your UI expects)
function ui_importRSS_7days_v2()  {
  var res = ingest_importRSS_(7);
  res._sig = "IngestService.ui_importRSS_7days_v2 @ 2026-01-13";
  return res;
}
function ui_importRSS_30days_v2() {
  var res = ingest_importRSS_(30);
  res._sig = "IngestService.ui_importRSS_30days_v2 @ 2026-01-13";
  return res;
}
function ui_importRSS_all_v2()    {
  var res = ingest_importRSS_(null);
  res._sig = "IngestService.ui_importRSS_all_v2 @ 2026-01-13";
  return res;
}

function ui_importRSS_ping() {
  return { ok:true, _sig:"IngestService.ping @ 2026-01-13", now: new Date().toISOString() };
}

/* ========= MAIN INGEST ========= */

/**
 * Main ingest runner
 * Returns: { ok, message, stats, errors[], debug? }
 */
function ingest_importRSS_(daysBack) {
  var t0 = Date.now();

  var stats = {
    daysBack: daysBack,
    feedsActive: 0,
    feedsProcessed: 0,
    feedErrors: 0,
    entriesSeen: 0,
    imported: 0,
    skippedDuplicate: 0,
    skippedDate: 0,
    skippedNoMatch: 0,
    skippedMissingFields: 0
  };

  var errors = [];

  try {
    repo_resetRawArticles_();

    var themeRules = repo_getActiveThemeRules_();
    var feeds = repo_getActiveFeeds_();
    stats.feedsActive = feeds.length;

    var debug = {
      spreadsheetIdUsed: ING_CFG.SPREADSHEET_ID || "(active)",
      controlSheet: ING_CFG.CONTROL_SHEET,
      feedsSheet: ING_CFG.FEEDS_SHEET,
      rawSheet: ING_CFG.RAW_SHEET,
      themeRulesActive: Array.isArray(themeRules) ? themeRules.length : -1,
      feedsActive: Array.isArray(feeds) ? feeds.length : -1,
      firstFeed: feeds && feeds[0] ? (feeds[0].url || "") : "",
      ingestAll: ingest_shouldIngestAll_()
    };

    if (!feeds.length) {
      return { ok:false, message:'No active feeds found in "' + ING_CFG.FEEDS_SHEET + '".', stats:stats, errors:errors, debug:debug };
    }

    if (!themeRules.length && !ingest_shouldIngestAll_()) {
      return { ok:false, message:'No active ThemeRules found in "' + ING_CFG.CONTROL_SHEET + '" (and INGEST_ALL_ARTICLES is false).', stats:stats, errors:errors, debug:debug };
    }

    var cutoff = null;
    if (daysBack !== null) {
      cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - Number(daysBack));
    }

    var seenLinks = new Set();

    feeds.forEach(function (feedObj) {
      stats.feedsProcessed++;

      var fixedUrl = ingest_fixGoogleNewsRssUrl_(feedObj.url);

      try {
        var response = UrlFetchApp.fetch(fixedUrl, {
          muteHttpExceptions: true,
          followRedirects: true,
          timeout: 25000,
          headers: {
            "User-Agent": "Mozilla/5.0",
            "Accept": "application/rss+xml, application/xml;q=0.9, text/xml;q=0.8, */*;q=0.7"
          }
        });

        var status = response.getResponseCode();
        if (status < 200 || status >= 300) {
          stats.feedErrors++;
          errors.push("Feed HTTP " + status + ": " + fixedUrl);
          return;
        }

        var xml = String(response.getContentText() || "").trim();
        if (!xml) {
          stats.feedErrors++;
          errors.push("Feed empty body: " + fixedUrl);
          return;
        }

        var safeXml = xml.replace(/^\uFEFF/, "");
        var doc;
        try {
          doc = XmlService.parse(safeXml);
        } catch (parseErr) {
          stats.feedErrors++;
          errors.push("XML parse failed: " + fixedUrl);
          return;
        }

        var root = doc.getRootElement();
        var entries = ingest_getFeedEntries_(root);
        if (!entries.length) return;

        var rowsToWrite = [];

        entries.forEach(function (entry) {
          stats.entriesSeen++;

          var title = ingest_getChildTextAny_(entry, ["title"]);
          var rawLink = ingest_getEntryLink_(entry);
          var linkKey = ingest_normalizeLink_(rawLink);

          var pubDateText = ingest_getChildTextAny_(entry, [
            "pubDate",
            "published",
            "updated",
            "date",
            "created",
            "issued",
            "modified"
          ]);
          var desc = ingest_getChildTextAny_(entry, [
            "description",
            "summary",
            "content",
            "encoded"
          ]) || "";

          if (!title || !linkKey) {
            stats.skippedMissingFields++;
            return;
          }

          if (seenLinks.has(linkKey)) {
            stats.skippedDuplicate++;
            return;
          }
          seenLinks.add(linkKey);

          var pubDate = null;
          if (pubDateText) pubDate = new Date(pubDateText);

          if (cutoff) {
            if (pubDate && !isNaN(pubDate.getTime())) {
              if (pubDate < cutoff) {
                stats.skippedDate++;
                return;
              }
            }
          }

          var source =
            ingest_detectSourceFromEntry_(entry) ||
            feedObj.sourceName ||
            "Unknown";

          var blob = (title + " " + desc).toLowerCase();

          var theme = "";
          var poi = "";
          var matchedKeywords = [];

          if (!ingest_shouldIngestAll_()) {
            var detected = ingest_detectFromRules_(blob, themeRules);
            theme = detected[0];
            poi = detected[1];
            matchedKeywords = detected[2];

            if (!theme) {
              stats.skippedNoMatch++;
              return;
            }
          }

          rowsToWrite.push([
            title,
            rawLink,
            pubDateText,
            source,
            theme,
            poi,
            matchedKeywords.join(", ")
          ]);

          stats.imported++;
        });

        if (rowsToWrite.length) {
          repo_appendRawArticles_(rowsToWrite);
        }

      } catch (e) {
        stats.feedErrors++;
        errors.push("Feed error: " + fixedUrl + " → " + (e && e.message ? e.message : String(e)));
      }
    });

    var ms = Date.now() - t0;
    return {
      ok: true,
      message: "RSS import complete: imported " + stats.imported + " articles from " +
        stats.feedsProcessed + "/" + stats.feedsActive + " active feeds in " + ms + "ms.",
      stats: stats,
      errors: errors.slice(0, 50),
      debug: debug
    };

  } catch (err) {
    return {
      ok: false,
      message: (err && err.message) ? err.message : String(err),
      stats: stats,
      errors: errors.slice(0, 50),
      debug: { where: "top-level catch", _err: String(err) }
    };
  }
}

/* ========= REPO (SHEETS I/O) ========= */

function repo_getSpreadsheet_() {
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  if (ING_CFG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(ING_CFG.SPREADSHEET_ID);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}

function repo_getSheet_(name) {
  var ss = repo_getSpreadsheet_();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet not found: "' + name + '"');
  return sh;
}

function repo_resetRawArticles_() {
  var sh = repo_getSheet_(ING_CFG.RAW_SHEET);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();

  if (lastRow >= 2 && lastCol >= 1) {
    sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
  return sh;
}

function repo_appendRawArticles_(rows) {
  if (!rows || !rows.length) return 0;
  var sh = repo_getSheet_(ING_CFG.RAW_SHEET);
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  return rows.length;
}

/**
 * RSS Feeds sheet expected:
 * Row 1 header
 * Col A: Feed URL
 * Col B: Active? (checkbox TRUE/FALSE)
 * Col C: Source Name (optional)
 */
function repo_getActiveFeeds_() {
  var sh = repo_getSheet_(ING_CFG.FEEDS_SHEET);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  var feeds = [];

  values.forEach(function (r) {
    var url = String(r[0] || "").trim();
    var active = (r[1] === true);
    var sourceName = String(r[2] || "").trim();

    if (!active) return;
    if (!url) return;

    feeds.push({ url: url, sourceName: sourceName });
  });

  return feeds;
}

/**
 * ThemeRules expected:
 * Starts at row 5 (default)
 * Col A: Theme
 * Col B: POI
 * Col C: Keywords
 * Col D: Active? checkbox
 */
function repo_getActiveThemeRules_() {
  var sh = repo_getSheet_(ING_CFG.CONTROL_SHEET);
  var startRow = ING_CFG.THEME_RULES_START_ROW || 5;

  var lastRow = sh.getLastRow();
  if (lastRow < startRow) return [];

  var values = sh.getRange(startRow, 1, lastRow - startRow + 1, 4).getValues();
  var rules = [];

  values.forEach(function (row) {
    var theme = String(row[0] || "").trim();
    var poi = String(row[1] || "").trim();
    var keywordString = String(row[2] || "").trim();
    var active = (row[3] === true);

    if (!active) return;
    if (!theme || !keywordString) return;

    var keywords = ingest_parseKeywords_(keywordString);
    if (!keywords.length) return;

    var keywordRegexes = [];
    keywords.forEach(function (k) {
      var re = ingest_buildKeywordRegex_(k);
      if (re) keywordRegexes.push({ k: k, re: re });
    });

    rules.push({ theme: theme, poi: poi, keywords: keywords, keywordRegexes: keywordRegexes });
  });

  return rules;
}

/* ========= KEYWORDS / MATCHING ========= */

function ingest_shouldIngestAll_() {
  return ING_CFG.INGEST_ALL_ARTICLES === true;
}

function ingest_parseKeywords_(s) {
  return String(s || "")
    .split(/[,;\n]/g)
    .map(function (x) { return String(x || "").trim().toLowerCase(); })
    .filter(function (x) { return !!x; });
}

function ingest_detectFromRules_(textLower, rules) {
  var best = null;
  var bestScore = 0;
  var matchedKeywords = [];

  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    var score = 0;
    var hits = [];

    var regs = r.keywordRegexes || [];
    for (var j = 0; j < regs.length; j++) {
      var obj = regs[j];
      if (obj.re.test(textLower)) {
        score++;
        hits.push(obj.k);
      }
    }

    if (score > bestScore) {
      bestScore = score;
      best = r;
      matchedKeywords = hits;
    }
  }

  if (!best || bestScore < 1) return [null, null, []];
  return [best.theme, best.poi, matchedKeywords];
}

function ingest_escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Keyword -> regex
 * Supports:
 * - phrases (space => \\s+)
 * - wildcard * (e.g. "layoff*" matches layoff, layoffs, layoff-related)
 */
function ingest_buildKeywordRegex_(kw) {
  var raw = String(kw || "").trim().toLowerCase();
  if (!raw) return null;

  var escaped = ingest_escapeRegex_(raw);
  escaped = escaped.replace(/\\\*/g, "[a-z0-9_-]*");

  var phrase = escaped.replace(/\s+/g, "\\s+");

  var isShortSingleToken = raw.length <= 3 && !/\s/.test(raw) && raw.indexOf("*") === -1;
  if (isShortSingleToken) {
    return new RegExp("(^|[^a-z0-9_])" + phrase + "([^a-z0-9_]|$)", "i");
  }

  var startsWithWord = /^[a-z0-9_]/i.test(raw);
  var endsWithWord   = /[a-z0-9_]$/i.test(raw);

  var left  = startsWithWord ? "\\b" : "";
  var right = endsWithWord ? "\\b" : "";

  return new RegExp(left + phrase + right, "i");
}

function ingest_normalizeLink_(link) {
  if (!link) return "";
  return String(link).trim()
    .replace(/^http:\/\//i, "https://")
    .replace(/\/+$/, "");
}

/* ========= GOOGLE NEWS RSS URL FIX ========= */

function ingest_fixGoogleNewsRssUrl_(url) {
  url = String(url || "").trim();
  if (!url) return url;

  if (!/^https?:\/\/news\.google\.com\/rss\/search/i.test(url)) return url;

  var qMatch    = url.match(/[?&]q=([^&]*)/i);
  var hlMatch   = url.match(/[?&]hl=([^&]*)/i);
  var glMatch   = url.match(/[?&]gl=([^&]*)/i);
  var ceidMatch = url.match(/[?&]ceid=([^&]*)/i);

  var qRaw = qMatch ? qMatch[1] : "";
  var qDecoded = qRaw;
  try { qDecoded = decodeURIComponent(qRaw.replace(/\+/g, "%20")); } catch (e) {}

  var hl = hlMatch ? hlMatch[1] : "en-SG";
  var gl = glMatch ? glMatch[1] : "SG";
  var ceid = ceidMatch ? ceidMatch[1] : "SG:en";

  return (
    "https://news.google.com/rss/search" +
    "?q=" + encodeURIComponent(qDecoded) +
    "&hl=" + encodeURIComponent(hl) +
    "&gl=" + encodeURIComponent(gl) +
    "&ceid=" + encodeURIComponent(ceid)
  );
}

/* ========= RSS/ATOM PARSING ========= */

function ingest_getFeedEntries_(root) {
  var channel = ingest_getChildByLocalName_(root, "channel");
  if (channel) {
    var items = ingest_getChildrenByLocalName_(channel, "item");
    if (items.length) return items;
  }

  var rootItems = ingest_getChildrenByLocalName_(root, "item");
  if (rootItems.length) return rootItems;

  var entries = ingest_getChildrenByLocalName_(root, "entry");
  return entries || [];
}

function ingest_getChildByLocalName_(el, localName) {
  if (!el) return null;
  var kids = el.getChildren();
  var target = String(localName || "").toLowerCase();
  for (var i = 0; i < kids.length; i++) {
    var k = kids[i];
    if (!k.getName) continue;
    var kName = ingest_getLocalName_(k);
    if (kName && kName.toLowerCase() === target) return k;
  }
  return null;
}

function ingest_getChildrenByLocalName_(el, localName) {
  if (!el) return [];
  var out = [];
  var kids = el.getChildren();
  var target = String(localName || "").toLowerCase();
  for (var i = 0; i < kids.length; i++) {
    var k = kids[i];
    if (!k.getName) continue;
    var kName = ingest_getLocalName_(k);
    if (kName && kName.toLowerCase() === target) out.push(k);
  }
  return out;
}

function ingest_getChildTextAny_(element, names) {
  try {
    for (var i = 0; i < names.length; i++) {
      var child = ingest_getChildByLocalName_(element, names[i]);
      if (child) {
        var t = String(child.getText() || "").trim();
        if (t) return t;
      }
    }
    return "";
  } catch (e) {
    return "";
  }
}

function ingest_getEntryLink_(entry) {
  var rssLink = ingest_getChildTextAny_(entry, ["link"]);
  if (rssLink) return rssLink;

  var linkEls = ingest_getChildrenByLocalName_(entry, "link");
  if (linkEls.length) {
    var preferred = null;
    for (var i = 0; i < linkEls.length; i++) {
      var linkEl = linkEls[i];
      var rel = "";
      try {
        var relAttr = linkEl.getAttribute("rel");
        rel = relAttr ? String(relAttr.getValue() || "").trim().toLowerCase() : "";
      } catch (e) {}

      if (!rel || rel === "alternate") {
        preferred = linkEl;
        break;
      }

      if (!preferred) preferred = linkEl;
    }

    if (preferred) {
      try {
        var hrefAttr = preferred.getAttribute("href");
        if (hrefAttr) {
          var v = String(hrefAttr.getValue() || "").trim();
          if (v) return v;
        }
      } catch (e) {}

      var t = String(preferred.getText() || "").trim();
      if (t) return t;
    }
  }

  var guid = ingest_getChildTextAny_(entry, ["guid", "id"]);
  return guid || "";
}

function ingest_detectSourceFromEntry_(entry) {
  try {
    var src = ingest_getChildByLocalName_(entry, "source");
    if (src) {
      var t1 = String(src.getText() || "").trim();
      if (t1) return t1;

      var srcTitle = ingest_getChildByLocalName_(src, "title");
      if (srcTitle) {
        var t2 = String(srcTitle.getText() || "").trim();
        if (t2) return t2;
      }
    }

    var author = ingest_getChildByLocalName_(entry, "author");
    if (author) {
      var name = ingest_getChildByLocalName_(author, "name");
      if (name) {
        var t3 = String(name.getText() || "").trim();
        if (t3) return t3;
      }
    }

    return "";
  } catch (e) {
    return "";
  }
}

function ingest_getLocalName_(el) {
  if (!el || !el.getName) return "";
  var name = String(el.getName() || "");
  if (!name) return "";
  var parts = name.split(":");
  return parts[parts.length - 1];
}
