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

  var fetchTimeoutMs = (typeof INGEST_FETCH_TIMEOUT_MS !== "undefined") ? INGEST_FETCH_TIMEOUT_MS : 15000;
  var maxRuntimeMs = (typeof INGEST_MAX_RUNTIME_MS !== "undefined") ? INGEST_MAX_RUNTIME_MS : 330000;
  var fetchBatchSize = (typeof INGEST_FETCH_BATCH_SIZE !== "undefined") ? INGEST_FETCH_BATCH_SIZE : 15;
  var dailyDaysBack = (typeof INGEST_DAILY_DAYS_BACK !== "undefined") ? INGEST_DAILY_DAYS_BACK : 1;
  var pruneDays = (typeof INGEST_PRUNE_DAYS !== "undefined") ? INGEST_PRUNE_DAYS : 14;

  // If your UI.gs defines SPREADSHEET_ID, reuse it.
  var spreadsheetId = (typeof SPREADSHEET_ID !== "undefined") ? SPREADSHEET_ID : "";

  return {
    CONTROL_SHEET: controlSheet,
    FEEDS_SHEET: feedsSheet,
    RAW_SHEET: rawSheet,
    THEME_RULES_START_ROW: rulesStartRow,
    INGEST_ALL_ARTICLES: ingestAll,
    INGEST_FETCH_TIMEOUT_MS: fetchTimeoutMs,
    INGEST_MAX_RUNTIME_MS: maxRuntimeMs,
    INGEST_FETCH_BATCH_SIZE: fetchBatchSize,
    INGEST_DAILY_DAYS_BACK: dailyDaysBack,
    INGEST_PRUNE_DAYS: pruneDays,
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
function ui_importRSS_14days_v1() {
  var res = ingest_importRSS_(14);
  res._sig = "IngestService.ui_importRSS_14days_v1 @ 2026-02-19";
  return res;
}
function ui_importRSS_daily_v1() {
  var res;
  try {
    res = ingest_dailyRawArticles_();
  } catch (err) {
    var msg = (err && err.message) ? err.message : String(err || "Unknown error");
    return { ok: false, message: msg, errors: [msg], _sig: "IngestService.ui_importRSS_daily_v1 @ 2026-01-13" };
  }
  if (!res || typeof res.ok === "undefined") {
    return {
      ok: false,
      message: "Daily ingest returned no response object.",
      errors: ["Daily ingest returned no response object."],
      _sig: "IngestService.ui_importRSS_daily_v1 @ 2026-01-13"
    };
  }
  res._sig = "IngestService.ui_importRSS_daily_v1 @ 2026-01-13";
  return res;
}
function ui_importRSS_30days_v2() {
  var res = ingest_importRSS_(14);
  res._sig = "IngestService.ui_importRSS_30days_v2 @ 2026-02-19";
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

/**
 * Daily ingest runner (intended for 12:00 daily trigger)
 * - Prunes raw articles older than INGEST_PRUNE_DAYS
 * - Appends new articles (no duplicates)
 */
function ingest_dailyRawArticles_() {
  var pruneStats = repo_pruneRawArticlesByAge_(ING_CFG.INGEST_PRUNE_DAYS);
  pruneStats.maxAgeDays = ING_CFG.INGEST_PRUNE_DAYS;
  var ingestRes = ingest_importRSS_(ING_CFG.INGEST_DAILY_DAYS_BACK);
  ingestRes.prune = pruneStats;

  if (ingestRes && ingestRes.ok && ingestRes.stats && Number(ingestRes.stats.imported || 0) > 0) {
    ingestRes.llm = ingest_runDailyLlmScoringForNewArticles_({ timeBudgetMs: 45000 });

    if (ingestRes.llm && ingestRes.llm.status === "in_progress") {
      ingestRes.llm.queue = ingest_scheduleDailyLlmScoring_();
    }

    if (ingestRes.llm && ingestRes.llm.ok === false) {
      if (!Array.isArray(ingestRes.errors)) ingestRes.errors = [];
      ingestRes.errors.push("Daily LLM scoring warning: " + String(ingestRes.llm.message || "Unknown LLM scoring error"));
    }
  } else {
    ingestRes.llm = {
      ok: true,
      status: "skipped",
      message: "No new articles imported; LLM scoring skipped."
    };
  }

  ingestRes._sig = "IngestService.ingest_dailyRawArticles_ @ 2026-02-19";
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty("RAW_INGEST_LAST_RUN", new Date().toISOString());
    props.setProperty("RAW_INGEST_LAST_STATUS", ingestRes.ok ? "ok" : "error");
    props.setProperty("RAW_INGEST_LAST_MESSAGE", String(ingestRes.message || ""));
    props.setProperty(
      "RAW_INGEST_LAST_RUNTIME_MS",
      String((ingestRes.stats && ingestRes.stats.runtimeMs) || "")
    );
    props.setProperty(
      "RAW_INGEST_LAST_STOPPED_EARLY",
      (ingestRes.stats && ingestRes.stats.stoppedEarly) ? "true" : "false"
    );
    props.setProperty(
      "RAW_INGEST_LAST_ERRORS",
      Array.isArray(ingestRes.errors) ? ingestRes.errors.slice(0, 3).join(" | ") : ""
    );
  } catch (e) {
    console.log("Failed to record RAW_INGEST_LAST_RUN:", e);
  }
  return ingestRes;
}

function ingest_scheduleDailyLlmScoring_() {
  var handlerName = "ingest_runDailyLlmScoringJob_";
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var removed = 0;
    triggers.forEach(function (t) {
      if (t.getHandlerFunction && t.getHandlerFunction() === handlerName) {
        ScriptApp.deleteTrigger(t);
        removed += 1;
      }
    });

    ScriptApp.newTrigger(handlerName)
      .timeBased()
      .after(5000)
      .create();

    return {
      ok: true,
      status: "queued",
      handler: handlerName,
      replacedTriggers: removed,
      message: "Daily LLM scoring queued to run in background."
    };
  } catch (err) {
    return {
      ok: false,
      status: "queue_error",
      handler: handlerName,
      message: String(err && err.message ? err.message : err)
    };
  }
}

function ingest_runDailyLlmScoringJob_() {
  var handlerName = "ingest_runDailyLlmScoringJob_";
  var res = ingest_runDailyLlmScoringForNewArticles_({});

  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction && t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (res && res.ok && res.status === "in_progress") {
    try {
      ScriptApp.newTrigger(handlerName)
        .timeBased()
        .after(15000)
        .create();
    } catch (e) {
      if (!Array.isArray(res.errors)) res.errors = [];
      res.errors.push("Failed to queue next LLM scoring pass: " + String(e && e.message ? e.message : e));
    }
  }

  return res;
}

function ingest_runDailyLlmScoringForNewArticles_(opts) {
  if (typeof ui_runRawArticlesLlmRank_v2 !== "function") {
    return { ok: false, message: "LLM scoring function ui_runRawArticlesLlmRank_v2 is unavailable." };
  }

  var prompt = (typeof RAW_LLM_SCORE_ONLY_GUIDE_TEXT !== "undefined" && RAW_LLM_SCORE_ONLY_GUIDE_TEXT)
    ? RAW_LLM_SCORE_ONLY_GUIDE_TEXT
    : "Score each article for executive horizon-scanning relevance and set publish_recommendation to publish, maybe, or skip.";

  var lastRes = ui_runRawArticlesLlmRank_v2({
    prompt: prompt,
    maxRows: 0,
    fetchText: false,
    sectorBy: "theme",
    timeBudgetMs: Number(opts && opts.timeBudgetMs || 0)
  });

  if (!lastRes || lastRes.ok !== true) {
    return {
      ok: false,
      message: (lastRes && lastRes.message) ? lastRes.message : "LLM scoring returned no response.",
      response: lastRes || null
    };
  }

  return {
    ok: true,
    status: lastRes.status || "done",
    meta: lastRes.meta || null,
    errors: Array.isArray(lastRes.errors) ? lastRes.errors : []
  };
}

function ui_pruneRawArticlesNow_v1(payload) {
  var daysRaw = payload && payload.days;
  var days = Number(daysRaw || ING_CFG.INGEST_PRUNE_DAYS || 14);
  if (!isFinite(days) || days <= 0) days = ING_CFG.INGEST_PRUNE_DAYS || 14;
  days = Math.max(1, Math.round(days));

  var stats = repo_pruneRawArticlesByAge_(days);
  stats.maxAgeDays = days;
  return {
    ok: true,
    message: "Pruned Raw Articles older than " + days + " days.",
    prune: stats,
    _sig: "IngestService.ui_pruneRawArticlesNow_v1 @ 2026-02-19"
  };
}

/**
 * One-time helper to install a 12:00 daily trigger for ingest_dailyRawArticles_.
 */
function ingest_setupDailyRawIngestTrigger_() {
  var handlerName = "ingest_dailyRawArticles_";
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction && t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  return { ok: true, handler: handlerName, atHour: 12 };
}

/**
 * Ensures a daily trigger exists for ingest_dailyRawArticles_.
 * Safe to call on open; only creates a trigger if missing.
 */
function ingest_ensureDailyRawIngestTrigger_() {
  var handlerName = "ingest_dailyRawArticles_";
  var triggers = ScriptApp.getProjectTriggers();
  var existing = triggers.some(function (t) {
    return t.getHandlerFunction && t.getHandlerFunction() === handlerName;
  });
  if (existing) {
    return { ok: true, handler: handlerName, created: false };
  }
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  return { ok: true, handler: handlerName, created: true, atHour: 12 };
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
    skippedMissingFields: 0,
    stoppedEarly: false,
    runtimeMs: 0
  };

  var errors = [];

  try {
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
      ingestAll: ingest_shouldIngestAll_(),
      fetchTimeoutMs: ING_CFG.INGEST_FETCH_TIMEOUT_MS,
      maxRuntimeMs: ING_CFG.INGEST_MAX_RUNTIME_MS,
      fetchBatchSize: ING_CFG.INGEST_FETCH_BATCH_SIZE
    };

    if (!feeds.length) {
      stats.runtimeMs = Date.now() - t0;
      return { ok:false, message:'No active feeds found in "' + ING_CFG.FEEDS_SHEET + '".', stats:stats, errors:errors, debug:debug };
    }

    if (!themeRules.length && !ingest_shouldIngestAll_()) {
      stats.runtimeMs = Date.now() - t0;
      return { ok:false, message:'No active ThemeRules found in "' + ING_CFG.CONTROL_SHEET + '" (and INGEST_ALL_ARTICLES is false).', stats:stats, errors:errors, debug:debug };
    }

    var cutoff = null;
    if (daysBack !== null) {
      cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - Number(daysBack));
    }

    var seenLinks = ingest_buildExistingLinkSet_();

    var allRowsToWrite = [];
    var maxRuntimeMs = ING_CFG.INGEST_MAX_RUNTIME_MS;
    var batchSize = Math.max(1, ING_CFG.INGEST_FETCH_BATCH_SIZE || 1);

    for (var start = 0; start < feeds.length; start += batchSize) {
      if (maxRuntimeMs && (Date.now() - t0) > maxRuntimeMs) {
        stats.stoppedEarly = true;
        break;
      }

      var batch = feeds.slice(start, start + batchSize).map(function (feedObj) {
        return {
          feedObj: feedObj,
          fixedUrl: ingest_fixGoogleNewsRssUrl_(feedObj.url)
        };
      });

      var fetchOptions = {
        muteHttpExceptions: true,
        followRedirects: true,
        timeout: ING_CFG.INGEST_FETCH_TIMEOUT_MS,
        headers: {
          "User-Agent": "Mozilla/5.0",
          "Accept": "application/rss+xml, application/xml;q=0.9, text/xml;q=0.8, */*;q=0.7"
        }
      };

      var requests = batch.map(function (item) {
        var req = {};
        Object.keys(fetchOptions).forEach(function (k) {
          req[k] = fetchOptions[k];
        });
        req.url = item.fixedUrl;
        return req;
      });

      var responses;
      try {
        responses = UrlFetchApp.fetchAll(requests);
      } catch (batchErr) {
        responses = batch.map(function (item) {
          try {
            return UrlFetchApp.fetch(item.fixedUrl, fetchOptions);
          } catch (singleErr) {
            return { __error: singleErr };
          }
        });
      }

      responses.forEach(function (response, idx) {
        var batchItem = batch[idx];
        var feedObj = batchItem.feedObj;
        var fixedUrl = batchItem.fixedUrl;
        stats.feedsProcessed++;

        if (response && response.__error) {
          stats.feedErrors++;
          errors.push("Feed error: " + fixedUrl + " → " + (response.__error.message || String(response.__error)));
          return;
        }

        try {
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

            allRowsToWrite.push([
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
        } catch (e) {
          stats.feedErrors++;
          errors.push("Feed error: " + fixedUrl + " → " + (e && e.message ? e.message : String(e)));
        }
      });
    }

    if (allRowsToWrite.length) {
      repo_appendRawArticles_(allRowsToWrite);
    }

    var ms = Date.now() - t0;
    stats.runtimeMs = ms;
    return {
      ok: true,
      message: "RSS import complete: imported " + stats.imported + " articles from " +
        stats.feedsProcessed + "/" + stats.feedsActive + " active feeds in " + ms + "ms." +
        (stats.stoppedEarly ? " (Stopped early to avoid execution time limit.)" : ""),
      stats: stats,
      errors: errors.slice(0, 50),
      debug: debug
    };

  } catch (err) {
    stats.runtimeMs = Date.now() - t0;
    return {
      ok: false,
      message: (err && err.message) ? err.message : String(err),
      stats: stats,
      errors: errors.slice(0, 50),
      debug: { where: "top-level catch", _err: String(err) }
    };
  }
}

function ingest_buildExistingLinkSet_() {
  var links = repo_getRawArticleLinks_();
  var set = new Set();
  links.forEach(function (rawLink) {
    var linkKey = ingest_normalizeLink_(rawLink);
    if (linkKey) set.add(linkKey);
  });
  return set;
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
