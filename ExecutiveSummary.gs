/*********************************
 * EXECUTIVE SUMMARY (AI)
 *
 * Reads from: Weekly Picks
 * Writes to: Executive Summary
 *
 * - No manual priority column
 * - Uses Weekly Picks as curated input
 * - Sorts by Score if present, else by Date
 * - AI selects 3–5 articles
 *********************************/

// ===== Sheet names =====
const ES_PROMPT_SHEET = "PromptConfig";
const ES_RAW_SHEET    = "Weekly Picks";
const ES_EXEC_SHEET   = "Executive Summary";

// How many Weekly Picks to send to AI
const ES_MAX_CANDIDATES = 20;

/**
 * Reads a header row and returns a map: normalizedHeader -> 1-based column index
 */
function headerMap_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const key = String(h || "").trim().toLowerCase();
    if (key) map[key] = i + 1;
  });
  return map;
}

/**
 * Safe parse date to ms
 */
function safeDateMs_(v) {
  const d = new Date(v);
  const t = d.getTime();
  return isNaN(t) ? 0 : t;
}

/**
 * Reads prompt config from PromptConfig sheet
 * Supports:
 *  A) A2=persona, B2=weekly_instructions
 *  B) key/value table in columns A:B
 */
function getPromptConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ES_PROMPT_SHEET);
  if (!sh) throw new Error(`Prompt config sheet not found: ${ES_PROMPT_SHEET}`);

  const personaA2 = (sh.getRange("A2").getDisplayValue() || "").trim();
  const weeklyB2  = (sh.getRange("B2").getDisplayValue() || "").trim();

  if (personaA2 || weeklyB2) {
    return { persona: personaA2, weekly_instructions: weeklyB2 };
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { persona: "", weekly_instructions: "" };

  const vals = sh.getRange(2, 1, lastRow - 1, 2).getValues();
  const map = {};
  vals.forEach(r => {
    const key = (r[0] || "").toString().trim();
    const value = (r[1] || "").toString();
    if (key) map[key] = value;
  });

  return {
    persona: (map["persona"] || "").trim(),
    weekly_instructions: (map["weekly_instructions"] || "").trim()
  };
}

/**
 * Builds the final Gemini prompt
 */
function buildExecutivePrompt_(persona, weeklyInstructions, articlesForModel) {
  const modelInputJson = JSON.stringify(articlesForModel);

  return [
    persona,
    "",
    "You are selecting 3–5 articles most relevant for manpower, labour market,",
    "workforce policy, leadership decisions, and forward-looking policy insight.",
    "",
    weeklyInstructions,
    "",
    "Candidate articles (JSON array):",
    modelInputJson,
    "",
    "Return ONLY a JSON array like:",
    "[",
    "  {\"index\": 0, \"rationale\": \"Why this article is important for policymakers\"}",
    "]",
    "",
    "Where index refers to the index in the candidate list above."
  ].join("\n");
}

// ===== EXECUTIVE SUMMARY (AI ONLY) =====
function generateExecutiveSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(ES_RAW_SHEET);
  if (!sh) {
    SpreadsheetApp.getUi().alert(`"${ES_RAW_SHEET}" sheet not found.`);
    return;
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert(`No rows in "${ES_RAW_SHEET}".`);
    return;
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const hmap = headerMap_(headers);

  const colTitle  = hmap["title"] || 1;
  const colLink   = hmap["link"] || 2;
  const colDate   = hmap["date"] || 3;
  const colSource = hmap["source"] || 4;
  const colTheme  = hmap["theme"] || 5;
  const colPoi    = hmap["point of interest"] || hmap["poi"] || 0;
  const colScore  = hmap["score"] || 0;

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  let rows = data.map(r => {
    const title = r[colTitle - 1];
    const link  = r[colLink - 1];
    if (!title || !link) return null;

    return {
      title: String(title),
      link: String(link),
      date: r[colDate - 1],
      source: String(r[colSource - 1] || ""),
      theme: String(r[colTheme - 1] || ""),
      poi: colPoi ? String(r[colPoi - 1] || "") : "",
      score: colScore ? Number(r[colScore - 1]) || 0 : null
    };
  }).filter(Boolean);

  if (!rows.length) {
    SpreadsheetApp.getUi().alert("No usable rows found in Weekly Picks.");
    return;
  }

  const hasScore = colScore > 0;
  rows.sort((a, b) => {
    if (hasScore && b.score !== a.score) return b.score - a.score;
    return safeDateMs_(b.date) - safeDateMs_(a.date);
  });

  const candidates = rows.slice(0, ES_MAX_CANDIDATES);

  const articlesForModel = candidates.map((a, idx) => ({
    index: idx,
    title: a.title,
    link: a.link,
    date: a.date,
    source: a.source,
    theme: a.theme,
    pointOfInterest: a.poi,
    valueType: hasScore ? `Score ${a.score}` : "Weekly Pick",
    contentSnippet: fetchArticleText(a.link)
  }));

  const cfg = getPromptConfig_();
  if (!cfg.persona || !cfg.weekly_instructions) {
    SpreadsheetApp.getUi().alert(`Missing persona or weekly instructions in "${ES_PROMPT_SHEET}".`);
    return;
  }

  const prompt = buildExecutivePrompt_(cfg.persona, cfg.weekly_instructions, articlesForModel);

  let selectionText = "";
  try {
    selectionText = aiGenerateText_(prompt, {
      temperature: 0.2,
      maxOutputTokens: 800
    });
  } catch (e) {
    Logger.log("AI API failed: " + e);
  }

  let selected = [];
  try {
    const m = selectionText.match(/(\[[\s\S]*\])/);
    selected = JSON.parse(m ? m[1] : selectionText);
  } catch (_) {}

  if (!Array.isArray(selected) || !selected.length) {
    selected = candidates.slice(0, 5).map((_, i) => ({
      index: i,
      rationale: "Fallback: highest-ranked Weekly Pick."
    }));
  }

  let exec = ss.getSheetByName(ES_EXEC_SHEET);
  if (!exec) exec = ss.insertSheet(ES_EXEC_SHEET);
  else exec.clearContents();

  const header = ["Title","Link","Date","Source","Theme"];
  if (colPoi) header.push("Point of Interest");
  if (colScore) header.push("Score");
  header.push("AI Rationale");

  exec.appendRow(header);

  const out = selected.map(sel => {
    const a = candidates[sel.index];
    const row = [a.title, a.link, a.date, a.source, a.theme];
    if (colPoi) row.push(a.poi);
    if (colScore) row.push(a.score);
    row.push(sel.rationale || "");
    return row;
  });

  exec.getRange(2, 1, out.length, header.length).setValues(out);

  SpreadsheetApp.getUi().alert(
    `Executive Summary generated.\nCandidates: ${candidates.length}\nSelected: ${out.length}`
  );
}

/********************************************
 * FETCH & CLEAN ARTICLE CONTENT
 ********************************************/
function fetchArticleText(url) {
  if (!url) return "";

  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "User-Agent": "Mozilla/5.0" }
    });

    if (response.getResponseCode() >= 300) return "";

    let cleaned = response.getContentText()
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<noscript[\s\S]*?<\/noscript>/gi, "")
      .replace(/<!--[\s\S]*?-->/g, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/\s+/g, " ")
      .trim();

    return cleaned.substring(0, 3000);
  } catch (e) {
    Logger.log("fetchArticleText error: " + e);
    return "";
  }
}
