/*********************************
 * RawLlmSearch.gs — LLM-based ranking for Raw Articles
 * - Uses AIService.aiGenerateJson_()
 * - Optionally fetches article text
 *********************************/

const WEEKLY_PICKS_LLM_GUIDE_TEXT = [
  "You are “FutureScans@MOM Article Rater”.",
  "",
  "Goal:",
  "Given ONE news article (title + url + date + source + text/snippet + existing Theme/POI labels), decide if it is suitable for a weekly workforce/workplace/future-of-work horizon scan. Prefer structural/futures signals, evidence-backed reporting, and clear “levers” (policy/employer/institution actions). Avoid pure politics, celebrity gossip, stock/earnings-only pieces, and generic career self-help.",
  "",
  "The Theme/POI labels are already assigned; DO NOT re-classify or change them. Focus on scoring and recommendation only.",
  "",
  "Output MUST be valid JSON only (no markdown, no commentary). If article text is missing or too short, still score using what is available and set uncertainty=\"high\".",
  "",
  "========================",
  "GATES",
  "========================",
  "Hard gates (gate_pass must be true to recommend publish/maybe):",
  "1) topic_relevance must be >= 2 (0–3 scale)",
  "2) article is not primarily unrelated politics/celebrity/earnings-only",
  "3) has at least ONE of: lever OR credible evidence OR clear workforce/workplace policy implication",
  "",
  "topic_relevance scale:",
  "0 = not about workforce/workplace/future-of-work",
  "1 = tangential (weak linkage)",
  "2 = clearly relevant",
  "3 = strongly relevant and central",
  "",
  "========================",
  "SCORING (0–100)",
  "========================",
  "Compute final_score using:",
  "- policy_or_operational_lever (0–20): concrete lever/action (law, regulation, programme, employer practice)",
  "- evidence_strength (0–20): credible stats/surveys/reports/official sources (not just opinions)",
  "- futures_signal (0–15): structural shift / forward-looking risk/opportunity",
  "- distributional_impact (0–10): identifies affected groups (e.g., older workers, women, low-wage, migrants, youth)",
  "- transferability_to_singapore_MOM (0–15): plausible learning/applicability to Singapore/MOM levers",
  "- novelty (0–10): new/developing, not evergreen advice",
  "- source_credibility (0–10): reputable outlet and/or citing primary sources",
  "",
  "publish_recommendation:",
  "- \"publish\" if final_score >= 70 AND gate_pass=true",
  "- \"maybe\" if 55–69 AND gate_pass=true",
  "- \"skip\" otherwise",
  "",
  "========================",
  "WRITING RULES",
  "========================",
  "- relevance_sg_mom must be 30–40 words, plain English, no “This article says…”.",
  "- summary must be 80–90 words, neutral tone, include 1–2 bolded phrases using **double asterisks**.",
  "- score_reasons: max 5 bullets, each max 14 words, reference concrete signals.",
  "",
  "========================",
  "RETURN JSON ONLY (schema)",
  "========================",
  "{",
  "  \"title\": string,",
  "  \"url\": string,",
  "  \"date\": string,",
  "  \"source\": string,",
  "  \"theme\": string,",
  "  \"poi\": string,",
  "",
  "  \"gate_pass\": boolean,",
  "  \"topic_relevance\": 0|1|2|3,",
  "",
  "  \"affected_groups\": string[],",
  "  \"levers\": string[],",
  "  \"evidence_signals\": string[],",
  "",
  "  \"policy_or_operational_lever\": number,",
  "  \"evidence_strength\": number,",
  "  \"futures_signal\": number,",
  "  \"distributional_impact\": number,",
  "  \"transferability_to_singapore_MOM\": number,",
  "  \"novelty\": number,",
  "  \"source_credibility\": number,",
  "  \"final_score\": number,",
  "",
  "  \"publish_recommendation\": \"publish\"|\"maybe\"|\"skip\",",
  "  \"uncertainty\": \"low\"|\"medium\"|\"high\",",
  "",
  "  \"relevance_sg_mom\": string,",
  "  \"summary\": string,",
  "  \"score_reasons\": string[]",
  "}",
  "",
  "Now rate the article using ONLY the provided inputs.",
  "If the article is clearly irrelevant, keep relevance_sg_mom and summary short but valid strings."
].join("\n");

const RAW_LLM_GUIDE_TEXT = [
  "You are “FutureScans@MOM Article Rater”.",
  "",
  "Goal:",
  "Given ONE news article (title + url + date + source + text/snippet + existing Theme/POI labels), decide if it is suitable for a weekly workforce/workplace/future-of-work horizon scan. Prefer structural/futures signals, evidence-backed reporting, and clear “levers” (policy/employer/institution actions). Avoid pure politics, celebrity gossip, stock/earnings-only pieces, and generic career self-help.",
  "",
  "The Theme/POI labels are already assigned; DO NOT re-classify or change them. Focus on scoring and recommendation only.",
  "",
  "Output MUST be valid JSON only (no markdown, no commentary). If article text is missing or too short, still score using what is available and set uncertainty=\"high\".",
  "",
  "========================",
  "GATES",
  "========================",
  "Hard gates (gate_pass must be true to recommend publish/maybe):",
  "1) topic_relevance must be >= 2 (0–3 scale)",
  "2) article is not primarily unrelated politics/celebrity/earnings-only",
  "3) has at least ONE of: lever OR credible evidence OR clear workforce/workplace policy implication",
  "",
  "topic_relevance scale:",
  "0 = not about workforce/workplace/future-of-work",
  "1 = tangential (weak linkage)",
  "2 = clearly relevant",
  "3 = strongly relevant and central",
  "",
  "========================",
  "SCORING (0–100)",
  "========================",
  "Compute final_score using:",
  "- policy_or_operational_lever (0–20): concrete lever/action (law, regulation, programme, employer practice)",
  "- evidence_strength (0–20): credible stats/surveys/reports/official sources (not just opinions)",
  "- futures_signal (0–15): structural shift / forward-looking risk/opportunity",
  "- distributional_impact (0–10): identifies affected groups (e.g., older workers, women, low-wage, migrants, youth)",
  "- transferability_to_singapore_MOM (0–15): plausible learning/applicability to Singapore/MOM levers",
  "- novelty (0–10): new/developing, not evergreen advice",
  "- source_credibility (0–10): reputable outlet and/or citing primary sources",
  "",
  "publish_recommendation:",
  "- \"publish\" if final_score >= 70 AND gate_pass=true",
  "- \"maybe\" if 55–69 AND gate_pass=true",
  "- \"skip\" otherwise",
  "",
  "========================",
  "WRITING RULES",
  "========================",
  "- relevance_sg_mom must be 30–40 words, plain English, no “This article says…”.",
  "- summary must be 80–90 words, neutral tone, include 1–2 bolded phrases using **double asterisks**.",
  "- score_reasons: max 5 bullets, each max 14 words, reference concrete signals.",
  "",
  "========================",
  "RETURN JSON ONLY (schema)",
  "========================",
  "{",
  "  \"title\": string,",
  "  \"url\": string,",
  "  \"date\": string,",
  "  \"source\": string,",
  "  \"theme\": string,",
  "  \"poi\": string,",
  "",
  "  \"gate_pass\": boolean,",
  "  \"topic_relevance\": 0|1|2|3,",
  "",
  "  \"affected_groups\": string[],",
  "  \"levers\": string[],",
  "  \"evidence_signals\": string[],",
  "",
  "  \"policy_or_operational_lever\": number,",
  "  \"evidence_strength\": number,",
  "  \"futures_signal\": number,",
  "  \"distributional_impact\": number,",
  "  \"transferability_to_singapore_MOM\": number,",
  "  \"novelty\": number,",
  "  \"source_credibility\": number,",
  "  \"final_score\": number,",
  "",
  "  \"publish_recommendation\": \"publish\"|\"maybe\"|\"skip\",",
  "  \"uncertainty\": \"low\"|\"medium\"|\"high\",",
  "",
  "  \"relevance_sg_mom\": string,",
  "  \"summary\": string,",
  "  \"score_reasons\": string[]",
  "}",
  "",
  "Now rate the article using ONLY the provided inputs.",
  "If the article is clearly irrelevant, keep relevance_sg_mom and summary short but valid strings."
].join("\n");

const RAW_LLM_SCORE_ONLY_GUIDE_TEXT = [
  "You are “FutureScans@MOM Article Rater”.",
  "",
  "Goal:",
  "Given ONE news article (title + url + date + source + text/snippet + existing Theme/POI labels), decide if it is suitable for a weekly workforce/workplace/future-of-work horizon scan. Prefer structural/futures signals, evidence-backed reporting, and clear “levers” (policy/employer/institution actions). Avoid pure politics, celebrity gossip, stock/earnings-only pieces, and generic career self-help.",
  "",
  "The Theme/POI labels are already assigned; DO NOT re-classify or change them. Focus on scoring and recommendation only.",
  "",
  "Output MUST be valid JSON only (no markdown, no commentary). If article text is missing or too short, still score using what is available and set uncertainty=\"high\".",
  "",
  "========================",
  "GATES",
  "========================",
  "Hard gates (gate_pass must be true to recommend publish/maybe):",
  "1) topic_relevance must be >= 2 (0–3 scale)",
  "2) article is not primarily unrelated politics/celebrity/earnings-only",
  "3) has at least ONE of: lever OR credible evidence OR clear workforce/workplace policy implication",
  "",
  "topic_relevance scale:",
  "0 = not about workforce/workplace/future-of-work",
  "1 = tangential (weak linkage)",
  "2 = clearly relevant",
  "3 = strongly relevant and central",
  "",
  "========================",
  "SCORING (0–100)",
  "========================",
  "Compute final_score using:",
  "- policy_or_operational_lever (0–20): concrete lever/action (law, regulation, programme, employer practice)",
  "- evidence_strength (0–20): credible stats/surveys/reports/official sources (not just opinions)",
  "- futures_signal (0–15): structural shift / forward-looking risk/opportunity",
  "- distributional_impact (0–10): identifies affected groups (e.g., older workers, women, low-wage, migrants, youth)",
  "- transferability_to_singapore_MOM (0–15): plausible learning/applicability to Singapore/MOM levers",
  "- novelty (0–10): new/developing, not evergreen advice",
  "- source_credibility (0–10): reputable outlet and/or citing primary sources",
  "",
  "publish_recommendation:",
  "- \"publish\" if final_score >= 70 AND gate_pass=true",
  "- \"maybe\" if 55–69 AND gate_pass=true",
  "- \"skip\" otherwise",
  "",
  "========================",
  "RETURN JSON ONLY (schema)",
  "========================",
  "{",
  "  \"gate_pass\": boolean,",
  "  \"topic_relevance\": 0|1|2|3,",
  "  \"final_score\": number,",
  "  \"publish_recommendation\": \"publish\"|\"maybe\"|\"skip\",",
  "  \"uncertainty\": \"low\"|\"medium\"|\"high\"",
  "}",
  "",
  "Now rate the article using ONLY the provided inputs."
].join("\n");

const RAW_LLM_MAX_TEXT_CHARS = 6000;
const RAW_LLM_SCORE_MAX_TEXT_CHARS = 1200;
const RAW_LLM_RANK_CACHE_KEY = "RAW_LLM_RANK_CACHE_V1";
const RAW_LLM_RANK_PROGRESS_KEY = "RAW_LLM_RANK_PROGRESS_V1";
const RAW_LLM_RANK_TIME_BUDGET_MS = 250000;
const RAW_LLM_RANK_SHEET = "Raw LLM Rank Cache";
const RAW_LLM_RANK_META_SHEET = "Raw LLM Rank Meta";

function ui_runWeeklyPicksLlmRank_v1(payload) {
  try {
    const prompt = String(payload?.prompt || "").trim();
    if (!prompt) return { ok: false, message: "Missing prompt." };

    const rows = Array.isArray(payload?.rows) ? payload.rows : [];
    const maxRows = Math.max(1, Math.min(60, Number(payload?.maxRows || rows.length || 20)));
    const fetchText = payload?.fetchText === true;

    const results = [];
    const errors = [];
    const target = rows.slice(0, maxRows);
    let promptChars = 0;

    target.forEach((row) => {
      try {
        const article = weekly_buildLlmArticle_(row, fetchText);
        const llmPrompt = weekly_buildLlmPrompt_(prompt, article);
        promptChars += llmPrompt.length;
        const responseText = aiGenerateJson_(llmPrompt, { maxOutputTokens: 1200 });
        const parsed = feeds_safeParseJsonObject_(responseText);
        if (!parsed) {
          errors.push(`No JSON parsed for: ${article.url || article.title}`);
          return;
        }

        const finalScore = Number(parsed.final_score);
        const llm = Object.assign({}, parsed, {
          final_score: Number.isFinite(finalScore) ? finalScore : 0
        });

        results.push({
          key: row.key,
          title: article.title,
          url: article.url,
          llm
        });
      } catch (err) {
        errors.push(String(err?.message || err));
      }
    });

    return {
      ok: true,
      meta: {
        prompt,
        requested: target.length,
        scored: results.length,
        prompt_chars_total: promptChars,
        tokens_estimate: estimateTokensFromChars_(promptChars)
      },
      results,
      errors
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function ui_runRawArticlesLlmRank_v2(payload) {
  try {
    const prompt = String(payload?.prompt || "").trim();
    if (!prompt) return { ok: false, message: "Missing prompt." };

    const topN = Math.max(1, Number(payload?.maxRows || 20));
    const fetchText = payload?.fetchText === true;
    const scoreFetchText = false;

    const raw = ui_getRawArticles_bootstrap_v1();
    if (!raw || raw.ok !== true) {
      return { ok: false, message: raw?.message || "Failed to load raw articles." };
    }
    const rows = Array.isArray(raw.rows) ? raw.rows : [];

    const startedAt = Date.now();
    const results = [];
    const errors = [];
    const progress = raw_getLlmRankProgress_();
    const sameJob = progress &&
      progress.prompt === prompt &&
      progress.topN === topN &&
      progress.fetchText === fetchText;

    if (!sameJob && progress) {
      raw_clearLlmRankProgress_();
    }

    const scored = sameJob && Array.isArray(progress?.scored) ? progress.scored : [];
    let promptChars = sameJob ? Number(progress?.promptChars || 0) : 0;
    let nextIndex = sameJob ? Number(progress?.nextIndex || 0) : 0;

    for (let i = nextIndex; i < rows.length; i += 1) {
      if (Date.now() - startedAt > RAW_LLM_RANK_TIME_BUDGET_MS) {
        nextIndex = i;
        break;
      }
      const row = rows[i];
      try {
        const article = raw_buildLlmArticle_(row, scoreFetchText);
        const llmPrompt = raw_buildLlmScorePrompt_(prompt, article);
        promptChars += llmPrompt.length;
        const responseText = aiGenerateJson_(llmPrompt, { maxOutputTokens: 600 });
        const parsed = feeds_safeParseJsonObject_(responseText);
        if (!parsed) {
          errors.push(`No JSON parsed for: ${article.url || article.title}`);
          continue;
        }

        const finalScore = Number(parsed.final_score);
        const llm = Object.assign({}, parsed, {
          final_score: Number.isFinite(finalScore) ? finalScore : 0
        });

        scored.push({
          key: row.link || row.title,
          title: article.title,
          url: article.url,
          llm,
          row_index: i
        });
      } catch (err) {
        errors.push(String(err?.message || err));
      }
      nextIndex = i + 1;
    }

    if (nextIndex < rows.length) {
      raw_setLlmRankProgress_({
        prompt,
        topN,
        fetchText,
        nextIndex,
        promptChars,
        scored,
        errors: errors.slice(0, 50)
      });
      return {
        ok: true,
        status: "in_progress",
        meta: {
          prompt,
          requested_top_n: topN,
          scored_total: scored.length,
          total_rows: rows.length,
          next_index: nextIndex,
          prompt_chars_total: promptChars,
          tokens_estimate: estimateTokensFromChars_(promptChars),
          source: "Raw Articles",
          savedAt: new Date().toISOString()
        },
        results: [],
        errors
      };
    }

    const topCandidates = scored
      .slice()
      .sort((a, b) => (Number(b?.llm?.final_score || 0) - Number(a?.llm?.final_score || 0)))
      .slice(0, topN);

    topCandidates.forEach((candidate) => {
      try {
        const article = raw_buildLlmArticle_(rows[candidate.row_index], fetchText);
        const llmPrompt = raw_buildLlmPrompt_(prompt, article);
        promptChars += llmPrompt.length;
        const responseText = aiGenerateJson_(llmPrompt, { maxOutputTokens: 1200 });
        const parsed = feeds_safeParseJsonObject_(responseText);
        if (!parsed) {
          errors.push(`No JSON parsed for: ${article.url || article.title}`);
          return;
        }

        const finalScore = Number(parsed.final_score);
        const llm = Object.assign({}, parsed, {
          final_score: Number.isFinite(finalScore) ? finalScore : 0
        });

        results.push({
          key: candidate.key,
          title: article.title,
          url: article.url,
          llm
        });
      } catch (err) {
        errors.push(String(err?.message || err));
      }
    });

    const meta = {
      prompt,
      requested_top_n: topN,
      scored_total: scored.length,
      scored_top_n: results.length,
      total_rows: rows.length,
      score_fetch_text: scoreFetchText,
      top_n_fetch_text: fetchText,
      prompt_chars_total: promptChars,
      tokens_estimate: estimateTokensFromChars_(promptChars),
      source: "Raw Articles",
      savedAt: new Date().toISOString()
    };
    raw_clearLlmRankProgress_();
    raw_saveLlmRankCache_({ results, meta });

    return {
      ok: true,
      status: "done",
      meta,
      results,
      errors
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function ui_getRawArticlesLlmRankCache_v1() {
  try {
    const sheetPayload = raw_readLlmRankSheet_();
    if (sheetPayload) {
      return {
        ok: true,
        results: sheetPayload.results,
        meta: sheetPayload.meta
      };
    }

    const raw = PropertiesService.getScriptProperties().getProperty(RAW_LLM_RANK_CACHE_KEY);
    if (!raw) return { ok: true, results: [], meta: null };
    const parsed = JSON.parse(raw);
    return {
      ok: true,
      results: Array.isArray(parsed?.results) ? parsed.results : [],
      meta: parsed?.meta || null
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function ui_clearRawArticlesLlmRankCache_v1() {
  raw_clearLlmRankSheets_();
  PropertiesService.getScriptProperties().deleteProperty(RAW_LLM_RANK_CACHE_KEY);
  return { ok: true };
}

function estimateTokensFromChars_(chars) {
  const n = Number(chars || 0);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return Math.max(1, Math.round(n / 4));
}

function raw_saveLlmRankCache_(payload) {
  const safe = payload || {};
  const serialized = JSON.stringify({
    results: Array.isArray(safe.results) ? safe.results : [],
    meta: safe.meta || null
  });
  PropertiesService.getScriptProperties().setProperty(RAW_LLM_RANK_CACHE_KEY, serialized);
  raw_writeLlmRankSheet_(safe);
}

function raw_getLlmRankProgress_() {
  const raw = PropertiesService.getScriptProperties().getProperty(RAW_LLM_RANK_PROGRESS_KEY);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function raw_setLlmRankProgress_(payload) {
  const safe = payload || {};
  const serialized = JSON.stringify({
    prompt: safe.prompt || "",
    topN: Number(safe.topN || 0),
    fetchText: safe.fetchText === true,
    nextIndex: Number(safe.nextIndex || 0),
    promptChars: Number(safe.promptChars || 0),
    scored: Array.isArray(safe.scored) ? safe.scored : [],
    errors: Array.isArray(safe.errors) ? safe.errors.slice(0, 50) : []
  });
  PropertiesService.getScriptProperties().setProperty(RAW_LLM_RANK_PROGRESS_KEY, serialized);
}

function raw_clearLlmRankProgress_() {
  PropertiesService.getScriptProperties().deleteProperty(RAW_LLM_RANK_PROGRESS_KEY);
}

function raw_getOrCreateSheet_(name) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function raw_clearLlmRankSheets_() {
  const ss = getSpreadsheet_();
  const cacheSheet = ss.getSheetByName(RAW_LLM_RANK_SHEET);
  if (cacheSheet) cacheSheet.clearContents();
  const metaSheet = ss.getSheetByName(RAW_LLM_RANK_META_SHEET);
  if (metaSheet) metaSheet.clearContents();
}

function raw_writeLlmRankSheet_(payload) {
  const results = Array.isArray(payload?.results) ? payload.results : [];
  const meta = payload?.meta || null;

  const cacheSheet = raw_getOrCreateSheet_(RAW_LLM_RANK_SHEET);
  const headers = ["key", "title", "url", "llm_json"];
  cacheSheet.clearContents();
  cacheSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (results.length) {
    const rows = results.map((item) => ([
      item?.key || "",
      item?.title || "",
      item?.url || "",
      JSON.stringify(item?.llm || {})
    ]));
    cacheSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  const metaSheet = raw_getOrCreateSheet_(RAW_LLM_RANK_META_SHEET);
  metaSheet.clearContents();
  metaSheet.getRange(1, 1, 1, 2).setValues([["key", "value"]]);
  if (meta) {
    const metaRows = Object.keys(meta).map((key) => [key, String(meta[key])]);
    if (metaRows.length) metaSheet.getRange(2, 1, metaRows.length, 2).setValues(metaRows);
  }
}

function raw_readLlmRankSheet_() {
  const ss = getSpreadsheet_();
  const cacheSheet = ss.getSheetByName(RAW_LLM_RANK_SHEET);
  if (!cacheSheet) return null;
  const lastRow = cacheSheet.getLastRow();
  const lastCol = cacheSheet.getLastColumn();
  if (lastRow < 2 || lastCol < 4) return null;

  const values = cacheSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map((h) => String(h || "").trim());
  const idx = {
    key: headers.indexOf("key"),
    title: headers.indexOf("title"),
    url: headers.indexOf("url"),
    llm: headers.indexOf("llm_json")
  };

  const results = [];
  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    const key = idx.key >= 0 ? row[idx.key] : "";
    if (!key) continue;
    const llmRaw = idx.llm >= 0 ? String(row[idx.llm] || "") : "";
    let llm = {};
    if (llmRaw) {
      try {
        llm = JSON.parse(llmRaw);
      } catch (err) {
        llm = {};
      }
    }
    results.push({
      key: String(key),
      title: idx.title >= 0 ? String(row[idx.title] || "") : "",
      url: idx.url >= 0 ? String(row[idx.url] || "") : "",
      llm
    });
  }

  let meta = null;
  const metaSheet = ss.getSheetByName(RAW_LLM_RANK_META_SHEET);
  if (metaSheet && metaSheet.getLastRow() >= 2) {
    const metaValues = metaSheet.getRange(1, 1, metaSheet.getLastRow(), 2).getValues();
    meta = {};
    for (let i = 1; i < metaValues.length; i += 1) {
      const key = String(metaValues[i][0] || "").trim();
      if (!key) continue;
      meta[key] = String(metaValues[i][1] || "");
    }
  }

  return { results, meta };
}

function weekly_buildLlmArticle_(row, fetchText) {
  const title = String(row?.title || "").trim();
  const url = String(row?.link || row?.url || "").trim();
  const date = String(row?.date || row?.dateDisplay || "").trim();
  const source = String(row?.source || "").trim();
  const theme = String(row?.theme || "").trim();
  const poi = String(row?.poi || "").trim();
  const snippet = String(row?.snippet || "").trim();

  const fallbackParts = [];
  if (theme) fallbackParts.push(`Theme: ${theme}`);
  if (poi) fallbackParts.push(`POI: ${poi}`);
  if (snippet) fallbackParts.push(`Snippet: ${snippet}`);
  const fallbackSnippet = fallbackParts.join(" | ");

  let textSnippet = "";

  if (fetchText && url) {
    try {
      const meta = fetchArticleMeta_(url);
      const desc = String(meta?.description || "").trim();
      const text = String(meta?.text || "").trim();
      textSnippet = [desc, text].filter(Boolean).join("\n\n");
    } catch (err) {
      textSnippet = "";
    }
  }

  if (!textSnippet) textSnippet = fallbackSnippet || title;

  return {
    title,
    url,
    date,
    source,
    theme,
    poi,
    text: textSnippet
  };
}

function raw_buildLlmArticle_(row, fetchText) {
  const title = String(row?.title || "").trim();
  const url = String(row?.link || row?.url || "").trim();
  const date = String(row?.date || row?.dateDisplay || "").trim();
  const source = String(row?.source || "").trim();
  const theme = String(row?.theme || "").trim();
  const poi = String(row?.poi || "").trim();
  const keywords = String(row?.keywords || "").trim();

  const fallbackParts = [];
  if (theme) fallbackParts.push(`Theme: ${theme}`);
  if (poi) fallbackParts.push(`POI: ${poi}`);
  if (keywords) fallbackParts.push(`Keywords: ${keywords}`);
  const fallbackSnippet = fallbackParts.join(" | ");

  let textSnippet = "";

  if (fetchText && url) {
    try {
      const meta = fetchArticleMeta_(url);
      const desc = String(meta?.description || "").trim();
      const text = String(meta?.text || "").trim();
      textSnippet = [desc, text].filter(Boolean).join("\n\n");
    } catch (err) {
      textSnippet = "";
    }
  }

  if (!textSnippet) textSnippet = fallbackSnippet || title;

  return {
    title,
    url,
    date,
    source,
    theme,
    poi,
    text: textSnippet
  };
}

function raw_buildLlmPrompt_(userPrompt, article) {
  const clippedText = clampText_(article.text, RAW_LLM_MAX_TEXT_CHARS);
  return [
    RAW_LLM_GUIDE_TEXT,
    "",
    `User prompt (sorting intent): ${userPrompt}`,
    "",
    "Article:",
    `Title: ${article.title}`,
    `URL: ${article.url}`,
    `Date: ${article.date}`,
    `Source: ${article.source}`,
    `Theme (preassigned): ${article.theme}`,
    `POI (preassigned): ${article.poi}`,
    `Text/snippet: ${clippedText}`
  ].join("\n");
}

function raw_buildLlmScorePrompt_(userPrompt, article) {
  const clippedText = clampText_(article.text, RAW_LLM_SCORE_MAX_TEXT_CHARS);
  return [
    RAW_LLM_SCORE_ONLY_GUIDE_TEXT,
    "",
    `User prompt (sorting intent): ${userPrompt}`,
    "",
    "Article:",
    `Title: ${article.title}`,
    `URL: ${article.url}`,
    `Date: ${article.date}`,
    `Source: ${article.source}`,
    `Theme (preassigned): ${article.theme}`,
    `POI (preassigned): ${article.poi}`,
    `Text/snippet: ${clippedText}`
  ].join("\n");
}

function weekly_buildLlmPrompt_(userPrompt, article) {
  const clippedText = clampText_(article.text, RAW_LLM_MAX_TEXT_CHARS);
  return [
    WEEKLY_PICKS_LLM_GUIDE_TEXT,
    "",
    `User prompt (sorting intent): ${userPrompt}`,
    "",
    "Article:",
    `Title: ${article.title}`,
    `URL: ${article.url}`,
    `Date: ${article.date}`,
    `Source: ${article.source}`,
    `Theme (preassigned): ${article.theme}`,
    `POI (preassigned): ${article.poi}`,
    `Text/snippet: ${clippedText}`
  ].join("\n");
}

function clampText_(text, maxChars) {
  const limit = Number(maxChars);
  const raw = String(text || "");
  if (!Number.isFinite(limit) || limit <= 0 || raw.length <= limit) return raw;
  return raw.slice(0, limit).trim() + "…";
}
