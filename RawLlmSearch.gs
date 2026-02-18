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
const RAW_LLM_RANK_PROGRESS_SHEET = "Raw LLM Rank Progress";
const RAW_LLM_SCORE_COL = "LLM Score";
const RAW_LLM_REC_COL = "LLM Recommendation";
const RAW_LLM_SUMMARY_COL = "LLM Summary";
const RAW_LLM_REASONS_COL = "LLM Reasons";

function raw_isLlmRecommended_(llm) {
  const raw = String(llm?.publish_recommendation || llm?.recommendation || "").trim().toLowerCase();
  if (!raw) return false;
  return ["publish", "maybe", "recommend", "recommended", "feature", "include", "yes", "y"].some((v) => raw === v || raw.includes(v));
}

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

    const topN = 0;
    const fetchText = payload?.fetchText === true;
    const scoreFetchText = false;
    const sectorBy = raw_normalizeSectorBy_(payload?.sectorBy);

    const raw = ui_getRawArticles_bootstrap_v1();
    if (!raw || raw.ok !== true) {
      return { ok: false, message: raw?.message || "Failed to load raw articles." };
    }
    const rows = Array.isArray(raw.rows) ? raw.rows : [];

    const existingCachePayload = raw_readLlmRankSheet_();
    const existingResults = Array.isArray(existingCachePayload?.results) ? existingCachePayload.results : [];
    const scoredByLink = new Map();
    existingResults.forEach((item) => {
      const linkNorm = feeds_normalizeLink_(item?.url || item?.link);
      if (linkNorm) scoredByLink.set(linkNorm, item);
    });

    const pendingRows = [];
    const scored = [];
    const scoredKeys = new Set();
    const pushScored = (item) => {
      const rowIndex = Number(item?.row_index);
      const keyBase = Number.isFinite(rowIndex) ? `idx:${rowIndex}` : `key:${String(item?.key || "")}`;
      if (scoredKeys.has(keyBase)) return;
      scoredKeys.add(keyBase);
      scored.push(item);
    };

    rows.forEach((row, idx) => {
      const linkNorm = feeds_normalizeLink_(row?.link || row?.url);
      const cached = (linkNorm && scoredByLink.get(linkNorm)) || null;
      const cachedLlm = cached?.llm || {};
      const cachedHasScore = Number.isFinite(Number(cachedLlm?.final_score));
      const cachedHasRec = String(cachedLlm?.publish_recommendation || cachedLlm?.recommendation || "").trim().length > 0;
      const cachedHasSummary = String(cachedLlm?.summary || "").trim().length > 0;

      if (cached && cachedHasScore && cachedHasRec && cachedHasSummary) {
        pushScored({
          key: row.link || row.title,
          title: String(cached?.title || row?.title || ""),
          url: String(cached?.url || row?.link || ""),
          theme: String(cached?.theme || row?.theme || ""),
          poi: String(cached?.poi || row?.poi || ""),
          llm: cached?.llm || {},
          row_index: idx
        });
      } else {
        pendingRows.push({ row, row_index: idx, sector: raw_getRowSectorKey_(row, sectorBy) });
      }
    });

    const startedAt = Date.now();
    const rawBudgetMs = Number(payload?.timeBudgetMs || 0);
    const effectiveTimeBudgetMs = Number.isFinite(rawBudgetMs) && rawBudgetMs > 0
      ? Math.max(10000, Math.min(rawBudgetMs, RAW_LLM_RANK_TIME_BUDGET_MS))
      : RAW_LLM_RANK_TIME_BUDGET_MS;
    const errors = [];
    const progress = raw_getLlmRankProgress_();
    const sameJob = progress &&
      progress.prompt === prompt &&
      progress.topN === topN &&
      progress.fetchText === fetchText &&
      raw_normalizeSectorBy_(progress.sectorBy) === sectorBy;

    if (!sameJob && progress) {
      raw_clearLlmRankProgress_();
    }

    const progressScored = sameJob && Array.isArray(progress?.scored) ? progress.scored : [];
    if (progressScored.length) progressScored.forEach((item) => pushScored(item));

    let promptChars = sameJob ? Number(progress?.promptChars || 0) : 0;

    const sectors = new Map();
    pendingRows.forEach((entry) => {
      const key = String(entry?.sector || "Unspecified");
      if (!sectors.has(key)) sectors.set(key, []);
      sectors.get(key).push(entry);
    });
    const sectorOrder = Array.from(sectors.keys()).sort((a, b) => a.localeCompare(b));

    let sectorCursor = sameJob ? Number(progress?.sectorCursor || 0) : 0;
    let nextIndex = sameJob ? Number(progress?.nextIndex || 0) : 0;
    if (!Number.isFinite(sectorCursor) || sectorCursor < 0) sectorCursor = 0;
    if (!Number.isFinite(nextIndex) || nextIndex < 0) nextIndex = 0;

    while (sectorCursor < sectorOrder.length) {
      const currentSector = sectorOrder[sectorCursor] || "Unspecified";
      const sectorRows = sectors.get(currentSector) || [];
      if (nextIndex >= sectorRows.length) {
        sectorCursor += 1;
        nextIndex = 0;
        continue;
      }

      for (let i = nextIndex; i < sectorRows.length; i += 1) {
        if (Date.now() - startedAt > effectiveTimeBudgetMs) {
          nextIndex = i;
          break;
        }

        const entry = sectorRows[i] || {};
        const row = entry.row || {};
        try {
          const article = raw_buildLlmArticle_(row, scoreFetchText);
          const llmPrompt = raw_buildLlmScorePrompt_(prompt, article);
          promptChars += llmPrompt.length;
          const responseText = aiGenerateJson_(llmPrompt, { maxOutputTokens: 1200 });
          const parsed = feeds_safeParseJsonObject_(responseText);
          if (!parsed) {
            errors.push(`[${currentSector}] No JSON parsed for: ${article.url || article.title}`);
            continue;
          }

          const finalScore = Number(parsed.final_score);
          const llm = Object.assign({}, parsed, {
            final_score: Number.isFinite(finalScore) ? finalScore : 0
          });

          const scoredRowIndex = Number(entry.row_index || 0);
          pushScored({
            key: row.link || row.title,
            title: article.title,
            url: article.url,
            theme: article.theme,
            poi: article.poi,
            llm,
            row_index: scoredRowIndex
          });
        } catch (err) {
          errors.push(`[${currentSector}] ${String(err?.message || err)}`);
        }
        nextIndex = i + 1;
      }

      if (Date.now() - startedAt > effectiveTimeBudgetMs) {
        break;
      }

      if (nextIndex >= sectorRows.length) {
        sectorCursor += 1;
        nextIndex = 0;
      }
    }

    if (sectorCursor < sectorOrder.length) {
      const currentSector = sectorOrder[sectorCursor] || "Unspecified";
      try {
        raw_writeScoredLlmToRawSheet_(scored);
      } catch (err) {
        errors.push(`Sheet write warning: ${String(err?.message || err)}`);
      }
      raw_setLlmRankProgress_({
        prompt,
        topN,
        fetchText,
        sectorBy,
        sectorCursor,
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
          requested_top_n: topN || null,
          sector_by: sectorBy,
          sector_current: currentSector,
          sector_cursor: sectorCursor,
          sector_total: sectorOrder.length,
          sector_pending_rows: (sectors.get(currentSector) || []).length,
          recommended_total: scored.filter((item) => raw_isLlmRecommended_(item?.llm || {})).length,
          scored_total: scored.length,
          total_rows: rows.length,
          pending_rows: pendingRows.length,
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

    const recommendedCandidates = scored
      .filter((item) => raw_isLlmRecommended_(item?.llm || {}))
      .sort((a, b) => (Number(b?.llm?.final_score || 0) - Number(a?.llm?.final_score || 0)));

    const results = recommendedCandidates.map((item) => ({
      key: item?.key || "",
      title: item?.title || "",
      url: item?.url || "",
      theme: item?.theme || "",
      poi: item?.poi || "",
      llm: item?.llm || {}
    }));

    const meta = {
      prompt,
      requested_top_n: topN || null,
      sector_by: sectorBy,
      sector_total: sectorOrder.length,
      recommended_total: recommendedCandidates.length,
      scored_total: scored.length,
      scored_top_n: results.length,
      total_rows: rows.length,
      pending_rows: pendingRows.length,
      score_fetch_text: scoreFetchText,
      top_n_fetch_text: false,
      prompt_chars_total: promptChars,
      tokens_estimate: estimateTokensFromChars_(promptChars),
      source: "Raw Articles",
      savedAt: new Date().toISOString()
    };
    try {
      raw_writeScoredLlmToRawSheet_(scored);
    } catch (err) {
      errors.push(`Sheet write warning: ${String(err?.message || err)}`);
    }
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

function ui_runRawArticleLlmThoughts_v1(payload) {
  try {
    const id = Number(payload?.id || 0);
    if (!Number.isFinite(id) || id < 2) return { ok: false, message: "Invalid article id." };

    const prompt = String(payload?.prompt || "").trim() ||
      "Score each article for executive horizon-scanning relevance and set publish_recommendation to publish, maybe, or skip.";
    const fetchText = payload?.fetchText === true;

    const raw = ui_getRawArticles_bootstrap_v1();
    if (!raw || raw.ok !== true) {
      return { ok: false, message: raw?.message || "Failed to load raw articles." };
    }

    const rows = Array.isArray(raw.rows) ? raw.rows : [];
    const rowIndex = rows.findIndex((r) => Number(r?.id || 0) === id);
    if (rowIndex < 0) return { ok: false, message: `Article not found for id ${id}.` };

    const row = rows[rowIndex] || {};
    const article = raw_buildLlmArticle_(row, fetchText);
    const llmPrompt = raw_buildLlmPrompt_(prompt, article);
    const responseText = aiGenerateJson_(llmPrompt, { maxOutputTokens: 1200 });
    const parsed = feeds_safeParseJsonObject_(responseText);
    if (!parsed) return { ok: false, message: "No JSON parsed from LLM response." };

    const finalScore = Number(parsed.final_score);
    const llm = Object.assign({}, parsed, {
      final_score: Number.isFinite(finalScore) ? finalScore : 0
    });

    const scoredItem = {
      key: row.link || row.title,
      title: article.title,
      url: article.url,
      theme: article.theme,
      poi: article.poi,
      llm,
      row_index: rowIndex
    };

    raw_writeScoredLlmToRawSheet_([scoredItem]);

    const llmRecommendation = String(llm.publish_recommendation || llm.recommendation || "").trim();
    const llmSummary = String(llm.summary || "").trim();
    const llmReasons = Array.isArray(llm.score_reasons)
      ? llm.score_reasons.map((x) => String(x || "").trim()).filter(Boolean)
      : [];

    return {
      ok: true,
      row: {
        id,
        llmScore: Number(llm.final_score || 0),
        llmRecommendation,
        llmRecommended: raw_isLlmRecommended_(llm),
        llmSummary,
        llmReasons
      },
      meta: {
        source: "Raw Articles",
        updated_row_id: id,
        savedAt: new Date().toISOString()
      }
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function raw_normalizeSectorBy_(value) {
  const key = String(value || "theme").trim().toLowerCase();
  if (["theme", "poi", "source"].includes(key)) return key;
  return "theme";
}

function raw_getRowSectorKey_(row, sectorBy) {
  const safe = row || {};
  if (sectorBy === "poi") return String(safe.poi || "Unspecified").trim() || "Unspecified";
  if (sectorBy === "source") return String(safe.source || "Unspecified").trim() || "Unspecified";
  return String(safe.theme || "Unspecified").trim() || "Unspecified";
}

function raw_writeScoredLlmToRawSheet_(scoredItems) {
  const items = Array.isArray(scoredItems) ? scoredItems : [];
  if (!items.length) return;

  const ss = getSpreadsheet_();
  const sheetName = (typeof getRawArticlesSheetName_ === "function") ? getRawArticlesSheetName_() : "Raw Articles";
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const lastCol = Math.max(1, sh.getLastColumn());
  if (lastRow < 1) return;

  const headerValues = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map((v) => String(v || "").trim());
  const llmCols = raw_ensureLlmColumns_(sh, headerValues);
  const rowColCount = sh.getLastColumn();

  const bodyRows = Math.max(0, sh.getLastRow() - 1);
  if (!bodyRows) return;
  const body = sh.getRange(2, 1, bodyRows, rowColCount).getDisplayValues();
  const rowByLink = new Map();

  body.forEach((row, idx) => {
    const rowNo = idx + 2;
    const link = String((row[1] || "")).trim();
    const linkNorm = feeds_normalizeLink_(link);
    if (linkNorm && !rowByLink.has(linkNorm)) rowByLink.set(linkNorm, rowNo);
  });

  const updates = [];
  items.forEach((item) => {
    const llm = item?.llm || {};
    const linkNorm = feeds_normalizeLink_(item?.url || item?.link);
    const rowNo = (linkNorm && rowByLink.get(linkNorm));
    if (!rowNo) return;

    const reasons = Array.isArray(llm.score_reasons)
      ? llm.score_reasons.map((x) => String(x || "").trim()).filter(Boolean).join(" | ")
      : "";
    const scoreNum = Number(llm.final_score);
    const score = Number.isFinite(scoreNum) ? scoreNum : "";
    const recommendation = String(llm.publish_recommendation || llm.recommendation || "").trim();
    const summary = String(llm.summary || "").trim();

    updates.push({
      rowNo,
      score,
      recommendation,
      summary,
      reasons
    });
  });

  updates.forEach((u) => {
    sh.getRange(u.rowNo, llmCols.score, 1, 1).setValue(u.score);
    sh.getRange(u.rowNo, llmCols.recommendation, 1, 1).setValue(u.recommendation);
    sh.getRange(u.rowNo, llmCols.summary, 1, 1).setValue(u.summary);
    sh.getRange(u.rowNo, llmCols.reasons, 1, 1).setValue(u.reasons);
  });
}

function raw_ensureLlmColumns_(sheet, headerValues) {
  const sh = sheet;
  let headers = Array.isArray(headerValues) ? headerValues.slice() : [];
  const norm = (s) => String(s || "").trim().toLowerCase().replace(/\s+/g, " ");

  const findIdx = (names) => {
    const targets = names.map(norm);
    for (let i = 0; i < headers.length; i += 1) {
      if (targets.includes(norm(headers[i]))) return i + 1;
    }
    return 0;
  };

  const appendCol = (name) => {
    const col = headers.length + 1;
    sh.getRange(1, col).setValue(name);
    headers.push(name);
    return col;
  };

  let scoreCol = findIdx([RAW_LLM_SCORE_COL, "llm_score"]);
  if (!scoreCol) scoreCol = appendCol(RAW_LLM_SCORE_COL);

  let recCol = findIdx([RAW_LLM_REC_COL, "llm_recommendation", "publish_recommendation"]);
  if (!recCol) recCol = appendCol(RAW_LLM_REC_COL);

  let summaryCol = findIdx([RAW_LLM_SUMMARY_COL, "llm_summary"]);
  if (!summaryCol) summaryCol = appendCol(RAW_LLM_SUMMARY_COL);

  let reasonsCol = findIdx([RAW_LLM_REASONS_COL, "llm_reasons", "score_reasons"]);
  if (!reasonsCol) reasonsCol = appendCol(RAW_LLM_REASONS_COL);

  return {
    score: scoreCol,
    recommendation: recCol,
    summary: summaryCol,
    reasons: reasonsCol
  };
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
  try {
    PropertiesService.getScriptProperties().setProperty(RAW_LLM_RANK_CACHE_KEY, serialized);
  } catch (err) {
    // Large ranking payloads can exceed Script Properties size limits.
    // Sheet persistence is the source of truth and supports larger volumes.
  }
  raw_writeLlmRankSheet_(safe);
}

function raw_getLlmRankProgress_() {
  const sheetPayload = raw_readLlmRankProgressSheet_();
  if (sheetPayload) return sheetPayload;

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
  const progressPayload = {
    prompt: safe.prompt || "",
    topN: Number(safe.topN || 0),
    fetchText: safe.fetchText === true,
    sectorBy: raw_normalizeSectorBy_(safe.sectorBy),
    sectorCursor: Number(safe.sectorCursor || 0),
    nextIndex: Number(safe.nextIndex || 0),
    promptChars: Number(safe.promptChars || 0),
    scored: Array.isArray(safe.scored) ? safe.scored : [],
    errors: Array.isArray(safe.errors) ? safe.errors.slice(0, 50) : []
  };

  raw_writeLlmRankProgressSheet_(progressPayload);

  const serialized = JSON.stringify(Object.assign({}, progressPayload, {
    scored: []
  }));
  try {
    PropertiesService.getScriptProperties().setProperty(RAW_LLM_RANK_PROGRESS_KEY, serialized);
  } catch (err) {
    // Best-effort only. Sheet persistence is the source of truth.
  }
}

function raw_clearLlmRankProgress_() {
  raw_clearLlmRankProgressSheet_();
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
  raw_clearLlmRankProgressSheet_();
}

function raw_clearLlmRankProgressSheet_() {
  const ss = getSpreadsheet_();
  const progressSheet = ss.getSheetByName(RAW_LLM_RANK_PROGRESS_SHEET);
  if (progressSheet) progressSheet.clearContents();
}

function raw_writeLlmRankProgressSheet_(payload) {
  const safe = payload || {};
  const scored = Array.isArray(safe.scored) ? safe.scored : [];

  const progressSheet = raw_getOrCreateSheet_(RAW_LLM_RANK_PROGRESS_SHEET);
  progressSheet.clearContents();

  const headers = [
    "prompt",
    "topN",
    "fetchText",
    "sectorBy",
    "sectorCursor",
    "nextIndex",
    "promptChars",
    "errors_json"
  ];
  progressSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  progressSheet.getRange(2, 1, 1, headers.length).setValues([[
    String(safe.prompt || ""),
    Number(safe.topN || 0),
    safe.fetchText === true,
    raw_normalizeSectorBy_(safe.sectorBy),
    Number(safe.sectorCursor || 0),
    Number(safe.nextIndex || 0),
    Number(safe.promptChars || 0),
    JSON.stringify(Array.isArray(safe.errors) ? safe.errors.slice(0, 50) : [])
  ]]);

  const scoredHeaders = ["key", "title", "url", "theme", "poi", "row_index", "llm_json"];
  progressSheet.getRange(4, 1, 1, scoredHeaders.length).setValues([scoredHeaders]);
  if (scored.length) {
    const scoredRows = scored.map((item) => ([
      item?.key || "",
      item?.title || "",
      item?.url || "",
      item?.theme || "",
      item?.poi || "",
      Number(item?.row_index || 0),
      JSON.stringify(item?.llm || {})
    ]));
    progressSheet.getRange(5, 1, scoredRows.length, scoredHeaders.length).setValues(scoredRows);
  }
}

function raw_readLlmRankProgressSheet_() {
  const ss = getSpreadsheet_();
  const progressSheet = ss.getSheetByName(RAW_LLM_RANK_PROGRESS_SHEET);
  if (!progressSheet || progressSheet.getLastRow() < 2) return null;

  const metaValues = progressSheet.getRange(2, 1, 1, 8).getValues()[0] || [];
  const prompt = String(metaValues[0] || "");
  if (!prompt) return null;

  let errors = [];
  const errorsRaw = String(metaValues[7] || "").trim();
  if (errorsRaw) {
    try {
      const parsed = JSON.parse(errorsRaw);
      errors = Array.isArray(parsed) ? parsed : [];
    } catch (err) {
      errors = [];
    }
  }

  const scored = [];
  if (progressSheet.getLastRow() >= 5) {
    const scoredRows = progressSheet.getRange(5, 1, progressSheet.getLastRow() - 4, 7).getValues();
    scoredRows.forEach((row) => {
      const key = String(row[0] || "").trim();
      const llmRaw = String(row[6] || "").trim();
      if (!key && !llmRaw) return;

      let llm = {};
      if (llmRaw) {
        try {
          llm = JSON.parse(llmRaw);
        } catch (err) {
          llm = {};
        }
      }

      scored.push({
        key,
        title: String(row[1] || ""),
        url: String(row[2] || ""),
        theme: String(row[3] || ""),
        poi: String(row[4] || ""),
        row_index: Number(row[5] || 0),
        llm
      });
    });
  }

  return {
    prompt,
    topN: Number(metaValues[1] || 0),
    fetchText: metaValues[2] === true || String(metaValues[2]).toLowerCase() === "true",
    sectorBy: raw_normalizeSectorBy_(metaValues[3]),
    sectorCursor: Number(metaValues[4] || 0),
    nextIndex: Number(metaValues[5] || 0),
    promptChars: Number(metaValues[6] || 0),
    scored,
    errors
  };
}

function raw_writeLlmRankSheet_(payload) {
  const results = Array.isArray(payload?.results) ? payload.results : [];
  const meta = payload?.meta || null;

  const cacheSheet = raw_getOrCreateSheet_(RAW_LLM_RANK_SHEET);
  const headers = ["key", "title", "url", "theme", "poi", "llm_json"];
  cacheSheet.clearContents();
  cacheSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (results.length) {
    const rows = results.map((item) => ([
      item?.key || "",
      item?.title || "",
      item?.url || "",
      item?.theme || "",
      item?.poi || "",
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
    theme: headers.indexOf("theme"),
    poi: headers.indexOf("poi"),
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
      theme: idx.theme >= 0 ? String(row[idx.theme] || "") : "",
      poi: idx.poi >= 0 ? String(row[idx.poi] || "") : "",
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
