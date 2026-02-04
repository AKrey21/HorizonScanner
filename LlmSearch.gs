/*********************************
 * RawLlmSearch.gs — LLM-based ranking for Raw Articles
 * - Uses AIService.aiGenerateJson_()
 * - Optionally fetches article text
 *********************************/

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

function ui_runRawArticlesLlmRank_v2(payload) {
  try {
    const prompt = String(payload?.prompt || "").trim();
    if (!prompt) return { ok: false, message: "Missing prompt." };

    const maxRows = Math.max(1, Math.min(60, Number(payload?.maxRows || 20)));
    const fetchText = payload?.fetchText === true;

    const raw = ui_getRawArticles_bootstrap_v1();
    if (!raw || raw.ok !== true) {
      return { ok: false, message: raw?.message || "Failed to load raw articles." };
    }
    const rows = Array.isArray(raw.rows) ? raw.rows : [];

    const results = [];
    const errors = [];
    const target = rows.slice(0, maxRows);
    let promptChars = 0;

    target.forEach((row) => {
      try {
        const article = raw_buildLlmArticle_(row, fetchText);
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
          key: row.link || row.title,
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
        tokens_estimate: estimateTokensFromChars_(promptChars),
        source: "Raw Articles"
      },
      results,
      errors
    };
  } catch (err) {
    return { ok: false, message: err?.message || String(err) };
  }
}

function estimateTokensFromChars_(chars) {
  const n = Number(chars || 0);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return Math.max(1, Math.round(n / 4));
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
    `Text/snippet: ${article.text}`
  ].join("\n");
}
