/**************************************
 * AIService.gs — AI helpers + AI scoring config generator
 *
 * What this file includes:
 * 1) aiGenerateText_() — provider-aware caller (Gemini/OpenAI)
 * 2) geminiGenerateText_() — direct Gemini caller (optionally schema-enforced JSON)
 * 3) openaiGenerateText_() — direct OpenAI caller (optionally schema-enforced JSON)
 * 4) aiGenerateJson_() — helper for JSON responses
 * 5) ui_generateScoringConfigFromAI_v1() — UI endpoint for config generation
 * 6) ai_generateScoringConfig_() + full sanitiser + helpers
 *
 * Dependencies expected elsewhere in your project:
 * - sc_getConfigSafe_()   (returns baseline config object)
 * - CONTROL_SHEET constant (e.g. "ThemeRules")
 **************************************/

const AI_PROVIDER_DEFAULT = "gemini";
const AI_DEFAULT_MAX_OUTPUT_TOKENS = 600; // conservative default
const GEMINI_DEFAULT_MODEL = "gemini-2.0-flash";
const OPENAI_DEFAULT_MODEL = "gpt-4o-mini";

function getAiProvider_() {
  const raw = PropertiesService.getScriptProperties().getProperty("AI_PROVIDER");
  const provider = String(raw || AI_PROVIDER_DEFAULT).trim().toLowerCase();
  return (provider === "openai") ? "openai" : "gemini";
}

function getAiModel_(provider, overrideModel) {
  if (overrideModel) return overrideModel;
  const props = PropertiesService.getScriptProperties();
  const shared = props.getProperty("AI_MODEL");
  const providerKey = (provider === "openai") ? "OPENAI_MODEL" : "GEMINI_MODEL";
  const providerModel = props.getProperty(providerKey);
  if (providerModel) return providerModel;
  if (shared) return shared;
  return (provider === "openai") ? OPENAI_DEFAULT_MODEL : GEMINI_DEFAULT_MODEL;
}

function getAiMaxOutputTokens_(requested) {
  const props = PropertiesService.getScriptProperties();
  const capRaw = Number(props.getProperty("AI_MAX_OUTPUT_TOKENS"));
  const cap = Number.isFinite(capRaw) ? capRaw : null;
  const fallback = AI_DEFAULT_MAX_OUTPUT_TOKENS;
  let value = (requested != null) ? Number(requested) : fallback;
  if (!Number.isFinite(value)) value = fallback;
  if (cap && value > cap) value = cap;
  return Math.max(1, Math.round(value));
}

/* =========================
 * Gemini helpers
 * ========================= */

function getGeminiApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!key) throw new Error('Missing Script Property "GEMINI_API_KEY".');
  return key;
}

function getOpenAiApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) throw new Error('Missing Script Property "OPENAI_API_KEY".');
  return key;
}

/**
 * Generic Gemini call.
 * Returns model text response (string).
 *
 * opts:
 * - model (string) default GEMINI_DEFAULT_MODEL
 * - temperature (number)
 * - maxOutputTokens (int)
 * - responseMimeType ("application/json" etc)
 * - responseSchema (object) — ONLY when responseMimeType is application/json (schema-enforced)
 */
function geminiGenerateText_(prompt, opts) {
  opts = opts || {};
  const model = getAiModel_("gemini", opts.model);

  const payload = {
    contents: [{ role: "user", parts: [{ text: String(prompt || "") }] }],
    generationConfig: {
      temperature: (opts.temperature != null) ? opts.temperature : 0.4,
      maxOutputTokens: getAiMaxOutputTokens_(opts.maxOutputTokens),
      responseMimeType: opts.responseMimeType,
      responseSchema: opts.responseSchema
    }
  };

  // Remove undefined/null fields that Gemini may reject
  if (!payload.generationConfig.responseMimeType) delete payload.generationConfig.responseMimeType;
  if (!payload.generationConfig.responseSchema) delete payload.generationConfig.responseSchema;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`;
  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { "x-goog-api-key": getGeminiApiKey_() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code >= 300) throw new Error(`Gemini error ${code}: ${body}`);

  const json = JSON.parse(body);
  const text =
    json?.candidates?.[0]?.content?.parts?.map(p => p.text).join("") || "";

  return text;
}

/**
 * Generic OpenAI call.
 * Returns model text response (string).
 *
 * opts:
 * - model (string) default OPENAI_DEFAULT_MODEL
 * - temperature (number)
 * - maxOutputTokens (int)
 * - responseMimeType ("application/json" etc)
 * - responseSchema (object) — JSON schema if responseMimeType is application/json
 */
function openaiGenerateText_(prompt, opts) {
  opts = opts || {};
  const model = getAiModel_("openai", opts.model);
  const responseFormat = (opts.responseMimeType === "application/json")
    ? (opts.responseSchema
        ? { type: "json_schema", json_schema: { name: "response", schema: opts.responseSchema, strict: true } }
        : { type: "json_object" })
    : undefined;

  const payload = {
    model,
    input: [
      { role: "user", content: [{ type: "text", text: String(prompt || "") }] }
    ],
    temperature: (opts.temperature != null) ? opts.temperature : 0.4,
    max_output_tokens: getAiMaxOutputTokens_(opts.maxOutputTokens),
    response_format: responseFormat
  };

  if (!payload.response_format) delete payload.response_format;

  const res = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    contentType: "application/json",
    headers: { "Authorization": "Bearer " + getOpenAiApiKey_() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code >= 300) throw new Error(`OpenAI error ${code}: ${body}`);

  const json = JSON.parse(body);
  if (json && typeof json.output_text === "string") return json.output_text;

  const output = Array.isArray(json?.output) ? json.output : [];
  const content = output.flatMap(o => Array.isArray(o?.content) ? o.content : []);
  const text = content.map(c => c?.text || "").join("");
  return text || "";
}

/**
 * Provider-aware AI call.
 * Returns model text response (string).
 */
function aiGenerateText_(prompt, opts) {
  opts = opts || {};
  const provider = String(opts.provider || getAiProvider_()).toLowerCase();
  if (provider === "openai") return openaiGenerateText_(prompt, opts);
  return geminiGenerateText_(prompt, opts);
}

/**
 * Legacy helper (compatible with your old gemini.gs usage pattern):
 * Returns JSON text (string). Does NOT parse it for you.
 */
function geminiGenerateJson_(prompt, modelOpt) {
  return geminiGenerateText_(prompt, {
    model: modelOpt || GEMINI_DEFAULT_MODEL,
    temperature: 0.4,
    maxOutputTokens: 700,
    responseMimeType: "application/json"
  });
}

/**
 * Provider-aware JSON helper:
 * Returns JSON text (string). Does NOT parse it for you.
 */
function aiGenerateJson_(prompt, opts) {
  opts = opts || {};
  return aiGenerateText_(prompt, {
    model: opts.model,
    temperature: (opts.temperature != null) ? opts.temperature : 0.4,
    maxOutputTokens: (opts.maxOutputTokens != null) ? opts.maxOutputTokens : 700,
    responseMimeType: "application/json",
    responseSchema: opts.responseSchema
  });
}

/* =========================
 * AI Scoring Config Generator
 * ========================= */

/**
 * UI wrapper: called by HTML
 * Returns { ok: true, jsonText: "..." }
 */
function ui_generateScoringConfigFromAI_v1(userPrompt, baselineJsonText) {
  const existing = (baselineJsonText && String(baselineJsonText).trim())
    ? JSON.parse(String(baselineJsonText))
    : sc_getConfigSafe_(); // expected to exist elsewhere

  const taxonomy = getTaxonomyFromThemeRules_();
  const cfg = ai_generateScoringConfig_(userPrompt, existing, taxonomy);
  return { ok: true, jsonText: JSON.stringify(cfg, null, 2) };
}

/**
 * Gemini: Generate / adjust FULL scoring config JSON for scoring.gs
 * - Strict sanitisation: unknown keys dropped, missing keys filled, numbers clamped
 * - IMPORTANT: responseSchema object properties must be NON-EMPTY for OBJECT type
 */
function ai_generateScoringConfig_(userPrompt, existingCfg, taxonomy) {
  const GROUP_KEYS = [
    "policy","programme","leadership","labourMarket","skills",
    "industrialRelations","evidence","institutions",
    "emergingTrends","foresight","aiContext","keywordDensity"
  ];

  const baseline = sanitizeScoringConfigFull_(existingCfg || {}, true);

  // ✅ Build NON-EMPTY schema properties for nested objects
  const weightsProps = {};
  const groupsProps = {};
  GROUP_KEYS.forEach(k => {
    weightsProps[k] = { type: "number" };
    groupsProps[k]  = { type: "array", items: { type: "string" } };
  });

  const schema = {
    type: "object",
    properties: {
      version: { type: "string" },
      scopeGateRegex: { type: "string" },

      // params used by scoring.gs
      topN: { type: "integer" },
      minScore: { type: "number" },
      recencyDays: { type: "integer" },
      capPerTheme: { type: "integer" },
      capPerPOI: { type: "integer" },
      capPerSource: { type: "integer" },
      titleWeight: { type: "number" },
      descWeight: { type: "number" },

      // ✅ MUST be object with NON-EMPTY properties
      weights: {
        type: "object",
        properties: weightsProps
      },
      keywordGroups: {
        type: "object",
        properties: groupsProps
      },

      notes: { type: "string" }
    },
    required: [
      "scopeGateRegex",
      "topN","minScore","recencyDays","capPerTheme","capPerPOI","capPerSource","titleWeight","descWeight",
      "weights","keywordGroups"
    ]
  };

  const system = [
    "You generate a FULL scoring config JSON for a labour-market horizon scanning tool.",
    "Return ONLY a single JSON object. No markdown. No explanation text.",
    "Make small, explainable changes. Avoid extreme values.",
    "",
    "CRITICAL RULES:",
    `1) weights keys LIMITED to: ${GROUP_KEYS.join(", ")}.`,
    `2) keywordGroups keys LIMITED to: ${GROUP_KEYS.join(", ")}.`,
    "3) keywordGroups values MUST be arrays of strings (keywords/phrases).",
    "4) Do NOT add new top-level keys beyond the schema.",
    "5) scopeGateRegex must stay a REGEX STRING that matches manpower/workplace/labour-market content.",
    "",
    "If the user prompt is empty, return the baseline with minor safe improvements."
  ].join("\n");

  const context = [
    "Baseline config (use this as starting point):",
    JSON.stringify(baseline),
    "",
    "Active taxonomy (awareness only; do not invent new themes/pois):",
    "Themes: " + JSON.stringify((taxonomy && taxonomy.themes) || []),
    "POIs: " + JSON.stringify((taxonomy && taxonomy.pois) || []),
    "",
    "User request:",
    String(userPrompt || "").trim()
  ].join("\n");

  const outText = aiGenerateText_(system + "\n\n" + context, {
    temperature: 0.2,
    maxOutputTokens: 800,
    responseMimeType: "application/json",
    responseSchema: schema
  });

  let cfgRaw;
  try {
    cfgRaw = JSON.parse(outText);
  } catch (e) {
    const extracted = extractFirstJsonObject_(outText);
    if (!extracted) throw new Error("Gemini output was not valid JSON.");
    cfgRaw = extracted;
  }

  if (!cfgRaw || typeof cfgRaw !== "object" || Array.isArray(cfgRaw)) {
    throw new Error("Gemini output is not a JSON object.");
  }

  // Merge onto baseline so we always return complete config
  const merged = deepMerge_(baseline, cfgRaw);

  // Final strict sanitise + clamps
  return sanitizeScoringConfigFull_(merged, true);
}

/**
 * Strict sanitiser for FULL scoring config.
 * - Drops unknown top-level keys
 * - Forces weights + keywordGroups keys to GROUP_KEYS only
 * - Ensures required fields exist if fallback=true
 * - Clamps numeric fields to safe ranges
 */
function sanitizeScoringConfigFull_(cfg, fallback) {
  const GROUP_KEYS = [
    "policy","programme","leadership","labourMarket","skills",
    "industrialRelations","evidence","institutions",
    "emergingTrends","foresight","aiContext","keywordDensity"
  ];

  const defaults = {
    version: "1.0",
    scopeGateRegex: "\\b(job|jobs|employment|workforce|wage|salary|skills|training|retren(ch|chment)|layoff|unemployment|union|tripartite|migrant|retirement|wsh|workplace)\\b",

    // params
    topN: 15,
    minScore: 8,
    recencyDays: 365,
    capPerTheme: 3,
    capPerPOI: 3,
    capPerSource: 2,
    titleWeight: 1.0,
    descWeight: 1.0,

    // scoring groups
    weights: {
      policy: 3, programme: 2, leadership: 2, labourMarket: 2, skills: 2,
      industrialRelations: 2, evidence: 2, institutions: 1,
      emergingTrends: 2, foresight: 3, aiContext: 1, keywordDensity: 1
    },

    keywordGroups: {
      policy: ["bill", "act", "regulation", "policy", "guideline", "framework"],
      programme: ["grant", "scheme", "pilot", "initiative", "programme", "program"],
      leadership: ["minister", "agency", "tripartite", "employer federation", "union leaders"],
      labourMarket: ["wage", "salary", "jobs", "hiring", "retrench", "layoff", "unemployment"],
      skills: ["training", "reskill", "upskill", "certification", "apprenticeship"],
      industrialRelations: ["union", "collective", "industrial action", "strike", "dispute"],
      evidence: ["survey", "dataset", "study", "report", "evaluation", "impact"],
      institutions: ["ilo", "oecd", "world bank", "think tank", "ministry"],
      emergingTrends: ["automation", "platform work", "gig", "hybrid work", "four-day week"],
      foresight: ["by 2026", "2026", "2027", "2030", "forecast", "outlook", "scenario", "projection"],
      aiContext: ["ai", "genai", "automation", "machine learning"],
      keywordDensity: []
    },

    notes: ""
  };

  const TOP_KEYS = Object.keys(defaults);
  const out = {};

  TOP_KEYS.forEach(k => {
    if (cfg && (k in cfg)) out[k] = cfg[k];
    else if (fallback) out[k] = defaults[k];
  });

  // Normalise strings
  out.version = String(out.version || defaults.version);
  out.scopeGateRegex = String(out.scopeGateRegex || defaults.scopeGateRegex);
  out.notes = String(out.notes || "");

  // Clamp params
  out.topN = clampInt_(out.topN, 5, 50);
  out.minScore = clampNum_(out.minScore, 0, 999);
  out.recencyDays = clampInt_(out.recencyDays, 1, 3650);
  out.capPerTheme = clampInt_(out.capPerTheme, 1, 50);
  out.capPerPOI = clampInt_(out.capPerPOI, 1, 50);
  out.capPerSource = clampInt_(out.capPerSource, 1, 50);
  out.titleWeight = clampNum_(out.titleWeight, 0, 5);
  out.descWeight = clampNum_(out.descWeight, 0, 5);

  // Weights: keep only GROUP_KEYS
  out.weights = normalizeWeights_(out.weights, defaults.weights, GROUP_KEYS);

  // KeywordGroups: keep only GROUP_KEYS, ensure arrays of strings
  out.keywordGroups = normalizeKeywordGroups_(out.keywordGroups, defaults.keywordGroups, GROUP_KEYS);

  // Validate regex (don’t crash scoring if invalid)
  if (!isValidRegex_(out.scopeGateRegex)) {
    out.scopeGateRegex = defaults.scopeGateRegex;
    out.notes = (out.notes ? out.notes + "\n" : "") +
      "NOTE: Invalid scopeGateRegex returned by AI; reverted to default.";
  }

  return out;
}

/* ---------- Normalisers ---------- */

function normalizeWeights_(w, fallbackWeights, groupKeys) {
  const out = {};
  const src = (w && typeof w === "object" && !Array.isArray(w)) ? w : {};
  groupKeys.forEach(k => {
    const v = Number(src[k]);
    out[k] = Number.isFinite(v) && v >= 0 ? v : Number(fallbackWeights[k] || 0);
  });
  return out;
}

function normalizeKeywordGroups_(g, fallbackGroups, groupKeys) {
  const out = {};
  const src = (g && typeof g === "object" && !Array.isArray(g)) ? g : {};
  groupKeys.forEach(k => {
    const arr = src[k];
    if (Array.isArray(arr)) {
      out[k] = arr.map(x => String(x || "").trim()).filter(Boolean);
    } else {
      out[k] = (fallbackGroups[k] || []).map(x => String(x || "").trim()).filter(Boolean);
    }
  });
  return out;
}

/* ---------- Utilities ---------- */

function clampInt_(n, lo, hi) {
  n = Math.round(Number(n));
  if (!Number.isFinite(n)) n = lo;
  return Math.max(lo, Math.min(hi, n));
}
function clampNum_(n, lo, hi) {
  n = Number(n);
  if (!Number.isFinite(n)) n = lo;
  return Math.max(lo, Math.min(hi, n));
}

function isValidRegex_(pattern) {
  try { new RegExp(String(pattern || ""), "i"); return true; } catch (_) { return false; }
}

/**
 * Extract the first JSON object from a string, if Gemini includes extra chars.
 */
function extractFirstJsonObject_(text) {
  const s = String(text || "");
  const start = s.indexOf("{");
  if (start < 0) return null;

  let depth = 0;
  for (let i = start; i < s.length; i++) {
    const ch = s[i];
    if (ch === "{") depth++;
    if (ch === "}") depth--;
    if (depth === 0) {
      const candidate = s.slice(start, i + 1);
      try { return JSON.parse(candidate); } catch (_) { return null; }
    }
  }
  return null;
}

/**
 * Deep merge objects (patch onto base) without losing nested objects
 */
function deepMerge_(base, patch) {
  const b = (base && typeof base === "object") ? base : {};
  const p = (patch && typeof patch === "object") ? patch : {};
  const out = Array.isArray(b) ? b.slice() : Object.assign({}, b);

  Object.keys(p).forEach(k => {
    const pv = p[k];
    const bv = out[k];
    if (pv && typeof pv === "object" && !Array.isArray(pv) &&
        bv && typeof bv === "object" && !Array.isArray(bv)) {
      out[k] = deepMerge_(bv, pv);
    } else {
      out[k] = pv;
    }
  });

  return out;
}

/**
 * Returns taxonomy from ThemeRules (active rows only)
 * { themes: string[], pois: string[] }
 */
function getTaxonomyFromThemeRules_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONTROL_SHEET); // expected to exist elsewhere (e.g., "ThemeRules")
  if (!sh) throw new Error(`Control sheet not found: ${CONTROL_SHEET}`);

  const lastRow = sh.getLastRow();
  if (lastRow < 5) return { themes: [], pois: [] };

  const rows = sh.getRange(5, 1, lastRow - 4, 4).getValues(); // A:D
  const themes = new Set();
  const pois = new Set();

  rows.forEach(r => {
    const theme = String(r[0] || "").trim();
    const poi   = String(r[1] || "").trim();
    const keys  = String(r[2] || "").trim();
    const active = (r[3] === true);
    if (!active || !theme || !poi || !keys) return;
    themes.add(theme);
    pois.add(poi);
  });

  return {
    themes: Array.from(themes).sort((a,b) => String(a).localeCompare(String(b))),
    pois: Array.from(pois).sort((a,b) => String(a).localeCompare(String(b)))
  };
}
