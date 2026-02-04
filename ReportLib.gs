/******************************
 * ReportLib.gs — Futurescan core helpers
 *
 * Responsibilities:
 * - Parse links
 * - Resolve Google News wrapper → publisher URL
 * - Fetch article meta (via your existing fetchArticleMeta_)
 * - AI summarisation (via your existing aiGenerateJson_)
 * - Optional PDF hook (pdf_makeArticlePdf_) kept but gated by caller
 * - Preview block rendering helpers
 ******************************/

/* =========================================================================
 * Public helper: build topics from user links
 * Returns array of topic objects:
 * { topicNo, title, relevance20, summaryHtml, imageUrl, pdfUrl, articleUrl, sectionTitle }
 * ========================================================================= */

function rpt_buildTopicsFromLinks_(linksText, opts) {
  const options = opts || {};
  const includePdf = !!options.includePdf;

  const links = rpt_parseLinks_(linksText);
  if (links.length < 3) throw new Error("Please provide at least 3 links.");

  const topics = [];

  for (let i = 0; i < links.length; i++) {
    const idx = i + 1;
    const inputUrl = links[i];

    const resolved = resolveUrlIfGoogleNews_(inputUrl);
    const finalUrl = resolved || inputUrl;

    // fetch meta safely
    const a = rpt_fetchArticleMetaSafe_(finalUrl, inputUrl);

    // AI
    const ai = generateAiBits_(a);
    const sectionTitle = generateAiTopicHeader_(a);

    // optional PDF (offline article replica)
    let pdfUrl = "";
    if (includePdf && typeof pdf_makeArticlePdf_ === "function") {
      try {
        pdfUrl = rpt_makeArticlePdfFromUrl_(finalUrl) || "";
      } catch (e) {
        pdfUrl = "";
      }
    }

    topics.push({
      topicNo: idx,
      title: a.title,
      relevance20: ai.relevance20,
      summaryHtml: ai.summaryHtml,
      imageUrl: a.ogImage || "",
      pdfUrl: pdfUrl,
      articleUrl: a.url,
      sectionTitle
    });
  }

  return topics;
}

function rpt_parseLinks_(linksText) {
  const maxN = rpt_getMaxReportLinks_();

  return (linksText || "")
    .split("\n")
    .map(s => String(s || "").trim())
    .filter(Boolean)
    .slice(0, maxN);
}

function rpt_getMaxReportLinks_() {
  try {
    if (typeof MAX_REPORT_LINKS !== "undefined") {
      const n = Number(MAX_REPORT_LINKS);
      if (!isNaN(n) && n > 0) return Math.min(20, Math.floor(n));
    }
  } catch (e) {}
  return 5;
}

/* =========================================================================
 * Article meta fetch wrapper + sanitisation
 * ========================================================================= */

function rpt_fetchArticleMetaSafe_(finalUrl, inputUrl) {
  let a;
  try {
    a = fetchArticleMeta_(finalUrl); // dependency elsewhere
  } catch (e) {
    a = { url: finalUrl, title: "Untitled", description: "", ogImage: "", text: "" };
  }

  // Ensure we have a sane URL
  a.url = (a.url || finalUrl || inputUrl || "").trim();
  if (isBadPublisherCandidate_(a.url)) {
    a.url = finalUrl;
    a.title = a.title || "Untitled";
    a.description = "";
    a.text = "";
    a.ogImage = "";
  }

  // Sanitize fields so we never display binary junk
  a.title = safeTitle_(a.title, a.url);
  a.description = safeTextSnippet_(a.description, 600);
  a.text = safeTextSnippet_(a.text, 12000);
  a.ogImage = (a.ogImage || "").trim();

  return a;
}

/* =========================================================================
 * Optional: PDF helper (offline article replica)
 * ========================================================================= */

function rpt_makeArticlePdfFromUrl_(articleUrl) {
  if (!articleUrl) return "";
  if (typeof pdf_makeArticlePdf_ !== "function") return "";

  const opts = (typeof pdf_defaultEmailOpts_ === "function")
    ? pdf_defaultEmailOpts_()
    : {};

  const res = pdf_makeArticlePdf_(articleUrl, opts);
  if (!res) return "";
  if (res.fileUrl) return res.fileUrl;
  if (!res.pdfBlob) return "";

  const pdfName = res.pdfName || "Article.pdf";
  let file;

  if (typeof PDF_FOLDER_ID !== "undefined" &&
      PDF_FOLDER_ID &&
      PDF_FOLDER_ID !== "PASTE_PDF_FOLDER_ID_HERE") {
    file = DriveApp.getFolderById(PDF_FOLDER_ID).createFile(res.pdfBlob.setName(pdfName));
  } else {
    file = DriveApp.createFile(res.pdfBlob.setName(pdfName));
  }

  return file.getUrl();
}

/* =========================================================================
 * 1) Google News resolver (robust)
 * ========================================================================= */

function resolveUrlIfGoogleNews_(url) {
  if (!url) return "";

  const u = String(url).trim();
  if (!isGoogleNewsWrapper_(u)) return u;

  // Primary: c-wiz[data-p] + batchexecute rpcid=Fbv4je
  const viaBatch = tryResolveGoogleNewsViaBatchexecute_(u);
  if (isGoodPublisherCandidate_(viaBatch)) return viaBatch;

  // Fallback A: decode token
  const token = extractGoogleNewsToken_(u);
  if (token) {
    const decoded = decodeGoogleNewsTokenToUrl_(token);
    if (isGoodPublisherCandidate_(decoded)) return decoded;
  }

  // Fallback B: redirect hops
  const byRedirect = resolveByRedirectHops_(u, 6);
  if (isGoodPublisherCandidate_(byRedirect)) return byRedirect;

  // Fallback C: HTML sniffing
  const scraped = scrapePublisherFromWrapperHtml_(u);
  if (isGoodPublisherCandidate_(scraped)) return scraped;

  return u;
}

function isGoogleNewsWrapper_(u) {
  const s = String(u || "").toLowerCase();
  return (
    s.includes("news.google.com/rss/articles/") ||
    s.includes("news.google.com/articles/") ||
    s.includes("news.google.com/read/") ||
    s.includes("news.google.com/rss/")
  );
}

/**
 * ✅ StackOverflow method (ported to Apps Script)
 */
function tryResolveGoogleNewsViaBatchexecute_(googleNewsUrl) {
  try {
    const resp = UrlFetchApp.fetch(googleNewsUrl, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
      },
      timeout: 20000
    });

    const code = resp.getResponseCode();
    if (code < 200 || code >= 400) return "";

    const html = resp.getContentText() || "";

    // Extract c-wiz[data-p]
    const m = html.match(/<c-wiz[^>]+data-p="([^"]+)"/i);
    if (!m || !m[1]) return "";

    const dataP = decodeHtmlEntities_(m[1]);

    // Convert special payload into valid JSON
    const jsonStr = dataP.replace("%.@.", '["garturlreq",');
    let obj = JSON.parse(jsonStr);

    // obj[:-6] + obj[-2:]
    if (Array.isArray(obj)) obj = obj.slice(0, -6).concat(obj.slice(-2));

    const fReq = JSON.stringify([
      [["Fbv4je", JSON.stringify(obj), "null", "generic"]]
    ]);

    const post = UrlFetchApp.fetch("https://news.google.com/_/DotsSplashUi/data/batchexecute", {
      method: "post",
      muteHttpExceptions: true,
      followRedirects: true,
      contentType: "application/x-www-form-urlencoded;charset=UTF-8",
      headers: { "User-Agent": "Mozilla/5.0" },
      payload: { "f.req": fReq },
      timeout: 20000
    });

    const t = String(post.getContentText() || "");
    const cleaned = t.replace(/^\)\]\}'\s*/, "");
    const outer = JSON.parse(cleaned);

    const innerStr = outer && outer[0] && outer[0][2] ? outer[0][2] : "";
    if (!innerStr) return "";

    const inner = JSON.parse(innerStr);
    const articleUrl = inner && inner[1] ? String(inner[1]) : "";
    if (isGoodPublisherCandidate_(articleUrl)) return articleUrl;

    return "";
  } catch (e) {
    return "";
  }
}

function extractGoogleNewsToken_(u) {
  let m = String(u).match(/\/rss\/articles\/([^?]+)/i);
  if (m && m[1]) return m[1];

  m = String(u).match(/\/articles\/([^?]+)/i);
  if (m && m[1]) return m[1];

  return "";
}

function decodeGoogleNewsTokenToUrl_(token) {
  try {
    const bytes = Utilities.base64DecodeWebSafe(token);
    const bin = Utilities.newBlob(bytes).getDataAsString("ISO-8859-1");

    const urls = extractAllHttpUrlsFromBinary_(bin);
    const best = pickBestPublisherUrl_(urls);
    return best || "";
  } catch (e) {
    return "";
  }
}

function extractAllHttpUrlsFromBinary_(bin) {
  const out = [];
  if (!bin) return out;

  const re = /https?:\/\/[^\s"'<>\\\u0000\u0001\u001f]+/g;
  let m;

  while ((m = re.exec(bin)) !== null) {
    let u = m[0];
    u = u.replace(/[^\x20-\x7E]+$/g, "").trim();
    u = u.split("\u0000")[0].split("\u001f")[0].split("\u0001")[0].trim();
    if (u.startsWith("http")) out.push(u);
  }

  return Array.from(new Set(out));
}

function pickBestPublisherUrl_(urls) {
  if (!urls || !urls.length) return "";

  const good = urls.filter(isGoodPublisherCandidate_);
  if (!good.length) return "";

  good.sort((a, b) => b.length - a.length);
  return good[0];
}

function resolveByRedirectHops_(startUrl, maxHops) {
  let cur = startUrl;

  for (let i = 0; i < maxHops; i++) {
    const res = UrlFetchApp.fetch(cur, {
      followRedirects: false,
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" },
      timeout: 15000
    });

    const code = res.getResponseCode();
    const headers = res.getAllHeaders ? res.getAllHeaders() : {};
    const loc = headers.Location || headers.location || "";

    if ((code === 301 || code === 302 || code === 303 || code === 307 || code === 308) && loc) {
      const next = absolutizeUrl_(cur, loc);
      if (isBadPublisherCandidate_(next)) return startUrl;
      cur = next;
      continue;
    }

    return cur;
  }

  return cur;
}

function scrapePublisherFromWrapperHtml_(u) {
  try {
    const res = UrlFetchApp.fetch(u, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
      },
      timeout: 20000
    });

    const body = res.getContentText() || "";

    let m = body.match(/data-n-au="([^"]+)"/i);
    if (m && m[1]) {
      const candidate = decodeHtmlEntities_(m[1]);
      if (isGoodPublisherCandidate_(candidate)) return candidate;
    }

    m = body.match(/https?:\/\/www\.google\.com\/url\?[^"']*url=([^&"']+)/i);
    if (m && m[1]) {
      const candidate = decodeURIComponent(m[1]);
      if (isGoodPublisherCandidate_(candidate)) return candidate;
    }

    const all = body.match(/https?:\/\/[^"' <>\]]+/g) || [];
    const publisher = all.find(isGoodPublisherCandidate_);
    return publisher || "";
  } catch (e) {
    return "";
  }
}

function absolutizeUrl_(base, maybeRelative) {
  if (!maybeRelative) return base;
  if (/^https?:\/\//i.test(maybeRelative)) return maybeRelative;

  const m = String(base).match(/^(https?:\/\/[^\/]+)/i);
  const origin = m ? m[1] : "";
  if (maybeRelative.startsWith("/")) return origin + maybeRelative;
  return origin + "/" + maybeRelative;
}

function decodeHtmlEntities_(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

/* ----------------------------
 * Candidate filtering
 * ---------------------------- */

function isGoodPublisherCandidate_(u) {
  return u && String(u).startsWith("http") && !isBadPublisherCandidate_(u);
}

function isBadPublisherCandidate_(u) {
  const s = String(u || "").toLowerCase();
  if (!s) return true;

  if (s.includes("news.google.com")) return true;

  if (s.includes("lh3.googleusercontent.com") || s.includes("lh4.googleusercontent.com") || s.includes("lh5.googleusercontent.com")) return true;
  if (s.includes("gstatic.com")) return true;

  if (s.includes("google-analytics.com")) return true;
  if (s.includes("googletagmanager.com")) return true;
  if (s.includes("doubleclick.net")) return true;
  if (s.includes("googlesyndication.com")) return true;
  if (s.includes("adservice.google.com")) return true;

  if (s.match(/\.(png|jpg|jpeg|webp|gif|bmp|svg)(\?|#|$)/)) return true;
  if (s.match(/\.(js|css|json|xml|txt|map)(\?|#|$)/)) return true;

  if (s.includes("=w16") || s.includes("=w32") || s.includes("=w64")) return true;

  return false;
}

/* =========================================================================
 * 2) AI generation (robust JSON + safe rendering)
 * ========================================================================= */

function generateAiBits_(article) {
  const rawSource = (article.description || article.text || "");
  const sourceText = cleanForModel_(rawSource).slice(0, 9000);

  if (!sourceText || sourceText.length < 80) {
    return fallbackAi_(article);
  }

  const prompt = `
You are writing a MOM Singapore FutureScans weekly brief.

Given the article:
TITLE: ${article.title}
URL: ${article.url}
TEXT (may be partial): ${sourceText}

Return STRICT JSON only with:
{
  "relevance20": "Exactly 20 words, how this is relevant to Singapore/MOM (labour, manpower, employment, workplace).",
  "summary80": "About 80 words total. Include 1-2 key points that are important and measurable.",
  "boldPhrases": ["1-2 short phrases to bold inside summary80 (must appear verbatim in summary80)"]
}

Rules:
- Output JSON only (no markdown fences, no commentary).
- Use double quotes for JSON keys/strings.
- Do not invent statistics.
- Keep it policy/workforce focused.
`.trim();

  let raw = "";
  try {
    raw = aiGenerateJson_(prompt); // dependency elsewhere
  } catch (e) {
    return fallbackAi_(article);
  }

  let obj;
  try {
    obj = safeJsonParse_(raw);
  } catch (e) {
    console.log("AI JSON parse failed:", e && e.message ? e.message : e);
    console.log("RAW AI OUTPUT (first 800 chars):", String(raw).slice(0, 800));
    return fallbackAi_(article);
  }

  let relevance20 = String(obj.relevance20 || "").trim();
  let summary80 = String(obj.summary80 || "").trim();
  const phrases = Array.isArray(obj.boldPhrases) ? obj.boldPhrases.slice(0, 2) : [];

  if (!relevance20) relevance20 = fallbackRelevance20_();
  if (!summary80) summary80 = safeTextSnippet_(article.description || article.text || "", 420);

  let summaryHtml = escapeHtml_(summary80);

  phrases.forEach(p => {
    const phrase = String(p || "").trim();
    if (!phrase) return;

    const safePhrase = escapeHtml_(phrase);
    if (summaryHtml.includes(safePhrase)) {
      summaryHtml = summaryHtml.replace(safePhrase, `<b>${safePhrase}</b>`);
    }
  });

  return { relevance20, summaryHtml };
}

/* =========================================================================
 * 2b) Topic header classification
 * ========================================================================= */

function generateAiTopicHeader_(article) {
  const options = getReportTopicOptions_();
  const fallback = "Other Labour Markets";
  const sourceText = cleanForModel_(article.description || article.text || "").slice(0, 4000);
  const title = String(article.title || "").trim();

  if (!title && sourceText.length < 80) {
    return fallback;
  }

  const prompt = `
You are classifying a news article for a MOM Singapore FutureScans brief.

Choose the single best topic label from this list:
${options.map(o => `- ${o}`).join("\n")}

Article:
TITLE: ${title}
URL: ${article.url}
TEXT (may be partial): ${sourceText}

Return STRICT JSON only:
{ "topic": "one label from the list above" }

Rules:
- Output JSON only (no markdown fences).
- Use the label exactly as written in the list.
`.trim();

  try {
    const raw = aiGenerateJson_(prompt);
    const obj = safeJsonParse_(raw);
    const picked = String(obj.topic || "").trim();
    if (isReportTopicOption_(picked, options)) {
      return picked;
    }
  } catch (e) {
    // fall through to heuristic
  }

  const heuristic = classifyTopicHeuristic_(article, options);
  return heuristic || fallback;
}

function getReportTopicOptions_() {
  return [
    "Jobs/ Employment and skills",
    "Minimum wage/ Low-wage workers",
    "Self-employed persons/ Platform workers",
    "Senior workers",
    "Retirement adequacy",
    "Migrant workers/ Foreign workers policy",
    "Women at work",
    "Union/ Labour movement",
    "Workplace safety and health/ Workplace injury",
    "Behaviour Insights",
    "Organisational culture/ Organisational health/ Work culture",
    "Hybrid working/ Remote working/ Flexible work arrangements",
    "Artificial intelligence/ Generative AI",
    "Tech/ Automation"
  ];
}

function isReportTopicOption_(value, options) {
  if (!value) return false;
  const v = String(value).trim().toLowerCase();
  return (options || []).some(opt => String(opt).trim().toLowerCase() === v);
}

function classifyTopicHeuristic_(article, options) {
  const text = `${article.title || ""} ${article.description || ""} ${article.text || ""}`.toLowerCase();
  const pick = label => (isReportTopicOption_(label, options) ? label : "");

  if (text.match(/\b(ai|artificial intelligence|generative ai|llm|chatgpt)\b/)) {
    return pick("Artificial intelligence/ Generative AI");
  }
  if (text.match(/\b(automation|robotics|automated|autonomous|digitisation|digitalization|technology|tech)\b/)) {
    return pick("Tech/ Automation");
  }
  if (text.match(/\b(remote|hybrid|flexible work|work from home|telework)\b/)) {
    return pick("Hybrid working/ Remote working/ Flexible work arrangements");
  }
  if (text.match(/\b(behaviour|behavioral|behavioural|nudges?|insight)\b/)) {
    return pick("Behaviour Insights");
  }
  if (text.match(/\b(culture|organisational|organizational|workplace culture)\b/)) {
    return pick("Organisational culture/ Organisational health/ Work culture");
  }
  if (text.match(/\b(safety|injury|accident|osha|wsq|workplace safety)\b/)) {
    return pick("Workplace safety and health/ Workplace injury");
  }
  if (text.match(/\b(union|labour movement|labor movement|collective bargaining)\b/)) {
    return pick("Union/ Labour movement");
  }
  if (text.match(/\b(women|female|gender)\b/)) {
    return pick("Women at work");
  }
  if (text.match(/\b(migrant|foreign worker|work permit|s pass|employment pass)\b/)) {
    return pick("Migrant workers/ Foreign workers policy");
  }
  if (text.match(/\b(retirement|pension|cpf|elderly|aged|older workers)\b/)) {
    return pick("Retirement adequacy");
  }
  if (text.match(/\b(senior worker|older worker|ageing workforce|aging workforce)\b/)) {
    return pick("Senior workers");
  }
  if (text.match(/\b(self-employed|self employed|platform worker|gig worker|freelancer)\b/)) {
    return pick("Self-employed persons/ Platform workers");
  }
  if (text.match(/\b(minimum wage|low-wage|low wage|workfare)\b/)) {
    return pick("Minimum wage/ Low-wage workers");
  }
  if (text.match(/\b(job|employment|skills|training|reskilling|upskilling|labour market|labor market)\b/)) {
    return pick("Jobs/ Employment and skills");
  }

  return "";
}

function safeJsonParse_(raw) {
  if (!raw) throw new Error("Empty model response.");

  let t = String(raw).trim().replace(/```json|```/g, "").trim();

  try { return JSON.parse(t); } catch (e) {}

  const start = t.indexOf("{");
  const end = t.lastIndexOf("}");
  if (start >= 0 && end > start) {
    const candidate = t.slice(start, end + 1).trim();
    try { return JSON.parse(candidate); } catch (e) {}
  }

  let repaired = t
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/,\s*}/g, "}")
    .replace(/,\s*]/g, "]");

  const s2 = repaired.indexOf("{");
  const e2 = repaired.lastIndexOf("}");
  if (s2 >= 0 && e2 > s2) repaired = repaired.slice(s2, e2 + 1).trim();

  return JSON.parse(repaired);
}

function fallbackAi_(article) {
  const relevance20 = fallbackRelevance20_();
  const base = article.description
    ? safeTextSnippet_(article.description, 420)
    : safeTextSnippet_(article.text, 420);

  const summary = base || "Could not retrieve readable article text. If this is Google News, paste the publisher link if available.";
  return {
    relevance20,
    summaryHtml: escapeHtml_(summary)
  };
}

function fallbackRelevance20_() {
  return "Relevant to Singapore workforce planning and labour market monitoring, offering signals on hiring, wages, job quality, and policy responses.";
}

/* =========================================================================
 * 3) Images (best-effort)
 * ========================================================================= */

function tryFetchImageBlob_(imgUrl) {
  if (!imgUrl) return null;

  try {
    const res = UrlFetchApp.fetch(imgUrl, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" },
      timeout: 15000
    });

    const code = res.getResponseCode();
    if (code < 200 || code >= 300) return null;

    const blob = res.getBlob();
    const bytes = blob.getBytes();
    if (bytes && bytes.length > 1200 * 1024) return null;

    return blob;
  } catch (e) {
    return null;
  }
}

/* =========================================================================
 * 4) Rendering (UI preview blocks)
 * ========================================================================= */

function renderTopicBlock_(topicNo, title, relevance20, summaryHtml, imgUrl, pdfUrl, articleUrl) {
  const barColor = (topicNo === 2) ? "#F57C00" : "#0D47A1";
  const linkStyle = "color:#1a73e8; text-decoration:underline;";

  const imgHtml = imgUrl
    ? `<img src="${escapeHtml_(imgUrl)}" style="width:220px; max-height:130px; object-fit:cover; border-radius:8px; border:1px solid #eee;" />`
    : `<div style="width:220px; height:130px; border-radius:8px; border:1px dashed #ccc; display:flex; align-items:center; justify-content:center; color:#777;">No image</div>`;

  const pdfLine = pdfUrl
    ? `<a style="${linkStyle}" href="${escapeHtml_(pdfUrl)}" target="_blank" rel="noopener">Article ${topicNo} – Offline PDF</a>`
    : `<span style="color:#a7afc0; font-size:12px;">PDF generation on hold (fast mode)</span>`;

  return `
  <div style="margin-bottom:18px; border:1px solid #eee;">
    <div style="background:${barColor}; color:#fff; padding:6px 10px; font-weight:bold;">
      &lt;Topic ${topicNo}&gt;
    </div>

    <div style="padding:10px 12px;">
      <div style="font-weight:bold; margin-bottom:6px;">[${escapeHtml_(title)}]</div>

      <div style="margin:8px 0; color:#333;">
        ${escapeHtml_(relevance20)}
      </div>

      <div style="display:flex; gap:18px; align-items:flex-start; margin-top:10px;">
        <div style="flex:1; min-width:260px;">
          <div>${summaryHtml}</div>
          <div style="margin-top:12px;">
            Click <a style="${linkStyle}" href="${escapeHtml_(articleUrl)}" target="_blank" rel="noopener">here</a> to read more.
          </div>
        </div>

        <div style="width:240px;">
          ${imgHtml}
          <div style="margin-top:8px; font-size:12px; color:#333;">
            ${pdfLine}
          </div>
        </div>
      </div>
    </div>
  </div>`;
}

/* =========================================================================
 * 5) Text safety helpers
 * ========================================================================= */

function safeTitle_(title, url) {
  const t = String(title || "").trim();
  if (t && t.toLowerCase() !== "untitled") return t;

  try {
    const host = String(url || "").replace(/^https?:\/\//, "").split("/")[0];
    return host ? `Untitled (${host})` : "Untitled";
  } catch (e) {
    return "Untitled";
  }
}

function safeTextSnippet_(s, maxLen) {
  const cleaned = cleanForModel_(s);
  if (!cleaned) return "";
  return cleaned.length > maxLen ? cleaned.slice(0, maxLen) + "…" : cleaned;
}

function cleanForModel_(s) {
  if (!s) return "";
  let t = String(s);
  t = t.replace(/[^\x09\x0A\x0D\x20-\x7E]/g, " ");
  t = t.replace(/\s+/g, " ").trim();

  if (t.length < 20) return "";
  return t;
}

function escapeHtml_(s) {
  return (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
