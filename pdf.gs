/*********************************
 * pdf.gs — Article → Offline PDF (Replica-first)
 *
 * What’s improved vs your current version:
 * 1) Replica mode defaults to "article-only replica" (keeps site CSS/head + article body)
 * 2) Better CSS inlining:
 *    - Handles rel=stylesheet AND rel=preload as=style
 *    - Resolves a few @import rules (capped)
 *    - Rewrites CSS url(...) to absolute
 *    - Optional: inline small CSS background images (capped)
 * 3) Better image inlining:
 *    - Parses srcset and tries largest → smaller until under maxImageBytes
 *    - Handles common lazy attributes (data-src, data-original, data-lazy-src, data-srcset)
 * 4) Conversion fallback ladder:
 *    - replica(article) → replica(full-lite) → reader
 *
 * IMPORTANT LIMITATIONS:
 * - Paywalls / cookie-gated content cannot be bypassed.
 * - Some sites block UrlFetchApp (403/451). We try browser-ish headers.
 * - HtmlService PDF conversion is not a full browser; some modern CSS may not render 1:1.
 *********************************/

// ✅ Put a Drive folder to store PDFs (recommended)
const PDF_FOLDER_ID = "PASTE_PDF_FOLDER_ID_HERE"; // optional

const PDF_USER_AGENT =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36";

/**
 * Public UI endpoint (use from your web app / sidebar)
 */
function ui_makeArticlePdfFromUrl_v2(url) {
  try {
    const res = pdf_makeArticlePdf_(url, pdf_defaultEmailOpts_());
    return {
      ok: true,
      message: "PDF created",
      title: res.title,
      fileId: res.fileId || null,
      fileUrl: res.fileUrl || null,
      pdfName: res.pdfName
    };
  } catch (err) {
    return { ok: false, message: String(err && err.message ? err.message : err) };
  }
}

/**
 * Recommended defaults when generating PDFs for email attachments.
 * (Good likeness without exploding PDF size.)
 */
function pdf_defaultEmailOpts_() {
  return {
    mode: "replica",              // "replica" | "reader"
    replicaScope: "article",      // "article" | "full"
    inlineImages: true,
    inlineCss: true,
    inlineCssBgImages: true,      // inline small background images in CSS (helps likeness)
    hideCommonJunk: true,         // CSS hide for ads/nav/subscribe overlays (best-effort)

    // Caps / safety
    maxImages: 10,
    maxImageBytes: 650000,        // ~650KB per image
    maxTotalInlineImageBytes: 3500000, // total bytes across ALL inlined images (avoid giant HTML)
    maxCssFiles: 4,
    maxCssBytes: 260000,
    maxCssImports: 3,
    maxCssBgImages: 10,
    maxCssBgImageBytes: 120000,   // background images are usually small icons/textures
    maxHtmlChars: 1800000         // if HTML grows beyond this, we fallback down
  };
}

/**
 * Core: Make an offline PDF for an article URL.
 * Returns: { pdfBlob, pdfName, title, fileId?, fileUrl? }
 */
function pdf_makeArticlePdf_(url, opts) {
  opts = Object.assign(pdf_defaultEmailOpts_(), opts || {});

  const normalizedUrl = pdf_normalizeUrl_(url);
  const fetched = pdf_fetchHtml_(normalizedUrl);
  const baseUrl = fetched.finalUrl || normalizedUrl;

  let meta = pdf_extractMeta_(fetched.html, baseUrl);
  const title = meta.title || "Article";
  const pdfName = pdf_safeFilename_(title) + ".pdf";

  // 1) Try: Replica (article scope) if requested
  if (opts.mode === "replica") {
    try {
      const replicaHtml = pdf_renderReplicaHtml_(fetched.html, baseUrl, meta, opts);
      const pdfBlob1 = pdf_htmlToPdfBlob_(replicaHtml, pdfName);
      return pdf_maybeSave_(pdfBlob1, pdfName, title, meta.siteName || "");
    } catch (e1) {
      // continue fallback ladder
    }

    // 2) Fallback: Replica full but "lite" (less inlining)
    try {
      const lite = Object.assign({}, opts, {
        replicaScope: "full",
        inlineCssBgImages: false,
        maxImages: Math.min(opts.maxImages || 10, 6),
        maxCssFiles: Math.min(opts.maxCssFiles || 4, 2),
        maxTotalInlineImageBytes: Math.min(opts.maxTotalInlineImageBytes || 3500000, 2000000)
      });
      const replicaLiteHtml = pdf_renderReplicaHtml_(fetched.html, baseUrl, meta, lite);
      const pdfBlob2 = pdf_htmlToPdfBlob_(replicaLiteHtml, pdfName);
      return pdf_maybeSave_(pdfBlob2, pdfName, title, meta.siteName || "");
    } catch (e2) {
      // proceed to reader fallback
    }
  }

  // 3) Reader mode fallback (most reliable)
  const readerHtml = pdf_renderReaderHtml_(fetched.html, baseUrl, meta, Object.assign({}, opts, { mode: "reader" }));
  const pdfBlob3 = pdf_htmlToPdfBlob_(readerHtml, pdfName);
  return pdf_maybeSave_(pdfBlob3, pdfName, title, meta.siteName || "");
}

function pdf_maybeSave_(pdfBlob, pdfName, title, siteName) {
  let fileId = null, fileUrl = null;
  if (PDF_FOLDER_ID && PDF_FOLDER_ID !== "PASTE_PDF_FOLDER_ID_HERE") {
    const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
    const file = folder.createFile(pdfBlob.setName(pdfName));
    fileId = file.getId();
    fileUrl = file.getUrl();
  }
  return { pdfBlob, pdfName, title, siteName, fileId, fileUrl };
}

/* =========================
 * Fetch + URL Normalisation
 * ========================= */

function pdf_normalizeUrl_(url) {
  if (!url) throw new Error("Missing URL");
  let u = String(url).trim();

  // Menlo Safe wrapping:
  // https://safe.menlosecurity.com/https://example.com/path
  u = u.replace(/^safehttps:\/\//i, "https://");
  const menloIdx = u.indexOf("safe.menlosecurity.com/");
  if (menloIdx >= 0) {
    const httpsIdx = u.indexOf("https://", menloIdx);
    if (httpsIdx >= 0) u = u.slice(httpsIdx);
  }

  if (!/^https?:\/\//i.test(u)) u = "https://" + u;
  return u;
}

function pdf_fetchHtml_(url) {
  const resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": PDF_USER_AGENT,
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
      "Accept-Language": "en-GB,en;q=0.9",
      "Cache-Control": "no-cache"
    }
  });

  const code = resp.getResponseCode();
  const finalUrl = (resp.getFinalUrl && resp.getFinalUrl()) ? resp.getFinalUrl() : url;
  const html = resp.getContentText();

  if (code >= 400) {
    if (!html || html.length < 200) throw new Error(`Failed to fetch article (HTTP ${code}). URL: ${finalUrl}`);
  }

  return { html, code, finalUrl };
}

/* =========================
 * Meta + Extraction
 * ========================= */

function pdf_extractMeta_(html, baseUrl) {
  const meta = {
    title: pdf_findMeta_(html, /property=["']og:title["']/i) || pdf_findTag_(html, /<title[^>]*>([\s\S]*?)<\/title>/i),
    siteName: pdf_findMeta_(html, /property=["']og:site_name["']/i),
    author: pdf_findMeta_(html, /name=["']author["']/i) || pdf_findMeta_(html, /property=["']article:author["']/i),
    publishedTime:
      pdf_findMeta_(html, /property=["']article:published_time["']/i) ||
      pdf_findMeta_(html, /name=["']pubdate["']/i) ||
      pdf_findMeta_(html, /name=["']date["']/i),
    description: pdf_findMeta_(html, /name=["']description["']/i) || pdf_findMeta_(html, /property=["']og:description["']/i),
    canonical: pdf_findLinkHref_(html, /rel=["']canonical["']/i) || baseUrl,
    ogImage: pdf_findMeta_(html, /property=["']og:image["']/i)
  };

  Object.keys(meta).forEach(k => {
    if (typeof meta[k] === "string") meta[k] = pdf_decodeHtmlEntities_(meta[k]).trim();
  });

  if (meta.ogImage) {
    try { meta.ogImage = new URL(meta.ogImage, baseUrl).href; } catch (e) {}
  }
  if (meta.canonical) {
    try { meta.canonical = new URL(meta.canonical, baseUrl).href; } catch (e) {}
  }

  return meta;
}

function pdf_findMeta_(html, attrRegex) {
  const re = new RegExp(`<meta[^>]*${attrRegex.source}[^>]*content=["']([^"']+)["'][^>]*>`, "i");
  const m = html.match(re);
  return m ? m[1] : "";
}

function pdf_findLinkHref_(html, attrRegex) {
  const re = new RegExp(`<link[^>]*${attrRegex.source}[^>]*href=["']([^"']+)["'][^>]*>`, "i");
  const m = html.match(re);
  return m ? m[1] : "";
}

function pdf_findTag_(html, re) {
  const m = html.match(re);
  if (!m) return "";
  return pdf_stripTags_(m[1] || "").trim();
}

/* =========================
 * Rendering (Replica / Reader)
 * ========================= */

function pdf_renderReplicaHtml_(rawHtml, baseUrl, meta, opts) {
  let html = String(rawHtml);

  // Remove scripts/iframes/noscripts (PDF conversion can choke on these)
  html = pdf_stripHeavyTags_(html);

  const htmlOpen = pdf_captureTagOpen_(html, "html") || "<html>";
  const bodyOpen = pdf_captureTagOpen_(html, "body") || "<body>";

  // Extract head inner HTML (keep site styles as much as possible)
  const headInner = pdf_extractHeadInner_(html);

  // Extract article body if replicaScope=article, else use body inner
  let bodyInner;
  if ((opts.replicaScope || "article") === "article") {
    bodyInner = pdf_extractArticleHtml_(html) || pdf_extractBodyInner_(html) || "";
  } else {
    bodyInner = pdf_extractBodyInner_(html) || "";
  }

  // Inline CSS links into a combined <style> (best-effort)
  let cssCombined = "";
  if (opts.inlineCss) {
    cssCombined = pdf_fetchAndInlineCss_(headInner, baseUrl, opts);
  }

  // Inject print constraints + optional junk-hiding
  const injectedCss = pdf_buildInjectedCss_(opts);

  // Build a controlled document:
  // - keep original html/body attributes (classes/data-theme can matter)
  // - keep head inner
  // - add base + injected CSS + inlined CSS
  let out = `
<!doctype html>
${htmlOpen}
<head>
  <meta charset="utf-8">
  <base href="${pdf_escapeAttr_(baseUrl)}">
  ${headInner || ""}
  ${injectedCss}
  ${cssCombined ? `<style>\n${cssCombined}\n</style>` : ""}
</head>
${bodyOpen}
  ${pdf_replicaHeaderBanner_(meta, baseUrl)}
  <div id="FS_REPLICA_ROOT">
    ${bodyInner || ""}
  </div>
</body>
</html>
`.trim();

  // Inline images inside the reconstructed document
  if (opts.inlineImages) out = pdf_inlineImages_(out, baseUrl, opts);

  // Optional: inline background images in *inline style=""* attributes too
  if (opts.inlineCssBgImages) out = pdf_inlineStyleBackgroundImages_(out, baseUrl, opts);

  // Guardrail: if HTML becomes too huge, throw so we drop to the fallback ladder
  if (opts.maxHtmlChars && out.length > opts.maxHtmlChars) {
    throw new Error(`Replica HTML too large (${out.length} chars). Falling back.`);
  }

  return out;
}

function pdf_buildInjectedCss_(opts) {
  const hideJunk = opts.hideCommonJunk
    ? `
/* best-effort hide of common junk */
header, nav, footer, aside { display: none !important; }
[class*="subscribe"], [id*="subscribe"],
[class*="paywall"], [id*="paywall"],
[class*="newsletter"], [id*="newsletter"],
[class*="modal"], [id*="modal"],
[class*="popup"], [id*="popup"],
[class*="overlay"], [id*="overlay"],
[class*="cookie"], [id*="cookie"],
[class*="ad-"], [id*="ad-"],
[class*="ads"], [id*="ads"],
[class*="promo"], [id*="promo"],
[class*="banner"], [id*="banner"]
{ display: none !important; }
`
    : "";

  return `
<style>
  @page { margin: 18mm 14mm; }
  html, body { width: 100%; }
  body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  img, video { max-width: 100% !important; height: auto !important; }
  a { color: inherit; text-decoration: none; }
  /* avoid fixed elements covering content */
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  [style*="position:fixed"], [style*="position: fixed"] { position: static !important; }
  ${hideJunk}
</style>
`.trim();
}

function pdf_replicaHeaderBanner_(meta, baseUrl) {
  // Small banner helps when article-only extraction loses the visible headline block.
  // Kept minimal so it doesn't “ruin” replica styling.
  const title = meta && meta.title ? meta.title : "";
  const site = meta && meta.siteName ? meta.siteName : "";
  const href = (meta && meta.canonical) ? meta.canonical : baseUrl;

  if (!title) return "";

  return `
<div style="font-family:Arial,sans-serif;margin:0 0 12px 0;padding:10px 12px;border-bottom:1px solid #ddd;">
  <div style="font-size:18px;font-weight:700;margin-bottom:4px;">${pdf_escapeHtml_(title)}</div>
  <div style="font-size:12px;color:#666;">
    ${site ? pdf_escapeHtml_(site) + " · " : ""}
    <a href="${pdf_escapeAttr_(href)}">${pdf_escapeHtml_(href)}</a>
  </div>
</div>
`.trim();
}

function pdf_renderReaderHtml_(rawHtml, baseUrl, meta, opts) {
  const articleHtml = pdf_extractArticleHtml_(rawHtml);
  let content = articleHtml || "<p>(Unable to extract article body. Showing link only.)</p>";

  if (opts.inlineImages) content = pdf_inlineImages_(content, baseUrl, opts);

  const title = meta.title || "Article";
  const bylineBits = [];
  if (meta.siteName) bylineBits.push(meta.siteName);
  if (meta.author) bylineBits.push(meta.author);
  if (meta.publishedTime) bylineBits.push(meta.publishedTime);
  const byline = bylineBits.join(" · ");

  return `
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <base href="${pdf_escapeAttr_(baseUrl)}">
  <style>
    @page { margin: 18mm 16mm; }
    body { font-family: Georgia, "Times New Roman", serif; line-height: 1.55; color: #111; }
    .wrap { max-width: 900px; margin: 0 auto; }
    h1 { font-family: Arial, sans-serif; font-size: 28px; line-height: 1.2; margin: 0 0 8px; }
    .meta { font-family: Arial, sans-serif; font-size: 12px; color: #666; margin-bottom: 14px; }
    .meta a { color: #666; text-decoration: none; }
    .deck { font-size: 15px; color: #333; margin: 0 0 16px; }
    img { max-width: 100%; height: auto; }
    figure { margin: 14px 0; }
    figcaption { font-size: 12px; color: #666; }
    p { margin: 0 0 12px; }
    h2, h3 { font-family: Arial, sans-serif; margin: 20px 0 10px; }
    blockquote { margin: 14px 0; padding-left: 12px; border-left: 3px solid #ddd; color: #444; }
    hr { border: none; border-top: 1px solid #eee; margin: 18px 0; }
    a { color: #0b57d0; }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>${pdf_escapeHtml_(title)}</h1>
    <div class="meta">
      ${byline ? pdf_escapeHtml_(byline) + "<br>" : ""}
      <a href="${pdf_escapeAttr_(meta.canonical || baseUrl)}">${pdf_escapeHtml_(meta.canonical || baseUrl)}</a>
    </div>
    ${meta.description ? `<p class="deck">${pdf_escapeHtml_(meta.description)}</p>` : ""}
    ${meta.ogImage ? `<figure><img src="${pdf_escapeAttr_(meta.ogImage)}"></figure>` : ""}
    <hr>
    <div class="article">${content}</div>
  </div>
</body>
</html>
`.trim();
}

/* =========================
 * Extract / Sanitize helpers
 * ========================= */

function pdf_stripHeavyTags_(html) {
  let h = String(html);
  h = h.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
  h = h.replace(/<noscript\b[^<]*(?:(?!<\/noscript>)<[^<]*)*<\/noscript>/gi, "");
  h = h.replace(/<iframe\b[^>]*>[\s\S]*?<\/iframe>/gi, "");
  // Remove videos (often heavy / broken in PDF)
  h = h.replace(/<video\b[^>]*>[\s\S]*?<\/video>/gi, "");
  return h;
}

function pdf_captureTagOpen_(html, tagName) {
  const re = new RegExp(`<${tagName}\\b[^>]*>`, "i");
  const m = String(html).match(re);
  return m ? m[0] : "";
}

function pdf_extractHeadInner_(html) {
  const m = String(html).match(/<head\b[^>]*>([\s\S]*?)<\/head>/i);
  return m ? m[1] : "";
}

function pdf_extractBodyInner_(html) {
  const m = String(html).match(/<body\b[^>]*>([\s\S]*?)<\/body>/i);
  return m ? m[1] : "";
}

function pdf_extractArticleHtml_(html) {
  let h = String(html);

  h = h.replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");
  h = h.replace(/<!--([\s\S]*?)-->/g, "");

  const art = h.match(/<article\b[^>]*>([\s\S]*?)<\/article>/i);
  if (art && art[1] && pdf_wordCount_(art[1]) > 180) return pdf_sanitizeContentHtml_(art[1]);

  const main = h.match(/<main\b[^>]*>([\s\S]*?)<\/main>/i);
  if (main && main[1] && pdf_wordCount_(main[1]) > 180) return pdf_sanitizeContentHtml_(main[1]);

  // content-ish div/section
  const candidates = [];
  const re = /<(div|section)\b[^>]*(class|id)=["'][^"']*(content|article|post|story|main|body|entry|page)[^"']*["'][^>]*>([\s\S]*?)<\/\1>/gi;
  let m;
  while ((m = re.exec(h)) !== null) {
    const chunk = m[4] || "";
    const wc = pdf_wordCount_(chunk);
    if (wc > 180) candidates.push({ wc, chunk });
  }
  if (candidates.length) {
    candidates.sort((a, b) => b.wc - a.wc);
    return pdf_sanitizeContentHtml_(candidates[0].chunk);
  }

  return "";
}

function pdf_sanitizeContentHtml_(contentHtml) {
  let c = String(contentHtml);

  c = c.replace(/<(nav|footer|aside|form|button)\b[^>]*>[\s\S]*?<\/\1>/gi, "");
  c = c.replace(/<[^>]*(class|id)=["'][^"']*(ad-|ads|advert|promo|subscribe|newsletter|paywall|cookie|modal|overlay)[^"']*["'][^>]*>[\s\S]*?<\/[^>]+>/gi, "");
  // remove inline scripts that survived
  c = c.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");

  return c.trim();
}

/* =========================
 * CSS inlining (stylesheets + imports + url rewriting)
 * ========================= */

function pdf_fetchAndInlineCss_(headInner, baseUrl, opts) {
  let h = String(headInner || "");
  const cssLinks = [];

  // Match rel="stylesheet"
  const linkRe1 = /<link\b[^>]*rel=["']stylesheet["'][^>]*>/gi;
  // Match rel="preload" as="style" href="..." (common pattern)
  const linkRe2 = /<link\b[^>]*rel=["']preload["'][^>]*as=["']style["'][^>]*>/gi;

  const extractHref = (tag) => pdf_getAttr_(tag, "href");

  // Collect stylesheet links
  h.replace(linkRe1, (tag) => {
    const href = extractHref(tag);
    if (href) cssLinks.push(href);
    return tag;
  });

  h.replace(linkRe2, (tag) => {
    const href = extractHref(tag);
    if (href) cssLinks.push(href);
    return tag;
  });

  // Cap number of CSS files
  const maxFiles = opts.maxCssFiles || 3;
  const unique = [];
  cssLinks.forEach(u => { if (unique.indexOf(u) === -1) unique.push(u); });

  const chosen = unique.slice(0, maxFiles);

  let combined = "";
  for (let i = 0; i < chosen.length; i++) {
    const href = chosen[i];
    let abs;
    try { abs = new URL(href, baseUrl).href; } catch (e) { continue; }

    const cssTxt = pdf_fetchText_(abs, opts.maxCssBytes || 200000);
    if (!cssTxt) continue;

    let expanded = pdf_expandCssImports_(cssTxt, abs, opts);
    expanded = pdf_rewriteCssUrlsToAbsolute_(expanded, abs);
    if (opts.inlineCssBgImages) expanded = pdf_inlineCssBackgroundImages_(expanded, abs, opts);

    combined += `\n/* inlined: ${abs} */\n` + expanded + "\n";
  }

  return combined;
}

function pdf_fetchText_(url, maxChars) {
  try {
    const r = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "User-Agent": PDF_USER_AGENT, "Accept": "text/css,*/*;q=0.8" }
    });
    const code = r.getResponseCode();
    if (code >= 400) return "";
    const txt = r.getContentText();
    if (!txt) return "";
    if (maxChars && txt.length > maxChars) return "";
    return txt;
  } catch (e) {
    return "";
  }
}

function pdf_expandCssImports_(cssText, cssBaseUrl, opts) {
  const maxImports = opts.maxCssImports || 2;
  let importsDone = 0;

  // @import "x";  or  @import url("x");
  const importRe = /@import\s+(?:url\()?["']([^"']+)["']\)?\s*;/gi;

  return String(cssText).replace(importRe, (m, importUrl) => {
    if (importsDone >= maxImports) return ""; // drop extra imports
    importsDone++;

    let abs;
    try { abs = new URL(importUrl, cssBaseUrl).href; } catch (e) { return ""; }
    const txt = pdf_fetchText_(abs, opts.maxCssBytes || 200000);
    if (!txt) return "";
    const expanded = pdf_rewriteCssUrlsToAbsolute_(txt, abs);
    return `\n/* inlined import: ${abs} */\n${expanded}\n`;
  });
}

function pdf_rewriteCssUrlsToAbsolute_(cssText, cssBaseUrl) {
  // Rewrite url(...) to absolute, but keep data: and about: and # anchors
  return String(cssText).replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (m, q, u) => {
    const raw = String(u || "").trim();
    if (!raw) return m;
    if (/^(data:|about:|#)/i.test(raw)) return m;
    try {
      const abs = new URL(raw, cssBaseUrl).href;
      return `url("${abs}")`;
    } catch (e) {
      return m;
    }
  });
}

function pdf_inlineCssBackgroundImages_(cssText, cssBaseUrl, opts) {
  const maxImgs = opts.maxCssBgImages || 8;
  const maxBytes = opts.maxCssBgImageBytes || 120000;
  let done = 0;

  return String(cssText).replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (m, q, u) => {
    if (done >= maxImgs) return m;

    const raw = String(u || "").trim();
    if (!raw) return m;
    if (/^(data:|about:|#)/i.test(raw)) return m;
    if (!/\.(png|jpg|jpeg|webp|gif|svg)(\?|#|$)/i.test(raw)) return m;

    let abs;
    try { abs = new URL(raw, cssBaseUrl).href; } catch (e) { return m; }

    const b = pdf_fetchBinary_(abs, maxBytes);
    if (!b || !b.bytes || !b.ct) return m;

    done++;
    const b64 = Utilities.base64Encode(b.bytes);
    return `url("data:${b.ct};base64,${b64}")`;
  });
}

function pdf_fetchBinary_(url, maxBytes) {
  try {
    const r = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { "User-Agent": PDF_USER_AGENT, "Accept": "image/*,*/*;q=0.8" }
    });
    const code = r.getResponseCode();
    if (code >= 400) return null;

    const headers = r.getHeaders();
    const ct = (headers["Content-Type"] || headers["content-type"] || "").toLowerCase();
    if (ct.indexOf("image/") !== 0 && ct.indexOf("font/") !== 0 && ct.indexOf("application/font") !== 0) return null;

    const bytes = r.getContent();
    if (maxBytes && bytes.length > maxBytes) return null;

    return { bytes: bytes, ct: ct };
  } catch (e) {
    return null;
  }
}

/* =========================
 * Inlining: HTML Images + inline style background images
 * ========================= */

function pdf_inlineImages_(html, baseUrl, opts) {
  let h = String(html);

  let count = 0;
  let totalBytes = 0;

  h = h.replace(/<img\b[^>]*>/gi, (tag) => {
    if (count >= (opts.maxImages || 10)) return tag;

    const candidates = pdf_collectImgCandidates_(tag);
    if (!candidates.length) return tag;

    // Try candidates from best → smaller until we find one that fits maxImageBytes and total cap
    let chosen = null;
    for (let i = 0; i < candidates.length; i++) {
      const src = candidates[i];
      if (!src || /^data:/i.test(src)) { chosen = src; break; }

      let abs;
      try { abs = new URL(src, baseUrl).href; } catch (e) { continue; }

      // Fetch without exceeding caps
      const b = pdf_fetchBinary_(abs, opts.maxImageBytes || 600000);
      if (!b || !b.bytes || !b.ct || b.ct.indexOf("image/") !== 0) {
        // If fetch failed, at least absolutize src
        continue;
      }

      if (opts.maxTotalInlineImageBytes && (totalBytes + b.bytes.length) > opts.maxTotalInlineImageBytes) {
        // total cap exceeded, stop inlining more
        break;
      }

      const b64 = Utilities.base64Encode(b.bytes);
      const dataUri = `data:${b.ct};base64,${b64}`;
      chosen = dataUri;
      totalBytes += b.bytes.length;
      break;
    }

    if (!chosen) {
      // no inline; just absolutize best candidate
      try {
        const abs = new URL(candidates[0], baseUrl).href;
        return pdf_setAttr_(tag, "src", abs);
      } catch (e) {
        return tag;
      }
    }

    count++;

    let newTag = pdf_setAttr_(tag, "src", chosen);
    // Drop srcset/sizes/lazy attrs after inlining
    newTag = newTag.replace(/\s(srcset|sizes|data-src|data-original|data-lazy-src|data-srcset|data-sizes)=["'][^"']*["']/gi, "");
    return newTag;
  });

  return h;
}

function pdf_collectImgCandidates_(imgTag) {
  const tag = String(imgTag);

  // 1) data-src / lazy variants
  const dataSrc = pdf_getAttr_(tag, "data-src") ||
                  pdf_getAttr_(tag, "data-lazy-src") ||
                  pdf_getAttr_(tag, "data-original");
  // 2) src
  const src = pdf_getAttr_(tag, "src");

  // 3) data-srcset / srcset
  const srcset = pdf_getAttr_(tag, "data-srcset") || pdf_getAttr_(tag, "srcset");

  const list = [];

  // Prefer srcset-based largest first (but keep src fallback)
  const fromSrcset = pdf_parseSrcset_(srcset);
  fromSrcset.forEach(u => { if (u) list.push(u); });

  if (dataSrc) list.push(dataSrc);
  if (src && src !== "#" && !/^about:blank/i.test(src)) list.push(src);

  // Deduplicate while preserving order
  const out = [];
  list.forEach(u => {
    const uu = String(u).trim();
    if (!uu) return;
    if (out.indexOf(uu) === -1) out.push(uu);
  });

  return out;
}

function pdf_parseSrcset_(srcset) {
  const s = String(srcset || "").trim();
  if (!s) return [];

  // Parse like: url1 320w, url2 640w, url3 1024w
  const parts = s.split(",").map(x => x.trim()).filter(Boolean);
  const items = parts.map(p => {
    const bits = p.split(/\s+/).filter(Boolean);
    const url = bits[0];
    let score = 0;
    const desc = bits[1] || "";
    const mW = desc.match(/^(\d+)w$/i);
    const mX = desc.match(/^(\d+(\.\d+)?)x$/i);
    if (mW) score = parseInt(mW[1], 10);
    else if (mX) score = Math.round(parseFloat(mX[1]) * 1000);
    return { url, score };
  });

  // Sort largest → smallest
  items.sort((a, b) => (b.score || 0) - (a.score || 0));

  return items.map(it => it.url);
}

function pdf_inlineStyleBackgroundImages_(html, baseUrl, opts) {
  // Inline background-image URLs in style="" attributes (small only)
  const maxBytes = opts.maxCssBgImageBytes || 120000;
  let done = 0;
  const maxImgs = opts.maxCssBgImages || 10;

  return String(html).replace(/\sstyle=["']([^"']*)["']/gi, (m, styleVal) => {
    if (done >= maxImgs) return m;
    let s = String(styleVal);

    s = s.replace(/url\(\s*(['"]?)([^'")]+)\1\s*\)/gi, (mm, q, u) => {
      if (done >= maxImgs) return mm;
      const raw = String(u || "").trim();
      if (!raw || /^(data:|about:|#)/i.test(raw)) return mm;
      if (!/\.(png|jpg|jpeg|webp|gif|svg)(\?|#|$)/i.test(raw)) return mm;

      let abs;
      try { abs = new URL(raw, baseUrl).href; } catch (e) { return mm; }

      const b = pdf_fetchBinary_(abs, maxBytes);
      if (!b || !b.bytes || !b.ct || b.ct.indexOf("image/") !== 0) return mm;

      done++;
      const b64 = Utilities.base64Encode(b.bytes);
      return `url("data:${b.ct};base64,${b64}")`;
    });

    return ` style="${pdf_escapeAttr_(s)}"`;
  });
}

/* =========================
 * Convert HTML → PDF
 * ========================= */

function pdf_htmlToPdfBlob_(html, pdfName) {
  const out = HtmlService.createHtmlOutput(html);
  return out.getAs(MimeType.PDF).setName(pdfName);
}

/* =========================
 * Small helpers
 * ========================= */

function pdf_getAttr_(tag, attr) {
  const re = new RegExp(`\\b${attr}\\s*=\\s*["']([^"']*)["']`, "i");
  const m = String(tag).match(re);
  return m ? m[1] : "";
}

function pdf_setAttr_(tag, attr, value) {
  const has = new RegExp(`\\b${attr}\\s*=`, "i").test(tag);
  if (has) {
    return String(tag).replace(
      new RegExp(`\\b${attr}\\s*=\\s*["'][^"']*["']`, "i"),
      `${attr}="${pdf_escapeAttr_(value)}"`
    );
  }
  return String(tag).replace(/>$/, ` ${attr}="${pdf_escapeAttr_(value)}">`);
}

function pdf_wordCount_(html) {
  const text = pdf_stripTags_(html).replace(/\s+/g, " ").trim();
  if (!text) return 0;
  return text.split(" ").length;
}

function pdf_stripTags_(s) {
  return String(s).replace(/<[^>]+>/g, " ");
}

function pdf_decodeHtmlEntities_(s) {
  return String(s)
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function pdf_escapeHtml_(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function pdf_escapeAttr_(s) {
  return pdf_escapeHtml_(s).replace(/`/g, "&#96;");
}

function pdf_safeFilename_(name) {
  const cleaned = String(name || "Article")
    .replace(/[\\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.slice(0, 140) || "Article";
}
