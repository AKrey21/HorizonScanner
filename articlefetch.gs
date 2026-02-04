/******************************
 * ArticleFetch.gs
 ******************************/
function fetchArticleMeta_(url) {
  const html = UrlFetchApp.fetch(url, {
    followRedirects: true,
    muteHttpExceptions: true,
    headers: { "User-Agent": "Mozilla/5.0" }
  }).getContentText();

  // Title: prefer og:title, then <title>
  const ogTitle = matchMeta_(html, /<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["'][^>]*>/i);
  const titleTag = match_(html, /<title[^>]*>([^<]+)<\/title>/i);
  const title = cleanText_(ogTitle || titleTag || "Untitled");

  // Image candidates (best for report image)
  const imageCandidates = extractImageCandidates_(html);
  const ogImage = imageCandidates.length ? resolveUrl_(imageCandidates[0], url) : "";

  // Description (optional)
  const desc = matchMeta_(html, /<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)["'][^>]*>/i)
            || matchMeta_(html, /<meta[^>]+property=["']og:description["'][^>]+content=["']([^"']+)["'][^>]*>/i);

  // Main text extraction: lightweight heuristic
  // (If you already have weekly picks description/content, you can replace this with that.)
  const bodyText = extractReadableText_(html);

  return {
    url,
    title,
    description: cleanText_(desc || ""),
    ogImage: ogImage || "",
    text: bodyText
  };
}

function extractImageCandidates_(html) {
  const candidates = [];
  const add = (value) => {
    const clean = cleanText_(value || "");
    if (!clean) return;
    if (clean.startsWith("data:")) return;
    candidates.push(clean);
  };

  add(matchMeta_(html, /<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["'][^>]*>/i));
  add(matchMeta_(html, /<meta[^>]+property=["']og:image:secure_url["'][^>]+content=["']([^"']+)["'][^>]*>/i));
  add(matchMeta_(html, /<meta[^>]+name=["']twitter:image["'][^>]+content=["']([^"']+)["'][^>]*>/i));
  add(matchMeta_(html, /<meta[^>]+name=["']twitter:image:src["'][^>]+content=["']([^"']+)["'][^>]*>/i));
  add(matchMeta_(html, /<meta[^>]+itemprop=["']image["'][^>]+content=["']([^"']+)["'][^>]*>/i));
  add(match_(html, /<link[^>]+rel=["']image_src["'][^>]+href=["']([^"']+)["'][^>]*>/i));

  return candidates;
}

function resolveUrl_(value, baseUrl) {
  const raw = cleanText_(value || "");
  if (!raw) return "";
  try {
    return new URL(raw, baseUrl).toString();
  } catch (e) {
    return raw;
  }
}

function extractReadableText_(html) {
  // strip scripts/styles
  let t = html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ");

  // remove nav/footer/aside-ish blocks crudely
  t = t.replace(/<(nav|footer|aside)[\s\S]*?<\/\1>/gi, " ");

  // convert <br> and </p> into newlines
  t = t.replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>/gi, "\n");

  // remove tags
  t = t.replace(/<[^>]+>/g, " ");

  // clean whitespace
  t = t.replace(/[ \t]+/g, " ").replace(/\n{3,}/g, "\n\n").trim();

  // keep it bounded (Gemini + PDF)
  if (t.length > 12000) t = t.slice(0, 12000) + " ...";
  return t;
}

function match_(s, re) {
  const m = s.match(re);
  return m ? m[1] : "";
}

function matchMeta_(s, re) {
  const m = s.match(re);
  return m ? m[1] : "";
}

function cleanText_(s) {
  return (s || "").replace(/\s+/g, " ").trim();
}
