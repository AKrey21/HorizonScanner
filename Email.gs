/******************************
 * Email.gs — Email HTML builders + EML builder
 *
 * v1: simple HTML email (external images)
 * v2: EmailTemplateBody.html (Apps Script template) + CID inline images
 * v3: Adds OFFLINE PDF attachments (via pdf.gs) + EML multipart/mixed
 *
 * Notes:
 * - Attachments in Outlook EML require multipart/mixed (v3 builder below).
 * - CID inline images require multipart/related (nested inside mixed).
 * - PDF generation is best-effort; failures are SKIPPED per-topic (won’t break the whole email),
 *   but we CAPTURE + LOG the PDF error so you can see what failed.
 *
 * Banner:
 * - NO Base64
 * - NO BannerConfig.gs
 * - Banner is a Drive file (by ID) attached as CID "fs_banner"
 * - EmailTemplateBody.html should reference it as: <img src="cid:fs_banner" ...>
 ******************************/

// Optional: where to save generated .eml drafts (Drive folder)
const EML_OUTPUT_FOLDER_ID = "PASTE_EML_OUTPUT_FOLDER_ID_HERE"; // optional

// ===== FutureScans banner (Drive asset) =====
const FS_BANNER_FILE_ID = "13g5m1yI6V5Fcjvl57aEUZcHYq1lfnk6g"; // Banner.jpg in Drive
const FS_BANNER_FOLDER_ID = "1yuFMdayHpj9UB2qM14CORFii6LUCVJxp"; // Banner folder in Drive

// Browser-ish UA for image fetch fallback
const EMAIL_UA =
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36";

// ===== Typography (Outlook / Word engine) =====
const FS_FONT_STACK = `"Aptos","Aptos Display","Segoe UI",Arial,sans-serif`;
const FS_SECTION_PT = "12pt";
const FS_BODY_PT    = "11pt";

/**
 * Safety net: enforce Aptos + 11pt body + 12pt section bars
 * Useful if any Calibri/inline font sneaks in (e.g. from AI HTML).
 *
 * Note: Your EmailTemplateBody.html should already set these; this just makes it harder
 * for Outlook/Word to override.
 */
function enforceFsTypography_(html) {
  let out = String(html || "");

  // Replace Calibri occurrences from Word-ish HTML
  out = out.replace(/Calibri/gi, "Aptos");

  // If template already contains our fs-section CSS, don't inject duplicates
  const alreadyHasFsCss = out.includes(".fs-section") && out.toLowerCase().includes("aptos");
  if (!alreadyHasFsCss && !out.includes("FS_TYPO_GUARD_V1")) {
    const css =
      `\n<!-- FS_TYPO_GUARD_V1 -->\n` +
      `<style type="text/css">\n` +
      `  body, table, td, div, p, span, li, a {\n` +
      `    font-family: ${FS_FONT_STACK} !important;\n` +
      `    font-size: ${FS_BODY_PT} !important;\n` +
      `  }\n` +
      `  .fs-section { font-size: ${FS_SECTION_PT} !important; font-weight: 700 !important; }\n` +
      `</style>\n`;

    if (/<\/head>/i.test(out)) out = out.replace(/<\/head>/i, css + "</head>");
    else out = css + out;
  }

  return out;
}

/* =========================================================================
 * EMAIL HTML (Outlook-friendly): v1 (simple)
 * - Updated to match template typography: Aptos, 12pt headers, 11pt body
 * ========================================================================= */

function buildFuturescansEmailHtml_v1_(topics, weekLabel) {
  const headerImgUrl = ""; // optional external header image

  const intro =
    "FutureScans@MOM is a weekly collection of articles covering topics relevant to MOM’s work, curated by SPTD. " +
    "The summaries are AI-generated and reviewed by the editorial team.";

  const blocks = (topics || []).map(t => renderEmailTopicBlock_v1_(t)).join("\n");

  return `
<div style="margin:0;padding:0;background:#ffffff;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
    <tr>
      <td align="center" style="padding:0;">
        <table role="presentation" width="900" cellpadding="0" cellspacing="0" style="border-collapse:collapse;max-width:900px;width:100%;">

          ${headerImgUrl ? `
          <tr>
            <td style="padding:0;">
              <img src="${escapeHtml_(headerImgUrl)}" width="900" style="width:100%;max-width:900px;display:block;border:0;" alt="FutureScans@MOM" />
            </td>
          </tr>` : ``}

          <tr>
            <td style="padding:14px 14px 8px 14px;font-family:${FS_FONT_STACK};color:#111;font-size:${FS_BODY_PT};">
              <div style="line-height:1.4;">${escapeHtml_(intro)}</div>
              <div style="margin-top:6px;color:#555;"><b>Week of ${escapeHtml_(weekLabel)}</b></div>
            </td>
          </tr>

          <tr>
            <td style="padding:0 14px 14px 14px;">
              ${blocks}
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</div>`;
}

function renderEmailTopicBlock_v1_(t) {
  const topicNo = t.topicNo || 1;

  // Topic 2 orange, others blue
  const barColor = (topicNo === 2) ? "#F57C00" : "#0D47A1";
  const linkStyle = "color:#1a73e8;text-decoration:underline;";

  const imgCell = t.imageUrl
    ? `<img src="${escapeHtml_(t.imageUrl)}" width="220" style="display:block;border-radius:10px;border:1px solid #ddd;max-height:130px;object-fit:cover;" alt="" />`
    : `<div style="width:220px;height:130px;border-radius:10px;border:1px dashed #ccc;color:#777;font-family:${FS_FONT_STACK};font-size:${FS_BODY_PT};line-height:130px;text-align:center;">No image</div>`;

  // Optional: if you pass { pdfAttached: true, pdfName: "..."} into t
  const pdfNote = t.pdfAttached
    ? `<div style="font-size:${FS_BODY_PT};color:#555;margin-top:10px;">
         <b>Offline PDF attached:</b> ${escapeHtml_(t.pdfName || "Article.pdf")}
       </div>`
    : ``;

  return `
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;border:1px solid #e6e6e6;border-radius:12px;overflow:hidden;margin:14px 0;">
    <tr>
      <td style="background:${barColor};color:#fff;font-weight:700;padding:8px 12px;font-family:${FS_FONT_STACK};font-size:${FS_SECTION_PT};">
        &lt;Topic ${topicNo}&gt;
      </td>
    </tr>

    <tr>
      <td style="padding:0;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
          <tr>
            <td valign="top" style="padding:12px 12px;font-family:${FS_FONT_STACK};color:#111;font-size:${FS_BODY_PT};">
              <div style="font-weight:700;margin-bottom:6px;">
                [${escapeHtml_(t.title || "Untitled")}]
              </div>

              <div style="color:#333;margin:8px 0;">
                ${escapeHtml_(t.relevance20 || "")}
              </div>

              <div style="line-height:1.45;color:#111;margin-top:10px;">
                ${sanitizeHtmlBasic_(t.summaryHtml || "")}
              </div>

              ${pdfNote}

              <div style="margin-top:12px;">
                Click <a href="${escapeHtml_(t.articleUrl || "#")}" style="${linkStyle}" target="_blank" rel="noopener">here</a> to read more.
              </div>
            </td>

            <td width="240" valign="top" style="padding:12px 12px;">
              ${imgCell}
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>`;
}

/* =========================================================================
 * EMAIL HTML (Outlook-friendly): v2 (EmailTemplateBody.html + CID images)
 *  - Generates PDF attachments per topic (offline)
 *  - Captures + logs PDF errors per-topic instead of silently skipping
 *
 * Banner:
 *  - attached from Drive ID as CID "fs_banner"
 *  - EmailTemplateBody.html should reference <img src="cid:fs_banner" ...>
 * ========================================================================= */

function buildEmailFromWordTemplate_v2_(topics, weekLabel, options) {
  const opts = Object.assign(
    {
      // PDF generation
      makePdf: true,
      pdfMode: "replica",        // "replica" | "reader"
      pdfInlineImages: true,
      pdfInlineCss: true,
      pdfMaxImages: 10,
      pdfMaxImageBytes: 650000,
      pdfMaxCssFiles: 2,
      pdfMaxCssBytes: 220000,

      // Inline images (CID)
      includeTopicImages: true
    },
    options || {}
  );

  const inlineImages = [];
  const attachments = []; // PDFs live here

  // ---- Banner (Drive asset → CID) ----
  // Always attach banner, and force JPG mime so Outlook renders it reliably.
  try {
    const bannerBlobSource = getFsBannerBlob_();
    if (!bannerBlobSource) throw new Error("No banner file found.");

    let bannerBlob = bannerBlobSource.getBlob();
    // Force correct metadata (Drive sometimes returns application/octet-stream)
    bannerBlob = bannerBlob
      .setName("fs_banner.jpg")
      .setContentType("image/jpeg");

    inlineImages.push({
      cid: "fs_banner",
      blob: bannerBlob,
      filename: "fs_banner" // ext added later by guessExtFromMime_
    });
  } catch (e) {
    Logger.log("[BANNER] Failed to attach banner: " + formatErr_(e));
  }

  // ---- Section labels for template ----
  const sectionLabels = [
    "Labour Market",
    "Gig Workers",
    "Jobs and Skills",
    "Workplace",
    "Other Labour Markets"
  ];

  const tplTopics = [];
  (topics || []).forEach((t, i) => {
    const topicNo = t.topicNo || (i + 1);
    const sectionTitle = sectionLabels[i] || `Topic ${topicNo}`;
    const barColor = (topicNo === 2) ? "#FF7300" : "#002A7B";

    // Topic image -> CID
    let imgCid = "";
    if (opts.includeTopicImages && t.imageUrl) {
      const blob = tryFetchImageBlobSafe_(t.imageUrl);
      if (blob) {
        const mime = (blob.getContentType && blob.getContentType()) || "";
        if (isOutlookInlineImageMime_(mime)) {
          imgCid = `fs_img_${topicNo}`;
          inlineImages.push({ cid: imgCid, blob: blob, filename: imgCid });
        }
      }
    }

    // Offline PDF attachment (via pdf.gs)
    let pdfName = "";
    let hasPdf = false;
    let pdfError = "";

    if (opts.makePdf && t.articleUrl) {
      if (typeof pdf_makeArticlePdf_ !== "function") {
        pdfError = "pdf_makeArticlePdf_() not found (pdf.gs missing or function name mismatch).";
      } else {
        try {
          const pdfRes = pdf_makeArticlePdf_(t.articleUrl, {
            mode: opts.pdfMode,
            inlineImages: opts.pdfInlineImages,
            inlineCss: opts.pdfInlineCss,
            maxImages: opts.pdfMaxImages,
            maxImageBytes: opts.pdfMaxImageBytes,
            maxCssFiles: opts.pdfMaxCssFiles,
            maxCssBytes: opts.pdfMaxCssBytes
          });

          if (pdfRes && pdfRes.pdfBlob) {
            const sourceName = pdfRes.siteName || email_sourceNameFromUrl_(t.articleUrl);
            const baseName = `Article ${topicNo} - ${t.title || "Untitled"}` +
              (sourceName ? ` - ${sourceName}` : "");
            pdfName = `${safeFilename_(baseName)}.pdf`;
            attachments.push({
              filename: pdfName,
              blob: pdfRes.pdfBlob
            });
            hasPdf = true;
          } else {
            pdfError = "PDF generator returned no pdfBlob (site may block fetch / paywall / empty HTML).";
          }
        } catch (e) {
          pdfError = formatErr_(e);
        }
      }
    }

    if (!hasPdf && pdfError) {
      Logger.log(
        `[PDF FAIL] Topic ${topicNo} — ${t.title || "Untitled"} — ${t.articleUrl || ""} — ${pdfError}`
      );
    }

    const attachmentLabel = hasPdf ? (pdfName || `Article ${topicNo} - ${t.title || "Untitled"}.pdf`) : "";

    tplTopics.push({
      topicNo,
      sectionTitle,
      barColor,
      title: t.title || "Untitled",
      relevance20: t.relevance20 || "",
      summaryHtml: sanitizeHtmlBasic_(t.summaryHtml || ""),
      articleUrl: t.articleUrl || "#",
      imgCid,
      imgSrc: imgCid ? `cid:${imgCid}` : "",
      hasPdf,
      attachmentLabel,
      pdfError
    });
  });

  // ---- Render EmailTemplateBody.html as Apps Script Template ----
  const tpl = HtmlService.createTemplateFromFile("EmailTemplateBody");
  tpl.weekOf = weekLabel;
  tpl.bannerSrc = "cid:fs_banner";
  tpl.topics = tplTopics;

  // ✅ Post-process safety net (prevents Calibri + enforces 11/12pt if needed)
  let htmlBody = tpl.evaluate().getContent();
  htmlBody = enforceFsTypography_(htmlBody);

  return { htmlBody, inlineImages, attachments };
}

/* =========================================================================
 * Preview HTML builder (browser-safe)
 * ========================================================================= */

function buildFuturescansPreviewHtml_(topics, weekLabel, options) {
  const opts = Object.assign(
    {
      bannerSrc: buildPreviewBannerDataUrl_(),
      imagePlaceholderSrc: buildPreviewImageDataUrl_("Article image")
    },
    options || {}
  );

  const sectionLabels = [
    "Labour Market",
    "Gig Workers",
    "Jobs and Skills",
    "Workplace",
    "Other Labour Markets"
  ];

  const tplTopics = (topics || []).map((t, i) => {
    const topicNo = t.topicNo || (i + 1);
    const sectionTitle = sectionLabels[i] || `Topic ${topicNo}`;
    const barColor = (topicNo === 2) ? "#FF7300" : "#002A7B";
    const attachmentLabel = t.attachmentLabel ||
      `Article ${topicNo} – ${(t.title || "Untitled").slice(0, 40)}.pdf`;

    return {
      topicNo,
      sectionTitle,
      barColor,
      title: t.title || "Untitled",
      relevance20: t.relevance20 || "",
      summaryHtml: sanitizeHtmlBasic_(t.summaryHtml || ""),
      articleUrl: t.articleUrl || "#",
      imgSrc: t.imageUrl || opts.imagePlaceholderSrc,
      attachmentLabel
    };
  });

  const tpl = HtmlService.createTemplateFromFile("EmailTemplateBody");
  tpl.weekOf = weekLabel;
  tpl.bannerSrc = opts.bannerSrc;
  tpl.topics = tplTopics;

  return tpl.evaluate().getContent();
}

function buildPreviewBannerDataUrl_() {
  try {
    const bannerFile = getFsBannerBlob_();
    if (bannerFile) {
      const blob = bannerFile.getBlob();
      const bytes = blob.getBytes();
      const b64 = Utilities.base64Encode(bytes);
      return `data:${blob.getContentType()};base64,${b64}`;
    }
  } catch (e) {
    Logger.log("[PREVIEW] Banner load failed: " + formatErr_(e));
  }

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="966" height="278">` +
    `<defs><linearGradient id="g" x1="0" x2="1"><stop offset="0" stop-color="#002A7B"/>` +
    `<stop offset="1" stop-color="#0D47A1"/></linearGradient></defs>` +
    `<rect width="966" height="278" fill="url(#g)"/>` +
    `<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" ` +
    `font-family="Aptos,Segoe UI,Arial" font-size="36" fill="#ffffff">` +
    `FutureScans@MOM</text></svg>`;
  return "data:image/svg+xml;base64," + Utilities.base64Encode(svg);
}

function buildPreviewImageDataUrl_(label) {
  const safeLabel = String(label || "Image").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="246" height="140">` +
    `<rect width="246" height="140" fill="#f2f4f8" stroke="#c7c9cc"/>` +
    `<text x="50%" y="50%" dominant-baseline="middle" text-anchor="middle" ` +
    `font-family="Aptos,Segoe UI,Arial" font-size="12" fill="#4a4a4a">` +
    `${safeLabel}</text></svg>`;
  return "data:image/svg+xml;base64," + Utilities.base64Encode(svg);
}

function getFsBannerBlob_() {
  if (FS_BANNER_FOLDER_ID && FS_BANNER_FOLDER_ID !== "PASTE_BANNER_FOLDER_ID_HERE") {
    const folder = DriveApp.getFolderById(FS_BANNER_FOLDER_ID);
    const files = folder.getFiles();
    if (files.hasNext()) return files.next();
  }

  if (FS_BANNER_FILE_ID && FS_BANNER_FILE_ID !== "PASTE_BANNER_FILE_ID_HERE") {
    return DriveApp.getFileById(FS_BANNER_FILE_ID);
  }

  return null;
}

/**
 * Unified “package” builder for Outlook draft
 * Returns everything you need to generate an EML with inline images + PDFs.
 */
function buildFuturescansEmailPackage_v3_(subject, topics, weekLabel, options) {
  const built = buildEmailFromWordTemplate_v2_(topics, weekLabel, options);
  return {
    subject,
    htmlBody: built.htmlBody,
    inlineImages: built.inlineImages || [],
    attachments: built.attachments || []
  };
}

/* =========================================================================
 * EML builder v2 (legacy): multipart/related + CID inline images
 * ========================================================================= */

function buildEmlMultipartRelated_v2_(subject, htmlBody, inlineImages) {
  const relBoundary = "REL_" + Utilities.getUuid().replace(/-/g, "");
  const altBoundary = "ALT_" + Utilities.getUuid().replace(/-/g, "");
  const CRLF = "\r\n";

  const altPart =
    `--${altBoundary}${CRLF}` +
    `Content-Type: text/html; charset="UTF-8"${CRLF}` +
    `Content-Transfer-Encoding: quoted-printable${CRLF}${CRLF}` +
    toQuotedPrintable_(htmlBody) + CRLF +
    `--${altBoundary}--${CRLF}`;

  let eml =
    `Subject: ${subject}${CRLF}` +
    `X-Unsent: 1${CRLF}` +
    `MIME-Version: 1.0${CRLF}` +
    `Content-Type: multipart/related; boundary="${relBoundary}"${CRLF}` +
    `${CRLF}` +

    `--${relBoundary}${CRLF}` +
    `Content-Type: multipart/alternative; boundary="${altBoundary}"${CRLF}${CRLF}` +
    altPart + CRLF;

  (inlineImages || []).forEach((it) => {
    if (!it || !it.cid || !it.blob) return;

    const blob = it.blob;
    const contentType = (blob.getContentType && blob.getContentType()) || "application/octet-stream";
    const ext = guessExtFromMime_(contentType);
    const filename = (it.filename || it.cid || "image") + ext;

    const bytes = blob.getBytes();
    const b64 = wrapBase64_(Utilities.base64Encode(bytes));

    eml +=
      `--${relBoundary}${CRLF}` +
      `Content-Type: ${contentType}; name="${filename}"${CRLF}` +
      `Content-Transfer-Encoding: base64${CRLF}` +
      `Content-ID: <${it.cid}>${CRLF}` +
      `Content-Location: ${filename}${CRLF}` +
      `Content-Disposition: inline; filename="${filename}"${CRLF}${CRLF}` +
      b64 + CRLF;
  });

  eml += `--${relBoundary}--${CRLF}`;
  return eml;
}

/* =========================================================================
 * EML builder v3: multipart/mixed
 *  - Part 1: multipart/related (HTML + CID images)
 *  - Part 2..n: attachments (PDFs)
 * ========================================================================= */

function buildEmlMultipartMixed_v3_(subject, htmlBody, inlineImages, attachments) {
  const mixBoundary = "MIX_" + Utilities.getUuid().replace(/-/g, "");
  const relBoundary = "REL_" + Utilities.getUuid().replace(/-/g, "");
  const altBoundary = "ALT_" + Utilities.getUuid().replace(/-/g, "");
  const CRLF = "\r\n";

  const altPart =
    `--${altBoundary}${CRLF}` +
    `Content-Type: text/html; charset="UTF-8"${CRLF}` +
    `Content-Transfer-Encoding: quoted-printable${CRLF}${CRLF}` +
    toQuotedPrintable_(htmlBody) + CRLF +
    `--${altBoundary}--${CRLF}`;

  let eml =
    `Subject: ${subject}${CRLF}` +
    `X-Unsent: 1${CRLF}` +
    `MIME-Version: 1.0${CRLF}` +
    `Content-Type: multipart/mixed; boundary="${mixBoundary}"${CRLF}` +
    `${CRLF}`;

  // Part 1: related (html + inline images)
  eml +=
    `--${mixBoundary}${CRLF}` +
    `Content-Type: multipart/related; boundary="${relBoundary}"${CRLF}${CRLF}` +

    `--${relBoundary}${CRLF}` +
    `Content-Type: multipart/alternative; boundary="${altBoundary}"${CRLF}${CRLF}` +
    altPart + CRLF;

  (inlineImages || []).forEach((it) => {
    if (!it || !it.cid || !it.blob) return;

    const blob = it.blob;
    const contentType = (blob.getContentType && blob.getContentType()) || "application/octet-stream";
    const ext = guessExtFromMime_(contentType);
    const filename = (it.filename || it.cid || "image") + ext;

    const bytes = blob.getBytes();
    const b64 = wrapBase64_(Utilities.base64Encode(bytes));

    eml +=
      `--${relBoundary}${CRLF}` +
      `Content-Type: ${contentType}; name="${filename}"${CRLF}` +
      `Content-Transfer-Encoding: base64${CRLF}` +
      `Content-ID: <${it.cid}>${CRLF}` +
      `Content-Location: ${filename}${CRLF}` +              // ✅ ADDED
      `Content-Disposition: inline; filename="${filename}"${CRLF}${CRLF}` +
      b64 + CRLF;
  });

  eml += `--${relBoundary}--${CRLF}`;

  // Attachments (PDFs)
  (attachments || []).forEach((a) => {
    if (!a || !a.blob) return;
    const blob = a.blob;

    let contentType = (blob.getContentType && blob.getContentType()) || "application/octet-stream";
    if (!contentType || contentType === "application/octet-stream") {
      const maybeName = String(a.filename || blob.getName() || "").toLowerCase();
      if (maybeName.endsWith(".pdf")) contentType = "application/pdf";
    }

    const rawName = a.filename || blob.getName() || "attachment";
    const filename = safeFilename_(rawName);

    const bytes = blob.getBytes();
    const b64 = wrapBase64_(Utilities.base64Encode(bytes));

    eml +=
      `--${mixBoundary}${CRLF}` +
      `Content-Type: ${contentType}; name="${filename}"${CRLF}` +
      `Content-Transfer-Encoding: base64${CRLF}` +
      `Content-Disposition: attachment; filename="${filename}"${CRLF}${CRLF}` +
      b64 + CRLF;
  });

  eml += `--${mixBoundary}--${CRLF}`;
  return eml;
}

/**
 * Convenience: create an Outlook draft .eml file in Drive
 * Returns {fileId, fileUrl, name}
 */
function saveOutlookDraftEml_v3_(subject, topics, weekLabel, options) {
  const pkg = buildFuturescansEmailPackage_v3_(subject, topics, weekLabel, options);

  const eml = buildEmlMultipartMixed_v3_(
    pkg.subject,
    pkg.htmlBody,
    pkg.inlineImages,
    pkg.attachments
  );

  const name = safeFilename_(`${subject}.eml`);
  const blob = Utilities.newBlob(eml, "message/rfc822", name);

  let file;
  if (EML_OUTPUT_FOLDER_ID && EML_OUTPUT_FOLDER_ID !== "PASTE_EML_OUTPUT_FOLDER_ID_HERE") {
    file = DriveApp.getFolderById(EML_OUTPUT_FOLDER_ID).createFile(blob);
  } else {
    file = DriveApp.createFile(blob);
  }

  return { fileId: file.getId(), fileUrl: file.getUrl(), name };
}

/**
 * Quick sanity test: run once from Apps Script editor.
 * Confirms PDFs attach (if pdf.gs is present).
 */
function test_saveOutlookDraftEml_v3_() {
  const topics = [{
    topicNo: 1,
    title: "Test Article",
    relevance20: "Relevance line here",
    summaryHtml: "<p>Hello from FutureScans.</p>",
    articleUrl: "https://example.com",
    imageUrl: ""
  }];

  const res = saveOutlookDraftEml_v3_(
    "FutureScans Test Draft",
    topics,
    "14 Jan 2026",
    { makePdf: true, pdfMode: "reader" }
  );

  Logger.log(res);
}

/* =========================================================================
 * Helpers
 * ========================================================================= */

function formatErr_(e) {
  try {
    if (!e) return "Unknown error";
    const msg = (e && e.message) ? e.message : String(e);
    return String(msg).slice(0, 500);
  } catch (_) {
    return "Unknown error";
  }
}

function email_sourceNameFromUrl_(url) {
  if (!url) return "";
  try {
    const host = new URL(url).hostname || "";
    const clean = host.replace(/^www\./i, "");
    if (!clean) return "";

    const parts = clean.split(".").filter(Boolean);
    if (!parts.length) return "";

    const tld = parts[parts.length - 1];
    const sld = parts[parts.length - 2];
    const knownSecondLevel = ["co", "com", "org", "net", "gov", "edu"];
    const base =
      (tld && tld.length === 2 && knownSecondLevel.includes(sld) && parts.length >= 3)
        ? parts[parts.length - 3]
        : (sld || parts[0]);

    if (!base) return clean;

    return base
      .replace(/[-_]+/g, " ")
      .split(" ")
      .map(w => w ? w[0].toUpperCase() + w.slice(1) : "")
      .join(" ")
      .trim();
  } catch (e) {
    return "";
  }
}

function isOutlookInlineImageMime_(mime) {
  const m = String(mime || "").toLowerCase();
  return (
    m.includes("image/png") ||
    m.includes("image/jpeg") ||
    m.includes("image/jpg") ||
    m.includes("image/gif")
  );
}

function guessExtFromMime_(mime) {
  const m = String(mime || "").toLowerCase();
  if (m.includes("png")) return ".png";
  if (m.includes("gif")) return ".gif";
  if (m.includes("jpeg") || m.includes("jpg")) return ".jpg";
  if (m.includes("webp")) return ".webp";
  return "";
}

function wrapBase64_(b64) {
  const CRLF = "\r\n";
  return String(b64 || "").replace(/(.{76})/g, "$1" + CRLF);
}

function toQuotedPrintable_(str) {
  const s = String(str || "");
  const bytes = Utilities.newBlob(s, "UTF-8").getBytes();
  let out = "";
  let lineLen = 0;

  for (let i = 0; i < bytes.length; i++) {
    let b = bytes[i];
    if (b < 0) b += 256;

    let chunk;

    if ((b >= 33 && b <= 60) || (b >= 62 && b <= 126) || b === 9 || b === 32) {
      chunk = String.fromCharCode(b);
    } else if (b === 13) {
      continue;
    } else if (b === 10) {
      out += "\r\n";
      lineLen = 0;
      continue;
    } else {
      chunk = "=" + b.toString(16).toUpperCase().padStart(2, "0");
    }

    if (lineLen + chunk.length > 72) {
      out += "=\r\n";
      lineLen = 0;
    }

    out += chunk;
    lineLen += chunk.length;
  }

  return out;
}

function escapeHtml_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function sanitizeHtmlBasic_(html) {
  let h = String(html || "");
  h = h.replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "");
  h = h.replace(/<iframe\b[^>]*>[\s\S]*?<\/iframe>/gi, "");
  h = h.replace(/\son\w+=["'][^"']*["']/gi, "");
  return h;
}

function safeFilename_(name) {
  const cleaned = String(name || "file")
    .replace(/[\\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return (cleaned.slice(0, 140) || "file");
}

function tryFetchImageBlobSafe_(url) {
  if (!url) return null;

  try {
    if (typeof tryFetchImageBlob_ === "function") {
      return tryFetchImageBlob_(url);
    }
  } catch (e) {
    // ignore and fallback
  }

  try {
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        "User-Agent": EMAIL_UA,
        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8"
      }
    });
    if (resp.getResponseCode() >= 400) return null;
    return resp.getBlob();
  } catch (e2) {
    return null;
  }
}
