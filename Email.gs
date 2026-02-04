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
 * - Falls back to newest image in FS_BANNER_FOLDER_ID if ID is missing/invalid
 * - Resolves Drive shortcuts and can use thumbnails for non-image Drive types
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
    let bannerBlob = getFsBannerBlob_();
    if (!bannerBlob) throw new Error("No banner file found.");
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
    const sectionTitle = t.sectionTitle || sectionLabels[i] || `Topic ${topicNo}`;
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
    const sectionTitle = t.sectionTitle || sectionLabels[i] || `Topic ${topicNo}`;
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
    const bannerBlob = getFsBannerBlob_();
    if (bannerBlob) {
      let mime = bannerBlob.getContentType() || "image/jpeg";
      if (!mime || mime === "application/octet-stream" || !/^image\//i.test(mime)) {
        mime = "image/jpeg";
      }
      const b64 = Utilities.base64Encode(bannerBlob.getBytes());
      return `data:${mime};base64,${b64}`;
    }
  } catch (e) {
    // fall back to placeholder banner
  }

  const embeddedBanner =
    "data:image/jpeg;base64,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<snip>";
  if (embeddedBanner) {
    return embeddedBanner;
  }

  const bannerId = getFsBannerFileId_();
  if (bannerId) {
    return `https://drive.google.com/thumbnail?id=${encodeURIComponent(bannerId)}&sz=w2000`;
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

function normalizeDriveFileId_(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (!raw.includes("drive.google.com")) return raw;

  const match =
    raw.match(/\/d\/([a-zA-Z0-9_-]+)/) ||
    raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : raw;
}

function getFsBannerFileId_() {
  const id = normalizeDriveFileId_(FS_BANNER_FILE_ID);
  if (id && id !== "PASTE_BANNER_FILE_ID_HERE") {
    try {
      const file = DriveApp.getFileById(id);
      const resolved = resolveBannerFile_(file);
      if (resolved) return resolved.getId();
    } catch (e) {
      // fall through to folder lookup
    }
  }

  const folderId = normalizeDriveFileId_(FS_BANNER_FOLDER_ID);
  if (!folderId || folderId === "PASTE_BANNER_FOLDER_ID_HERE") return "";

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    let newest = null;
    let newestTime = 0;

    while (files.hasNext()) {
      const file = files.next();
      const resolved = resolveBannerFile_(file);
      if (!resolved) continue;

      const updated = resolved.getLastUpdated();
      const updatedTime = updated ? updated.getTime() : 0;
      if (!newest || updatedTime > newestTime) {
        newest = resolved;
        newestTime = updatedTime;
      }
    }

    return newest ? newest.getId() : "";
  } catch (e) {
    return "";
  }
}

function getFsBannerBlob_() {
  const bannerId = getFsBannerFileId_();
  if (!bannerId) return null;

  try {
    const file = DriveApp.getFileById(bannerId);
    const resolved = resolveBannerFile_(file);
    if (!resolved) return null;

    const mime = String(resolved.getMimeType() || "").toLowerCase();
    if (mime.startsWith("image/")) return resolved.getBlob();

    if (typeof resolved.getThumbnail === "function") {
      const thumb = resolved.getThumbnail();
      if (thumb) {
        return thumb.setName("fs_banner.jpg").setContentType("image/jpeg");
      }
    }
  } catch (e) {
    return null;
  }

  return null;
}

function resolveBannerFile_(file) {
  if (!file) return null;

  const mime = String(file.getMimeType() || "").toLowerCase();
  if (mime.startsWith("image/")) return file;

  if (mime === "application/vnd.google-apps.shortcut" && typeof file.getTargetId === "function") {
    try {
      const target = DriveApp.getFileById(file.getTargetId());
      return resolveBannerFile_(target);
    } catch (e) {
      return null;
    }
  }

  return file;
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
