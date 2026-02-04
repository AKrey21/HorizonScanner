/******************************
 * Report.gs — Futurescan UI Entry Points
 *
 * Keeps UI functions small.
 * Heavy lifting lives in:
 * - ReportLib.gs (link parsing, resolver, AI, preview rendering, helpers)
 * - Email.gs     (v1/v2 email HTML + EML multipart builder)
 *
 * Optional dependency:
 * - BannerConfig.gs: const FS_BANNER_FILE_ID = "YOUR_DRIVE_FILE_ID";
 ******************************/

// Optional: cap the number of links processed (kept at 5)
const MAX_REPORT_LINKS = 5;

/* =========================================================================
 * UI: Generate preview HTML (your existing report blocks)
 * ========================================================================= */

function ui_generateFuturescanReportFromLinks(linksText) {
  const topics = rpt_buildTopicsFromLinks_(linksText, {
    includePdf: ENABLE_PDF_GENERATION
  });

  const tz = Session.getScriptTimeZone();
  const weekLabel = Utilities.formatDate(new Date(), tz, "dd MMM yyyy");

  return buildFuturescansPreviewHtml_(topics, weekLabel);
}

function ui_getFuturescansPreviewSample() {
  const weekLabel = "DD MMM 2025";
  const topics = [
    {
      topicNo: 1,
      sectionTitle: "<Topic>",
      title: "[Title]",
      relevance20: "&lt;How are these topics or issues relevant to Singapore or MOM, 20 words&gt;",
      summaryHtml: "[AI-generated summary with 1-2 key points bolded, 80 words]",
      imageUrl: "",
      articleUrl: "#",
      attachmentLabel: "Article 1 \u2013 xxx.pdf"
    },
    {
      topicNo: 2,
      sectionTitle: "<Topic>",
      title: "[Title]",
      relevance20: "&lt;How are these topics or issues relevant to Singapore or MOM, 20 words&gt;",
      summaryHtml: "[AI-generated summary with 1-2 key points bolded, 80 words]",
      imageUrl: "",
      articleUrl: "#",
      attachmentLabel: "Article 2 \u2013 xxx.pdf"
    },
    {
      topicNo: 3,
      sectionTitle: "<Topic>",
      title: "[Title]",
      relevance20: "&lt;How are these topics or issues relevant to Singapore or MOM, 20 words&gt;",
      summaryHtml: "[AI-generated summary with 1-2 key points bolded, 80 words]",
      imageUrl: "",
      articleUrl: "#",
      attachmentLabel: "Article 3 \u2013 xxx.pdf"
    },
    {
      topicNo: 4,
      sectionTitle: "<Topic>",
      title: "[Title]",
      relevance20: "&lt;How are these topics or issues relevant to Singapore or MOM, 20 words&gt;",
      summaryHtml: "[AI-generated summary with 1-2 key points bolded, 80 words]",
      imageUrl: "",
      articleUrl: "#",
      attachmentLabel: "Article 4 \u2013 xxx.pdf"
    }
  ];

  return buildFuturescansPreviewHtml_(topics, weekLabel);
}

/* =========================================================================
 * UI: Generate Outlook draft (.eml) v1 — simple HTML (external images)
 * Returns: { filename, b64 }
 * ========================================================================= */

function ui_buildOutlookEmlFromLinks_v1(linksText) {
  const topics = rpt_buildTopicsFromLinks_(linksText, { includePdf: false });

  const tz = Session.getScriptTimeZone();
  const weekLabel = Utilities.formatDate(new Date(), tz, "dd MMM yyyy");
  const subject = `FutureScans@MOM — Week of ${weekLabel}`;

  const htmlBody = buildFuturescansEmailHtml_v1_(topics, weekLabel);

  // Minimal RFC822 EML content
  const eml =
    `Subject: ${subject}\r\n` +
    `X-Unsent: 1\r\n` +
    `MIME-Version: 1.0\r\n` +
    `Content-Type: text/html; charset="UTF-8"\r\n` +
    `Content-Transfer-Encoding: 8bit\r\n` +
    `\r\n` +
    htmlBody;

  const b64 = Utilities.base64EncodeWebSafe(eml);
  const filename = `FutureScans_${weekLabel.replace(/ /g, "_")}.eml`;
  return { filename, b64 };
}

/* =========================================================================
 * UI: Generate Outlook draft (.eml) v2 — EmailTemplateBody.html + CID images + PDFs
 * Returns: { filename, b64 }
 * ========================================================================= */

function ui_buildOutlookEmlFromLinks_v2(linksText) {
  const topics = rpt_buildTopicsFromLinks_(linksText, { includePdf: false });

  const tz = Session.getScriptTimeZone();
  const weekLabel = Utilities.formatDate(new Date(), tz, "dd MMM yyyy");
  const subject = `FutureScans@MOM — Week of ${weekLabel}`;

  const pkg = buildFuturescansEmailPackage_v3_(subject, topics, weekLabel, {
    makePdf: true,
    pdfMode: "replica"
  });
  const eml = buildEmlMultipartMixed_v3_(
    pkg.subject,
    pkg.htmlBody,
    pkg.inlineImages,
    pkg.attachments
  );

  const b64 = Utilities.base64EncodeWebSafe(eml);
  const filename = `FutureScans_${weekLabel.replace(/ /g, "_")}.eml`;
  return { filename, b64 };
}

/* =========================================================================
 * Optional: Banner debug helper (v2)
 * ========================================================================= */

function ui_debugBanner_v2() {
  const id = getFsBannerFileId_();
  if (!id) {
    return {
      ok: false,
      reason: "FS_BANNER_FILE_ID missing/empty (BannerConfig.gs not loaded or const not set)"
    };
  }

  try {
    const f = DriveApp.getFileById(id);
    const b = f.getBlob();
    return {
      ok: true,
      fileName: f.getName(),
      mime: b.getContentType(),
      sizeBytes: b.getBytes().length,
      outlookInlineAllowed: isOutlookInlineImageMime_(b.getContentType())
    };
  } catch (e) {
    return { ok: false, reason: String(e && e.message ? e.message : e) };
  }
}
