/* Runs inside outlook.office.com / outlook.live.com pages.
 * Best-effort scraping of the currently open message in the reading pane.
 * Outlook Web's DOM is not officially documented and its class names are
 * generated/obfuscated, so this uses several fallback strategies. Whatever
 * it can't find is simply left blank for the user to fill in manually in
 * the popup - nothing here is required to be perfect.
 */

function findReadingPane() {
  // Confirmed via live inspection: the reading pane is a landmark
  // role="main" labelled "Reading Pane" in current Outlook Web
  // (outlook.cloud.microsoft and outlook.office.com both use this).
  return (
    document.querySelector('[role="main"][aria-label="Reading Pane" i]') ||
    document.querySelector('[role="main"]') ||
    document.body
  );
}

function findMessageBodyElement() {
  const pane = findReadingPane();

  // Confirmed via live inspection: each message in the pane has a
  // role="document" element labelled "Message body". In a conversation
  // thread there can be several (one per message) - the first one is the
  // topmost/currently-open message, which is what we want.
  const bodies = pane.querySelectorAll('[role="document"][aria-label="Message body" i]');
  if (bodies.length > 0) return bodies[0];

  // Fallback for older Outlook Web builds: message body inside an iframe
  // whose id starts with "UniqueMessageBody".
  const iframes = document.querySelectorAll('iframe[id^="UniqueMessageBody"]');
  for (const f of iframes) {
    try {
      if (f.contentDocument && f.contentDocument.body && f.contentDocument.body.innerText.trim()) {
        return f.contentDocument.body;
      }
    } catch (e) {
      /* cross-origin iframe, skip */
    }
  }

  const docRole = pane.querySelector('[role="document"]');
  if (docRole) return docRole;

  // Last resort: the largest text block inside the reading pane.
  let best = null;
  let bestLen = 0;
  pane.querySelectorAll("div, section").forEach((el) => {
    const len = (el.innerText || "").length;
    if (len > bestLen && el.querySelectorAll("div,section").length < 400) {
      bestLen = len;
      best = el;
    }
  });
  return best;
}

function findSubject() {
  const pane = findReadingPane();

  // Confirmed via live inspection: the first role="heading" inside the
  // Reading Pane landmark is the message subject. Scoping to the pane
  // (rather than the whole document) is what avoids accidentally picking
  // up "Navigation pane" / "Inbox" headings from the sidebar.
  const heading = pane.querySelector('[role="heading"]');
  const text = heading ? (heading.innerText || "").trim() : "";
  if (text) return text;

  // Fallback: document.title is often "<something> - Outlook", not always
  // the subject, so only use it as a last resort.
  const title = document.title.replace(/\s*-\s*Outlook\s*$/i, "").trim();
  return title || "";
}

function findAttachmentNames() {
  // NOTE: not yet confirmed against a live email with real attachments -
  // this is a best-effort guess at likely markup patterns. If it comes up
  // empty on an email you know has attachments, that's expected for now;
  // the file picker in the popup still works regardless. Open dev tools on
  // such an email, inspect the attachment row, and share the markup to get
  // this tightened up.
  const pane = findReadingPane();
  const names = new Set();
  const exts =
    /\.(pdf|docx?|xlsx?|pptx?|png|jpe?g|gif|txt|csv|zip|msg|eml|heic|mp4|mov)$/i;

  const containers = pane.querySelectorAll('[aria-label*="attachment" i], [class*="attachment" i]');
  containers.forEach((c) => {
    const text = (c.getAttribute("title") || c.innerText || "").trim();
    if (text && exts.test(text) && text.length < 200) {
      names.add(text.split("\n")[0].trim());
    }
    c.querySelectorAll("[title]").forEach((el) => {
      const t = el.getAttribute("title").trim();
      if (exts.test(t) && t.length < 200) names.add(t);
    });
  });

  return Array.from(names);
}

// Best-effort sign-off / signature detector. Looks for a line that is
// *itself* essentially just a closing greeting ("Best regards," "Lep
// pozdrav," "Mit freundlichen Grüßen," etc.), a "sent from my phone" marker,
// a "--" delimiter, or the start of a confidentiality disclaimer. Deliberately
// anchored to whole-line matches (not substring matches inside real
// sentences) to avoid false-triggering on something like "Thanks for sending
// the report, see attached below."
const SIGN_OFF_LINE_PATTERNS = [
  /^--+\s*$/,
  /^(best|kind|kindest|warm|warmest|many)\s+regards[,.!]?$/i,
  /^regards[,.!]?$/i,
  /^best[,.!]?$/i,
  /^b\.?\s?r\.?[,.!]?$/i,
  /^sincerely[,.!]?$/i,
  /^cheers[,.!]?$/i,
  /^thanks?( you)?( in advance)?[,.!]?$/i,
  /^sent from my (iphone|android|mobile|samsung|ipad)\b/i,
  // Slovenian
  /^lep(o)? pozdrav(ljeni)?[,.!]?$/i,
  /^s spo[sš]tovanjem[,.!]?$/i,
  /^lepi? pozdravi[,.!]?$/i,
  // German
  /^mit freundlichen gr[uü][sß]en[,.!]?$/i,
  /^(viele|liebe|beste) gr[uü][sß]e[,.!]?$/i,
  // Croatian
  /^s poštovanjem[,.!]?$/i,
  /^lijep pozdrav[,.!]?$/i,
];

// Substring (not whole-line) patterns for the start of legal disclaimers,
// which usually open mid-paragraph rather than as a standalone line.
const DISCLAIMER_PATTERNS = [
  /this\s+(e-?mail|message|communication)\b.{0,80}(confidential|privileged|intended (solely|only) for)/i,
  /confidentiality notice/i,
  /disclaimer\s*:/i,
];

function isSignOffLine(text) {
  return SIGN_OFF_LINE_PATTERNS.some((p) => p.test(text)) || DISCLAIMER_PATTERNS.some((p) => p.test(text));
}

// Finds the actual DOM element containing the sign-off line, so images that
// appear after it (a logo, social icons, a banner in the signature block)
// can be excluded by document position - not just by text. Outlook Web
// commonly puts each visual line in its own block-level element, so we scan
// small, mostly-leaf elements in document order and test each one's own
// text against the same patterns used for the plain-text trim.
function findSignatureBoundaryElement(bodyEl) {
  if (!bodyEl) return null;
  const candidates = bodyEl.querySelectorAll("div, p, li, td, span");
  for (const el of candidates) {
    const text = (el.textContent || "").replace(/\u00A0/g, " ").trim();
    if (!text || text.length > 200) continue;
    if (isSignOffLine(text)) return el;
  }
  return null;
}

async function extractInlineImages(bodyEl) {
  if (!bodyEl) return [];
  const boundary = findSignatureBoundaryElement(bodyEl);
  const imgs = Array.from(bodyEl.querySelectorAll("img")).filter((img) => {
    if (!boundary) return true;
    // compareDocumentPosition(img) describes img's position relative to
    // boundary. DOCUMENT_POSITION_PRECEDING (2) means img comes *before*
    // boundary in the document - i.e. before the signature - so keep it.
    // Anything else (following, or no relation) is at/after the signature
    // and gets excluded.
    const rel = boundary.compareDocumentPosition(img);
    return !!(rel & Node.DOCUMENT_POSITION_PRECEDING);
  });
  const results = [];
  let counter = 0;

  for (const img of imgs) {
    if (results.length >= 15) break; // sanity cap

    try {
      const w = img.naturalWidth || img.width || 0;
      const h = img.naturalHeight || img.height || 0;
      // Skip 1x1-ish tracking pixels / spacer gifs.
      if (w > 0 && h > 0 && w <= 3 && h <= 3) continue;

      const src = img.currentSrc || img.src;
      if (!src) continue;

      let dataUrl;
      if (src.startsWith("data:")) {
        dataUrl = src;
      } else {
        // Same-origin (outlook.*) image URLs carry the page's session
        // cookies automatically; cross-origin images (e.g. remote-hosted
        // marketing images) will fail here and are simply skipped.
        const resp = await fetch(src, { credentials: "include" });
        if (!resp.ok) continue;
        const blob = await resp.blob();
        if (blob.size < 200) continue; // likely a tracking pixel / spacer
        if (blob.size > 20 * 1024 * 1024) continue; // skip anything oversized
        dataUrl = await blobToDataUrl(blob);
      }

      counter++;
      const alt = (img.getAttribute("alt") || "").trim();
      const baseName = sanitizeFileName(alt || `inline-image-${counter}`);
      const name = baseName.includes(".") ? baseName : baseName + guessExtension(dataUrl);
      results.push({ name, dataUrl });
    } catch (e) {
      /* cross-origin or unreadable image - skip it, not fatal */
    }
  }

  return results;
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("could not read image blob"));
    reader.readAsDataURL(blob);
  });
}

function guessExtension(dataUrl) {
  const m = /^data:(image\/[a-zA-Z0-9.+-]+)/.exec(dataUrl || "");
  const map = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/webp": ".webp",
  };
  return (m && map[m[1]]) || ".png";
}

function sanitizeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, "_").slice(0, 150) || "inline-image";
}

function trimSignature(text) {
  if (!text) return text;
  const lines = text.split(/\r?\n/);

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    if (isSignOffLine(line)) {
      return lines.slice(0, i).join("\n").trimEnd();
    }
  }
  return text;
}

// Outlook Web frequently wraps each line of a message in its own <div>,
// and uses non-breaking spaces for indentation/padding. innerText on that
// structure often produces doubled-up blank lines between paragraphs and
// literal double spaces. Normalize both without touching real content.
function normalizeBodyText(text) {
  if (!text) return text;
  return text
    .replace(/\r\n?/g, "\n") // normalize line endings
    .replace(/\u00A0/g, " ") // non-breaking spaces -> regular spaces
    .split("\n")
    .map((line) => line.replace(/[ \t]+/g, " ").trim())
    .join("\n")
    .replace(/\n{3,}/g, "\n\n") // collapse 2+ blank lines down to a single blank line
    .trim();
}

async function extractEmailData() {
  const bodyEl = findMessageBodyElement();
  const inlineImages = await extractInlineImages(bodyEl);
  const bodyTextFull = normalizeBodyText(bodyEl ? bodyEl.innerText : "");
  const bodyTextTrimmed = trimSignature(bodyTextFull);
  return {
    subject: findSubject(),
    bodyHtml: bodyEl ? bodyEl.innerHTML : "",
    bodyText: bodyTextTrimmed,
    bodyTextFull,
    signatureTrimmed: bodyTextTrimmed !== bodyTextFull,
    attachmentNames: findAttachmentNames(),
    inlineImages,
    pageUrl: location.href,
  };
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg && msg.type === "GET_EMAIL_DATA") {
    extractEmailData()
      .then((data) => sendResponse({ ok: true, data }))
      .catch((e) => sendResponse({ ok: false, error: e.message }));
    return true; // keep the message channel open for the async response
  }
});
