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

function findAttachmentNames(scope) {
  // NOTE: not yet confirmed against a live email with real attachments -
  // this is a best-effort guess at likely markup patterns. If it comes up
  // empty on an email you know has attachments, that's expected for now;
  // the file picker in the popup still works regardless. Open dev tools on
  // such an email, inspect the attachment row, and share the markup to get
  // this tightened up.
  const pane = scope || findReadingPane();
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

/* =====================================================================
 * Per-message "Send to Azure DevOps" button
 *
 * Instead of always acting on whichever message happens to be topmost,
 * inject a small button next to each message's own "Reply" button so the
 * user can pick exactly which email in a conversation thread to send.
 * Clicking it opens an in-page panel (built in a shadow root so Outlook's
 * page styles can't bleed into it) scoped to that specific message.
 * ===================================================================== */

const WORK_ITEM_TYPES = ["Task", "Bug", "User Story", "Issue"];

function findMessageContainer(bodyEl) {
  if (!bodyEl) return findReadingPane();
  return bodyEl.closest('[aria-label="Email message" i]') || bodyEl;
}

function findMessageBodyElementNear(replyEl) {
  // Confirmed via live inspection: each open message is wrapped in a
  // div[aria-label="Email message"] that contains both that message's own
  // Message body and its own Reply/Reply all/Forward action row - so
  // climbing from the Reply node to that wrapper scopes us to the right
  // message even when a conversation thread has several.
  const container = replyEl.closest('[aria-label="Email message" i]');
  const body = container && container.querySelector('[role="document"][aria-label="Message body" i]');
  return body || findMessageBodyElement();
}

function findReplyButtons() {
  // Confirmed via live inspection: the per-message quick-action row (the
  // one with the reaction/theme icons next to Reply/Reply all/Forward,
  // distinct from the ribbon's Respond group) renders Reply as
  // role="menuitem" aria-label="Reply" - not a <button>. The "i" flag on
  // the attribute selector makes the match case-insensitive while still
  // requiring the whole value to be exactly "Reply" (so "Reply all" is
  // correctly excluded).
  //
  // A second, unrelated element also matches this selector: the mode
  // switcher (Reply / Reply all / Forward tabs) inside an open inline
  // reply-compose editor, which Outlook nests inside the same
  // div[aria-label="Email message"] wrapper as the message being replied
  // to - so that ancestor check alone doesn't filter it out. It renders as
  // an actual <button>, whereas the quick-action row's Reply is a <div>
  // carrying Fluent's overflow-tracking attribute (data-overflow-item);
  // requiring both distinguishes the two.
  return Array.from(document.querySelectorAll('[role="menuitem"][aria-label="Reply" i]')).filter(
    (el) => el.tagName === "DIV" && el.hasAttribute("data-overflow-item") && el.closest('[aria-label="Email message" i]')
  );
}

// A clipboard-with-checkmark glyph, in the same visual language as an
// "add work item" action. Drawn with currentColor so it follows whatever
// text color the surrounding quick-action row is using (light or dark theme).
const ADO_ICON_SVG = `
  <svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="3" y="2.5" width="10" height="12" rx="1.5" stroke="currentColor" stroke-width="1.3"/>
    <path d="M6 2.5V1.8a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1v0.7" stroke="currentColor" stroke-width="1.3"/>
    <path d="M5.7 8.7l1.4 1.4L10.4 6.7" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>
  </svg>
`;

function createAdoTriggerButton() {
  // Sized to match the sibling icon-only menu items in this same row (the
  // sun/theme and smiley/reaction toggles are 28x28, borderless, icon
  // only - confirmed via live inspection of their computed style).
  //
  // IMPORTANT: this button is never inserted into Outlook's own toolbar
  // DOM. That toolbar is a Fluent UI "Toolbar" with overflow behaviour -
  // it measures its own children's widths and moves whatever doesn't fit
  // into a "..." menu. Adding any extra child (even a small one) changes
  // that measurement and can bump Reply/Reply all/Forward themselves into
  // the overflow menu - confirmed live: a wider icon+text version of this
  // button pushed all three out of view. Instead, this button is
  // absolutely positioned on top of the page (see positionAdoButton) so
  // it never participates in Outlook's own layout at all.
  const btn = document.createElement("button");
  btn.type = "button";
  btn.title = "Create an Azure DevOps work item from this email";
  btn.setAttribute("aria-label", "Send this email to Azure DevOps");
  btn.innerHTML = ADO_ICON_SVG;
  Object.assign(btn.style, {
    position: "fixed",
    width: "28px",
    height: "28px",
    padding: "0",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: "4px",
    border: "none",
    background: "transparent",
    cursor: "pointer",
    zIndex: "2147483000",
  });
  btn.addEventListener("mouseenter", () => (btn.style.background = "rgba(128,128,128,0.25)"));
  btn.addEventListener("mouseleave", () => (btn.style.background = "transparent"));
  return btn;
}

// Maps a live Reply menuitem element to the floating button anchored to it.
const adoButtonsByReplyEl = new Map();

function positionAdoButton(replyEl, btn) {
  const rect = replyEl.getBoundingClientRect();
  const visible = rect.width > 0 && rect.height > 0 && !!replyEl.offsetParent;
  btn.style.display = visible ? "flex" : "none";
  if (!visible) return;
  // Anchored just to the right of Reply itself, matching its height, so it
  // reads as part of that action row without being structurally inside it.
  btn.style.top = `${Math.round(rect.top + (rect.height - 28) / 2)}px`;
  btn.style.left = `${Math.round(rect.right + 4)}px`;
  // Sampling Reply's own text color keeps the icon legible in both
  // Outlook's light and dark themes without relying on CSS inheritance,
  // which breaks once this button is moved out of Outlook's DOM subtree.
  btn.style.color = getComputedStyle(replyEl).color;
}

function injectAdoButtons() {
  const liveReplyEls = new Set(findReplyButtons());

  // Drop badges whose Reply element got removed/replaced (message closed,
  // navigated away from, etc).
  for (const [replyEl, btn] of adoButtonsByReplyEl) {
    if (!liveReplyEls.has(replyEl) || !document.contains(replyEl)) {
      btn.remove();
      adoButtonsByReplyEl.delete(replyEl);
    }
  }

  liveReplyEls.forEach((replyEl) => {
    let btn = adoButtonsByReplyEl.get(replyEl);
    if (!btn) {
      const bodyEl = findMessageBodyElementNear(replyEl);
      if (!bodyEl) return;
      btn = createAdoTriggerButton();
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        openAdoPanel(bodyEl);
      });
      document.body.appendChild(btn);
      adoButtonsByReplyEl.set(replyEl, btn);
    }
    positionAdoButton(replyEl, btn);
  });
}

function debounce(fn, delayMs) {
  let handle;
  return (...args) => {
    clearTimeout(handle);
    handle = setTimeout(() => fn(...args), delayMs);
  };
}

function repositionAllAdoButtons() {
  adoButtonsByReplyEl.forEach((btn, replyEl) => positionAdoButton(replyEl, btn));
}

function startAdoButtonInjection() {
  injectAdoButtons();
  const rerun = debounce(injectAdoButtons, 300);
  const observer = new MutationObserver(rerun);
  observer.observe(document.body, { childList: true, subtree: true });

  const reposition = debounce(repositionAllAdoButtons, 50);
  window.addEventListener("scroll", reposition, { capture: true, passive: true });
  window.addEventListener("resize", reposition, { passive: true });
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", startAdoButtonInjection);
} else {
  startAdoButtonInjection();
}

/* ---------------- In-page panel ---------------- */

function getStoredSettings() {
  return new Promise((resolve) => {
    chrome.storage.local.get(
      ["ado_org", "ado_project", "ado_pat", "ado_witype", "ado_default_areapath"],
      resolve
    );
  });
}

function sendToBackground(type, payload) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage({ type, payload }, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }
      if (!response || !response.ok) {
        reject(new Error((response && response.error) || "no response from extension"));
        return;
      }
      resolve(response.data);
    });
  });
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result).split(",")[1] || "");
    reader.onerror = () => reject(new Error(`could not read "${file.name}"`));
    reader.readAsDataURL(file);
  });
}

function dataUrlToBase64(dataUrl) {
  return (dataUrl || "").split(",")[1] || "";
}

const AREA_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 1 day

function getCache(key) {
  return new Promise((resolve) => chrome.storage.local.get([key], (r) => resolve(r[key] || null)));
}

function setCache(key, value) {
  return new Promise((resolve) => chrome.storage.local.set({ [key]: value }, resolve));
}

async function loadAreaPathsIntoSelect(select, hintEl, settings, force) {
  const { ado_org: org, ado_project: project, ado_pat: pat, ado_default_areapath: defaultAreaPath } = settings;
  const cacheKey = `areapaths_${org}_${project}`;

  const populate = (paths) => {
    select.innerHTML = '<option value="">(project default)</option>';
    paths.forEach(({ path, depth }) => {
      const opt = document.createElement("option");
      opt.value = path;
      opt.textContent = " ".repeat(depth) + path.split("\\").pop();
      select.appendChild(opt);
    });
    if (defaultAreaPath && paths.some((p) => p.path === defaultAreaPath)) select.value = defaultAreaPath;
  };

  if (!force) {
    const cached = await getCache(cacheKey);
    if (cached && Date.now() - cached.fetchedAt < AREA_CACHE_TTL_MS) {
      populate(cached.paths);
      hintEl.textContent = `${cached.paths.length} area paths (cached).`;
      return;
    }
  }

  hintEl.textContent = "Loading area paths…";
  try {
    const paths = await sendToBackground("ADO_FETCH_AREA_PATHS", { org, project, pat });
    await setCache(cacheKey, { paths, fetchedAt: Date.now() });
    populate(paths);
    hintEl.textContent = `${paths.length} area paths loaded.`;
  } catch (e) {
    hintEl.textContent = "Couldn't load area paths: " + e.message;
  }
}

const PANEL_STYLES = `
  :host { all: initial; }
  * { box-sizing: border-box; font-family: "Segoe UI", Tahoma, Arial, sans-serif; }
  .overlay {
    position: fixed; inset: 0; background: rgba(0,0,0,0.35);
    z-index: 2147483647; display: flex; align-items: flex-start;
    justify-content: center; padding-top: 6vh;
  }
  .panel {
    width: 420px; max-height: 88vh; overflow-y: auto; background: #fff;
    border-radius: 8px; box-shadow: 0 8px 30px rgba(0,0,0,0.3);
    padding: 16px; font-size: 13px; color: #242424;
  }
  .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
  .header h2 { font-size: 15px; margin: 0; }
  .close-btn { background: none; border: none; font-size: 18px; line-height: 1; cursor: pointer; color: #605e5c; }
  label { display: block; font-weight: 600; margin: 10px 0 4px; }
  input[type="text"], select, textarea {
    width: 100%; box-sizing: border-box; padding: 6px 8px; font-size: 13px;
    border: 1px solid #c8c6c4; border-radius: 4px; font-weight: normal;
  }
  input[type="file"] { width: 100%; font-weight: normal; font-size: 12px; }
  textarea { font-family: inherit; resize: vertical; }
  .checkbox-label { display: flex; align-items: center; gap: 6px; font-weight: normal; margin-top: 10px; }
  .checkbox-label input { margin: 0; }
  .hint { font-size: 11px; color: #605e5c; font-weight: normal; margin: 6px 0 0; }
  .actions { display: flex; gap: 8px; margin-top: 14px; }
  .primary-btn, .cancel-btn {
    border-radius: 4px; padding: 8px 14px; font-size: 13px; font-weight: 600; cursor: pointer; flex: 1;
  }
  .primary-btn { background: #0078d4; color: #fff; border: none; }
  .primary-btn:hover { background: #106ebe; }
  .primary-btn:disabled { background: #c8c6c4; cursor: not-allowed; }
  .cancel-btn { background: #fff; color: #242424; border: 1px solid #c8c6c4; }
  #status { margin-top: 10px; padding: 8px; border-radius: 4px; font-size: 12px; white-space: pre-wrap; }
  #status.success { background: #dff6dd; color: #107c10; }
  #status.error { background: #fde7e9; color: #a80000; }
  #status.info { background: #f3f2f1; color: #323130; }
`;

const PANEL_MARKUP = `
  <div class="overlay">
    <div class="panel">
      <div class="header">
        <h2>Send email to Azure DevOps</h2>
        <button class="close-btn" title="Close">×</button>
      </div>
      <label>Title<input type="text" class="title" /></label>
      <label>Work item type<select class="witype"></select></label>
      <label>Area path<select class="areapath"><option value="">(project default)</option></select></label>
      <p class="hint areapath-hint"></p>
      <label>Description<textarea class="description" rows="8"></textarea></label>
      <label class="checkbox-label"><input type="checkbox" class="trim-sig" checked /> Trim signature / sign-off from description</label>
      <p class="hint sig-hint"></p>
      <label>Attach files<input type="file" multiple class="file-input" /></label>
      <p class="hint attach-hint"></p>
      <label class="checkbox-label"><input type="checkbox" class="inline-images" checked /> Include inline images from email body</label>
      <p class="hint inline-hint"></p>
      <div id="status"></div>
      <div class="actions">
        <button class="cancel-btn">Cancel</button>
        <button class="primary-btn create-btn">Create work item</button>
      </div>
    </div>
  </div>
`;

let adoPanelHost = null;

function closeAdoPanel() {
  if (adoPanelHost) {
    adoPanelHost.remove();
    adoPanelHost = null;
  }
}

async function openAdoPanel(bodyEl) {
  closeAdoPanel();

  adoPanelHost = document.createElement("div");
  document.body.appendChild(adoPanelHost);
  const shadow = adoPanelHost.attachShadow({ mode: "open" });
  const style = document.createElement("style");
  style.textContent = PANEL_STYLES;
  shadow.appendChild(style);
  const wrapper = document.createElement("div");
  wrapper.innerHTML = PANEL_MARKUP;
  shadow.appendChild(wrapper);

  const q = (sel) => shadow.querySelector(sel);
  const els = {
    overlay: q(".overlay"),
    close: q(".close-btn"),
    cancel: q(".cancel-btn"),
    create: q(".create-btn"),
    title: q(".title"),
    witype: q(".witype"),
    areapath: q(".areapath"),
    areapathHint: q(".areapath-hint"),
    description: q(".description"),
    trimSig: q(".trim-sig"),
    sigHint: q(".sig-hint"),
    fileInput: q(".file-input"),
    attachHint: q(".attach-hint"),
    inlineImages: q(".inline-images"),
    inlineHint: q(".inline-hint"),
    status: q("#status"),
  };

  els.close.addEventListener("click", closeAdoPanel);
  els.cancel.addEventListener("click", closeAdoPanel);
  els.overlay.addEventListener("click", (e) => {
    if (e.target === els.overlay) closeAdoPanel();
  });

  WORK_ITEM_TYPES.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t;
    opt.textContent = t;
    els.witype.appendChild(opt);
  });

  const settings = await getStoredSettings();
  if (!settings.ado_org || !settings.ado_project || !settings.ado_pat) {
    els.status.textContent =
      'Set Organization, Project, and PAT first — open the extension icon\'s "Settings".';
    els.status.className = "error";
    els.create.disabled = true;
  } else {
    els.witype.value = settings.ado_witype || "Task";
    loadAreaPathsIntoSelect(els.areapath, els.areapathHint, settings, false);
  }

  const container = findMessageContainer(bodyEl);
  const inlineImages = await extractInlineImages(bodyEl);
  const bodyTextFull = normalizeBodyText(bodyEl ? bodyEl.innerText : "");
  const bodyTextTrimmed = trimSignature(bodyTextFull);

  els.title.value = findSubject();

  const applyTrim = () => {
    els.description.value = els.trimSig.checked ? bodyTextTrimmed : bodyTextFull;
  };
  applyTrim();
  els.trimSig.addEventListener("change", applyTrim);
  els.sigHint.textContent =
    bodyTextTrimmed !== bodyTextFull
      ? "Signature/sign-off detected and trimmed — toggle off to see the full text."
      : "No signature/sign-off detected to trim.";

  const attachmentNames = findAttachmentNames(container);
  els.attachHint.textContent = attachmentNames.length
    ? "Detected on this email: " + attachmentNames.join(", ") + " — use the file picker to actually attach them."
    : "";

  els.inlineHint.textContent = inlineImages.length
    ? `${inlineImages.length} inline image(s) found in this message (${inlineImages.map((i) => i.name).join(", ")}).`
    : "No inline images detected in this message.";

  // Using .onclick (not addEventListener) here because onCreateFromPanel
  // repoints this same handler to close the panel once the work item is
  // created - a second addEventListener would stack instead of replacing.
  els.create.onclick = () => onCreateFromPanel(els, inlineImages);
}

function panelShowStatus(els, msg, kind) {
  els.status.textContent = msg;
  els.status.className = kind || "info";
}

async function onCreateFromPanel(els, inlineImages) {
  const settings = await getStoredSettings();
  const { ado_org: org, ado_project: project, ado_pat: pat } = settings;
  if (!org || !project || !pat) {
    panelShowStatus(els, "Missing Organization/Project/PAT — configure them via the extension icon first.", "error");
    return;
  }

  const title = els.title.value.trim();
  if (!title) {
    panelShowStatus(els, "Title is required.", "error");
    return;
  }

  const witype = els.witype.value;
  const areaPath = els.areapath.value;
  const description = els.description.value;
  const files = els.fileInput.files;

  els.create.disabled = true;
  try {
    panelShowStatus(els, "Creating work item…", "info");
    const workItem = await sendToBackground("ADO_CREATE_WORK_ITEM", {
      org,
      project,
      pat,
      witype,
      title,
      description,
      areaPath,
    });

    for (let i = 0; i < files.length; i++) {
      panelShowStatus(els, `Work item #${workItem.id} created.\nUploading ${i + 1}/${files.length}: ${files[i].name}…`, "info");
      try {
        const bytesBase64 = await fileToBase64(files[i]);
        await sendToBackground("ADO_UPLOAD_ATTACHMENT", {
          org,
          project,
          pat,
          workItemId: workItem.id,
          fileName: files[i].name,
          bytesBase64,
        });
      } catch (e) {
        panelShowStatus(els, `Work item #${workItem.id} created, but "${files[i].name}" failed to attach: ${e.message}`, "error");
      }
    }

    if (els.inlineImages.checked && inlineImages.length > 0) {
      for (let i = 0; i < inlineImages.length; i++) {
        const img = inlineImages[i];
        panelShowStatus(
          els,
          `Work item #${workItem.id} created.\nUploading inline image ${i + 1}/${inlineImages.length}: ${img.name}…`,
          "info"
        );
        try {
          await sendToBackground("ADO_UPLOAD_ATTACHMENT", {
            org,
            project,
            pat,
            workItemId: workItem.id,
            fileName: img.name,
            bytesBase64: dataUrlToBase64(img.dataUrl),
          });
        } catch (e) {
          panelShowStatus(els, `Work item #${workItem.id} created, but inline image "${img.name}" failed to attach: ${e.message}`, "error");
        }
      }
    }

    const orgUrl = org.startsWith("http") ? org.replace(/\/$/, "") : `https://dev.azure.com/${org}`;
    const link = `${orgUrl}/${encodeURIComponent(project)}/_workitems/edit/${workItem.id}`;
    panelShowStatus(els, `Done. Work item #${workItem.id} created.\n${link}`, "success");

    els.create.textContent = "Close";
    els.create.disabled = false;
    els.create.onclick = closeAdoPanel;
    els.cancel.textContent = "Open work item";
    els.cancel.onclick = () => window.open(link, "_blank", "noopener");
  } catch (err) {
    panelShowStatus(els, "Failed: " + err.message, "error");
    els.create.disabled = false;
  }
}
