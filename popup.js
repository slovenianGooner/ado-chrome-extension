const WORK_ITEM_TYPES = ["Task", "Bug", "User Story", "Issue"];
const API_VERSION = "7.1";

let detectedInlineImages = [];
let lastBodyTextFull = "";
let lastBodyTextTrimmed = "";

document.addEventListener("DOMContentLoaded", async () => {
  populateTypeSelects();
  await loadSettings();
  wireButtons();
  await scanActiveEmail();
  await loadAreaPaths(false, true);
});

function wireButtons() {
  document.getElementById("toggle-settings").addEventListener("click", () => {
    document.getElementById("settings").classList.toggle("hidden");
  });
  document.getElementById("save-settings").addEventListener("click", onSaveSettingsClicked);
  document.getElementById("create-btn").addEventListener("click", onCreateClicked);
  document.getElementById("rescan-btn").addEventListener("click", scanActiveEmail);
  document.getElementById("refresh-areas-btn").addEventListener("click", () => loadAreaPaths(true));
  document.getElementById("trim-signature").addEventListener("change", applySignatureTrimPreference);
}

function applySignatureTrimPreference() {
  const trim = document.getElementById("trim-signature").checked;
  document.getElementById("description").value = trim ? lastBodyTextTrimmed : lastBodyTextFull;
}

function populateTypeSelects() {
  const a = document.getElementById("witype-setting");
  const b = document.getElementById("witype");
  [a, b].forEach((sel) => {
    sel.innerHTML = "";
    WORK_ITEM_TYPES.forEach((t) => {
      const opt = document.createElement("option");
      opt.value = t;
      opt.textContent = t;
      sel.appendChild(opt);
    });
  });
}

/* ---------------- Settings ---------------- */

function loadSettings() {
  return new Promise((resolve) => {
    chrome.storage.local.get(["ado_org", "ado_project", "ado_pat", "ado_witype"], (r) => {
      document.getElementById("org").value = r.ado_org || "";
      document.getElementById("project").value = r.ado_project || "";
      document.getElementById("pat").value = r.ado_pat || "";
      const t = r.ado_witype || "Task";
      document.getElementById("witype-setting").value = t;
      document.getElementById("witype").value = t;
      if (!r.ado_org || !r.ado_project || !r.ado_pat) {
        document.getElementById("settings").classList.remove("hidden");
      }
      resolve();
    });
  });
}

function onSaveSettingsClicked() {
  saveSettings().then(() => loadAreaPaths(true, true));
}

function saveSettings() {
  return new Promise((resolve) => {
    const settings = {
      ado_org: document.getElementById("org").value.trim(),
      ado_project: document.getElementById("project").value.trim(),
      ado_pat: document.getElementById("pat").value.trim(),
      ado_witype: document.getElementById("witype-setting").value,
      ado_default_areapath: document.getElementById("default-area-path").value,
    };
    chrome.storage.local.set(settings, () => {
      document.getElementById("witype").value = settings.ado_witype;
      showStatus("Settings saved.", "success");
      document.getElementById("settings").classList.add("hidden");
      resolve();
    });
  });
}

function getSettings() {
  return new Promise((resolve) => {
    chrome.storage.local.get(["ado_org", "ado_project", "ado_pat", "ado_default_areapath"], resolve);
  });
}

/* ---------------- Area paths ---------------- */

const AREA_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 1 day

async function loadAreaPaths(force, seedDefault) {
  const select = document.getElementById("area-path");
  const defaultSelect = document.getElementById("default-area-path");
  const hint = document.getElementById("area-path-hint");
  const {
    ado_org: org,
    ado_project: project,
    ado_pat: pat,
    ado_default_areapath: defaultAreaPath,
  } = await getSettings();

  if (!org || !project || !pat) {
    hint.textContent = "Set Organization, Project, and PAT in Settings to load area paths.";
    return;
  }

  const cacheKey = `areapaths_${org}_${project}`;

  if (!force) {
    const cached = await getCache(cacheKey);
    if (cached && Date.now() - cached.fetchedAt < AREA_CACHE_TTL_MS) {
      populateAreaSelect(select, cached.paths, seedDefault ? defaultAreaPath : null);
      populateAreaSelect(defaultSelect, cached.paths, defaultAreaPath);
      hint.textContent = `${cached.paths.length} area paths (cached).`;
      return;
    }
  }

  hint.textContent = "Loading area paths…";
  try {
    const paths = await fetchAreaPaths({ org, project, pat });
    await setCache(cacheKey, { paths, fetchedAt: Date.now() });
    populateAreaSelect(select, paths, seedDefault ? defaultAreaPath : null);
    populateAreaSelect(defaultSelect, paths, defaultAreaPath);
    hint.textContent = `${paths.length} area paths loaded.`;
  } catch (e) {
    hint.textContent = "Couldn't load area paths: " + e.message;
  }
}

function getCache(key) {
  return new Promise((resolve) => {
    chrome.storage.local.get([key], (r) => resolve(r[key] || null));
  });
}

function setCache(key, value) {
  return new Promise((resolve) => {
    chrome.storage.local.set({ [key]: value }, resolve);
  });
}

async function fetchAreaPaths({ org, project, pat }) {
  const url = `${buildOrgUrl(org)}/${encodeURIComponent(
    project
  )}/_apis/wit/classificationnodes/Areas?$depth=15&api-version=${API_VERSION}`;

  const resp = await fetch(url, {
    headers: { Authorization: authHeader(pat) },
  });
  if (!resp.ok) throw new Error(`ADO API ${resp.status}: ${await safeText(resp)}`);
  const root = await resp.json();

  const results = [];
  function walk(node, parentPath, depth) {
    const fullPath = parentPath ? `${parentPath}\\${node.name}` : node.name;
    results.push({ path: fullPath, depth });
    (node.children || []).forEach((child) => walk(child, fullPath, depth + 1));
  }
  walk(root, "", 0);
  return results;
}

function populateAreaSelect(select, paths, seedValue) {
  const previousValue = select.value;
  select.innerHTML = '<option value="">(project default)</option>';
  paths.forEach(({ path, depth }) => {
    const opt = document.createElement("option");
    opt.value = path;
    opt.textContent = "\u2003".repeat(depth) + path.split("\\").pop();
    select.appendChild(opt);
  });
  if (previousValue && paths.some((p) => p.path === previousValue)) {
    select.value = previousValue;
  } else if (seedValue && paths.some((p) => p.path === seedValue)) {
    select.value = seedValue;
  }
}

/* ---------------- Scan the open email ---------------- */

async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

function isOutlookUrl(url) {
  return /^https:\/\/(outlook\.office\.com|outlook\.office365\.com|outlook\.live\.com|outlook\.cloud\.microsoft)\//.test(
    url || ""
  );
}

function sendMessageToTab(tabId, message) {
  return new Promise((resolve) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      resolve({ response, error: chrome.runtime.lastError });
    });
  });
}

async function scanActiveEmail() {
  const tab = await getActiveTab();
  if (!tab || !isOutlookUrl(tab.url)) {
    document.getElementById("not-outlook").classList.remove("hidden");
    document.getElementById("form").classList.remove("hidden");
    return;
  }
  document.getElementById("not-outlook").classList.add("hidden");

  let { response, error } = await sendMessageToTab(tab.id, { type: "GET_EMAIL_DATA" });

  // Most common cause of failure here: this tab was already open before the
  // extension was installed/reloaded, so Chrome never injected content.js
  // into it (that only happens automatically on navigation). Inject it now
  // and retry once, rather than requiring the user to refresh the tab.
  if (error || !response || !response.ok) {
    try {
      await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ["content.js"] });
      ({ response, error } = await sendMessageToTab(tab.id, { type: "GET_EMAIL_DATA" }));
    } catch (e) {
      showStatus(
        "Couldn't access this tab to read the email (" + e.message + "). Fill in Title/Description manually.",
        "info"
      );
      return;
    }
  }

  if (error || !response || !response.ok) {
    const detail = error?.message || response?.error || "no response from page";
    showStatus(
      `Couldn't auto-read the open email (${detail}). Fill in Title/Description manually.`,
      "info"
    );
    return;
  }

  const data = response.data;
  if (data.subject) document.getElementById("title").value = data.subject;

  lastBodyTextFull = data.bodyTextFull || data.bodyText || "";
  lastBodyTextTrimmed = data.bodyText || "";
  applySignatureTrimPreference();

  const sigHint = document.getElementById("signature-hint");
  sigHint.textContent = data.signatureTrimmed
    ? "Looks like a sign-off/signature was found and trimmed — toggle off to see the full text."
    : "No signature/sign-off detected to trim.";

  const detected = document.getElementById("detected-attachments");
  if (data.attachmentNames && data.attachmentNames.length) {
    detected.textContent =
      "Detected on this email: " +
      data.attachmentNames.join(", ") +
      " — use the file picker above to actually attach them.";
  } else {
    detected.textContent = "";
  }

  detectedInlineImages = data.inlineImages || [];
  const inlineHint = document.getElementById("inline-images-hint");
  inlineHint.textContent = detectedInlineImages.length
    ? `${detectedInlineImages.length} inline image(s) found in the body (${detectedInlineImages
        .map((i) => i.name)
        .join(", ")}).`
    : "No inline images detected in the body.";
}

/* ---------------- Create work item ---------------- */

async function onCreateClicked() {
  const { ado_org: org, ado_project: project, ado_pat: pat } = await getSettings();
  if (!org || !project || !pat) {
    showStatus("Fill in Organization, Project, and PAT under Settings first.", "error");
    document.getElementById("settings").classList.remove("hidden");
    return;
  }

  const title = document.getElementById("title").value.trim();
  const description = document.getElementById("description").value;
  const witype = document.getElementById("witype").value;
  const areaPath = document.getElementById("area-path").value;
  const files = document.getElementById("file-input").files;

  if (!title) {
    showStatus("Title is required.", "error");
    return;
  }

  const createBtn = document.getElementById("create-btn");
  createBtn.disabled = true;

  try {
    showStatus("Creating work item…", "info");
    const workItem = await createWorkItem({ org, project, pat, witype, title, description, areaPath });

    if (files.length > 0) {
      for (let i = 0; i < files.length; i++) {
        showStatus(`Work item #${workItem.id} created.\nUploading ${i + 1}/${files.length}: ${files[i].name}…`, "info");
        try {
          const bytes = await files[i].arrayBuffer();
          await uploadAttachment({ org, project, pat, workItemId: workItem.id, fileName: files[i].name, bytes });
        } catch (e) {
          showStatus(`Work item #${workItem.id} created, but "${files[i].name}" failed to attach: ${e.message}`, "error");
        }
      }
    }

    const includeInline = document.getElementById("include-inline-images").checked;
    if (includeInline && detectedInlineImages.length > 0) {
      for (let i = 0; i < detectedInlineImages.length; i++) {
        const img = detectedInlineImages[i];
        showStatus(
          `Work item #${workItem.id} created.\nUploading inline image ${i + 1}/${detectedInlineImages.length}: ${img.name}…`,
          "info"
        );
        try {
          const bytes = dataUrlToUint8Array(img.dataUrl);
          await uploadAttachment({ org, project, pat, workItemId: workItem.id, fileName: img.name, bytes });
        } catch (e) {
          showStatus(`Work item #${workItem.id} created, but inline image "${img.name}" failed to attach: ${e.message}`, "error");
        }
      }
    }

    const orgUrl = buildOrgUrl(org);
    const link = `${orgUrl}/${encodeURIComponent(project)}/_workitems/edit/${workItem.id}`;
    showStatus(`Done. Work item #${workItem.id} created.\n${link}`, "success");
  } catch (err) {
    showStatus("Failed: " + err.message, "error");
  } finally {
    createBtn.disabled = false;
  }
}

function buildOrgUrl(org) {
  if (org.startsWith("http")) return org.replace(/\/$/, "");
  return `https://dev.azure.com/${org}`;
}

function authHeader(pat) {
  return "Basic " + btoa(":" + pat);
}

// Azure DevOps' System.Description field is HTML. Plain newlines are
// invisible to an HTML renderer, so without this the work item shows all
// paragraphs run together even though the popup's textarea looked fine.
function textToHtml(text) {
  const escaped = (text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return escaped.split("\n").join("<br>");
}

async function createWorkItem({ org, project, pat, witype, title, description, areaPath }) {
  const url = `${buildOrgUrl(org)}/${encodeURIComponent(project)}/_apis/wit/workitems/$${encodeURIComponent(
    witype
  )}?api-version=${API_VERSION}`;

  const patchDoc = [
    { op: "add", path: "/fields/System.Title", value: title },
    { op: "add", path: "/fields/System.Description", value: textToHtml(description) },
  ];
  if (areaPath) {
    patchDoc.push({ op: "add", path: "/fields/System.AreaPath", value: areaPath });
  }

  const resp = await fetch(url, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json-patch+json",
      Authorization: authHeader(pat),
    },
    body: JSON.stringify(patchDoc),
  });

  if (!resp.ok) throw new Error(`ADO API ${resp.status}: ${await safeText(resp)}`);
  return resp.json();
}

async function uploadAttachment({ org, project, pat, workItemId, fileName, bytes }) {
  const uploadUrl = `${buildOrgUrl(org)}/${encodeURIComponent(project)}/_apis/wit/attachments?fileName=${encodeURIComponent(
    fileName
  )}&api-version=${API_VERSION}`;

  const uploadResp = await fetch(uploadUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/octet-stream",
      Authorization: authHeader(pat),
    },
    body: bytes,
  });
  if (!uploadResp.ok) throw new Error(`upload failed (${uploadResp.status}): ${await safeText(uploadResp)}`);
  const uploaded = await uploadResp.json();

  const patchUrl = `${buildOrgUrl(org)}/${encodeURIComponent(project)}/_apis/wit/workitems/${workItemId}?api-version=${API_VERSION}`;
  const patchDoc = [
    {
      op: "add",
      path: "/relations/-",
      value: {
        rel: "AttachedFile",
        url: uploaded.url,
        attributes: { comment: "Added from Outlook email" },
      },
    },
  ];
  const linkResp = await fetch(patchUrl, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json-patch+json",
      Authorization: authHeader(pat),
    },
    body: JSON.stringify(patchDoc),
  });
  if (!linkResp.ok) throw new Error(`link failed (${linkResp.status}): ${await safeText(linkResp)}`);
}

function dataUrlToUint8Array(dataUrl) {
  const base64 = dataUrl.split(",")[1] || "";
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

async function safeText(resp) {
  try {
    return await resp.text();
  } catch {
    return "(no response body)";
  }
}

function showStatus(msg, kind) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = kind || "info";
}
