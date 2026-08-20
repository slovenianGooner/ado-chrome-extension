/* Service worker. Content scripts run inside outlook.office.com's page
 * context, where fetches to dev.azure.com can be blocked by Outlook's own
 * Content-Security-Policy. The extension's background context isn't subject
 * to that page CSP, so the in-page "send this email" panel routes all Azure
 * DevOps API calls through here via chrome.runtime.sendMessage.
 */

const API_VERSION = "7.1";

function buildOrgUrl(org) {
  if (org.startsWith("http")) return org.replace(/\/$/, "");
  return `https://dev.azure.com/${org}`;
}

function authHeader(pat) {
  return "Basic " + btoa(":" + pat);
}

async function safeText(resp) {
  try {
    return await resp.text();
  } catch {
    return "(no response body)";
  }
}

// Azure DevOps' System.Description field is HTML; plain newlines are
// invisible to an HTML renderer.
function textToHtml(text) {
  const escaped = (text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  return escaped.split("\n").join("<br>");
}

async function fetchAreaPaths({ org, project, pat }) {
  const url = `${buildOrgUrl(org)}/${encodeURIComponent(
    project
  )}/_apis/wit/classificationnodes/Areas?$depth=15&api-version=${API_VERSION}`;

  const resp = await fetch(url, { headers: { Authorization: authHeader(pat) } });
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

async function createWorkItem({ org, project, pat, witype, title, description, areaPath }) {
  const url = `${buildOrgUrl(org)}/${encodeURIComponent(project)}/_apis/wit/workitems/$${encodeURIComponent(
    witype
  )}?api-version=${API_VERSION}`;

  const patchDoc = [
    { op: "add", path: "/fields/System.Title", value: title },
    { op: "add", path: "/fields/System.Description", value: textToHtml(description) },
  ];
  if (areaPath) patchDoc.push({ op: "add", path: "/fields/System.AreaPath", value: areaPath });

  const resp = await fetch(url, {
    method: "PATCH",
    headers: { "Content-Type": "application/json-patch+json", Authorization: authHeader(pat) },
    body: JSON.stringify(patchDoc),
  });
  if (!resp.ok) throw new Error(`ADO API ${resp.status}: ${await safeText(resp)}`);
  return resp.json();
}

function base64ToUint8Array(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

async function uploadAttachment({ org, project, pat, workItemId, fileName, bytesBase64 }) {
  const bytes = base64ToUint8Array(bytesBase64);
  const uploadUrl = `${buildOrgUrl(org)}/${encodeURIComponent(project)}/_apis/wit/attachments?fileName=${encodeURIComponent(
    fileName
  )}&api-version=${API_VERSION}`;

  const uploadResp = await fetch(uploadUrl, {
    method: "POST",
    headers: { "Content-Type": "application/octet-stream", Authorization: authHeader(pat) },
    body: bytes,
  });
  if (!uploadResp.ok) throw new Error(`upload failed (${uploadResp.status}): ${await safeText(uploadResp)}`);
  const uploaded = await uploadResp.json();

  const patchUrl = `${buildOrgUrl(org)}/${encodeURIComponent(project)}/_apis/wit/workitems/${workItemId}?api-version=${API_VERSION}`;
  const patchDoc = [
    {
      op: "add",
      path: "/relations/-",
      value: { rel: "AttachedFile", url: uploaded.url, attributes: { comment: "Added from Outlook email" } },
    },
  ];
  const linkResp = await fetch(patchUrl, {
    method: "PATCH",
    headers: { "Content-Type": "application/json-patch+json", Authorization: authHeader(pat) },
    body: JSON.stringify(patchDoc),
  });
  if (!linkResp.ok) throw new Error(`link failed (${linkResp.status}): ${await safeText(linkResp)}`);
}

const HANDLERS = {
  ADO_FETCH_AREA_PATHS: fetchAreaPaths,
  ADO_CREATE_WORK_ITEM: createWorkItem,
  ADO_UPLOAD_ATTACHMENT: uploadAttachment,
};

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const handler = msg && HANDLERS[msg.type];
  if (!handler) return false;
  handler(msg.payload || {})
    .then((data) => sendResponse({ ok: true, data }))
    .catch((e) => sendResponse({ ok: false, error: e.message }));
  return true; // async response
});
