const WORK_ITEM_TYPES = ["Task", "Bug", "User Story", "Issue"];
const AREA_CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 1 day

document.addEventListener("DOMContentLoaded", async () => {
  populateTypeSelect();
  await loadSettings();
  wireButtons();
  await loadAreaPaths(false);
});

function wireButtons() {
  document.getElementById("save-settings").addEventListener("click", onSaveSettingsClicked);
  document.getElementById("refresh-areas-btn").addEventListener("click", () => loadAreaPaths(true));
}

function populateTypeSelect() {
  const select = document.getElementById("witype-setting");
  select.innerHTML = "";
  WORK_ITEM_TYPES.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t;
    opt.textContent = t;
    select.appendChild(opt);
  });
}

/* ---------------- Settings ---------------- */

function loadSettings() {
  return new Promise((resolve) => {
    chrome.storage.local.get(["ado_org", "ado_project", "ado_pat", "ado_witype"], (r) => {
      document.getElementById("org").value = r.ado_org || "";
      document.getElementById("project").value = r.ado_project || "";
      document.getElementById("pat").value = r.ado_pat || "";
      document.getElementById("witype-setting").value = r.ado_witype || "Task";
      resolve();
    });
  });
}

function onSaveSettingsClicked() {
  saveSettings().then(() => loadAreaPaths(true));
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
      showStatus("Settings saved.", "success");
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

async function loadAreaPaths(force) {
  const select = document.getElementById("default-area-path");
  const hint = document.getElementById("area-path-hint");
  const { ado_org: org, ado_project: project, ado_pat: pat, ado_default_areapath: defaultAreaPath } = await getSettings();

  if (!org || !project || !pat) {
    hint.textContent = "Set Organization, Project, and PAT to load area paths.";
    return;
  }

  const cacheKey = `areapaths_${org}_${project}`;

  if (!force) {
    const cached = await getCache(cacheKey);
    if (cached && Date.now() - cached.fetchedAt < AREA_CACHE_TTL_MS) {
      populateAreaSelect(select, cached.paths, defaultAreaPath);
      hint.textContent = `${cached.paths.length} area paths (cached).`;
      return;
    }
  }

  hint.textContent = "Loading area paths…";
  try {
    const paths = await sendToBackground("ADO_FETCH_AREA_PATHS", { org, project, pat });
    await setCache(cacheKey, { paths, fetchedAt: Date.now() });
    populateAreaSelect(select, paths, defaultAreaPath);
    hint.textContent = `${paths.length} area paths loaded.`;
  } catch (e) {
    hint.textContent = "Couldn't load area paths: " + e.message;
  }
}

function getCache(key) {
  return new Promise((resolve) => chrome.storage.local.get([key], (r) => resolve(r[key] || null)));
}

function setCache(key, value) {
  return new Promise((resolve) => chrome.storage.local.set({ [key]: value }, resolve));
}

function populateAreaSelect(select, paths, seedValue) {
  select.innerHTML = '<option value="">(project default)</option>';
  paths.forEach(({ path, depth }) => {
    const opt = document.createElement("option");
    opt.value = path;
    opt.textContent = " ".repeat(depth) + path.split("\\").pop();
    select.appendChild(opt);
  });
  if (seedValue && paths.some((p) => p.path === seedValue)) {
    select.value = seedValue;
  }
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

function showStatus(msg, kind) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = kind || "info";
}
