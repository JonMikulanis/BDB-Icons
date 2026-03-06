// ── Config ────────────────────────────────────────────────────────────────────
const GITHUB_API     = "https://api.github.com";
const DEFAULT_OWNER  = "JonMikulanis";
const DEFAULT_REPO   = "BDB-Icons";
const DEFAULT_BRANCH = "main";

// ── State ─────────────────────────────────────────────────────────────────────
let assets         = [];   // loaded from library.json
let filterText     = "";
let activeCategory = "All";
let cfg            = {};   // { token, owner, repo, branch }

// ── Office Init ───────────────────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    initSettings();
    loadLibrary();
  } else {
    toast("Open this add-in inside PowerPoint.", "error");
  }
});

// ── Settings ──────────────────────────────────────────────────────────────────
function initSettings() {
  cfg = {
    token:  localStorage.getItem("gh_token")  || "",
    owner:  localStorage.getItem("gh_owner")  || DEFAULT_OWNER,
    repo:   localStorage.getItem("gh_repo")   || DEFAULT_REPO,
    branch: localStorage.getItem("gh_branch") || DEFAULT_BRANCH,
  };
  document.getElementById("set-token").value  = cfg.token;
  document.getElementById("set-owner").value  = cfg.owner;
  document.getElementById("set-repo").value   = cfg.repo;
  document.getElementById("set-branch").value = cfg.branch;
}

function saveSettings() {
  cfg.token  = document.getElementById("set-token").value.trim();
  cfg.owner  = document.getElementById("set-owner").value.trim()  || DEFAULT_OWNER;
  cfg.repo   = document.getElementById("set-repo").value.trim()   || DEFAULT_REPO;
  cfg.branch = document.getElementById("set-branch").value.trim() || DEFAULT_BRANCH;
  localStorage.setItem("gh_token",  cfg.token);
  localStorage.setItem("gh_owner",  cfg.owner);
  localStorage.setItem("gh_repo",   cfg.repo);
  localStorage.setItem("gh_branch", cfg.branch);
  closePanel("settings-panel");
  toast("Settings saved.", "success");
  loadLibrary();
}

function pagesBase() {
  return "https://" + cfg.owner + ".github.io/" + cfg.repo + "/";
}

function apiBase() {
  return GITHUB_API + "/repos/" + cfg.owner + "/" + cfg.repo;
}

// ── Load library from GitHub Pages ───────────────────────────────────────────
async function loadLibrary() {
  setBusy(true);
  try {
    // Cache-bust so GitHub Pages CDN doesn't serve stale JSON
    const url = pagesBase() + "library.json?_=" + Date.now();
    const res = await fetch(url);
    if (!res.ok) throw new Error("HTTP " + res.status);
    assets = await res.json();
    renderCategoryChips();
    renderGrid();
  } catch (e) {
    toast("Could not load library. Check Settings or try Refresh.", "error");
    assets = [];
    renderCategoryChips();
    renderGrid();
  } finally {
    setBusy(false);
  }
}

function setBusy(on) {
  document.getElementById("refresh-btn").classList.toggle("spinning", on);
}

// ── Insert asset onto active slide ────────────────────────────────────────────
async function insertAsset(idx) {
  const asset = assets[idx];
  if (!asset) return;

  const url = pagesBase() + "assets/" + asset.file;

  try {
    if (asset.type === "svg") {
      const res = await fetch(url);
      if (!res.ok) throw new Error("Could not fetch SVG (HTTP " + res.status + ")");
      const svgText = await res.text();
      Office.context.document.setSelectedDataAsync(
        svgText,
        { coercionType: Office.CoercionType.XmlSvg },
        function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            toast("Inserted: " + asset.name, "success");
          } else {
            toast("Insert failed — SVG may be malformed.", "error");
          }
        }
      );
    } else {
      // PNG / JPG — fetch as blob, convert to base64, then insert
      const res = await fetch(url);
      if (!res.ok) throw new Error("Could not fetch image (HTTP " + res.status + ")");
      const blob = await res.blob();
      const b64  = await blobToBase64(blob);
      Office.context.document.setSelectedDataAsync(
        b64,
        { coercionType: Office.CoercionType.Image },
        function (r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            toast("Inserted: " + asset.name, "success");
          } else {
            toast("Insert failed — check image format.", "error");
          }
        }
      );
    }
  } catch (e) {
    toast("Could not load asset: " + (e.message || "unknown error"), "error");
  }
}

function blobToBase64(blob) {
  return new Promise(function (resolve, reject) {
    const reader = new FileReader();
    reader.onload  = function () { resolve(reader.result.split(",")[1]); };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// ── Upload flow (GitHub API) ──────────────────────────────────────────────────
function handleUpload(input) {
  const files = Array.from(input.files);
  input.value = "";
  if (!files.length) return;

  if (!cfg.token) {
    toast("Add a GitHub token in ⚙ Settings first.", "error");
    return;
  }

  showUploadModal(files[0]);
}

function showUploadModal(file) {
  const ext      = file.name.split(".").pop().toLowerCase();
  const baseName = file.name.replace(/\.[^.]+$/, "").replace(/[-_]/g, " ");
  const safeName = file.name.toLowerCase().replace(/\s+/g, "-");

  document.getElementById("up-display-name").value  = toTitleCase(baseName);
  document.getElementById("up-filename").value       = safeName;
  document.getElementById("up-filetype").textContent = ext.toUpperCase();
  document.getElementById("up-category").value       = "";

  // Populate existing categories in the datalist
  const cats = [...new Set(assets.map(function (a) { return a.category; }).filter(Boolean))];
  const dl   = document.getElementById("up-cat-list");
  dl.innerHTML = cats.map(function (c) {
    return "<option value=\"" + xmlEsc(c) + "\">";
  }).join("");

  window._pendingFile = file;
  openPanel("upload-modal");
}

async function confirmUpload() {
  const file = window._pendingFile;
  if (!file) return;

  const displayName = document.getElementById("up-display-name").value.trim();
  const filename    = document.getElementById("up-filename").value.trim().toLowerCase().replace(/\s+/g, "-");
  const category    = document.getElementById("up-category").value.trim() || "General";

  if (!displayName) { toast("Display name is required.", "error"); return; }
  if (!filename)    { toast("Filename is required.", "error");     return; }

  const ext  = filename.split(".").pop().toLowerCase();
  const type = (ext === "svg") ? "svg" : "image";

  closePanel("upload-modal");
  setBusy(true);
  toast("Uploading to GitHub…", "info");

  try {
    // 1. Read file as base64
    const fileB64 = await fileToBase64(file);

    // 2. Upload the asset file to the repo
    await ghPut("contents/assets/" + filename, {
      message: "Add asset: " + displayName,
      content: fileB64,
      branch:  cfg.branch,
    });

    // 3. Fetch current library.json from GitHub API to get SHA (needed for update)
    let currentLib = [];
    let libSha     = null;
    try {
      const libData = await ghGet("contents/library.json");
      libSha      = libData.sha;
      // GitHub API returns base64 content with line-breaks — strip them before decoding
      currentLib  = JSON.parse(atob(libData.content.replace(/\s/g, "")));
    } catch (e) {
      // library.json doesn't exist yet — will be created fresh
    }

    // 4. Append the new entry
    currentLib.push({ name: displayName, file: filename, type: type, category: category });

    // 5. Commit the updated library.json
    const libContent = btoa(unescape(encodeURIComponent(JSON.stringify(currentLib, null, 2))));
    const putBody    = {
      message: "Update library.json: add " + displayName,
      content: libContent,
      branch:  cfg.branch,
    };
    if (libSha) putBody.sha = libSha;
    await ghPut("contents/library.json", putBody);

    toast("Uploaded! Library will refresh in ~4 s…", "success");
    // GitHub Pages CDN takes a few seconds to propagate the commit
    setTimeout(loadLibrary, 4000);

  } catch (e) {
    toast("Upload failed: " + (e.message || "unknown error"), "error");
  } finally {
    setBusy(false);
    window._pendingFile = null;
  }
}

// ── GitHub API helpers ────────────────────────────────────────────────────────
async function ghGet(path) {
  const res = await fetch(apiBase() + "/" + path, {
    headers: {
      Authorization: "token " + cfg.token,
      Accept:        "application/vnd.github+json",
    }
  });
  if (!res.ok) {
    const err = await res.json().catch(function () { return {}; });
    throw new Error(err.message || "HTTP " + res.status);
  }
  return res.json();
}

async function ghPut(path, body) {
  const res = await fetch(apiBase() + "/" + path, {
    method:  "PUT",
    headers: {
      Authorization:  "token " + cfg.token,
      Accept:         "application/vnd.github+json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.json().catch(function () { return {}; });
    throw new Error(err.message || "HTTP " + res.status);
  }
  return res.json();
}

function fileToBase64(file) {
  return new Promise(function (resolve, reject) {
    const reader = new FileReader();
    reader.onload  = function () { resolve(reader.result.split(",")[1]); };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── Category chips ────────────────────────────────────────────────────────────
function renderCategoryChips() {
  const raw  = assets.map(function (a) { return a.category; }).filter(Boolean);
  const cats = ["All"].concat([...new Set(raw)].sort());

  // If active category was removed from the library, reset to All
  if (!cats.includes(activeCategory)) activeCategory = "All";

  document.getElementById("category-chips").innerHTML = cats.map(function (c) {
    return (
      "<button class=\"chip" + (c === activeCategory ? " active" : "") +
      "\" onclick=\"setCategory('" + c.replace(/'/g, "\\'") + "')\">" +
      xmlEsc(c) + "</button>"
    );
  }).join("");
}

function setCategory(cat) {
  activeCategory = cat;
  renderCategoryChips();
  renderGrid();
}

// ── Filter / search ────────────────────────────────────────────────────────────
function filterAssets(value) {
  filterText = value.toLowerCase().trim();
  renderGrid();
}

// ── Render grid ────────────────────────────────────────────────────────────────
function renderGrid() {
  const grid    = document.getElementById("icon-grid");
  const empty   = document.getElementById("empty-state");
  const counter = document.getElementById("icon-count");

  counter.textContent = assets.length;

  let visible = assets.slice();
  if (activeCategory !== "All") {
    visible = visible.filter(function (a) { return a.category === activeCategory; });
  }
  if (filterText) {
    visible = visible.filter(function (a) {
      return (
        a.name.toLowerCase().includes(filterText) ||
        a.file.toLowerCase().includes(filterText)
      );
    });
  }

  if (assets.length === 0) {
    empty.style.display = "flex";
    grid.style.display  = "none";
    grid.innerHTML      = "";
    return;
  }

  empty.style.display = "none";
  grid.style.display  = "grid";

  if (visible.length === 0) {
    grid.innerHTML =
      "<p class=\"no-results\">No assets match \"<strong>" +
      xmlEsc(filterText || activeCategory) + "</strong>\"</p>";
    return;
  }

  const base = pagesBase() + "assets/";
  grid.innerHTML = visible.map(function (a) {
    const realIdx = assets.indexOf(a);
    return (
      "<div class=\"icon-card\" onclick=\"insertAsset(" + realIdx + ")\" title=\"Click to insert\">" +
        "<div class=\"icon-preview\">" +
          "<img src=\"" + base + a.file + "\" alt=\"" + xmlEsc(a.name) + "\" loading=\"lazy\" />" +
        "</div>" +
        "<div class=\"icon-name\" title=\"" + xmlEsc(a.name) + "\">" + xmlEsc(a.name) + "</div>" +
      "</div>"
    );
  }).join("");
}

// ── Panel & modal helpers ─────────────────────────────────────────────────────
function openPanel(id) {
  document.getElementById("overlay").classList.add("show");
  document.getElementById(id).classList.add("open");
}

function closePanel(id) {
  document.getElementById(id).classList.remove("open");
  // Only hide overlay if no other panels or modals remain open
  var anyOpen = document.querySelectorAll(".side-panel.open, .modal.open").length > 0;
  if (!anyOpen) document.getElementById("overlay").classList.remove("show");
}

function closeAllPanels() {
  document.querySelectorAll(".side-panel, .modal").forEach(function (el) {
    el.classList.remove("open");
  });
  document.getElementById("overlay").classList.remove("show");
  window._pendingFile = null;
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function toTitleCase(str) {
  return str.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}

function xmlEsc(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ── Toast ─────────────────────────────────────────────────────────────────────
var toastTimer;
function toast(msg, type) {
  var el = document.getElementById("toast");
  el.textContent = msg;
  el.className   = "toast " + (type || "info") + " show";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.classList.remove("show"); }, 3200);
}
