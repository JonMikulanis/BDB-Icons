// ── Config ────────────────────────────────────────────────────────────────────
const GITHUB_API     = "https://api.github.com";
const DEFAULT_OWNER  = "JonMikulanis";
const DEFAULT_REPO   = "BDB-Icons";
const DEFAULT_BRANCH = "main";

// ── State ─────────────────────────────────────────────────────────────────────
let assets          = [];
let filterText      = "";
let activeCategory  = "All";
let cfg             = {};
let localCategories = [];
let activeColor     = null;   // null = no tint; hex string = tint SVGs on insert
let savedColors     = [];

// Bentgo brand palette — update hex values here if needed
const DEFAULT_COLORS = [
  "#3D8BCD",  // Blue
  "#F5A623",  // Gold
  "#3B9E8F",  // Teal
  "#C27B4A",  // Terracotta
  "#8B77C8",  // Purple
  "#D47B70",  // Coral
  "#F07840",  // Orange
  "#7D9B65",  // Sage
  "#484850",  // Charcoal
  "#F0EBE0",  // Cream
];

// ── Office Init ───────────────────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    initSettings();
    initLocalCategories();
    initColors();
    initDragDrop();
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

// ── Local categories (pre-created folders, stored in localStorage) ────────────
function initLocalCategories() {
  try {
    localCategories = JSON.parse(localStorage.getItem("local_cats") || "[]");
  } catch (e) {
    localCategories = [];
  }
}

function saveLocalCategories() {
  localStorage.setItem("local_cats", JSON.stringify(localCategories));
}

function addLocalCategory(name) {
  name = name.trim();
  if (!name) return;
  // Don't duplicate an existing library or local category
  var allCats = assets.map(function (a) { return a.category; }).concat(localCategories);
  if (allCats.map(function (c) { return c.toLowerCase(); }).includes(name.toLowerCase())) {
    toast("\"" + name + "\" already exists.", "info");
    return;
  }
  localCategories.push(name);
  saveLocalCategories();
  renderCategoryChips();
  // Auto-select the new category
  setCategory(name);
  toast("Folder \"" + name + "\" created. Upload assets to populate it.", "success");
}

function showAddCategoryInput() {
  var wrap = document.getElementById("add-cat-wrap");
  var input = document.getElementById("add-cat-input");
  wrap.classList.add("open");
  input.value = "";
  input.focus();
}

function hideAddCategoryInput() {
  document.getElementById("add-cat-wrap").classList.remove("open");
}

function handleAddCatKey(e) {
  if (e.key === "Enter") {
    addLocalCategory(document.getElementById("add-cat-input").value);
    hideAddCategoryInput();
  } else if (e.key === "Escape") {
    hideAddCategoryInput();
  }
}

// ── Colors ────────────────────────────────────────────────────────────────────
function initColors() {
  try {
    savedColors = JSON.parse(localStorage.getItem("saved_colors") || "null") || DEFAULT_COLORS.slice();
  } catch (e) {
    savedColors = DEFAULT_COLORS.slice();
  }
  activeColor = localStorage.getItem("active_color") || null;
  renderColorStrip();
}

function setActiveColor(color) {
  activeColor = (activeColor === color) ? null : color;  // toggle off if re-clicked
  localStorage.setItem("active_color", activeColor || "");
  renderColorStrip();
}

function clearActiveColor() {
  activeColor = null;
  localStorage.setItem("active_color", "");
  renderColorStrip();
}

function addCustomColor(hex) {
  hex = hex.trim().toUpperCase();
  if (!hex.startsWith("#")) hex = "#" + hex;
  if (!/^#[0-9A-F]{6}$/i.test(hex)) {
    toast("Enter a valid 6-digit hex color, e.g. #FF6B00", "error");
    return;
  }
  if (!savedColors.map(function (c) { return c.toUpperCase(); }).includes(hex)) {
    savedColors.push(hex);
    localStorage.setItem("saved_colors", JSON.stringify(savedColors));
  }
  setActiveColor(hex);
  hideColorInput();
}

function showColorInput() {
  var wrap  = document.getElementById("color-input-wrap");
  var input = document.getElementById("color-hex-input");
  wrap.classList.add("open");
  input.value = activeColor || "";
  input.focus();
  input.select();
}

function hideColorInput() {
  document.getElementById("color-input-wrap").classList.remove("open");
}

function handleColorKey(e) {
  if (e.key === "Enter") {
    addCustomColor(document.getElementById("color-hex-input").value);
  } else if (e.key === "Escape") {
    hideColorInput();
  }
}

function renderColorStrip() {
  var strip = document.getElementById("color-swatches");
  if (!strip) return;

  // "Original / no tint" toggle
  var html = (
    "<button class=\"color-swatch swatch-none" + (activeColor === null ? " active" : "") +
    "\" onclick=\"clearActiveColor()\" title=\"Original colors (no tint)\">" +
    "<svg viewBox=\"0 0 14 14\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"1.8\">" +
    "<line x1=\"3\" y1=\"11\" x2=\"11\" y2=\"3\"/>" +
    "</svg></button>"
  );

  // Brand / saved color swatches
  html += savedColors.map(function (c) {
    var light = isLightColor(c);
    return (
      "<button class=\"color-swatch" + (activeColor === c.toUpperCase() || activeColor === c ? " active" : "") + "\"" +
      " style=\"background:" + c + ";" + (light ? "border-color:rgba(0,0,0,0.18);" : "") + "\"" +
      " onclick=\"setActiveColor('" + c + "')\"" +
      " title=\"" + c + "\"></button>"
    );
  }).join("");

  strip.innerHTML = html;

  // Update the active-color label beneath the strip
  var label = document.getElementById("color-active-label");
  if (label) {
    if (activeColor) {
      label.textContent = activeColor.toUpperCase();
      label.style.color = activeColor;
      label.style.display = "inline";
    } else {
      label.style.display = "none";
    }
  }
}

// Returns true if a hex color is light enough to need a dark border
function isLightColor(hex) {
  var r = parseInt(hex.slice(1, 3), 16);
  var g = parseInt(hex.slice(3, 5), 16);
  var b = parseInt(hex.slice(5, 7), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) > 190;
}

// Walk an SVG's elements and replace all non-none fills/strokes with color
function applyColorToSvg(svgText, color) {
  if (!color) return svgText;
  try {
    var parser = new DOMParser();
    var doc    = parser.parseFromString(svgText, "image/svg+xml");
    var root   = doc.documentElement;
    var all    = [root].concat(Array.from(root.querySelectorAll("*")));
    all.forEach(function (el) {
      var fill = el.getAttribute("fill");
      if (fill && fill.toLowerCase() !== "none") el.setAttribute("fill", color);
      var stroke = el.getAttribute("stroke");
      if (stroke && stroke.toLowerCase() !== "none") el.setAttribute("stroke", color);
      // Handle inline style fill/stroke
      var style = el.getAttribute("style");
      if (style) {
        style = style.replace(/(fill\s*:\s*)(?!none\b)[^;"}]+/gi,   "$1" + color);
        style = style.replace(/(stroke\s*:\s*)(?!none\b)[^;"}]+/gi, "$1" + color);
        el.setAttribute("style", style);
      }
    });
    return new XMLSerializer().serializeToString(doc);
  } catch (e) {
    return svgText;  // if parsing fails, insert original
  }
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
      const svgText    = await res.text();
      const coloredSvg = applyColorToSvg(svgText, activeColor);
      Office.context.document.setSelectedDataAsync(
        coloredSvg,
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

// ── Drag and drop ─────────────────────────────────────────────────────────────
function initDragDrop() {
  var zone    = document.getElementById("drop-zone");
  var overlay = document.getElementById("drag-overlay");

  zone.addEventListener("dragenter", function (e) {
    e.preventDefault();
    zone.classList.add("drag-over");
  });

  zone.addEventListener("dragover", function (e) {
    e.preventDefault();
    zone.classList.add("drag-over");
  });

  zone.addEventListener("dragleave", function (e) {
    // Only clear if leaving the zone entirely (not entering a child)
    if (!zone.contains(e.relatedTarget)) {
      zone.classList.remove("drag-over");
    }
  });

  zone.addEventListener("drop", function (e) {
    e.preventDefault();
    zone.classList.remove("drag-over");
    var files = Array.from(e.dataTransfer.files).filter(function (f) {
      return /\.(svg|png|jpg|jpeg)$/i.test(f.name);
    });
    if (!files.length) {
      toast("Drop SVG, PNG, or JPG files only.", "error");
      return;
    }
    if (!cfg.token) {
      toast("Add a GitHub token in ⚙ Settings first.", "error");
      return;
    }
    showBatchModal(files);
  });
}

// ── Upload flow — batch (handles 1 or many files) ─────────────────────────────
function handleUpload(input) {
  var files = Array.from(input.files);
  input.value = "";
  if (!files.length) return;
  if (!cfg.token) {
    toast("Add a GitHub token in ⚙ Settings first.", "error");
    return;
  }
  showBatchModal(files);
}

function showBatchModal(files) {
  window._pendingFiles = files;

  // Populate category datalist with existing categories
  var cats = [...new Set(assets.map(function (a) { return a.category; }).filter(Boolean))];
  document.getElementById("batch-cat-list").innerHTML = cats.map(function (c) {
    return "<option value=\"" + xmlEsc(c) + "\">";
  }).join("");
  document.getElementById("batch-category").value = "";

  // Build the file rows
  var list = document.getElementById("batch-file-list");
  list.innerHTML = files.map(function (f, i) {
    var ext      = f.name.split(".").pop().toLowerCase();
    var baseName = toTitleCase(f.name.replace(/\.[^.]+$/, "").replace(/[-_]/g, " "));
    return (
      "<div class=\"batch-item\">" +
        "<span class=\"file-type-badge\">" + ext.toUpperCase() + "</span>" +
        "<div class=\"batch-item-body\">" +
          "<div class=\"batch-filename\">" + xmlEsc(f.name) + "</div>" +
          "<input type=\"text\" class=\"field-input batch-name\" data-index=\"" + i + "\"" +
            " value=\"" + xmlEsc(baseName) + "\" placeholder=\"Display name\" />" +
        "</div>" +
      "</div>"
    );
  }).join("");

  // Update button label
  var n = files.length;
  document.getElementById("batch-header-count").textContent =
    "Upload " + n + " File" + (n > 1 ? "s" : "");
  document.getElementById("batch-btn-label").textContent =
    "Upload " + n + " File" + (n > 1 ? "s" : "");

  openPanel("batch-modal");
}

async function confirmBatchUpload() {
  var files    = window._pendingFiles;
  if (!files || !files.length) return;

  var category = document.getElementById("batch-category").value.trim() || "General";

  // Collect names from the editable fields
  var items = files.map(function (file, i) {
    var nameEl      = document.querySelector(".batch-name[data-index=\"" + i + "\"]");
    var displayName = nameEl ? nameEl.value.trim() : toTitleCase(file.name.replace(/\.[^.]+$/, "").replace(/[-_]/g, " "));
    var filename    = file.name.toLowerCase().replace(/\s+/g, "-");
    var ext         = filename.split(".").pop().toLowerCase();
    var type        = (ext === "svg") ? "svg" : "image";
    return { file: file, displayName: displayName, filename: filename, type: type };
  });

  closePanel("batch-modal");
  setBusy(true);

  // ── Phase 1: Upload all asset files sequentially ──
  var uploaded  = [];
  var failCount = 0;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    toast("Uploading " + (i + 1) + " of " + items.length + ": " + item.displayName + "…", "info");
    try {
      var fileB64 = await fileToBase64(item.file);
      await ghPut("contents/assets/" + item.filename, {
        message: "Add asset: " + item.displayName,
        content: fileB64,
        branch:  cfg.branch,
      });
      uploaded.push(item);
    } catch (e) {
      failCount++;
      console.error("Failed to upload " + item.filename, e);
    }
  }

  // ── Phase 2: Update library.json once for all uploaded files ──
  if (uploaded.length > 0) {
    try {
      toast("Updating library…", "info");
      var currentLib = [];
      var libSha     = null;
      try {
        var libData = await ghGet("contents/library.json");
        libSha      = libData.sha;
        currentLib  = JSON.parse(atob(libData.content.replace(/\s/g, "")));
      } catch (e) {
        // library.json doesn't exist yet — will be created
      }

      uploaded.forEach(function (it) {
        currentLib.push({ name: it.displayName, file: it.filename, type: it.type, category: category });
      });

      var libContent = btoa(unescape(encodeURIComponent(JSON.stringify(currentLib, null, 2))));
      var putBody    = {
        message: "Update library.json: add " + uploaded.length + " asset(s)",
        content: libContent,
        branch:  cfg.branch,
      };
      if (libSha) putBody.sha = libSha;
      await ghPut("contents/library.json", putBody);
    } catch (e) {
      toast("Files uploaded but library.json update failed: " + (e.message || ""), "error");
      setBusy(false);
      window._pendingFiles = null;
      return;
    }
  }

  setBusy(false);
  window._pendingFiles = null;

  if (failCount === 0) {
    toast("All " + uploaded.length + " files uploaded! Refreshing…", "success");
  } else {
    toast(uploaded.length + " uploaded, " + failCount + " failed. Refreshing…", "error");
  }
  setTimeout(loadLibrary, 4000);
}

// ── GitHub API helpers ────────────────────────────────────────────────────────
async function ghGet(path) {
  var res = await fetch(apiBase() + "/" + path, {
    headers: {
      Authorization: "token " + cfg.token,
      Accept:        "application/vnd.github+json",
    }
  });
  if (!res.ok) {
    var err = await res.json().catch(function () { return {}; });
    throw new Error(err.message || "HTTP " + res.status);
  }
  return res.json();
}

async function ghPut(path, body) {
  var res = await fetch(apiBase() + "/" + path, {
    method:  "PUT",
    headers: {
      Authorization:  "token " + cfg.token,
      Accept:         "application/vnd.github+json",
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    var err = await res.json().catch(function () { return {}; });
    throw new Error(err.message || "HTTP " + res.status);
  }
  return res.json();
}

function fileToBase64(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload  = function () { resolve(reader.result.split(",")[1]); };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── Category chips ────────────────────────────────────────────────────────────
function renderCategoryChips() {
  var libCats   = [...new Set(assets.map(function (a) { return a.category; }).filter(Boolean))].sort();

  // Prune local categories that have graduated into library.json
  localCategories = localCategories.filter(function (c) {
    return !libCats.map(function (l) { return l.toLowerCase(); }).includes(c.toLowerCase());
  });
  saveLocalCategories();

  var allCats = ["All"].concat(libCats);

  if (!allCats.includes(activeCategory) && !localCategories.includes(activeCategory)) {
    activeCategory = "All";
  }

  // Render library chips
  var chipsHtml = allCats.map(function (c) {
    return (
      "<button class=\"chip" + (c === activeCategory ? " active" : "") +
      "\" onclick=\"setCategory('" + c.replace(/'/g, "\\'") + "')\">" +
      xmlEsc(c) + "</button>"
    );
  }).join("");

  // Render local-only (empty) chips with dashed style
  chipsHtml += localCategories.map(function (c) {
    return (
      "<button class=\"chip chip-empty" + (c === activeCategory ? " active" : "") +
      "\" onclick=\"setCategory('" + c.replace(/'/g, "\\'") + "')\">" +
      xmlEsc(c) + "</button>"
    );
  }).join("");

  document.getElementById("category-chips").innerHTML = chipsHtml;
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
  var grid    = document.getElementById("icon-grid");
  var empty   = document.getElementById("empty-state");
  var counter = document.getElementById("icon-count");
  counter.textContent = assets.length;

  var visible = assets.slice();
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

  var base = pagesBase() + "assets/";
  grid.innerHTML = visible.map(function (a) {
    var realIdx = assets.indexOf(a);
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
  var anyOpen = document.querySelectorAll(".side-panel.open, .modal.open").length > 0;
  if (!anyOpen) document.getElementById("overlay").classList.remove("show");
}

function closeAllPanels() {
  document.querySelectorAll(".side-panel, .modal").forEach(function (el) {
    el.classList.remove("open");
  });
  document.getElementById("overlay").classList.remove("show");
  window._pendingFiles = null;
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
