// ── Config ────────────────────────────────────────────────────────────────────
const GITHUB_API     = "https://api.github.com";
const DEFAULT_OWNER  = "JonMikulanis";
const DEFAULT_REPO   = "BDB-Icons";
const DEFAULT_BRANCH = "main";

// ── Insert size configs ───────────────────────────────────────────────────────
// px = SVG width/height attribute; pt = Office image points (72pt ≈ 1 inch)
const INSERT_SIZES = [
  { label: "Orig", px: null, pt: null  },
  { label: "S",    px: 72,   pt: 72    },
  { label: "M",    px: 144,  pt: 144   },
  { label: "L",    px: 288,  pt: 288   },
];

// ── State ─────────────────────────────────────────────────────────────────────
let assets          = [];
let filterText      = "";
let activeCategory  = "All";
let cfg             = {};
let localCategories = [];
let activeColor     = null;
let savedColors     = [];
let recentlyUsed    = [];       // filenames, most-recent first, max 10
let favorites       = new Set();// filenames
let sortOrder       = "default";// "default" | "az" | "za"
let insertSize      = 0;        // index into INSERT_SIZES
let selectMode      = false;
let selectedFiles   = new Set();
let ctxMenuIdx      = -1;       // assets[] index for right-click context menu
let previewIdx      = -1;       // assets[] index for preview modal
let renameIdx       = -1;
let moveIdx         = -1;
let deleteIdx       = -1;

// Bentgo brand palette
const DEFAULT_COLORS = [
  "#3D8BCD","#F5A623","#3B9E8F","#C27B4A","#8B77C8",
  "#D47B70","#F07840","#7D9B65","#484850","#F0EBE0",
];

// ── Office Init ────────────────────────────────────────────────────────────────
Office.onReady(function(info) {
  if (info.host === Office.HostType.PowerPoint) {
    initSettings();
    initLocalCategories();
    initColors();
    initRecentlyUsed();
    initFavorites();
    initInsertSize();
    initDragDrop();
    loadLibrary();
    // Close context menu when clicking anywhere outside it
    document.addEventListener("click", function(e) {
      if (!e.target.closest("#context-menu")) hideContextMenu();
    });
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

// ── Local categories ──────────────────────────────────────────────────────────
function initLocalCategories() {
  try { localCategories = JSON.parse(localStorage.getItem("local_cats") || "[]"); }
  catch(e) { localCategories = []; }
}
function saveLocalCategories() {
  localStorage.setItem("local_cats", JSON.stringify(localCategories));
}
function addLocalCategory(name) {
  name = name.trim();
  if (!name) return;
  var allCats = assets.map(function(a) { return a.category; }).concat(localCategories);
  if (allCats.map(function(c) { return c.toLowerCase(); }).includes(name.toLowerCase())) {
    toast('"' + name + '" already exists.', "info"); return;
  }
  localCategories.push(name);
  saveLocalCategories();
  renderCategoryChips();
  setCategory(name);
  toast('Folder "' + name + '" created. Upload assets to populate it.', "success");
}
function showAddCategoryInput() {
  var wrap = document.getElementById("add-cat-wrap");
  var input = document.getElementById("add-cat-input");
  wrap.classList.add("open"); input.value = ""; input.focus();
}
function hideAddCategoryInput() {
  document.getElementById("add-cat-wrap").classList.remove("open");
}
function handleAddCatKey(e) {
  if (e.key === "Enter") { addLocalCategory(document.getElementById("add-cat-input").value); hideAddCategoryInput(); }
  else if (e.key === "Escape") { hideAddCategoryInput(); }
}

// ── Recently Used ─────────────────────────────────────────────────────────────
function initRecentlyUsed() {
  try { recentlyUsed = JSON.parse(localStorage.getItem("recently_used") || "[]"); }
  catch(e) { recentlyUsed = []; }
}
function trackRecentlyUsed(filename) {
  recentlyUsed = recentlyUsed.filter(function(f) { return f !== filename; });
  recentlyUsed.unshift(filename);
  if (recentlyUsed.length > 10) recentlyUsed = recentlyUsed.slice(0, 10);
  localStorage.setItem("recently_used", JSON.stringify(recentlyUsed));
  // Re-render chips to show updated Recent count
  renderCategoryChips();
}

// ── Favorites ─────────────────────────────────────────────────────────────────
function initFavorites() {
  try { favorites = new Set(JSON.parse(localStorage.getItem("favorites") || "[]")); }
  catch(e) { favorites = new Set(); }
}
function toggleFavorite(filename) {
  if (favorites.has(filename)) {
    favorites.delete(filename);
    toast("Removed from Favorites.", "info");
  } else {
    favorites.add(filename);
    toast("Added to Favorites ★", "success");
  }
  localStorage.setItem("favorites", JSON.stringify(Array.from(favorites)));
  renderCategoryChips();
  renderGrid();
}

// ── Sort ──────────────────────────────────────────────────────────────────────
function toggleSort() {
  var order = ["default", "az", "za"];
  var idx = order.indexOf(sortOrder);
  sortOrder = order[(idx + 1) % order.length];
  updateSortBtn();
  renderGrid();
}
function updateSortBtn() {
  var btn = document.getElementById("sort-btn");
  if (!btn) return;
  var icons = { default: "↕", az: "↑", za: "↓" };
  var labels = { default: "Default order", az: "A → Z", za: "Z → A" };
  var ind = document.getElementById("sort-indicator");
  if (ind) ind.textContent = icons[sortOrder];
  btn.title = "Sort: " + labels[sortOrder];
}
function getSortedAssets(list) {
  if (sortOrder === "az") return list.slice().sort(function(a, b) { return a.name.localeCompare(b.name); });
  if (sortOrder === "za") return list.slice().sort(function(a, b) { return b.name.localeCompare(a.name); });
  return list;
}

// ── Insert size ───────────────────────────────────────────────────────────────
function initInsertSize() {
  insertSize = parseInt(localStorage.getItem("insert_size") || "0") || 0;
  syncSizeBtns();
}
function setInsertSize(idx) {
  insertSize = idx;
  localStorage.setItem("insert_size", idx);
  syncSizeBtns();
}
function syncSizeBtns() {
  document.querySelectorAll(".size-btn").forEach(function(btn) {
    btn.classList.toggle("active", parseInt(btn.dataset.size) === insertSize);
  });
}

// ── Select mode ───────────────────────────────────────────────────────────────
function toggleSelectMode() {
  selectMode = !selectMode;
  selectedFiles.clear();
  var btn = document.getElementById("select-btn");
  if (btn) btn.classList.toggle("active", selectMode);
  updateSelectBar();
  renderGrid();
}
function toggleSelectAsset(filename) {
  if (selectedFiles.has(filename)) selectedFiles.delete(filename);
  else selectedFiles.add(filename);
  updateSelectBar();
  renderGrid();
}
function updateSelectBar() {
  var bar = document.getElementById("select-bar");
  if (!bar) return;
  if (selectMode && selectedFiles.size > 0) {
    bar.classList.add("show");
    document.getElementById("select-count").textContent = selectedFiles.size + " selected";
  } else {
    bar.classList.remove("show");
  }
}
function cancelSelectMode() {
  selectMode = false;
  selectedFiles.clear();
  var btn = document.getElementById("select-btn");
  if (btn) btn.classList.remove("active");
  updateSelectBar();
  renderGrid();
}

// ── Colors ────────────────────────────────────────────────────────────────────
function initColors() {
  try { savedColors = JSON.parse(localStorage.getItem("saved_colors") || "null") || DEFAULT_COLORS.slice(); }
  catch(e) { savedColors = DEFAULT_COLORS.slice(); }
  activeColor = localStorage.getItem("active_color") || null;
  renderColorStrip();
}
function setActiveColor(color) {
  activeColor = (activeColor === color) ? null : color;
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
  if (!/^#[0-9A-F]{6}$/i.test(hex)) { toast("Enter a valid 6-digit hex, e.g. #FF6B00", "error"); return; }
  if (!savedColors.map(function(c) { return c.toUpperCase(); }).includes(hex)) {
    savedColors.push(hex);
    localStorage.setItem("saved_colors", JSON.stringify(savedColors));
  }
  setActiveColor(hex);
  hideColorInput();
}
function showColorInput() {
  var wrap = document.getElementById("color-input-wrap");
  var input = document.getElementById("color-hex-input");
  wrap.classList.add("open"); input.value = activeColor || ""; input.focus(); input.select();
}
function hideColorInput() { document.getElementById("color-input-wrap").classList.remove("open"); }
function handleColorKey(e) {
  if (e.key === "Enter") addCustomColor(document.getElementById("color-hex-input").value);
  else if (e.key === "Escape") hideColorInput();
}
function renderColorStrip() {
  var strip = document.getElementById("color-swatches");
  if (!strip) return;
  var html = '<button class="color-swatch swatch-none' + (activeColor === null ? " active" : "") +
    '" onclick="clearActiveColor()" title="Original colors (no tint)">' +
    '<svg viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.8">' +
    '<line x1="3" y1="11" x2="11" y2="3"/></svg></button>';
  html += savedColors.map(function(c) {
    var light = isLightColor(c);
    return '<button class="color-swatch' + (activeColor === c.toUpperCase() || activeColor === c ? " active" : "") + '"' +
      ' style="background:' + c + ';' + (light ? "border-color:rgba(0,0,0,0.18);" : "") + '"' +
      ' onclick="setActiveColor(\'' + c + '\')" title="' + c + '"></button>';
  }).join("");
  strip.innerHTML = html;
  var label = document.getElementById("color-active-label");
  if (label) {
    if (activeColor) { label.textContent = activeColor.toUpperCase(); label.style.color = activeColor; label.style.display = "inline"; }
    else { label.style.display = "none"; }
  }
}
function isLightColor(hex) {
  var r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return (0.299*r + 0.587*g + 0.114*b) > 190;
}

// ── SVG helpers ───────────────────────────────────────────────────────────────
function applyColorToSvg(svgText, color) {
  if (!color) return svgText;
  try {
    var parser = new DOMParser();
    var doc = parser.parseFromString(svgText, "image/svg+xml");
    var root = doc.documentElement;
    var all = [root].concat(Array.from(root.querySelectorAll("*")));
    all.forEach(function(el) {
      var fill = el.getAttribute("fill");
      if (fill && fill.toLowerCase() !== "none") el.setAttribute("fill", color);
      var stroke = el.getAttribute("stroke");
      if (stroke && stroke.toLowerCase() !== "none") el.setAttribute("stroke", color);
      var style = el.getAttribute("style");
      if (style) {
        style = style.replace(/(fill\s*:\s*)(?!none\b)[^;"}]+/gi, "$1" + color);
        style = style.replace(/(stroke\s*:\s*)(?!none\b)[^;"}]+/gi, "$1" + color);
        el.setAttribute("style", style);
      }
    });
    return new XMLSerializer().serializeToString(doc);
  } catch(e) { return svgText; }
}

function applySizeToSvg(svgText, px) {
  if (!px) return svgText;
  try {
    var parser = new DOMParser();
    var doc = parser.parseFromString(svgText, "image/svg+xml");
    var root = doc.documentElement;
    root.setAttribute("width", px);
    root.setAttribute("height", px);
    return new XMLSerializer().serializeToString(doc);
  } catch(e) { return svgText; }
}

function pagesBase() { return "https://" + cfg.owner + ".github.io/" + cfg.repo + "/"; }
function apiBase()   { return GITHUB_API + "/repos/" + cfg.owner + "/" + cfg.repo; }

// ── Load library ──────────────────────────────────────────────────────────────
async function loadLibrary() {
  setBusy(true);
  try {
    const url = pagesBase() + "library.json?_=" + Date.now();
    const res = await fetch(url);
    if (!res.ok) throw new Error("HTTP " + res.status);
    assets = await res.json();
    renderCategoryChips();
    renderGrid();
  } catch(e) {
    toast("Could not load library. Check Settings or try Refresh.", "error");
    assets = [];
    renderCategoryChips();
    renderGrid();
  } finally { setBusy(false); }
}
function setBusy(on) {
  document.getElementById("refresh-btn").classList.toggle("spinning", on);
}

// ── Insert asset onto slide ───────────────────────────────────────────────────
async function insertAsset(idx) {
  var asset = assets[idx];
  if (!asset) return;
  var url = pagesBase() + "assets/" + asset.file;
  var sizeConfig = INSERT_SIZES[insertSize] || INSERT_SIZES[0];
  try {
    if (asset.type === "svg") {
      var res = await fetch(url);
      if (!res.ok) throw new Error("Could not fetch SVG (HTTP " + res.status + ")");
      var svgText = await res.text();
      svgText = applyColorToSvg(svgText, activeColor);
      svgText = applySizeToSvg(svgText, sizeConfig.px);
      Office.context.document.setSelectedDataAsync(
        svgText, { coercionType: Office.CoercionType.XmlSvg },
        function(r) {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            trackRecentlyUsed(asset.file);
            toast("Inserted: " + asset.name, "success");
          } else { toast("Insert failed — SVG may be malformed.", "error"); }
        }
      );
    } else {
      var res = await fetch(url);
      if (!res.ok) throw new Error("Could not fetch image (HTTP " + res.status + ")");
      var blob = await res.blob();
      var b64 = await blobToBase64(blob);
      var opts = { coercionType: Office.CoercionType.Image };
      if (sizeConfig.pt) { opts.imageWidth = sizeConfig.pt; opts.imageHeight = sizeConfig.pt; }
      Office.context.document.setSelectedDataAsync(b64, opts, function(r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          trackRecentlyUsed(asset.file);
          toast("Inserted: " + asset.name, "success");
        } else { toast("Insert failed — check image format.", "error"); }
      });
    }
  } catch(e) { toast("Could not load asset: " + (e.message || "unknown error"), "error"); }
}

function blobToBase64(blob) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload  = function() { resolve(reader.result.split(",")[1]); };
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// ── Preview modal ─────────────────────────────────────────────────────────────
function previewAsset(idx) {
  var asset = assets[idx];
  if (!asset) return;
  previewIdx = idx;
  document.getElementById("preview-name").textContent = asset.name;
  document.getElementById("preview-category").textContent = asset.category || "General";
  document.getElementById("preview-type").textContent = asset.type === "svg" ? "SVG" : asset.file.split(".").pop().toUpperCase();
  document.getElementById("preview-img").src = pagesBase() + "assets/" + asset.file;
  syncSizeBtns();
  openPanel("preview-modal");
}
function insertFromPreview() {
  closePanel("preview-modal");
  if (previewIdx >= 0) insertAsset(previewIdx);
}

// ── Right-click context menu ──────────────────────────────────────────────────
function showContextMenu(e, idx) {
  e.preventDefault();
  e.stopPropagation();
  ctxMenuIdx = idx;
  var asset = assets[idx];
  if (!asset) return;
  // Update the favorite label
  document.getElementById("ctx-fav-label").textContent =
    favorites.has(asset.file) ? "Remove from Favorites" : "Add to Favorites";
  // Position the menu
  var menu = document.getElementById("context-menu");
  menu.style.left = "-999px"; menu.style.top = "-999px";
  menu.classList.add("show");
  var menuW = menu.offsetWidth, menuH = menu.offsetHeight;
  var vpW = window.innerWidth, vpH = window.innerHeight;
  var x = e.clientX, y = e.clientY;
  if (x + menuW > vpW - 4) x = vpW - menuW - 4;
  if (y + menuH > vpH - 4) y = vpH - menuH - 4;
  menu.style.left = x + "px";
  menu.style.top  = y + "px";
}
function hideContextMenu() {
  document.getElementById("context-menu").classList.remove("show");
  ctxMenuIdx = -1;
}
function ctxPreview()  { var i = ctxMenuIdx; hideContextMenu(); if (i >= 0) previewAsset(i); }
function ctxFavorite() { var i = ctxMenuIdx; hideContextMenu(); if (i >= 0) toggleFavorite(assets[i].file); }
function ctxRename()   { var i = ctxMenuIdx; hideContextMenu(); if (i >= 0) showRenameModal(i); }
function ctxMove()     { var i = ctxMenuIdx; hideContextMenu(); if (i >= 0) showMoveModal(i); }
function ctxDelete()   { var i = ctxMenuIdx; hideContextMenu(); if (i >= 0) deleteAssetConfirm(i); }

// ── Rename modal ──────────────────────────────────────────────────────────────
function showRenameModal(idx) {
  renameIdx = idx;
  document.getElementById("rename-input").value = assets[idx].name;
  openPanel("rename-modal");
  setTimeout(function() {
    var el = document.getElementById("rename-input");
    el.focus(); el.select();
  }, 60);
}
function handleRenameKey(e) {
  if (e.key === "Enter") confirmRename();
  else if (e.key === "Escape") closePanel("rename-modal");
}
async function confirmRename() {
  var newName = document.getElementById("rename-input").value.trim();
  if (!newName || renameIdx < 0) { closePanel("rename-modal"); return; }
  if (newName === assets[renameIdx].name) { closePanel("rename-modal"); return; }
  closePanel("rename-modal");
  await updateLibraryEntry(renameIdx, { name: newName });
}

// ── Move modal ────────────────────────────────────────────────────────────────
function showMoveModal(idx) {
  moveIdx = idx;
  var cats = [...new Set(assets.map(function(a) { return a.category; }).filter(Boolean))];
  document.getElementById("move-cat-list").innerHTML = cats.map(function(c) { return "<option value=\"" + xmlEsc(c) + "\">"; }).join("");
  document.getElementById("move-cat-input").value = assets[idx].category || "";
  openPanel("move-modal");
  setTimeout(function() { document.getElementById("move-cat-input").focus(); }, 60);
}
function handleMoveCatKey(e) {
  if (e.key === "Enter") confirmMove();
  else if (e.key === "Escape") closePanel("move-modal");
}
async function confirmMove() {
  var newCat = document.getElementById("move-cat-input").value.trim();
  if (!newCat || moveIdx < 0) { closePanel("move-modal"); return; }
  if (newCat === assets[moveIdx].category) { closePanel("move-modal"); return; }
  closePanel("move-modal");
  await updateLibraryEntry(moveIdx, { category: newCat });
}

// ── Update a library.json entry (rename / move) ───────────────────────────────
async function updateLibraryEntry(idx, changes) {
  if (!cfg.token) { toast("Add a GitHub token in ⚙ Settings first.", "error"); return; }
  setBusy(true);
  try {
    var libData = await ghGet("contents/library.json");
    var lib = JSON.parse(atob(libData.content.replace(/\s/g, "")));
    var filename = assets[idx].file;
    var libEntry = lib.find(function(e) { return e.file === filename; });
    if (!libEntry) throw new Error("Asset not found in library.json");
    Object.assign(libEntry, changes);
    var libContent = btoa(unescape(encodeURIComponent(JSON.stringify(lib, null, 2))));
    await ghPut("contents/library.json", {
      message: "Update asset: " + filename,
      content: libContent,
      sha: libData.sha,
      branch: cfg.branch,
    });
    toast("Updated! Refreshing…", "success");
    setTimeout(loadLibrary, 4000);
  } catch(e) { toast("Update failed: " + (e.message || ""), "error"); }
  finally { setBusy(false); }
}

// ── Delete single asset ───────────────────────────────────────────────────────
function deleteAssetConfirm(idx) {
  deleteIdx = idx;
  document.getElementById("delete-asset-name").textContent = assets[idx].name;
  openPanel("delete-modal");
}
async function confirmDeleteAsset() {
  closePanel("delete-modal");
  if (deleteIdx < 0) return;
  var idx = deleteIdx;
  deleteIdx = -1;
  await deleteAssetsFromGitHub([assets[idx].file]);
}

// ── Bulk delete ───────────────────────────────────────────────────────────────
function deleteSelectedConfirm() {
  if (selectedFiles.size === 0) return;
  document.getElementById("delete-bulk-count").textContent = selectedFiles.size;
  openPanel("delete-bulk-modal");
}
async function confirmDeleteBulk() {
  closePanel("delete-bulk-modal");
  var filenames = Array.from(selectedFiles);
  selectedFiles.clear();
  updateSelectBar();
  await deleteAssetsFromGitHub(filenames);
  // Exit select mode after bulk delete
  selectMode = false;
  var btn = document.getElementById("select-btn");
  if (btn) btn.classList.remove("active");
}

// Core delete: removes files from GitHub + updates library.json
async function deleteAssetsFromGitHub(filenames) {
  if (!cfg.token) { toast("Add a GitHub token in ⚙ Settings first.", "error"); return; }
  setBusy(true);
  toast("Deleting " + filenames.length + " file(s)…", "info");
  var deleted = [], failCount = 0;
  for (var i = 0; i < filenames.length; i++) {
    var fn = filenames[i];
    try {
      var fileData = await ghGet("contents/assets/" + fn);
      await ghDelete("contents/assets/" + fn, fileData.sha);
      deleted.push(fn);
    } catch(e) { failCount++; console.error("Failed to delete " + fn, e); }
  }
  if (deleted.length > 0) {
    try {
      toast("Updating library…", "info");
      var libData = await ghGet("contents/library.json");
      var lib = JSON.parse(atob(libData.content.replace(/\s/g, "")));
      lib = lib.filter(function(e) { return !deleted.includes(e.file); });
      var libContent = btoa(unescape(encodeURIComponent(JSON.stringify(lib, null, 2))));
      await ghPut("contents/library.json", {
        message: "Delete " + deleted.length + " asset(s)",
        content: libContent,
        sha: libData.sha,
        branch: cfg.branch,
      });
    } catch(e) {
      toast("Files deleted but library.json update failed: " + (e.message || ""), "error");
      setBusy(false); return;
    }
  }
  setBusy(false);
  if (failCount === 0) toast("Deleted " + deleted.length + " file(s). Refreshing…", "success");
  else toast(deleted.length + " deleted, " + failCount + " failed. Refreshing…", "error");
  setTimeout(loadLibrary, 4000);
}

// ── Drag and drop ─────────────────────────────────────────────────────────────
function initDragDrop() {
  var zone = document.getElementById("drop-zone");
  zone.addEventListener("dragenter", function(e) { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragover",  function(e) { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", function(e) {
    if (!zone.contains(e.relatedTarget)) zone.classList.remove("drag-over");
  });
  zone.addEventListener("drop", function(e) {
    e.preventDefault();
    zone.classList.remove("drag-over");
    var files = Array.from(e.dataTransfer.files).filter(function(f) { return /\.(svg|png|jpg|jpeg)$/i.test(f.name); });
    if (!files.length) { toast("Drop SVG, PNG, or JPG files only.", "error"); return; }
    if (!cfg.token) { toast("Add a GitHub token in ⚙ Settings first.", "error"); return; }
    showBatchModal(files);
  });
}

// ── Upload flow ───────────────────────────────────────────────────────────────
function handleUpload(input) {
  var files = Array.from(input.files);
  input.value = "";
  if (!files.length) return;
  if (!cfg.token) { toast("Add a GitHub token in ⚙ Settings first.", "error"); return; }
  showBatchModal(files);
}

function showBatchModal(files) {
  window._pendingFiles = files;
  var existingFilenames = assets.map(function(a) { return a.file.toLowerCase(); });
  var cats = [...new Set(assets.map(function(a) { return a.category; }).filter(Boolean))];
  document.getElementById("batch-cat-list").innerHTML = cats.map(function(c) { return "<option value=\"" + xmlEsc(c) + "\">"; }).join("");
  document.getElementById("batch-category").value = "";

  var dupCount = 0;
  var list = document.getElementById("batch-file-list");
  list.innerHTML = files.map(function(f, i) {
    var ext = f.name.split(".").pop().toLowerCase();
    var baseName = toTitleCase(f.name.replace(/\.[^.]+$/, "").replace(/[-_]/g, " "));
    var filename = f.name.toLowerCase().replace(/\s+/g, "-");
    var isDup = existingFilenames.includes(filename);
    if (isDup) dupCount++;
    return (
      "<div class=\"batch-item" + (isDup ? " batch-item-dup" : "") + "\">" +
        "<span class=\"file-type-badge\">" + ext.toUpperCase() + "</span>" +
        "<div class=\"batch-item-body\">" +
          "<div class=\"batch-filename\">" + xmlEsc(f.name) + (isDup ? " <span class=\"dup-badge\">⚠ duplicate</span>" : "") + "</div>" +
          "<input type=\"text\" class=\"field-input batch-name\" data-index=\"" + i + "\"" +
          " value=\"" + xmlEsc(baseName) + "\" placeholder=\"Display name\" />" +
        "</div>" +
      "</div>"
    );
  }).join("");

  var n = files.length;
  var hdr = "Upload " + n + " File" + (n > 1 ? "s" : "");
  if (dupCount > 0) hdr += " — " + dupCount + " already exist" + (dupCount > 1 ? " (will overwrite)" : " (will overwrite)");
  document.getElementById("batch-header-count").textContent = hdr;
  document.getElementById("batch-btn-label").textContent = "Upload " + n + " File" + (n > 1 ? "s" : "");
  openPanel("batch-modal");
}

async function confirmBatchUpload() {
  var files = window._pendingFiles;
  if (!files || !files.length) return;
  var category = document.getElementById("batch-category").value.trim() || "General";
  var items = files.map(function(file, i) {
    var nameEl = document.querySelector(".batch-name[data-index=\"" + i + "\"]");
    var displayName = nameEl ? nameEl.value.trim() : toTitleCase(file.name.replace(/\.[^.]+$/, "").replace(/[-_]/g, " "));
    var filename = file.name.toLowerCase().replace(/\s+/g, "-");
    var ext = filename.split(".").pop().toLowerCase();
    var type = (ext === "svg") ? "svg" : "image";
    return { file: file, displayName: displayName, filename: filename, type: type };
  });
  closePanel("batch-modal");
  setBusy(true);

  var uploaded = [], failCount = 0;
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    toast("Uploading " + (i + 1) + " of " + items.length + ": " + item.displayName + "…", "info");
    try {
      var fileB64 = await fileToBase64(item.file);
      var putBody = { message: "Add asset: " + item.displayName, content: fileB64, branch: cfg.branch };
      // Fetch existing SHA for overwrite support
      try { var existing = await ghGet("contents/assets/" + item.filename); putBody.sha = existing.sha; } catch(e) {}
      await ghPut("contents/assets/" + item.filename, putBody);
      uploaded.push(item);
    } catch(e) { failCount++; console.error("Failed to upload " + item.filename, e); }
  }

  if (uploaded.length > 0) {
    try {
      toast("Updating library…", "info");
      var currentLib = [], libSha = null;
      try {
        var libData = await ghGet("contents/library.json");
        libSha = libData.sha;
        currentLib = JSON.parse(atob(libData.content.replace(/\s/g, "")));
      } catch(e) {}
      uploaded.forEach(function(it) {
        var existing = currentLib.find(function(e) { return e.file === it.filename; });
        if (existing) { existing.name = it.displayName; existing.category = category; }
        else currentLib.push({ name: it.displayName, file: it.filename, type: it.type, category: category });
      });
      var libContent = btoa(unescape(encodeURIComponent(JSON.stringify(currentLib, null, 2))));
      var putBody = {
        message: "Update library.json: add/update " + uploaded.length + " asset(s)",
        content: libContent, branch: cfg.branch,
      };
      if (libSha) putBody.sha = libSha;
      await ghPut("contents/library.json", putBody);
    } catch(e) {
      toast("Files uploaded but library.json update failed: " + (e.message || ""), "error");
      setBusy(false); window._pendingFiles = null; return;
    }
  }

  setBusy(false);
  window._pendingFiles = null;
  if (failCount === 0) toast("All " + uploaded.length + " files uploaded! Refreshing…", "success");
  else toast(uploaded.length + " uploaded, " + failCount + " failed. Refreshing…", "error");
  setTimeout(loadLibrary, 4000);
}

// ── GitHub API helpers ────────────────────────────────────────────────────────
async function ghGet(path) {
  var res = await fetch(apiBase() + "/" + path, {
    headers: { Authorization: "token " + cfg.token, Accept: "application/vnd.github+json" }
  });
  if (!res.ok) { var err = await res.json().catch(function() { return {}; }); throw new Error(err.message || "HTTP " + res.status); }
  return res.json();
}
async function ghPut(path, body) {
  var res = await fetch(apiBase() + "/" + path, {
    method: "PUT",
    headers: { Authorization: "token " + cfg.token, Accept: "application/vnd.github+json", "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) { var err = await res.json().catch(function() { return {}; }); throw new Error(err.message || "HTTP " + res.status); }
  return res.json();
}
async function ghDelete(path, sha) {
  var res = await fetch(apiBase() + "/" + path, {
    method: "DELETE",
    headers: { Authorization: "token " + cfg.token, Accept: "application/vnd.github+json", "Content-Type": "application/json" },
    body: JSON.stringify({ message: "Delete " + path.split("/").pop(), sha: sha, branch: cfg.branch }),
  });
  if (!res.ok) { var err = await res.json().catch(function() { return {}; }); throw new Error(err.message || "HTTP " + res.status); }
  return res.json();
}
function fileToBase64(file) {
  return new Promise(function(resolve, reject) {
    var reader = new FileReader();
    reader.onload  = function() { resolve(reader.result.split(",")[1]); };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ── Category chips (with counts) ──────────────────────────────────────────────
function renderCategoryChips() {
  var libCats = [...new Set(assets.map(function(a) { return a.category; }).filter(Boolean))].sort();
  localCategories = localCategories.filter(function(c) {
    return !libCats.map(function(l) { return l.toLowerCase(); }).includes(c.toLowerCase());
  });
  saveLocalCategories();

  // Ensure activeCategory is still valid
  var validCats = ["All", "Recent", "Favorites"].concat(libCats).concat(localCategories);
  if (!validCats.includes(activeCategory)) activeCategory = "All";

  var counts = {};
  assets.forEach(function(a) { counts[a.category] = (counts[a.category] || 0) + 1; });

  var html = "";

  // ── Recent chip
  var recentInLib = recentlyUsed.filter(function(fn) { return assets.some(function(a) { return a.file === fn; }); });
  if (recentInLib.length > 0) {
    html += '<button class="chip chip-special' + (activeCategory === "Recent" ? " active" : "") +
      '" onclick="setCategory(\'Recent\')" title="Recently inserted">⏱ Recent <span class="chip-count">' + recentInLib.length + '</span></button>';
  }

  // ── Favorites chip
  var favInLib = Array.from(favorites).filter(function(fn) { return assets.some(function(a) { return a.file === fn; }); });
  if (favInLib.length > 0) {
    html += '<button class="chip chip-special' + (activeCategory === "Favorites" ? " active" : "") +
      '" onclick="setCategory(\'Favorites\')" title="Favorites">★ Favs <span class="chip-count">' + favInLib.length + '</span></button>';
  }

  // ── All chip
  html += '<button class="chip' + (activeCategory === "All" ? " active" : "") + '" onclick="setCategory(\'All\')">' +
    'All <span class="chip-count">' + assets.length + '</span></button>';

  // ── Library category chips
  html += libCats.map(function(c) {
    var count = counts[c] || 0;
    return '<button class="chip' + (c === activeCategory ? " active" : "") +
      '" onclick="setCategory(\'' + c.replace(/'/g, "\\'") + '\')">' +
      xmlEsc(c) + ' <span class="chip-count">' + count + '</span></button>';
  }).join("");

  // ── Local-only (empty) chips
  html += localCategories.map(function(c) {
    return '<button class="chip chip-empty' + (c === activeCategory ? " active" : "") +
      '" onclick="setCategory(\'' + c.replace(/'/g, "\\'") + '\')">' + xmlEsc(c) + '</button>';
  }).join("");

  document.getElementById("category-chips").innerHTML = html;
}

function setCategory(cat) {
  activeCategory = cat;
  renderCategoryChips();
  renderGrid();
}

// ── Filter / search ───────────────────────────────────────────────────────────
function filterAssets(value) {
  filterText = value.toLowerCase().trim();
  renderGrid();
}

// ── Render grid ───────────────────────────────────────────────────────────────
function renderGrid() {
  var grid    = document.getElementById("icon-grid");
  var empty   = document.getElementById("empty-state");
  var counter = document.getElementById("icon-count");
  counter.textContent = assets.length;

  var visible;
  if (activeCategory === "Recent") {
    // Keep recent order
    var recentFilenames = recentlyUsed.filter(function(fn) { return assets.some(function(a) { return a.file === fn; }); });
    visible = recentFilenames.map(function(fn) { return assets.find(function(a) { return a.file === fn; }); }).filter(Boolean);
  } else if (activeCategory === "Favorites") {
    visible = assets.filter(function(a) { return favorites.has(a.file); });
    visible = getSortedAssets(visible);
  } else {
    visible = assets.slice();
    if (activeCategory !== "All") {
      visible = visible.filter(function(a) { return a.category === activeCategory; });
    }
    visible = getSortedAssets(visible);
  }

  if (filterText) {
    visible = visible.filter(function(a) {
      return a.name.toLowerCase().includes(filterText) || a.file.toLowerCase().includes(filterText);
    });
  }

  if (assets.length === 0) {
    empty.style.display = "flex"; grid.style.display = "none"; grid.innerHTML = ""; return;
  }
  empty.style.display = "none"; grid.style.display = "grid";

  if (visible.length === 0) {
    grid.innerHTML = '<p class="no-results">No assets match "<strong>' + xmlEsc(filterText || activeCategory) + '</strong>"</p>';
    return;
  }

  var base = pagesBase() + "assets/";
  grid.innerHTML = visible.map(function(a) {
    var realIdx = assets.indexOf(a);
    var isSelected = selectedFiles.has(a.file);
    var isFav = favorites.has(a.file);
    return (
      '<div class="icon-card' + (isSelected ? " selected" : "") + '"' +
        ' onclick="' + (selectMode ? "toggleSelectAsset('" + a.file.replace(/'/g, "\\'") + "')" : "insertAsset(" + realIdx + ")") + '"' +
        ' oncontextmenu="showContextMenu(event, ' + realIdx + ')"' +
        ' title="' + (selectMode ? "Click to select" : "Click to insert · Right-click for options") + '">' +
        (selectMode ? '<div class="card-checkbox">' + (isSelected ? "✓" : "") + '</div>' : "") +
        (isFav ? '<div class="card-fav">★</div>' : "") +
        '<div class="icon-preview">' +
          '<img src="' + base + a.file + '" alt="' + xmlEsc(a.name) + '" loading="lazy" />' +
        '</div>' +
        '<div class="icon-name" title="' + xmlEsc(a.name) + '">' + xmlEsc(a.name) + '</div>' +
      '</div>'
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
  document.querySelectorAll(".side-panel, .modal").forEach(function(el) { el.classList.remove("open"); });
  document.getElementById("overlay").classList.remove("show");
  window._pendingFiles = null;
}

// ── Utilities ─────────────────────────────────────────────────────────────────
function toTitleCase(str) { return str.replace(/\b\w/g, function(c) { return c.toUpperCase(); }); }
function xmlEsc(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ── Toast ─────────────────────────────────────────────────────────────────────
var toastTimer;
function toast(msg, type) {
  var el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "toast " + (type || "info") + " show";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function() { el.classList.remove("show"); }, 3200);
}
