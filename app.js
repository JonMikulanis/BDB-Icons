// ── Configuration ────────────────────────────────────────────────────────────
// UPDATE THIS to your GitHub Pages base URL
const BASE_URL = "https://JonMikulanis.github.io/BDB-Icons";
const LIBRARY_URL = BASE_URL + "/library.json";

// ── State ────────────────────────────────────────────────────────────────────
let assets = [];
let filterText = "";

// ── Office Init ──────────────────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    loadLibrary();
  } else {
    toast("Open this add-in inside PowerPoint.", "error");
  }
});

// ── Load asset library from GitHub ──────────────────────────────────────────
function loadLibrary() {
  setStatus("Loading library…");

  fetch(LIBRARY_URL + "?t=" + Date.now())
    .then(function (res) {
      if (!res.ok) throw new Error("library.json not found");
      return res.json();
    })
    .then(function (data) {
      assets = (data.assets || []).map(function (a) {
        return {
          name: a.name,
          file: a.file,
          url:  BASE_URL + "/assets/" + a.file,
          type: a.type || guessType(a.file)
        };
      });
      clearStatus();
      renderGrid();
    })
    .catch(function () {
      clearStatus();
      showSetupState();
    });
}

// ── Insert asset onto active PowerPoint slide ─────────────────────────────────
function insertAsset(index) {
  const asset = assets[index];
  if (!asset) return;

  setStatus("Inserting " + asset.name + "…");

  fetch(asset.url)
    .then(function (res) {
      if (!res.ok) throw new Error("Could not fetch asset");
      if (asset.type === "svg") {
        return res.text().then(function (text) { return { kind: "svg", data: text }; });
      } else {
        return res.blob().then(function (blob) {
          return new Promise(function (resolve, reject) {
            const reader = new FileReader();
            reader.onload = function (e) { resolve({ kind: "image", data: e.target.result }); };
            reader.onerror = reject;
            reader.readAsDataURL(blob);
          });
        });
      }
    })
    .then(function (result) {
      clearStatus();
      if (result.kind === "svg") {
        Office.context.document.setSelectedDataAsync(
          result.data,
          { coercionType: Office.CoercionType.XmlSvg },
          function (res) {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              toast("Inserted: " + asset.name, "success");
            } else {
              toast("SVG insert failed — try re-saving as basic SVG.", "error");
            }
          }
        );
      } else {
        const base64 = result.data.split(",")[1];
        Office.context.document.setSelectedDataAsync(
          base64,
          { coercionType: Office.CoercionType.Image },
          function (res) {
            if (res.status === Office.AsyncResultStatus.Succeeded) {
              toast("Inserted: " + asset.name, "success");
            } else {
              toast("Image insert failed.", "error");
            }
          }
        );
      }
    })
    .catch(function () {
      clearStatus();
      toast("Could not load asset from GitHub.", "error");
    });
}

// ── Filter ───────────────────────────────────────────────────────────────────
function filterAssets(value) {
  filterText = value.toLowerCase().trim();
  renderGrid();
}

// ── Refresh ──────────────────────────────────────────────────────────────────
function refreshLibrary() {
  assets = [];
  renderGrid();
  loadLibrary();
}

// ── Render grid ──────────────────────────────────────────────────────────────
function renderGrid() {
  const grid    = document.getElementById("asset-grid");
  const empty   = document.getElementById("empty-state");
  const counter = document.getElementById("icon-count");
  const setup   = document.getElementById("setup-state");

  if (setup) setup.style.display = "none";
  counter.textContent = assets.length;

  const filtered = filterText
    ? assets.filter(function (a) {
        return a.name.toLowerCase().includes(filterText) ||
               a.file.toLowerCase().includes(filterText);
      })
    : assets;

  if (assets.length === 0) {
    empty.style.display = "flex";
    grid.style.display  = "none";
    grid.innerHTML = "";
    return;
  }

  empty.style.display = "none";
  grid.style.display  = "grid";

  if (filtered.length === 0) {
    grid.innerHTML = '<p class="no-results">No assets match "<strong>' + escHtml(filterText) + '</strong>"</p>';
    return;
  }

  grid.innerHTML = filtered.map(function (asset, i) {
    const label = asset.name || asset.file.replace(/\.[^.]+$/, "");
    const isSvg = asset.type === "svg";
    const imgStyle = isSvg
      ? "width:36px;height:36px;object-fit:contain;"
      : "width:44px;height:36px;object-fit:cover;border-radius:4px;";
    return (
      '<div class="icon-card" onclick="insertAsset(' + assets.indexOf(asset) + ')" title="Click to insert: ' + escHtml(asset.file) + '">' +
        '<div class="icon-preview">' +
          '<img src="' + asset.url + '" style="' + imgStyle + '" onerror="this.style.opacity=0.25" />' +
        '</div>' +
        '<div class="icon-name">' + escHtml(label) + '</div>' +
        '<div class="asset-badge ' + asset.type + '">' + asset.type.toUpperCase() + '</div>' +
      '</div>'
    );
  }).join("");
}

// ── Setup state ──────────────────────────────────────────────────────────────
function showSetupState() {
  const empty = document.getElementById("empty-state");
  const grid  = document.getElementById("asset-grid");
  let   setup = document.getElementById("setup-state");

  empty.style.display = "none";
  grid.style.display  = "none";

  if (!setup) {
    setup = document.createElement("div");
    setup.id = "setup-state";
    setup.className = "empty-state";
    setup.innerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" style="width:48px;height:48px;opacity:0.35">' +
        '<circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/>' +
      '</svg>' +
      '<p class="empty-title">Library not set up yet</p>' +
      '<p class="empty-sub">Upload your assets to GitHub and add a <strong>library.json</strong> file. See the README for instructions.</p>';
    document.querySelector(".content").appendChild(setup);
  }
  setup.style.display = "flex";
}

// ── Helpers ──────────────────────────────────────────────────────────────────
function guessType(filename) {
  const ext = (filename.split(".").pop() || "").toLowerCase();
  return ext === "svg" ? "svg" : "img";
}

function escHtml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

function setStatus(msg) {
  const el = document.getElementById("status-bar");
  if (el) { el.textContent = msg; el.style.display = "block"; }
}

function clearStatus() {
  const el = document.getElementById("status-bar");
  if (el) el.style.display = "none";
}

let toastTimer;
function toast(msg, type) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "toast " + (type || "info") + " show";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { el.classList.remove("show"); }, 2800);
}
