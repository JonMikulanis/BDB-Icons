// ── Constants ──────────────────────────────────────────────────────────────
const NS = "http://ppt-icon-library.local/v1";
const ROOT_TAG = "iconLibrary";

// ── State ───────────────────────────────────────────────────────────────────
let icons = [];          // { id, name, svgData }
let filterText = "";

// ── Office Init ─────────────────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    loadIcons();
  } else {
    toast("Open this add-in inside PowerPoint.", "error");
  }
});

// ── Load icons from custom XML embedded in the .pptx ────────────────────────
function loadIcons() {
  Office.context.document.customXmlParts.getByNamespaceAsync(NS, function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      renderGrid();
      return;
    }
    const parts = res.value;
    if (!parts || parts.length === 0) {
      renderGrid();
      return;
    }
    parts[0].getXmlAsync(function (xmlRes) {
      if (xmlRes.status === Office.AsyncResultStatus.Succeeded) {
        parseXml(xmlRes.value);
      }
      renderGrid();
    });
  });
}

// ── Parse XML → icons array ──────────────────────────────────────────────────
function parseXml(xmlString) {
  try {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlString, "text/xml");
    const nodes = doc.querySelectorAll("icon");
    icons = [];
    nodes.forEach(function (el) {
      const encoded = el.getAttribute("data");
      if (!encoded) return;
      icons.push({
        id:      el.getAttribute("id")   || uid(),
        name:    el.getAttribute("name") || "icon",
        svgData: b64decode(encoded)
      });
    });
  } catch (e) {
    icons = [];
  }
}

// ── Serialise icons array → XML string ──────────────────────────────────────
function buildXml() {
  const inner = icons.map(function (ic) {
    return '<icon id="' + ic.id +
           '" name="' + xmlEsc(ic.name) +
           '" data="' + b64encode(ic.svgData) + '"/>';
  }).join("");
  return '<' + ROOT_TAG + ' xmlns="' + NS + '">' + inner + '</' + ROOT_TAG + '>';
}

// ── Persist icons to the .pptx custom XML part ──────────────────────────────
function saveIcons(callback) {
  const xml = buildXml();
  Office.context.document.customXmlParts.getByNamespaceAsync(NS, function (res) {
    if (res.status === Office.AsyncResultStatus.Succeeded && res.value.length > 0) {
      res.value[0].deleteAsync(function () { writeXml(xml, callback); });
    } else {
      writeXml(xml, callback);
    }
  });
}

function writeXml(xml, callback) {
  Office.context.document.customXmlParts.addAsync(xml, function (res) {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      toast("Could not save icons to file.", "error");
    }
    if (typeof callback === "function") callback();
  });
}

// ── Handle SVG file upload ───────────────────────────────────────────────────
function handleUpload(input) {
  const files = Array.from(input.files).filter(function (f) {
    return f.name.toLowerCase().endsWith(".svg");
  });

  if (files.length === 0) {
    toast("Please select .svg files only.", "error");
    input.value = "";
    return;
  }

  let done = 0;
  const added = [];

  files.forEach(function (file) {
    const reader = new FileReader();
    reader.onload = function (e) {
      const svgText = e.target.result;
      // Basic sanity check — must contain <svg
      if (!svgText.includes("<svg") && !svgText.includes("<SVG")) {
        toast(file.name + " doesn't look like a valid SVG.", "error");
      } else {
        const clean = sanitizeSvg(svgText);
        added.push({ id: uid(), name: file.name, svgData: clean });
      }
      done++;
      if (done === files.length) {
        icons = icons.concat(added);
        saveIcons(function () {
          renderGrid();
          toast(added.length + " icon" + (added.length !== 1 ? "s" : "") + " added.", "success");
        });
      }
    };
    reader.onerror = function () {
      done++;
      toast("Could not read " + file.name, "error");
    };
    reader.readAsText(file);
  });

  input.value = "";
}

// ── Insert SVG onto active PowerPoint slide ──────────────────────────────────
function insertIcon(id) {
  const ic = icons.find(function (i) { return i.id === id; });
  if (!ic) return;

  Office.context.document.setSelectedDataAsync(
    ic.svgData,
    { coercionType: Office.CoercionType.XmlSvg },
    function (res) {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        toast("Inserted: " + ic.name.replace(".svg", ""), "success");
      } else {
        toast("Insert failed — check the SVG file.", "error");
      }
    }
  );
}

// ── Delete one icon ──────────────────────────────────────────────────────────
function deleteIcon(event, id) {
  event.stopPropagation();
  icons = icons.filter(function (i) { return i.id !== id; });
  saveIcons(function () {
    renderGrid();
    toast("Icon removed.", "info");
  });
}

// ── Clear all icons ──────────────────────────────────────────────────────────
function clearAll() {
  if (icons.length === 0) return;
  if (!confirm("Remove all " + icons.length + " icon(s) from this file?")) return;
  icons = [];
  saveIcons(function () {
    renderGrid();
    toast("Library cleared.", "info");
  });
}

// ── Filter / search ──────────────────────────────────────────────────────────
function filterIcons(value) {
  filterText = value.toLowerCase().trim();
  renderGrid();
}

// ── Render grid ──────────────────────────────────────────────────────────────
function renderGrid() {
  const grid    = document.getElementById("icon-grid");
  const empty   = document.getElementById("empty-state");
  const counter = document.getElementById("icon-count");

  counter.textContent = icons.length;

  const filtered = filterText
    ? icons.filter(function (ic) { return ic.name.toLowerCase().includes(filterText); })
    : icons;

  if (icons.length === 0) {
    empty.style.display  = "flex";
    grid.style.display   = "none";
    grid.innerHTML = "";
    return;
  }

  empty.style.display = "none";
  grid.style.display  = "grid";

  if (filtered.length === 0) {
    grid.innerHTML = '<p class="no-results">No icons match "<strong>' + filterText + '</strong>"</p>';
    return;
  }

  grid.innerHTML = filtered.map(function (ic) {
    const label = ic.name.replace(/\.svg$/i, "");
    // Inject SVG inline for preview — strip width/height to let CSS control size
    const preview = normaliseSvg(ic.svgData);
    return (
      '<div class="icon-card" onclick="insertIcon(\'' + ic.id + '\')" title="Click to insert">' +
        '<div class="icon-preview">' + preview + "</div>" +
        '<div class="icon-name" title="' + xmlEsc(ic.name) + '">' + xmlEsc(label) + "</div>" +
        '<button class="delete-btn" onclick="deleteIcon(event,\'' + ic.id + '\')" title="Remove">×</button>' +
      "</div>"
    );
  }).join("");
}

// ── SVG helpers ──────────────────────────────────────────────────────────────

// Strip dangerous tags/attrs for safe innerHTML preview
function sanitizeSvg(svgText) {
  // Remove script tags and event handlers — simple but effective for trusted company assets
  return svgText
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/\son\w+="[^"]*"/g, "")
    .replace(/\son\w+='[^']*'/g, "");
}

// Strip fixed w/h so CSS controls the preview size; preserve viewBox
function normaliseSvg(svgText) {
  return svgText
    .replace(/(<svg[^>]*?)\s+width="[^"]*"/i,  "$1")
    .replace(/(<svg[^>]*?)\s+height="[^"]*"/i, "$1")
    .replace(/(<svg[^>]*?)\s+width='[^']*'/i,  "$1")
    .replace(/(<svg[^>]*?)\s+height='[^']*'/i, "$1");
}

// ── Utilities ────────────────────────────────────────────────────────────────
function uid() {
  return "ic_" + Math.random().toString(36).slice(2, 11);
}

function xmlEsc(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function b64encode(str) {
  try {
    return btoa(unescape(encodeURIComponent(str)));
  } catch (e) {
    return btoa(str);
  }
}

function b64decode(str) {
  try {
    return decodeURIComponent(escape(atob(str)));
  } catch (e) {
    return atob(str);
  }
}

// ── Toast ────────────────────────────────────────────────────────────────────
let toastTimer;
function toast(msg, type) {
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "toast " + (type || "info") + " show";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () {
    el.classList.remove("show");
  }, 2600);
}
