const tools = [
  {
    id: "merge",
    label: "Merge PDF",
    endpoint: "/api/merge",
    fields: [{ name: "files", label: "PDFs", type: "file", multiple: true, accept: ".pdf", required: true }],
  },
  {
    id: "split",
    label: "Split PDF",
    endpoint: "/api/split",
    fields: [
      { name: "file", label: "PDF", type: "file", accept: ".pdf", required: true },
      { name: "chunk_size", label: "Pages par fichier", type: "number", value: 1, min: 1, required: true },
    ],
  },
  {
    id: "extract",
    label: "Extract Pages",
    endpoint: "/api/extract",
    fields: [
      { name: "file", label: "PDF", type: "file", accept: ".pdf", required: true },
      { name: "pages", label: "Pages (ex: 1,3-6)", type: "text", placeholder: "1,2,8-10", required: true },
    ],
  },
  {
    id: "rotate",
    label: "Rotate PDF",
    endpoint: "/api/rotate",
    fields: [
      { name: "file", label: "PDF", type: "file", accept: ".pdf", required: true },
      { name: "angle", label: "Angle (90/180/270)", type: "number", value: 90, required: true },
      { name: "pages", label: "Pages (vide = toutes)", type: "text", placeholder: "2,4-8" },
    ],
  },
  {
    id: "watermark",
    label: "Watermark",
    endpoint: "/api/watermark",
    fields: [
      { name: "file", label: "PDF", type: "file", accept: ".pdf", required: true },
      { name: "text", label: "Texte", type: "text", placeholder: "CONFIDENTIEL", required: true },
      { name: "opacity", label: "Opacite (0.1-1)", type: "number", value: 0.15, step: 0.05, min: 0.1, max: 1, required: true },
    ],
  },
  {
    id: "images",
    label: "Images -> PDF",
    endpoint: "/api/images-to-pdf",
    fields: [
      {
        name: "files",
        label: "Images",
        type: "file",
        multiple: true,
        accept: ".png,.jpg,.jpeg,.webp,.bmp",
        required: true,
      },
    ],
  },
  {
    id: "compress",
    label: "Compress PDF",
    endpoint: "/api/compress",
    fields: [
      { name: "file", label: "PDF", type: "file", accept: ".pdf", required: true },
      { name: "level", label: "Niveau (light/balanced/aggressive)", type: "text", value: "balanced", required: true },
    ],
  },
  {
    id: "remove_blank",
    label: "Remove Blank",
    endpoint: "/api/remove-blank",
    fields: [
      { name: "file", label: "PDF", type: "file", accept: ".pdf", required: true },
      { name: "content_threshold", label: "Sensibilite (0-300)", type: "number", value: 80, min: 0, required: true },
    ],
  },
  {
    id: "ocr",
    label: "OCR Text",
    endpoint: "/api/ocr-text",
    fields: [
      { name: "file", label: "PDF ou Image", type: "file", accept: ".pdf,.png,.jpg,.jpeg,.webp,.bmp", required: true },
      { name: "lang", label: "Langues OCR (fra+eng)", type: "text", value: "fra+eng", required: true },
      { name: "min_chars", label: "Seuil texte natif", type: "number", value: 40, min: 0, required: true },
    ],
  },
  {
    id: "video_extract",
    label: "Video -> MP4",
    endpoint: "/api/video-extract",
    fields: [
      { name: "video_url", label: "Lien video", type: "text", placeholder: "https://...", required: true },
      { name: "source", label: "Source (youtube/twitter/tiktok)", type: "text", value: "youtube", required: true },
      { name: "owns_rights", label: "J'ai les droits de telechargement", type: "checkbox", required: true },
    ],
  },
  {
    id: "convert",
    label: "Convert",
    endpoint: "/api/convert",
    fields: [
      { name: "file", label: "Fichier", type: "file", accept: ".pdf,.docx,.xlsx,.pptx,.png,.jpg,.jpeg,.webp,.bmp,.tiff,.tif", required: true },
      {
        name: "mode",
        label: "Mode de conversion",
        type: "select",
        value: "pdf_to_docx",
        required: true,
        options: [
          { value: "pdf_to_docx", label: "PDF -> DOCX" },
          { value: "pdf_to_images", label: "PDF -> Images (ZIP)" },
          { value: "pdf_to_excel", label: "PDF -> Excel (editable)" },
          { value: "image_to_pdf", label: "Image -> PDF" },
          { value: "office_to_pdf", label: "DOCX/XLSX/PPTX -> PDF" },
        ],
      },
    ],
  },
];

const tabsEl = document.getElementById("tabs");
const fieldsEl = document.getElementById("fields");
const formEl = document.getElementById("tool-form");
const runBtn = document.getElementById("run-btn");
const statusEl = document.getElementById("status");
const toolIndicatorEl = document.getElementById("tool-indicator");
const dropzoneEl = document.getElementById("dropzone");
const queueEl = document.getElementById("file-queue");
const clearQueueBtn = document.getElementById("clear-queue-btn");
const metricsEl = document.getElementById("size-metrics");
const apiBaseInput = document.getElementById("api-base");
const saveApiBtn = document.getElementById("save-api-btn");
const barInputEl = document.getElementById("bar-input");
const barOutputEl = document.getElementById("bar-output");

const savedToolId = localStorage.getItem("pdf_nova_active_tool");
let activeTool = tools.find((t) => t.id === savedToolId) || tools[0];
const toolQueues = {};
let lastOutputBytes = null;
let capabilities = { ocr_available: true, ocr_note: "" };
const storedApiBase = (localStorage.getItem("pdf_nova_api_base") || "").trim();
const isLocalHost = ["127.0.0.1", "localhost"].includes(window.location.hostname);
const cloudBackendFallback = "https://pdf-nova-api.onrender.com";
let apiBase = storedApiBase || (isLocalHost ? "" : cloudBackendFallback);
const localBackendFallback = "http://127.0.0.1:8091";
if (apiBaseInput) apiBaseInput.value = apiBase;

function normalizeApiBase(url) {
  const value = (url || "").trim();
  if (!value) return "";
  return value.replace(/\/+$/, "");
}

function apiUrl(path) {
  if (!apiBase) return path;
  return `${apiBase}${path}`;
}

async function tryAutoConnectLocalBackend() {
  if (isLocalHost || apiBase) return false;
  try {
    const resp = await fetch(`${localBackendFallback}/api/health`, { method: "GET" });
    if (!resp.ok) return false;
    const data = await resp.json();
    if (data && data.app === "pdf_nova") {
      apiBase = localBackendFallback;
      localStorage.setItem("pdf_nova_api_base", apiBase);
      if (apiBaseInput) apiBaseInput.value = apiBase;
      setStatus(`Backend local detecte automatiquement: ${apiBase}`);
      return true;
    }
  } catch (_e) {}
  return false;
}

function setStatus(text, isError = false) {
  statusEl.textContent = text || "";
  statusEl.style.color = isError ? "#a80000" : "var(--muted)";
}

function formatBytes(bytes) {
  if (!bytes || bytes <= 0) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  let value = bytes;
  let idx = 0;
  while (value >= 1024 && idx < units.length - 1) {
    value /= 1024;
    idx += 1;
  }
  return `${value.toFixed(value >= 10 || idx === 0 ? 0 : 1)} ${units[idx]}`;
}

function getPrimaryFileField(tool) {
  return tool.fields.find((f) => f.type === "file") || null;
}

function ensureQueue(toolId) {
  if (!toolQueues[toolId]) toolQueues[toolId] = [];
  return toolQueues[toolId];
}

function inputForPrimaryFile() {
  const field = getPrimaryFileField(activeTool);
  if (!field) return null;
  return formEl.querySelector(`[name="${field.name}"]`);
}

function syncQueueToInput() {
  const field = getPrimaryFileField(activeTool);
  const input = inputForPrimaryFile();
  if (!field || !input) return;
  const queue = ensureQueue(activeTool.id);
  const dt = new DataTransfer();
  const files = field.multiple ? queue : queue.slice(0, 1);
  for (const file of files) dt.items.add(file);
  input.files = dt.files;
}

function renderQueue() {
  const queue = ensureQueue(activeTool.id);
  queueEl.innerHTML = "";
  if (!queue.length) {
    const li = document.createElement("li");
    li.textContent = "Aucun fichier dans la queue.";
    queueEl.appendChild(li);
    return;
  }
  queue.forEach((file, index) => {
    const li = document.createElement("li");
    li.innerHTML = `<span>${index + 1}. ${file.name}</span><span>${formatBytes(file.size)}</span>`;
    queueEl.appendChild(li);
  });
}

function renderMetrics(outputBytes = lastOutputBytes) {
  const inputBytes = ensureQueue(activeTool.id).reduce((sum, f) => sum + f.size, 0);
  const ratio = outputBytes && inputBytes ? ((outputBytes / inputBytes) * 100).toFixed(1) : "-";
  metricsEl.innerHTML = `
    <span>Entree: ${formatBytes(inputBytes)}</span>
    <span>Sortie: ${outputBytes ? formatBytes(outputBytes) : "-"}</span>
    <span>Ratio: ${ratio === "-" ? "-" : `${ratio}%`}</span>
  `;
  const maxVal = Math.max(inputBytes, outputBytes || 0, 1);
  if (barInputEl) barInputEl.style.width = `${Math.min(100, (inputBytes / maxVal) * 100)}%`;
  if (barOutputEl) barOutputEl.style.width = `${Math.min(100, ((outputBytes || 0) / maxVal) * 100)}%`;
}

function updateDropzoneHint() {
  const field = getPrimaryFileField(activeTool);
  if (!field) {
    dropzoneEl.innerHTML = "<p>Ce module ne prend pas de fichier.</p>";
    return;
  }
  const mode = field.multiple ? "plusieurs fichiers" : "un seul fichier";
  dropzoneEl.innerHTML = `<p>Glisse-depose ${mode} ici (${field.accept || "tout type"}).</p>`;
}

function isAccepted(file, acceptValue) {
  if (!acceptValue) return true;
  const accepted = acceptValue
    .split(",")
    .map((x) => x.trim().toLowerCase())
    .filter(Boolean);
  if (!accepted.length) return true;
  const name = file.name.toLowerCase();
  return accepted.some((rule) => name.endsWith(rule));
}

function updateQueueFromFiles(fileList) {
  const field = getPrimaryFileField(activeTool);
  if (!field) return;
  let incoming = Array.from(fileList).filter((f) => isAccepted(f, field.accept));
  if (!incoming.length) return;
  if (!field.multiple) incoming = incoming.slice(0, 1);
  toolQueues[activeTool.id] = field.multiple ? incoming : [incoming[0]];
  lastOutputBytes = null;
  syncQueueToInput();
  renderQueue();
  renderMetrics();
}

function renderTabs() {
  tabsEl.innerHTML = "";
  for (const tool of tools) {
    const btn = document.createElement("button");
    btn.type = "button";
    const disabled = tool.id === "ocr" && !capabilities.ocr_available;
    btn.className = `tab ${tool.id === activeTool.id ? "active" : ""} ${disabled ? "disabled" : ""}`;
    btn.textContent = tool.label;
    if (disabled) btn.title = capabilities.ocr_note || "OCR indisponible";
    btn.addEventListener("click", () => {
      if (disabled) {
        setStatus(capabilities.ocr_note || "OCR indisponible.", true);
        return;
      }
      activeTool = tool;
      localStorage.setItem("pdf_nova_active_tool", activeTool.id);
      renderTabs();
      renderFields();
      updateDropzoneHint();
      renderQueue();
      renderMetrics(null);
      setStatus("");
    });
    tabsEl.appendChild(btn);
  }
}

function createInput(field) {
  const wrapper = document.createElement("div");
  wrapper.className = "field";
  const label = document.createElement("label");
  label.textContent = field.label;
  wrapper.appendChild(label);
  const input = document.createElement(field.type === "select" ? "select" : "input");
  if (field.type !== "select") input.type = field.type;
  input.name = field.name;
  if (field.required) input.required = true;
  if (field.accept) input.accept = field.accept;
  if (field.multiple) input.multiple = true;
  if (field.placeholder) input.placeholder = field.placeholder;
  if (field.value !== undefined) input.value = field.value;
  if (field.step !== undefined) input.step = field.step;
  if (field.min !== undefined) input.min = field.min;
  if (field.max !== undefined) input.max = field.max;
  if (field.type === "select" && Array.isArray(field.options)) {
    for (const opt of field.options) {
      const option = document.createElement("option");
      option.value = opt.value;
      option.textContent = opt.label;
      if (field.value === opt.value) option.selected = true;
      input.appendChild(option);
    }
  }
  if (field.type === "checkbox") input.value = "true";
  if (field.type === "file") {
    input.addEventListener("change", () => {
      if (input.files && input.files.length) updateQueueFromFiles(input.files);
    });
  }
  wrapper.appendChild(input);
  return wrapper;
}

function renderFields() {
  fieldsEl.innerHTML = "";
  for (const field of activeTool.fields) fieldsEl.appendChild(createInput(field));
  if (toolIndicatorEl) toolIndicatorEl.textContent = `Mode actif: ${activeTool.label} (${activeTool.endpoint})`;
  runBtn.textContent = `Lancer ${activeTool.label}`;
  syncQueueToInput();
}

function buildFormData(form, tool) {
  const fd = new FormData();
  for (const field of tool.fields) {
    const input = form.querySelector(`[name="${field.name}"]`);
    if (!input) continue;
    if (field.type === "file") {
      if (!input.files || input.files.length === 0) continue;
      if (field.multiple) {
        for (const file of input.files) fd.append(field.name, file);
      } else {
        fd.append(field.name, input.files[0]);
      }
    } else if (field.type === "checkbox") {
      fd.append(field.name, input.checked ? "true" : "false");
    } else {
      fd.append(field.name, input.value);
    }
  }
  return fd;
}

function inferExtension(contentType, toolId) {
  const ct = (contentType || "").toLowerCase();
  if (ct.includes("application/pdf")) return "pdf";
  if (ct.includes("application/vnd.openxmlformats-officedocument.wordprocessingml.document")) return "docx";
  if (ct.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) return "xlsx";
  if (ct.includes("application/zip")) return "zip";
  if (ct.includes("text/plain")) return "txt";
  if (toolId === "convert") return "pdf";
  return "dat";
}

function inferExtensionByTool(toolId) {
  const map = {
    merge: "pdf",
    split: "zip",
    extract: "pdf",
    rotate: "pdf",
    watermark: "pdf",
    images: "pdf",
    compress: "pdf",
    remove_blank: "pdf",
    ocr: "txt",
    video_extract: "mp4",
  };
  return map[toolId] || "dat";
}

function inferConvertExt(form) {
  const mode = (form.querySelector('[name="mode"]')?.value || "").trim().toLowerCase();
  const map = {
    pdf_to_docx: "docx",
    pdf_to_images: "zip",
    pdf_to_excel: "xlsx",
    image_to_pdf: "pdf",
    office_to_pdf: "pdf",
  };
  return map[mode] || "pdf";
}

async function runTool(evt) {
  evt.preventDefault();
  setStatus("Traitement...");
  runBtn.disabled = true;
  syncQueueToInput();
  try {
    const formData = buildFormData(formEl, activeTool);
    const response = await fetch(apiUrl(activeTool.endpoint), { method: "POST", body: formData });
    if (!response.ok) {
      let detail = `${response.status}`;
      try {
        const body = await response.json();
        detail = body.detail || detail;
      } catch (_e) {}
      if (detail.includes("Ajoute au moins 2 PDFs") && activeTool.id !== "merge") {
        detail = `${detail} | Onglet incoherent detecte: recharge la page (Ctrl+F5).`;
      }
      throw new Error(detail);
    }
    const blob = await response.blob();
    lastOutputBytes = blob.size;
    renderMetrics(lastOutputBytes);
    const disposition = response.headers.get("content-disposition") || "";
    const fileMatch = disposition.match(/filename="([^"]+)"/);
    const contentType = response.headers.get("content-type") || "";
    let ext = inferExtension(contentType, activeTool.id);
    if (ext === "dat") {
      ext = activeTool.id === "convert" ? inferConvertExt(formEl) : inferExtensionByTool(activeTool.id);
    }
    const filename = fileMatch ? fileMatch[1] : `pdf_nova_${activeTool.id}.${ext}`;
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = filename;
    link.click();
    URL.revokeObjectURL(url);
    setStatus(`OK -> ${filename}`);
  } catch (error) {
    setStatus(`Erreur: ${error.message}`, true);
  } finally {
    runBtn.disabled = false;
  }
}

dropzoneEl.addEventListener("dragover", (evt) => {
  evt.preventDefault();
  dropzoneEl.classList.add("dragging");
});

dropzoneEl.addEventListener("dragleave", () => {
  dropzoneEl.classList.remove("dragging");
});

dropzoneEl.addEventListener("drop", (evt) => {
  evt.preventDefault();
  dropzoneEl.classList.remove("dragging");
  if (evt.dataTransfer && evt.dataTransfer.files) updateQueueFromFiles(evt.dataTransfer.files);
});

clearQueueBtn.addEventListener("click", () => {
  toolQueues[activeTool.id] = [];
  lastOutputBytes = null;
  syncQueueToInput();
  renderQueue();
  renderMetrics();
  setStatus("Queue videe.");
});

renderTabs();
renderFields();
updateDropzoneHint();
renderQueue();
renderMetrics();
formEl.addEventListener("submit", runTool);

if (saveApiBtn) {
  saveApiBtn.addEventListener("click", () => {
    apiBase = normalizeApiBase(apiBaseInput ? apiBaseInput.value : "");
    localStorage.setItem("pdf_nova_api_base", apiBase);
    setStatus(apiBase ? `API backend configuree: ${apiBase}` : "API backend locale utilisee.");
    fetch(apiUrl("/api/capabilities"))
      .then((r) => r.json())
      .then((data) => {
        capabilities = data || capabilities;
        renderTabs();
      })
      .catch(() => {
        setStatus("Impossible de joindre ce backend API.", true);
      });
  });
}

(async () => {
  await tryAutoConnectLocalBackend();
  fetch(apiUrl("/api/capabilities"))
    .then((r) => r.json())
    .then((data) => {
      capabilities = data || capabilities;
      if (!capabilities.ocr_available && activeTool.id === "ocr") {
        activeTool = tools[0];
        renderFields();
        updateDropzoneHint();
        renderQueue();
        renderMetrics();
      }
      renderTabs();
      if (!capabilities.ocr_available) {
        setStatus(capabilities.ocr_note || "OCR indisponible.");
      }
    })
    .catch(() => {
      setStatus(
        "Backend API indisponible. Renseigne un backend dans le champ 'API backend', puis clique 'Sauver'.",
        true
      );
    });
})();
