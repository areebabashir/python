const state = {
  files: [],
  jobId: null,
  poll: null,
  meta: null,
  resultKey: "",
  phase: "pick",
};

const el = (id) => document.getElementById(id);

const statusLine = el("status-line");
const banner = el("banner");
const fileList = el("file-list");
const skipNote = el("skip-note");
const analyseBtn = el("analyse-btn");
const extractBtn = el("extract-btn");
const searchPanel = el("search-panel");
const progress = el("progress");
const barFill = el("bar-fill");
const progressCopy = el("progress-copy");
const results = el("results");
const tagsInput = el("tags");

function setStatus(text) {
  statusLine.textContent = text;
}

function showBanner(text) {
  if (!text) {
    banner.hidden = true;
    banner.textContent = "";
    return;
  }
  banner.hidden = false;
  banner.textContent = text;
}

function accepted(name) {
  const lower = name.toLowerCase();
  return lower.endsWith(".pdf") || lower.endsWith(".zip");
}

function formatBytes(size) {
  if (size < 1024) return size + " B";
  if (size < 1048576) return (size / 1024).toFixed(1) + " KB";
  return (size / 1048576).toFixed(1) + " MB";
}

function formatCount(n) {
  return Number(n || 0).toLocaleString();
}

function addFiles(fileListLike, fallbackPath) {
  const incoming = Array.from(fileListLike || []);
  let skipped = 0;
  for (const file of incoming) {
    const rel = (file.webkitRelativePath || file.relPath || fallbackPath || file.name || "").replace(/\\/g, "/");
    if (!accepted(rel)) {
      skipped += 1;
      continue;
    }
    const key = rel + ":" + file.size;
    if (state.files.some((item) => item.key === key)) continue;
    state.files.push({ key, file, rel });
  }
  if (skipped) {
    skipNote.hidden = false;
    skipNote.textContent = skipped + " file(s) were left out because they are not PDFs or ZIPs.";
  }
  renderFiles();
}

function renderFiles() {
  const busy = Boolean(state.jobId);
  analyseBtn.hidden = busy;
  el("pick-files").disabled = busy;
  el("pick-folder").disabled = busy;
  el("drop").style.pointerEvents = busy ? "none" : "";
  fileList.replaceChildren();
  if (!state.files.length) {
    fileList.hidden = true;
    analyseBtn.disabled = true;
    return;
  }
  fileList.hidden = false;
  analyseBtn.disabled = busy;
  for (const item of state.files) {
    const li = document.createElement("li");
    const name = document.createElement("span");
    name.className = "name";
    name.textContent = item.rel;
    const side = document.createElement("span");
    side.textContent = formatBytes(item.file.size);
    li.append(name, side);
    if (!state.jobId) {
      const remove = document.createElement("button");
      remove.type = "button";
      remove.textContent = "Remove";
      remove.addEventListener("click", () => {
        state.files = state.files.filter((other) => other !== item);
        renderFiles();
      });
      li.append(remove);
    }
    fileList.append(li);
  }
}

function renderAnalysed(analysed) {
  fileList.hidden = false;
  fileList.replaceChildren();
  for (const item of analysed) {
    const li = document.createElement("li");
    const name = document.createElement("span");
    name.className = "name";
    name.textContent = item.name;
    const side = document.createElement("span");
    side.textContent = formatCount(item.pages) + " pages · " + formatCount(item.items) + " items";
    li.append(name, side);
    fileList.append(li);
  }
}

function refreshExtract() {
  extractBtn.disabled = state.phase !== "ready" || !tagsInput.value.trim();
}

function setSearchInputs(locked) {
  searchPanel.classList.toggle("is-locked", locked);
  tagsInput.disabled = locked;
  for (const input of searchPanel.querySelectorAll("input")) {
    if (input !== tagsInput) input.disabled = locked || input.dataset.required === "1";
  }
}

function lockSearch(locked) {
  state.phase = locked ? "pick" : "ready";
  setSearchInputs(locked);
  refreshExtract();
}

function openSearchDraft() {
  setSearchInputs(false);
  state.phase = "analyse";
  refreshExtract();
}

function selectedColumns() {
  return ["Query", "Tag Number", "Model or Item Code", "Serial No", "Reference File", "File Name", "Page"];
}

function renderStats(stats) {
  const box = el("stats");
  box.replaceChildren();
  if (!stats) {
    box.hidden = true;
    return;
  }
  box.hidden = false;
  const cards = [
    ["Files", stats.files],
    ["Pages", stats.pages],
  ];
  for (const kind of stats.kinds) cards.push([kind.name, kind.count]);
  for (const [label, count] of cards) {
    const card = document.createElement("div");
    card.className = "stat";
    const strong = document.createElement("b");
    strong.textContent = formatCount(count);
    const span = document.createElement("span");
    span.textContent = label;
    card.append(strong, span);
    box.append(card);
  }
}

function stopPoll() {
  if (state.poll) {
    clearInterval(state.poll);
    state.poll = null;
  }
}

function startPoll() {
  stopPoll();
  const tick = async () => {
    if (!state.jobId) return;
    try {
      const response = await fetch("/api/jobs/" + state.jobId);
      const body = await response.json();
      if (!response.ok) throw new Error(body.error || "The session ended.");
      applyJob(body);
    } catch (error) {
      stopPoll();
      showBanner(error.message);
    }
  };
  tick();
  state.poll = setInterval(tick, 800);
}

function applyJob(job) {
  if (job.status === "analyzing" || job.status === "searching") {
    showBanner("");
    showProgress(job);
    if (job.status === "analyzing") openSearchDraft();
    if (job.status === "searching") {
      state.phase = "search";
      extractBtn.disabled = true;
      extractBtn.textContent = "Extracting…";
    }
    return;
  }
  hideProgress();
  extractBtn.textContent = "Extract to Excel";
  if (job.status === "cancelled") {
    stopPoll();
    state.jobId = null;
    hideProgress();
    lockSearch(true);
    renderFiles();
    setStatus("Reading was cancelled. You can start again.");
    return;
  }
  if (job.status === "error") {
    stopPoll();
    state.jobId = null;
    hideProgress();
    el("reset-btn").hidden = true;
    lockSearch(true);
    showBanner(job.error || "The drawings could not be read.");
    setStatus("Choose the drawings to read.");
    renderFiles();
    return;
  }
  if (job.status === "ready") {
    stopPoll();
    hideProgress();
    showBanner("");
    lockSearch(false);
    renderStats(job.stats);
    if (job.stats && job.stats.file_list) renderAnalysed(job.stats.file_list);
    const pages = job.stats ? formatCount(job.stats.pages) : "0";
    const files = job.stats ? formatCount(job.stats.files) : "0";
    setStatus("Ready. " + files + " file(s), " + pages + " page(s). Search as many times as you need.");
    el("search-error").hidden = !job.search_error;
    el("search-error").textContent = job.search_error || "";
    if (job.result) showResult(job.result);
    else {
      results.hidden = true;
      state.resultKey = "";
    }
  }
}

function showProgress(job) {
  const info = job.progress || {};
  progress.hidden = false;
  analyseBtn.disabled = true;
  el("ocr-field").disabled = true;
  const pages = Number(info.pages || 0);
  const page = Number(info.page || 0);
  const fileCount = Number(info.file_count || 0);
  const fileIndex = Number(info.file_index || 0);
  let fraction = 0;
  if (info.phase === "search") {
    barFill.classList.add("indet");
    barFill.style.width = "";
  } else if (fileCount && pages) {
    barFill.classList.remove("indet");
    fraction = ((fileIndex - 1) + page / pages) / fileCount;
    barFill.style.width = Math.max(0, Math.min(1, fraction)) * 100 + "%";
  } else {
    barFill.classList.add("indet");
    barFill.style.width = "";
  }
  if (info.phase === "search") {
    progressCopy.textContent = "Matching tags and writing the workbook.";
    setStatus("Writing the workbook.");
    return;
  }
  const name = info.file_name ? " · " + info.file_name : "";
  const pageBit = pages ? " page " + page + "/" + pages : "";
  const countBit = fileCount ? " file " + fileIndex + "/" + fileCount : "";
  const itemBit = info.items ? " · " + formatCount(info.items) + " items" : "";
  const etaBit = info.eta ? " · about " + info.eta + "s left on this file" : "";
  const cached = info.cached ? " · loaded from a previous reading" : "";
  progressCopy.textContent = (info.message || "Reading drawings") + countBit + pageBit + name + itemBit + etaBit + cached;
  setStatus("Reading drawings" + (fileCount ? " (" + fileIndex + " of " + fileCount + ")" : "") + ".");
}

function hideProgress() {
  progress.hidden = true;
  barFill.classList.remove("indet");
  el("ocr-field").disabled = Boolean(state.jobId);
}

function showResult(result) {
  const key = (result.excel_name || "none") + ":" + result.row_count + ":" + (result.totals || []).map((item) => item.Query + item.Matches).join("|");
  results.hidden = false;
  const lead = el("result-lead");
  if (!result.row_count) {
    lead.textContent = "Nothing matched. Try another spelling, or a wildcard such as *PIT*.";
  } else if (result.truncated) {
    lead.textContent = "Showing the first " + formatCount(result.rows.length) + " of " + formatCount(result.row_count) + " rows. The Excel file has every row.";
  } else {
    lead.textContent = formatCount(result.row_count) + " row(s) in this search.";
  }
  const download = el("download-link");
  if (result.download) {
    download.hidden = false;
    download.href = result.download;
    download.setAttribute("download", result.excel_name || "extracted_tags.xlsx");
  } else {
    download.hidden = true;
  }
  renderTotals(result.totals || []);
  renderSummary(result.summary || []);
  renderData(result.columns || [], result.rows || []);
  if (key !== state.resultKey) {
    state.resultKey = key;
    const motion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    results.scrollIntoView({ behavior: motion ? "auto" : "smooth", block: "start" });
  }
}

function renderTotals(totals) {
  const box = el("totals");
  box.replaceChildren();
  for (const item of totals) {
    const card = document.createElement("article");
    card.className = "total" + (item.Status === "FOUND" ? "" : " is-miss");
    const title = document.createElement("b");
    title.textContent = item.Query;
    const detail = document.createElement("div");
    detail.textContent = formatCount(item.Matches) + " matches in " + formatCount(item.Files) + " file(s)";
    const flag = document.createElement("em");
    flag.textContent = item.Status;
    card.append(title, detail, flag);
    box.append(card);
  }
}

function numericColumn(name) {
  return ["Matches", "Pages", "Occurrences", "Files"].includes(name) || name.includes("(pt)");
}

function fillTable(table, headers, records) {
  table.replaceChildren();
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (const header of headers) {
    const cell = document.createElement("th");
    cell.textContent = header;
    if (numericColumn(header)) cell.className = "num";
    headRow.append(cell);
  }
  thead.append(headRow);
  const tbody = document.createElement("tbody");
  for (const record of records) {
    const row = document.createElement("tr");
    for (const header of headers) {
      const cell = document.createElement("td");
      cell.textContent = record[header] == null ? "" : String(record[header]);
      if (numericColumn(header)) cell.className = "num";
      row.append(cell);
    }
    tbody.append(row);
  }
  table.append(thead, tbody);
}

function renderSummary(summary) {
  const wrap = el("summary-wrap");
  if (!summary.length) {
    wrap.hidden = true;
    return;
  }
  wrap.hidden = false;
  fillTable(el("summary-table"), ["Query", "File Name", "Matches", "Pages", "Status", "Page list"], summary);
}

function renderData(columns, rows) {
  const wrap = el("data-wrap");
  if (!rows.length) {
    wrap.hidden = true;
    return;
  }
  wrap.hidden = false;
  fillTable(el("data-table"), columns, rows);
}

async function onDrop(event) {
  event.preventDefault();
  el("drop").classList.remove("is-hot");
  if (state.jobId) return;
  const items = Array.from(event.dataTransfer.items || []);
  const files = [];
  if (items.some((item) => item.webkitGetAsEntry)) {
    for (const item of items) {
      const entry = item.webkitGetAsEntry && item.webkitGetAsEntry();
      if (!entry) continue;
      files.push(...await collectEntry(entry, ""));
    }
  } else {
    addFiles(event.dataTransfer.files);
    return;
  }
  addFiles(files);
}

function collectEntry(entry, prefix) {
  return new Promise((resolve) => {
    if (entry.isFile) {
      entry.file((file) => {
        const rel = prefix + entry.name;
        try {
          Object.defineProperty(file, "webkitRelativePath", { value: rel, configurable: true });
        } catch (_err) {
          file.relPath = rel;
        }
        resolve([file]);
      }, () => resolve([]));
      return;
    }
    if (!entry.isDirectory) {
      resolve([]);
      return;
    }
    const reader = entry.createReader();
    const all = [];
    const read = () => {
      reader.readEntries(async (batch) => {
        if (!batch.length) {
          const nested = [];
          for (const child of all) nested.push(...await collectEntry(child, prefix + entry.name + "/"));
          resolve(nested);
          return;
        }
        all.push(...batch);
        read();
      }, () => resolve([]));
    };
    read();
  });
}

async function analyse() {
  if (!state.files.length || state.jobId) return;
  showBanner("");
  results.hidden = true;
  state.resultKey = "";
  const body = new FormData();
  const paths = [];
  for (const item of state.files) {
    body.append("files", item.file, item.file.name);
    paths.push(item.rel);
  }
  body.append("paths", JSON.stringify(paths));
  body.append("ocr", document.querySelector('input[name="ocr"]:checked').value);
  analyseBtn.disabled = true;
  setStatus("Uploading drawings.");
  progress.hidden = false;
  barFill.classList.add("indet");
  progressCopy.textContent = "Uploading…";
  try {
    const payload = await upload(body);
    state.jobId = payload.id;
    el("reset-btn").hidden = false;
    renderFiles();
    if (payload.skipped && payload.skipped.length) {
      skipNote.hidden = false;
      skipNote.textContent = "Left out: " + payload.skipped.join(", ");
    }
    startPoll();
  } catch (error) {
    hideProgress();
    analyseBtn.disabled = false;
    showBanner(error.message);
    setStatus("Choose the drawings to read.");
  }
}

function upload(body) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/jobs");
    xhr.upload.onprogress = (event) => {
      if (!event.lengthComputable) return;
      barFill.classList.remove("indet");
      barFill.style.width = (event.loaded / event.total) * 100 + "%";
      progressCopy.textContent = "Uploading " + Math.round((event.loaded / event.total) * 100) + "%";
    };
    xhr.onload = () => {
      let payload = {};
      try { payload = JSON.parse(xhr.responseText || "{}"); } catch (_error) { payload = {}; }
      if (xhr.status >= 200 && xhr.status < 300) resolve(payload);
      else reject(new Error(payload.error || "The upload did not succeed."));
    };
    xhr.onerror = () => reject(new Error("Could not reach the app on this computer."));
    xhr.send(body);
  });
}

async function extract() {
  const tags = tagsInput.value.trim();
  if (state.phase !== "ready" || !state.jobId || !tags) return;
  el("search-error").hidden = true;
  state.phase = "search";
  extractBtn.disabled = true;
  extractBtn.textContent = "Extracting…";
  try {
    const response = await fetch("/api/jobs/" + state.jobId + "/search", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ tags, columns: selectedColumns() }),
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "The search did not start.");
    results.hidden = true;
    state.resultKey = "";
    startPoll();
  } catch (error) {
    state.phase = "ready";
    refreshExtract();
    extractBtn.textContent = "Extract to Excel";
    el("search-error").hidden = false;
    el("search-error").textContent = error.message;
  }
}

async function resetJob() {
  if (state.jobId) {
    await fetch("/api/jobs/" + state.jobId + "/cancel", { method: "POST" }).catch(() => {});
  }
  stopPoll();
  state.jobId = null;
  state.files = [];
  state.resultKey = "";
  hideProgress();
  results.hidden = true;
  el("stats").hidden = true;
  el("reset-btn").hidden = true;
  skipNote.hidden = true;
  showBanner("");
  lockSearch(true);
  renderFiles();
  setStatus("Choose the drawings to read.");
}

function bind() {
  const drop = el("drop");
  const fileInput = el("file-input");
  const folderInput = el("folder-input");
  drop.addEventListener("click", () => {
    if (!state.jobId) fileInput.click();
  });
  drop.addEventListener("keydown", (event) => {
    if (state.jobId) return;
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      fileInput.click();
    }
  });
  ["dragenter", "dragover"].forEach((name) => {
    drop.addEventListener(name, (event) => {
      event.preventDefault();
      if (!state.jobId) drop.classList.add("is-hot");
    });
  });
  drop.addEventListener("dragleave", () => drop.classList.remove("is-hot"));
  drop.addEventListener("drop", onDrop);
  window.addEventListener("dragover", (event) => event.preventDefault());
  window.addEventListener("drop", (event) => event.preventDefault());
  el("pick-files").addEventListener("click", () => fileInput.click());
  el("pick-folder").addEventListener("click", () => folderInput.click());
  fileInput.addEventListener("change", () => {
    addFiles(fileInput.files);
    fileInput.value = "";
  });
  folderInput.addEventListener("change", () => {
    addFiles(folderInput.files);
    folderInput.value = "";
  });
  analyseBtn.addEventListener("click", analyse);
  extractBtn.addEventListener("click", extract);
  el("reset-btn").addEventListener("click", resetJob);
  el("cancel-btn").addEventListener("click", resetJob);
  tagsInput.addEventListener("input", refreshExtract);
}

async function init() {
  bind();
  lockSearch(true);
  try {
    const response = await fetch("/api/meta");
    state.meta = await response.json();
    lockSearch(true);
  } catch (_error) {
    showBanner("The page could not load its settings from the local app.");
  }
}

init();
