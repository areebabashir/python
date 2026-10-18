const el = (id) => document.getElementById(id);
const state = { files: [], jobId: null, timer: null, running: false };

function showBanner(text) {
  const box = el("banner");
  box.hidden = !text;
  box.textContent = text;
}

function renderFiles() {
  const list = el("file-list");
  list.replaceChildren();
  state.files.forEach((file) => {
    const item = document.createElement("li");
    item.textContent = file.webkitRelativePath || file.name;
    list.append(item);
  });
  list.hidden = state.files.length === 0;
  el("start").disabled = state.running || state.files.length === 0;
}

function addFiles(list) {
  const next = Array.from(list || []).filter((file) => file.name.toLowerCase().endsWith(".pdf"));
  const seen = new Set(state.files.map((file) => (file.webkitRelativePath || file.name) + file.size));
  next.forEach((file) => {
    const key = (file.webkitRelativePath || file.name) + file.size;
    if (!seen.has(key)) state.files.push(file);
  });
  renderFiles();
}

function mode() {
  const picked = document.querySelector("input[name=mode]:checked");
  return picked ? picked.value : "normal";
}

function stopPoll() {
  if (state.timer) clearTimeout(state.timer);
  state.timer = null;
}

async function poll() {
  if (!state.jobId) return;
  try {
    const response = await fetch("/api/jobs/" + state.jobId);
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "The extraction stopped.");
    paint(payload);
    if (payload.status === "reading") state.timer = setTimeout(poll, 1000);
    else finish(payload);
  } catch (error) {
    finish({ status: "error", error: error.message, log: [] });
  }
}

function paint(payload) {
  const progress = el("progress");
  const log = el("log");
  progress.hidden = false;
  el("bar").style.width = (payload.percent || 0) + "%";
  el("progress-copy").textContent = payload.total
    ? "Reading page " + payload.done + " of " + payload.total
    : "Opening the PDFs…";
  log.hidden = !payload.log || payload.log.length === 0;
  log.textContent = (payload.log || []).join("\n");
}

function finish(payload) {
  stopPoll();
  state.running = false;
  renderFiles();
  if (payload.status === "error") {
    showBanner(payload.error || "The extraction did not finish.");
    return;
  }
  showBanner("");
  const results = el("results");
  results.hidden = false;
  el("result-title").textContent = payload.row_count + " equipment row" + (payload.row_count === 1 ? "" : "s");
  el("result-note").textContent = "Yellow rows need a look. The Excel file also has All Findings, Duplicates, and Files Log.";
  const link = el("download");
  link.hidden = false;
  link.href = payload.download;
  renderGrid(payload.rows || []);
}

function renderGrid(rows) {
  const table = el("grid");
  table.replaceChildren();
  if (!rows.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.textContent = "No equipment rows were found.";
    row.append(cell);
    table.append(row);
    return;
  }
  const columns = Object.keys(rows[0]);
  const head = document.createElement("tr");
  columns.forEach((name) => {
    const cell = document.createElement("th");
    cell.textContent = name;
    head.append(cell);
  });
  table.append(head);
  rows.forEach((record) => {
    const row = document.createElement("tr");
    if (record.Review) row.className = "review";
    columns.forEach((name) => {
      const cell = document.createElement("td");
      cell.textContent = record[name] || "";
      row.append(cell);
    });
    table.append(row);
  });
}

async function start() {
  if (!state.files.length || state.running) return;
  state.running = true;
  state.jobId = null;
  stopPoll();
  showBanner("");
  el("results").hidden = true;
  el("start").disabled = true;
  el("progress").hidden = false;
  el("bar").style.width = "8%";
  el("progress-copy").textContent = "Uploading…";
  const body = new FormData();
  state.files.forEach((file) => body.append("files", file, file.webkitRelativePath || file.name));
  const taglist = el("taglist").files[0];
  if (taglist) body.append("taglist", taglist, taglist.name);
  body.append("mode", mode());
  try {
    const response = await fetch("/api/jobs", { method: "POST", body });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "The upload did not succeed.");
    state.jobId = payload.id;
    paint(payload);
    poll();
  } catch (error) {
    state.running = false;
    renderFiles();
    showBanner(error.message);
  }
}

function bind() {
  const drop = el("drop");
  const fileInput = el("file-input");
  const folderInput = el("folder-input");
  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      fileInput.click();
    }
  });
  ["dragenter", "dragover"].forEach((name) => drop.addEventListener(name, (event) => {
    event.preventDefault();
    drop.classList.add("is-hot");
  }));
  drop.addEventListener("dragleave", () => drop.classList.remove("is-hot"));
  drop.addEventListener("drop", (event) => {
    event.preventDefault();
    drop.classList.remove("is-hot");
    addFiles(event.dataTransfer.files);
  });
  el("pick-files").addEventListener("click", () => fileInput.click());
  el("pick-folder").addEventListener("click", () => folderInput.click());
  fileInput.addEventListener("change", () => { addFiles(fileInput.files); fileInput.value = ""; });
  folderInput.addEventListener("change", () => { addFiles(folderInput.files); folderInput.value = ""; });
  el("taglist").addEventListener("change", () => {
    const file = el("taglist").files[0];
    el("taglist-name").textContent = file ? file.name : "A list of real tag numbers makes a scanned reading more accurate.";
  });
  el("start").addEventListener("click", start);
}

async function init() {
  bind();
  try {
    const response = await fetch("/api/meta");
    const payload = await response.json();
    el("ocr-note").hidden = !!payload.tesseract;
  } catch (_error) {
    showBanner("The page could not reach the extractor on this computer.");
  }
}

init();
