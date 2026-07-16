let currentLibrary = null;
let currentDeck = null;
let activeTab = "meta";
let sourceType = "folder";
let sortState = { key: "imported_at", dir: "desc" };

const $ = (id) => document.getElementById(id);

function payload() {
  return {
    source_type: sourceType,
    library_root: $("libraryRoot").value.trim(),
    output_root: $("outputRoot").value.trim(),
    exclude: $("exclude").value.split(",").map((s) => s.trim()).filter(Boolean),
    rendered_html: $("renderedHtml").checked,
    rendered_standalone_html: $("renderedStandalone").checked,
    standalone_html: $("standaloneHtml").checked,
  };
}

function log(data) {
  $("log").textContent = typeof data === "string" ? data : JSON.stringify(data, null, 2);
}

function setProgress(progress = {}) {
  const percent = Number.isFinite(progress.percent) ? progress.percent : 0;
  $("progressFill").style.width = `${Math.max(0, Math.min(100, percent))}%`;
  $("progressPercent").textContent = `${Math.round(percent)}%`;
  $("progressLabel").textContent = progress.label || "准备中";
}

function setSourceType(nextType) {
  sourceType = nextType;
  document.querySelectorAll(".source-mode").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.source === sourceType);
  });
  const isOnline = sourceType === "online";
  $("sourceLabel").textContent = isOnline ? "在线 PPT 链接" : (sourceType === "file" ? "PPT 文件" : "PPT 文件夹");
  $("libraryRoot").placeholder = isOnline ? "粘贴可直接下载 PPTX 的链接，例如 https://.../deck.pptx" : "";
  $("browseSourceBtn").style.display = isOnline ? "none" : "";
  $("scanBtn").textContent = isOnline ? "扫描链接" : (sourceType === "file" ? "扫描文件" : "扫描文件夹");
  $("scanBtnSecondary").textContent = isOnline ? "仅扫描链接" : (sourceType === "file" ? "仅扫描文件" : "仅扫描");
  $("convertBtn").textContent = isOnline ? "解析转换" : "开始转换";
  $("convertBtnSecondary").textContent = isOnline ? "解析转换" : "开始转换";
}

async function browse(kind) {
  const result = await api("/api/browse", { method: "POST", body: JSON.stringify({ kind }) });
  if (result.cancelled) {
    $("taskHint").textContent = "已取消选择";
    return;
  }
  if (kind === "output_folder") {
    $("outputRoot").value = result.path;
  } else {
    $("libraryRoot").value = result.path;
  }
}

function switchView(name) {
  document.querySelectorAll(".view").forEach((v) => v.classList.remove("active"));
  document.querySelectorAll(".nav").forEach((v) => v.classList.toggle("active", v.dataset.view === name));
  $(`view-${name}`).classList.add("active");
}

async function api(path, options = {}) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || res.statusText);
  return data;
}

async function start(kind) {
  const endpoint = kind === "scan" ? "/api/scan" : "/api/convert";
  const task = await api(endpoint, { method: "POST", body: JSON.stringify(payload()) });
  $("taskStatus").textContent = kind === "scan" ? "Scanning" : "Converting";
  $("taskHint").textContent = "任务已启动，正在等待后台输出";
  setProgress(task.progress || { percent: 0, label: "排队中" });
  log(task);
  pollTask(task.id);
}

async function pollTask(id) {
  const task = await api(`/api/task?id=${encodeURIComponent(id)}`);
  $("taskStatus").textContent = task.status;
  const progress = task.progress || {};
  $("taskHint").textContent = progress.file ? `${task.kind || "task"} · ${progress.status || task.status}` : `${task.kind || "task"} · ${task.started_at || "queued"}`;
  setProgress(progress);
  log({ status: task.status, progress, stdout: task.stdout, stderr: task.stderr, cmd: task.cmd });
  if (task.status === "running" || task.status === "queued") {
    setTimeout(() => pollTask(id), 1200);
    return;
  }
  if (task.library) {
    currentLibrary = task.library;
    renderLibrary(task.library);
    switchView("library");
  }
}

async function refreshLibrary() {
  const out = encodeURIComponent($("outputRoot").value.trim());
  const lib = await api(`/api/library?output_root=${out}`);
  currentLibrary = lib;
  renderLibrary(lib);
  $("taskStatus").textContent = "Ready";
  $("taskHint").textContent = "知识库已刷新";
  switchView("library");
}

function renderLibrary(lib) {
  const m = lib.manifest || {};
  const decks = lib.decks || [];
  const total = m.pptx_count || lib.catalog.length || decks.length || 0;
  const rendered = m.visual_rendered_count ?? decks.filter((d) => d.has_rendered_html).length;
  const failed = m.failed_count ?? (lib.failed || []).length ?? 0;
  const chunks = m.all_chunks_count || 0;
  $("librarySummary").textContent = `${total} files · ${m.processed_count || 0} processed · ${m.skipped_existing_count || 0} skipped · ${m.duplicate_count || 0} duplicate · ${failed} failed`;
  $("statsGrid").innerHTML = [
    ["Total", total, "PPT files"],
    ["Rendered", rendered, "visual HTML"],
    ["Chunks", chunks, "retrieval rows"],
    ["Failed", failed, "needs review"],
  ].map(([label, value, caption]) => `
    <article>
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(String(value))}</strong>
      <small>${escapeHtml(caption)}</small>
    </article>
  `).join("");
  renderRows();
}

function renderRows() {
  const q = $("search").value.trim().toLowerCase();
  const rows = (currentLibrary?.decks || []).filter((d) => {
    const text = `${d.title || ""} ${d.filename || ""} ${d.status || ""} ${d.source_file || ""}`.toLowerCase();
    return !q || text.includes(q);
  }).sort(compareDecks);
  $("deckRows").innerHTML = rows.map((d) => {
    const assets = [
      d.has_rendered_html ? ["visual", "Visual"] : ["muted", "No visual"],
      d.has_chunks ? ["ok", "Chunks"] : ["muted", "No chunks"],
      d.has_markdown ? ["ok", "MD"] : ["muted", "No MD"],
    ];
    return `
      <tr data-dir="${escapeHtml(d.output_dir)}">
        <td><span class="badge ${escapeHtml(d.status || "")}">${escapeHtml(statusLabel(d.status))}</span></td>
        <td class="deck-cell">
          <strong>${escapeHtml(d.title || d.filename || "")}</strong>
          <small>${escapeHtml(d.source_file || d.filename || "")}</small>
        </td>
        <td>
          <div class="asset-pills">
            ${assets.map(([tone, label]) => `<span class="asset ${tone}">${escapeHtml(label)}</span>`).join("")}
          </div>
        </td>
        <td class="date-cell">
          <strong>${escapeHtml(formatDate(d.imported_at))}</strong>
          <small>${escapeHtml(formatTime(d.imported_at))}</small>
        </td>
        <td>${escapeHtml(String(d.slide_count || ""))}</td>
        <td><code>${escapeHtml((d.sha256 || "").slice(0, 12))}</code></td>
      </tr>
    `;
  }).join("");
  document.querySelectorAll("#deckRows tr").forEach((tr) => {
    tr.addEventListener("click", () => loadDeck(tr.dataset.dir));
  });
  updateSortUi();
}

function compareDecks(a, b) {
  const dir = sortState.dir === "asc" ? 1 : -1;
  const key = sortState.key;
  const av = sortValue(a, key);
  const bv = sortValue(b, key);
  if (av < bv) return -1 * dir;
  if (av > bv) return 1 * dir;
  return String(a.title || a.filename || "").localeCompare(String(b.title || b.filename || ""), "zh-CN");
}

function sortValue(deck, key) {
  if (key === "title") return String(deck.title || deck.filename || "").toLowerCase();
  if (key === "status") return statusRank(deck.status);
  if (key === "slide_count") return Number(deck.slide_count || 0);
  if (key === "has_rendered_html") return deck.has_rendered_html ? 1 : 0;
  if (key === "has_chunks") return deck.has_chunks ? 1 : 0;
  if (key === "sha256") return String(deck.sha256 || "");
  if (key === "imported_at") return Date.parse(deck.imported_at || "") || 0;
  return String(deck[key] || "");
}

function statusRank(status) {
  const rank = { failed: 0, pending: 1, processed: 2, skipped_existing: 3, duplicate: 4 };
  return rank[status] ?? 9;
}

function setSort(key) {
  if (sortState.key === key) {
    sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
  } else {
    sortState.key = key;
    sortState.dir = key === "title" || key === "status" ? "asc" : "desc";
  }
  $("sortKey").value = sortState.key;
  renderRows();
}

function updateSortUi() {
  $("sortKey").value = sortState.key;
  $("sortDir").textContent = sortState.dir === "asc" ? "升序" : directionLabel(sortState.key);
  document.querySelectorAll(".th-sort").forEach((btn) => {
    const active = btn.dataset.sort === sortState.key;
    btn.classList.toggle("active", active);
    btn.dataset.dir = active ? sortState.dir : "";
  });
}

function directionLabel(key) {
  if (key === "imported_at") return "新到旧";
  if (key === "slide_count") return "多到少";
  if (key === "has_rendered_html" || key === "has_chunks") return "有优先";
  return "降序";
}

function formatDate(value) {
  const date = parseDate(value);
  if (!date) return "-";
  return new Intl.DateTimeFormat("zh-CN", { month: "2-digit", day: "2-digit" }).format(date);
}

function formatTime(value) {
  const date = parseDate(value);
  if (!date) return "";
  return new Intl.DateTimeFormat("zh-CN", { hour: "2-digit", minute: "2-digit", hour12: false }).format(date);
}

function parseDate(value) {
  const time = Date.parse(value || "");
  return Number.isFinite(time) ? new Date(time) : null;
}

function statusLabel(status) {
  const labels = {
    processed: "processed",
    skipped_existing: "existing",
    duplicate: "duplicate",
    failed: "failed",
    pending: "pending",
  };
  return labels[status] || status || "-";
}

async function loadDeck(deckDir) {
  const detail = await api(`/api/deck?deck_dir=${encodeURIComponent(deckDir)}`);
  currentDeck = detail;
  const metadata = detail.metadata || {};
  $("deckTitle").textContent = metadata.title || "Deck";
  $("deckTitle").title = metadata.title || "Deck";
  $("deckMeta").textContent = `${metadata.slide_count || 0} slides · ${(metadata.sha256 || "").slice(0, 16)}`;
  $("deckSource").textContent = metadata.source_file ? `Source: ${metadata.source_file}` : "";
  $("deckSource").title = metadata.source_file || "";
  renderArtifactActions(detail.links || {});
  const links = detail.links || {};
  $("deckFrame").src = links.rendered_html || links.html || links.standalone_html || "about:blank";
  renderDetail();
  switchView("deck");
}

function renderArtifactActions(links = {}) {
  const warnings = currentDeck?.metadata?.warnings || [];
  const renderWarning = warnings.find((w) => w.type === "rendered_html_skipped");
  const visualItems = [
    artifactLink("rendered_html", "原版式 HTML", links.rendered_html, {
      primary: true,
      disabledReason: renderWarning ? `未生成：${renderWarning.error || renderWarning.message}` : "未生成：转换时没有选择原版式 HTML",
    }),
    artifactLink("html", "知识 HTML", links.html),
    artifactLink("standalone_html", "单文件知识 HTML", links.standalone_html),
    artifactLink("rendered_standalone_html", "单文件原版式 HTML", links.rendered_standalone_html),
  ];
  const dataItems = [
    artifactLink("content", "Markdown", links.content),
    artifactLink("slides", "Slides JSON", links.slides),
    artifactLink("chunks", "Chunks", links.chunks),
    artifactLink("metadata", "Metadata", links.metadata),
  ];
  $("artifactActions").innerHTML = `
    <div class="artifact-group">
      <span>浏览输出</span>
      <div>${visualItems.join("")}</div>
    </div>
    <div class="artifact-group">
      <span>数据文件</span>
      <div>${dataItems.join("")}</div>
    </div>
  `;
}

function artifactLink(key, label, href, options = {}) {
  const classes = ["artifact-link"];
  if (options.primary && href) classes.push("primary-link");
  if (!href) classes.push("disabled");
  const title = href ? label : options.disabledReason || "未生成";
  if (!href) {
    return `<span class="${classes.join(" ")}" title="${escapeHtml(title)}">${escapeHtml(label)}</span>`;
  }
  return `<a class="${classes.join(" ")}" href="${escapeHtml(href)}" target="_blank" rel="noreferrer" title="${escapeHtml(title)}">${escapeHtml(label)}</a>`;
}

function renderDetail() {
  if (!currentDeck) return;
  if (activeTab === "md") {
    $("detailPane").textContent = currentDeck.markdown || "";
  } else {
    $("detailPane").textContent = JSON.stringify(currentDeck.metadata || {}, null, 2);
  }
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[c]));
}

document.querySelectorAll(".nav").forEach((btn) => btn.addEventListener("click", () => switchView(btn.dataset.view)));
document.querySelectorAll(".source-mode").forEach((btn) => btn.addEventListener("click", () => setSourceType(btn.dataset.source)));
document.querySelectorAll(".tab").forEach((btn) => btn.addEventListener("click", () => {
  activeTab = btn.dataset.tab;
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b === btn));
  renderDetail();
}));
$("scanBtn").addEventListener("click", () => start("scan").catch((e) => log(e.message)));
$("scanBtnSecondary").addEventListener("click", () => start("scan").catch((e) => log(e.message)));
$("convertBtn").addEventListener("click", () => start("convert").catch((e) => log(e.message)));
$("convertBtnSecondary").addEventListener("click", () => start("convert").catch((e) => log(e.message)));
$("refreshBtn").addEventListener("click", () => refreshLibrary().catch((e) => log(e.message)));
$("browseSourceBtn").addEventListener("click", () => browse(sourceType === "file" ? "source_file" : "source_folder").catch((e) => log(e.message)));
$("browseOutputBtn").addEventListener("click", () => browse("output_folder").catch((e) => log(e.message)));
$("search").addEventListener("input", renderRows);
$("sortKey").addEventListener("change", (event) => {
  sortState.key = event.target.value;
  sortState.dir = sortState.key === "title" || sortState.key === "status" ? "asc" : "desc";
  renderRows();
});
$("sortDir").addEventListener("click", () => {
  sortState.dir = sortState.dir === "asc" ? "desc" : "asc";
  renderRows();
});
document.querySelectorAll(".th-sort").forEach((btn) => btn.addEventListener("click", () => setSort(btn.dataset.sort)));

setSourceType("folder");
refreshLibrary().catch(() => {});
