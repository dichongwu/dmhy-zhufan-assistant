const elements = {
  headerStatus: document.querySelector("#headerStatus"),
  updateButton: document.querySelector("#updateButton"),
  runStatus: document.querySelector("#runStatus"),
  statusMessage: document.querySelector("#statusMessage"),
  statusMeta: document.querySelector("#statusMeta"),
  progressBar: document.querySelector("#progressBar"),
  errorMessage: document.querySelector("#errorMessage"),
  downloadButton: document.querySelector("#downloadButton"),
  recentPresets: document.querySelector("#recentPresets"),
  recentDaysInput: document.querySelector("#recentDaysInput"),
  recentSearchInput: document.querySelector("#recentSearchInput"),
  siteSearchToggle: document.querySelector("#siteSearchToggle"),
  recentRefresh: document.querySelector("#recentRefresh"),
  recentGroupFilter: document.querySelector('[aria-label="近期字幕组筛选"]'),
  recentCount: document.querySelector("#recentCount"),
  recentRows: document.querySelector("#recentRows"),
  recentEmpty: document.querySelector("#recentEmpty"),
  recentMeta: document.querySelector("#recentMeta"),
  recentPagination: document.querySelector("#recentPagination"),
  recentPrev: document.querySelector("#recentPrev"),
  recentNext: document.querySelector("#recentNext"),
  recentPageInfo: document.querySelector("#recentPageInfo"),
  copyRecentButton: document.querySelector("#copyRecentButton"),
  followsCount: document.querySelector("#followsCount"),
  followsRows: document.querySelector("#followsRows"),
  followsEmpty: document.querySelector("#followsEmpty"),
  followsColumns: document.querySelector("#followsColumns"),
  followsDoneCol: document.querySelector("#followsDoneCol"),
  followsDoneCopy: document.querySelector("#followsDoneCopy"),
  followsAllCol: document.querySelector("#followsAllCol"),
  followsPresets: document.querySelector("#followsPresets"),
  followsDoneLabel: document.querySelector("#followsDoneLabel"),
  copyAllFollowsButton: document.querySelector("#copyAllFollowsButton"),
  versionButton: document.querySelector("#versionButton"),
  versionModal: document.querySelector("#versionModal"),
  versionClose: document.querySelector("#versionClose"),
  versionCurrent: document.querySelector("#versionCurrent"),
  versionEntries: document.querySelector("#versionEntries"),
  logOutput: document.querySelector("#logOutput"),
  toast: document.querySelector("#toast"),
};

let currentRecent = [];
let followRecent = [];
let recentDays = 7;
let followsDays = 7;
let recentPage = 1;
let recentRequestId = 0;
let recentQuery = "";
let recentSearchTimer;
const RECENT_PAGE_SIZE = 30;
const selectedRecent = new Set();
const selectedFollows = new Set();
let currentFollowItems = [];
let currentDoneItems = [];
const FOLLOWS_STORAGE_KEY = "dmhy.follows";
const SITE_SEARCH_KEY = "dmhy.siteSearch";
let siteSearch = false;
try {
  siteSearch = localStorage.getItem(SITE_SEARCH_KEY) === "1";
} catch {
  // storage may be unavailable
}
let lastStatus = "";
let toastTimer;
const COPIED_STORAGE_KEY = "dmhy.copiedMagnets";

function loadFollows() {
  try {
    const value = JSON.parse(localStorage.getItem(FOLLOWS_STORAGE_KEY) || "[]");
    return new Set(Array.isArray(value) ? value : []);
  } catch {
    return new Set();
  }
}

let follows = loadFollows();

function saveFollows() {
  try {
    localStorage.setItem(FOLLOWS_STORAGE_KEY, JSON.stringify([...follows]));
  } catch {
    // storage may be unavailable
  }
}

function loadCopiedMagnets() {
  try {
    const value = JSON.parse(localStorage.getItem(COPIED_STORAGE_KEY) || "[]");
    return new Set(Array.isArray(value) ? value : []);
  } catch {
    return new Set();
  }
}

let copiedMagnets = loadCopiedMagnets();

function saveCopiedMagnets() {
  try {
    localStorage.setItem(COPIED_STORAGE_KEY, JSON.stringify([...copiedMagnets]));
  } catch {
    // storage may be unavailable; status still updates for this session
  }
}

function markCopied(magnet) {
  copiedMagnets.add(magnet);
  saveCopiedMagnets();
  renderRecent();
}

function showToast(message) {
  elements.toast.textContent = message;
  elements.toast.classList.add("visible");
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => elements.toast.classList.remove("visible"), 1800);
}

function statusMeta(state) {
  if (state.status === "running") {
    const labels = state.groups.map((id) => id === "7acg" ? "7³ACG" : "LoliHouse").join("、");
    return `${labels} · 最近 ${state.days} 天`;
  }
  if (state.finished_at) {
    return `完成于 ${state.finished_at.slice(11, 16)}`;
  }
  if (state.workbook) {
    return state.workbook.split("/").pop();
  }
  return "等待下一次更新";
}

function extractEncoding(title) {
  const text = String(title || "");
  const rules = [
    [/av1/i, "AV1"],
    [/x265|hevc|h\.?265/i, "X265"],
    [/x264|avc|h\.?264/i, "X264"],
    [/vp9/i, "VP9"],
    [/vp8/i, "VP8"],
  ];
  for (const [pattern, label] of rules) {
    if (pattern.test(text)) return label;
  }
  return "—";
}

function extractLanguage(title) {
  const text = String(title || "");
  const hasSimplified = /简/.test(text);
  const hasTraditional = /繁/.test(text);
  const hasJapanese =
    /简\s*繁\s*日|繁\s*简\s*日|简\s*日|繁\s*日|日(?:语|文|字)/.test(text);
  if (hasSimplified) return "简体";
  if (hasTraditional) return "繁体";
  if (hasJapanese) return "日文";
  return "无字幕";
}

function createResultRow(item, selectedSet, onToggle) {
  const row = document.createElement("div");
  row.className = "result-row";
  row.setAttribute("role", "row");

  const selectCell = document.createElement("span");
  selectCell.className = "select-cell";
  const selectBox = document.createElement("input");
  selectBox.type = "checkbox";
  selectBox.checked = selectedSet.has(item.magnet);
  selectBox.addEventListener("change", () => {
    if (selectBox.checked) {
      selectedSet.add(item.magnet);
    } else {
      selectedSet.delete(item.magnet);
    }
    if (onToggle) onToggle();
  });
  selectCell.append(selectBox);

  const badge = document.createElement("span");
  badge.className = "group-badge";
  badge.textContent = item.group;
  badge.title = item.group;

  const main = document.createElement("div");
  main.className = "release-main";
  const series = document.createElement("strong");
  series.textContent = item.series;
  series.title = item.title || item.series;
  const episode = document.createElement("small");
  episode.textContent = `${item.episode} · ${item.category || "未分类"}`;
  main.append(series, episode);

  const encoding = document.createElement("span");
  encoding.className = "release-encoding";
  encoding.textContent = extractEncoding(item.title);

  const language = document.createElement("span");
  language.className = "release-language";
  language.textContent = extractLanguage(item.title);

  const date = document.createElement("span");
  date.className = "release-date";
  date.textContent = item.published_at;

  const size = document.createElement("span");
  size.className = "release-size";
  size.textContent = item.size || "—";

  const status = document.createElement("span");
  const used = copiedMagnets.has(item.magnet);
  status.className = `copy-status${used ? " used" : ""}`;
  status.textContent = used ? "已使用" : "未使用";

  const actions = document.createElement("div");
  actions.className = "row-actions";
  const copyButton = document.createElement("button");
  copyButton.className = "row-action";
  copyButton.type = "button";
  copyButton.textContent = "复制磁链";
  copyButton.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(item.magnet);
      markCopied(item.magnet);
      showToast("磁链已复制");
    } catch {
      const area = document.createElement("textarea");
      area.value = item.magnet;
      document.body.append(area);
      area.select();
      document.execCommand("copy");
      area.remove();
      markCopied(item.magnet);
      showToast("磁链已复制");
    }
  });
  const detailLink = document.createElement("a");
  detailLink.className = "row-action";
  detailLink.href = item.detail_url;
  detailLink.target = "_blank";
  detailLink.rel = "noreferrer";
  detailLink.textContent = "详情页";
  const followButton = document.createElement("button");
  followButton.className = "row-action";
  followButton.type = "button";
  followButton.textContent = follows.has(item.series) ? "已追番" : "追番";
  followButton.classList.toggle("follow-active", follows.has(item.series));
  followButton.addEventListener("click", () => toggleFollow(item, followButton));
  actions.append(copyButton, detailLink, followButton);

  row.append(selectCell, badge, main, encoding, language, date, size, status, actions);
  return row;
}

function toggleFollow(item, button) {
  const series = item.series;
  if (follows.has(series)) {
    follows.delete(series);
    showToast("已取消追番");
  } else {
    follows.add(series);
    showToast("已加入追番");
  }
  saveFollows();
  if (button) {
    button.textContent = follows.has(series) ? "已追番" : "追番";
    button.classList.toggle("follow-active", follows.has(series));
  }
  renderFollows();
}

function renderFollows() {
  const source = followRecent.length ? followRecent : currentRecent;
  const now = new Date();
  const cutoff = new Date(now.getTime() - followsDays * 24 * 60 * 60 * 1000);
  const inWindow = (item) => {
    const parsed = new Date((item.published_at || "").replace(" ", "T"));
    return !Number.isNaN(parsed.getTime()) && parsed >= cutoff;
  };
  const within = source.filter(inWindow);

  // 追番更新中：所选“几日内”窗口内，每个追番作品最新一条未复制发布
  const latestBySeries = new Map();
  const completedBySeries = new Map();
  for (const item of within) {
    if (!follows.has(item.series)) continue;
    const target = copiedMagnets.has(item.magnet) ? completedBySeries : latestBySeries;
    const current = target.get(item.series);
    if (!current || item.published_at > current.published_at) {
      target.set(item.series, item);
    }
  }
  const items = [...latestBySeries.values()].sort(
    (a, b) => (b.published_at > a.published_at ? 1 : -1)
  );
  const followsTable = elements.followsRows.closest(".result-table");
  if (items.length > 0) {
    currentFollowItems = items;
    if (followsTable) followsTable.hidden = false;
    elements.followsRows.replaceChildren(
      ...items.map((item) => createResultRow(item, selectedFollows, updateSelectionButtons))
    );
    elements.followsCount.textContent = String(items.length);
    elements.followsEmpty.hidden = true;
    elements.followsColumns.hidden = true;
    return;
  }
  if (followsTable) followsTable.hidden = true;

  const done3d = [...completedBySeries.values()]
    .filter(inWindow)
    .sort((a, b) => (b.published_at > a.published_at ? 1 : -1));
  currentDoneItems = done3d;

  // 已追番：全部追番作品，不受天数 / 字幕组等任何筛选影响（每个作品取最新一条）
  const allBySeries = new Map();
  for (const item of source) {
    if (!follows.has(item.series)) continue;
    const current = allBySeries.get(item.series);
    if (!current || item.published_at > current.published_at) {
      allBySeries.set(item.series, item);
    }
  }
  const allItems = [...allBySeries.values()].sort(
    (a, b) => (b.published_at > a.published_at ? 1 : -1)
  );

  // 复制所有磁链仍按所选“几日内”：窗口内每个追番作品最新一条
  const windowLatest = new Map();
  for (const item of within) {
    if (!follows.has(item.series)) continue;
    const current = windowLatest.get(item.series);
    if (!current || item.published_at > current.published_at) {
      windowLatest.set(item.series, item);
    }
  }
  currentFollowItems = [...windowLatest.values()];

  const showColumns = done3d.length > 0 || allItems.length > 0;
  document.querySelectorAll(".recent-panel .filter-hideable").forEach((group) => {
    group.hidden = showColumns;
  });
  const renderColItem = (item, compact) => {
    const row = document.createElement("div");
    row.className = compact ? "follow-done-item compact" : "follow-done-item";
    const main = document.createElement("div");
    main.className = "follow-done-main";
    const series = document.createElement("strong");
    series.textContent = item.series;
    series.title = item.title || item.series;
    main.append(series);
    if (!compact) {
      const meta = document.createElement("small");
      meta.textContent = `${item.episode} · ${item.published_at}`;
      main.append(meta);
      const magnetLink = document.createElement("button");
      magnetLink.className = "row-action";
      magnetLink.type = "button";
      magnetLink.textContent = "复制磁链";
      magnetLink.addEventListener("click", async () => {
        await copyText(item.magnet);
        showToast("磁链已复制");
      });
      const done = document.createElement("span");
      done.className = "follow-done-badge";
      done.textContent = copiedMagnets.has(item.magnet) ? "已完成" : "追番中";
      const actions = document.createElement("div");
      actions.className = "follow-actions";
      actions.append(magnetLink, done);
      row.append(main, actions);
    } else {
      const unfollow = document.createElement("button");
      unfollow.className = "row-action";
      unfollow.type = "button";
      unfollow.textContent = "取消追番";
      unfollow.addEventListener("click", () => toggleFollow(item, null));
      row.append(main, unfollow);
    }
    return row;
  };
  elements.followsDoneCol.replaceChildren(...done3d.map((item) => renderColItem(item, false)));
  elements.followsAllCol.replaceChildren(
    ...allItems.map((item) => renderColItem(item, true))
  );
  elements.followsColumns.hidden = done3d.length === 0 && allItems.length === 0;
  elements.followsRows.replaceChildren();
  elements.followsCount.textContent = allItems.length ? String(allItems.length) : "0";
  elements.followsEmpty.hidden = done3d.length > 0 || allItems.length > 0;
}

async function openVersionLog() {
  try {
    const response = await fetch("/api/version", { cache: "no-store" });
    if (!response.ok) throw new Error("无法读取更新日志");
    const data = await response.json();
    elements.versionCurrent.textContent = `当前版本 ${data.current || ""}`;
    const entries = Array.isArray(data.entries) ? data.entries : [];
    elements.versionEntries.replaceChildren(
      ...entries
        .slice()
        .reverse()
        .map((entry) => {
          const item = document.createElement("div");
          item.className = "version-entry";
          const head = document.createElement("div");
          head.className = "version-entry-head";
          const version = document.createElement("strong");
          version.textContent = `v${entry.version}`;
          const time = document.createElement("span");
          time.textContent = entry.time || "";
          head.append(version, time);
          const summary = document.createElement("p");
          summary.textContent = entry.summary || "";
          item.append(head, summary);
          if (Array.isArray(entry.details) && entry.details.length > 0) {
            const list = document.createElement("ul");
            entry.details.forEach((detail) => {
              const li = document.createElement("li");
              li.textContent = detail;
              list.append(li);
            });
            item.append(list);
          }
          return item;
        })
    );
    elements.versionModal.hidden = false;
  } catch (error) {
    showToast("更新日志不可用，请先重启服务：launchctl kickstart -k gui/$(id -u)/com.dmhy.zhufan");
  }
}

function closeVersionLog() {
  elements.versionModal.hidden = true;
}

function updateSelectionButtons() {
  elements.copyRecentButton.textContent = selectedRecent.size
    ? `复制选中（${selectedRecent.size}）`
    : "复制选中";
  elements.copyRecentButton.disabled = selectedRecent.size === 0;
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const area = document.createElement("textarea");
    area.value = text;
    document.body.append(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
}

function batchMarkCopied(magnets) {
  magnets.forEach((magnet) => copiedMagnets.add(magnet));
  saveCopiedMagnets();
  renderRecent();
}

async function copySelected(set) {
  const magnets = [...set];
  if (magnets.length === 0) return;
  await copyText(magnets.join("\n"));
  set.clear();
  batchMarkCopied(magnets);
  updateSelectionButtons();
  showToast(`已复制 ${magnets.length} 条磁链`);
}

async function loadRecent(days, afterLoad) {
  const requestId = ++recentRequestId;
  recentDays = days;
  recentPage = 1;
  elements.recentDaysInput.value = String(days);
  elements.recentPresets.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("active", Number(button.dataset.days) === Number(days));
  });
  const groups = document.querySelector('input[name="recentGroup"]:checked')?.value || "";
  const query = recentQuery.trim();
  const useSiteSearch = siteSearch && query;
  elements.recentRefresh.disabled = true;
  elements.recentRefresh.textContent = useSiteSearch ? "搜索中…" : "读取中…";
  try {
    const requestUrl = useSiteSearch
      ? `/api/search?q=${encodeURIComponent(query)}&groups=${encodeURIComponent(groups)}`
      : `/api/recent?days=${encodeURIComponent(days)}&groups=${encodeURIComponent(groups)}&q=${encodeURIComponent(query)}`;
    const response = await fetch(requestUrl, { cache: "no-store" });
    if (!response.ok) {
      let message = useSiteSearch ? "全站搜索失败" : "无法读取近期发布";
      try {
        const payload = await response.json();
        if (payload && payload.error) message = payload.error;
      } catch {
        // keep the generic message
      }
      throw new Error(message);
    }
    const data = await response.json();
    if (requestId !== recentRequestId) return;
    currentRecent = Array.isArray(data.results) ? data.results : [];
    renderRecent();
    if (afterLoad) afterLoad();
    if (useSiteSearch) {
      const suffix = data.source === "rss"
        ? "（DMHY 搜索服务故障，仅覆盖最近约 500 条）"
        : "（来自 DMHY 全站）";
      elements.recentMeta.textContent = `全站搜索「${data.q}」共 ${data.total} 条${suffix}`;
    } else if (data.q) {
      elements.recentMeta.textContent = `搜索「${data.q}」共 ${data.total} 条（已加载的全部数据）`;
    } else {
      const groupLabel = data.groups && data.groups.length ? data.groups.map((id) => id === "7acg" ? "7³ACG" : "LoliHouse").join("、") : "全部";
      elements.recentMeta.textContent = `读取最近 ${data.days} 天（${groupLabel}）的发布记录，共 ${data.total} 条`;
    }
  } catch (error) {
    if (requestId !== recentRequestId) return;
    elements.recentMeta.textContent = error.message;
  } finally {
    if (requestId !== recentRequestId) return;
    elements.recentRefresh.disabled = false;
    elements.recentRefresh.textContent = "刷新";
  }
}

async function loadFollowRecent() {
  try {
    const response = await fetch("/api/recent?all=1&groups=", { cache: "no-store" });
    if (!response.ok) throw new Error("无法读取追番数据");
    const data = await response.json();
    followRecent = Array.isArray(data.results) ? data.results : [];
  } catch (error) {
    // keep the last known follows data on failure
  }
  renderFollows();
}

function setFollowsDays(days) {
  followsDays = days;
  elements.followsPresets.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("active", Number(button.dataset.fdays) === Number(days));
  });
  elements.followsDoneLabel.textContent = `${days} 日内完成`;
  renderFollows();
}

function filteredRecent() {
  const encoding = document.querySelector('input[name="recentEncoding"]:checked')?.value || "";
  const language = document.querySelector('input[name="recentLanguage"]:checked')?.value || "";
  const group = document.querySelector('input[name="recentGroup"]:checked')?.value || "";
  return currentRecent.filter((item) => {
    if (encoding && extractEncoding(item.title) !== encoding) return false;
    if (language && extractLanguage(item.title) !== language) return false;
    if (group && item.group_id !== group) return false;
    return true;
  });
}

function renderRecent() {
  syncEncodingOptions("recentEncoding", currentRecent);
  syncLanguageOptions("recentLanguage", currentRecent);
  const filtered = filteredRecent();
  const total = filtered.length;
  const pageCount = Math.max(1, Math.ceil(total / RECENT_PAGE_SIZE));
  recentPage = Math.min(Math.max(1, recentPage), pageCount);
  const start = (recentPage - 1) * RECENT_PAGE_SIZE;
  const pageItems = filtered.slice(start, start + RECENT_PAGE_SIZE);
  elements.recentRows.replaceChildren(...pageItems.map((item) => createResultRow(item, selectedRecent, updateSelectionButtons)));
  elements.recentCount.textContent = String(total);
  elements.recentEmpty.hidden = total > 0;
  const emptyTitle = elements.recentEmpty.querySelector("strong");
  const emptyHint = elements.recentEmpty.querySelector("span");
  if (recentQuery.trim() && siteSearch) {
    emptyTitle.textContent = `全站没有找到「${recentQuery.trim()}」的相关资源`;
    emptyHint.textContent = "换个关键词试试，或关闭全站查看已加载的数据。";
  } else if (recentQuery.trim()) {
    emptyTitle.textContent = `已加载数据中没有「${recentQuery.trim()}」`;
    emptyHint.textContent = "试试开启「全站」，直接搜索 DMHY 站内资源。";
  } else {
    emptyTitle.textContent = "近期没有新增发布";
    emptyHint.textContent = "选择其他天数或先执行一次更新。";
  }
  elements.recentPagination.hidden = pageCount <= 1;
  elements.recentPageInfo.textContent = `第 ${recentPage} / ${pageCount} 页`;
  elements.recentPrev.disabled = recentPage <= 1;
  elements.recentNext.disabled = recentPage >= pageCount;
  updateSelectionButtons();
}

function syncEncodingOptions(name, items) {
  const present = new Set(
    items.map((item) => extractEncoding(item.title)).filter((value) => value !== "—")
  );
  const checked = document.querySelector(`input[name="${name}"]:checked`);
  let selected = checked?.value || "";
  if (selected && !present.has(selected)) {
    selected = "";
  }
  document.querySelectorAll(`input[name="${name}"]`).forEach((radio) => {
    const label = radio.closest("label");
    const show = radio.value === "" || present.has(radio.value);
    if (label) label.hidden = !show;
    radio.checked = radio.value === selected;
  });
}

function syncLanguageOptions(name, items) {
  const present = new Set(
    items.map((item) => extractLanguage(item.title)).filter((value) => value !== "—")
  );
  const checked = document.querySelector(`input[name="${name}"]:checked`);
  let selected = checked?.value || "";
  if (selected && !present.has(selected)) {
    selected = "";
  }
  document.querySelectorAll(`input[name="${name}"]`).forEach((radio) => {
    const label = radio.closest("label");
    const show = radio.value === "" || present.has(radio.value);
    if (label) label.hidden = !show;
    radio.checked = radio.value === selected;
  });
}

function renderState(state) {
  elements.headerStatus.dataset.status = state.status;
  elements.headerStatus.querySelector("span:last-child").textContent =
    state.status === "running" ? "更新中" : state.status === "error" ? "更新失败" : state.status === "success" ? "已完成" : "准备就绪";
  elements.runStatus.dataset.status = state.status;
  elements.statusMessage.textContent = state.message;
  elements.statusMeta.textContent = statusMeta(state);
  elements.progressBar.style.width = `${Math.max(0, Math.min(100, state.progress || 0))}%`;
  elements.updateButton.disabled = state.status === "running";
  elements.updateButton.textContent = state.status === "running" ? "正在更新…" : "更新";
  elements.errorMessage.hidden = !state.error;
  elements.errorMessage.textContent = state.error || "";
  elements.downloadButton.classList.toggle("disabled", !state.workbook);
  elements.logOutput.textContent = state.logs.length ? state.logs.join("\n") : "尚无运行记录";

  if (lastStatus === "running" && state.status === "success") {
    showToast(state.message);
    loadRecent(recentDays, () => loadFollowRecent());
  } else if (lastStatus === "running" && state.status === "error") {
    showToast("更新失败，请查看状态信息");
  }
  lastStatus = state.status;
}

async function fetchStatus() {
  try {
    const response = await fetch("/api/status", { cache: "no-store" });
    if (!response.ok) throw new Error("无法读取状态");
    renderState(await response.json());
  } catch (error) {
    elements.statusMessage.textContent = error.message;
    elements.runStatus.dataset.status = "error";
  }
}

async function startUpdate() {
  const days = Number(elements.recentDaysInput.value);
  const groupValue = document.querySelector('input[name="recentGroup"]:checked')?.value || "";
  const groups = groupValue ? [groupValue] : ["lolihouse", "7acg"];
  if (!Number.isInteger(days) || days < 1 || days > 365) {
    showToast("更新时间需为 1 到 365 天");
    elements.recentDaysInput.focus();
    return;
  }
  if (groups.length === 0) {
    showToast("请至少选择一个字幕组");
    return;
  }
  elements.updateButton.disabled = true;
  try {
    const response = await fetch("/api/update", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ days, groups }),
    });
    const state = await response.json();
    if (!response.ok) throw new Error(state.error || "无法开始更新");
    renderState(state);
  } catch (error) {
    showToast(error.message);
    elements.updateButton.disabled = false;
  }
}

function clearSearch() {
  window.clearTimeout(recentSearchTimer);
  if (elements.recentSearchInput.value || recentQuery) {
    elements.recentSearchInput.value = "";
    recentQuery = "";
  }
}

function updateSearchPlaceholder() {
  if (elements.recentSearchInput) {
    elements.recentSearchInput.placeholder = siteSearch
      ? "搜索 DMHY 全站资源…"
      : "搜索番剧名称 / 集数…";
  }
}

elements.updateButton.addEventListener("click", startUpdate);
elements.recentPresets.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-days]");
  if (button) {
    clearSearch();
    loadRecent(Number(button.dataset.days));
  }
});
elements.recentSearchInput.addEventListener("input", () => {
  window.clearTimeout(recentSearchTimer);
  recentSearchTimer = window.setTimeout(() => {
    recentQuery = elements.recentSearchInput.value.trim();
    loadRecent(recentDays);
  }, siteSearch ? 500 : 260);
});
elements.siteSearchToggle?.addEventListener("change", () => {
  siteSearch = elements.siteSearchToggle.checked;
  try {
    localStorage.setItem(SITE_SEARCH_KEY, siteSearch ? "1" : "0");
  } catch {
    // storage may be unavailable
  }
  updateSearchPlaceholder();
  if (recentQuery.trim()) loadRecent(recentDays);
});
elements.recentSearchInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    window.clearTimeout(recentSearchTimer);
    recentQuery = elements.recentSearchInput.value.trim();
    loadRecent(recentDays);
  }
});
elements.recentSearchInput.addEventListener("search", () => {
  window.clearTimeout(recentSearchTimer);
  recentQuery = elements.recentSearchInput.value.trim();
  loadRecent(recentDays);
});
elements.recentRefresh.addEventListener("click", () => {
  const days = Number(elements.recentDaysInput.value);
  if (!Number.isInteger(days) || days < 1 || days > 365) {
    showToast("天数需为 1 到 365");
    elements.recentDaysInput.focus();
    return;
  }
  loadRecent(days, () => loadFollowRecent());
});
elements.followsPresets.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-fdays]");
  if (button) setFollowsDays(Number(button.dataset.fdays));
});
elements.recentDaysInput.addEventListener("change", () => {
  const days = Number(elements.recentDaysInput.value);
  if (Number.isInteger(days) && days >= 1 && days <= 365) {
    clearSearch();
    loadRecent(days);
  }
});
elements.recentGroupFilter.addEventListener("change", () => {
  renderRecent();
  loadRecent(recentDays);
});
elements.recentPrev.addEventListener("click", () => {
  if (recentPage > 1) {
    recentPage -= 1;
    renderRecent();
    elements.recentRows.closest(".recent-scroll")?.scrollTo({ top: 0 });
  }
});
elements.recentNext.addEventListener("click", () => {
  const pageCount = Math.max(1, Math.ceil(filteredRecent().length / RECENT_PAGE_SIZE));
  if (recentPage < pageCount) {
    recentPage += 1;
    renderRecent();
    elements.recentRows.closest(".recent-scroll")?.scrollTo({ top: 0 });
  }
});
elements.copyRecentButton.addEventListener("click", () => copySelected(selectedRecent));
elements.copyAllFollowsButton.addEventListener("click", async () => {
  if (currentFollowItems.length === 0) {
    showToast("追番更新中没有可复制的磁链");
    return;
  }
  const magnets = currentFollowItems.map((item) => item.magnet);
  await copyText(magnets.join("\n"));
  batchMarkCopied(magnets);
  renderFollows();
  showToast(`已复制 ${magnets.length} 条磁链`);
});
elements.versionButton.addEventListener("click", openVersionLog);
elements.versionClose.addEventListener("click", closeVersionLog);
elements.versionModal.addEventListener("click", (event) => {
  if (event.target === elements.versionModal) closeVersionLog();
});
elements.followsDoneCopy.addEventListener("click", async () => {
  if (currentDoneItems.length === 0) {
    showToast("左栏没有可复制的磁链");
    return;
  }
  const magnets = currentDoneItems.map((item) => item.magnet);
  await copyText(magnets.join("\n"));
  showToast(`已复制 ${magnets.length} 条磁链`);
});
document.querySelectorAll('input[name="recentEncoding"]').forEach((radio) => {
  radio.addEventListener("click", () => {
    recentPage = 1;
    renderRecent();
  });
});
document.querySelectorAll('input[name="recentLanguage"]').forEach((radio) => {
  radio.addEventListener("click", () => {
    recentPage = 1;
    renderRecent();
  });
});

if (elements.siteSearchToggle) {
  elements.siteSearchToggle.checked = siteSearch;
}
updateSearchPlaceholder();
fetchStatus();
loadRecent(7);
loadFollowRecent();
window.setInterval(fetchStatus, 1000);
