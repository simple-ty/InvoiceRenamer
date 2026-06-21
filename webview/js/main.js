/* Invoice Renamer WebView UI 交互逻辑（HTTP fetch 方案） */

const state = {
  version: "",
  fieldOrder: [],
  fieldEnabled: {},
  fieldLabels: {},
  customValue: "",
  previewMode: true,
  records: [],
  cloud: { enabled: false, configured: false, secret_id: "" },
  processing: false,
  scanning: false,
  renameHistory: false,
  sortCol: null,
  sortReverse: false,
  _initialized: false,
  _polling: false,
};

const STAT_STYLE = {
  total: { title: "文件总数", accent: "#191919" },
  complete: { title: "完整识别", accent: "#07C160" },
  partial: { title: "部分识别", accent: "#FA9D3B" },
  failed: { title: "未识别/异常", accent: "#FA5151" },
  not_invoice: { title: "非发票", accent: "#B0B0B0" },
  cloud_not_configured: { title: "云端未启用", accent: "#3FA9F5" },
};

const $ = (id) => document.getElementById(id);

// ── fetch 通信 ──────────────────────────────────────────────────────────

async function apiGet(endpoint) {
  const resp = await fetch("/api/" + endpoint);
  if (!resp.ok) throw new Error("HTTP " + resp.status);
  return resp.json();
}

async function apiPost(endpoint, params = {}) {
  const resp = await fetch("/api/" + endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(params),
  });
  if (!resp.ok) throw new Error("HTTP " + resp.status);
  return resp.json();
}

// ── 事件轮询 ────────────────────────────────────────────────────────────

function startPolling() {
  if (state._polling) return;
  state._polling = true;
  setInterval(async () => {
    try {
      const data = await apiGet("poll");
      if (data.events && data.events.length > 0) {
        data.events.forEach(evt => {
          if (window.__onPyEvent__) window.__onPyEvent__(evt);
        });
      }
    } catch (e) { /* 静默 */ }
  }, 50);
}

// ── 初始化 ──────────────────────────────────────────────────────────────

async function init() {
  if (state._initialized) return;
  state._initialized = true;

  try {
    const init = await apiGet("get_init_state");
    Object.assign(state, {
      version: init.version || "",
      fieldOrder: init.field_order || [],
      fieldEnabled: init.field_enabled || {},
      fieldLabels: init.field_labels || {},
      customValue: init.custom_value || "",
      previewMode: init.preview_mode !== false,
      cloud: init.cloud || { enabled: false, configured: false, secret_id: "" },
    });

    $("custom-input").value = state.customValue;

    renderTemplateRows();
    renderStats(init.stats || {});
    renderTable(init.records || []);
    updateCloudButton();
    updateRenameButton();
    updatePreviewSwitch();
    syncCustomInputState();
    bindEvents();
    startPolling();
    updateStatus("就绪");
    checkUpdate();  // 静默检查更新
  } catch (e) {
    showToast("初始化失败", e.message, "error");
    updateStatus("初始化失败");
    state._initialized = false;
    state._initAttempts = (state._initAttempts || 0) + 1;
    if (state._initAttempts < 10) {
      setTimeout(() => init(), 1000);
    } else {
      updateStatus("初始化失败，请重启程序");
    }
  }
}

function bindEvents() {
  $("source-btn").addEventListener("click", (e) => {
    e.stopPropagation();
    const menu = $("source-menu");
    const rect = $("source-btn").getBoundingClientRect();
    menu.style.left = rect.left + "px";
    menu.style.top = (rect.bottom + 4) + "px";
    menu.classList.toggle("show");
  });

  document.addEventListener("click", () => $("source-menu").classList.remove("show"));

  $("source-menu").addEventListener("click", (e) => {
    e.stopPropagation();
    const item = e.target.closest(".dropdown-item");
    if (!item) return;
    handleSourceAction(item.dataset.action);
  });

  $("custom-input").addEventListener("input", debounce(updateTemplate, 200));

  $("cloud-row").addEventListener("click", openCloudSettings);
  $("cloud-status-btn").addEventListener("click", (e) => {
    e.stopPropagation();
    if (state.cloud.configured) toggleCloudEnabled();
    else openCloudSettings();
  });

  $("preview-switch").addEventListener("click", togglePreviewMode);
  $("rescan-btn").addEventListener("click", rescanFiles);
  $("export-btn").addEventListener("click", exportExcel);
  $("rename-btn").addEventListener("click", onRenameClick);

  $("cancel-settings-btn").addEventListener("click", closeCloudSettings);
  $("save-settings-btn").addEventListener("click", saveCloudSettings);
  $("clear-key-btn").addEventListener("click", clearCloudSettings);
  $("verify-key-btn").addEventListener("click", verifyCloudSettings);
  $("eye-btn").addEventListener("click", toggleSecretKeyVisibility);
  $("get-key-link").addEventListener("click", (e) => {
    e.preventDefault();
    apiPost("open_browser", { url: "https://console.cloud.tencent.com/cam/capi" });
  });
  $("secret-id-input").addEventListener("input", updateSwitchState);
  $("secret-key-input").addEventListener("input", updateSwitchState);

  $("update-close").addEventListener("click", () => {
    $("update-banner").classList.remove("show");
  });
  $("update-link").addEventListener("click", (e) => {
    e.preventDefault();
    apiPost("open_browser", { url: $("update-link").href });
  });

  // 手动检查更新按钮
  $("check-update-btn").addEventListener("click", () => checkUpdate(true));

  $("settings-overlay").addEventListener("click", (e) => {
    if (e.target === $("settings-overlay")) closeCloudSettings();
  });

  document.querySelectorAll(".th").forEach(th => {
    th.addEventListener("click", () => sortTable(th.dataset.col));
  });

  // 表格内编辑 — 点击 ✎ 切换输入
  $("table-body").addEventListener("click", (e) => {
    const icon = e.target.closest(".edit-icon");
    if (!icon) return;
    startEditName(icon.parentElement);
  });
}

// ── 模板渲染与拖拽 ──────────────────────────────────────────────────────

function renderTemplateRows() {
  const container = $("template-rows");
  container.innerHTML = "";
  state.fieldOrder.forEach((key, idx) => {
    const row = document.createElement("div");
    row.className = "template-row";
    row.dataset.key = key;
    row.draggable = true;
    row.innerHTML = `
      <div class="row-badge">${String(idx + 1).padStart(2, "0")}</div>
      <input type="checkbox" class="row-checkbox" ${state.fieldEnabled[key] ? "checked" : ""}>
      <div class="row-label">${state.fieldLabels[key] || key}</div>
      <div class="row-handle">≡</div>
    `;
    const checkbox = row.querySelector(".row-checkbox");
    checkbox.addEventListener("change", () => {
      state.fieldEnabled[key] = checkbox.checked;
      if (key === "custom") syncCustomInputState();
      updateTemplate();
    });
    row.addEventListener("dragstart", onDragStart);
    row.addEventListener("dragend", onDragEnd);
    row.addEventListener("dragover", onDragOver);
    row.addEventListener("drop", onDrop);
    container.appendChild(row);
  });
}

let dragKey = null;
function onDragStart(e) { dragKey = this.dataset.key; this.classList.add("dragging"); e.dataTransfer.effectAllowed = "move"; }
function onDragEnd() { this.classList.remove("dragging"); dragKey = null; }
function onDragOver(e) { e.preventDefault(); e.dataTransfer.dropEffect = "move"; }
function onDrop(e) {
  e.preventDefault();
  if (!dragKey) return;
  const targetKey = this.dataset.key;
  if (dragKey === targetKey) return;
  const fromIdx = state.fieldOrder.indexOf(dragKey);
  const toIdx = state.fieldOrder.indexOf(targetKey);
  state.fieldOrder.splice(fromIdx, 1);
  state.fieldOrder.splice(toIdx, 0, dragKey);
  renderTemplateRows();
  updateTemplate();
}

function syncCustomInputState() {
  const input = $("custom-input");
  const enabled = state.fieldEnabled.custom === true;
  input.disabled = !enabled;
  if (!enabled) input.value = "";
}

async function updateTemplate() {
  state.customValue = $("custom-input").value;
  const result = await apiPost("update_template", {
    field_order: state.fieldOrder,
    field_enabled: state.fieldEnabled,
    custom_value: state.customValue,
  });
  if (result && result.records) {
    state.records = result.records;
    renderTable(state.records);
  }
}

// ── 来源选择 ────────────────────────────────────────────────────────────

async function handleSourceAction(action) {
  $("source-menu").classList.remove("show");
  if (action === "clear_source") {
    const result = await apiGet("clear_source");
    if (result) applyState(result.state);
    return;
  }
  if (action === "choose_folder") {
    updateStatus("请选择文件夹...");
    const result = await apiPost("choose_folder");
    if (result && result.ok) {
      $("path-input").value = result.path;
      apiPost("scan_files");
      state.scanning = true;
      updateStatus("识别中...");
    } else {
      updateStatus(result?.error || "已取消");
    }
  } else if (action === "choose_files") {
    updateStatus("请选择文件...");
    const result = await apiPost("choose_files");
    if (result && result.ok) {
      $("path-input").value = result.path;
      apiPost("scan_files");
      state.scanning = true;
      updateStatus("识别中...");
    } else {
      updateStatus(result?.error || "已取消");
    }
  }
}

async function rescanFiles() {
  apiPost("scan_files");
  state.scanning = true;
  updateStatus("识别中...");
}

// ── 表格渲染 ────────────────────────────────────────────────────────────

function buildRow(r) {
  const row = document.createElement("div");
  row.className = "table-row";
  row.dataset.idx = r.idx;
  row.dataset.status = r.status;
  const statusText = r.manual_override ? "手动重命名"
    : { complete: "完整识别", partial: "部分识别",
        failed: "解析失败", cloud_error: "云端异常",
        not_invoice: "非发票", cloud_not_configured: "云端未启用",
      }[r.status] || r.status;
  const statusClass = r.manual_override ? "manual_rename" : r.status;
  const rowClass = r.manual_override ? ""
    : ["failed", "cloud_error"].includes(r.status) ? "row-error"
    : r.status === "not_invoice" ? "row-weak"
    : r.status === "cloud_not_configured" ? "row-info"
    : r.status === "partial" ? "row-warning"
    : "";
  const editable = ["partial", "failed", "cloud_error", "cloud_not_configured"].includes(r.status);
  const newNameClass = r.manual_override ? "col-new manual-override" : "col-new";
  const newNameHtml = editable
    ? `<div class="td ${newNameClass} editable-cell" title="${esc(r.new_name)}" data-new-name="${esc(r.new_name)}">
         <span class="edit-icon">✎ 编辑</span>
         <span class="edit-text${r.manual_override ? '' : ' muted'}">${esc(r.new_name)}</span>
       </div>`
    : `<div class="td col-new" title="${esc(r.new_name)}">${esc(r.new_name)}</div>`;
  row.innerHTML = `
    <div class="td col-idx">${r.idx}</div>
    <div class="td col-org" title="${esc(r.source_name)}">${esc(r.source_name)}</div>
    ${newNameHtml}
    <div class="td col-type">${esc(r.type)}</div>
    <div class="td col-seller" title="${esc(r.seller)}">${esc(r.seller)}</div>
    <div class="td col-amount">${esc(r.amount)}</div>
    <div class="td col-status ${statusClass}">${esc(statusText)}</div>
  `;
  if (rowClass) row.classList.add(rowClass);
  return row;
}

function renderTable(records) {
  state.records = records || [];
  const tbody = $("table-body");
  if (!state.records.length) {
    tbody.innerHTML = '<div class="empty-tip">暂无识别结果</div>';
    return;
  }
  tbody.innerHTML = "";
  state.records.forEach(r => {
    tbody.appendChild(buildRow(r));
  });
}

function sortTable(col) {
  if (state.sortCol === col) state.sortReverse = !state.sortReverse;
  else { state.sortCol = col; state.sortReverse = false; }
  document.querySelectorAll(".th").forEach(th => {
    th.classList.remove("sort-asc", "sort-desc");
    if (th.dataset.col === col) th.classList.add(state.sortReverse ? "sort-desc" : "sort-asc");
  });
  const colMap = { org: "source_name", new: "new_name", type: "type", seller: "seller", amount: "amount", status: "status" };
  const key = colMap[col] || col;
  const sorted = [...state.records].sort((a, b) => {
    let va = a[key] || "", vb = b[key] || "";
    if (typeof va === "number" && typeof vb === "number")
      return state.sortReverse ? vb - va : va - vb;
    return state.sortReverse
      ? String(vb).localeCompare(String(va), "zh-CN")
      : String(va).localeCompare(String(vb), "zh-CN");
  });
  renderTable(sorted);
}

// ── 统计 ────────────────────────────────────────────────────────────────

function _calcStats(records) {
  const stats = { total: 0, complete: 0, partial: 0, failed: 0, not_invoice: 0, cloud_not_configured: 0 };
  records.forEach(r => {
    stats.total++;
    const key = r.status || "";
    if (stats[key] !== undefined) stats[key]++;
    else if (key === "cloud_error") stats.failed++;  // 归入失败
  });
  return stats;
}

function renderStats(stats) {
  const container = $("stats-area");
  container.innerHTML = "";
  ["total", "complete", "partial", "failed", "not_invoice", "cloud_not_configured"].forEach(key => {
    const count = stats[key] || 0;
    if (key === "cloud_not_configured" && count === 0) return;  // 无需要时不显示
    const style = STAT_STYLE[key];
    const chip = document.createElement("div");
    chip.className = "stat-chip";
    chip.innerHTML = `<span class="stat-title">${style.title}</span><span class="stat-value" style="color:${style.accent}">${stats[key] || 0}</span>`;
    container.appendChild(chip);
  });
}

// ── 操作栏 ──────────────────────────────────────────────────────────────

function updateStatus(msg) { $("status-text").textContent = msg; }
function setProgress(v) { $("progress-fill").style.width = (Math.max(0, Math.min(1, v)) * 100) + "%"; }

function togglePreviewMode() {
  state.previewMode = !state.previewMode;
  updatePreviewSwitch();
  apiPost("set_preview_mode", { preview_mode: state.previewMode });
  updateRenameButton();
}

function updatePreviewSwitch() {
  $("preview-switch").classList.toggle("on", state.previewMode);
}

function updateRenameButton() {
  const btn = $("rename-btn");
  if (state.renameHistory) {
    btn.textContent = "撤销重命名";
    btn.className = "btn";
    btn.style.background = "#F9D65C";
    btn.style.color = "#191919";
  } else {
    btn.textContent = "开始重命名";
    btn.className = "btn btn-primary";
    btn.style.background = "";
    btn.style.color = "";
  }
  btn.disabled = state.processing || state.scanning || state.previewMode || state.records.length === 0;
}

async function onRenameClick() {
  if (state.renameHistory) {
    apiPost("on_rename_button_click");
    state.processing = true;
    updateStatus("撤销中...");
  } else {
    apiPost("on_rename_button_click");
    state.processing = true;
    updateStatus("重命名中...");
  }
  updateRenameButton();
}

async function exportExcel() {
  const result = await apiGet("export_excel");
  if (result && result.ok) {
    showToast("导出成功", result.path);
    updateStatus("导出成功");
  } else {
    showToast("导出失败", result?.error || "导出失败", "error");
    updateStatus("导出失败");
  }
}

function showToast(title, sub, type) {
  type = type || "success";
  $("toast-title").textContent = title;
  $("toast-sub").textContent = sub || "";
  var toast = $("toast");
  toast.className = "toast " + type;
  toast.classList.add("show");
  clearTimeout(toast._timer);
  toast._timer = setTimeout(function () { toast.classList.remove("show"); }, 4000);
}

// ── 云端 OCR 设置 ───────────────────────────────────────────────────────

async function openCloudSettings() {
  const s = await apiGet("get_cloud_settings");
  $("secret-id-input").value = s.secret_id || "";
  $("secret-key-input").value = s.secret_key || "";
  $("cloud-enable-input").checked = s.enabled || false;
  updateUsageInfo(s.usage || {});
  updateSwitchState();
  $("settings-overlay").classList.add("show");
}

function closeCloudSettings() { $("settings-overlay").classList.remove("show"); }

async function saveCloudSettings() {
  const sid = $("secret-id-input").value.trim();
  const skey = $("secret-key-input").value.trim();
  if (!sid || !skey) { alert("请先输入 SecretId 和 SecretKey"); return; }
  const result = await apiPost("save_cloud_settings", {
    secret_id: sid, secret_key: skey, enabled: $("cloud-enable-input").checked,
  });
  if (result && result.ok) {
    state.cloud = result.cloud;
    updateCloudButton();
    updateRenameButton();
    closeCloudSettings();
  }
}

async function clearCloudSettings() {
  const result = await apiGet("clear_cloud_settings");
  if (result && result.ok) {
    state.cloud = result.cloud;
    $("secret-id-input").value = "";
    $("secret-key-input").value = "";
    $("cloud-enable-input").checked = false;
    updateSwitchState();
    updateCloudButton();
    updateRenameButton();
  }
}

async function verifyCloudSettings() {
  const sid = $("secret-id-input").value.trim();
  const skey = $("secret-key-input").value.trim();
  if (!sid || !skey) { showVerifyResult(false, "请先输入密钥"); return; }
  $("verify-key-btn").disabled = true;
  $("verify-key-btn").textContent = "验证中...";
  const result = await apiPost("verify_cloud_credentials", { secret_id: sid, secret_key: skey });
  $("verify-key-btn").disabled = false;
  $("verify-key-btn").textContent = "验证密钥";
  showVerifyResult(result && result.ok, result ? result.message : "验证失败");
}

function showVerifyResult(valid, message) {
  const el = $("verify-result");
  el.textContent = valid ? "密钥有效" : (message || "密钥无效");
  el.className = "verify-result show " + (valid ? "valid" : "invalid");
  setTimeout(() => { el.className = "verify-result"; }, 3000);
}

function toggleSecretKeyVisibility() {
  const input = $("secret-key-input");
  const openIcon = document.querySelector(".eye-open");
  const closedIcon = document.querySelector(".eye-closed");
  if (input.type === "password") {
    input.type = "text";
    openIcon.style.display = "none";
    closedIcon.style.display = "block";
  } else {
    input.type = "password";
    openIcon.style.display = "block";
    closedIcon.style.display = "none";
  }
}

async function toggleCloudEnabled() {
  const result = await     apiPost("toggle_cloud_enabled");;
  if (result && result.ok) {
    state.cloud = result.cloud;
    updateCloudButton();
    updateRenameButton();
  }
}

function updateCloudButton() {
  const btn = $("cloud-status-btn");
  btn.className = "cloud-status-btn";
  if (!state.cloud.configured) { btn.classList.add("unconfigured"); btn.textContent = "● 未配置"; }
  else if (state.cloud.enabled) { btn.classList.add("configured-enabled"); btn.textContent = "● 已启用"; }
  else { btn.classList.add("configured-disabled"); btn.textContent = "● 未启用"; }
}

function updateUsageInfo(usage) {
  const used = usage.used || 0, limit = usage.limit || 1000;
  $("usage-info").innerHTML = `<div>本月已调用 ${used} 次</div><div>免费额度 ${limit} 次/月</div>`;
}

function updateSwitchState() {
  const sid = $("secret-id-input").value.trim();
  const skey = $("secret-key-input").value.trim();
  const input = $("cloud-enable-input");
  input.disabled = !sid || !skey;
  if (!sid || !skey) input.checked = false;
}

// ── Python 事件处理 ─────────────────────────────────────────────────────

window.__onPyEvent__ = function(payload) {
  if (!payload) return;
  const { event, data } = payload;
  switch (event) {
    case "scan_started":
      state.scanning = true; state.records = [];
      renderTable([]); renderStats({}); setProgress(0);
      updateStatus("识别中..."); break;
    case "scan_progress":
      setProgress(data.current / data.total);
      updateStatus(`识别中... (${data.current}/${data.total})`); break;
    case "scan_item":
      const tb = $("table-body");
      if (tb.querySelector(".empty-tip")) tb.innerHTML = "";
      state.records.push(data);
      tb.appendChild(buildRow(data));
      renderStats(_calcStats(state.records)); break;
    case "scan_finished":
      state.scanning = false;
      // 兜底：若实时推送未生效，从完整列表渲染
      if (data.records && data.records.length > state.records.length) {
        state.records = data.records;
        renderTable(state.records);
      }
      renderStats(data.stats || _calcStats(state.records));
      setProgress(1); updateStatus(data.message || "识别完成");
      updateRenameButton(); break;
    case "rename_started":
      state.processing = true; setProgress(0);
      updateStatus("重命名中..."); updateRenameButton(); break;
    case "rename_progress":
      setProgress(data.current / data.total);
      updateStatus(`重命名中... (${data.current}/${data.total})`);
      renderStats({ total: data.total, complete: data.success, partial: 0,
        failed: data.failed, not_invoice: data.skipped }); break;
    case "rename_item_done": updateRow(data.idx, data); break;
    case "rename_finished":
      state.processing = false; state.records = data.records || state.records;
      state.renameHistory = data.can_undo;
      renderTable(state.records); renderStats(data.stats || {});
      setProgress(1); updateStatus(data.message || "重命名完成");
      if (data.error_detail) showToast("重命名部分失败", data.error_detail, "error");
      updateRenameButton(); break;
    case "undo_started":
      state.processing = true; updateStatus("撤销中..."); updateRenameButton(); break;
    case "undo_finished":
      state.processing = false; state.records = data.records || state.records;
      state.renameHistory = data.can_undo;
      renderTable(state.records); renderStats(data.stats || {});
      updateStatus(data.message || "撤销完成");
      if (data.error_detail) showToast("撤销部分失败", data.error_detail, "error");
      updateRenameButton(); break;
    case "status": updateStatus(data.message || ""); break;
  }
};

// ── 行内编辑文件名 ──────────────────────────────────────────────────────

function startEditName(cell) {
  if (cell.classList.contains("editing")) return;
  const row = cell.closest(".table-row");
  const idx = row.dataset.idx;
  const curName = cell.dataset.newName;
  cell.classList.add("editing");
  cell.innerHTML = `
    <input class="edit-input" value="${esc(curName)}" maxlength="200">
    <span class="edit-confirm">✓</span>
    <span class="edit-cancel">✕</span>
  `;
  const input = cell.querySelector(".edit-input");
  input.focus();
  input.select();
  const commit = () => {
    const name = input.value.trim();
    if (!name) { cancelEdit(cell, curName); return; }
    cell.dataset.newName = name;
    apiPost("update_record_name", { idx: idx, new_name: name }).then(r => {
      if (r.ok && r.record) {
        const rec = state.records.find(v => v.idx == idx);
        if (rec) { rec.new_name = name; rec.manual_override = true; }
        cell.classList.remove("editing");
        cell.innerHTML = `<span class="edit-icon">✎ 编辑</span><span class="edit-text">${esc(name)}</span>`;
        cell.classList.add("manual-override");
      } else {
        cancelEdit(cell, curName);
      }
    }).catch(() => cancelEdit(cell, curName));
  };
  input.addEventListener("keydown", (e) => { if (e.key === "Enter") commit(); });
  cell.querySelector(".edit-confirm").addEventListener("click", commit);
  cell.querySelector(".edit-cancel").addEventListener("click", () => cancelEdit(cell, curName));
}

function cancelEdit(cell, restoreName) {
  cell.classList.remove("editing");
  cell.dataset.newName = restoreName;
  cell.innerHTML = `<span class="edit-icon">✎ 编辑</span><span class="edit-text muted">${esc(restoreName)}</span>`;
}

function updateRow(idx, data) {
  const row = document.querySelector(`.table-row[data-idx="${idx}"]`);
  if (!row) return;
  const newCell = row.querySelector(".col-new");
  const statusCell = row.querySelector(".col-status");
  if (newCell) newCell.textContent = data.new_name || newCell.textContent;
  if (statusCell) {
    statusCell.textContent = data.status === "success" ? "完成" : (data.error || "失败");
    statusCell.className = "td col-status " + (data.status === "success" ? "success" : "error");
  }
}

function applyState(newState) {
  state.fieldOrder = newState.field_order || state.fieldOrder;
  state.fieldEnabled = newState.field_enabled || state.fieldEnabled;
  state.customValue = newState.custom_value || "";
  state.cloud = newState.cloud || state.cloud;
  $("custom-input").value = state.customValue;
  $("path-input").value = "";
  renderTemplateRows();
  renderTable(newState.records || []);
  renderStats(newState.stats || {});
  updateCloudButton();
  updateRenameButton();
  syncCustomInputState();
}

// ── 工具函数 ────────────────────────────────────────────────────────────

function esc(text) {
  if (!text) return "";
  return String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function debounce(fn, delay) {
  let timer = null;
  return function(...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

// ── 版本更新检查 ──────────────────────────────────────────────────────

async function checkUpdate(force = false) {
  const btn = $("check-update-btn");
  if (force) {
    btn.disabled = true;
    btn.textContent = "检查中…";
    btn.classList.add("checking");
  }

  try {
    const query = force ? "?force=1" : "";
    const result = await apiGet("check_update" + query);

    if (result && result.has_update) {
      $("update-version").textContent = result.latest;
      $("update-link").href = result.url || "https://github.com/simple-ty/InvoiceRenamer/releases";
      $("update-banner").classList.add("show");
    }

    // 手动检查显示结果
    if (force) {
      if (result && result.has_update) {
        showToast("发现新版本", result.latest + "，点击上方提示条查看详情");
      } else if (result && result.error) {
        showToast("检查失败", result.error, "error");
      } else {
        showToast("已是最新版本", result.current);
      }
    }
  } catch (e) {
    if (force) showToast("检查失败", "网络不可用，请稍后重试", "error");
  } finally {
    if (force) {
      btn.disabled = false;
      btn.textContent = "检查更新";
      btn.classList.remove("checking");
      setTimeout(() => { btn.blur(); }, 200);
    }
  }
}

// ── 启动 ────────────────────────────────────────────────────────────────

// 不再依赖 pywebview JS bridge，页面加载后直接 fetch
if (document.readyState === "complete" || document.readyState === "interactive") {
  init();
} else {
  document.addEventListener("DOMContentLoaded", init);
}
