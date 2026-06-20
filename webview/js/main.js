/* Invoice Renamer WebView UI 交互逻辑 */

// 全局状态
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
};

// 复用配置
const TAG_COLORS = {
  success: "#191919",
  partial: "#FA9D3B",
  error: "#FA5151",
  idle: "#8A8A8A",
  not_invoice: "#B0B0B0",
};

const STAT_STYLE = {
  total: { title: "文件总数", accent: "#191919" },
  complete: { title: "完整识别", accent: "#07C160" },
  partial: { title: "部分识别", accent: "#FA9D3B" },
  failed: { title: "未识别/异常", accent: "#FA5151" },
  not_invoice: { title: "非发票", accent: "#B0B0B0" },
};

// DOM 元素缓存
const $ = (id) => document.getElementById(id);

// 桥接调用
async function py(method, ...args) {
  try {
    if (window.pywebview && window.pywebview.api) {
      return await window.pywebview.api[method](...args);
    }
  } catch (e) {
    console.error("pywebview call failed:", e);
  }
  return null;
}

// 初始化
async function init() {
  if (state._initialized) return;
  state._initialized = true;

  const init = await py("get_init_state") || {};
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
  $("summary-text").textContent = init.result_hint || "请选择文件或文件夹开始识别。";

  renderTemplateRows();
  renderStats(init.stats || {});
  renderTable(init.records || []);
  updateCloudButton();
  updateRenameButton();
  updatePreviewSwitch();
  bindEvents();
}

function bindEvents() {
  // 来源按钮下拉菜单
  $("source-btn").addEventListener("click", (e) => {
    e.stopPropagation();
    const menu = $("source-menu");
    const rect = $("source-btn").getBoundingClientRect();
    menu.style.left = rect.left + "px";
    menu.style.top = (rect.bottom + 4) + "px";
    menu.classList.toggle("show");
  });

  document.addEventListener("click", () => {
    $("source-menu").classList.remove("show");
  });

  $("source-menu").addEventListener("click", (e) => {
    e.stopPropagation();
    const item = e.target.closest(".dropdown-item");
    if (!item) return;
    const action = item.dataset.action;
    handleSourceAction(action);
  });

  // 自定义字段
  $("custom-input").addEventListener("input", debounce(updateTemplate, 200));

  // 云端设置入口
  $("cloud-row").addEventListener("click", () => {
    openCloudSettings();
  });
  $("cloud-status-btn").addEventListener("click", (e) => {
    e.stopPropagation();
    if (state.cloud.configured) {
      toggleCloudEnabled();
    } else {
      openCloudSettings();
    }
  });

  // 底部操作
  $("preview-switch").addEventListener("click", togglePreviewMode);
  $("rescan-btn").addEventListener("click", rescanFiles);
  $("export-btn").addEventListener("click", exportExcel);
  $("rename-btn").addEventListener("click", onRenameClick);

  // 设置弹窗
  $("cancel-settings-btn").addEventListener("click", closeCloudSettings);
  $("save-settings-btn").addEventListener("click", saveCloudSettings);
  $("clear-key-btn").addEventListener("click", clearCloudSettings);
  $("verify-key-btn").addEventListener("click", verifyCloudSettings);
  $("toggle-key-btn").addEventListener("click", toggleSecretKeyVisibility);
  $("get-key-link").addEventListener("click", (e) => {
    e.preventDefault();
    // 通过 Python 打开浏览器链接
    if (window.pywebview && window.pywebview.api) {
      window.pywebview.api.open_browser("https://console.cloud.tencent.com/cam/capi");
    } else {
      window.open("https://console.cloud.tencent.com/cam/capi", "_blank");
    }
  });
  $("secret-id-input").addEventListener("input", updateSwitchState);
  $("secret-key-input").addEventListener("input", updateSwitchState);

  $("settings-overlay").addEventListener("click", (e) => {
    if (e.target === $("settings-overlay")) closeCloudSettings();
  });

  // 表格列头排序
  document.querySelectorAll(".th").forEach(th => {
    th.addEventListener("click", () => {
      const col = th.dataset.col;
      sortTable(col);
    });
  });
}

// ── 模板渲染与拖拽 ───────────────────────────────────────────

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

function onDragStart(e) {
  dragKey = this.dataset.key;
  this.classList.add("dragging");
  e.dataTransfer.effectAllowed = "move";
}

function onDragEnd() {
  this.classList.remove("dragging");
  dragKey = null;
}

function onDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = "move";
}

function onDrop(e) {
  e.preventDefault();
  if (!dragKey) return;
  const targetKey = this.dataset.key;
  if (dragKey === targetKey) return;

  const fromIndex = state.fieldOrder.indexOf(dragKey);
  const toIndex = state.fieldOrder.indexOf(targetKey);
  state.fieldOrder.splice(fromIndex, 1);
  state.fieldOrder.splice(toIndex, 0, dragKey);

  renderTemplateRows();
  updateTemplate();
}

async function updateTemplate() {
  state.customValue = $("custom-input").value;
  const result = await py("update_template", state.fieldOrder, state.fieldEnabled, state.customValue);
  if (result && result.records) {
    state.records = result.records;
    renderTable(state.records);
  }
}

// ── 来源选择 ───────────────────────────────────────────────

async function handleSourceAction(action) {
  $("source-menu").classList.remove("show");
  if (action === "clear_source") {
    const result = await py("clear_source");
    if (result) applyState(result.state);
    return;
  }
  if (action === "choose_folder") {
    const result = await py("choose_folder");
    if (result && result.ok) {
      $("path-input").value = result.path;
      py("scan_files");
      state.scanning = true;
      updateStatus("识别中...");
    }
  } else if (action === "choose_files") {
    const result = await py("choose_files");
    if (result && result.ok) {
      $("path-input").value = result.path;
      py("scan_files");
      state.scanning = true;
      updateStatus("识别中...");
    }
  }
}

async function rescanFiles() {
  py("scan_files");
  state.scanning = true;
  updateStatus("识别中...");
}

// ── 表格渲染 ───────────────────────────────────────────────

function renderTable(records) {
  state.records = records;
  const tbody = $("table-body");
  if (!records || records.length === 0) {
    tbody.innerHTML = '<div class="empty-tip">暂无识别结果</div>';
    return;
  }

  tbody.innerHTML = "";
  records.forEach(r => {
    const row = document.createElement("div");
    row.className = "table-row";
    row.dataset.idx = r.idx;

    const statusClass = r.status || "idle";
    const statusText = {
      complete: "完成",
      partial: "部分",
      failed: r.error || "未识别",
      not_invoice: "非发票",
      idle: "待处理",
    }[r.status] || r.status;

    row.innerHTML = `
      <div class="td col-idx ${statusClass}">${r.idx}</div>
      <div class="td col-org ${statusClass}" title="${escapeHtml(r.source_name)}">${escapeHtml(r.source_name)}</div>
      <div class="td col-new ${statusClass}" title="${escapeHtml(r.new_name)}">${escapeHtml(r.new_name)}</div>
      <div class="td col-type ${statusClass}">${escapeHtml(r.type || "")}</div>
      <div class="td col-seller ${statusClass}" title="${escapeHtml(r.seller || "")}">${escapeHtml(r.seller || "")}</div>
      <div class="td col-amount ${statusClass}">${escapeHtml(r.amount || "")}</div>
      <div class="td col-status ${statusClass}">${escapeHtml(statusText)}</div>
    `;
    tbody.appendChild(row);
  });
}

function sortTable(col) {
  if (state.sortCol === col) {
    state.sortReverse = !state.sortReverse;
  } else {
    state.sortCol = col;
    state.sortReverse = false;
  }

  document.querySelectorAll(".th").forEach(th => {
    th.classList.remove("sort-asc", "sort-desc");
    if (th.dataset.col === col) {
      th.classList.add(state.sortReverse ? "sort-desc" : "sort-asc");
    }
  });

  const sorted = [...state.records].sort((a, b) => {
    let va = a[col] || "";
    let vb = b[col] || "";
    if (typeof va === "number" && typeof vb === "number") {
      return state.sortReverse ? vb - va : va - vb;
    }
    va = String(va);
    vb = String(vb);
    if (state.sortReverse) {
      return vb.localeCompare(va, "zh-CN");
    }
    return va.localeCompare(vb, "zh-CN");
  });

  renderTable(sorted);
}

// ── 统计卡片 ───────────────────────────────────────────────

function renderStats(stats) {
  const container = $("stats-area");
  container.innerHTML = "";
  const keys = ["total", "complete", "partial", "failed", "not_invoice"];
  keys.forEach(key => {
    const style = STAT_STYLE[key];
    const chip = document.createElement("div");
    chip.className = `stat-chip ${key === "failed" ? "failed" : ""}`;
    chip.innerHTML = `
      <span class="stat-title">${style.title}</span>
      <span class="stat-value" style="color:${style.accent}">${stats[key] || 0}</span>
    `;
    container.appendChild(chip);
  });
}

// ── 操作栏 ──────────────────────────────────────────────────

function updateStatus(msg) {
  $("status-text").textContent = msg;
}

function setProgress(value) {
  $("progress-fill").style.width = (Math.max(0, Math.min(1, value)) * 100) + "%";
}

function togglePreviewMode() {
  state.previewMode = !state.previewMode;
  updatePreviewSwitch();
  py("set_preview_mode", state.previewMode);
  updateRenameButton();
}

function updatePreviewSwitch() {
  const track = $("preview-switch");
  if (state.previewMode) {
    track.classList.add("on");
  } else {
    track.classList.remove("on");
  }
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
  const disabled = state.processing || state.scanning || state.previewMode || state.records.length === 0;
  btn.disabled = disabled;
}

async function onRenameClick() {
  if (state.renameHistory) {
    py("undo_rename");
    state.processing = true;
    updateStatus("撤销中...");
  } else {
    py("on_rename_button_click");
    state.processing = true;
    updateStatus("重命名中...");
  }
  updateRenameButton();
}

async function exportExcel() {
  const result = await py("export_excel");
  if (result && result.ok) {
    updateStatus(`已导出: ${result.path}`);
  } else {
    updateStatus(result?.error || "导出失败");
  }
}

// ── 云端 OCR 设置 ───────────────────────────────────────────

async function openCloudSettings() {
  const settings = await py("get_cloud_settings") || {};
  $("secret-id-input").value = settings.secret_id || "";
  $("secret-key-input").value = settings.secret_key || "";
  $("cloud-enable-input").checked = settings.enabled || false;
  updateUsageInfo(settings.usage || { used: 0, limit: 1000, remaining: 1000 });
  updateSwitchState();
  $("settings-overlay").classList.add("show");
}

function closeCloudSettings() {
  $("settings-overlay").classList.remove("show");
}

async function saveCloudSettings() {
  const sid = $("secret-id-input").value.trim();
  const skey = $("secret-key-input").value.trim();
  if (!sid || !skey) {
    alert("请先输入 SecretId 和 SecretKey");
    return;
  }
  const enabled = $("cloud-enable-input").checked;
  const result = await py("save_cloud_settings", sid, skey, enabled);
  if (result && result.ok) {
    state.cloud = result.cloud;
    updateCloudButton();
    updateRenameButton();
    closeCloudSettings();
  }
}

async function clearCloudSettings() {
  const result = await py("clear_cloud_settings");
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
  if (!sid || !skey) {
    alert("请先输入 SecretId 和 SecretKey");
    return;
  }
  const result = await py("verify_cloud_credentials", sid, skey);
  alert(result?.message || "验证结果未知");
}

async function toggleCloudEnabled() {
  const result = await py("toggle_cloud_enabled");
  if (result && result.ok) {
    state.cloud = result.cloud;
    updateCloudButton();
    updateRenameButton();
  }
}

function updateCloudButton() {
  const btn = $("cloud-status-btn");
  btn.className = "cloud-status-btn";
  if (!state.cloud.configured) {
    btn.classList.add("unconfigured");
    btn.textContent = "● 未配置";
  } else if (state.cloud.enabled) {
    btn.classList.add("configured-enabled");
    btn.textContent = "● 已启用";
  } else {
    btn.classList.add("configured-disabled");
    btn.textContent = "● 未启用";
  }
}

function updateUsageInfo(usage) {
  const used = usage.used || 0;
  const remaining = usage.remaining || 0;
  const limit = usage.limit || 1000;
  const el = $("usage-info");
  el.innerHTML = `<div>本月已调用 ${used} 次</div><div>免费额度 ${limit} 次/月</div>`;
  el.className = "usage-info";
  if (remaining <= 0) el.classList.add("danger");
  else if (remaining <= 100) el.classList.add("warning");
}

function updateSwitchState() {
  const sid = $("secret-id-input").value.trim();
  const skey = $("secret-key-input").value.trim();
  const input = $("cloud-enable-input");
  if (!sid || !skey) {
    input.disabled = true;
    input.checked = false;
  } else {
    input.disabled = false;
  }
}

function toggleSecretKeyVisibility() {
  const input = $("secret-key-input");
  const btn = $("toggle-key-btn");
  if (input.type === "password") {
    input.type = "text";
    btn.textContent = "隐藏";
  } else {
    input.type = "password";
    btn.textContent = "显示";
  }
}

// ── Python 事件推送处理 ─────────────────────────────────────

window.__onPyEvent__ = function(payload) {
  if (!payload) return;
  const { event, data } = payload;

  switch (event) {
    case "scan_started":
      state.scanning = true;
      state.records = [];
      renderTable([]);
      renderStats({});
      setProgress(0);
      updateStatus("识别中...");
      break;

    case "scan_progress":
      setProgress(data.current / data.total);
      updateStatus(`识别中... (${data.current}/${data.total})`);
      break;

    case "scan_finished":
      state.scanning = false;
      state.records = data.records || [];
      renderTable(state.records);
      renderStats(data.stats || {});
      setProgress(1);
      updateStatus(data.message || "识别完成");
      updateRenameButton();
      break;

    case "rename_started":
      state.processing = true;
      setProgress(0);
      updateStatus("重命名中...");
      updateRenameButton();
      break;

    case "rename_progress":
      setProgress(data.current / data.total);
      updateStatus(`重命名中... (${data.current}/${data.total})`);
      renderStats({
        total: data.total,
        complete: data.success,
        partial: 0,
        failed: data.failed,
        not_invoice: data.skipped,
      });
      break;

    case "rename_item_done":
      updateRow(data.idx, data);
      break;

    case "rename_finished":
      state.processing = false;
      state.records = data.records || state.records;
      state.renameHistory = data.can_undo;
      renderTable(state.records);
      renderStats(data.stats || {});
      setProgress(1);
      updateStatus(data.message || "重命名完成");
      updateRenameButton();
      break;

    case "undo_started":
      state.processing = true;
      updateStatus("撤销中...");
      updateRenameButton();
      break;

    case "undo_finished":
      state.processing = false;
      state.records = data.records || state.records;
      state.renameHistory = data.can_undo;
      renderTable(state.records);
      renderStats(data.stats || {});
      updateStatus(data.message || "撤销完成");
      updateRenameButton();
      break;

    case "status":
      updateStatus(data.message || "");
      break;
  }
};

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
}

// ── 工具函数 ─────────────────────────────────────────────────

function escapeHtml(text) {
  if (!text) return "";
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function debounce(fn, delay) {
  let timer = null;
  return function(...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

// 启动：等待 pywebview 桥接就绪
if (window.pywebview) {
  init();
} else {
  window.addEventListener("pywebviewready", init);
  document.addEventListener("pywebviewready", init);
  // 兜底：2 秒后如果仍未初始化，尝试直接调用
  setTimeout(() => {
    if (!state._initialized) init();
  }, 2000);
}
