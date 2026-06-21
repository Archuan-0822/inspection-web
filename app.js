const APP_VERSION = "OneDrive 送出版 v13";
const PAGE_LOAD_TIME = new Date();

const POWER_AUTOMATE_CONFIG = {
  enabled: true,
  submitEndpointUrl: "https://default8660009a103e42099ff744105f8112.c5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/07ade87a20454b35af289262c215f185/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fUNe20p34kXoGyyofm2JGkDdVbjYH_MTLNEHGLSRhwk",
};

const checkFields = [
  ["公用設施", "安全門及樓梯間無影響進出之障礙物"],
  ["公用設施", "升降機門開關正常，運行中無停頓或異常聲音"],
  ["公用設施", "地面及通道無可能造成人員滑倒之油漬或積水"],
  ["公用設施", "玻璃門、鐵捲門開啟至定位，無人員碰撞危險"],
  ["公用設施", "消防設備無受障礙物阻擋(使用期限)"],
  ["公用設施", "照明充足 / 空調系統正常運轉"],
  ["用電安全", "電源線、延長線無破損、糾結或超載使用情形"],
  ["用電安全", "潮濕場所電源端有裝置漏電斷路器"],
  ["用電安全", "施工場所臨時用電絕緣完整，線路配接良好(未使用應斷電)"],
  ["機械設備", "車輛機具依規定停放，並定期實施自動檢查"],
  ["機械設備", "起重機、堆高機、高空工作車等機具，由合格人員操作"],
  ["機械設備", "機械設備之安全防護完整，不得任意拆除"],
  ["物料儲存", "物品置放牢靠，無掉落、崩塌傷人之危險"],
  ["物料儲存", "物料置放，無阻礙通道及可能受碰撞影響安全"],
  ["機邊作業", "航機周邊是否有適當隔離或防護，以避免碰撞"],
  ["機邊作業", "車輛有依規定設置輪擋、閃光警示燈、接地等安全防護措施"],
  ["機邊作業", "梯架、作業機具護欄架設完整，無墜落危險"],
  ["機邊作業", "滾帶車運轉中，人員不得踩踏於輸送帶上"],
  ["機邊作業", "作業人員依規定穿著反光背心及使用耳塞、耳罩等聽力防護具"],
  ["承攬商施工區", "施工區域有適當隔離、標示及防護措施"],
  ["承攬商施工區", "動火作業有申請動火許可，並執行安全措施(張貼告示、消防設備、易燃物隔離等)"],
  ["承攬商施工區", "高處作業有依規定設置上下設備及防墜(安全帽、安全帶或安全網、母索)措施，並確實使用"],
  ["承攬商施工區", "電焊機配置有自動電擊防止裝置"],
  ["承攬商施工區", "危險機械及其操作人員有合格證照"],
  ["承攬商施工區", "局限空間或其他危險作業，監視人員在場監督"],
  ["承攬商施工區", "作業人員有確實使用適當個人防護裝備"],
  ["承攬商施工區", "有機溶劑/化學品有妥善管制"],
  ["作業環境及其他應注意事項", "吸煙區以外區域，禁止吸煙"],
  ["作業環境及其他應注意事項", "作業人員不安全行為或動作"],
  ["作業環境及其他應注意事項", "餐廳環境衛生及食品安全管理"],
  ["作業環境及其他應注意事項", "行人及車輛交通安全"],
].map(([category, label], index) => ({
  key: `check-${index + 1}`,
  category,
  label,
}));

const CHECK_OPTIONS = ["正常", "異常", "不適用"];
const FIXED_LOCATIONS = ["華航園區", "1航廈", "2航廈"];
const EMPLOYEE_BY_NAME = {
  房家儀: "630951",
  張春輝: "633653",
  詹子穗: "630628",
  馬智仁: "635529",
  張煒: "645788",
  許澤泉: "646425",
};
const STORAGE_KEYS = [
  "department-inspection-records-v3",
  "department-inspection-records-v2",
  "department-inspection-records-v1",
];
const PRIMARY_STORAGE_KEY = STORAGE_KEYS[0];

const form = document.querySelector("#inspectionForm");
const queryForm = document.querySelector("#queryForm");
const checkGroups = document.querySelector("#checkGroups");
const checkSummary = document.querySelector("#checkSummary");
const formStatus = document.querySelector("#formStatus");
const connectionText = document.querySelector("#connectionText");
const resultBody = document.querySelector("#resultBody");
const resultCount = document.querySelector("#resultCount");
const exportRecordsButton = document.querySelector("#exportRecords");
const clearQueryButton = document.querySelector("#clearQuery");
const otherLocationField = document.querySelector("#otherLocationField");
const detailDialog = document.querySelector("#detailDialog");
const detailContent = document.querySelector("#detailContent");
const successDialog = document.querySelector("#successDialog");

let recordsCache = [];

function today() {
  return formatInputDate(new Date());
}

function formatInputDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDateTime(date) {
  return new Intl.DateTimeFormat("zh-TW", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).format(date);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function toCanonicalDate(value) {
  if (!value) return "";
  const match = String(value).trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (!match) return "";
  const [, year, month, day] = match;
  return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function normalizeDate(value) {
  const canonical = toCanonicalDate(value);
  return canonical ? canonical.replaceAll("-", "/") : "";
}

function assertDateValue(value, label) {
  if (!toCanonicalDate(value)) throw new Error(`請選擇${label}。`);
}

function updateEmployeeId() {
  const name = form.elements.inspectorName.value;
  form.elements.employeeId.value = EMPLOYEE_BY_NAME[name] || "";
}

function updateQueryEmployeeId() {
  const name = queryForm.elements.inspectorName.value;
  queryForm.elements.employeeId.value = EMPLOYEE_BY_NAME[name] || "";
}

function setStatus(message, type = "") {
  formStatus.textContent = message;
  formStatus.className = type;
}

function readRecordsFromKey(key) {
  try {
    const records = JSON.parse(localStorage.getItem(key) || "[]");
    if (Array.isArray(records)) return records;
    return records?.id ? [records] : [];
  } catch {
    return [];
  }
}

function dedupeRecords(records) {
  const seen = new Set();
  return records.filter((record) => {
    if (!record?.id || seen.has(record.id)) return false;
    seen.add(record.id);
    return true;
  });
}

function normalizeStoredRecord(record) {
  return {
    ...record,
    photos: (record.photos || []).map((photo) => ({
      name: photo.name || photo.storedName || "巡檢照片",
      storedName: photo.storedName || photo.name || "",
      url: "",
      previewUrl: String(photo.previewUrl || "").startsWith("data:image/") ? photo.previewUrl : "",
    })),
  };
}

function getStoredRecords() {
  const records = dedupeRecords(STORAGE_KEYS.flatMap((key) => readRecordsFromKey(key)))
    .map(normalizeStoredRecord);
  records.sort((a, b) => String(b.created || "").localeCompare(String(a.created || "")));
  localStorage.setItem(PRIMARY_STORAGE_KEY, JSON.stringify(records));
  return records;
}

function saveStoredRecord(record) {
  const records = dedupeRecords([record, ...getStoredRecords()]);
  localStorage.setItem(PRIMARY_STORAGE_KEY, JSON.stringify(records));
  recordsCache = records;
}

function groupByCategory(fields) {
  return fields.reduce((groups, field) => {
    if (!groups.has(field.category)) groups.set(field.category, []);
    groups.get(field.category).push(field);
    return groups;
  }, new Map());
}

function renderFields(fields) {
  checkGroups.innerHTML = "";
  for (const [category, items] of groupByCategory(fields)) {
    const group = document.createElement("section");
    group.className = "check-group";
    group.innerHTML = `<h3>${escapeHtml(category)}</h3>`;

    for (const item of items) {
      const row = document.createElement("div");
      row.className = "check-item";
      row.innerHTML = `
        <div class="check-label">${escapeHtml(item.label)}</div>
        <div class="segmented" role="radiogroup" aria-label="${escapeHtml(item.label)}">
          ${CHECK_OPTIONS.map((option, index) => `
            <label>
              <input type="radio" name="${escapeHtml(item.key)}" value="${option}" ${option === "不適用" ? "checked" : ""}>
              <span>${option}</span>
            </label>
          `).join("")}
        </div>
      `;
      group.append(row);
    }

    checkGroups.append(group);
  }
}

function markAll(value) {
  for (const field of checkFields) {
    const input = form.querySelector(`input[name="${CSS.escape(field.key)}"][value="${value}"]`);
    if (input) input.checked = true;
  }
}

function setCheckMode(mode) {
  const isDetailMode = mode === "異常";
  checkGroups.hidden = !isDetailMode;
  checkGroups.classList.toggle("hidden-field", !isDetailMode);

  if (mode === "全部正常") {
    markAll("正常");
    checkSummary.textContent = "目前已設為全部正常。若需逐項標示異常，請點選「異常」。";
  } else if (mode === "不適用") {
    markAll("不適用");
    checkSummary.textContent = "目前預設為不適用。若有異常，請點選「異常」展開逐項檢查。";
  } else {
    checkSummary.textContent = "請在下方逐項選擇正常、異常或不適用。";
  }
}

function switchView(viewId) {
  document.querySelectorAll(".view").forEach((view) => view.classList.toggle("active", view.id === viewId));
  document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === viewId));
  if (viewId === "queryView") loadAndRenderRecords();
}

function updateLocationOtherVisibility() {
  const isOther = form.elements.locationPreset.value === "其它";
  otherLocationField.classList.toggle("hidden-field", !isOther);
  otherLocationField.hidden = !isOther;
  form.elements.locationOther.required = isOther;
  if (!isOther) form.elements.locationOther.value = "";
}

function getSelectedLocation(data) {
  const preset = String(data.get("locationPreset") || "").trim();
  const other = String(data.get("locationOther") || "").trim();
  return preset === "其它" ? other : preset;
}

function getCheckResults(data) {
  return checkFields.map((field) => ({
    key: field.key,
    category: field.category,
    label: field.label,
    value: String(data.get(field.key) || "正常"),
  }));
}

function getFormPayload() {
  const data = new FormData(form);
  const checks = getCheckResults(data);
  const location = getSelectedLocation(data);
  const hasAbnormal = checks.some((item) => item.value === "異常");
  const abnormalDescription = String(data.get("abnormalDescription") || "").trim();
  const photos = Array.from(form.elements.photos.files || []);

  if (!location) throw new Error("請選擇或填寫巡檢地點。");
  assertDateValue(data.get("inspectionDate"), "巡檢日期");
  if (hasAbnormal && !abnormalDescription) {
    throw new Error("有異常項目時，請填寫異常情況描述或辦理情形。");
  }
  if (hasAbnormal && photos.length === 0) {
    throw new Error("有異常項目時，請至少上傳 1 張照片。");
  }

  return {
    inspectorName: String(data.get("inspectorName") || "").trim(),
    employeeId: String(data.get("employeeId") || "").trim(),
    inspectionDate: normalizeDate(data.get("inspectionDate")),
    location,
    abnormalDescription,
    hasAbnormal,
    checks,
    photos,
  };
}

async function fileToBase64(file) {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

async function buildRemotePayload(payload) {
  const id = `pa-${Date.now()}`;
  const photos = [];
  for (const file of payload.photos) {
    photos.push({
      name: file.name,
      type: file.type || "application/octet-stream",
      contentBase64: await fileToBase64(file),
    });
  }

  return {
    id,
    submittedAt: new Date().toISOString(),
    record: {
      inspectorName: payload.inspectorName,
      employeeId: payload.employeeId,
      inspectionDate: payload.inspectionDate,
      location: payload.location,
      hasAbnormal: payload.hasAbnormal,
      abnormalDescription: payload.abnormalDescription,
      checks: payload.checks,
    },
    photos,
  };
}

async function submitToPowerAutomate(remotePayload) {
  if (!/[?&](sig|code)=/i.test(POWER_AUTOMATE_CONFIG.submitEndpointUrl)) {
    throw new Error("Power Automate URL 不完整，請重新複製完整 HTTP POST URL。");
  }

  const response = await fetch(POWER_AUTOMATE_CONFIG.submitEndpointUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(remotePayload),
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`送出到 Power Automate 失敗 (${response.status})：${text.slice(0, 220)}`);
  }
}

function toLocalRecord(remotePayload) {
  const photoNames = remotePayload.photos.map((photo) => `${remotePayload.id}-${photo.name}`);
  return {
    id: remotePayload.id,
    title: `${remotePayload.record.inspectionDate} ${remotePayload.record.location} ${remotePayload.record.inspectorName}`,
    inspectorName: remotePayload.record.inspectorName,
    employeeId: remotePayload.record.employeeId,
    inspectionDate: remotePayload.record.inspectionDate,
    location: remotePayload.record.location,
    hasAbnormal: remotePayload.record.hasAbnormal,
    abnormalDescription: remotePayload.record.abnormalDescription,
    checks: remotePayload.record.checks,
    photos: remotePayload.photos.map((photo, index) => ({
      name: photo.name,
      storedName: photoNames[index],
      url: "",
      previewUrl: photo.type.startsWith("image/") ? `data:${photo.type};base64,${photo.contentBase64}` : "",
    })),
    created: remotePayload.submittedAt,
  };
}

function showSuccessDialog() {
  successDialog.showModal();
}

async function handleSubmit(event) {
  event.preventDefault();
  setStatus("正在送出巡檢紀錄...", "");

  try {
    const payload = getFormPayload();
    const remotePayload = await buildRemotePayload(payload);
    await submitToPowerAutomate(remotePayload);
    saveStoredRecord(toLocalRecord(remotePayload));

    form.reset();
    renderFields(checkFields);
    form.elements.inspectionDate.value = today();
    updateEmployeeId();
    updateLocationOtherVisibility();
    setCheckMode("不適用");
    setStatus("已完成填報。", "success");
    showSuccessDialog();
    loadAndRenderRecords(false);
  } catch (error) {
    setStatus(error.message || "送出失敗，請稍後再試。", "error");
  }
}

function exportRecords() {
  const records = filterRecords(recordsCache);
  const headers = ["巡檢日期", "巡檢地點", "姓名", "員工編號", "異常狀態", "異常說明", "照片數", "照片檔名", "送出時間", "檢查項目"];
  const rows = records.map((record) => [
    normalizeDate(record.inspectionDate),
    record.location || "",
    record.inspectorName || "",
    record.employeeId || "",
    record.hasAbnormal ? "有異常" : "正常",
    record.abnormalDescription || "",
    String((record.photos || []).length),
    (record.photos || []).map((photo) => photo.storedName || photo.name).join("、"),
    record.created ? formatDateTime(new Date(record.created)) : "",
    (record.checks || []).map((item) => `${item.category} - ${item.label}: ${item.value}`).join("\n"),
  ]);
  const tableRows = [headers, ...rows].map((row) => (
    `<tr>${row.map((cell) => `<td>${escapeHtml(cell).replaceAll("\n", "<br>")}</td>`).join("")}</tr>`
  )).join("");
  const html = `<!doctype html><html><head><meta charset="utf-8"></head><body><table>${tableRows}</table></body></html>`;
  const blob = new Blob(["\ufeff", html], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const date = today().replaceAll("-", "");
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `巡檢紀錄-${date}.xls`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function filterRecords(records) {
  const data = new FormData(queryForm);
  const dateFrom = toCanonicalDate(data.get("dateFrom"));
  const dateTo = toCanonicalDate(data.get("dateTo"));
  const location = String(data.get("location") || "").trim();
  const inspectorName = String(data.get("inspectorName") || "").trim().toLowerCase();
  const employeeId = String(data.get("employeeId") || "").trim().toLowerCase();
  const abnormal = String(data.get("abnormal") || "");

  return records.filter((record) => {
    const date = toCanonicalDate(record.inspectionDate);
    if (dateFrom && date < dateFrom) return false;
    if (dateTo && date > dateTo) return false;
    if (location === "其它" && FIXED_LOCATIONS.includes(String(record.location || ""))) return false;
    if (location && location !== "其它" && record.location !== location) return false;
    if (inspectorName && !String(record.inspectorName || "").toLowerCase().includes(inspectorName)) return false;
    if (employeeId && !String(record.employeeId || "").toLowerCase().includes(employeeId)) return false;
    if (abnormal === "yes" && !record.hasAbnormal) return false;
    if (abnormal === "no" && record.hasAbnormal) return false;
    return true;
  });
}

function renderResults(records, options = {}) {
  const filtered = filterRecords(records);
  const visible = options.latestOnly ? filtered.slice(0, 1) : filtered;
  resultCount.textContent = options.latestOnly && filtered.length > 1
    ? `顯示最新 1 筆，共 ${filtered.length} 筆`
    : `共 ${filtered.length} 筆`;

  if (filtered.length === 0) {
    resultBody.innerHTML = `<tr><td colspan="7">沒有符合條件的巡檢紀錄。</td></tr>`;
    return;
  }

  resultBody.innerHTML = visible.map((record) => {
    const photos = record.photos || [];
    return `
      <tr>
        <td>${escapeHtml(normalizeDate(record.inspectionDate))}</td>
        <td>${escapeHtml(record.location)}</td>
        <td>${escapeHtml(record.inspectorName)}</td>
        <td>${escapeHtml(record.employeeId)}</td>
        <td><span class="badge ${record.hasAbnormal ? "bad" : "ok"}">${record.hasAbnormal ? "有異常" : "正常"}</span></td>
        <td>${photos.length ? `${photos.length} 張` : "無"}</td>
        <td><button class="link-button" data-detail-id="${escapeHtml(record.id)}" type="button">明細</button></td>
      </tr>
    `;
  }).join("");
}

function loadAndRenderRecords(showLoading = true) {
  if (showLoading) {
    resultCount.textContent = "載入中...";
    resultBody.innerHTML = `<tr><td colspan="7">正在讀取本瀏覽器送出的巡檢紀錄。</td></tr>`;
  }
  recordsCache = getStoredRecords();
  renderResults(recordsCache, { latestOnly: true });
}

function renderPhoto(photo) {
  const name = escapeHtml(photo.name || "巡檢照片");
  const storedName = escapeHtml(photo.storedName || photo.name || "");
  if (photo.previewUrl) {
    return `
      <a class="photo-preview" href="${escapeHtml(photo.previewUrl)}" target="_blank" rel="noopener">
        <img src="${escapeHtml(photo.previewUrl)}" alt="${name}">
        <span>${name}</span>
      </a>
    `;
  }

  return `
    <div class="photo-file">
      <strong>${name}</strong>
      <span>照片已存至 OneDrive / 巡檢網頁 / photos / ${storedName}</span>
      <small>這筆舊資料沒有保留可直接預覽的照片內容。</small>
    </div>
  `;
}

function renderDetail(record) {
  const photos = record.photos || [];
  detailContent.innerHTML = `
    <div class="detail-grid">
      <div class="detail-box"><strong>紀錄編號</strong>${escapeHtml(record.id)}</div>
      <div class="detail-box"><strong>巡檢日期</strong>${escapeHtml(normalizeDate(record.inspectionDate))}</div>
      <div class="detail-box"><strong>巡檢地點</strong>${escapeHtml(record.location)}</div>
      <div class="detail-box"><strong>姓名</strong>${escapeHtml(record.inspectorName)}</div>
      <div class="detail-box"><strong>員工編號</strong>${escapeHtml(record.employeeId)}</div>
      <div class="detail-box"><strong>送出時間</strong>${escapeHtml(formatDateTime(new Date(record.created)))}</div>
    </div>
    <div class="detail-box wide"><strong>異常情況描述或其它辦理情形</strong>${escapeHtml(record.abnormalDescription || "無")}</div>
    <h3 class="detail-title">巡檢項目</h3>
    <div class="detail-list">
      ${(record.checks || []).map((item) => `
        <div class="detail-row">
          <div>
            <strong>${escapeHtml(item.category)}</strong>
            <span>${escapeHtml(item.label)}</span>
          </div>
          <span class="badge ${item.value === "異常" ? "bad" : "ok"}">${escapeHtml(item.value)}</span>
        </div>
      `).join("")}
    </div>
    <h3 class="detail-title">照片</h3>
    <div class="photos">
      ${photos.length ? photos.map(renderPhoto).join("") : "<p>沒有照片。</p>"}
    </div>
  `;
}

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => switchView(tab.dataset.view));
});

document.querySelector("#markAllNormal").addEventListener("click", () => setCheckMode("全部正常"));
document.querySelector("#markAllNA").addEventListener("click", () => setCheckMode("不適用"));
document.querySelector("#showAbnormalChecks").addEventListener("click", () => setCheckMode("異常"));
document.querySelector("#closeDialog").addEventListener("click", () => detailDialog.close());
document.querySelector("#closeSuccessDialog").addEventListener("click", () => successDialog.close());
form.elements.inspectorName.addEventListener("change", updateEmployeeId);
queryForm.elements.inspectorName.addEventListener("change", updateQueryEmployeeId);
form.elements.locationPreset.addEventListener("change", updateLocationOtherVisibility);
exportRecordsButton.addEventListener("click", exportRecords);
clearQueryButton.addEventListener("click", () => {
  queryForm.reset();
  updateQueryEmployeeId();
  renderResults(recordsCache, { latestOnly: true });
});

form.addEventListener("submit", handleSubmit);
queryForm.addEventListener("submit", (event) => {
  event.preventDefault();
  renderResults(recordsCache);
});

resultBody.addEventListener("click", (event) => {
  const button = event.target.closest("[data-detail-id]");
  if (!button) return;
  const record = recordsCache.find((item) => String(item.id) === button.dataset.detailId);
  if (!record) return;
  renderDetail(record);
  detailDialog.showModal();
});

renderFields(checkFields);
form.elements.inspectionDate.value = today();
updateEmployeeId();
updateQueryEmployeeId();
setCheckMode("不適用");
otherLocationField.hidden = true;
updateLocationOtherVisibility();
recordsCache = getStoredRecords();
connectionText.textContent = `Power Automate 模式：送出資料將由流程存入 OneDrive｜${APP_VERSION}｜頁面載入 ${formatDateTime(PAGE_LOAD_TIME)}`;
loadAndRenderRecords(false);
