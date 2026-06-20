const APP_VERSION = "OneDrive 送出版 2026-06-20 22:25";

const SHAREPOINT_CONFIG = {
  enabled: false,
  recordsListTitle: "巡檢紀錄",
  photosLibraryTitle: "巡檢照片",
  pageSize: 200,
  fieldMap: {
    title: "Title",
    inspectorName: "InspectorName",
    employeeId: "EmployeeId",
    inspectionDate: "InspectionDate",
    location: "InspectionLocation",
    hasAbnormal: "HasAbnormal",
    abnormalDescription: "AbnormalDescription",
    checkItemsJson: "CheckItemsJson",
    photoLinksJson: "PhotoLinksJson",
  },
};

const ONEDRIVE_CONFIG = {
  enabled: false,
  clientId: "",
  authority: "https://login.microsoftonline.com/common",
  storageFolder: "巡檢網頁",
  recordsFileName: "inspection-records.json",
  photoFolderName: "photos",
  scopes: ["User.Read", "Files.ReadWrite"],
};

const POWER_AUTOMATE_CONFIG = {
  enabled: true,
  submitEndpointUrl: "https://default8660009a103e42099ff744105f8112.c5.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/07ade87a20454b35af289262c215f185/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=fUNe20p34kXoGyyofm2JGkDdVbjYH_MTLNEHGLSRhwk",
  queryEndpointUrl: "",
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
const LOCAL_STORAGE_KEY = "department-inspection-records-v1";

const form = document.querySelector("#inspectionForm");
const queryForm = document.querySelector("#queryForm");
const checkGroups = document.querySelector("#checkGroups");
const formStatus = document.querySelector("#formStatus");
const connectionText = document.querySelector("#connectionText");
const resultBody = document.querySelector("#resultBody");
const resultCount = document.querySelector("#resultCount");
const exportRecordsButton = document.querySelector("#exportRecords");
const dialog = document.querySelector("#detailDialog");
const detailContent = document.querySelector("#detailContent");

let recordsCache = [];
let service;

function today() {
  return new Date().toISOString().slice(0, 10);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizeDate(value) {
  if (!value) return "";
  return String(value).slice(0, 10);
}

function setStatus(message, type = "") {
  formStatus.textContent = message;
  formStatus.className = type;
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
              <input type="radio" name="${escapeHtml(item.key)}" value="${option}" ${index === 0 ? "checked" : ""}>
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

function switchView(viewId) {
  document.querySelectorAll(".view").forEach((view) => view.classList.toggle("active", view.id === viewId));
  document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("active", tab.dataset.view === viewId));
  if (viewId === "queryView") {
    loadAndRenderRecords();
  }
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
  const hasAbnormal = checks.some((item) => item.value === "異常");
  const abnormalDescription = String(data.get("abnormalDescription") || "").trim();
  const photos = Array.from(form.elements.photos.files || []);

  if (hasAbnormal && !abnormalDescription) {
    throw new Error("有異常項目時，請填寫異常情況描述或辦理情形。");
  }
  if (hasAbnormal && photos.length === 0) {
    throw new Error("有異常項目時，請至少上傳 1 張照片。");
  }

  return {
    inspectorName: String(data.get("inspectorName") || "").trim(),
    employeeId: String(data.get("employeeId") || "").trim(),
    inspectionDate: String(data.get("inspectionDate") || ""),
    location: String(data.get("location") || "").trim(),
    abnormalDescription,
    hasAbnormal,
    checks,
    photos,
  };
}

async function handleSubmit(event) {
  event.preventDefault();
  setStatus("正在送出巡檢紀錄...", "");

  try {
    const payload = getFormPayload();
    await service.createRecord(payload);
    form.reset();
    renderFields(checkFields);
    form.elements.inspectionDate.value = today();
    const message = service.mode === "local"
      ? "巡檢紀錄已暫存在本機瀏覽器。接上 Power Automate 後才會寫入 OneDrive。"
      : "巡檢紀錄已送出。";
    setStatus(message, "success");
    await loadAndRenderRecords(false);
  } catch (error) {
    setStatus(error.message || "送出失敗，請稍後再試。", "error");
  }
}

function exportRecords() {
  const data = JSON.stringify(recordsCache, null, 2);
  const blob = new Blob([data], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const date = today().replaceAll("-", "");
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `巡檢紀錄-${date}.json`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function filterRecords(records) {
  const data = new FormData(queryForm);
  const dateFrom = String(data.get("dateFrom") || "");
  const dateTo = String(data.get("dateTo") || "");
  const location = String(data.get("location") || "").trim().toLowerCase();
  const inspectorName = String(data.get("inspectorName") || "").trim().toLowerCase();
  const employeeId = String(data.get("employeeId") || "").trim().toLowerCase();
  const abnormal = String(data.get("abnormal") || "");

  return records.filter((record) => {
    const date = normalizeDate(record.inspectionDate);
    if (dateFrom && date < dateFrom) return false;
    if (dateTo && date > dateTo) return false;
    if (location && !String(record.location || "").toLowerCase().includes(location)) return false;
    if (inspectorName && !String(record.inspectorName || "").toLowerCase().includes(inspectorName)) return false;
    if (employeeId && !String(record.employeeId || "").toLowerCase().includes(employeeId)) return false;
    if (abnormal === "yes" && !record.hasAbnormal) return false;
    if (abnormal === "no" && record.hasAbnormal) return false;
    return true;
  });
}

function renderResults(records) {
  const filtered = filterRecords(records);
  resultCount.textContent = `共 ${filtered.length} 筆`;

  if (filtered.length === 0) {
    resultBody.innerHTML = `<tr><td colspan="7">沒有符合條件的巡檢紀錄。</td></tr>`;
    return;
  }

  resultBody.innerHTML = filtered.map((record) => {
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

async function loadAndRenderRecords(showLoading = true) {
  if (showLoading) {
    resultCount.textContent = "載入中...";
    resultBody.innerHTML = `<tr><td colspan="7">正在讀取巡檢紀錄。</td></tr>`;
  }

  try {
    recordsCache = await service.getRecords();
    renderResults(recordsCache);
  } catch (error) {
    resultCount.textContent = "讀取失敗";
    resultBody.innerHTML = `<tr><td colspan="7">${escapeHtml(error.message || "無法讀取巡檢資料。")}</td></tr>`;
  }
}

function renderDetail(record) {
  const photos = record.photos || [];
  const abnormalDescription = record.abnormalDescription || "無";
  detailContent.innerHTML = `
    <div class="detail-grid">
      <div class="detail-box"><strong>巡檢日期</strong>${escapeHtml(normalizeDate(record.inspectionDate))}</div>
      <div class="detail-box"><strong>巡檢地點</strong>${escapeHtml(record.location)}</div>
      <div class="detail-box"><strong>姓名</strong>${escapeHtml(record.inspectorName)}</div>
      <div class="detail-box"><strong>員工編號</strong>${escapeHtml(record.employeeId)}</div>
    </div>
    <div class="detail-box wide"><strong>異常情況描述或其它辦理情形</strong>${escapeHtml(abnormalDescription)}</div>
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
      ${photos.length ? photos.map((photo) => `
        <a href="${escapeHtml(photo.url)}" target="_blank" rel="noreferrer">
          ${photo.previewUrl ? `<img src="${escapeHtml(photo.previewUrl)}" alt="${escapeHtml(photo.name)}">` : ""}
          <span>${escapeHtml(photo.name || "巡檢照片")}</span>
        </a>
      `).join("") : "<p>沒有照片。</p>"}
    </div>
  `;
}

function getSharePointWebUrl() {
  if (window._spPageContextInfo?.webAbsoluteUrl) return window._spPageContextInfo.webAbsoluteUrl;
  const match = window.location.href.match(/^(https:\/\/[^/]+\/sites\/[^/]+)/i);
  return match ? match[1] : "";
}

function sanitizeFileName(value) {
  return String(value || "photo")
    .replace(/[\\/:*?"<>|#%{}~&]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 90);
}

function odataText(value) {
  return String(value).replaceAll("'", "''");
}

function parseJsonField(value, fallback) {
  if (!value) return fallback;
  try {
    return JSON.parse(value);
  } catch {
    return fallback;
  }
}

async function fileToBase64(file) {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

class LocalInspectionService {
  constructor() {
    this.mode = "local";
  }

  async getRecords() {
    return JSON.parse(localStorage.getItem(LOCAL_STORAGE_KEY) || "[]");
  }

  async createRecord(payload) {
    const records = await this.getRecords();
    const id = `local-${Date.now()}`;
    const photos = payload.photos.map((file) => ({
      name: file.name,
      url: URL.createObjectURL(file),
      previewUrl: URL.createObjectURL(file),
    }));
    records.unshift({
      id,
      title: `${payload.inspectionDate} ${payload.location} ${payload.inspectorName}`,
      inspectorName: payload.inspectorName,
      employeeId: payload.employeeId,
      inspectionDate: payload.inspectionDate,
      location: payload.location,
      hasAbnormal: payload.hasAbnormal,
      abnormalDescription: payload.abnormalDescription,
      checks: payload.checks,
      photos,
      created: new Date().toISOString(),
    });
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(records));
  }
}

class PowerAutomateInspectionService {
  constructor(config) {
    this.mode = "powerautomate";
    this.config = config;
    this.local = new LocalInspectionService();
  }

  async getRecords() {
    if (!this.config.queryEndpointUrl) return this.local.getRecords();

    const response = await fetch(this.config.queryEndpointUrl, {
      method: "GET",
      cache: "no-store",
    });
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`讀取 Power Automate 紀錄失敗 (${response.status})：${text.slice(0, 220)}`);
    }
    const records = await response.json();
    return Array.isArray(records) ? records : [];
  }

  async createRecord(payload) {
    if (!/[?&](sig|code)=/i.test(this.config.submitEndpointUrl)) {
      throw new Error("Power Automate URL 看起來不完整。請重新複製 manual 觸發器的完整 HTTP POST URL，網址通常會包含 sig= 或 code= 參數。");
    }

    const id = `pa-${Date.now()}`;
    const photos = [];
    for (const file of payload.photos) {
      photos.push({
        name: file.name,
        type: file.type || "application/octet-stream",
        contentBase64: await fileToBase64(file),
      });
    }

    const remotePayload = {
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

    const body = JSON.stringify(remotePayload);
    const response = await fetch(this.config.submitEndpointUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json; charset=utf-8" },
      body,
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`送出到 Power Automate 失敗 (${response.status})：${text.slice(0, 220)}`);
    }

    await this.local.createRecord({
      ...payload,
      photos: [],
    });
  }
}

class SharePointInspectionService {
  constructor(webUrl, config) {
    this.mode = "sharepoint";
    this.webUrl = webUrl.replace(/\/$/, "");
    this.config = config;
    this.digest = "";
    this.entityType = "";
    this.libraryRootUrl = "";
  }

  async request(path, options = {}) {
    const response = await fetch(`${this.webUrl}${path}`, {
      credentials: "same-origin",
      headers: {
        Accept: "application/json;odata=verbose",
        ...(options.body && !(options.body instanceof ArrayBuffer) ? { "Content-Type": "application/json;odata=verbose" } : {}),
        ...(options.method && options.method !== "GET" ? { "X-RequestDigest": await this.getDigest() } : {}),
        ...(options.headers || {}),
      },
      ...options,
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`SharePoint 回應失敗 (${response.status})：${text.slice(0, 220)}`);
    }

    if (response.status === 204) return null;
    return response.json();
  }

  async getDigest() {
    if (this.digest) return this.digest;
    const response = await fetch(`${this.webUrl}/_api/contextinfo`, {
      method: "POST",
      credentials: "same-origin",
      headers: { Accept: "application/json;odata=verbose" },
    });
    if (!response.ok) throw new Error("無法取得 SharePoint 寫入權杖。");
    const json = await response.json();
    this.digest = json.d.GetContextWebInformation.FormDigestValue;
    return this.digest;
  }

  async ensureMetadata() {
    if (!this.entityType) {
      const listInfo = await this.request(
        `/_api/web/lists/getbytitle('${odataText(this.config.recordsListTitle)}')?$select=ListItemEntityTypeFullName`,
      );
      this.entityType = listInfo.d.ListItemEntityTypeFullName;
    }
    if (!this.libraryRootUrl) {
      const libraryInfo = await this.request(
        `/_api/web/lists/getbytitle('${odataText(this.config.photosLibraryTitle)}')/RootFolder?$select=ServerRelativeUrl`,
      );
      this.libraryRootUrl = libraryInfo.d.ServerRelativeUrl;
    }
  }

  mapItem(item) {
    const fields = this.config.fieldMap;
    return {
      id: String(item.Id),
      title: item[fields.title],
      inspectorName: item[fields.inspectorName],
      employeeId: item[fields.employeeId],
      inspectionDate: item[fields.inspectionDate],
      location: item[fields.location],
      hasAbnormal: Boolean(item[fields.hasAbnormal]),
      abnormalDescription: item[fields.abnormalDescription],
      checks: parseJsonField(item[fields.checkItemsJson], []),
      photos: parseJsonField(item[fields.photoLinksJson], []),
      created: item.Created,
      author: item.Author?.Title || "",
    };
  }

  async getRecords() {
    const fields = this.config.fieldMap;
    const select = [
      "Id",
      "Created",
      "Author/Title",
      fields.title,
      fields.inspectorName,
      fields.employeeId,
      fields.inspectionDate,
      fields.location,
      fields.hasAbnormal,
      fields.abnormalDescription,
      fields.checkItemsJson,
      fields.photoLinksJson,
    ].join(",");
    const json = await this.request(
      `/_api/web/lists/getbytitle('${odataText(this.config.recordsListTitle)}')/items` +
      `?$select=${select}&$expand=Author&$orderby=${fields.inspectionDate} desc,Id desc&$top=${this.config.pageSize}`,
    );
    return json.d.results.map((item) => this.mapItem(item));
  }

  async createRecord(payload) {
    await this.ensureMetadata();
    const fields = this.config.fieldMap;
    const title = `${payload.inspectionDate} ${payload.location} ${payload.inspectorName}`;
    const body = {
      __metadata: { type: this.entityType },
      [fields.title]: title,
      [fields.inspectorName]: payload.inspectorName,
      [fields.employeeId]: payload.employeeId,
      [fields.inspectionDate]: payload.inspectionDate,
      [fields.location]: payload.location,
      [fields.hasAbnormal]: payload.hasAbnormal,
      [fields.abnormalDescription]: payload.abnormalDescription,
      [fields.checkItemsJson]: JSON.stringify(payload.checks),
      [fields.photoLinksJson]: "[]",
    };

    const created = await this.request(
      `/_api/web/lists/getbytitle('${odataText(this.config.recordsListTitle)}')/items`,
      { method: "POST", body: JSON.stringify(body) },
    );

    const recordId = created.d.Id;
    const photos = await this.uploadPhotos(recordId, payload.photos);
    if (photos.length > 0) {
      await this.request(
        `/_api/web/lists/getbytitle('${odataText(this.config.recordsListTitle)}')/items(${recordId})`,
        {
          method: "POST",
          headers: {
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
          },
          body: JSON.stringify({
            __metadata: { type: this.entityType },
            [fields.photoLinksJson]: JSON.stringify(photos),
          }),
        },
      );
    }
  }

  async uploadPhotos(recordId, files) {
    const photos = [];
    for (const [index, file] of files.entries()) {
      const extension = file.name.includes(".") ? file.name.slice(file.name.lastIndexOf(".")) : "";
      const baseName = sanitizeFileName(file.name.replace(extension, ""));
      const fileName = `${recordId}-${index + 1}-${baseName}${extension}`;
      const buffer = await file.arrayBuffer();
      const json = await this.request(
        `/_api/web/GetFolderByServerRelativeUrl('${odataText(this.libraryRootUrl)}')/Files/add(url='${odataText(fileName)}',overwrite=true)`,
        {
          method: "POST",
          body: buffer,
          headers: { "Content-Type": "application/octet-stream" },
        },
      );
      photos.push({
        name: file.name,
        url: json.d.ServerRelativeUrl,
        previewUrl: json.d.ServerRelativeUrl,
      });
    }
    return photos;
  }
}

class OneDriveInspectionService {
  constructor(config) {
    this.mode = "onedrive";
    this.config = config;
    this.account = null;
    this.client = new window.msal.PublicClientApplication({
      auth: {
        clientId: config.clientId,
        authority: config.authority,
        redirectUri: window.location.origin + window.location.pathname,
      },
      cache: {
        cacheLocation: "localStorage",
      },
    });
  }

  async ensureAccount() {
    const response = await this.client.handleRedirectPromise();
    if (response?.account) this.account = response.account;

    const accounts = this.client.getAllAccounts();
    if (!this.account && accounts.length > 0) this.account = accounts[0];

    if (!this.account) {
      const login = await this.client.loginPopup({ scopes: this.config.scopes });
      this.account = login.account;
    }
  }

  async getAccessToken() {
    await this.ensureAccount();
    try {
      const response = await this.client.acquireTokenSilent({
        account: this.account,
        scopes: this.config.scopes,
      });
      return response.accessToken;
    } catch {
      const response = await this.client.acquireTokenPopup({
        account: this.account,
        scopes: this.config.scopes,
      });
      return response.accessToken;
    }
  }

  async graph(path, options = {}) {
    const token = await this.getAccessToken();
    const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        ...(options.headers || {}),
      },
    });

    if (response.status === 404 && options.allowNotFound) return null;
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`OneDrive 回應失敗 (${response.status})：${text.slice(0, 220)}`);
    }
    if (response.status === 204) return null;
    const contentType = response.headers.get("content-type") || "";
    return contentType.includes("application/json") ? response.json() : response.text();
  }

  recordsPath() {
    return `/me/drive/root:/${encodeURIComponent(this.config.storageFolder)}/${encodeURIComponent(this.config.recordsFileName)}:/content`;
  }

  photoPath(fileName) {
    return `/me/drive/root:/${encodeURIComponent(this.config.storageFolder)}/${encodeURIComponent(this.config.photoFolderName)}/${encodeURIComponent(fileName)}:/content`;
  }

  async getRecords() {
    const content = await this.graph(this.recordsPath(), { allowNotFound: true });
    if (!content) return [];
    const records = typeof content === "string" ? JSON.parse(content) : content;
    return Array.isArray(records) ? records : [];
  }

  async saveRecords(records) {
    await this.graph(this.recordsPath(), {
      method: "PUT",
      body: JSON.stringify(records, null, 2),
      headers: { "Content-Type": "application/json; charset=utf-8" },
    });
  }

  async createRecord(payload) {
    const records = await this.getRecords();
    const id = `od-${Date.now()}`;
    const photos = await this.uploadPhotos(id, payload.photos);
    records.unshift({
      id,
      title: `${payload.inspectionDate} ${payload.location} ${payload.inspectorName}`,
      inspectorName: payload.inspectorName,
      employeeId: payload.employeeId,
      inspectionDate: payload.inspectionDate,
      location: payload.location,
      hasAbnormal: payload.hasAbnormal,
      abnormalDescription: payload.abnormalDescription,
      checks: payload.checks,
      photos,
      created: new Date().toISOString(),
    });
    await this.saveRecords(records);
  }

  async uploadPhotos(recordId, files) {
    const photos = [];
    for (const [index, file] of files.entries()) {
      const extension = file.name.includes(".") ? file.name.slice(file.name.lastIndexOf(".")) : "";
      const baseName = sanitizeFileName(file.name.replace(extension, ""));
      const fileName = `${recordId}-${index + 1}-${baseName}${extension}`;
      const item = await this.graph(this.photoPath(fileName), {
        method: "PUT",
        body: await file.arrayBuffer(),
        headers: { "Content-Type": file.type || "application/octet-stream" },
      });
      photos.push({
        name: file.name,
        url: item.webUrl,
        previewUrl: item.webUrl,
      });
    }
    return photos;
  }
}

function createService() {
  const webUrl = getSharePointWebUrl();
  if (POWER_AUTOMATE_CONFIG.enabled && POWER_AUTOMATE_CONFIG.submitEndpointUrl) {
    return new PowerAutomateInspectionService(POWER_AUTOMATE_CONFIG);
  }
  if (ONEDRIVE_CONFIG.enabled && ONEDRIVE_CONFIG.clientId && window.msal) {
    return new OneDriveInspectionService(ONEDRIVE_CONFIG);
  }
  if (SHAREPOINT_CONFIG.enabled && webUrl && window.location.protocol.startsWith("http")) {
    return new SharePointInspectionService(webUrl, SHAREPOINT_CONFIG);
  }
  return new LocalInspectionService();
}

document.querySelectorAll(".tab").forEach((tab) => {
  tab.addEventListener("click", () => switchView(tab.dataset.view));
});

document.querySelector("#markAllNormal").addEventListener("click", () => markAll("正常"));
document.querySelector("#markAllNA").addEventListener("click", () => markAll("不適用"));
document.querySelector("#closeDialog").addEventListener("click", () => dialog.close());
exportRecordsButton.addEventListener("click", exportRecords);

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
  dialog.showModal();
});

renderFields(checkFields);
form.elements.inspectionDate.value = today();
service = createService();
connectionText.textContent = service.mode === "sharepoint"
  ? `SharePoint 模式：將寫入「${SHAREPOINT_CONFIG.recordsListTitle}」清單`
  : service.mode === "powerautomate"
    ? `Power Automate 模式：送出資料將由流程存入 OneDrive｜${APP_VERSION}`
  : service.mode === "onedrive"
    ? `OneDrive 模式：將寫入目前登入帳號的「${ONEDRIVE_CONFIG.storageFolder}」資料夾｜${APP_VERSION}`
  : `本機預覽模式：資料暫存在這台電腦的瀏覽器中，共 ${checkFields.length} 個巡檢項目｜${APP_VERSION}`;
loadAndRenderRecords(false);
