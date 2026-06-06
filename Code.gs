/**********************************************************************
 * 學校教師事件紀錄系統 — 附件雲端後端（Google Apps Script）
 *
 * 功能：接收前端上傳的照片 / 影片 / 檔案，存進 Google Drive，
 *      並把「附件清單」寫進這份試算表的「附件」工作表，供跨裝置調閱。
 *
 * ── 部署步驟（只做一次）──────────────────────────────────────
 * 1. 開一個新的 Google 試算表（這就是你的附件清單資料庫）。
 * 2. 在 Drive 建一個資料夾（例如「教師事件附件」），打開它，
 *    網址列 .../folders/【這串就是 FOLDER_ID】，複製貼到下面。
 * 3. 試算表 → 擴充功能 → Apps Script，把本檔內容整段貼上。
 * 4. 修改下面的 FOLDER_ID。
 * 5. 右上「部署」→「新增部署作業」→ 類型選「網頁應用程式」：
 *      - 執行身分：我（你的帳號）
 *      - 誰可以存取：所有人
 *    按「部署」，授權後複製產生的 /exec 網址。
 * 6. 回到系統 → 系統設定 → Google Drive 雲端附件，貼上該網址並儲存。
 *
 * ⚠ 個人 Gmail 注意：上傳的檔案會設為「知道連結者可檢視」。
 *   本系統含學生影像等敏感個資，請限定人員存取、妥善保管網址，
 *   正式上線建議改用學校 Google Workspace 帳號。
 *********************************************************************/

const FOLDER_ID = "在這裡貼上你的_Drive_資料夾_ID";
const SHEET_NAME = "附件";
const DATA_SHEET = "資料";          // 存所有紀錄與帳號（多裝置同步）
const TOKEN_TTL_MS = 12 * 3600 * 1000; // 登入 token 有效 12 小時
const TZ = "Asia/Taipei";
const DATA_ENTITIES = ["teachers","events","counsel","students","stuEvents","audit"];

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["id", "owner", "name", "mimeType", "size", "fileId", "url", "thumb", "by", "at"]);
  }
  return sh;
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    switch (req.action) {
      // ── 附件（既有）──
      case "upload":  out = uploadFile(req); break;
      case "list":    out = listOwner(req.owner); break;
      case "listAll": out = listAll(); break;
      case "delete":  out = deleteFile(req.id); break;
      // ── 資料同步 / 帳號（新增）──
      case "status":  out = apiStatus(); break;
      case "init":    out = apiInit(req); break;
      case "login":   out = apiLogin(req); break;
      case "logout":  out = apiLogout(req); break;
      case "pull":    out = apiPull(req); break;
      case "push":    out = apiPush(req); break;
      case "remove":  out = apiRemove(req); break;
      default:        out = { ok: false, error: "未知的動作：" + req.action };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
                       .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput("教師事件附件後端運作中 ✔");
}

function uploadFile(req) {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const bytes = Utilities.base64Decode(req.data);
  const blob = Utilities.newBlob(bytes, req.mimeType || "application/octet-stream", req.name || "file");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const fileId = file.getId();
  const id = "att" + Date.now() + Math.floor(Math.random() * 10000);
  const url = "https://drive.google.com/uc?export=view&id=" + fileId;
  const thumb = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w240";
  const at = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Taipei", "yyyy/MM/dd HH:mm:ss");

  getSheet().appendRow([id, req.owner, req.name, req.mimeType, req.size, fileId, url, thumb, req.by, at]);

  return {
    ok: true,
    item: { id: id, owner: req.owner, name: req.name, type: req.mimeType, size: req.size,
            fileId: fileId, url: url, thumb: thumb, by: req.by, at: at, store: "cloud" }
  };
}

function rowsToItems(values) {
  return values.map(function (r) {
    return { id: r[0], owner: r[1], name: r[2], type: r[3], size: r[4],
             fileId: r[5], url: r[6], thumb: r[7], by: r[8], at: r[9], store: "cloud" };
  });
}

function listAll() {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, items: [] };
  const vals = sh.getRange(2, 1, last - 1, 10).getValues();
  return { ok: true, items: rowsToItems(vals) };
}

function listOwner(owner) {
  return { ok: true, items: listAll().items.filter(function (i) { return i.owner === owner; }) };
}

function deleteFile(id) {
  const sh = getSheet();
  const last = sh.getLastRow();
  if (last < 2) return { ok: false, error: "無資料" };
  const vals = sh.getRange(2, 1, last - 1, 10).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === id) {
      try { DriveApp.getFileById(vals[i][5]).setTrashed(true); } catch (e) {}
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  return { ok: false, error: "找不到該附件" };
}

/* =====================================================================
 * 資料同步 + 帳號登入（多裝置共用）
 * 「資料」工作表欄位：entity | id | json | updatedAt | deleted
 * 帳號以 entity="account" 存放，json 內含 passHash（SHA-256，由前端雜湊）
 * =================================================================== */
function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || TZ, "yyyy/MM/dd HH:mm:ss");
}
function getDataSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DATA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(DATA_SHEET);
    sh.appendRow(["entity", "id", "json", "updatedAt", "deleted"]);
  }
  return sh;
}
function readDataRows_() {
  const sh = getDataSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 5).getValues();
  return vals.map(function (r, i) {
    return { row: i + 2, entity: String(r[0]), id: String(r[1]), json: r[2], updatedAt: r[3], deleted: String(r[4]) };
  });
}
function upsertData_(entity, id, jsonStr) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const sh = getDataSheet_();
    const rows = readDataRows_();
    const at = nowStr_();
    let found = null;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].entity === entity && rows[i].id === String(id)) { found = rows[i]; break; }
    }
    if (found) sh.getRange(found.row, 3, 1, 3).setValues([[jsonStr, at, ""]]);
    else sh.appendRow([entity, String(id), jsonStr, at, ""]);
  } finally { lock.releaseLock(); }
}
function markDeletedData_(entity, id) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const sh = getDataSheet_();
    const rows = readDataRows_();
    for (let i = 0; i < rows.length; i++) {
      if (rows[i].entity === entity && rows[i].id === String(id)) {
        sh.getRange(rows[i].row, 4, 1, 2).setValues([[nowStr_(), "1"]]);
        return;
      }
    }
  } finally { lock.releaseLock(); }
}
function hasAccounts_() {
  return readDataRows_().some(function (r) { return r.entity === "account" && !r.deleted; });
}

/* ---- token（存於 Script Properties）---- */
function makeToken_(acc) {
  const token = Utilities.getUuid().replace(/-/g, "");
  const payload = { u: acc.username, r: acc.role, n: acc.name, exp: Date.now() + TOKEN_TTL_MS };
  PropertiesService.getScriptProperties().setProperty("tok_" + token, JSON.stringify(payload));
  return token;
}
function checkToken_(token) {
  if (!token) return null;
  const raw = PropertiesService.getScriptProperties().getProperty("tok_" + token);
  if (!raw) return null;
  let t; try { t = JSON.parse(raw); } catch (e) { return null; }
  if (Date.now() > t.exp) { PropertiesService.getScriptProperties().deleteProperty("tok_" + token); return null; }
  return t;
}

/* ---- API ---- */
function apiStatus() {
  return { ok: true, hasAccounts: hasAccounts_() };
}
function apiInit(req) {
  if (hasAccounts_()) return { ok: false, error: "雲端已初始化，無法重複初始化" };
  const lock = LockService.getScriptLock();
  lock.tryLock(20000);
  try {
    const sh = getDataSheet_();
    const at = nowStr_();
    const batch = [];
    (req.accounts || []).forEach(function (a) { batch.push(["account", a.username, JSON.stringify(a), at, ""]); });
    const data = req.data || {};
    DATA_ENTITIES.forEach(function (en) {
      (data[en] || []).forEach(function (o) { batch.push([en, o.id, JSON.stringify(o), at, ""]); });
    });
    const ck = data.checklist || {};
    Object.keys(ck).forEach(function (k) { batch.push(["checklist", k, JSON.stringify(ck[k]), at, ""]); });
    if (batch.length) sh.getRange(sh.getLastRow() + 1, 1, batch.length, 5).setValues(batch);
    return { ok: true, count: batch.length };
  } finally { lock.releaseLock(); }
}
function apiLogin(req) {
  const accs = readDataRows_().filter(function (r) { return r.entity === "account" && !r.deleted; });
  let match = null;
  for (let i = 0; i < accs.length; i++) {
    let a; try { a = JSON.parse(accs[i].json); } catch (e) { continue; }
    if (a.username && a.username.toLowerCase() === String(req.username || "").toLowerCase()) { match = a; break; }
  }
  if (!match || match.passHash !== req.passHash) return { ok: false, error: "帳號或密碼錯誤" };
  const token = makeToken_(match);
  return { ok: true, token: token, user: { username: match.username, name: match.name, role: match.role } };
}
function apiLogout(req) {
  if (req.token) PropertiesService.getScriptProperties().deleteProperty("tok_" + req.token);
  return { ok: true };
}
function apiPull(req) {
  const t = checkToken_(req.token);
  if (!t) return { ok: false, auth: false, error: "未授權或登入已逾時，請重新登入" };
  const rows = readDataRows_().filter(function (r) { return !r.deleted; });
  const data = { teachers: [], events: [], counsel: [], students: [], stuEvents: [], audit: [] };
  const checklist = {};
  const accounts = [];
  rows.forEach(function (r) {
    let obj; try { obj = JSON.parse(r.json); } catch (e) { return; }
    if (r.entity === "checklist") checklist[r.id] = obj;
    else if (r.entity === "account") accounts.push(obj); // 含 passHash，供各裝置離線登入
    else if (data[r.entity]) data[r.entity].push(obj);
  });
  data.checklist = checklist;
  data.accounts = accounts;
  return { ok: true, data: data };
}
function apiPush(req) {
  const t = checkToken_(req.token);
  if (!t) return { ok: false, auth: false, error: "未授權" };
  if (req.entity === "account" && t.r !== "admin") return { ok: false, error: "僅系統管理員可變更帳號" };
  if (!req.entity || req.id == null) return { ok: false, error: "缺少 entity / id" };
  upsertData_(req.entity, req.id, JSON.stringify(req.record));
  return { ok: true };
}
function apiRemove(req) {
  const t = checkToken_(req.token);
  if (!t) return { ok: false, auth: false, error: "未授權" };
  if (req.entity === "account" && t.r !== "admin") return { ok: false, error: "僅系統管理員可變更帳號" };
  markDeletedData_(req.entity, req.id);
  return { ok: true };
}

/** 在編輯器手動執行一次此函式，以授權 Google 雲端硬碟存取權限。 */
function grantPermissions() {
  DriveApp.getRootFolder().getName();
  SpreadsheetApp.getActiveSpreadsheet().getName();
}
