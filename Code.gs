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
      case "upload":  out = uploadFile(req); break;
      case "list":    out = listOwner(req.owner); break;
      case "listAll": out = listAll(); break;
      case "delete":  out = deleteFile(req.id); break;
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
