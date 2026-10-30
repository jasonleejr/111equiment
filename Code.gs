/**
 * 設備借用／歸還後台（Google Apps Script）
 *
 * 呼叫方式：
 *   借用：    ?id=1001&action=borrow  → C 欄「可借用」改為「借出中」
 *   歸還：    ?id=1001&action=return  → C 欄「借出中」改為「可借用」
 *   設備清單：?action=list            → 回傳 JSON（給 status.html 用）
 *   借還紀錄：?action=log&limit=50    → 回傳 JSON（給 status.html 用）
 *   未帶 action 時預設為 borrow（借用）
 *
 * 工作表 Equipment_List 欄位：
 *   A 欄：設備編號   B 欄：名稱   C 欄：狀態   E 欄：廠牌
 *
 * 每次借用／歸還成功時，會同時在「借還紀錄」分頁新增一列。
 * 該分頁不存在時會自動建立。
 */

var LIST_SHEET_NAME = "Equipment_List";
var LOG_SHEET_NAME = "借還紀錄";

function doGet(e) {
  var equipmentId = e.parameter.id;
  var action = (e.parameter.action || "borrow").toLowerCase();

  // ---- 查詢類（給狀態總覽網頁用）----
  if (action === "list") return jsonOutput(getEquipmentList());
  if (action === "log")  return jsonOutput(getLog(parseInt(e.parameter.limit, 10) || 50));

  // ---- 借用／歸還 ----
  if (!equipmentId) {
    return ContentService.createTextOutput("請提供設備 ID，例如：?id=1001&action=borrow 或 ?id=1001&action=return");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LIST_SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(equipmentId).trim()) {
      var currentStatus = String(data[i][2]).trim();
      var eqName = data[i][1];
      var eqBrand = data[i][4];
      var statusCell = sheet.getRange(i + 1, 3); // C 欄

      // 借用：可借用 → 借出中
      if (action === "borrow") {
        if (currentStatus === "可借用") {
          statusCell.setValue("借出中");
          appendLog(ss, equipmentId, eqName, eqBrand, "借用", "借出中");
          return ContentService.createTextOutput("借用成功！\n" + eqBrand + " " + eqName + "\n狀態已更新為：借出中");
        } else {
          return ContentService.createTextOutput("借用失敗！該設備目前狀態為：" + currentStatus);
        }
      }

      // 歸還：借出中 → 可借用
      if (action === "return") {
        if (currentStatus === "借出中") {
          statusCell.setValue("可借用");
          appendLog(ss, equipmentId, eqName, eqBrand, "歸還", "可借用");
          return ContentService.createTextOutput("歸還成功！\n" + eqBrand + " " + eqName + "\n狀態已更新為：可借用");
        } else {
          return ContentService.createTextOutput("歸還失敗！該設備目前狀態為：" + currentStatus + "（未在借出狀態）");
        }
      }

      return ContentService.createTextOutput("未知的動作：" + action + "（請使用 borrow 或 return）");
    }
  }
  return ContentService.createTextOutput("找不到設備編號：" + equipmentId);
}

/** 回傳 JSON 格式 */
function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 取得所有設備與狀態 */
function getEquipmentList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LIST_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === "") continue;
    items.push({
      id: String(data[i][0]).trim(),
      name: String(data[i][1] || ""),
      status: String(data[i][2] || "").trim(),
      brand: String(data[i][4] || "")
    });
  }
  return { updatedAt: formatTime(new Date()), items: items };
}

/** 取得最近 N 筆借還紀錄（新到舊） */
function getLog(limit) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!logSheet || logSheet.getLastRow() < 2) return { records: [] };

  var lastRow = logSheet.getLastRow();
  var count = Math.min(limit, lastRow - 1);
  var startRow = lastRow - count + 1;
  var data = logSheet.getRange(startRow, 1, count, 6).getValues();

  var records = [];
  for (var i = data.length - 1; i >= 0; i--) {
    records.push({
      time: data[i][0] instanceof Date ? formatTime(data[i][0]) : String(data[i][0]),
      id: String(data[i][1]),
      name: String(data[i][2]),
      brand: String(data[i][3]),
      action: String(data[i][4]),
      status: String(data[i][5])
    });
  }
  return { records: records };
}

/** 在「借還紀錄」分頁新增一列。分頁不存在時自動建立並寫入標題列。 */
function appendLog(ss, equipmentId, eqName, eqBrand, actionText, newStatus) {
  var logSheet = ss.getSheetByName(LOG_SHEET_NAME);

  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(["時間", "設備編號", "名稱", "廠牌", "動作", "更新後狀態"]);
    logSheet.getRange(1, 1, 1, 6).setFontWeight("bold");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidth(1, 160);
  }

  logSheet.appendRow([new Date(), String(equipmentId), eqName, eqBrand, actionText, newStatus]);
  logSheet.getRange(logSheet.getLastRow(), 1).setNumberFormat("yyyy/MM/dd HH:mm:ss");
}

function formatTime(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
}
