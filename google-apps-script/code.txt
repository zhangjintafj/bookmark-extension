/**
 * Google Apps Script - 浏览器书签同步服务 (4列大纲统一版)
 */

var API_TOKEN = ""; 
var DEFAULT_SHEET_NAME = "Bookmarks";

function doGet(e) {
  try {
    var params = (e && e.parameter) || {};
    var token = params.token || "";
    if (API_TOKEN && API_TOKEN.trim() !== "" && token !== API_TOKEN) {
      return jsonResponse({ status: "error", message: "Token 验证失败，请检查配置" });
    }

    var action = params.action || "test";
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "test") {
      return jsonResponse({
        status: "success",
        message: "Google Sheets 连接成功！",
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    if (action === "getSheets") {
      var sheets = ss.getSheets();
      var sheetNames = [];
      for (var s = 0; s < sheets.length; s++) {
        sheetNames.push(sheets[s].getName());
      }
      return jsonResponse({ status: "success", sheets: sheetNames });
    }

    if (action === "getCategories") {
      var sheetName = params.sheetName || DEFAULT_SHEET_NAME;
      var sheet = ss.getSheetByName(sheetName);
      var categories = [];
      if (sheet && sheet.getLastRow() > 1) {
        var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        var catSet = {};
        for (var i = 0; i < data.length; i++) {
          var cat = (data[i][0] || "").toString().trim();
          if (cat && !catSet[cat]) { catSet[cat] = true; categories.push(cat); }
        }
      }
      return jsonResponse({ status: "success", categories: categories });
    }

    return jsonResponse({ status: "success", message: "服务运行正常" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "请求体为空" });
    }

    var payload = JSON.parse(e.postData.contents);
    var token = payload.token || "";

    if (API_TOKEN && API_TOKEN.trim() !== "" && token !== API_TOKEN) {
      return jsonResponse({ status: "error", message: "Token 验证失败，请检查配置" });
    }

    var action = payload.action || "saveBookmark";
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "test") {
      return jsonResponse({
        status: "success",
        message: "Google Sheets 连接成功！",
        spreadsheetName: ss.getName(),
        spreadsheetUrl: ss.getUrl()
      });
    }

    if (action === "saveBookmark") {
      return handleSaveSingleBookmark(ss, payload);
    }

    if (action === "batchImport") {
      return handleBatchImport(ss, payload);
    }

    if (action === "getSheets") {
      var sheets = ss.getSheets();
      var sheetNames = [];
      for (var s = 0; s < sheets.length; s++) {
        sheetNames.push(sheets[s].getName());
      }
      return jsonResponse({ status: "success", sheets: sheetNames });
    }

    if (action === "getCategories") {
      var sheetName = (payload.options && payload.options.sheetName) || DEFAULT_SHEET_NAME;
      var sheet = ss.getSheetByName(sheetName);
      var categories = [];
      if (sheet && sheet.getLastRow() > 1) {
        var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        var catSet = {};
        for (var i = 0; i < data.length; i++) {
          var cat = (data[i][0] || "").toString().trim();
          if (cat && !catSet[cat]) { catSet[cat] = true; categories.push(cat); }
        }
      }
      return jsonResponse({ status: "success", categories: categories });
    }

    return jsonResponse({ status: "error", message: "未知的 action: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * 保存单个书签（与导入完全一致的 4 列表头与大纲格式）
 */
function handleSaveSingleBookmark(ss, payload) {
  var bookmark = payload.bookmark || {};
  var options = payload.options || {};
  var sheetName = options.sheetName || DEFAULT_SHEET_NAME;

  var sheet = getOrCreateSheet(ss, sheetName);
  initHeaderIfEmpty(sheet); // 初始化统一的 4 列表头

  var url = (bookmark.url || "").trim();
  var title = (bookmark.title || "未命名书签").trim();
  var folderPath = (bookmark.folderPath || "书签栏").trim();
  var notes = (bookmark.notes || "").trim();
  var dateAdded = bookmark.dateAdded || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  if (!url) {
    return jsonResponse({ status: "error", message: "书签 URL 不能为空" });
  }

  // 查重：第 2 列是 URL
  if (options.preventDuplicates !== false) {
    var existingRow = findRowByUrl(sheet, url, 2);
    if (existingRow > 0) {
      return jsonResponse({
        status: "duplicate",
        message: "该网页已存在于工作表「" + sheet.getName() + "」第 " + existingRow + " 行",
        row: existingRow,
        sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
      });
    }
  }

  var lastRow = sheet.getLastRow();
  var folderExists = false;

  // 检查是否已有该文件夹目录行
  if (lastRow >= 1) {
    var colA = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var text = (colA[i][0] || "").toString().trim();
      if (text.indexOf("📂") !== -1 && text.indexOf(folderPath) !== -1) {
        folderExists = true;
        break;
      }
    }
  }

  var itemTitle = "    🔖 " + title;
  var formula = '=HYPERLINK("' + url.replace(/"/g, '""') + '", "' + itemTitle.replace(/"/g, '""') + '")';

  // 若尚未存在该文件夹分组，先创建灰底粗体的文件夹标题行
  if (!folderExists) {
    sheet.appendRow(["📂 " + folderPath, "", "", ""]);
    var fRowNum = sheet.getLastRow();
    var fRange = sheet.getRange(fRowNum, 1, 1, 4);
    fRange.setBackground("#F1F3F4");
    fRange.setFontWeight("bold");
    fRange.setFontColor("#202124");
    sheet.setRowHeight(fRowNum, 30);
  }

  // 追加书签行: [ 目录层级/书签名称(超链接), 链接 URL, 添加时间, 备注 ]
  sheet.appendRow([formula, url, dateAdded, notes]);
  var bmRowNum = sheet.getLastRow();
  var bmRange = sheet.getRange(bmRowNum, 1);
  bmRange.setFontColor("#1a73e8");
  sheet.setRowHeight(bmRowNum, 26);

  // 自动将书签行加入原生折叠行组
  try {
    sheet.getRange(bmRowNum, 1, 1, 1).shiftRowGroupDepth(1);
  } catch (e) {}

  return jsonResponse({
    status: "success",
    message: "书签已成功按大纲目录存入工作表「" + sheet.getName() + "」！",
    sheetName: sheet.getName(),
    row: bmRowNum,
    sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
  });
}

function handleBatchImport(ss, payload) {
  var bookmarks = payload.bookmarks || [];
  var options = payload.options || {};
  var sheetName = options.sheetName || DEFAULT_SHEET_NAME;
  var viewMode = options.viewMode || "treeOutline";
  var clearSheet = options.clearSheet === true;

  if (!bookmarks || bookmarks.length === 0) {
    return jsonResponse({ status: "error", message: "导入的书签列表为空" });
  }

  var sheet = ss.getSheetByName(sheetName);
  if (clearSheet && sheet) {
    try {
      var maxRows = sheet.getMaxRows();
      if (maxRows > 1) {
        sheet.getRange(1, 1, maxRows, sheet.getMaxColumns()).shiftRowGroupDepth(-8);
      }
    } catch(e) {}
    sheet.clear();
  }

  if (!sheet) sheet = ss.insertSheet(sheetName);

  return importAsTreeOutline(ss, sheet, bookmarks, options);
}

function importAsTreeOutline(ss, sheet, bookmarks, options) {
  var startRow = sheet.getLastRow();
  if (startRow <= 1) {
    initHeaderIfEmpty(sheet);
    startRow = 1;
  }

  var rows = [];
  var rowMetadata = [];
  var lastFolder = null;

  for (var i = 0; i < bookmarks.length; i++) {
    var b = bookmarks[i];
    var folderPath = b.folderPath || "书签栏";
    var title = b.title || "未命名";
    var url = b.url || "";
    var dateAdded = b.dateAdded || "";
    var level = b.level || 1;

    if (folderPath !== lastFolder) {
      lastFolder = folderPath;
      var indentSpaces = "";
      for (var l = 0; l < (level - 1); l++) indentSpaces += "    ";
      rows.push([indentSpaces + "📂 " + folderPath, "", "", ""]);
      rowMetadata.push({ type: "folder", level: level, path: folderPath });
    }

    var indent = "";
    for (var l2 = 0; l2 < level; l2++) indent += "    ";
    var itemTitle = indent + "🔖 " + title;
    var hyperlinkFormula = url ? '=HYPERLINK("' + url.replace(/"/g, '""') + '", "' + itemTitle.replace(/"/g, '""') + '")' : itemTitle;
    
    rows.push([hyperlinkFormula, url, dateAdded, b.notes || ""]);
    rowMetadata.push({ type: "bookmark", level: level, url: url });
  }

  if (rows.length === 0) {
    return jsonResponse({ status: "success", count: 0, message: "无数据可写入" });
  }

  var writeStartRow = sheet.getLastRow() + 1;
  var range = sheet.getRange(writeStartRow, 1, rows.length, 4);
  range.setValues(rows);
  range.setVerticalAlignment("middle");

  for (var r = 0; r < rowMetadata.length; r++) {
    var currentRowIndex = writeStartRow + r;
    var meta = rowMetadata[r];
    if (meta.type === "folder") {
      var folderRange = sheet.getRange(currentRowIndex, 1, 1, 4);
      folderRange.setBackground("#F1F3F4");
      folderRange.setFontWeight("bold");
      folderRange.setFontColor("#202124");
      sheet.setRowHeight(currentRowIndex, 30);
    } else {
      var bmRange = sheet.getRange(currentRowIndex, 1);
      bmRange.setFontColor("#1a73e8");
      sheet.setRowHeight(currentRowIndex, 26);
    }
  }

  try {
    var groupStart = -1;
    for (var m = 0; m < rowMetadata.length; m++) {
      var rowNum = writeStartRow + m;
      if (rowMetadata[m].type === "folder") {
        if (groupStart !== -1 && rowNum - 1 >= groupStart) {
          sheet.getRange(groupStart, 1, (rowNum - groupStart), 1).shiftRowGroupDepth(1);
        }
        groupStart = rowNum + 1;
      }
    }
    if (groupStart !== -1 && (writeStartRow + rowMetadata.length - 1) >= groupStart) {
      sheet.getRange(groupStart, 1, (writeStartRow + rowMetadata.length - groupStart), 1).shiftRowGroupDepth(1);
    }
  } catch (groupErr) {
    Logger.log("Row group error: " + groupErr);
  }

  sheet.setColumnWidth(1, 420);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 160);
  sheet.setColumnWidth(4, 200);

  return jsonResponse({
    status: "success",
    count: bookmarks.length,
    message: "成功以【大纲折叠树】模式导入 " + bookmarks.length + " 条书签！",
    sheetName: sheet.getName(),
    sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
  });
}

/**
 * 统一初始化 4 列表头（与截图 100% 一致）
 */
function initHeaderIfEmpty(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      ["📁 目录层级 / 书签名称", "🔗 链接 URL", "📅 添加时间", "📝 备注"]
    ];
    var hRange = sheet.getRange(1, 1, 1, 4);
    hRange.setValues(headers);
    hRange.setBackground("#0F9D58");
    hRange.setFontColor("#FFFFFF");
    hRange.setFontWeight("bold");
    hRange.setHorizontalAlignment("left");
    hRange.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 35);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 420);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 160);
    sheet.setColumnWidth(4, 200);
  }
}

function findRowByUrl(sheet, targetUrl, colIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  var values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().trim() === targetUrl) {
      return i + 2;
    }
  }
  return -1;
}

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}