/**
 * Google Apps Script - 浏览器书签同步服务 (Chrome Extension Backend)
 * 
 * 功能特点：
 * 1. 接收单条书签存储请求，精准存入已有目录内，保持结构顺序；目录不存在则自动新建
 * 2. 接收 HTML 书签文件解析后的批量导入请求
 * 3. 完美保持目录分组：
 *    - 无论单条新增还是新建分组，所有子书签统一严格归入所属文件夹的同一分组折叠内
 *    - 采用动态深度检测，自动压平阶梯嵌套，杜绝逐级递增子折叠
 * 4. 智能提取已有工作表列表与已有目录列表供插件实时联动选择
 * 5. 自动修复并保持合法的 📂 图标，防止任何字符乱码
 * 6. 支持 URL 查重去重
 * 7. 自动格式化表头、自动设置超链接公式、美化样式
 */

// 安全配置：如果设置了 API_TOKEN，插件请求时必须携带该 Token（留空则不验证）
var API_TOKEN = ""; 

// 默认配置
var DEFAULT_SHEET_NAME = "Bookmarks";

/**
 * 辅助函数：标准化路径
 * 1. 彻底清除 📂 等 Emoji 符号以及历史残留的孤立代理对乱码
 * 2. 去除每个分段的多余空格，规范化层级斜杠（如 "开发 / 前端" -> "开发/前端"）
 */
function normalizePath(p) {
  if (!p) return "";
  var clean = p.toString()
    .replace(/📂/g, "")
    .replace(/[\uD800-\uDFFF]/g, "")
    .replace(/[📁🗂🗃🔖]/g, "");
  return clean.split("/").map(function(s) {
    return s.trim();
  }).filter(Boolean).join("/");
}

/**
 * 辅助函数：精准确保某一行处于指定的行折叠深度
 * targetDepth: 0 表示未折叠根行（目录标题行），1 表示第 1 层行折叠分组（书签行）
 */
function ensureRowGroupDepth(sheet, rowNum, targetDepth) {
  try {
    var depth = sheet.getRowGroupDepth(rowNum);
    if (depth !== targetDepth) {
      sheet.getRange(rowNum, 1, 1, 1).shiftRowGroupDepth(targetDepth - depth);
    }
  } catch (e) {
    Logger.log("Row group error on row " + rowNum + ": " + e);
  }
}

/**
 * 辅助函数：确保某行为第 1 层行折叠深度
 */
function ensureRowGroupDepth1(sheet, rowNum) {
  ensureRowGroupDepth(sheet, rowNum, 1);
}

/**
 * 辅助函数：从指定工作表中读取所有已存在的分类目录列表
 */
function fetchCategoriesFromSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  var categories = [];
  if (!sheet) return categories;

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return categories;

  var firstHeader = sheet.getRange(1, 1).getValue().toString().trim();
  var isFlat = (firstHeader.indexOf("目录路径") !== -1);
  var colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var catSet = {};

  for (var i = 0; i < colA.length; i++) {
    var raw = (colA[i][0] || "").toString().trim();
    if (!raw) continue;

    // 排除书签行（书签行包含 🔖）
    if (raw.indexOf("🔖") !== -1) continue;

    // 目录行的特征：包含 📂、📁、历史遗留字符或扁平表首列
    if (raw.indexOf("📂") !== -1 || raw.indexOf("📁") !== -1 || raw.indexOf("\ud83d") !== -1 || raw.indexOf("\udcc2") !== -1 || raw.indexOf("\udcc1") !== -1 || isFlat) {
      var normCat = normalizePath(raw);
      if (normCat && !catSet[normCat]) {
        catSet[normCat] = true;
        categories.push(normCat);
      }
    }
  }
  return categories;
}

/**
 * 处理 GET 请求（用于插件测试连接及获取当前已有的目录/工作表列表）
 */
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
      var categories = fetchCategoriesFromSheet(ss, sheetName);
      return jsonResponse({
        status: "success",
        categories: categories
      });
    }

    return jsonResponse({ status: "success", message: "服务运行正常" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * 处理 POST 请求（保存单条书签、批量导入书签、获取工作表/目录）
 */
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
      var categories = fetchCategoriesFromSheet(ss, sheetName);
      return jsonResponse({ status: "success", categories: categories });
    }

    return jsonResponse({ status: "error", message: "未知的 action: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * 保存单个书签
 * 智能定位目录位置，精准存入已有目录内；统一同级分组，绝不逐级递增子折叠
 */
function handleSaveSingleBookmark(ss, payload) {
  var bookmark = payload.bookmark || {};
  var options = payload.options || {};
  var sheetName = options.sheetName || DEFAULT_SHEET_NAME;

  var sheet = getOrCreateSheet(ss, sheetName);
  initHeaderIfEmpty(sheet);

  var url = (bookmark.url || "").trim();
  var title = (bookmark.title || "未命名书签").trim();
  var folderPath = (bookmark.folderPath || "书签栏").trim();
  var notes = (bookmark.notes || "").trim();
  var dateAdded = bookmark.dateAdded || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  if (!url) {
    return jsonResponse({ status: "error", message: "书签 URL 不能为空" });
  }

  // 判断工作表类型是扁平表还是大纲折叠树
  var isFlat = false;
  if (sheet.getLastRow() >= 1) {
    var firstHeader = sheet.getRange(1, 1).getValue().toString().trim();
    if (firstHeader.indexOf("目录路径") !== -1) {
      isFlat = true;
    }
  }

  // 查重处理（扁平表第 4 列是 URL，大纲树第 2 列是 URL）
  if (options.preventDuplicates !== false) {
    var checkCol = isFlat ? 4 : 2;
    var existingRow = findRowByUrl(sheet, url, checkCol);
    if (existingRow > 0) {
      return jsonResponse({
        status: "duplicate",
        message: "该网页已存在于工作表「" + sheet.getName() + "」第 " + existingRow + " 行",
        row: existingRow,
        sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
      });
    }
  }

  var targetNorm = normalizePath(folderPath);
  if (!targetNorm) targetNorm = "书签栏";

  var lastRow = sheet.getLastRow();

  if (isFlat) {
    // 扁平表结构处理：[目录路径, 一级分类, 书签名称, URL, 备注, 标签, 添加时间]
    var parts = targetNorm.split("/");
    var primary = parts[0] || "根目录";
    var formula = '=HYPERLINK("' + url.replace(/"/g, '""') + '", "' + title.replace(/"/g, '""') + '")';
    var newRowData = [targetNorm, primary, formula, url, notes, "", dateAdded];

    var targetInsertRow = -1;
    if (lastRow > 1) {
      var col1Vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var r = col1Vals.length - 1; r >= 0; r--) {
        if (normalizePath((col1Vals[r][0] || "").toString()) === targetNorm) {
          targetInsertRow = r + 2;
          break;
        }
      }
    }

    var insertedRow = -1;
    if (targetInsertRow !== -1) {
      sheet.insertRowAfter(targetInsertRow);
      insertedRow = targetInsertRow + 1;
      var range = sheet.getRange(insertedRow, 1, 1, 7);
      range.setValues([newRowData]);
      range.setVerticalAlignment("middle");
    } else {
      sheet.appendRow(newRowData);
      insertedRow = sheet.getLastRow();
      sheet.getRange(insertedRow, 1, 1, 7).setVerticalAlignment("middle");
    }

    var titleCell = sheet.getRange(insertedRow, 3);
    titleCell.setFontColor("#1a73e8");
    titleCell.setFontLine("underline");

    return jsonResponse({
      status: "success",
      message: "书签已成功存入工作表「" + sheet.getName() + "」的「" + targetNorm + "」分类下！",
      sheetName: sheet.getName(),
      row: insertedRow,
      sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
    });
  }

  // 大纲折叠树视图处理 (Tree Outline)
  var targetFolderRow = -1;
  var nextFolderRow = -1;

  if (lastRow >= 2) {
    var colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < colA.length; i++) {
      var currentRow = i + 2;
      var cellText = (colA[i][0] || "").toString().trim();
      
      // 必须是目录行，且排除书签行（书签行包含 🔖）
      var isFolderLine = (cellText.indexOf("📂") !== -1 || cellText.indexOf("📁") !== -1 || cellText.indexOf("\ud83d") !== -1 || cellText.indexOf("\udcc2") !== -1 || cellText.indexOf("\udcc1") !== -1) && cellText.indexOf("🔖") === -1;

      if (isFolderLine) {
        var normFolder = normalizePath(cellText);

        if (targetFolderRow === -1) {
          if (normFolder === targetNorm) {
            targetFolderRow = currentRow;
          }
        } else {
          // 已匹配到目标目录，后续遇到的第一个目录行即为边界
          nextFolderRow = currentRow;
          break;
        }
      }
    }
  }

  var parts = targetNorm.split("/");
  var level = Math.max(1, parts.length);
  var bmRowNum = -1;

  if (targetFolderRow !== -1) {
    // 1. 已有该目录行：定位该目录下的最后一行插入位置（下一个目录行之前）
    var insertAfterRow = (nextFolderRow !== -1) ? (nextFolderRow - 1) : sheet.getLastRow();

    // 继承该目录的缩进格式
    var folderRowText = (sheet.getRange(targetFolderRow, 1).getValue() || "").toString();
    var matchSpaces = folderRowText.match(/^(\s*)/);
    var baseIndent = matchSpaces ? matchSpaces[1] : "";

    // 自动修复并保持该文件夹标题行的完整 📂 图标
    sheet.getRange(targetFolderRow, 1).setValue(baseIndent + "📂 " + targetNorm);

    // 确保该目录标题行自身处于根层级 (depth = 0)
    ensureRowGroupDepth(sheet, targetFolderRow, 0);

    var itemTitle = baseIndent + "    🔖 " + title;
    var formula = '=HYPERLINK("' + url.replace(/"/g, '""') + '", "' + itemTitle.replace(/"/g, '""') + '")';

    sheet.insertRowAfter(insertAfterRow);
    bmRowNum = insertAfterRow + 1;

    var newRange = sheet.getRange(bmRowNum, 1, 1, 4);
    newRange.setValues([[formula, url, dateAdded, notes]]);
    newRange.setVerticalAlignment("middle");

    var bmCell = sheet.getRange(bmRowNum, 1);
    bmCell.setFontColor("#1a73e8");
    sheet.setRowHeight(bmRowNum, 26);

    // 关键分组控制：确保该目录下所有书签统一处于第 1 层折叠组，绝不产生嵌套阶梯
    var actualFolderEnd = (nextFolderRow !== -1) ? nextFolderRow : sheet.getLastRow();
    for (var r = targetFolderRow + 1; r <= actualFolderEnd; r++) {
      ensureRowGroupDepth(sheet, r, 1);
    }

  } else {
    // 2. 目录不存在：在工作表末尾新建该目录分组及书签
    var folderIndent = "";
    for (var l = 0; l < (level - 1); l++) {
      folderIndent += "    ";
    }
    var folderDisplayName = folderIndent + "📂 " + targetNorm;

    sheet.appendRow([folderDisplayName, "", "", ""]);
    var fRowNum = sheet.getLastRow();
    var fRange = sheet.getRange(fRowNum, 1, 1, 4);
    fRange.setBackground("#F1F3F4");
    fRange.setFontWeight("bold");
    fRange.setFontColor("#202124");
    sheet.setRowHeight(fRowNum, 30);
    // 确保新建的目录标题行为未折叠根行 (depth = 0)
    ensureRowGroupDepth(sheet, fRowNum, 0);

    var itemIndent = "";
    for (var l2 = 0; l2 < level; l2++) {
      itemIndent += "    ";
    }
    var itemTitle = itemIndent + "🔖 " + title;
    var formula = '=HYPERLINK("' + url.replace(/"/g, '""') + '", "' + itemTitle.replace(/"/g, '""') + '")';

    sheet.appendRow([formula, url, dateAdded, notes]);
    bmRowNum = sheet.getLastRow();

    var bmCell = sheet.getRange(bmRowNum, 1);
    bmCell.setFontColor("#1a73e8");
    sheet.setRowHeight(bmRowNum, 26);

    // 新建分类的书签加入第 1 层行折叠 (depth = 1)
    ensureRowGroupDepth(sheet, bmRowNum, 1);
  }

  return jsonResponse({
    status: "success",
    message: "书签已成功存入工作表「" + sheet.getName() + "」的「" + targetNorm + "」分类下！",
    sheetName: sheet.getName(),
    row: bmRowNum,
    sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
  });
}

/**
 * 批量导入书签（支持原生大纲折叠树或扁平数据表）
 */
function handleBatchImport(ss, payload) {
  var bookmarks = payload.bookmarks || [];
  var options = payload.options || {};
  var sheetName = options.sheetName || DEFAULT_SHEET_NAME;
  var viewMode = options.viewMode || "treeOutline"; // 默认大纲折叠
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

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  if (viewMode === "treeOutline") {
    return importAsTreeOutline(ss, sheet, bookmarks, options);
  } else {
    return importAsFlatTable(ss, sheet, bookmarks, options);
  }
}

/**
 * 模式一：树状大纲折叠分组视图 (Native Row Groups)
 */
function importAsTreeOutline(ss, sheet, bookmarks, options) {
  var startRow = sheet.getLastRow();
  if (startRow <= 1) {
    sheet.clear();
    var headers = [["📁 目录层级 / 书签名称", "🔗 链接 URL", "📅 添加时间", "📝 备注"]];
    var hRange = sheet.getRange(1, 1, 1, 4);
    hRange.setValues(headers);
    hRange.setBackground("#0F9D58");
    hRange.setFontColor("#FFFFFF");
    hRange.setFontWeight("bold");
    hRange.setHorizontalAlignment("left");
    hRange.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 35);
    sheet.setFrozenRows(1);
    startRow = 1;
  }

  var rows = [];
  var rowMetadata = [];
  var lastFolder = null;

  for (var i = 0; i < bookmarks.length; i++) {
    var b = bookmarks[i];
    var folderPath = normalizePath(b.folderPath || "书签栏");
    var title = b.title || "未命名";
    var url = b.url || "";
    var dateAdded = b.dateAdded || "";
    var level = b.level || 1;

    // 当遇到新文件夹时，插入文件夹分组行
    if (folderPath !== lastFolder) {
      lastFolder = folderPath;
      var indentSpaces = "";
      for (var l = 0; l < (level - 1); l++) {
        indentSpaces += "    ";
      }
      var folderDisplayName = indentSpaces + "📂 " + folderPath;
      rows.push([folderDisplayName, "", "", ""]);
      rowMetadata.push({ type: "folder", level: level, path: folderPath });
    }

    var indent = "";
    for (var l2 = 0; l2 < level; l2++) {
      indent += "    ";
    }
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

  // 美化和样式：文件夹行加粗和背景色，书签行链接高亮
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

  // 建立 Google Sheets 原生行折叠分组 (Row Groups)
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
 * 模式二：标准扁平结构化表格 (Flat Table)
 */
function importAsFlatTable(ss, sheet, bookmarks, options) {
  initHeaderIfEmpty(sheet, "flat");

  var rows = [];
  var existingUrls = {};
  if (options.preventDuplicates) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var urls = sheet.getRange(2, 4, lastRow - 1, 1).getValues();
      for (var u = 0; u < urls.length; u++) {
        var existing = (urls[u][0] || "").toString().trim();
        if (existing) existingUrls[existing] = true;
      }
    }
  }

  var skippedCount = 0;
  for (var i = 0; i < bookmarks.length; i++) {
    var b = bookmarks[i];
    var url = (b.url || "").trim();
    if (!url) continue;

    if (options.preventDuplicates && existingUrls[url]) {
      skippedCount++;
      continue;
    }
    existingUrls[url] = true;

    var title = (b.title || "未命名书签").trim();
    var folderPath = normalizePath(b.folderPath || "根目录");
    var parts = folderPath.split("/").map(function(p){ return p.trim(); }).filter(Boolean);
    var primary = parts[0] || "根目录";
    var formula = '=HYPERLINK("' + url.replace(/"/g, '""') + '", "' + title.replace(/"/g, '""') + '")';
    var dateAdded = b.dateAdded || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    rows.push([
      folderPath,
      primary,
      formula,
      url,
      b.notes || "",
      b.tags || "",
      dateAdded
    ]);
  }

  if (rows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, 7).setValues(rows);
    sheet.getRange(startRow, 1, rows.length, 7).setVerticalAlignment("middle");
    var titleRange = sheet.getRange(startRow, 3, rows.length, 1);
    titleRange.setFontColor("#1a73e8");
    titleRange.setFontLine("underline");
  }

  return jsonResponse({
    status: "success",
    count: rows.length,
    skipped: skippedCount,
    message: "成功导入 " + rows.length + " 条书签" + (skippedCount > 0 ? "（已跳过 " + skippedCount + " 条重复）" : ""),
    sheetName: sheet.getName(),
    sheetUrl: ss.getUrl() + "#gid=" + sheet.getSheetId()
  });
}

/**
 * 初始化表头
 */
function initHeaderIfEmpty(sheet, mode) {
  if (sheet.getLastRow() === 0) {
    if (mode === "flat") {
      var headers = [
        ["目录路径", "一级分类", "书签名称", "URL", "备注", "标签", "添加时间"]
      ];
      var hRange = sheet.getRange(1, 1, 1, 7);
      hRange.setValues(headers);
      hRange.setBackground("#0F9D58");
      hRange.setFontColor("#FFFFFF");
      hRange.setFontWeight("bold");
      hRange.setHorizontalAlignment("left");
      hRange.setVerticalAlignment("middle");
      sheet.setRowHeight(1, 35);
      sheet.setFrozenRows(1);
    } else {
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
}

/**
 * 根据 URL 查重
 */
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

/**
 * 辅助获取或新建 Sheet
 */
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * 输出标准 JSON 响应
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}