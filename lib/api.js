(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.BookmarkApi = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {

  var DEFAULT_SETTINGS = {
    webAppUrl: '',
    apiToken: '',
    defaultSheetName: 'Bookmarks',
    defaultFolder: '未分类',
    viewMode: 'treeOutline',
    preventDuplicates: true
  };

  function getSettings() {
    return new Promise(function (resolve) {
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.sync) {
        chrome.storage.sync.get(DEFAULT_SETTINGS, function (items) {
          resolve(items || DEFAULT_SETTINGS);
        });
      } else {
        try {
          var saved = localStorage.getItem('gsheet_bookmark_settings');
          resolve(saved ? Object.assign({}, DEFAULT_SETTINGS, JSON.parse(saved)) : DEFAULT_SETTINGS);
        } catch (e) {
          resolve(DEFAULT_SETTINGS);
        }
      }
    });
  }

  function saveSettings(settings) {
    return new Promise(function (resolve) {
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.sync) {
        chrome.storage.sync.set(settings, function () {
          resolve(true);
        });
      } else {
        try {
          localStorage.setItem('gsheet_bookmark_settings', JSON.stringify(settings));
          resolve(true);
        } catch (e) {
          resolve(false);
        }
      }
    });
  }

  async function testConnection(url, token) {
    if (!url || !url.trim()) throw new Error('请先填写 Web App URL！');
    var response = await fetch(url.trim(), {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'test', token: token || '' })
    });
    if (!response.ok) throw new Error('网络请求失败，HTTP 状态码: ' + response.status);
    var result = await response.json();
    if (result.status === 'error') throw new Error(result.message || '连接测试返回错误');
    return result;
  }

  async function fetchSheetNames() {
    var settings = await getSettings();
    if (!settings.webAppUrl) return [];
    try {
      var response = await fetch(settings.webAppUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'getSheets',
          token: settings.apiToken || ''
        })
      });
      var result = await response.json();
      if (result.status === 'success' && Array.isArray(result.sheets)) {
        return result.sheets;
      }
    } catch (e) {
      console.warn('获取工作表列表失败:', e);
    }
    return [];
  }

  async function fetchCategories() {
    var settings = await getSettings();
    if (!settings.webAppUrl) return [];
    try {
      var response = await fetch(settings.webAppUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'getCategories',
          token: settings.apiToken || '',
          options: { sheetName: settings.defaultSheetName || 'Bookmarks' }
        })
      });
      var result = await response.json();
      if (result.status === 'success' && Array.isArray(result.categories)) return result.categories;
    } catch (e) {}
    return [];
  }

  async function saveBookmark(bookmarkData, customOptions) {
    var settings = await getSettings();
    if (!settings.webAppUrl) {
      throw new Error('请先在插件设置中配置 Google Apps Script Web App URL！');
    }

    var options = Object.assign({
      sheetName: settings.defaultSheetName || 'Bookmarks',
      preventDuplicates: settings.preventDuplicates !== false
    }, customOptions || {});

    var payload = {
      action: 'saveBookmark',
      token: settings.apiToken || '',
      bookmark: bookmarkData,
      options: options
    };

    var response = await fetch(settings.webAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      throw new Error('保存失败，HTTP 状态码: ' + response.status);
    }

    return await response.json();
  }

  async function batchImportBookmarks(bookmarks, options, onProgress) {
    var settings = await getSettings();
    var webAppUrl = (options && options.webAppUrl) || settings.webAppUrl;
    var token = (options && options.apiToken) || settings.apiToken;
    var sheetName = (options && options.sheetName) || settings.defaultSheetName || 'Bookmarks';
    var viewMode = (options && options.viewMode) || settings.viewMode || 'treeOutline';
    var clearSheet = options && options.clearSheet === true;

    if (!webAppUrl) throw new Error('请先配置 Web App URL！');
    var total = bookmarks.length;
    if (total === 0) return { total: 0, imported: 0, sheetUrl: '' };

    var batchSize = viewMode === 'treeOutline' ? 300 : 200;
    var chunks = [];
    for (var i = 0; i < total; i += batchSize) {
      chunks.push(bookmarks.slice(i, i + batchSize));
    }

    var totalImported = 0;
    var finalSheetUrl = '';

    for (var chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
      var currentChunk = chunks[chunkIndex];
      var response = await fetch(webAppUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'batchImport',
          token: token || '',
          bookmarks: currentChunk,
          options: {
            sheetName: sheetName,
            viewMode: viewMode,
            clearSheet: (chunkIndex === 0) ? clearSheet : false,
            preventDuplicates: options.preventDuplicates
          }
        })
      });

      var result = await response.json();
      if (result.status === 'error') throw new Error(result.message || '未知错误');
      totalImported += (result.count || currentChunk.length);
      if (result.sheetUrl) finalSheetUrl = result.sheetUrl;

      if (typeof onProgress === 'function') {
        onProgress({
          processed: Math.min((chunkIndex + 1) * batchSize, total),
          total: total,
          percentage: Math.round(((chunkIndex + 1) / chunks.length) * 100),
          chunkIndex: chunkIndex + 1,
          totalChunks: chunks.length
        });
      }
    }

    return { total: total, imported: totalImported, sheetUrl: finalSheetUrl };
  }

  return {
    getSettings: getSettings,
    saveSettings: saveSettings,
    testConnection: testConnection,
    fetchSheetNames: fetchSheetNames,
    fetchCategories: fetchCategories,
    saveBookmark: saveBookmark,
    batchImportBookmarks: batchImportBookmarks
  };
});