/**
 * 弹出窗口 (Popup) 交互逻辑
 */

document.addEventListener('DOMContentLoaded', async function () {
  var pageTitleInput = document.getElementById('pageTitle');
  var pageUrlInput = document.getElementById('pageUrl');
  var targetSheetSelect = document.getElementById('targetSheetSelect');
  var targetSheetInput = document.getElementById('targetSheetInput');
  var sheetSelectWrapper = document.getElementById('sheetSelectWrapper');
  var sheetInputWrapper = document.getElementById('sheetInputWrapper');
  var btnCancelNewSheet = document.getElementById('btnCancelNewSheet');
  var folderSelect = document.getElementById('folderSelect');
  var folderInput = document.getElementById('folderInput');
  var folderSelectWrapper = document.getElementById('folderSelectWrapper');
  var folderInputWrapper = document.getElementById('folderInputWrapper');
  var btnCancelNewFolder = document.getElementById('btnCancelNewFolder');
  var btnRefreshSheets = document.getElementById('btnRefreshSheets');
  var sheetLoadingTip = document.getElementById('sheetLoadingTip');
  var categoryLoadingTip = document.getElementById('categoryLoadingTip');
  var notesInput = document.getElementById('bookmarkNotes');
  var btnSave = document.getElementById('btnSaveBookmark');
  var btnText = btnSave.querySelector('.btn-text');
  var btnLoading = btnSave.querySelector('.btn-loading');
  var feedbackBox = document.getElementById('feedbackBox');
  var feedbackText = document.getElementById('feedbackText');
  var feedbackSheetLink = document.getElementById('feedbackSheetLink');
  var unconfiguredAlert = document.getElementById('unconfiguredAlert');
  var btnGoConfig = document.getElementById('btnGoConfig');
  var btnOpenOptions = document.getElementById('btnOpenOptions');
  var btnOpenImport = document.getElementById('btnOpenImport');
  var btnImportPageLink = document.getElementById('btnImportPageLink');
  var folderSuggestions = document.getElementById('folderSuggestions');

  // --- 字符清洗辅助函数：彻底去除 Emoji 及截断造成的孤立代理对乱码 ---
  function cleanCategoryString(cat) {
    if (!cat) return '';
    return cat.toString()
      .replace(/📂/g, '')
      .replace(/[\uD800-\uDFFF]/g, '')
      .replace(/[📁🗂🗃🔖]/g, '')
      .replace(/^[ \t\/\\]+/, '')
      .replace(/[ \t\/\\]+$/, '')
      .trim();
  }

  // --- 存储与缓存辅助函数 (SWR 策略，确保快速显示) ---
  function getLocalCache(key) {
    return new Promise(function (resolve) {
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
        chrome.storage.local.get([key], function (res) {
          resolve(res ? res[key] : null);
        });
      } else {
        try {
          var val = localStorage.getItem(key);
          resolve(val ? JSON.parse(val) : null);
        } catch (e) {
          resolve(null);
        }
      }
    });
  }

  function setLocalCache(key, val) {
    return new Promise(function (resolve) {
      if (typeof chrome !== 'undefined' && chrome.storage && chrome.storage.local) {
        var obj = {};
        obj[key] = val;
        chrome.storage.local.set(obj, resolve);
      } else {
        try {
          localStorage.setItem(key, JSON.stringify(val));
        } catch (e) {}
        resolve();
      }
    });
  }

  // 1. 读取基础设置与上次选中的工作表与目录
  var settings = await BookmarkApi.getSettings();
  if (!settings.webAppUrl) {
    unconfiguredAlert.style.display = 'flex';
  }

  var savedSheet = (await getLocalCache('last_target_sheet')) || settings.defaultSheetName || 'Bookmarks';
  var savedFolder = cleanCategoryString(await getLocalCache('last_folder_path') || settings.defaultFolder || '书签栏');

  // 渲染工作表下拉框（【➕ 新建工作表...】置顶在第一项，查找极方便）
  function renderSheetOptions(sheets, selectedName) {
    if (!sheets || sheets.length === 0) return;
    targetSheetSelect.innerHTML = '';

    // 1. 新建选项置顶第一项
    var newOpt = document.createElement('option');
    newOpt.value = '__NEW__';
    newOpt.textContent = '➕ 新建工作表...';
    targetSheetSelect.appendChild(newOpt);

    var hasMatch = false;
    sheets.forEach(function (name) {
      var opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      if (name === selectedName) {
        opt.selected = true;
        hasMatch = true;
      }
      targetSheetSelect.appendChild(opt);
    });

    // 默认选中上次或第一个已有工作表，避免自动跳到新建
    if (!hasMatch && sheets[0]) {
      targetSheetSelect.value = sheets[0];
    }
  }

  // 渲染分组目录下拉框（【➕ 新建目录...】置顶在第一项，查找极方便）
  function renderFolderOptions(categories, selectedName) {
    folderSelect.innerHTML = '';
    folderSuggestions.innerHTML = '';

    // 1. 新建选项置顶第一项
    var newOpt = document.createElement('option');
    newOpt.value = '__NEW__';
    newOpt.textContent = '➕ 新建目录...';
    folderSelect.appendChild(newOpt);

    var cleanCats = [];
    var catSet = new Set();
    if (categories && categories.length > 0) {
      categories.forEach(function (c) {
        var clean = cleanCategoryString(c);
        if (clean && !catSet.has(clean)) {
          catSet.add(clean);
          cleanCats.push(clean);
        }
      });
    }

    if (cleanCats.length === 0) {
      cleanCats = ['书签栏'];
    }

    var hasMatch = false;
    cleanCats.forEach(function (name) {
      var opt = document.createElement('option');
      opt.value = name;
      opt.textContent = name;
      if (name === selectedName) {
        opt.selected = true;
        hasMatch = true;
      }
      folderSelect.appendChild(opt);

      // 也放入 datalist 供手动输入时联想提示
      var dlOpt = document.createElement('option');
      dlOpt.value = name;
      folderSuggestions.appendChild(dlOpt);
    });

    // 默认选中匹配的已有目录，避免自动跳到新建
    if (!hasMatch && cleanCats[0]) {
      folderSelect.value = cleanCats[0];
    }
  }

  // 2. 优先从本地缓存渲染工作表与分类（0毫秒瞬间展示）
  var cachedSheets = await getLocalCache('cached_sheet_names');
  if (cachedSheets && cachedSheets.length > 0) {
    renderSheetOptions(cachedSheets, savedSheet);
  }

  var currentTargetSheet = targetSheetSelect.value || savedSheet;
  var cachedCategories = await getLocalCache('cached_categories_' + currentTargetSheet);
  if (cachedCategories && cachedCategories.length > 0) {
    renderFolderOptions(cachedCategories, savedFolder);
  }

  // 3. 异步并发刷新工作表与目录（SWR 保持最新）
  async function refreshSheetsAndCategories(isManual) {
    if (!settings.webAppUrl) return;

    if (isManual) {
      btnRefreshSheets.classList.add('spinning');
    }
    sheetLoadingTip.style.display = 'inline';
    sheetLoadingTip.textContent = '读取中...';
    categoryLoadingTip.style.display = 'inline';
    categoryLoadingTip.textContent = '读取中...';

    var curTargetSheet = (sheetInputWrapper.style.display !== 'none') ? targetSheetInput.value.trim() : targetSheetSelect.value;
    if (!curTargetSheet || curTargetSheet === '__NEW__') {
      curTargetSheet = savedSheet || 'Bookmarks';
    }

    try {
      var results = await Promise.allSettled([
        BookmarkApi.fetchSheetNames(),
        BookmarkApi.fetchCategories(curTargetSheet)
      ]);

      var sheetsRes = results[0];
      var catsRes = results[1];

      // 处理工作表结果
      if (sheetsRes.status === 'fulfilled' && Array.isArray(sheetsRes.value) && sheetsRes.value.length > 0) {
        var sheets = sheetsRes.value;
        await setLocalCache('cached_sheet_names', sheets);
        var curSelected = (sheetInputWrapper.style.display !== 'none') ? targetSheetInput.value.trim() : targetSheetSelect.value;
        renderSheetOptions(sheets, curSelected || savedSheet);
      }

      // 处理目录分类结果
      if (catsRes.status === 'fulfilled' && Array.isArray(catsRes.value)) {
        var cleanCats = [];
        var catSet = new Set();
        catsRes.value.forEach(function (c) {
          var clean = cleanCategoryString(c);
          if (clean && !catSet.has(clean)) {
            catSet.add(clean);
            cleanCats.push(clean);
          }
        });
        await setLocalCache('cached_categories_' + curTargetSheet, cleanCats);
        var curFolderSelected = (folderInputWrapper.style.display !== 'none') ? folderInput.value.trim() : folderSelect.value;
        renderFolderOptions(cleanCats, curFolderSelected || savedFolder);
      }
    } catch (e) {
      console.warn('刷新列表失败:', e);
    } finally {
      sheetLoadingTip.style.display = 'none';
      categoryLoadingTip.style.display = 'none';

      if (isManual) {
        btnRefreshSheets.classList.remove('spinning');
        sheetLoadingTip.style.display = 'inline';
        sheetLoadingTip.textContent = '✓ 已更新';
        setTimeout(function () {
          sheetLoadingTip.style.display = 'none';
        }, 1500);
      }
    }
  }

  // 后台自动执行一次刷新
  refreshSheetsAndCategories(false);

  // 点击刷新按钮手动强制更新
  btnRefreshSheets.addEventListener('click', function () {
    refreshSheetsAndCategories(true);
  });

  // 工作表下拉与新建输入切换
  targetSheetSelect.addEventListener('change', async function () {
    if (targetSheetSelect.value === '__NEW__') {
      sheetSelectWrapper.style.display = 'none';
      sheetInputWrapper.style.display = 'block';
      targetSheetInput.value = '';
      targetSheetInput.focus();
    } else {
      var chosenSheet = targetSheetSelect.value;
      await setLocalCache('last_target_sheet', chosenSheet);

      // 联动加载该工作表的分类：先读缓存，再后台刷新
      var cachedCats = await getLocalCache('cached_categories_' + chosenSheet);
      if (cachedCats) {
        renderFolderOptions(cachedCats, folderSelect.value || savedFolder);
      }
      categoryLoadingTip.style.display = 'inline';
      categoryLoadingTip.textContent = '读取中...';
      BookmarkApi.fetchCategories(chosenSheet).then(async function (cats) {
        categoryLoadingTip.style.display = 'none';
        if (cats) {
          var cleanCats = [];
          var catSet = new Set();
          cats.forEach(function (c) {
            var clean = cleanCategoryString(c);
            if (clean && !catSet.has(clean)) {
              catSet.add(clean);
              cleanCats.push(clean);
            }
          });
          await setLocalCache('cached_categories_' + chosenSheet, cleanCats);
          renderFolderOptions(cleanCats, folderSelect.value || savedFolder);
        }
      }).catch(function () {
        categoryLoadingTip.style.display = 'none';
      });
    }
  });

  btnCancelNewSheet.addEventListener('click', function () {
    sheetInputWrapper.style.display = 'none';
    sheetSelectWrapper.style.display = 'block';
    // 取消时切回已有列表中的第一个工作表（索引 1）
    targetSheetSelect.value = (targetSheetSelect.options[1] ? targetSheetSelect.options[1].value : 'Bookmarks');
  });

  // 分组目录下拉与新建输入切换
  folderSelect.addEventListener('change', function () {
    if (folderSelect.value === '__NEW__') {
      folderSelectWrapper.style.display = 'none';
      folderInputWrapper.style.display = 'block';
      folderInput.value = '';
      folderInput.focus();
    }
  });

  btnCancelNewFolder.addEventListener('click', function () {
    folderInputWrapper.style.display = 'none';
    folderSelectWrapper.style.display = 'block';
    // 取消时切回已有列表中的第一个目录（索引 1）
    folderSelect.value = (folderSelect.options[1] ? folderSelect.options[1].value : '书签栏');
  });

  // 4. 获取当前浏览器活动标签页信息
  if (typeof chrome !== 'undefined' && chrome.tabs && chrome.tabs.query) {
    chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
      if (tabs && tabs[0]) {
        var currentTab = tabs[0];
        pageTitleInput.value = currentTab.title || '';
        pageUrlInput.value = currentTab.url || '';
      }
    });
  }

  // 5. 事件绑定：保存书签
  btnSave.addEventListener('click', async function () {
    var title = pageTitleInput.value.trim();
    var url = pageUrlInput.value.trim();
    var currentSettings = await BookmarkApi.getSettings();

    // 判断目标工作表是下拉选择还是新建输入
    var isNewSheetMode = (sheetInputWrapper.style.display !== 'none');
    var targetSheet = '';
    if (isNewSheetMode) {
      targetSheet = targetSheetInput.value.trim();
      if (!targetSheet) targetSheet = 'Bookmarks';
    } else {
      targetSheet = targetSheetSelect.value;
      if (targetSheet === '__NEW__') targetSheet = 'Bookmarks';
    }

    // 判断分组目录是下拉选择还是新建输入
    var isNewFolderMode = (folderInputWrapper.style.display !== 'none');
    var folderPath = '';
    if (isNewFolderMode) {
      folderPath = cleanCategoryString(folderInput.value.trim());
      if (!folderPath) folderPath = '书签栏';
    } else {
      folderPath = cleanCategoryString(folderSelect.value);
      if (folderPath === '__NEW__' || !folderPath) folderPath = '书签栏';
    }

    var notes = notesInput.value.trim();

    if (!url) {
      showFeedback('error', '请输入有效的网址！');
      return;
    }

    if (!currentSettings.webAppUrl) {
      showFeedback('warning', '请先在插件设置中配置 Google Apps Script Web App URL！');
      openOptionsPage();
      return;
    }

    setLoading(true);

    try {
      var result = await BookmarkApi.saveBookmark(
        {
          title: title || '未命名书签',
          url: url,
          folderPath: folderPath,
          notes: notes,
          dateAdded: formatDateTime(new Date())
        },
        {
          sheetName: targetSheet
        }
      );

      // 记忆用户本次选择的工作表与目录
      await setLocalCache('last_target_sheet', targetSheet);
      await setLocalCache('last_folder_path', folderPath);

      if (result.status === 'success') {
        showFeedback('success', '✅ ' + (result.message || '书签已成功保存！'), result.sheetUrl);

        // 将本次输入的目录动态加入该工作表的分类缓存并重绘下拉框
        var curCats = (await getLocalCache('cached_categories_' + targetSheet)) || [];
        curCats = curCats.map(cleanCategoryString).filter(Boolean);
        if (!curCats.includes(folderPath)) {
          curCats.push(folderPath);
          await setLocalCache('cached_categories_' + targetSheet, curCats);
        }
        renderFolderOptions(curCats, folderPath);
        folderInputWrapper.style.display = 'none';
        folderSelectWrapper.style.display = 'block';
        folderSelect.value = folderPath;

        // 如果是新建的工作表，将其加入工作表列表缓存
        var curSheets = (await getLocalCache('cached_sheet_names')) || [];
        if (!curSheets.includes(targetSheet)) {
          curSheets.push(targetSheet);
          await setLocalCache('cached_sheet_names', curSheets);
        }
      } else if (result.status === 'duplicate') {
        showFeedback('warning', 'ℹ️ ' + (result.message || '该网页已在 Google Sheet 中'), result.sheetUrl);
      } else {
        showFeedback('error', '❌ 保存失败: ' + (result.message || '未知错误'));
      }
    } catch (err) {
      showFeedback('error', '❌ 网络错误或配置异常: ' + err.message);
    } finally {
      setLoading(false);
    }
  });

  // 快捷跳转
  function openOptionsPage() {
    if (typeof chrome !== 'undefined' && chrome.runtime && chrome.runtime.openOptionsPage) {
      chrome.runtime.openOptionsPage();
    } else {
      window.open('../options/options.html');
    }
  }

  function openImportPage() {
    if (typeof chrome !== 'undefined' && chrome.tabs && chrome.tabs.create) {
      chrome.tabs.create({ url: chrome.runtime.getURL('import/import.html') });
    } else {
      window.open('../import/import.html');
    }
  }

  btnGoConfig.addEventListener('click', openOptionsPage);
  btnOpenOptions.addEventListener('click', openOptionsPage);
  btnOpenImport.addEventListener('click', openImportPage);
  btnImportPageLink.addEventListener('click', openImportPage);

  function setLoading(isLoading) {
    btnSave.disabled = isLoading;
    if (isLoading) {
      btnText.style.display = 'none';
      btnLoading.style.display = 'inline-flex';
    } else {
      btnText.style.display = 'inline';
      btnLoading.style.display = 'none';
    }
  }

  function showFeedback(type, text, sheetUrl) {
    feedbackBox.style.display = 'flex';
    feedbackBox.className = 'feedback-box feedback-' + type;
    feedbackText.textContent = text;

    if (sheetUrl) {
      feedbackSheetLink.style.display = 'inline-flex';
      feedbackSheetLink.href = sheetUrl;
    } else {
      feedbackSheetLink.style.display = 'none';
    }
  }

  function formatDateTime(d) {
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getFullYear() + '-' +
      pad(d.getMonth() + 1) + '-' +
      pad(d.getDate()) + ' ' +
      pad(d.getHours()) + ':' +
      pad(d.getMinutes()) + ':' +
      pad(d.getSeconds());
  }
});