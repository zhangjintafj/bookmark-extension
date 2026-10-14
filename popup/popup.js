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
  var sheetLoadingTip = document.getElementById('sheetLoadingTip');
  var folderPathInput = document.getElementById('folderPath');
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

  // 1. 读取设置与上次选择的工作表
  var settings = await BookmarkApi.getSettings();
  if (!settings.webAppUrl) {
    unconfiguredAlert.style.display = 'flex';
  }

  var savedSheet = localStorage.getItem('last_target_sheet') || settings.defaultSheetName || 'Bookmarks';
  folderPathInput.value = settings.defaultFolder || '未分类';

  // 下拉切换与新建输入模式
  targetSheetSelect.addEventListener('change', function () {
    if (targetSheetSelect.value === '__NEW__') {
      sheetSelectWrapper.style.display = 'none';
      sheetInputWrapper.style.display = 'block';
      targetSheetInput.value = '';
      targetSheetInput.focus();
    }
  });

  btnCancelNewSheet.addEventListener('click', function () {
    sheetInputWrapper.style.display = 'none';
    sheetSelectWrapper.style.display = 'block';
    targetSheetSelect.value = targetSheetSelect.options[0] ? targetSheetSelect.options[0].value : 'Bookmarks';
  });

  // 2. 异步实时加载 Google 表格中的所有工作表供下拉选择
  if (settings.webAppUrl) {
    sheetLoadingTip.style.display = 'inline';
    BookmarkApi.fetchSheetNames().then(function (sheets) {
      sheetLoadingTip.style.display = 'none';
      if (sheets && sheets.length > 0) {
        targetSheetSelect.innerHTML = '';
        var hasMatch = false;
        sheets.forEach(function (name) {
          var opt = document.createElement('option');
          opt.value = name;
          opt.textContent = name;
          if (name === savedSheet) {
            opt.selected = true;
            hasMatch = true;
          }
          targetSheetSelect.appendChild(opt);
        });
        var newOpt = document.createElement('option');
        newOpt.value = '__NEW__';
        newOpt.textContent = '➕ 新建工作表...';
        targetSheetSelect.appendChild(newOpt);

        if (!hasMatch && sheets[0]) {
          targetSheetSelect.value = sheets[0];
        }
      }
    }).catch(function () {
      sheetLoadingTip.style.display = 'none';
    });
  }

  // 3. 获取当前浏览器活动标签页
  if (typeof chrome !== 'undefined' && chrome.tabs && chrome.tabs.query) {
    chrome.tabs.query({ active: true, currentWindow: true }, function (tabs) {
      if (tabs && tabs[0]) {
        var currentTab = tabs[0];
        pageTitleInput.value = currentTab.title || '';
        pageUrlInput.value = currentTab.url || '';
      }
    });
  }

  // 4. 异步尝试加载已有的分类列表
  if (settings.webAppUrl) {
    BookmarkApi.fetchCategories().then(function (categories) {
      if (categories && categories.length > 0) {
        folderSuggestions.innerHTML = '';
        categories.forEach(function (cat) {
          var option = document.createElement('option');
          option.value = cat;
          folderSuggestions.appendChild(option);
        });
      }
    }).catch(function () {});
  }

  // 5. 事件绑定：保存书签
  btnSave.addEventListener('click', async function () {
    var title = pageTitleInput.value.trim();
    var url = pageUrlInput.value.trim();
    var currentSettings = await BookmarkApi.getSettings();

    // 区分是下拉选择还是新建输入
    var isNewMode = (sheetInputWrapper.style.display !== 'none');
    var targetSheet = '';
    if (isNewMode) {
      targetSheet = targetSheetInput.value.trim() || 'Bookmarks';
    } else {
      targetSheet = targetSheetSelect.value;
      if (targetSheet === '__NEW__') targetSheet = 'Bookmarks';
    }

    var folderPath = folderPathInput.value.trim() || '书签栏';
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

      // 记忆用户本次选择的工作表
      localStorage.setItem('last_target_sheet', targetSheet);

      if (result.status === 'success') {
        showFeedback('success', '✅ ' + (result.message || '书签已成功保存！'), result.sheetUrl);
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