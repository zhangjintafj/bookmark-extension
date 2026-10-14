document.addEventListener('DOMContentLoaded', async function () {
  var urlInput = document.getElementById('webAppUrl');
  var sheetInput = document.getElementById('defaultSheetName');
  var folderInput = document.getElementById('defaultFolder');
  var statusBox = document.getElementById('statusBox');

  var cfg = await BookmarkApi.getSettings();
  urlInput.value = cfg.webAppUrl || '';
  sheetInput.value = cfg.defaultSheetName || 'Bookmarks';
  folderInput.value = cfg.defaultFolder || '未分类';

  document.getElementById('btnTest').onclick = async function () {
    statusBox.style.display = 'block';
    statusBox.style.background = '#E8F0FE';
    statusBox.textContent = '正在测试连接...';
    try {
      var res = await BookmarkApi.testConnection(urlInput.value.trim());
      statusBox.style.background = '#E6F4EA';
      statusBox.textContent = '✅ ' + (res.message || '连接成功！');
    } catch (e) {
      statusBox.style.background = '#FCE8E6';
      statusBox.textContent = '❌ 连接失败: ' + e.message;
    }
  };

  document.getElementById('btnSave').onclick = async function () {
    await BookmarkApi.saveSettings({
      webAppUrl: urlInput.value.trim(),
      defaultSheetName: sheetInput.value.trim() || 'Bookmarks',
      defaultFolder: folderInput.value.trim() || '未分类'
    });
    statusBox.style.display = 'block';
    statusBox.style.background = '#E6F4EA';
    statusBox.textContent = '✅ 设置保存成功！';
  };
});