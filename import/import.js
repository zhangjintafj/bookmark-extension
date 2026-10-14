document.addEventListener('DOMContentLoaded', async function () {
  var dropZone = document.getElementById('dropZone');
  var fileInput = document.getElementById('fileInput');
  var fileInfo = document.getElementById('fileInfo');
  var settingsCard = document.getElementById('settingsCard');
  var btnStartImport = document.getElementById('btnStartImport');
  var progressBox = document.getElementById('progressBox');
  var progressBar = document.getElementById('progressBar');
  var progressText = document.getElementById('progressText');
  var resultBox = document.getElementById('resultBox');
  var sheetLink = document.getElementById('sheetLink');

  var parsedData = null;

  dropZone.onclick = () => fileInput.click();
  dropZone.ondragover = (e) => { e.preventDefault(); };
  dropZone.ondrop = (e) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
  };
  fileInput.onchange = () => {
    if (fileInput.files[0]) handleFile(fileInput.files[0]);
  };

  function handleFile(file) {
    document.getElementById('fileName').textContent = file.name;
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        parsedData = BookmarkParser.parseHtmlBookmarks(e.target.result);
        document.getElementById('statBookmarks').textContent = parsedData.stats.totalBookmarks;
        document.getElementById('statFolders').textContent = parsedData.stats.totalFolders;
        fileInfo.style.display = 'block';
        settingsCard.style.display = 'block';
      } catch (err) {
        alert('解析失败: ' + err.message);
      }
    };
    reader.readAsText(file);
  }

  btnStartImport.onclick = async function () {
    if (!parsedData) return;
    var settings = await BookmarkApi.getSettings();
    if (!settings.webAppUrl) return alert('请先在插件设置中填写 Web App URL！');

    btnStartImport.disabled = true;
    progressBox.style.display = 'block';
    resultBox.style.display = 'none';

    var viewMode = document.querySelector('input[name="viewMode"]:checked').value;
    var sheetName = document.getElementById('targetSheetName').value.trim() || 'Bookmarks';
    var clearSheet = document.getElementById('chkClearSheet').checked;

    try {
      var res = await BookmarkApi.batchImportBookmarks(
        parsedData.bookmarks,
        { webAppUrl: settings.webAppUrl, sheetName: sheetName, viewMode: viewMode, clearSheet: clearSheet },
        function (p) {
          progressBar.style.width = p.percentage + '%';
          progressText.textContent = `正在导入: ${p.processed} / ${p.total} (${p.percentage}%)`;
        }
      );
      resultBox.style.display = 'block';
      sheetLink.href = res.sheetUrl;
    } catch (e) {
      alert('导入失败: ' + e.message);
    } finally {
      btnStartImport.disabled = false;
    }
  };
});