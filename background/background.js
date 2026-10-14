importScripts('../lib/api.js');

chrome.runtime.onInstalled.addListener(function () {
  chrome.contextMenus.create({
    id: "save-page-to-gsheet",
    title: "📑 存入当前页面到 Google Sheet",
    contexts: ["page"]
  });

  chrome.contextMenus.create({
    id: "save-link-to-gsheet",
    title: "🔗 存入该链接到 Google Sheet",
    contexts: ["link"]
  });
});

chrome.contextMenus.onClicked.addListener(async function (info, tab) {
  var url = (info.menuItemId === "save-page-to-gsheet") ? (info.pageUrl || tab?.url || "") : (info.linkUrl || "");
  var title = (info.menuItemId === "save-page-to-gsheet") ? (tab?.title || url) : (info.selectionText || url);
  if (!url) return;

  var settings = await BookmarkApi.getSettings();
  if (!settings.webAppUrl) {
    chrome.runtime.openOptionsPage();
    return;
  }

  try {
    var result = await BookmarkApi.saveBookmark({
      title: title,
      url: url,
      folderPath: settings.defaultFolder || "未分类",
      notes: "通过右键快捷菜单添加",
      tags: "快捷收藏",
      dateAdded: new Date().toISOString().replace('T', ' ').substring(0, 19)
    });
    if (chrome.notifications) {
      chrome.notifications.create({
        type: "basic",
        iconUrl: "../icons/icon48.png",
        title: result.status === "success" ? "保存成功" : "提示",
        message: result.message || "已存入 Google Sheet"
      });
    }
  } catch (err) {}
});