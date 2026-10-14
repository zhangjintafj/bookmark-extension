(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.BookmarkParser = factory();
  }
})(typeof self !== 'undefined' ? self : this, function () {

  function parseHtmlBookmarks(htmlContent) {
    if (!htmlContent || typeof htmlContent !== 'string') {
      throw new Error('书签文件内容为空或格式无效');
    }

    var tagRegex = /<H3(?:\s+[^>]*)?>([\s\S]*?)<\/H3>|<DL(?:\s+[^>]*)?>|<\/DL>|<A\s+([^>]+)>([\s\S]*?)<\/A>|<DD>([\s\S]*?)(?=(?:<DT|<DL|<\/DL|<p|$))/gi;

    var pathStack = [];
    var folderStack = [];
    var pendingFolderName = '';
    var bookmarks = [];
    var folderSet = new Set();
    var maxDepth = 0;

    var rootTree = {
      name: "全部书签",
      type: "folder",
      children: [],
      level: 0
    };
    var treeStack = [rootTree];
    var lastBookmark = null;
    var match;

    while ((match = tagRegex.exec(htmlContent)) !== null) {
      var fullMatch = match[0];

      if (match[1] !== undefined) {
        var rawName = match[1].replace(/<[^>]+>/g, '').trim();
        pendingFolderName = decodeHtmlEntities(rawName) || "未命名文件夹";
      } else if (/^<DL/i.test(fullMatch)) {
        if (pendingFolderName) {
          var folderName = pendingFolderName;
          pendingFolderName = '';
          pathStack.push(folderName);
          folderSet.add(pathStack.join(' / '));
          if (pathStack.length > maxDepth) maxDepth = pathStack.length;

          var folderNode = {
            name: folderName,
            type: "folder",
            path: pathStack.join(' / '),
            level: pathStack.length,
            children: []
          };

          treeStack[treeStack.length - 1].children.push(folderNode);
          treeStack.push(folderNode);
          folderStack.push(true);
        } else {
          folderStack.push(false);
        }
      } else if (/^<\/DL/i.test(fullMatch)) {
        var isNamedFolder = folderStack.pop();
        if (isNamedFolder) {
          if (pathStack.length > 0) pathStack.pop();
          if (treeStack.length > 1) treeStack.pop();
        }
      } else if (match[2] !== undefined && match[3] !== undefined) {
        var attrsStr = match[2];
        var rawTitle = match[3].replace(/<[^>]+>/g, '').trim();
        var title = decodeHtmlEntities(rawTitle) || "未命名书签";

        var hrefMatch = attrsStr.match(/HREF=(["'])(.*?)\1/i) || attrsStr.match(/HREF=([^\s>]+)/i);
        var url = hrefMatch ? (hrefMatch[2] || hrefMatch[1] || '').trim() : '';

        if (!url || url.toLowerCase().startsWith('javascript:')) continue;

        var addDateMatch = attrsStr.match(/ADD_DATE=(["']?)(\d+)\1/i);
        var dateAdded = '';
        if (addDateMatch) {
          var ts = parseInt(addDateMatch[2], 10);
          if (ts < 10000000000) ts *= 1000;
          var d = new Date(ts);
          if (!isNaN(d.getTime())) dateAdded = formatDateTime(d);
        }
        if (!dateAdded) dateAdded = formatDateTime(new Date());

        var currentPath = pathStack.slice();
        var folderPathStr = currentPath.length > 0 ? currentPath.join(' / ') : "根目录";

        var bm = {
          title: title,
          url: url,
          folderPath: folderPathStr,
          pathArray: currentPath,
          level: currentPath.length,
          dateAdded: dateAdded,
          notes: ''
        };

        bookmarks.push(bm);
        lastBookmark = bm;

        treeStack[treeStack.length - 1].children.push({
          name: title,
          type: "bookmark",
          url: url,
          dateAdded: dateAdded,
          level: currentPath.length
        });
      } else if (match[4] !== undefined) {
        if (lastBookmark) {
          var rawNotes = match[4].replace(/<[^>]+>/g, '').trim();
          lastBookmark.notes = decodeHtmlEntities(rawNotes);
        }
      }
    }

    return {
      bookmarks: bookmarks,
      tree: rootTree,
      stats: {
        totalBookmarks: bookmarks.length,
        totalFolders: folderSet.size,
        maxDepth: maxDepth
      }
    };
  }

  function decodeHtmlEntities(str) {
    if (!str) return '';
    return str
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ');
  }

  function formatDateTime(date) {
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };
    return date.getFullYear() + '-' + pad(date.getMonth() + 1) + '-' + pad(date.getDate()) + ' ' +
      pad(date.getHours()) + ':' + pad(date.getMinutes()) + ':' + pad(date.getSeconds());
  }

  return { parseHtmlBookmarks: parseHtmlBookmarks, formatDateTime: formatDateTime };
});