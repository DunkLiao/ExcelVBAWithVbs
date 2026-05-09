/**
 * frontend/app.js
 * Excel VBA 模組庫靜態網站前端邏輯
 *
 * 優先使用 window.VBA_DATA（由 app-data.js 嵌入，支援 file:// 協定）
 * 若不存在則 fetch data.json（GitHub Pages 等 HTTP 環境）
 */

(function () {
  "use strict";

  /* ── 狀態 ─────────────────────────────────────── */
  let allData = null;          // { modules: [...] }
  let activeFile = null;       // 目前選取的 file 物件
  let searchTimer = null;

  /* ── DOM refs ─────────────────────────────────── */
  const treePanel    = document.getElementById("tree-panel");
  const searchInput  = document.getElementById("search-input");
  const searchClear  = document.getElementById("search-clear");
  const searchResults = document.getElementById("search-results");
  const fileView     = document.getElementById("file-view");
  const statsEl      = document.getElementById("stats");
  const themeBtn     = document.getElementById("theme-toggle");

  /* ── 主題切換 ──────────────────────────────────── */
  var DARK = "dark";
  var STORAGE_KEY = "vba-theme";

  function applyTheme(theme) {
    if (theme === DARK) {
      document.documentElement.setAttribute("data-theme", DARK);
      themeBtn.textContent = "☀️";
      themeBtn.title = "切換白天模式";
    } else {
      document.documentElement.removeAttribute("data-theme");
      themeBtn.textContent = "🌙";
      themeBtn.title = "切換黑夜模式";
    }
  }

  (function initTheme() {
    var saved = localStorage.getItem(STORAGE_KEY) || "light";
    applyTheme(saved);
  })();

  themeBtn.addEventListener("click", function () {
    var current = document.documentElement.getAttribute("data-theme");
    var next = current === DARK ? "light" : DARK;
    applyTheme(next);
    localStorage.setItem(STORAGE_KEY, next);
  });

  /* ── 初始化 ────────────────────────────────────── */
  function init() {
    if (window.VBA_DATA) {
      loadData(window.VBA_DATA);
    } else {
      fetch("data.json")
        .then(function (r) { return r.json(); })
        .then(loadData)
        .catch(function (err) {
          treePanel.innerHTML =
            '<p style="color:#f38ba8;padding:16px">⚠ 無法載入資料，請先執行 build.py<br><small>' +
            err.message + "</small></p>";
        });
    }
  }

  function loadData(data) {
    allData = data;

    // 統計列
    statsEl.textContent =
      data.total_folders + " 資料夾　" + data.total_files + " 個檔案";

    renderTree(data.modules);
  }

  /* ── 目錄樹 ────────────────────────────────────── */
  function renderTree(modules) {
    treePanel.innerHTML = "";
    modules.forEach(function (folder) {
      treePanel.appendChild(buildFolderNode(folder));
    });
  }

  function buildFolderNode(folder) {
    var div = document.createElement("div");
    div.className = "tree-folder";

    var header = document.createElement("div");
    header.className = "folder-header";
    header.innerHTML =
      '<span class="folder-toggle">▾</span>' +
      '<span>📁 ' + escapeHtml(folder.folder) + "</span>" +
      '<span style="color:var(--text-dim);font-weight:400;margin-left:auto;font-size:0.75rem;">' +
      folder.files.length + "</span>";
    header.addEventListener("click", function () {
      div.classList.toggle("collapsed");
    });

    var fileList = document.createElement("div");
    fileList.className = "folder-files";

    folder.files.forEach(function (file) {
      var item = document.createElement("div");
      item.className = "file-item";
      item.textContent = "📄 " + file.name;
      item.title = file.path;
      item.addEventListener("click", function () {
        openFile(file, item);
      });
      fileList.appendChild(item);
    });

    div.appendChild(header);
    div.appendChild(fileList);
    return div;
  }

  /* ── 檔案預覽 ──────────────────────────────────── */
  function openFile(file, itemEl, highlightKeyword) {
    // 清除舊 active
    var prev = treePanel.querySelector(".file-item.active");
    if (prev) prev.classList.remove("active");
    if (itemEl) itemEl.classList.add("active");

    activeFile = file;

    var content = file.content || "";
    var displayContent = highlightKeyword
      ? highlightText(content, highlightKeyword)
      : escapeHtml(content);

    fileView.innerHTML =
      '<div class="file-header">' +
      '<div class="file-header-top">' +
      '<h2>📄 ' + escapeHtml(file.name) + "</h2>" +
      '<button class="btn-download" onclick="downloadFile()" title="下載原始檔案（CP950 編碼）">⬇ 下載</button>' +
      "</div>" +
      '<div class="file-meta">' + escapeHtml(file.path) + "</div>" +
      "</div>" +
      '<div class="code-block"><pre>' + displayContent + "</pre></div>";
  }

  /* ── 下載（保留原始 CP950 binary） ─────────────── */
  window.downloadFile = function () {
    if (!activeFile) return;
    var b64 = activeFile.content_b64;
    if (!b64) {
      alert("此檔案無法下載（缺少 content_b64，請重新執行 build.py）");
      return;
    }
    // atob → binary string → Uint8Array → Blob
    var binary = atob(b64);
    var bytes = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    var blob = new Blob([bytes], { type: "application/octet-stream" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = activeFile.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  /* ── 搜尋 ──────────────────────────────────────── */
  searchInput.addEventListener("input", function () {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(runSearch, 250);
    searchClear.hidden = !searchInput.value;
  });

  searchClear.addEventListener("click", function () {
    searchInput.value = "";
    searchClear.hidden = true;
    clearSearch();
  });

  function runSearch() {
    var kw = searchInput.value.trim();
    if (!kw) { clearSearch(); return; }

    var kwLower = kw.toLowerCase();
    var results = [];

    allData.modules.forEach(function (folder) {
      folder.files.forEach(function (file) {
        var nameMatch = file.name.toLowerCase().includes(kwLower);
        var contentIdx = file.content ? file.content.toLowerCase().indexOf(kwLower) : -1;
        var contentMatch = contentIdx >= 0;

        if (nameMatch || contentMatch) {
          results.push({
            file: file,
            nameMatch: nameMatch,
            snippet: contentMatch ? extractSnippet(file.content, contentIdx, kw.length) : null,
          });
        }
      });
    });

    renderSearchResults(results, kw);
  }

  function extractSnippet(content, idx, kwLen) {
    var start = Math.max(0, idx - 60);
    var end = Math.min(content.length, idx + kwLen + 120);
    var snippet = "";
    if (start > 0) snippet += "…";
    snippet += content.slice(start, end);
    if (end < content.length) snippet += "…";
    return snippet;
  }

  function renderSearchResults(results, kw) {
    fileView.hidden = true;
    searchResults.hidden = false;

    if (!results.length) {
      searchResults.innerHTML =
        '<h2>搜尋「' + escapeHtml(kw) + '」— 無符合結果</h2>';
      return;
    }

    var html =
      '<h2>搜尋「' + escapeHtml(kw) + '」— 共 ' + results.length + " 筆結果</h2>";

    results.forEach(function (r) {
      var snippetHtml = r.snippet
        ? '<div class="result-snippet">' + highlightText(r.snippet, kw) + "</div>"
        : "";
      html +=
        '<div class="result-item" data-path="' + escapeAttr(r.file.path) + '">' +
        '<div class="result-name">' + highlightText(r.file.name, kw) + "</div>" +
        '<div class="result-path">' + escapeHtml(r.file.path) + "</div>" +
        snippetHtml +
        "</div>";
    });

    searchResults.innerHTML = html;

    // 綁定點擊
    var items = searchResults.querySelectorAll(".result-item");
    items.forEach(function (el) {
      var path = el.getAttribute("data-path");
      el.addEventListener("click", function () {
        var file = findFileByPath(path);
        if (file) {
          showFileView();
          openFile(file, null, kw);
          // 捲動至目錄中的對應項目
          scrollToFileInTree(path);
        }
      });
    });
  }

  function clearSearch() {
    searchResults.hidden = true;
    fileView.hidden = false;
  }

  function showFileView() {
    searchResults.hidden = true;
    fileView.hidden = false;
  }

  /* ── 工具函式 ──────────────────────────────────── */
  function findFileByPath(path) {
    for (var i = 0; i < allData.modules.length; i++) {
      var folder = allData.modules[i];
      for (var j = 0; j < folder.files.length; j++) {
        if (folder.files[j].path === path) return folder.files[j];
      }
    }
    return null;
  }

  function scrollToFileInTree(path) {
    var items = treePanel.querySelectorAll(".file-item");
    for (var i = 0; i < items.length; i++) {
      var el = items[i];
      if (el.title === path) {
        // 展開父資料夾
        var parent = el.closest(".tree-folder");
        if (parent) parent.classList.remove("collapsed");
        // 移除舊 active，加上新 active
        var prev = treePanel.querySelector(".file-item.active");
        if (prev) prev.classList.remove("active");
        el.classList.add("active");
        el.scrollIntoView({ block: "nearest" });
        break;
      }
    }
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function escapeAttr(str) {
    return String(str).replace(/"/g, "&quot;");
  }

  function highlightText(text, keyword) {
    if (!keyword) return escapeHtml(text);
    var escaped = escapeHtml(text);
    var kwEsc = escapeHtml(keyword);
    // case-insensitive 替換
    var regex = new RegExp(
      "(" + kwEsc.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + ")",
      "gi"
    );
    return escaped.replace(regex, "<mark>$1</mark>");
  }

  /* ── 啟動 ──────────────────────────────────────── */
  init();
})();
