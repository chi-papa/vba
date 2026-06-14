// ユーザー指定の「Local Secure Notes - v13」HTML全コード
export const htmlNotebookCode = `<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Local Secure Notes - v13</title>
  <style>
    /* 4px Grid System & CSS Variables */
    :root {
      --color-bg: #ffffff;
      --color-sidebar: #f7f9fa;
      --color-surface: #ffffff;
      --color-border: #e1e4e8;
      --color-text-main: #24292f;
      --color-text-muted: #57606a;
      --color-accent: #0969da;
      --color-accent-hover: #0349b4;
      --color-danger: #cf222e;
      --color-danger-hover: #a40e26;
      --color-success: #1a7f37;
      --color-hover: #f3f4f6;
      --color-active: #e5e7eb;
      --color-modal-overlay: rgba(0, 0, 0, 0.4);
      
      --sp-4: 4px; --sp-8: 8px; --sp-12: 12px; --sp-16: 16px; --sp-24: 24px; --sp-32: 32px;
      --radius-sm: 4px; --radius-md: 8px; --radius-lg: 12px;
      --font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
    }

    [data-theme="dark"] {
      --color-bg: #0d1117;
      --color-sidebar: #161b22;
      --color-surface: #161b22;
      --color-border: #30363d;
      --color-text-main: #c9d1d9;
      --color-text-muted: #8b949e;
      --color-accent: #58a6ff;
      --color-accent-hover: #79c0ff;
      --color-danger: #f85149;
      --color-danger-hover: #da3633;
      --color-hover: #21262d;
      --color-active: #30363d;
      --color-modal-overlay: rgba(0, 0, 0, 0.7);
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body { font-family: var(--font-family); background-color: var(--color-bg); color: var(--color-text-main); line-height: 1.6; -webkit-font-smoothing: antialiased; height: 100vh; overflow: hidden; display: flex; }
    ::-webkit-scrollbar { width: 6px; height: 6px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: var(--color-border); border-radius: 8px; }
    ::-webkit-scrollbar-thumb:hover { background: var(--color-text-muted); }

    .sidebar { width: 340px; min-width: 280px; height: 100vh; background-color: var(--color-sidebar); border-right: 1px solid var(--color-border); display: flex; flex-direction: column; }
    .main-content { flex: 1; height: 100vh; display: flex; flex-direction: column; background-color: var(--color-bg); position: relative; }

    .sidebar-header { padding: var(--sp-16) var(--sp-16) 0 var(--sp-16); }
    .sidebar-header h1 { font-size: 14px; font-weight: 600; color: var(--color-text-muted); margin-bottom: var(--sp-12); display: flex; justify-content: space-between; align-items: center;}
    .search-box { width: 100%; padding: var(--sp-8); border: 1px solid var(--color-border); border-radius: var(--radius-sm); background-color: var(--color-surface); color: var(--color-text-main); font-size: 13px; outline: none; margin-bottom: var(--sp-12); }
    .search-box:focus { border-color: var(--color-accent); }
    .btn-new { width: 100%; padding: var(--sp-8); background-color: var(--color-accent); color: #fff; border: none; border-radius: var(--radius-sm); font-size: 13px; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: var(--sp-8); margin-bottom: var(--sp-12); }
    .btn-new:hover { background-color: var(--color-accent-hover); }
    .btn-danger-outline { background-color: transparent; border: 1px solid var(--color-danger); color: var(--color-danger); }
    .btn-danger-outline:hover { background-color: var(--color-danger); color: white; }

    .sidebar-tabs { display: flex; border-bottom: 1px solid var(--color-border); margin-bottom: var(--sp-8); }
    .tab-btn { flex: 1; padding: var(--sp-8); text-align: center; font-size: 13px; font-weight: 600; cursor: pointer; border: none; background: transparent; color: var(--color-text-muted); border-bottom: 2px solid transparent; transition: all 0.2s; }
    .tab-btn:hover { color: var(--color-text-main); }
    .tab-btn.active { color: var(--color-accent); border-bottom-color: var(--color-accent); }

    .sidebar-content { flex: 1; overflow-y: auto; padding: 0 var(--sp-8) var(--sp-8) var(--sp-8); }
    .section-title { font-size: 11px; font-weight: bold; color: var(--color-text-muted); padding: var(--sp-8); }
    
    .tree-node { margin-bottom: 2px; }
    .note-item { padding: 6px var(--sp-8); border-radius: var(--radius-sm); cursor: pointer; display: flex; align-items: center; gap: var(--sp-4); transition: background-color 0.1s; border-left: 3px solid transparent; position: relative; }
    .note-item:hover { background-color: var(--color-hover); }
    .note-item.active { background-color: var(--color-active); border-left-color: var(--color-accent); }
    
    .tree-toggle { width: 16px; height: 16px; display: flex; align-items: center; justify-content: center; font-size: 10px; color: var(--color-text-muted); cursor: pointer; user-select: none; border-radius: 4px; }
    .tree-toggle:hover { background-color: var(--color-border); color: var(--color-text-main); }
    .tree-toggle.empty { opacity: 0; pointer-events: none; }
    .tree-toggle.open { transform: rotate(90deg); }

    .note-item-content { flex: 1; min-width: 0; display: flex; flex-direction: column; }
    .note-title { font-size: 13px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; color: var(--color-text-main); }
    .note-item.pinned .note-title::before { content: "★ "; color: var(--color-accent); font-size: 11px; }

    .note-item-actions { display: flex; opacity: 0; gap: 4px; transition: opacity 0.1s; }
    .note-item:hover .note-item-actions, .flat-note-item:hover .note-item-actions { opacity: 1; }
    .btn-quick { padding: 2px 8px; font-size: 11px; font-weight: bold; background-color: var(--color-surface); border: 1px solid var(--color-border); color: var(--color-text-main); border-radius: 4px; cursor: pointer; display: flex; align-items: center; gap: 4px; }
    .btn-quick:hover { background-color: var(--color-hover); border-color: var(--color-text-muted); }
    .btn-quick-primary { color: var(--color-accent); border-color: var(--color-accent); }
    .btn-quick-primary:hover { background-color: var(--color-accent); color: white; border-color: var(--color-accent); }

    .tree-children { padding-left: 16px; margin-left: 6px; border-left: 1px solid var(--color-border); display: none; }
    .tree-children.open { display: block; }

    .flat-note-item { padding: var(--sp-8) var(--sp-12); margin-bottom: var(--sp-4); border-radius: var(--radius-sm); cursor: pointer; transition: background-color 0.1s; border-left: 3px solid transparent; display: flex; align-items: center; justify-content: space-between; gap: var(--sp-8); }
    .flat-note-item:hover { background-color: var(--color-hover); }
    .flat-note-item.active { background-color: var(--color-active); border-left-color: var(--color-accent); }
    .flat-note-content { flex: 1; min-width: 0; }
    .flat-note-path { font-size: 11px; color: var(--color-text-muted); margin-bottom: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .flat-note-title { font-size: 14px; font-weight: 600; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

    .top-bar { height: 50px; padding: 0 var(--sp-24); border-bottom: 1px solid var(--color-border); display: flex; justify-content: flex-end; align-items: center; gap: var(--sp-16); }
    .tool-btn { background: transparent; border: none; color: var(--color-text-muted); cursor: pointer; font-size: 14px; display: flex; align-items: center; gap: var(--sp-4); transition: color 0.2s; }
    .tool-btn:hover { color: var(--color-text-main); }
    
    .editor-container { flex: 1; padding: var(--sp-24) var(--sp-32) 100px var(--sp-32); overflow-y: auto; max-width: 900px; margin: 0 auto; width: 100%; }
    
    .breadcrumbs { font-size: 13px; color: var(--color-text-muted); margin-bottom: var(--sp-16); display: flex; flex-wrap: wrap; gap: 4px; align-items: center;}
    .breadcrumb-item { cursor: pointer; transition: color 0.2s; }
    .breadcrumb-item:hover { color: var(--color-accent); text-decoration: underline; }
    .breadcrumb-separator { color: var(--color-border); }

    .editor-title { width: 100%; font-size: 32px; font-weight: bold; border: none; outline: none; background: transparent; color: var(--color-text-main); margin-bottom: var(--sp-16); font-family: var(--font-family); }
    .editor-title::placeholder { color: var(--color-border); }
    .editor-title:disabled { color: var(--color-text-muted); }

    .editor-body { width: 100%; min-height: 60vh; border: none; outline: none; background: transparent; color: var(--color-text-main); font-size: 15px; line-height: 1.8; resize: none; font-family: var(--font-family); }
    .editor-body::placeholder { color: var(--color-border); }
    .editor-body:disabled { color: var(--color-text-muted); }

    .trash-banner { background-color: rgba(207, 34, 46, 0.1); border: 1px solid var(--color-danger); color: var(--color-danger); padding: var(--sp-12); border-radius: var(--radius-sm); margin-bottom: var(--sp-24); font-size: 13px; display: flex; align-items: center; gap: var(--sp-8); font-weight: 600; }

    .action-bar { position: absolute; bottom: var(--sp-24); right: var(--sp-24); display: flex; gap: var(--sp-8); background-color: var(--color-surface); padding: var(--sp-8); border-radius: var(--radius-md); box-shadow: 0 4px 12px rgba(0,0,0,0.1); border: 1px solid var(--color-border); }
    .btn-action { padding: var(--sp-8) var(--sp-16); border: none; border-radius: var(--radius-sm); font-size: 13px; font-weight: 600; cursor: pointer; background: transparent; color: var(--color-text-muted); }
    .btn-action:hover { background-color: var(--color-hover); color: var(--color-text-main); }
    .btn-save { background-color: var(--color-accent); color: white; }
    .btn-save:hover { background-color: var(--color-accent-hover); color: white; }
    .btn-danger:hover { color: var(--color-danger); }

    .toast-container { position: fixed; bottom: var(--sp-24); left: 360px; display: flex; flex-direction: column; gap: var(--sp-8); z-index: 1000; }
    .toast { padding: var(--sp-12) var(--sp-24); background-color: var(--color-text-main); color: var(--color-bg); border-radius: var(--radius-sm); font-size: 13px; opacity: 0; transform: translateY(10px); transition: all 0.3s; }
    .toast.show { opacity: 1; transform: translateY(0); }
    .empty-state { display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; color: var(--color-text-muted); font-size: 14px;}

    .modal-overlay { position: fixed; top: 0; left: 0; right: 0; bottom: 0; background-color: var(--color-modal-overlay); backdrop-filter: blur(2px); display: flex; align-items: center; justify-content: center; z-index: 2000; opacity: 0; pointer-events: none; transition: opacity 0.2s ease; }
    .modal-overlay.show { opacity: 1; pointer-events: auto; }
    .modal-content { background-color: var(--color-surface); width: 100%; max-width: 360px; padding: var(--sp-24); border-radius: var(--radius-lg); border: 1px solid var(--color-border); box-shadow: 0 12px 32px rgba(0,0,0,0.2); transform: translateY(20px); transition: transform 0.2s ease; }
    .modal-overlay.show .modal-content { transform: translateY(0); }
    .modal-title { font-size: 16px; font-weight: bold; margin-bottom: var(--sp-8); color: var(--color-text-main); }
    .modal-desc { font-size: 14px; color: var(--color-text-muted); margin-bottom: var(--sp-24); line-height: 1.5; }
    .modal-actions { display: flex; justify-content: flex-end; gap: var(--sp-12); }
    .btn-modal { padding: var(--sp-8) var(--sp-16); border-radius: var(--radius-sm); font-size: 13px; font-weight: 600; cursor: pointer; border: none; }
    .btn-modal-cancel { background-color: transparent; border: 1px solid var(--color-border); color: var(--color-text-main); }
    .btn-modal-cancel:hover { background-color: var(--color-hover); }
    .btn-modal-danger { background-color: var(--color-danger); color: white; }
    .btn-modal-danger:hover { background-color: var(--color-danger-hover); }

    /* スマホ最適化 (Responsive) */
    @media (max-width: 768px) {
      body { flex-direction: column; }
      .sidebar { width: 100%; height: 35vh; border-right: none; border-bottom: 1px solid var(--color-border); }
      .main-content { height: 65vh; }
      .top-bar { padding: 0 var(--sp-12); }
      .editor-container { padding: var(--sp-16) var(--sp-16) 80px var(--sp-16); }
      .toast-container { left: var(--sp-24); }
    }
  </style>
</head>
<body>

  <aside class="sidebar">
    <div class="sidebar-header">
      <h1>Secure Notes <span style="font-size:10px; font-weight:normal; color:var(--color-accent);">v13</span></h1>
      <button class="btn-new" id="newNoteBtn" onclick="createNewNote(null)">＋ 新規ノート作成</button>
      <button class="btn-new btn-danger-outline" id="emptyTrashBtn" style="display: none;" onclick="requestEmptyTrash()">🗑️ ゴミ箱を空にする</button>
      <input type="text" id="searchInput" class="search-box" placeholder="検索..." oninput="renderNoteList()">
    </div>
    
    <div class="sidebar-tabs">
      <button class="tab-btn active" id="tabActive" onclick="switchTab('active')">ノート</button>
      <button class="tab-btn" id="tabTrash" onclick="switchTab('trash')">ゴミ箱</button>
    </div>

    <div class="sidebar-content">
      <div id="pinnedSection" style="display: none; margin-bottom: var(--sp-8);">
        <div class="section-title">ピン留め</div>
        <div id="pinnedList"></div>
      </div>
      
      <div class="section-title" id="treeTitle" style="display: none;">すべてのノート</div>
      <div id="noteList"></div>
    </div>
  </aside>

  <main class="main-content">
    <div class="top-bar">
      <button class="tool-btn" onclick="document.getElementById('importFile').click()">📥 復元</button>
      <input type="file" id="importFile" accept=".json" style="display: none;" onchange="importData(event)">
      <button class="tool-btn" onclick="exportData()">📤 全バックアップ</button>
      <button class="tool-btn" onclick="toggleTheme()" id="themeToggle">🌙 ダークモード</button>
    </div>

    <div class="editor-container" id="editorContainer" style="display: none;">
      <div id="trashBanner" class="trash-banner" style="display: none;"></div>
      <div class="breadcrumbs" id="breadcrumbs"></div>
      <input type="text" id="noteTitle" class="editor-title" placeholder="無題のノート">
      <textarea id="noteBody" class="editor-body" placeholder="ここにメモやソースコードを記述します..."></textarea>
      
      <div class="action-bar" id="actionBar"></div>
    </div>

    <div class="empty-state" id="emptyState">
      <p>左側のリストからノートを選択するか、新しく作成してください。</p>
    </div>
  </main>

  <div class="modal-overlay" id="confirmModal" onclick="closeModal(event)">
    <div class="modal-content" onclick="event.stopPropagation()">
      <div class="modal-title" id="modalTitle">確認</div>
      <div class="modal-desc" id="modalDesc">この操作を実行してもよろしいですか？</div>
      <div class="modal-actions">
        <button class="btn-modal btn-modal-cancel" onclick="closeModal()">キャンセル</button>
        <button class="btn-modal btn-modal-danger" id="modalActionBtn">実行する</button>
      </div>
    </div>
  </div>

  <div class="toast-container" id="toastContainer"></div>

  <script>
    const STORAGE_KEY = 'chipapa_local_notes_v13';
    const EXPANDED_KEY = 'chipapa_expanded_folders';
    let notesData = [];
    let activeNoteId = null;
    let currentTab = 'active';
    let expandedFolders = new Set();
    let pendingModalAction = null;

    document.addEventListener("DOMContentLoaded", () => {
      initTheme();
      loadExpandedFolders();
      loadNotes();
      
      // ★ バグ修正⑥: 配列の先頭ではなく、最終更新日が最新のノートを初期選択する
      const activeNotes = notesData.filter(n => !n.trashed);
      if (activeNotes.length > 0) {
        activeNotes.sort((a, b) => new Date(b.updateDate || b.date) - new Date(a.updateDate || a.date));
        selectNote(activeNotes[0].id);
      } else {
        renderNoteList();
      }

      document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
          e.preventDefault();
          saveActiveNote(true);
        }
      });
    });

    // ★ バグ修正①： オートセーブ（サイレント保存）機能
    function autoSaveCurrentEdits() {
      if (!activeNoteId) return;
      const index = notesData.findIndex(n => n.id === activeNoteId);
      if (index === -1 || notesData[index].trashed || isParentTrashed(notesData[index].parentId)) return;
      
      const newTitle = document.getElementById('noteTitle').value.trim() || '無題';
      const newBody = document.getElementById('noteBody').value;
      
      // 変更があった場合のみ裏で保存
      if (notesData[index].title !== newTitle || notesData[index].body !== newBody) {
        notesData[index].title = newTitle;
        notesData[index].body = newBody;
        notesData[index].updateDate = new Date().toISOString();
        localStorage.setItem(STORAGE_KEY, JSON.stringify(notesData));
      }
    }

    function initTheme() {
      const savedTheme = localStorage.getItem('chipapa_theme') || 'dark';
      document.documentElement.setAttribute('data-theme', savedTheme);
      updateThemeBtn(savedTheme);
    }
    
    function toggleTheme() {
      const current = document.documentElement.getAttribute('data-theme');
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('chipapa_theme', next);
      updateThemeBtn(next);
    }
    
    function updateThemeBtn(theme) {
      document.getElementById('themeToggle').textContent = theme === 'dark' ? '☀️ ライトモード' : '🌙 ダークモード';
    }

    function loadExpandedFolders() {
      const saved = sessionStorage.getItem(EXPANDED_KEY);
      if (saved) expandedFolders = new Set(JSON.parse(saved));
    }

    // 省略されていたその他の関数実装...
    function saveExpandedFolders() {
      sessionStorage.setItem(EXPANDED_KEY, JSON.stringify(Array.from(expandedFolders)));
    }

    function loadNotes() {
      const saved = localStorage.getItem(STORAGE_KEY);
      const legacySavedV12 = localStorage.getItem('chipapa_local_notes_v12');

      if (saved) {
        try {
          notesData = JSON.parse(saved);
        } catch(e) {
          console.error('ノートデータの読み込みに失敗しました。データをリセットします。', e);
          notesData = [];
          localStorage.removeItem(STORAGE_KEY);
        }
      } else if (legacySavedV12) {
        try {
          notesData = JSON.parse(legacySavedV12);
          saveNotes();
        } catch(e) {
          console.error('v12データの読み込みに失敗しました。', e);
          notesData = [];
        }
      }
    }

    function saveNotes() {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(notesData));
      renderNoteList();
    }

    function getChildren(parentId) { return notesData.filter(n => n.parentId === parentId); }

    function getAllDescendants(parentId) {
      let descendants = [];
      const children = getChildren(parentId);
      for (const child of children) {
        descendants.push(child);
        descendants = descendants.concat(getAllDescendants(child.id));
      }
      return descendants;
    }

    function isParentTrashed(parentId, visited = new Set()) {
      if (!parentId || visited.has(parentId)) return false;
      visited.add(parentId);
      const parent = notesData.find(n => n.id === parentId);
      if (!parent) return false;
      if (parent.trashed) return true;
      return isParentTrashed(parent.parentId, visited);
    }

    function getBreadcrumbs(noteId) {
      const crumbs = [];
      const visited = new Set();
      let current = notesData.find(n => n.id === noteId);
      while (current && !visited.has(current.id)) {
        visited.add(current.id);
        crumbs.unshift(current);
        current = notesData.find(n => n.id === current.parentId);
      }
      return crumbs;
    }

    function switchTab(tab) {
      autoSaveCurrentEdits();
      currentTab = tab;
      document.getElementById('tabActive').classList.toggle('active', tab === 'active');
      document.getElementById('tabTrash').classList.toggle('active', tab === 'trash');
      document.getElementById('newNoteBtn').style.display = tab === 'active' ? 'flex' : 'none';
      document.getElementById('emptyTrashBtn').style.display = tab === 'trash' ? 'flex' : 'none';
      
      activeNoteId = null;
      document.getElementById('editorContainer').style.display = 'none';
      document.getElementById('emptyState').style.display = 'flex';
      renderNoteList();
    }

    function createNewNote(parentId = null) {
      autoSaveCurrentEdits();
      
      if (currentTab !== 'active') switchTab('active');
      document.getElementById('searchInput').value = '';
      
      const newId = 'note_' + Date.now();
      const newNote = {
        id: newId, parentId: parentId, title: '', body: '',
        date: new Date().toISOString(), updateDate: '', pinned: false, trashed: false
      };
      notesData.unshift(newNote);
      
      if (parentId) {
        expandedFolders.add(parentId);
        saveExpandedFolders();
      }
      
      saveNotes();
      selectNote(newId);
      document.getElementById('noteTitle').focus();
    }

    function selectNote(id) {
      autoSaveCurrentEdits();

      activeNoteId = id;
      const note = notesData.find(n => n.id === id);
      if (!note) return;

      document.getElementById('emptyState').style.display = 'none';
      document.getElementById('editorContainer').style.display = 'block';

      const crumbs = getBreadcrumbs(note.id);
      document.getElementById('breadcrumbs').innerHTML = crumbs.map((crumb, index) => {
        const isLast = index === crumbs.length - 1;
        return \`<span class="breadcrumb-item" onclick="selectNote('\${crumb.id}')" style="\${isLast ? 'color: var(--color-text-main); font-weight: 600;' : ''}">\${escapeHTML(crumb.title || '無題')}</span>\`;
      }).join('<span class="breadcrumb-separator">/</span>');

      const titleEl = document.getElementById('noteTitle');
      const bodyEl = document.getElementById('noteBody');
      const bannerEl = document.getElementById('trashBanner');

      titleEl.value = note.title;
      bodyEl.value = note.body;

      const implicitlyTrashed = isParentTrashed(note.parentId);
      const isTrashed = note.trashed || implicitlyTrashed;
      
      titleEl.disabled = isTrashed;
      bodyEl.disabled = isTrashed;
      
      bannerEl.style.display = isTrashed ? 'flex' : 'none';
      if (implicitlyTrashed && !note.trashed) {
        bannerEl.innerHTML = '⚠️ 親ノートがゴミ箱にあるため、このノートは読み取り専用です。';
      } else {
        bannerEl.innerHTML = '⚠️ このノートはゴミ箱に入っています。編集するには復元してください。';
      }

      renderActionBar(note, isTrashed, implicitlyTrashed);
      renderNoteList();
    }

    function renderActionBar(note, isTrashed, implicitlyTrashed) {
      const bar = document.getElementById('actionBar');
      if (isTrashed) {
        if (implicitlyTrashed && !note.trashed) {
          bar.innerHTML = \`<span style="font-size:12px; color:var(--color-text-muted);">親ノートを復元してください</span>\`;
        } else {
          bar.innerHTML = \`
            <button class="btn-action btn-danger" onclick="requestPermanentDelete()">🗑️ 完全に削除</button>
            <button class="btn-action btn-save" onclick="restoreNote()">⤴️ 復元する</button>
          \`;
        }
      } else {
        bar.innerHTML = \`
          <button class="btn-action" onclick="exportPartial('\${note.id}')" title="このノートと子ノートをファイル出力します">📤 共有出力</button>
          <button class="btn-action" onclick="copyActiveNote()" title="本文をコピー">📋 コピー</button>
          <button class="btn-action" onclick="togglePinActive()">\${note.pinned ? '★ ピン留め解除' : '☆ ピン留め'}</button>
          <button class="btn-action btn-danger" onclick="moveToTrash()">🗑️ ゴミ箱へ</button>
          <button class="btn-action btn-save" onclick="saveActiveNote(true)">保存 (Ctrl+S)</button>
        \`;
      }
    }

    function saveActiveNote(showNotification = false) {
      if (!activeNoteId) return;
      const index = notesData.findIndex(n => n.id === activeNoteId);
      if (index === -1 || notesData[index].trashed || isParentTrashed(notesData[index].parentId)) return;
      
      notesData[index].title = document.getElementById('noteTitle').value.trim() || '無題';
      notesData[index].body = document.getElementById('noteBody').value;
      notesData[index].updateDate = new Date().toISOString();

      saveNotes();
      if (showNotification) showToast("保存しました");
    }

    function moveToTrash() {
      autoSaveCurrentEdits();
      if (!activeNoteId) return;
      const index = notesData.findIndex(n => n.id === activeNoteId);
      if (index > -1) {
        notesData[index].trashed = true;
        notesData[index].pinned = false;
        const descendants = getAllDescendants(activeNoteId);
        descendants.forEach(d => {
           const dIndex = notesData.findIndex(n => n.id === d.id);
           notesData[dIndex].trashed = true;
           notesData[dIndex].pinned = false;
        });
        saveNotes();
        activeNoteId = null;
        document.getElementById('editorContainer').style.display = 'none';
        document.getElementById('emptyState').style.display = 'flex';
        showToast("ゴミ箱に移動しました");
      }
    }

    function restoreNote() {
      if (!activeNoteId) return;
      const index = notesData.findIndex(n => n.id === activeNoteId);
      if (index > -1) {
        notesData[index].trashed = false;
        let msg = "ノートを復元しました";

        if (isParentTrashed(notesData[index].parentId)) {
          notesData[index].parentId = null;
          msg = "親ノートがゴミ箱にあるため、一番上の階層に復元しました";
        }

        const descendants = getAllDescendants(activeNoteId);
        descendants.forEach(d => {
           const dIndex = notesData.findIndex(n => n.id === d.id);
           if (notesData[dIndex].trashed) notesData[dIndex].trashed = false;
        });

        saveNotes();
        switchTab('active');
        selectNote(notesData[index].id);
        showToast(msg);
      }
    }

    function showModal(title, desc, actionText, actionCallback) {
      document.getElementById('modalTitle').textContent = title;
      document.getElementById('modalDesc').innerHTML = desc;
      const btn = document.getElementById('modalActionBtn');
      btn.textContent = actionText;
      pendingModalAction = actionCallback;
      const newBtn = btn.cloneNode(true);
      btn.parentNode.replaceChild(newBtn, btn);
      newBtn.addEventListener('click', () => {
        if (pendingModalAction) pendingModalAction();
        closeModal();
      });
      document.getElementById('confirmModal').classList.add('show');
    }

    function closeModal(event) {
      if (event && event.target.closest('.modal-content')) return;
      document.getElementById('confirmModal').classList.remove('show');
      pendingModalAction = null;
    }

    function requestPermanentDelete() {
      showModal("完全に削除", "このノート（および子ノート）を完全に削除しますか？<br>この操作は取り消すことができません。", "削除する", () => {
        const descendants = getAllDescendants(activeNoteId).map(n => n.id);
        const toDelete = new Set([activeNoteId, ...descendants]);
        notesData = notesData.filter(n => !toDelete.has(n.id));
        saveNotes();
        activeNoteId = null;
        document.getElementById('editorContainer').style.display = 'none';
        document.getElementById('emptyState').style.display = 'flex';
        showToast("完全に削除しました");
      });
    }

    function requestEmptyTrash() {
      const trashCount = notesData.filter(n => n.trashed).length;
      if (trashCount === 0) return showToast("ゴミ箱は空です");
      showModal("ゴミ箱を空にする", \`ゴミ箱の中にある \${trashCount} 件のノートをすべて完全に削除しますか？<br>この操作は取り消すことができません。\`, "空にする", () => {
        notesData = notesData.filter(n => !n.trashed);
        saveNotes();
        activeNoteId = null;
        document.getElementById('editorContainer').style.display = 'none';
        document.getElementById('emptyState').style.display = 'flex';
        showToast("ゴミ箱を空にしました");
      });
    }

    function togglePinActive() {
      if (!activeNoteId) return;
      const index = notesData.findIndex(n => n.id === activeNoteId);
      if (index === -1 || notesData[index].trashed) return;
      notesData[index].pinned = !notesData[index].pinned;
      renderActionBar(notesData[index], false, false);
      saveNotes();
    }

    function toggleFolder(id, event) {
      event.stopPropagation();
      if (expandedFolders.has(id)) expandedFolders.delete(id);
      else expandedFolders.add(id);
      saveExpandedFolders();
      renderNoteList();
    }

    async function quickCopy(id) {
      const note = notesData.find(n => n.id === id);
      if (!note) return;
      try {
        await navigator.clipboard.writeText(note.body);
        showToast(\`「\${note.title || '無題'}」をコピーしました\`);
      } catch (err) {
        alert("コピーに失敗しました");
      }
    }

    function exportPartial(id) {
      autoSaveCurrentEdits();
      const rootNote = notesData.find(n => n.id === id);
      if (!rootNote) return;

      const descendants = getAllDescendants(id);
      const itemsToExport = [rootNote, ...descendants].filter(n => !n.trashed);

      if (itemsToExport.length === 0) return showToast("書き出すデータがありません");

      const exportData = JSON.parse(JSON.stringify(itemsToExport));
      const exportRoot = exportData.find(n => n.id === id);
      if (exportRoot) exportRoot.parentId = null;

      const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;

      const safeTitle = (rootNote.title || '無題').replace(/[\\\\/:*?"<>|]/g, '_');
      const now = new Date();
      const yyyy = now.getFullYear();
      const mm = String(now.getMonth() + 1).padStart(2, '0');
      const dd = String(now.getDate()).padStart(2, '0');

      a.download = \`share_\${safeTitle}_\${yyyy}\${mm}\${dd}.json\`;
      a.click();
      URL.revokeObjectURL(url);
      showToast(\`「\${rootNote.title || '無題'}」を共有用に書き出しました\`);
    }

    function renderNoteList() {
      const list = document.getElementById('noteList');
      const pinnedList = document.getElementById('pinnedList');
      const pinnedSection = document.getElementById('pinnedSection');
      const treeTitle = document.getElementById('treeTitle');
      const query = document.getElementById('searchInput').value.toLowerCase();
      
      list.innerHTML = '';
      pinnedList.innerHTML = '';

      if (query) {
        pinnedSection.style.display = 'none';
        treeTitle.style.display = 'none';

        let filtered = notesData.filter(n => {
          const shouldShow = currentTab === 'trash' ? n.trashed : !n.trashed;
          if (!shouldShow) return false;
          const inTitle = n.title.toLowerCase().includes(query);
          const inBody = n.body.toLowerCase().includes(query);
          return inTitle || inBody;
        });

        if (filtered.length === 0) {
          list.innerHTML = \`<div style="text-align:center; color:var(--color-text-muted); padding-top:24px; font-size:13px;">一致するノートがありません</div>\`;
          return;
        }

        filtered.forEach(note => {
          const div = document.createElement('div');
          div.className = \`flat-note-item \${note.id === activeNoteId ? 'active' : ''}\`;
          div.onclick = () => selectNote(note.id);
          const pathNames = getBreadcrumbs(note.id).map(c => c.title || '無題').join(' > ');

          div.innerHTML = \`
            <div class="flat-note-content">
              <div class="flat-note-path">\${escapeHTML(pathNames)}</div>
              <div class="flat-note-title">\${note.pinned ? '★ ' : ''}\${escapeHTML(note.title || '無題')}</div>
            </div>
            <div class="note-item-actions">
              <button class="btn-quick btn-quick-primary" onclick="event.stopPropagation(); quickCopy('\${note.id}')">コピー</button>
            </div>
          \`;
          list.appendChild(div);
        });
      } else {
        if (currentTab === 'active') {
          const pinnedNotes = notesData.filter(n => n.pinned && !n.trashed);
          if (pinnedNotes.length > 0) {
            pinnedSection.style.display = 'block';
            treeTitle.style.display = 'block';
            
            pinnedNotes.sort((a, b) => new Date(b.updateDate || b.date) - new Date(a.updateDate || a.date));
            pinnedNotes.forEach(note => {
              const div = document.createElement('div');
              div.className = \`note-item \${note.id === activeNoteId ? 'active' : ''} pinned\`;
              div.onclick = () => selectNote(note.id);
              div.innerHTML = \`
                <div class="note-item-content" style="padding-left: 6px;">
                  <div class="note-title">\${escapeHTML(note.title || '無題')}</div>
                </div>
                <div class="note-item-actions">
                  <button class="btn-quick btn-quick-primary" onclick="event.stopPropagation(); quickCopy('\${note.id}')">コピー</button>
                </div>
              \`;
              pinnedList.appendChild(div);
            });
          } else {
            pinnedSection.style.display = 'none';
            treeTitle.style.display = 'none';
          }
        } else {
          pinnedSection.style.display = 'none';
          treeTitle.style.display = 'none';
        }

        const rootNotes = notesData.filter(n => {
          if (currentTab === 'trash') {
            return n.trashed && (!n.parentId || !notesData.some(p => p.id === n.parentId && p.trashed));
          } else {
            return !n.trashed && (!n.parentId || !notesData.some(p => p.id === n.parentId && !p.trashed));
          }
        });

        rootNotes.sort((a, b) => {
          if (a.pinned !== b.pinned) return a.pinned ? -1 : 1;
          return new Date(b.updateDate || b.date) - new Date(a.updateDate || a.date);
        });

        if (rootNotes.length === 0) {
          list.innerHTML = \`<div style="text-align:center; color:var(--color-text-muted); padding-top:24px; font-size:13px;">\${currentTab === 'trash' ? 'ゴミ箱は空です' : 'ノートがありません'}</div>\`;
          return;
        }
        list.innerHTML = renderTree(rootNotes);
      }
    }

    function renderTree(nodes) {
      let html = '';
      nodes.forEach(note => {
        const children = notesData.filter(n => n.parentId === note.id && n.trashed === note.trashed);
        const hasChildren = children.length > 0;
        const isOpen = expandedFolders.has(note.id);
        children.sort((a, b) => new Date(b.updateDate || b.date) - new Date(a.updateDate || a.date));

        html += \`<div class="tree-node">\`;
        html += \`
          <div class="note-item \${note.id === activeNoteId ? 'active' : ''} \${note.pinned ? 'pinned' : ''}" onclick="selectNote('\${note.id}')">
            <div class="tree-toggle \${hasChildren ? '' : 'empty'} \${isOpen ? 'open' : ''}" onclick="toggleFolder('\${note.id}', event)">▶</div>
            <div class="note-item-content">
              <div class="note-title">\${escapeHTML(note.title || '無題')}</div>
            </div>
            <div class="note-item-actions">
              <button class="btn-quick btn-quick-primary" onclick="event.stopPropagation(); quickCopy('\${note.id}')">コピー</button>
              \${currentTab === 'active' ? \`<button class="btn-quick" onclick="event.stopPropagation(); createNewNote('\${note.id}')">＋ 子ノート</button>\` : ''}
            </div>
          </div>
        \`;
        if (hasChildren) html += \`<div class="tree-children \${isOpen ? 'open' : ''}">\${renderTree(children)}</div>\`;
        html += \`</div>\`;
      });
      return html;
    }

    function exportData() {
      autoSaveCurrentEdits();
      const activeNotes = notesData.filter(n => !n.trashed);
      if (activeNotes.length === 0) return showToast("バックアップするデータがありません");
      const blob = new Blob([JSON.stringify(activeNotes, null, 2)], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      
      const now = new Date();
      const yyyy = now.getFullYear();
      const mm = String(now.getMonth() + 1).padStart(2, '0');
      const dd = String(now.getDate()).padStart(2, '0');
      const hh = String(now.getHours()).padStart(2, '0');
      const min = String(now.getMinutes()).padStart(2, '0');
      
      a.download = \`notes_backup_ALL_\${yyyy}\${mm}\${dd}_\${hh}\${min}.json\`;
      a.click();
      URL.revokeObjectURL(url);
    }

    function importData(event) {
      const file = event.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = function(e) {
        try {
          const imported = JSON.parse(e.target.result);
          const isMerge = confirm("【復元方法の選択】\\n現在のノートを残したまま、データを「追加（マージ）」しますか？\\n\\n・[OK] ＝ 現在のデータに追加します\\n・[キャンセル] ＝ 上書きモードへ進む");
          
          if (isMerge) {
            const existingIds = new Set(notesData.map(n => n.id));
            const newItems = imported.filter(n => !existingIds.has(n.id));
            notesData = [...newItems, ...notesData];
            saveNotes();
            showToast(\`\${newItems.length}件のノートを追加（マージ）しました\`);
          } else {
            const isOverwrite = confirm("⚠️【警告】\\n現在のノートをすべて消去し、データで「完全上書き」しますか？\\n\\n・[OK] ＝ 完全上書きを実行します\\n・[キャンセル] ＝ 復元を中止します");
            if (isOverwrite) {
              notesData = imported;
              saveNotes();
              activeNoteId = null;
              document.getElementById('editorContainer').style.display = 'none';
              document.getElementById('emptyState').style.display = 'flex';
              switchTab('active');
              showToast("データを完全に上書きして復元しました");
            }
          }
        } catch (err) { alert("読み込みエラー: 正しいバックアップファイルを選択してください"); }
      };
      reader.readAsText(file);
      event.target.value = '';
    }

    function escapeHTML(str) { return str.replace(/[&<>'"]/g, tag => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[tag] || tag)); }

    async function copyActiveNote() {
      if (!activeNoteId) return;
      const body = document.getElementById('noteBody').value;
      try { await navigator.clipboard.writeText(body); showToast("コピーしました"); } catch (err) { alert("コピーに失敗しました"); }
    }

    function showToast(msg) {
      const container = document.getElementById('toastContainer');
      const toast = document.createElement('div');
      toast.className = 'toast'; toast.textContent = msg;
      container.appendChild(toast);
      void toast.offsetWidth;
      toast.classList.add('show');
      setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 300); }, 2000);
    }
  </script>
</body>
</html>
`;
