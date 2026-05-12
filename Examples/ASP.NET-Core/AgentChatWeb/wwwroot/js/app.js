/* ── State ──────────────────────────────────────────────────── */
const sessionId = crypto.randomUUID();
let currentFolder = null;

/* ── DOM refs ───────────────────────────────────────────────── */
const folderList        = document.getElementById('folderList');
const fileListContainer = document.getElementById('fileListContainer');
const fileList          = document.getElementById('fileList');
const folderBreadcrumb  = document.getElementById('folderBreadcrumb');
const backBtn           = document.getElementById('backBtn');
const uploadBtn         = document.getElementById('uploadBtn');
const uploadInput       = document.getElementById('uploadInput');
const chatMessages      = document.getElementById('chatMessages');
const chatInput         = document.getElementById('chatInput');
const sendBtn           = document.getElementById('sendBtn');
const clearBtn          = document.getElementById('clearBtn');

/* ══════════════════════════════════════════════════════════════
   FILE EXPLORER
══════════════════════════════════════════════════════════════ */

async function loadFolders() {
  folderList.innerHTML = '<li class="state-msg">Loading…</li>';
  try {
    const res = await fetch('/api/files/folders');
    if (!res.ok) throw new Error(await res.text());
    const folders = await res.json();   // [{name}]

    folderList.innerHTML = '';
    for (const folder of folders) {
      const li = document.createElement('li');
      li.className = 'folder-item';
      li.innerHTML = `<span class="folder-icon">${folderIcon(folder.name)}</span><span class="folder-name">${folder.name}</span>`;
      li.addEventListener('click', () => openFolder(folder.name));
      folderList.appendChild(li);
    }
  } catch (err) {
    folderList.innerHTML = `<li class="state-msg">Error: ${err.message}</li>`;
  }
}

async function openFolder(folderName) {
  currentFolder = folderName;
  folderBreadcrumb.textContent = `📂 ${folderName}`;
  fileList.innerHTML = '<li class="state-msg">Loading…</li>';
  folderList.parentElement.querySelector('.folder-list').style.display = 'none';
  fileListContainer.classList.remove('hidden');

  try {
    const res = await fetch(`/api/files/folders/${encodeURIComponent(folderName)}`);
    if (!res.ok) throw new Error(await res.text());
    const { files } = await res.json();  // [{name, fullPath?, isFolder?, size, extension, modified}]

    fileList.innerHTML = '';
    if (!files.length) {
      fileList.innerHTML = '<li class="state-msg">No files found.</li>';
      return;
    }
    for (const file of files) {
      const li = document.createElement('li');
      li.className = 'file-item';
      const displayName = file.name;
      const downloadPath = file.fullPath;
      
      li.innerHTML = `
        <span class="file-item-left" title="Download ${escHtml(displayName)}">
          <span class="file-icon">${fileIcon(file.extension)}</span>
          <span class="file-name">${escHtml(displayName)}</span>
          <span class="badge">${escHtml(file.extension)}</span>
        </span>
        <span class="file-item-right">
          <span class="file-meta">${escHtml(file.size)}</span>
          <button class="delete-btn" title="Delete ${escHtml(displayName)}" data-file="${escHtml(downloadPath)}">
            <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none"
              stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6l-1 14H6L5 6"></path>
              <path d="M10 11v6"></path>
              <path d="M14 11v6"></path>
              <path d="M9 6V4h6v2"></path>
            </svg>
          </button>
        </span>`;
      li.querySelector('.file-item-left').addEventListener('click', () => downloadFile(folderName, downloadPath));
      li.querySelector('.delete-btn').addEventListener('click', (e) => {
        e.stopPropagation();
        deleteFile(folderName, downloadPath, li);
      });
      fileList.appendChild(li);
    }
  } catch (err) {
    fileList.innerHTML = `<li class="state-msg">Error: ${err.message}</li>`;
  }
}

function downloadFile(folderName, fileName) {
  const url = `/api/files/download/${encodeURIComponent(folderName)}/${encodeURIComponent(fileName)}`;
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.appendChild(a);
  a.click();
  a.remove();
}

async function deleteFile(folderName, fileName, listItem) {
  if (!confirm(`Delete "${fileName}"?`)) return;
  try {
    const res = await fetch(
      `/api/files/delete/${encodeURIComponent(folderName)}/${encodeURIComponent(fileName)}`,
      { method: 'DELETE' }
    );
    if (!res.ok) throw new Error(await res.text());
    listItem.remove();
    if (!fileList.querySelector('.file-item')) {
      fileList.innerHTML = '<li class="state-msg">No files found.</li>';
    }
  } catch (err) {
    alert(`Failed to delete "${fileName}":\n${err.message}`);
  }
}

/* ── Upload ──────────────────────────────────────────────────── */
uploadBtn.addEventListener('click', () => uploadInput.click());

uploadInput.addEventListener('change', async () => {
  const files = uploadInput.files;
  if (!files || files.length === 0) return;

  const targetFolder = 'Input';
  const formData = new FormData();
  for (const file of files) formData.append('files', file);

  uploadBtn.disabled = true;
  uploadBtn.textContent = 'Uploading…';

  try {
    const res = await fetch('/api/files/upload', { method: 'POST', body: formData });
    if (!res.ok) throw new Error(await res.text());
    const data = await res.json();
    alert(`✅ Uploaded ${data.files.length} file(s) to ${targetFolder}.`);
    // Refresh the Input folder if it's currently open
    if (currentFolder === targetFolder) {
        await openFolder(targetFolder);
    }
  } catch (err) {
    alert(`Upload failed:\n${err.message}`);
  } finally {
    uploadBtn.disabled = false;
    uploadBtn.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none"
      stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
      <polyline points="16 16 12 12 8 16"></polyline>
      <line x1="12" y1="12" x2="12" y2="21"></line>
      <path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"></path>
    </svg> Upload`;
    uploadInput.value = '';
  }
});

backBtn.addEventListener('click', () => {
  fileListContainer.classList.add('hidden');
  folderList.parentElement.querySelector('.folder-list').style.display = '';
  currentFolder = null;
});

/* ══════════════════════════════════════════════════════════════
   CHAT BOT
══════════════════════════════════════════════════════════════ */

sendBtn.addEventListener('click', sendMessage);
chatInput.addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); sendMessage(); }
});
clearBtn.addEventListener('click', async () => {
  await fetch(`/api/chat/session/${sessionId}`, { method: 'DELETE' });
  chatMessages.innerHTML = `
    <div class="chat-welcome">
      <span class="robot-icon">🤖</span>
      <p>Hello! I can help you work with your documents.<br/>Ask me anything about Word, Excel, PDF, or PowerPoint files.</p>
    </div>`;
});

async function sendMessage() {
  const text = chatInput.value.trim();
  if (!text || sendBtn.disabled) return;

  // Clear welcome
  const welcome = chatMessages.querySelector('.chat-welcome');
  if (welcome) welcome.remove();

  appendMessage('user', text);
  chatInput.value = '';
  sendBtn.disabled = true;

  const botBubble = appendMessage('bot', '', true);

  try {
    const res = await fetch('/api/chat/send', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId, message: text })
    });

    if (!res.ok) {
      const err = await res.text();
      appendToBubble(botBubble, `❌ Error: ${err}`, false);
      setBubbleTyping(botBubble, false);
      return;
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = decoder.decode(value, { stream: true });
      appendToBubble(botBubble, chunk, true);
    }

    setBubbleTyping(botBubble, false);
  } catch (err) {
    appendToBubble(botBubble, `❌ ${err.message}`, false);
    setBubbleTyping(botBubble, false);
  } finally {
    sendBtn.disabled = false;
    chatInput.focus();
  }
}

function appendMessage(role, text, typing = false) {
  const div = document.createElement('div');
  div.className = `msg ${role}${typing ? ' typing' : ''}`;
  div.innerHTML = `
    <span class="msg-label">${role === 'user' ? 'You' : 'AI Assistant'}</span>
    <div class="msg-bubble">${escHtml(text)}</div>`;
  chatMessages.appendChild(div);
  chatMessages.scrollTop = chatMessages.scrollHeight;
  return div;
}

/**
 * Append a new chunk of text to the bot bubble, rendering inline markdown.
 * Each chunk is wrapped in a <span> so previous chunks are never re-rendered.
 */
function appendToBubble(div, chunk, typing) {
  div.className = `msg bot${typing ? ' typing' : ''}`;
  const bubble = div.querySelector('.msg-bubble');
  const span = document.createElement('span');
  span.innerHTML = renderMarkdown(chunk);
  bubble.appendChild(span);
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

/** Toggle the typing indicator without touching bubble content. */
function setBubbleTyping(div, typing) {
  div.className = `msg bot${typing ? ' typing' : ''}`;
}

/**
 * Minimal inline-markdown renderer.
 * Handles: **bold**, *italic*, `code`, and newlines → <br>.
 */
function renderMarkdown(text) {
  return escHtml(text)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\n/g, '<br>');
}

/* ══════════════════════════════════════════════════════════════
   HELPERS
══════════════════════════════════════════════════════════════ */

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function folderIcon(name) {
  // Standard folder icon for all folders
  return `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="#f5a623"
      stroke="#e09000" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round">
    <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path>
  </svg>`;
}

function fileIcon(ext) {
  switch ((ext || '').toUpperCase()) {
    case 'PDF':  return '📄';
    case 'DOCX':
    case 'DOC':  return '📝';
    case 'XLSX':
    case 'XLS':  return '📊';
    case 'PPTX':
    case 'PPT':  return '📽️';
    case 'RTF':
    case 'HTML': return '📝';
    case 'XLSM':
    case 'CSV':  return '📊';
    case 'JSON': return '🔧';
    case 'MD':   return '📋';
    case 'PNG':
    case 'JPG':
    case 'JPEG':
    case 'GIF':
    case 'BMP':
    case 'TIFF':
    case 'WEBP': return '🖼️';
    default:     return '📄';
  }
}

/* ── Init ───────────────────────────────────────────────────── */
loadFolders();
