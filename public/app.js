/**
 * FileShift Web Application
 * Client-Side Controller & Offline-Ready Converter Interface
 */

(function () {
  'use strict';

  // State Management
  const state = {
    selectedFile: null,
    sourceFormat: 'xlsx',
    targetFormat: 'csv',
    convertedBlob: null,
    convertedFilename: '',
    history: [],
    stats: {
      totalConverted: 0,
      totalBytes: 0,
      modeCounts: {}
    },
    isOnlineServer: false
  };

  // DOM Elements
  const elements = {
    themeToggle: document.getElementById('theme-toggle'),
    engineStatusText: document.getElementById('engine-status-text'),
    engineDot: document.querySelector('.status-dot'),
    
    // Formats & Presets
    badgeSourceFmt: document.getElementById('badge-source-fmt'),
    badgeTargetFmt: document.getElementById('badge-target-fmt'),
    btnSwapFormats: document.getElementById('btn-swap-formats'),
    presetItems: document.querySelectorAll('.preset-item'),
    
    // Dropzone
    dropzone: document.getElementById('dropzone'),
    fileInput: document.getElementById('file-input'),
    dropzoneEmptyState: document.getElementById('dropzone-empty-state'),
    dropzoneActiveState: document.getElementById('dropzone-active-state'),
    activeFileExt: document.getElementById('active-file-ext'),
    activeFileName: document.getElementById('active-file-name'),
    activeFileSize: document.getElementById('active-file-size'),
    btnRemoveFile: document.getElementById('btn-remove-file'),
    btnSelectFile: document.getElementById('btn-select-file'),
    
    // Options
    selectDelimiter: document.getElementById('select-delimiter'),
    selectSheet: document.getElementById('select-sheet'),
    optDelimiterGroup: document.getElementById('opt-delimiter-group'),
    optSheetGroup: document.getElementById('opt-sheet-group'),
    
    // Actions & States
    btnConvert: document.getElementById('btn-convert'),
    stateReady: document.getElementById('state-ready'),
    stateProcessing: document.getElementById('state-processing'),
    progressStepText: document.getElementById('progress-step-text'),
    progressPercentText: document.getElementById('progress-percent-text'),
    progressFill: document.getElementById('progress-fill'),
    stateSuccess: document.getElementById('state-success'),
    successMetaInfo: document.getElementById('success-meta-info'),
    btnDownloadResult: document.getElementById('btn-download-result'),
    btnResetConverter: document.getElementById('btn-reset-converter'),
    stateError: document.getElementById('state-error'),
    errorMessageText: document.getElementById('error-message-text'),
    btnDismissError: document.getElementById('btn-dismiss-error'),
    
    // History & Stats
    historyEmptyView: document.getElementById('history-empty-view'),
    historyItemsList: document.getElementById('history-items-list'),
    btnClearHistory: document.getElementById('btn-clear-history'),
    statTotalConverted: document.getElementById('stat-total-converted'),
    statTotalBytes: document.getElementById('stat-total-bytes'),
    statActiveMode: document.getElementById('stat-active-mode')
  };

  // Helper Functions
  function formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
  }

  function getFileExtension(filename) {
    if (!filename || !filename.includes('.')) return '';
    return filename.split('.').pop().toLowerCase();
  }

  // 1. Theme Toggle
  function initTheme() {
    const savedTheme = localStorage.getItem('fileshift_theme');
    const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches;
    const theme = savedTheme || (prefersDark ? 'dark' : 'light');
    document.documentElement.setAttribute('data-theme', theme);
  }

  elements.themeToggle.addEventListener('click', () => {
    const currentTheme = document.documentElement.getAttribute('data-theme') || 'dark';
    const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', newTheme);
    localStorage.setItem('fileshift_theme', newTheme);
  });

  // 2. Server Health Check
  async function checkServerHealth() {
    try {
      const res = await fetch('/api/health', { method: 'GET' });
      if (res.ok) {
        state.isOnlineServer = true;
        elements.engineDot.className = 'status-dot online';
        elements.engineStatusText.textContent = 'Mesin API Aktif';
        return;
      }
    } catch (e) {
      // Offline / client-only mode
    }
    state.isOnlineServer = false;
    elements.engineDot.className = 'status-dot online';
    elements.engineStatusText.textContent = 'Mode Lokal Mandiri';
  }

  // 3. Format & Presets Configuration
  function setConversionRoute(fromFmt, toFmt) {
    state.sourceFormat = fromFmt.toLowerCase();
    state.targetFormat = toFmt.toLowerCase();

    elements.badgeSourceFmt.textContent = state.sourceFormat.toUpperCase();
    elements.badgeTargetFmt.textContent = state.targetFormat.toUpperCase();

    // Update preset pills active state
    elements.presetItems.forEach(item => {
      const itemFrom = item.getAttribute('data-from');
      const itemTo = item.getAttribute('data-to');
      const isActive = itemFrom === state.sourceFormat && itemTo === state.targetFormat;
      item.classList.toggle('active', isActive);
      item.setAttribute('aria-checked', isActive ? 'true' : 'false');
    });

    // Update accept attribute on file picker
    elements.fileInput.setAttribute('accept', '.' + state.sourceFormat);

    // Toggle options visibility based on format
    const isSpreadsheet = (state.sourceFormat === 'csv' || state.targetFormat === 'csv' ||
                           state.sourceFormat === 'xlsx' || state.targetFormat === 'xlsx');
    elements.optDelimiterGroup.style.display = isSpreadsheet ? 'flex' : 'none';
    elements.optSheetGroup.style.display = (state.sourceFormat === 'xlsx') ? 'flex' : 'none';
  }

  elements.presetItems.forEach(item => {
    item.addEventListener('click', () => {
      const fromFmt = item.getAttribute('data-from');
      const toFmt = item.getAttribute('data-to');
      setConversionRoute(fromFmt, toFmt);
      if (state.selectedFile) {
        const ext = getFileExtension(state.selectedFile.name);
        if (ext !== state.sourceFormat) {
          clearSelectedFile();
        }
      }
    });
  });

  elements.btnSwapFormats.addEventListener('click', () => {
    const curFrom = state.sourceFormat;
    const curTo = state.targetFormat;
    setConversionRoute(curTo, curFrom);
    if (state.selectedFile) {
      const ext = getFileExtension(state.selectedFile.name);
      if (ext !== state.sourceFormat) {
        clearSelectedFile();
      }
    }
  });

  // 4. File Handling & Dropzone
  function handleFileSelected(file) {
    if (!file) return;

    const ext = getFileExtension(file.name);
    state.selectedFile = file;

    // Auto-select route if matches known format
    if (ext === 'xlsx') setConversionRoute('xlsx', 'csv');
    else if (ext === 'csv') setConversionRoute('csv', 'xlsx');
    else if (ext === 'pdf') setConversionRoute('pdf', 'docx');
    else if (ext === 'docx') setConversionRoute('docx', 'pdf');

    // Update Dropzone View
    elements.dropzoneEmptyState.style.display = 'none';
    elements.dropzoneActiveState.style.display = 'block';
    elements.activeFileExt.textContent = ext.toUpperCase();
    elements.activeFileName.textContent = file.name;
    elements.activeFileSize.textContent = formatBytes(file.size);

    // Enable Convert Button
    elements.btnConvert.disabled = false;
    showState('ready');

  }

  function clearSelectedFile() {
    state.selectedFile = null;
    elements.fileInput.value = '';
    elements.dropzoneEmptyState.style.display = 'block';
    elements.dropzoneActiveState.style.display = 'none';
    elements.btnConvert.disabled = true;
    showState('ready');
  }

  elements.dropzone.addEventListener('click', (e) => {
    if (e.target.closest('#btn-remove-file')) return;
    elements.fileInput.click();
  });

  elements.btnSelectFile.addEventListener('click', () => {
    elements.fileInput.click();
  });

  elements.dropzone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      elements.fileInput.click();
    }
  });

  elements.fileInput.addEventListener('change', (e) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      handleFileSelected(files[0]);
    }
  });

  elements.btnRemoveFile.addEventListener('click', (e) => {
    e.stopPropagation();
    clearSelectedFile();
  });

  // Drag and Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    elements.dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      elements.dropzone.classList.add('drag-over');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    elements.dropzone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      elements.dropzone.classList.remove('drag-over');
    });
  });

  elements.dropzone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt && dt.files && dt.files.length > 0) {
      handleFileSelected(dt.files[0]);
    }
  });

  // 6. UI State Management (ready, processing, success, error)
  function showState(viewName) {
    elements.stateReady.style.display = (viewName === 'ready') ? 'flex' : 'none';
    elements.stateProcessing.style.display = (viewName === 'processing') ? 'block' : 'none';
    elements.stateSuccess.style.display = (viewName === 'success') ? 'block' : 'none';
    elements.stateError.style.display = (viewName === 'error') ? 'flex' : 'none';
  }

  function updateProgress(stepText, percent) {
    elements.progressStepText.textContent = stepText;
    elements.progressPercentText.textContent = percent + '%';
    elements.progressFill.style.width = percent + '%';
  }

  elements.btnDismissError.addEventListener('click', () => {
    showState('ready');
  });

  elements.btnResetConverter.addEventListener('click', () => {
    clearSelectedFile();
  });

  // 7. Conversion Workflow
  elements.btnConvert.addEventListener('click', async () => {
    if (!state.selectedFile) return;

    showState('processing');
    updateProgress('Membaca file input...', 20);

    const file = state.selectedFile;
    const fromFmt = state.sourceFormat;
    const toFmt = state.targetFormat;
    const delimiter = elements.selectDelimiter.value === 'auto' ? '' : elements.selectDelimiter.value;
    const sheetIdx = elements.selectSheet.value || '0';

    try {
      updateProgress('Menghubungkan ke mesin konversi...', 45);

      const formData = new FormData();
      formData.append('file', file);
      formData.append('from', fromFmt);
      formData.append('to', toFmt);
      if (delimiter) formData.append('delimiter', delimiter);
      formData.append('sheet_index', sheetIdx);

      const response = await fetch('/api/convert', {
        method: 'POST',
        body: formData
      });

      if (!response.ok) {
        let errDesc = 'Gagal melakukan konversi berkas.';
        try {
          const errJson = await response.json();
          if (errJson && errJson.error) errDesc = errJson.error;
        } catch (_) {}
        throw new Error(errDesc);
      }

      updateProgress('Menyusun dokumen keluaran...', 85);

      const blob = await response.blob();
      const contentDisp = response.headers.get('Content-Disposition') || '';
      let outFilename = `converted_${file.name.replace(/\.[^/.]+$/, '')}.${toFmt}`;
      if (contentDisp.includes('filename=')) {
        const match = contentDisp.match(/filename="?([^";]+)"?/);
        if (match && match[1]) outFilename = match[1];
      }

      state.convertedBlob = blob;
      state.convertedFilename = outFilename;

      updateProgress('Selesai!', 100);

      // Success State Update
      elements.successMetaInfo.textContent = `${outFilename} (${formatBytes(blob.size)}) siap disimpan ke perangkat Anda.`;
      showState('success');

      // Record to history & stats
      addToHistory({
        sourceName: file.name,
        targetName: outFilename,
        fromFmt: fromFmt.toUpperCase(),
        toFmt: toFmt.toUpperCase(),
        sizeBytes: blob.size,
        blob: blob,
        timestamp: new Date().toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' })
      });


    } catch (err) {
      elements.errorMessageText.textContent = err.message || 'Terjadi kesalahan sistem saat pemrosesan.';
      showState('error');
    }
  });

  // 8. Download Result
  elements.btnDownloadResult.addEventListener('click', () => {
    if (!state.convertedBlob || !state.convertedFilename) return;
    downloadBlob(state.convertedBlob, state.convertedFilename);
  });

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // 9. History Management
  function loadHistory() {
    try {
      const raw = localStorage.getItem('fileshift_history');
      if (raw) {
        state.history = JSON.parse(raw);
        renderHistory();
      }
    } catch (_) {}
  }

  function saveHistory() {
    try {
      // Don't serialize blobs to localStorage
      const serializable = state.history.map(item => ({
        sourceName: item.sourceName,
        targetName: item.targetName,
        fromFmt: item.fromFmt,
        toFmt: item.toFmt,
        sizeBytes: item.sizeBytes,
        timestamp: item.timestamp
      }));
      localStorage.setItem('fileshift_history', JSON.stringify(serializable.slice(0, 20)));
    } catch (_) {}
  }

  function addToHistory(item) {
    state.history.unshift(item);
    renderHistory();
    saveHistory();
    updateStats(item);
  }

  function renderHistory() {
    if (!state.history || state.history.length === 0) {
      elements.historyEmptyView.style.display = 'flex';
      elements.historyItemsList.style.display = 'none';
      return;
    }

    elements.historyEmptyView.style.display = 'none';
    elements.historyItemsList.style.display = 'flex';

    elements.historyItemsList.innerHTML = state.history.slice(0, 10).map((item, idx) => {
      return `
        <li class="history-item">
          <div class="history-info">
            <span class="history-filename">${escapeHtml(item.targetName)}</span>
            <span class="history-meta">${escapeHtml(item.fromFmt)} ➔ ${escapeHtml(item.toFmt)} • ${formatBytes(item.sizeBytes)} • ${item.timestamp}</span>
          </div>
          ${item.blob ? `<button class="btn-download-sm" data-idx="${idx}">Unduh</button>` : ''}
        </li>
      `;
    }).join('');

    // Attach click listeners to download buttons
    elements.historyItemsList.querySelectorAll('.btn-download-sm').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.getAttribute('data-idx'), 10);
        const histItem = state.history[idx];
        if (histItem && histItem.blob) {
          downloadBlob(histItem.blob, histItem.targetName);
        }
      });
    });
  }

  elements.btnClearHistory.addEventListener('click', () => {
    state.history = [];
    localStorage.removeItem('fileshift_history');
    renderHistory();
  });

  // 10. Stats Tracking
  function updateStats(lastItem) {
    state.stats.totalConverted += 1;
    state.stats.totalBytes += lastItem.sizeBytes;
    const modeKey = `${lastItem.fromFmt} ➔ ${lastItem.toFmt}`;
    state.stats.modeCounts[modeKey] = (state.stats.modeCounts[modeKey] || 0) + 1;

    // Find most frequent mode
    let maxCount = 0;
    let mostFrequent = modeKey;
    for (const [k, v] of Object.entries(state.stats.modeCounts)) {
      if (v > maxCount) {
        maxCount = v;
        mostFrequent = k;
      }
    }

    elements.statTotalConverted.textContent = state.stats.totalConverted.toString();
    elements.statTotalBytes.textContent = formatBytes(state.stats.totalBytes);
    elements.statActiveMode.textContent = mostFrequent;
  }

  // 11. Keyboard Shortcuts
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (state.selectedFile) {
        clearSelectedFile();
      }
    }
  });

  // Initialize
  initTheme();
  setConversionRoute('xlsx', 'csv');
  checkServerHealth();
  loadHistory();

})();
