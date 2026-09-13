(function () {
  const API = ''; // same origin: static frontend + FastAPI

  // ---- Theme ----
  (function () {
    const t = document.querySelector('[data-theme-toggle]'), r = document.documentElement;
    const stored = localStorage.getItem('tf-theme');
    let d = stored === 'dark' || stored === 'light'
      ? stored
      : (matchMedia('(prefers-color-scheme:dark)').matches ? 'dark' : 'light');
    r.setAttribute('data-theme', d);
    t.addEventListener('click', () => {
      d = d === 'dark' ? 'light' : 'dark';
      r.setAttribute('data-theme', d);
      localStorage.setItem('tf-theme', d);
    });
  })();

  // ---- State ----
  let inputFiles = []; // File[]
  let resultBlob = null;
  let resultObjectUrl = null;

  // ---- Elements ----
  const el = {
    tplName: document.getElementById('tplName'),
    tplMeta: document.getElementById('tplMeta'),
    tplDownloadBtn: document.getElementById('tplDownloadBtn'),
    tplReplaceBtn: document.getElementById('tplReplaceBtn'),
    tplFileInput: document.getElementById('tplFileInput'),
    dropzone: document.getElementById('dropzone'),
    inputFilesEl: document.getElementById('inputFiles'),
    fileList: document.getElementById('fileList'),
    countryInput: document.getElementById('countryInput'),
    unitInput: document.getElementById('unitInput'),
    downloadBtn: document.getElementById('downloadBtn'),
    errorBox: document.getElementById('errorBox'),
    // элементы протокола ошибок и прогресса
    errorLog: document.getElementById('errorLog'),
    errorLogList: document.getElementById('errorLogList'),
    errorLogSummary: document.getElementById('errorLogSummary'),
    progressBar: document.getElementById('progressBar'),
    progressFill: document.getElementById('progressFill'),
    progressLabel: document.getElementById('progressLabel'),
    resultBox: document.getElementById('resultBox'),
    resultName: document.getElementById('resultName'),
    resultMeta: document.getElementById('resultMeta'),
    resultDownloadBtn: document.getElementById('resultDownloadBtn'),
    resultDeleteBtn: document.getElementById('resultDeleteBtn'),
  };

  function fmtBytes(n) {
    if (n < 1024) return n + ' Б';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' КБ';
    return (n / 1024 / 1024).toFixed(1) + ' МБ';
  }

  function fmtDate(iso) {
    if (!iso) return '';
    try {
      const d = new Date(iso);
      return d.toLocaleString('ru-RU', {
        day: '2-digit', month: '2-digit', year: 'numeric',
        hour: '2-digit', minute: '2-digit'
      });
    } catch { return iso; }
  }

  function showError(msg) {
    el.errorBox.textContent = msg;
    el.errorBox.style.display = 'block';
  }

  function clearError() {
    el.errorBox.style.display = 'none';
    el.errorBox.textContent = '';
    if (el.errorLog) {
      el.errorLog.style.display = 'none';
      el.errorLogList.innerHTML = '';
      el.errorLogSummary.textContent = '';
    }
  }

  // ---- Template info ----
  function sleep(ms) { return new Promise((r) => setTimeout(r, ms)); }

  async function fetchWithRetry(input, initOrFactory, attempts = 4, delayMs = 1500) {
    let lastErr;
    for (let i = 0; i < attempts; i++) {
      try {
        const init = typeof initOrFactory === 'function' ? initOrFactory() : initOrFactory;
        return await fetch(input, init);
      } catch (e) {
        lastErr = e;
        if (i < attempts - 1) await sleep(delayMs);
      }
    }
    throw lastErr;
  }

  async function loadTemplateInfo() {
    el.tplName.textContent = 'Загрузка…';
    el.tplMeta.textContent = '';
    try {
      const res = await fetchWithRetry(`${API}/api/template`);
      const data = await res.json();
      if (data.exists) {
        el.tplName.textContent = data.filename || 'template.xlsx';
        el.tplMeta.textContent = 'обновлён: ' + fmtDate(data.updated_at);
        el.tplDownloadBtn.disabled = false;
      } else {
        el.tplName.textContent = 'Шаблон не загружен';
        el.tplMeta.textContent = 'загрузите эталонный файл';
      }
    } catch (e) {
      el.tplName.textContent = 'Не удалось получить данные о шаблоне';
      const retryLink = document.createElement('a');
      retryLink.href = '#';
      retryLink.textContent = 'Повторить';
      retryLink.addEventListener('click', (ev) => { ev.preventDefault(); loadTemplateInfo(); });
      el.tplMeta.textContent = '';
      el.tplMeta.appendChild(retryLink);
    }
  }

  el.tplDownloadBtn.addEventListener('click', () => {
    window.open(`${API}/api/template/download`, '_blank');
  });

  el.tplReplaceBtn.addEventListener('click', () => el.tplFileInput.click());
  el.tplFileInput.addEventListener('change', async () => {
    const file = el.tplFileInput.files[0];
    if (!file) return;
    const fd = new FormData();
    fd.append('file', file);
    el.tplReplaceBtn.disabled = true;
    try {
      const res = await fetch(`${API}/api/template`, { method: 'POST', body: fd });
      if (!res.ok) throw new Error('Не удалось загрузить шаблон');
      await loadTemplateInfo();
      clearError();
    } catch (e) {
      showError('Ошибка при замене шаблона: ' + e.message);
    } finally {
      el.tplReplaceBtn.disabled = false;
      el.tplFileInput.value = '';
    }
  });

  // ---- Input files ----
  const FILE_ICON_SVG = '';
  const REMOVE_ICON_SVG = '';

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function showResult(blob, filename) {
    if (resultObjectUrl) URL.revokeObjectURL(resultObjectUrl);
    resultBlob = blob;
    resultObjectUrl = URL.createObjectURL(blob);
    el.resultName.textContent = filename;
    el.resultMeta.textContent = 'можно удалить и загрузить новую партию без перезапуска';
    el.resultBox.hidden = false;
  }

  function clearResult() {
    if (resultObjectUrl) URL.revokeObjectURL(resultObjectUrl);
    resultBlob = null;
    resultObjectUrl = null;
    if (el.resultBox) el.resultBox.hidden = true;
  }

  function startNewBatch() {
    inputFiles = [];
    renderFileList();
    clearResult();
    clearError();
    hideProgress();
    if (el.inputFilesEl) el.inputFilesEl.value = '';
  }

  function renderFileList() {
    el.fileList.innerHTML = '';
    inputFiles.forEach((f, i) => {
      const li = document.createElement('li');
      li.className = 'file-item';

      const iconWrap = document.createElement('span');
      iconWrap.innerHTML = FILE_ICON_SVG;
      if (iconWrap.firstElementChild) li.appendChild(iconWrap.firstElementChild);

      const nameSpan = document.createElement('span');
      nameSpan.className = 'fname';
      nameSpan.textContent = f.name;
      li.appendChild(nameSpan);

      const sizeSpan = document.createElement('span');
      sizeSpan.className = 'fsize';
      sizeSpan.textContent = fmtBytes(f.size);
      li.appendChild(sizeSpan);

      const removeBtn = document.createElement('button');
      removeBtn.className = 'remove';
      removeBtn.type = 'button';
      removeBtn.setAttribute('aria-label', 'Удалить файл');
      removeBtn.dataset.idx = String(i);
      const removeIconWrap = document.createElement('span');
      removeIconWrap.innerHTML = REMOVE_ICON_SVG;
      if (removeIconWrap.firstElementChild) {
        removeBtn.appendChild(removeIconWrap.firstElementChild);
      } else {
        removeBtn.textContent = '×';
      }
      li.appendChild(removeBtn);

      el.fileList.appendChild(li);
    });

    el.fileList.querySelectorAll('button.remove').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = Number(btn.dataset.idx);
        inputFiles.splice(idx, 1);
        renderFileList();
      });
    });
  }

  function addFiles(fileListObj) {
    const arr = Array.from(fileListObj).filter((f) => /\.xlsx$/i.test(f.name));
    inputFiles = inputFiles.concat(arr);
    renderFileList();
  }

  el.dropzone.addEventListener('click', () => el.inputFilesEl.click());
  el.dropzone.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') el.inputFilesEl.click();
  });
  el.inputFilesEl.addEventListener('change', () => {
    addFiles(el.inputFilesEl.files);
    el.inputFilesEl.value = '';
  });

  ['dragenter', 'dragover'].forEach((evt) =>
    el.dropzone.addEventListener(evt, (e) => { e.preventDefault(); el.dropzone.classList.add('dragover'); })
  );
  ['dragleave', 'drop'].forEach((evt) =>
    el.dropzone.addEventListener(evt, (e) => { e.preventDefault(); el.dropzone.classList.remove('dragover'); })
  );
  el.dropzone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
  });

  // ---- Build form data ----
  function buildFormData() {
    const fd = new FormData();
    inputFiles.forEach((f) => fd.append('inputs', f));
    fd.append('country', el.countryInput.value.trim() || 'CN');
    fd.append('unit', el.unitInput.value.trim() || 'шт');
    return fd;
  }

  function setBusy(btn, busy, labelWhileBusy) {
    if (busy) {
      btn.dataset.originalHtml = btn.innerHTML;
      btn.innerHTML = ` ${labelWhileBusy}`;
      btn.disabled = true;
    } else {
      btn.innerHTML = btn.dataset.originalHtml || btn.innerHTML;
      btn.disabled = false;
    }
  }

  // ---- Протокол ошибок и прогресс ----
  function setProgress(percent) {
    if (!el.progressBar) return;
    const p = Math.max(0, Math.min(100, percent));
    el.progressBar.style.display = 'block';
    el.progressFill.style.width = p + '%';
    el.progressLabel.textContent = p + '%';
  }

  function hideProgress() {
    if (!el.progressBar) return;
    el.progressBar.style.display = 'none';
  }

  // log: { errors?: string[], warnings?: string[], unknown?: string[] }
  function showErrorLog(log) {
    if (!el.errorLog) return;

    el.errorLogList.innerHTML = '';

    const addEntry = (type, message) => {
      const li = document.createElement('li');
      li.className = 'log-' + type;
      li.textContent = message;
      el.errorLogList.appendChild(li);
    };

    (log.errors || []).forEach((m) => addEntry('error', m));
    (log.warnings || []).forEach((m) => addEntry('warning', m));
    (log.unknown || []).forEach((m) => addEntry('unknown', m));

    if ((log.unknown || []).length) {
      el.errorLogSummary.textContent = 'Обнаружены неизвестные элементы в загружаемом файле.';
    } else if ((log.errors || []).length || (log.warnings || []).length) {
      el.errorLogSummary.textContent = 'Обнаружены ошибки/предупреждения, проверьте протокол.';
    } else {
      el.errorLogSummary.textContent = 'Ошибок не обнаружено, файл обработан корректно.';
    }

    el.errorLog.style.display = 'block';
  }

  // ---- Только выгрузка файла, без предпросмотра ----
  el.downloadBtn.addEventListener('click', async () => {
    clearError();

    if (!inputFiles.length) {
      showError('Добавьте хотя бы один входной файл.');
      return;
    }

    setBusy(el.downloadBtn, true, 'Формируем файл…');
    setProgress(5);

    let progressTimer = null;

    try {
      // Псевдо-прогресс до ~90%
      progressTimer = setInterval(() => {
        const current = parseInt(el.progressLabel.textContent, 10) || 0;
        if (current < 90) setProgress(current + 5);
      }, 500);

      const res = await fetch(`${API}/api/process`, { method: 'POST', body: buildFormData() });

      if (progressTimer) clearInterval(progressTimer);
      setProgress(100);

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        throw new Error(data.error || 'Ошибка формирования файла');
      }

      showErrorLog({ errors: [], warnings: [], unknown: [] });

      const blob = await res.blob();
      const filename = 'result.xlsx';
      showResult(blob, filename);
      triggerDownload(blob, filename);
    } catch (e) {
      if (progressTimer) clearInterval(progressTimer);
      showError('Не удалось скачать файл: ' + e.message);
    } finally {
      hideProgress();
      setBusy(el.downloadBtn, false);
    }
  });

  if (el.resultDownloadBtn) {
    el.resultDownloadBtn.addEventListener('click', () => {
      if (!resultBlob) return;
      triggerDownload(resultBlob, el.resultName.textContent || 'result.xlsx');
    });
  }

  if (el.resultDeleteBtn) {
    el.resultDeleteBtn.addEventListener('click', () => {
      startNewBatch();
    });
  }

  // ---- Инициализация ----
  loadTemplateInfo();
})();
