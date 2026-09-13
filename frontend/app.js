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

  // ---- Табы ----
  (function () {
    const tabs = document.querySelectorAll('.tab');
    const contents = document.querySelectorAll('.tab-content');
    tabs.forEach(function (tab) {
      tab.addEventListener('click', function () {
        var target = tab.dataset.tab;
        tabs.forEach(function (t) { t.classList.remove('active'); });
        contents.forEach(function (c) { c.classList.remove('active'); });
        tab.classList.add('active');
        var targetEl = document.getElementById('tab-' + target);
        if (targetEl) targetEl.classList.add('active');
      });
    });
  })();

  // ---- SVH Search ----
  var svhEl = {
    address: document.getElementById('svhAddress'),
    name: document.getElementById('svhName'),
    customs: document.getElementById('svhCustoms'),
    license: document.getElementById('svhLicense'),
    transport: document.getElementById('svhTransport'),
    searchBtn: document.getElementById('svhSearchBtn'),
    clearBtn: document.getElementById('svhClearBtn'),
    errorBox: document.getElementById('svhErrorBox'),
    resultsCard: document.getElementById('svhResultsCard'),
    resultsCount: document.getElementById('svhResultsCount'),
    resultsSource: document.getElementById('svhResultsSource'),
    loading: document.getElementById('svhLoading'),
    resultsList: document.getElementById('svhResultsList'),
    pagination: document.getElementById('svhPagination'),
  };

  var svhState = {
    allResults: [],
    currentPage: 1,
    perPage: 20,
    total: 0,
  };

  function showSvhError(msg) {
    svhEl.errorBox.textContent = msg;
    svhEl.errorBox.style.display = 'block';
  }
  function clearSvhError() {
    svhEl.errorBox.style.display = 'none';
    svhEl.errorBox.textContent = '';
  }

  function fmtDateShort(iso) {
    if (!iso) return '';
    try {
      var d = new Date(iso);
      return d.toLocaleDateString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric' });
    } catch (e) { return iso; }
  }

  function escapeHtml(str) {
    if (!str) return '';
    var div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
  }

  function renderSvhCard(card) {
    var div = document.createElement('div');
    div.className = 'svh-card';

    var typeHtml = '';
    if (card.type) {
      typeHtml = '<span class="svh-card-type">' + escapeHtml(card.type) + '</span>';
    }

    var fieldsHtml = '';
    var fields = [
      ['Адрес', card.address],
      ['Лицензия', card.license],
      ['Таможня', card.customs],
      ['Транспорт', card.transport],
      ['Телефон', card.phone],
      ['Email', card.email],
      ['ИНН', card.inn],
    ];
    fields.forEach(function (f) {
      if (f[1]) {
        var val = f[1];
        if (f[0] === 'Email') {
          val = '<a href="mailto:' + escapeHtml(f[1]) + '">' + escapeHtml(f[1]) + '</a>';
        } else if (f[0] === 'Телефон') {
          val = '<a href="tel:' + escapeHtml(f[1]) + '">' + escapeHtml(f[1]) + '</a>';
        } else {
          val = escapeHtml(f[1]);
        }
        fieldsHtml += '<div class="svh-card-field"><span class="svh-card-label">' + escapeHtml(f[0]) + ':</span><span class="svh-card-value">' + val + '</span></div>';
      }
    });

    var urlHtml = '';
    if (card.url) {
      urlHtml = '<a href="' + escapeHtml(card.url) + '" target="_blank" rel="noopener">Страница на alta.ru →</a>';
    }

    div.innerHTML =
      '<div class="svh-card-header">' +
        '<div class="svh-card-title">' + escapeHtml(card.name) + '</div>' +
        typeHtml +
      '</div>' +
      '<div class="svh-card-grid">' + fieldsHtml + '</div>' +
      (urlHtml ? '<div class="svh-card-actions">' + urlHtml + '</div>' : '');

    return div;
  }

  function renderSvhResults() {
    svhEl.resultsList.innerHTML = '';
    var start = (svhState.currentPage - 1) * svhState.perPage;
    var end = start + svhState.perPage;
    var pageItems = svhState.allResults.slice(start, end);

    pageItems.forEach(function (card) {
      svhEl.resultsList.appendChild(renderSvhCard(card));
    });

    svhEl.resultsCount.textContent = 'Найдено: ' + svhState.allResults.length;
    svhEl.resultsSource.textContent = '';
    svhEl.loading.style.display = 'none';
  }

  function renderPagination() {
    svhEl.pagination.innerHTML = '';
    if (svhState.totalPages <= 1) return;

    var total = svhState.totalPages;
    var current = svhState.currentPage;

    // Previous
    var prevBtn = document.createElement('button');
    prevBtn.className = 'svh-page-btn';
    prevBtn.textContent = '← Назад';
    prevBtn.disabled = current <= 1;
    prevBtn.addEventListener('click', function () {
      svhState.currentPage = current - 1;
      renderSvhResults();
      renderPagination();
      svhEl.resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
    svhEl.pagination.appendChild(prevBtn);

    // Page numbers
    var range = 3;
    var from = Math.max(1, current - range);
    var to = Math.min(total, current + range);
    for (var i = from; i <= to; i++) {
      var btn = document.createElement('button');
      btn.className = 'svh-page-btn' + (i === current ? ' active' : '');
      btn.textContent = i;
      (function (p) {
        btn.addEventListener('click', function () {
          svhState.currentPage = p;
          renderSvhResults();
          renderPagination();
          svhEl.resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
        });
      })(i);
      svhEl.pagination.appendChild(btn);
    }

    // Next
    var nextBtn = document.createElement('button');
    nextBtn.className = 'svh-page-btn';
    nextBtn.textContent = 'Вперёд →';
    nextBtn.disabled = current >= total;
    nextBtn.addEventListener('click', function () {
      svhState.currentPage = current + 1;
      renderSvhResults();
      renderPagination();
      svhEl.resultsCard.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
    svhEl.pagination.appendChild(nextBtn);

    svhEl.pagination.style.display = 'flex';
  }

  function doSvhSearch() {
    clearSvhError();
    svhEl.resultsCard.hidden = true;
    svhEl.resultsList.innerHTML = '';
    svhEl.pagination.innerHTML = '';
    svhEl.pagination.style.display = 'none';

    var address = (svhEl.address ? svhEl.address.value : '').trim();
    var name = (svhEl.name ? svhEl.name.value : '').trim();
    var customs = (svhEl.customs ? svhEl.customs.value : '').trim();
    var license = (svhEl.license ? svhEl.license.value : '').trim();
    var transport = (svhEl.transport ? svhEl.transport.value : '');

    console.log('[SVH] Search params:', { address: address, name: name, customs: customs, license: license, transport: transport });
    console.log('[SVH] DOM elements:', {
      addressEl: !!svhEl.address,
      nameEl: !!svhEl.name,
      customsEl: !!svhEl.customs,
      licenseEl: !!svhEl.license,
      transportEl: !!svhEl.transport,
    });

    var params = new URLSearchParams();
    if (address) params.set('s_adres', address);
    if (name) params.set('s_name', name);
    if (customs) params.set('s_tam', customs);
    if (license) params.set('s_nlic', license);
    if (transport) params.set('s_vidtrans', transport);

    var url = API + '/api/svh/search?' + params.toString();
    console.log('[SVH] Request URL:', url);

    svhEl.loading.style.display = 'flex';

    fetch(url)
      .then(function (res) {
        console.log('[SVH] Response status:', res.status);
        return res.json();
      })
      .then(function (data) {
        console.log('[SVH] Response data:', data);
        if (data.error) {
          showSvhError(data.error);
          svhEl.loading.style.display = 'none';
          return;
        }
        svhState.allResults = data.results || [];
        svhState.total = data.count || svhState.allResults.length;
        svhState.totalPages = Math.max(1, Math.ceil(svhState.total / svhState.perPage));
        svhState.currentPage = 1;
        svhEl.resultsCard.hidden = false;
        renderSvhResults();
        renderPagination();
      })
      .catch(function (e) {
        console.error('[SVH] Fetch error:', e);
        showSvhError('Ошибка при поиске: ' + e.message);
        svhEl.loading.style.display = 'none';
      });
  }

  svhEl.searchBtn.addEventListener('click', function () {
    doSvhSearch();
  });

  // Enter key triggers search
  [svhEl.address, svhEl.name, svhEl.customs, svhEl.license].forEach(function (input) {
    input.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); doSvhSearch(); }
    });
  });

  svhEl.clearBtn.addEventListener('click', function () {
    svhEl.address.value = '';
    svhEl.name.value = '';
    svhEl.customs.value = '';
    svhEl.license.value = '';
    svhEl.transport.value = '';
    clearSvhError();
    svhEl.resultsCard.hidden = true;
    svhEl.resultsList.innerHTML = '';
    svhEl.pagination.innerHTML = '';
    svhState.allResults = [];
  });

  // ---- AI Chat (Yandex GPT) ----
  var aiEl = {
    messages: document.getElementById('aiMessages'),
    input: document.getElementById('aiInput'),
    sendBtn: document.getElementById('aiSendBtn'),
    loading: document.getElementById('aiLoading'),
    errorBox: document.getElementById('aiErrorBox'),
  };
  var aiSession = '';

  function showAiError(msg) {
    if (aiEl.errorBox) {
      aiEl.errorBox.textContent = msg;
      aiEl.errorBox.style.display = 'block';
    }
  }
  function clearAiError() {
    if (aiEl.errorBox) {
      aiEl.errorBox.style.display = 'none';
      aiEl.errorBox.textContent = '';
    }
  }

  function addAiMessage(role, text) {
    if (!aiEl.messages) return;
    var div = document.createElement('div');
    div.className = 'ai-msg ' + role;
    div.textContent = text;
    aiEl.messages.appendChild(div);
    var box = document.getElementById('aiChatBox');
    if (box) box.scrollTop = box.scrollHeight;
  }

  function doAiChat() {
    var msg = (aiEl.input ? aiEl.input.value : '').trim();
    if (!msg) return;

    clearAiError();
    addAiMessage('user', msg);
    if (aiEl.input) aiEl.input.value = '';

    if (aiEl.loading) aiEl.loading.style.display = 'flex';
    if (aiEl.sendBtn) aiEl.sendBtn.disabled = true;

    var body = { message: msg };
    if (aiSession) body.session = aiSession;

    fetch(API + '/api/ai/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    })
      .then(function (res) { return res.json(); })
      .then(function (data) {
        if (aiEl.loading) aiEl.loading.style.display = 'none';
        if (aiEl.sendBtn) aiEl.sendBtn.disabled = false;

        if (data.error) {
          showAiError(data.error);
          return;
        }
        aiSession = data.session || aiSession;
        addAiMessage('assistant', data.reply);
      })
      .catch(function (e) {
        if (aiEl.loading) aiEl.loading.style.display = 'none';
        if (aiEl.sendBtn) aiEl.sendBtn.disabled = false;
        showAiError('Ошибка: ' + e.message);
      });
  }

  if (aiEl.sendBtn) {
    aiEl.sendBtn.addEventListener('click', doAiChat);
  }
  if (aiEl.input) {
    aiEl.input.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); doAiChat(); }
    });
  }

  // Welcome message
  (function () {
    var welcome = document.createElement('div');
    welcome.className = 'ai-msg system';
    welcome.textContent = '🤖 Привет! Я ИИ-ассистент по таможенным вопросам. Спрашивайте о декларировании, сертификатах, ТН ВЭД, СВХ и других таможенных процедурах.';
    if (aiEl.messages) aiEl.messages.appendChild(welcome);
  })();

  // ---- AI Analyze (Yandex GPT) ----
  var aiAnalyzeBtn = document.getElementById('aiAnalyzeBtn');
  if (aiAnalyzeBtn) {
    aiAnalyzeBtn.addEventListener('click', function () {
      if (!inputFiles.length) {
        showError('Сначала загрузите входные файлы.');
        return;
      }
      aiAnalyzeBtn.disabled = true;
      aiAnalyzeBtn.textContent = '🤖 ИИ анализирует…';
      setProgress(10);

      var fd = new FormData();
      inputFiles.forEach(function (f) { fd.append('inputs', f); });
      fd.append('country', el.countryInput.value.trim() || 'CN');
      fd.append('unit', el.unitInput.value.trim() || 'шт');

      fetch(API + '/api/ai/analyze', {
        method: 'POST',
        body: fd,
      })
        .then(function (res) {
          setProgress(100);
          if (!res.ok) return res.json().then(function (d) { throw new Error(d.error || 'Ошибка'); });
          return res.blob().then(function (blob) { return { blob: blob }; });
        })
        .then(function (data) {
          var filename = 'ai_result.xlsx';
          // Check if fallback mode (X-AI-Mode header might not be accessible in browser)
          aiAnalyzeBtn.textContent = '🤖 ИИ-анализ файлов';
          aiAnalyzeBtn.disabled = false;
          triggerDownload(data.blob, filename);
          showResult(data.blob, filename);
        })
        .catch(function (e) {
          showError('ИИ-анализ не удался: ' + e.message + '. Используется стандартный режим.');
          aiAnalyzeBtn.textContent = '🤖 ИИ-анализ файлов';
          aiAnalyzeBtn.disabled = false;
        })
        .finally(function () {
          hideProgress();
        });
    });
  }

  // ---- Инициализация ----
  loadTemplateInfo();
})();
