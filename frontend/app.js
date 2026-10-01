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
  let lastProtocol = null;

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
    compareOpenBtn: document.getElementById('compareOpenBtn'),
    compareModal: document.getElementById('compareModal'),
    compareCloseBtn: document.getElementById('compareCloseBtn'),
    compareFinding: document.getElementById('compareFinding'),
    compareSource: document.getElementById('compareSource'),
    compareResult: document.getElementById('compareResult'),
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
    lastProtocol = null;
    if (el.compareOpenBtn) el.compareOpenBtn.style.display = 'none';
    closeCompare();
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
    const arr = Array.from(fileListObj).filter((f) => /\.xlsx?$/i.test(f.name));
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

  // log: { summary?, errors?, warnings?, notes?, findings?, compare? }
  function parseProtocol(res) {
    var b64 = res.headers.get('X-Protocol');
    if (!b64) return null;
    try {
      var bytes = Uint8Array.from(atob(b64), function (c) { return c.charCodeAt(0); });
      return JSON.parse(new TextDecoder('utf-8').decode(bytes));
    } catch (e) {
      return null;
    }
  }

  function b64ToBlob(b64, type) {
    var bin = atob(b64);
    var arr = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], { type: type });
  }

  function readResultPayload(res) {
    var ctype = (res.headers.get('content-type') || '').toLowerCase();
    if (ctype.indexOf('json') >= 0) {
      return res.json().then(function (data) {
        if (data && data.error) throw new Error(data.error);
        if (!data || !data.file_b64) throw new Error('Сервер не вернул файл');
        return {
          blob: b64ToBlob(data.file_b64, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
          protocol: data.protocol || null,
          filename: data.filename || 'result.xlsx',
        };
      });
    }
    var protocol = parseProtocol(res);
    return res.blob().then(function (blob) {
      return { blob: blob, protocol: protocol, filename: 'result.xlsx' };
    });
  }

  function formatApiError(d, status) {
    if (d && typeof d.error === 'string' && d.error) return d.error;
    if (d && typeof d.message === 'string' && d.message) return d.message;
    if (d && typeof d.detail === 'string' && d.detail) return d.detail;
    if (d && Array.isArray(d.detail) && d.detail.length) {
      return d.detail.map(function (x) {
        if (typeof x === 'string') return x;
        return (x && (x.msg || x.message)) || JSON.stringify(x);
      }).join('; ');
    }
    return 'Ошибка сервера (HTTP ' + (status || '?') + ')';
  }

  function readErrorBody(res) {
    return res.text().then(function (text) {
      var d = {};
      try { d = JSON.parse(text); } catch (e) {}
      var msg = formatApiError(d, res.status);
      if ((!d || (!d.error && !d.detail && !d.message)) && text) {
        var snippet = text.replace(/\s+/g, ' ').trim().slice(0, 240);
        if (snippet && snippet.charAt(0) !== '{') msg += ': ' + snippet;
      }
      throw new Error(msg);
    });
  }

  function showErrorLog(log) {
    if (!el.errorLog) return;
    log = log || {};
    lastProtocol = log;

    el.errorLogList.innerHTML = '';

    const addEntry = (type, message, findingNo) => {
      const li = document.createElement('li');
      li.className = 'log-' + type;
      li.textContent = message;
      if (findingNo != null) {
        li.className += ' log-clickable';
        li.title = 'Открыть сравнение по этой позиции';
        li.addEventListener('click', function () { openCompare(findingNo); });
      }
      el.errorLogList.appendChild(li);
    };

    var findings = log.findings || [];
    if (findings.length) {
      findings.forEach(function (f) {
        addEntry(f.severity === 'error' ? 'error' : 'warning', f.message || '', f.no);
      });
    } else {
      (log.errors || []).forEach((m) => addEntry('error', m));
      (log.warnings || []).forEach((m) => addEntry('warning', m));
    }
    (log.unknown || []).forEach((m) => addEntry('unknown', m));
    (log.notes || []).forEach((m) => addEntry('note', m));

    const hasErrors = (log.errors || []).length > 0 || findings.some(function (f) { return f.severity === 'error'; });
    const hasWarnings = (log.warnings || []).length > 0 || (log.unknown || []).length > 0
      || findings.some(function (f) { return f.severity !== 'error'; });
    el.errorLog.classList.toggle('has-errors', hasErrors);
    el.errorLog.classList.toggle('has-issues', hasWarnings && !hasErrors);

    if (log.summary) {
      el.errorLogSummary.textContent = log.summary;
    } else if ((log.unknown || []).length) {
      el.errorLogSummary.textContent = 'Обнаружены неизвестные элементы в загружаемом файле.';
    } else if (hasErrors || hasWarnings) {
      el.errorLogSummary.textContent = 'Есть замечания, проверьте протокол.';
    } else {
      el.errorLogSummary.textContent = 'Ошибок не обнаружено, файл обработан корректно.';
    }

    var canCompare = !!(log.compare && (log.compare.rows || []).length);
    if (el.compareOpenBtn) {
      el.compareOpenBtn.style.display = canCompare ? 'inline-flex' : 'none';
    }

    el.errorLog.style.display = 'block';
  }

  function closeCompare() {
    if (el.compareModal) el.compareModal.hidden = true;
  }

  function _cellText(v) {
    if (v == null || v === '') return '';
    return String(v);
  }

  function _srcCaption(src) {
    if (!src || !src.sheet) return '';
    var t = 'Лист «' + src.sheet + '»';
    if (src.row) t += ', строка ' + src.row;
    return t;
  }

  function renderCompareTables(focusNo) {
    var cmp = (lastProtocol && lastProtocol.compare) || {};
    var fields = cmp.fields || ['article', 'name', 'qty', 'price', 'amount', 'net_weight', 'gross_weight'];
    var labels = cmp.labels || {};
    var rows = cmp.rows || [];
    var findings = (lastProtocol && lastProtocol.findings) || [];
    var issueByNo = {};
    findings.forEach(function (f) {
      if (f.no == null) return;
      if (!issueByNo[f.no]) issueByNo[f.no] = [];
      if (f.field && issueByNo[f.no].indexOf(f.field) < 0) issueByNo[f.no].push(f.field);
    });

    function tableHtml(side) {
      var html = '<table class="compare-table"><thead><tr><th>№</th>';
      fields.forEach(function (f) {
        html += '<th>' + (labels[f] || f) + '</th>';
      });
      html += '</tr></thead><tbody>';
      rows.forEach(function (row) {
        var src = row.source || {};
        var data = side === 'source' ? src : (row.result || {});
        if (side === 'source' && !(src.article || src.name || src.qty != null || src.price != null)) {
          data = row.source_pack || src;
        }
        var hits = issueByNo[row.no] || row.issues || [];
        var trClass = (focusNo != null && row.no === focusNo) || (hits.length && focusNo == null) ? ' is-hit' : '';
        html += '<tr class="' + trClass.trim() + '" data-no="' + row.no + '">';
        html += '<td>' + row.no;
        if (side === 'source') {
          var cap = _srcCaption(data.sheet ? data : row.source_pack);
          if (cap) html += '<span class="compare-src-meta">' + cap + '</span>';
        }
        html += '</td>';
        fields.forEach(function (f) {
          var hit = hits.indexOf(f) >= 0 ? ' cell-hit' : '';
          html += '<td class="' + hit.trim() + '">' + _cellText(data[f]) + '</td>';
        });
        html += '</tr>';
      });
      html += '</tbody></table>';
      return html;
    }

    if (el.compareSource) el.compareSource.innerHTML = tableHtml('source');
    if (el.compareResult) el.compareResult.innerHTML = tableHtml('result');

    var focusMsg = '';
    if (focusNo != null) {
      var f = findings.filter(function (x) { return x.no === focusNo; });
      focusMsg = f.map(function (x) { return x.message; }).join(' · ');
    }
    if (el.compareFinding) el.compareFinding.textContent = focusMsg;

    var hit = el.compareResult && el.compareResult.querySelector('tr.is-hit');
    if (hit && hit.scrollIntoView) hit.scrollIntoView({ block: 'center' });
    var hitL = el.compareSource && el.compareSource.querySelector('tr.is-hit');
    if (hitL && hitL.scrollIntoView) hitL.scrollIntoView({ block: 'center' });
  }

  function openCompare(focusNo) {
    if (!lastProtocol || !lastProtocol.compare) return;
    if (el.compareModal) el.compareModal.hidden = false;
    renderCompareTables(focusNo);
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
        throw new Error(formatApiError(data, res.status) || 'Ошибка формирования файла');
      }

      const payload = await readResultPayload(res);
      const filename = payload.filename || 'result.xlsx';
      showResult(payload.blob, filename);
      triggerDownload(payload.blob, filename);
      showErrorLog(payload.protocol);
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
        document.body.classList.toggle('catalog-wide', target === 'articles');
      });
    });
  })();

  // ---- Article catalog ----
  var articleState = { q: '', page: 1, pages: 1, fields: [], filters: {} };
  var articleEditingId = null;

  var articleEl = {
    search: document.getElementById('articleSearch'),
    searchBtn: document.getElementById('articleSearchBtn'),
    addBtn: document.getElementById('articleAddBtn'),
    error: document.getElementById('articleError'),
    meta: document.getElementById('articleMeta'),
    rows: document.getElementById('articleRows'),
    pager: document.getElementById('articlePager'),
    modal: document.getElementById('articleModal'),
    form: document.getElementById('articleForm'),
    title: document.getElementById('articleFormTitle'),
    fields: document.getElementById('articleFields'),
    formError: document.getElementById('articleFormError'),
    cancel: document.getElementById('articleCancelBtn'),
  };

  var loadTableEl = {
    file: document.getElementById('loadTableFile'),
    compareFile: document.getElementById('loadCompareFile'),
    button: document.getElementById('loadTableBtn'),
    compare: document.getElementById('loadCompareBtn'),
    error: document.getElementById('loadTableError'),
    compareError: document.getElementById('loadCompareError'),
  };
  var diffEl = {
    modal: document.getElementById('diffModal'),
    close: document.getElementById('diffCloseBtn'),
    meta: document.getElementById('diffMeta'),
    error: document.getElementById('diffError'),
    rows: document.getElementById('loadCompareRows'),
    form: document.getElementById('diffForm'),
    title: document.getElementById('diffFormTitle'),
    fields: document.getElementById('diffFields'),
    formError: document.getElementById('diffFormError'),
    cancel: document.getElementById('diffFormCancel'),
  };
  var compareDiffs = [];
  var diffEditingId = null;
  var COMPARE_KEYS = ['hs_code', 'origin_code', 'group_description'];
  var DIFF_FIELDS = [
    { key: 'article', label: 'Артикул' },
    { key: 'description', label: 'Описание товара' },
    { key: 'origin_code', label: 'Код страны происхождения' },
    { key: 'hs_code', label: 'Код товара' },
    { key: 'group_description', label: 'Описание группы' },
    { key: 'manufacturer', label: 'Наименование фирмы-изготовителя' },
    { key: 'brand', label: 'Марка' },
    { key: 'model', label: 'Модель' },
    { key: 'extra_code', label: 'Доп. код' },
  ];

  function showPanelError(box, msg) {
    if (!box) return;
    box.textContent = msg;
    box.style.display = 'block';
  }

  function clearPanelError(box) {
    if (!box) return;
    box.style.display = 'none';
    box.textContent = '';
  }

  function showLoadTableError(msg) {
    showPanelError(loadTableEl.error, msg);
  }

  function clearLoadTableError() {
    clearPanelError(loadTableEl.error);
  }

  function showCompareError(msg) {
    showPanelError(loadTableEl.compareError, msg);
  }

  function clearCompareError() {
    clearPanelError(loadTableEl.compareError);
  }

  function selectedFile(input) {
    return input && input.files && input.files[0];
  }

  function saveLoadTableBlob(blob) {
    var link = document.createElement('a');
    var objectUrl = URL.createObjectURL(blob);
    link.href = objectUrl;
    link.download = 'загрузочная таблица.xlsx';
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () { URL.revokeObjectURL(objectUrl); }, 1000);
  }

  function postLoadTable(download, onProgress, file) {
    return new Promise(function (resolve, reject) {
      if (!file) {
        resolve(null);
        return;
      }
      var fd = new FormData();
      fd.append('file', file);
      var xhr = new XMLHttpRequest();
      xhr.open('POST', API + '/api/articles/load-table' + (download ? '?download=1' : '?compare=1'));
      xhr.responseType = download ? 'blob' : 'json';
      if (onProgress && xhr.upload) {
        xhr.upload.onprogress = function (event) {
          if (!event.lengthComputable || !onProgress) return;
          onProgress(Math.min(40, Math.round(event.loaded / event.total * 40)));
        };
      }
      xhr.onload = function () {
        if (xhr.status >= 200 && xhr.status < 300) {
          resolve(xhr);
          return;
        }
        var fail = function (message) { reject(new Error(message || 'Не удалось собрать таблицу')); };
        if (download && xhr.response) {
          var reader = new FileReader();
          reader.onload = function () {
            try { fail(JSON.parse(reader.result).error); }
            catch (err) { fail(); }
          };
          reader.readAsText(xhr.response);
          return;
        }
        fail(xhr.response && xhr.response.error);
      };
      xhr.onerror = function () { reject(new Error('Сайт с базой артикулов недоступен')); };
      xhr.send(fd);
    });
  }

  function compareCell(diff, key) {
    var td = document.createElement('td');
    var incoming = (diff.file && diff.file[key]) || '';
    var current = (diff.db && diff.db[key]) || '';
    if (!incoming || incoming === current) return td;
    if (key === 'hs_code' && !(/^\d{10}$/.test(incoming) && /^\d{10}$/.test(current))) return td;
    var fileLine = document.createElement('span');
    fileLine.className = 'compare-file';
    fileLine.textContent = 'файл: ' + incoming;
    var dbLine = document.createElement('span');
    dbLine.className = 'compare-db';
    dbLine.textContent = 'база: ' + (current || '—');
    td.appendChild(fileLine);
    td.appendChild(dbLine);
    return td;
  }

  function showDiffError(msg) {
    if (!diffEl.error) return;
    diffEl.error.textContent = msg;
    diffEl.error.style.display = 'block';
  }

  function clearDiffError() {
    if (!diffEl.error) return;
    diffEl.error.style.display = 'none';
    diffEl.error.textContent = '';
  }

  function showDiffFormError(msg) {
    if (!diffEl.formError) return;
    diffEl.formError.textContent = msg;
    diffEl.formError.style.display = 'block';
  }

  function clearDiffFormError() {
    if (!diffEl.formError) return;
    diffEl.formError.style.display = 'none';
    diffEl.formError.textContent = '';
  }

  function closeDiffForm() {
    diffEditingId = null;
    if (diffEl.form) diffEl.form.hidden = true;
    if (diffEl.fields) diffEl.fields.innerHTML = '';
    clearDiffFormError();
    markDiffRow();
  }

  function closeDiffWindow() {
    closeDiffForm();
    if (diffEl.modal) diffEl.modal.hidden = true;
  }

  function markDiffRow() {
    if (!diffEl.rows) return;
    diffEl.rows.querySelectorAll('tr').forEach(function (tr) {
      tr.classList.toggle('diff-row-active', !!(diffEditingId && String(tr.dataset.id) === String(diffEditingId)));
    });
  }

  var manufacturerNamesRequest = null;

  function loadManufacturerNames() {
    if (manufacturerNamesRequest) return manufacturerNamesRequest;
    manufacturerNamesRequest = fetch(API + '/api/articles/manufacturers').then(function (res) {
      if (!res.ok) return [];
      return res.json();
    }).then(function (data) {
      var seen = {};
      var items = [];
      (data && data.items || []).forEach(function (row) {
        var name = String(row && row.name || '').trim().toUpperCase();
        if (!name || seen[name]) return;
        seen[name] = true;
        var countries = [];
        (row.countries || []).forEach(function (code) {
          var country = String(code || '').trim().toUpperCase();
          if (country && countries.indexOf(country) < 0) countries.push(country);
        });
        items.push({ name: name, countries: countries });
      });
      return items;
    }).catch(function () {
      manufacturerNamesRequest = null;
      return [];
    });
    return manufacturerNamesRequest;
  }

  function bindManufacturerPicker(input) {
    var box = document.createElement('span');
    box.className = 'manufacturer-picker';
    input.parentNode.insertBefore(box, input);
    box.appendChild(input);
    var arrow = document.createElement('button');
    arrow.type = 'button';
    arrow.className = 'manufacturer-arrow';
    arrow.setAttribute('aria-label', 'Выбрать изготовителя из базы');
    box.appendChild(arrow);
    var menu = document.createElement('div');
    menu.className = 'manufacturer-menu';
    menu.hidden = true;
    box.appendChild(menu);
    arrow.addEventListener('click', function (event) {
      event.preventDefault();
      event.stopPropagation();
      if (!menu.hidden) {
        menu.hidden = true;
        return;
      }
      document.querySelectorAll('.manufacturer-menu').forEach(function (other) {
        other.hidden = true;
      });
      loadManufacturerNames().then(function (items) {
        menu.innerHTML = '';
        if (!items.length) {
          var empty = document.createElement('div');
          empty.className = 'manufacturer-empty';
          empty.textContent = 'В базе пока нет изготовителей';
          menu.appendChild(empty);
        }
        items.forEach(function (row) {
          var item = document.createElement('button');
          item.type = 'button';
          item.className = 'manufacturer-option';
          var label = document.createElement('span');
          label.textContent = row.name;
          item.appendChild(label);
          if (row.countries.length) {
            var country = document.createElement('span');
            country.className = 'manufacturer-country';
            country.textContent = ' - ' + row.countries.join(', ');
            item.appendChild(country);
          }
          item.addEventListener('click', function (ev) {
            ev.preventDefault();
            input.value = row.name;
            menu.hidden = true;
          });
          menu.appendChild(item);
        });
        menu.hidden = false;
      });
    });
    document.addEventListener('click', function (event) {
      if (!box.contains(event.target)) menu.hidden = true;
    });
  }

  function openDiffForm(diff) {
    clearDiffFormError();
    diffEditingId = diff.id;
    var item = diff.db || {};
    if (diffEl.title) diffEl.title.textContent = diff.article || 'Изменить';
    diffEl.fields.innerHTML = '';
    DIFF_FIELDS.forEach(function (field) {
      var wrap = document.createElement('label');
      wrap.className = 'article-field';
      var caption = document.createElement('span');
      caption.textContent = field.label;
      var longText = field.key === 'description' || field.key === 'group_description';
      if (longText) wrap.classList.add('article-field-wide');
      var input = document.createElement(longText ? 'textarea' : 'input');
      input.name = field.key;
      input.value = item[field.key] ? String(item[field.key]).toUpperCase() : '';
      if (field.key === 'article') input.required = true;
      input.addEventListener('input', function () {
        var start = input.selectionStart;
        var end = input.selectionEnd;
        var upper = input.value.toUpperCase();
        if (upper === input.value) return;
        input.value = upper;
        try { input.setSelectionRange(start, end); } catch (err) {}
      });
      if (field.key === 'hs_code') {
        var originalCode = input.value.trim();
        var lookupTimer = null;
        input.addEventListener('input', function () {
          clearTimeout(lookupTimer);
          var code = input.value.trim();
          if (!code || code === originalCode) return;
          lookupTimer = setTimeout(function () {
            fetch(API + '/api/articles/common-description?hs_code=' + encodeURIComponent(code))
              .then(function (res) { return res.ok ? res.json() : null; })
              .then(function (data) {
                if (!data || !data.description || input.value.trim() !== code) return;
                var desc = diffEl.fields.querySelector('[name="description"]');
                if (desc) desc.value = String(data.description).toUpperCase();
              })
              .catch(function () {});
          }, 300);
        });
      }
      wrap.appendChild(caption);
      wrap.appendChild(input);
      if (field.key === 'manufacturer') bindManufacturerPicker(input);
      diffEl.fields.appendChild(wrap);
    });
    diffEl.form.hidden = false;
    markDiffRow();
  }

  function renderCompare() {
    if (!diffEl.rows || !diffEl.modal) return;
    diffEl.rows.innerHTML = '';
    if (!compareDiffs.length) closeDiffForm();
    compareDiffs.forEach(function (diff) {
      var tr = document.createElement('tr');
      tr.dataset.id = diff.id;
      var articleCell = document.createElement('td');
      articleCell.textContent = diff.article || '';
      tr.appendChild(articleCell);
      COMPARE_KEYS.forEach(function (key) { tr.appendChild(compareCell(diff, key)); });
      var action = document.createElement('td');
      var button = document.createElement('button');
      button.type = 'button';
      button.className = 'btn btn-ghost';
      button.textContent = 'Изменить';
      button.addEventListener('click', function () { openDiffForm(diff); });
      action.appendChild(button);
      tr.appendChild(action);
      diffEl.rows.appendChild(tr);
    });
    markDiffRow();
    diffEl.modal.hidden = false;
  }

  function dropSavedDiff(savedId) {
    compareDiffs = compareDiffs.filter(function (diff) {
      return String(diff.id) !== String(savedId);
    });
    if (diffEl.meta) {
      diffEl.meta.textContent = compareDiffs.length
        ? 'Расхождений: ' + compareDiffs.length + '.'
        : 'Расхождений не осталось.';
    }
    closeDiffForm();
    renderCompare();
  }

  if (loadTableEl.button) {
    loadTableEl.button.addEventListener('click', async function () {
      clearLoadTableError();
      var file = selectedFile(loadTableEl.file);
      if (!file) {
        showLoadTableError('Выберите файл Excel.');
        return;
      }
      var button = loadTableEl.button;
      var progress = 0;
      button.disabled = true;
      button.textContent = '0%';
      var timer = setInterval(function () {
        if (progress < 90) {
          progress += 2;
          button.textContent = progress + '%';
        }
      }, 200);
      try {
        var xhr = await postLoadTable(true, function (value) {
          if (value > progress) {
            progress = value;
            button.textContent = progress + '%';
          }
        }, file);
        if (!xhr) return;
        clearInterval(timer);
        timer = null;
        button.textContent = '100%';
        saveLoadTableBlob(xhr.response);
      } catch (err) {
        showLoadTableError(err.message || 'Не удалось собрать таблицу');
      } finally {
        if (timer) clearInterval(timer);
        setTimeout(function () {
          button.textContent = 'Собрать таблицу';
          button.disabled = false;
        }, 400);
      }
    });
  }

  if (loadTableEl.compare) {
    loadTableEl.compare.addEventListener('click', async function () {
      clearCompareError();
      var file = selectedFile(loadTableEl.compareFile);
      if (!file) {
        showCompareError('Выберите файл Excel.');
        return;
      }
      var button = loadTableEl.compare;
      var progress = 0;
      button.disabled = true;
      button.textContent = '0%';
      var timer = setInterval(function () {
        if (progress < 90) {
          progress += 2;
          button.textContent = progress + '%';
        }
      }, 200);
      try {
        var xhr = await postLoadTable(false, function (value) {
          if (value > progress) {
            progress = value;
            button.textContent = progress + '%';
          }
        }, file);
        if (!xhr) return;
        clearInterval(timer);
        timer = null;
        button.textContent = '100%';
        var data = xhr.response || {};
        compareDiffs = data.diffs || [];
        clearDiffError();
        closeDiffForm();
        if (diffEl.meta) {
          if (!compareDiffs.length && !data.found_count) {
            diffEl.meta.textContent = 'Артикулы из файла в базе не найдены.';
          } else if (!compareDiffs.length) {
            diffEl.meta.textContent = 'Расхождений нет.';
          } else {
            diffEl.meta.textContent = 'Расхождений: ' + compareDiffs.length + '.';
          }
        }
        renderCompare();
      } catch (err) {
        showCompareError(err.message || 'Не удалось сравнить файл с базой');
      } finally {
        if (timer) clearInterval(timer);
        setTimeout(function () {
          button.textContent = 'Сравнение';
          button.disabled = false;
        }, 400);
      }
    });
  }

  if (diffEl.close) diffEl.close.addEventListener('click', closeDiffWindow);
  if (diffEl.modal) {
    diffEl.modal.addEventListener('click', function (event) {
      if (event.target === diffEl.modal) closeDiffWindow();
    });
  }
  if (diffEl.cancel) diffEl.cancel.addEventListener('click', closeDiffForm);
  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape' && diffEl.modal && !diffEl.modal.hidden) closeDiffWindow();
  });
  if (diffEl.form) {
    diffEl.form.addEventListener('submit', function (event) {
      event.preventDefault();
      if (diffEditingId == null) return;
      clearDiffFormError();
      var payload = {};
      diffEl.fields.querySelectorAll('input, textarea').forEach(function (input) {
        payload[input.name] = input.value.toUpperCase();
      });
      var savedId = diffEditingId;
      fetch(API + '/api/articles/' + savedId, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      }).then(function (res) {
        if (!res.ok) return readErrorBody(res);
        return res.json();
      }).then(function () {
        dropSavedDiff(savedId);
        loadArticles(articleState.page);
      }).catch(function (err) {
        showDiffFormError(err.message || 'Не удалось сохранить артикул');
      });
    });
  }

  function showArticleError(msg) {
    if (!articleEl.error) return;
    articleEl.error.textContent = msg;
    articleEl.error.style.display = 'block';
  }

  function clearArticleError() {
    if (!articleEl.error) return;
    articleEl.error.style.display = 'none';
    articleEl.error.textContent = '';
  }

  function renderArticleForm(item) {
    articleEl.fields.innerHTML = '';
    (articleState.fields.length ? articleState.fields : [
      { key: 'article', label: 'Артикул' },
      { key: 'description', label: 'Описание товара' },
      { key: 'origin_code', label: 'Код страны происхождения' },
      { key: 'hs_code', label: 'Код товара' },
      { key: 'group_description', label: 'Описание группы' },
      { key: 'manufacturer', label: 'Наименование фирмы-изготовителя' },
      { key: 'brand', label: 'Марка' },
      { key: 'model', label: 'Модель' },
      { key: 'extra_code', label: 'Доп. код' },
    ]).forEach(function (field) {
      var wrap = document.createElement('label');
      wrap.className = 'article-field';
      var caption = document.createElement('span');
      caption.textContent = field.label;
      var longText = field.key === 'description' || field.key === 'group_description';
      if (longText) wrap.classList.add('article-field-wide');
      var input = document.createElement(longText ? 'textarea' : 'input');
      input.name = field.key;
      input.value = item && item[field.key] ? String(item[field.key]).toUpperCase() : '';
      if (field.key === 'article') input.required = true;
      input.addEventListener('input', function () {
        var start = input.selectionStart;
        var end = input.selectionEnd;
        var upper = input.value.toUpperCase();
        if (upper === input.value) return;
        input.value = upper;
        try { input.setSelectionRange(start, end); } catch (err) {}
      });
      if (field.key === 'hs_code') {
        var originalCode = input.value.trim();
        var lookupTimer = null;
        input.addEventListener('input', function () {
          clearTimeout(lookupTimer);
          var code = input.value.trim();
          if (!code || code === originalCode) return;
          lookupTimer = setTimeout(function () {
            fetch(API + '/api/articles/common-description?hs_code=' + encodeURIComponent(code))
              .then(function (res) {
                if (!res.ok) return readErrorBody(res);
                return res.json();
              })
              .then(function (data) {
                if (!data || !data.description || input.value.trim() !== code) return;
                var desc = articleEl.fields.querySelector('[name="description"]');
                if (desc) desc.value = String(data.description).toUpperCase();
              })
              .catch(function () {});
          }, 300);
        });
      }
      wrap.appendChild(caption);
      wrap.appendChild(input);
      if (field.key === 'manufacturer') bindManufacturerPicker(input);
      articleEl.fields.appendChild(wrap);
    });
  }

  function showFormError(msg) {
    if (!articleEl.formError) return;
    articleEl.formError.textContent = msg;
    articleEl.formError.style.display = 'block';
  }

  function clearFormError() {
    if (!articleEl.formError) return;
    articleEl.formError.style.display = 'none';
    articleEl.formError.textContent = '';
  }

  function openArticleForm(item) {
    clearFormError();
    articleEditingId = item && item.id ? item.id : null;
    articleEl.title.textContent = articleEditingId ? 'Изменить артикул' : 'Новый артикул';
    renderArticleForm(item);
    articleEl.modal.hidden = false;
    markEditingRow();
  }

  function closeArticleForm() {
    articleEl.modal.hidden = true;
    articleEditingId = null;
    markEditingRow();
  }

  function markEditingRow() {
    document.querySelectorAll('#articleRows tr').forEach(function (tr) {
      tr.classList.toggle('article-row-active', !!(articleEditingId && String(tr.dataset.id) === String(articleEditingId)));
    });
  }

  var COL_WIDTH_KEY = 'tfArticleColWidths';
  var ROW_HEIGHT_KEY = 'tfArticleRowHeights';
  var HEADER_HEIGHT_KEY = 'tfArticleHeaderHeight';

  function readStored(key, fallback) {
    try {
      var raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch (err) {
      return fallback;
    }
  }

  function writeStored(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (err) {}
  }

  function articleTable() {
    return document.querySelector('.article-table');
  }

  function applyColumnWidths(widths) {
    var table = articleTable();
    if (!table || !widths) return;
    var cols = table.querySelectorAll('col');
    if (widths.length !== cols.length) return;
    var sum = 0;
    cols.forEach(function (col, i) {
      var width = Math.max(48, Math.round(widths[i]));
      col.style.width = width + 'px';
      sum += width;
    });
    table.style.width = sum + 'px';
  }

  function applyHeaderHeight(height) {
    var inner = Math.max(24, Math.round(height) - 16);
    articleTable().querySelectorAll('thead .cell-clip').forEach(function (clip) {
      clip.style.maxHeight = inner + 'px';
      clip.style.minHeight = inner + 'px';
    });
  }

  function applyRowHeight(tr, height) {
    var inner = Math.max(18, Math.round(height) - 16);
    tr.querySelectorAll('.cell-clip').forEach(function (clip) {
      clip.style.maxHeight = inner + 'px';
      clip.style.minHeight = inner + 'px';
    });
  }

  function trackDrag(className, move, done) {
    document.body.classList.add(className);
    function stop() {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', stop);
      document.body.classList.remove(className);
      done();
    }
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', stop);
  }

  function buildArticleCell(text) {
    var td = document.createElement('td');
    var clip = document.createElement('div');
    clip.className = 'cell-clip';
    clip.textContent = text;
    td.appendChild(clip);
    var handle = document.createElement('div');
    handle.className = 'row-resizer';
    handle.title = 'Высота строки';
    handle.addEventListener('mousedown', function (e) {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      var tr = td.parentElement;
      var startY = e.clientY;
      var start = tr.getBoundingClientRect().height;
      trackDrag('article-row-resizing', function (ev) {
        applyRowHeight(tr, Math.max(28, start + ev.clientY - startY));
      }, function () {
        if (!tr.dataset.id) return;
        var heights = readStored(ROW_HEIGHT_KEY, {});
        heights[tr.dataset.id] = Math.round(tr.getBoundingClientRect().height);
        writeStored(ROW_HEIGHT_KEY, heights);
      });
    });
    td.appendChild(handle);
    return td;
  }

  function bindColumnResize() {
    var table = articleTable();
    if (!table || table.dataset.resizeReady) return;
    table.dataset.resizeReady = '1';
    var saved = readStored(COL_WIDTH_KEY, null);
    if (saved) applyColumnWidths(saved);
    table.querySelectorAll('thead th').forEach(function (th, index) {
      var clip = document.createElement('div');
      clip.className = 'cell-clip';
      while (th.firstChild) clip.appendChild(th.firstChild);
      th.appendChild(clip);
      var colHandle = document.createElement('div');
      colHandle.className = 'col-resizer';
      colHandle.title = 'Ширина столбца';
      colHandle.addEventListener('mousedown', function (e) {
        if (e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        var widths = Array.from(table.querySelectorAll('thead th')).map(function (cell) {
          return cell.getBoundingClientRect().width;
        });
        var startX = e.clientX;
        var start = widths[index];
        trackDrag('article-resizing', function (ev) {
          widths[index] = Math.max(48, start + ev.clientX - startX);
          applyColumnWidths(widths);
        }, function () {
          writeStored(COL_WIDTH_KEY, widths.map(function (width) { return Math.round(width); }));
        });
      });
      var rowHandle = document.createElement('div');
      rowHandle.className = 'row-resizer';
      rowHandle.title = 'Высота строки';
      rowHandle.addEventListener('mousedown', function (e) {
        if (e.button !== 0) return;
        e.preventDefault();
        e.stopPropagation();
        var startY = e.clientY;
        var start = th.parentElement.getBoundingClientRect().height;
        trackDrag('article-row-resizing', function (ev) {
          applyHeaderHeight(Math.max(36, start + ev.clientY - startY));
        }, function () {
          writeStored(HEADER_HEIGHT_KEY, Math.round(th.parentElement.getBoundingClientRect().height));
        });
      });
      th.appendChild(colHandle);
      th.appendChild(rowHandle);
    });
    var headerHeight = readStored(HEADER_HEIGHT_KEY, 0);
    if (headerHeight) applyHeaderHeight(headerHeight);
  }

  function renderArticleTable(data) {
    articleState.fields = data.fields || articleState.fields;
    articleState.pages = data.pages || 1;
    articleEl.rows.innerHTML = '';
    (data.items || []).forEach(function (item) {
      var tr = document.createElement('tr');
      tr.dataset.id = item.id ? String(item.id) : '';
      [item.article, item.description, item.origin_code, item.hs_code, item.group_description, item.manufacturer, item.brand, item.model, item.extra_code].forEach(function (value) {
        tr.appendChild(buildArticleCell(value || ''));
      });
      var action = buildArticleCell('');
      action.classList.add('cell-action');
      var btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn btn-ghost';
      btn.textContent = 'Изменить';
      btn.addEventListener('click', function () { openArticleForm(item); });
      action.querySelector('.cell-clip').appendChild(btn);
      tr.appendChild(action);
      articleEl.rows.appendChild(tr);
      var savedHeight = readStored(ROW_HEIGHT_KEY, {})[tr.dataset.id];
      if (savedHeight) applyRowHeight(tr, savedHeight);
    });
    if (!(data.items || []).length) {
      var empty = document.createElement('tr');
      var td = document.createElement('td');
      td.colSpan = 10;
      td.textContent = 'Ничего не найдено.';
      empty.appendChild(td);
      articleEl.rows.appendChild(empty);
    }
    articleEl.meta.textContent = 'Записей: ' + (data.total || 0);
    markEditingRow();
    articleEl.pager.innerHTML = '';
    if ((data.pages || 1) > 1) {
      var prev = document.createElement('button');
      prev.type = 'button';
      prev.className = 'btn btn-ghost';
      prev.textContent = 'Назад';
      prev.disabled = articleState.page <= 1;
      prev.addEventListener('click', function () { loadArticles(articleState.page - 1); });
      var next = document.createElement('button');
      next.type = 'button';
      next.className = 'btn btn-ghost';
      next.textContent = 'Дальше';
      next.disabled = articleState.page >= data.pages;
      next.addEventListener('click', function () { loadArticles(articleState.page + 1); });
      var label = document.createElement('span');
      label.textContent = articleState.page + ' / ' + data.pages;
      articleEl.pager.appendChild(prev);
      articleEl.pager.appendChild(label);
      articleEl.pager.appendChild(next);
    }
  }

  var filterEl = {
    box: document.getElementById('articleFilter'),
    search: document.getElementById('articleFilterSearch'),
    all: document.getElementById('articleFilterAll'),
    list: document.getElementById('articleFilterList'),
    ok: document.getElementById('articleFilterOk'),
    cancel: document.getElementById('articleFilterCancel'),
  };
  var filterField = null;
  var filterDraft = {};
  var filterValues = [];
  var FILTER_ICON = '<svg viewBox="0 0 16 16" width="14" height="14" aria-hidden="true"><path fill="currentColor" d="M1.5 2.5h13L9.2 8.4V13L6.4 14.2V8.4z"/></svg>';

  document.querySelectorAll('.article-filter-btn').forEach(function (btn) {
    btn.innerHTML = FILTER_ICON;
  });

  function markActiveFilters() {
    document.querySelectorAll('.article-filter-btn').forEach(function (btn) {
      btn.classList.toggle('active', Object.prototype.hasOwnProperty.call(articleState.filters, btn.dataset.field));
    });
  }

  function closeColumnFilter() {
    if (filterEl.box) filterEl.box.hidden = true;
    filterField = null;
  }

  function visibleFilterInputs() {
    return Array.from(filterEl.list.querySelectorAll('input[type="checkbox"]'));
  }

  function syncFilterAll() {
    var inputs = visibleFilterInputs();
    var checked = inputs.filter(function (input) { return input.checked; }).length;
    filterEl.all.checked = inputs.length > 0 && checked === inputs.length;
    filterEl.all.indeterminate = checked > 0 && checked < inputs.length;
  }

  function draftKey(value) {
    return value === '' ? '\0blank' : value;
  }

  function renderFilterList() {
    var needle = (filterEl.search.value || '').trim().toLowerCase();
    filterEl.list.innerHTML = '';
    filterValues.forEach(function (value) {
      var caption = value === '' ? '(Пустые)' : value;
      if (needle && caption.toLowerCase().indexOf(needle) === -1) return;
      var label = document.createElement('label');
      var input = document.createElement('input');
      input.type = 'checkbox';
      input._filterValue = value;
      input.checked = !!filterDraft[draftKey(value)];
      input.addEventListener('change', function () {
        filterDraft[draftKey(value)] = input.checked;
        syncFilterAll();
      });
      var text = document.createElement('span');
      text.textContent = caption;
      label.appendChild(input);
      label.appendChild(text);
      filterEl.list.appendChild(label);
    });
    syncFilterAll();
  }

  function selectFilterMatches() {
    var needle = (filterEl.search.value || '').trim().toLowerCase();
    filterValues.forEach(function (value) {
      var caption = value === '' ? '(Пустые)' : value;
      filterDraft[draftKey(value)] = !needle || caption.toLowerCase().indexOf(needle) !== -1;
    });
    renderFilterList();
  }

  function commitColumnFilter() {
    if (!filterField) return;
    var checked = filterValues.filter(function (value) { return filterDraft[draftKey(value)]; });
    if (!checked.length) {
      articleState.filters[filterField] = [];
    } else if (checked.length === filterValues.length) {
      delete articleState.filters[filterField];
    } else {
      articleState.filters[filterField] = checked;
    }
    closeColumnFilter();
  }

  function openColumnFilter(btn) {
    var field = btn.dataset.field;
    if (filterField === field && filterEl.box && !filterEl.box.hidden) {
      closeColumnFilter();
      return;
    }
    var others = {};
    Object.keys(articleState.filters).forEach(function (key) {
      if (key !== field) others[key] = articleState.filters[key];
    });
    fetch(API + '/api/articles/values', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ field: field, q: articleState.q, filters: others }),
    }).then(function (res) {
      if (!res.ok) return readErrorBody(res);
      return res.json();
    }).then(function (data) {
      filterField = field;
      filterValues = data.values || [];
      if (filterValues.indexOf('') === -1) filterValues.unshift('');
      var selected = articleState.filters[field];
      filterDraft = {};
      filterValues.forEach(function (value) {
        filterDraft[draftKey(value)] = !selected || selected.indexOf(value) !== -1;
      });
      filterEl.search.value = '';
      renderFilterList();
      var rect = btn.getBoundingClientRect();
      var width = 280;
      var left = Math.max(8, Math.min(rect.left, window.innerWidth - width - 8));
      filterEl.box.hidden = false;
      filterEl.box.style.left = left + 'px';
      filterEl.box.style.top = (rect.bottom + 4) + 'px';
      filterEl.search.focus();
    }).catch(function (err) {
      showArticleError(err.message || String(err));
    });
  }

  function loadArticles(page) {
    articleState.page = page || 1;
    clearArticleError();
    fetch(API + '/api/articles/query', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        q: articleState.q,
        page: articleState.page,
        per_page: 40,
        filters: articleState.filters,
      }),
    }).then(function (res) {
      if (!res.ok) return readErrorBody(res);
      return res.json();
    }).then(function (data) {
      renderArticleTable(data);
      markActiveFilters();
    }).catch(function (e) { showArticleError(e.message || String(e)); });
  }

  document.querySelectorAll('.article-filter-btn').forEach(function (btn) {
    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      openColumnFilter(btn);
    });
  });
  var resetFiltersBtn = document.getElementById('articleResetFilters');
  if (resetFiltersBtn) {
    resetFiltersBtn.addEventListener('click', function () {
      articleState.filters = {};
      articleState.q = '';
      if (articleEl.search) articleEl.search.value = '';
      closeColumnFilter();
      loadArticles(1);
    });
  }
  if (filterEl.search) {
    filterEl.search.addEventListener('input', selectFilterMatches);
    filterEl.search.addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      commitColumnFilter();
      loadArticles(1);
    });
  }
  if (filterEl.all) {
    filterEl.all.addEventListener('change', function () {
      visibleFilterInputs().forEach(function (input) {
        input.checked = filterEl.all.checked;
        filterDraft[draftKey(input._filterValue)] = filterEl.all.checked;
      });
      filterEl.all.indeterminate = false;
    });
  }
  if (filterEl.ok) {
    filterEl.ok.addEventListener('click', function () {
      commitColumnFilter();
      loadArticles(1);
    });
  }
  if (filterEl.cancel) filterEl.cancel.addEventListener('click', closeColumnFilter);
  document.addEventListener('click', function (e) {
    if (!filterEl.box || filterEl.box.hidden) return;
    if (filterEl.box.contains(e.target)) return;
    closeColumnFilter();
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeColumnFilter();
  });

  if (articleEl.searchBtn) {
    articleEl.searchBtn.addEventListener('click', function () {
      if (filterField) commitColumnFilter();
      articleState.q = (articleEl.search.value || '').trim();
      loadArticles(1);
    });
  }
  if (articleEl.search) {
    articleEl.search.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        articleState.q = articleEl.search.value.trim();
        loadArticles(1);
      }
    });
    articleEl.search.addEventListener('input', function () {
      if ((articleEl.search.value || '').trim()) return;
      if (!articleState.q) return;
      articleState.q = '';
      loadArticles(1);
    });
  }
  if (articleEl.addBtn) {
    articleEl.addBtn.addEventListener('click', function () { openArticleForm(null); });
  }
  if (articleEl.cancel) {
    articleEl.cancel.addEventListener('click', closeArticleForm);
  }
  if (articleEl.form) {
    articleEl.form.addEventListener('submit', function (e) {
      e.preventDefault();
      clearFormError();
      var payload = {};
      articleEl.fields.querySelectorAll('input, textarea').forEach(function (input) {
        payload[input.name] = input.value.toUpperCase();
      });
      var url = articleEditingId ? (API + '/api/articles/' + articleEditingId) : (API + '/api/articles');
      fetch(url, {
        method: articleEditingId ? 'PUT' : 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      }).then(function (res) {
        if (!res.ok) return readErrorBody(res);
        return res.json();
      }).then(function () {
        closeArticleForm();
        loadArticles(articleState.page);
      }).catch(function (err) {
        showFormError(err.message || String(err));
      });
    });
  }
  bindColumnResize();
  document.querySelectorAll('.tab').forEach(function (tab) {
    tab.addEventListener('click', function () {
      if (tab.dataset.tab === 'articles') loadArticles(articleState.page || 1);
    });
  });

  // ---- AI Analyze (YandexGPT column mapping) ----
  var aiAnalyzeBtn = document.getElementById('aiAnalyzeBtn');
  if (aiAnalyzeBtn) {
    aiAnalyzeBtn.addEventListener('click', function () {
      clearError();
      if (!inputFiles.length) {
        showError('Сначала загрузите входные файлы.');
        return;
      }
      aiAnalyzeBtn.disabled = true;
      var original = aiAnalyzeBtn.textContent;
      aiAnalyzeBtn.textContent = 'ИИ анализирует…';
      setProgress(10);

      var progressTimer = setInterval(function () {
        var current = parseInt(el.progressLabel.textContent, 10) || 0;
        if (current < 90) setProgress(current + 5);
      }, 700);

      var fd = new FormData();
      inputFiles.forEach(function (f) { fd.append('inputs', f); });
      fd.append('country', el.countryInput.value.trim() || 'CN');
      fd.append('unit', el.unitInput.value.trim() || 'шт');

      var controller = new AbortController();
      var timeoutId = setTimeout(function () { controller.abort(); }, 120000);

      fetch(API + '/api/ai/analyze', { method: 'POST', body: fd, signal: controller.signal })
        .then(function (res) {
          clearTimeout(timeoutId);
          clearInterval(progressTimer);
          setProgress(100);
          if (!res.ok) return readErrorBody(res);
          return readResultPayload(res);
        })
        .then(function (payload) {
          triggerDownload(payload.blob, payload.filename || 'result.xlsx');
          showResult(payload.blob, payload.filename || 'result.xlsx');
          showErrorLog(payload.protocol);
        })
        .catch(function (e) {
          clearTimeout(timeoutId);
          clearInterval(progressTimer);
          var msg = e && e.name === 'AbortError'
            ? 'ИИ-анализ превысил время ожидания. Попробуйте ещё раз или используйте «Скачать готовый файл».'
            : ('ИИ-анализ не удался: ' + (e && e.message ? e.message : e));
          showError(msg);
        })
        .finally(function () {
          hideProgress();
          aiAnalyzeBtn.textContent = original;
          aiAnalyzeBtn.disabled = false;
        });
    });
  }

  // ---- Инициализация ----
  if (el.compareOpenBtn) {
    el.compareOpenBtn.addEventListener('click', function () { openCompare(null); });
  }
  if (el.compareCloseBtn) {
    el.compareCloseBtn.addEventListener('click', closeCompare);
  }
  if (el.compareModal) {
    el.compareModal.addEventListener('click', function (e) {
      if (e.target === el.compareModal) closeCompare();
    });
  }
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeCompare();
  });

  loadTemplateInfo();
})();
