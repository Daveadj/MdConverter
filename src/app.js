/* =============================================================================
 * app.js — UI glue. Owns the queue, the previews and the downloads; knows
 * nothing about file formats, which all live in to-md.js / from-md.js.
 * ========================================================================== */
(function () {
  'use strict';

  const $ = sel => document.querySelector(sel);
  const $$ = sel => Array.from(document.querySelectorAll(sel));

  const state = {
    items: [],          // { id, file, name, status, markdown, warnings, error }
    selected: null,
    view: 'rendered',
    seq: 0
  };

  /* ------------------------------------------------------------ chrome --- */

  const THEME_KEY = 'mdconv.theme';
  try {
    const saved = localStorage.getItem(THEME_KEY);
    const dark = saved ? saved === 'dark'
                       : matchMedia('(prefers-color-scheme: dark)').matches;
    document.documentElement.dataset.theme = dark ? 'dark' : 'light';
  } catch (e) { /* storage blocked — light theme is a fine default */ }

  $('#themeToggle').addEventListener('click', () => {
    const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    try { localStorage.setItem(THEME_KEY, next); } catch (e) {}
  });

  $$('.tab').forEach(tab => tab.addEventListener('click', () => {
    $$('.tab').forEach(t => { t.classList.remove('is-active'); t.setAttribute('aria-selected', 'false'); });
    tab.classList.add('is-active');
    tab.setAttribute('aria-selected', 'true');
    $$('.panel').forEach(p => p.classList.remove('is-active'));
    $('#panel-' + tab.dataset.tab).classList.add('is-active');
  }));

  let toastTimer = null;
  function toast(message, isError) {
    const el = $('#toast');
    el.textContent = message;
    el.classList.toggle('is-err', !!isError);
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { el.hidden = true; }, isError ? 6000 : 3000);
  }

  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  const baseName = name => name.replace(/\.[^.]+$/, '') || 'document';

  // Without this the browser opens a dropped PDF instead of handing it over.
  ['dragover', 'drop'].forEach(evt =>
    document.addEventListener(evt, e => e.preventDefault()));

  function wireDropzone(zone, input, onFiles) {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
    });
    zone.addEventListener('dragenter', () => zone.classList.add('is-over'));
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('is-over'); });
    zone.addEventListener('dragleave', e => {
      if (!zone.contains(e.relatedTarget)) zone.classList.remove('is-over');
    });
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('is-over');
      onFiles(Array.from(e.dataTransfer.files || []));
    });
    input.addEventListener('change', () => {
      onFiles(Array.from(input.files || []));
      input.value = '';
    });
  }

  /* ================================================== TAB 1: to Markdown === */

  function readOptions() {
    return {
      embedImages:      $('#optEmbedImages').checked,
      pageBreaks:       $('#optPageBreaks').checked,
      detectTables:     $('#optDetectTables').checked,
      detectHeadings:   $('#optDetectHeadings').checked,
      headingPerSection:$('#optSheetPerHeading').checked
    };
  }

  wireDropzone($('#dropIn'), $('#fileIn'), addFiles);

  async function addFiles(files) {
    if (!files.length) return;
    const fresh = files.map(file => ({
      id: ++state.seq,
      file,
      name: file.name,
      status: 'queued',
      markdown: '',
      warnings: [],
      error: null,
      note: ''
    }));
    state.items.push(...fresh);
    $('#queueIn').hidden = false;
    renderQueue();
    if (!state.selected) select(fresh[0].id);

    for (const item of fresh) await runConversion(item);
    updateZipButton();
  }

  async function runConversion(item) {
    if (!ToMd.supports(item.name)) {
      item.status = 'error';
      item.error = 'Unsupported file type (.' + (ToMd.extOf(item.name) || '?') + ')';
      renderQueue();
      return;
    }
    item.status = 'working';
    item.note = '';
    renderQueue();

    try {
      const result = await ToMd.convert(item.file, readOptions(), (page, total) => {
        item.note = `page ${page} of ${total}`;
        renderQueue();
      });
      item.markdown = result.markdown;
      item.warnings = result.warnings || [];
      item.status = 'done';
      item.note = describe(result.markdown);
    } catch (err) {
      item.status = 'error';
      item.error = err && err.message ? err.message : String(err);
      console.error(item.name, err);
    }
    renderQueue();
    if (state.selected === item.id) renderPreviewIn();
  }

  function describe(md) {
    const kb = Math.max(1, Math.round(new Blob([md]).size / 1024));
    const words = (md.match(/\S+/g) || []).length;
    return `${words.toLocaleString()} words · ${kb} KB`;
  }

  function renderQueue() {
    const list = $('#listIn');
    list.textContent = '';

    state.items.forEach(item => {
      const li = document.createElement('li');
      li.className = 'file' + (state.selected === item.id ? ' is-selected' : '');
      li.tabIndex = 0;

      const ext = document.createElement('span');
      ext.className = 'file-ext';
      ext.textContent = ToMd.extOf(item.name) || '?';

      const main = document.createElement('div');
      main.className = 'file-main';
      const name = document.createElement('div');
      name.className = 'file-name';
      name.textContent = item.name;
      name.title = item.name;
      const note = document.createElement('div');
      note.className = 'file-note';
      if (item.status === 'error') { note.classList.add('is-err'); note.textContent = item.error; }
      else if (item.status === 'working') note.textContent = item.note || 'converting…';
      else if (item.status === 'queued') note.textContent = 'queued';
      else {
        note.textContent = item.note;
        if (item.warnings.length) { note.classList.add('is-warn'); note.textContent = item.warnings[0]; }
      }
      main.append(name, note);

      const side = document.createElement('div');
      side.className = 'file-side';
      if (item.status === 'working') {
        const sp = document.createElement('span');
        sp.className = 'spinner';
        side.append(sp);
      } else if (item.status === 'done') {
        const btn = document.createElement('button');
        btn.className = 'btn-mini';
        btn.textContent = 'Save .md';
        btn.addEventListener('click', e => {
          e.stopPropagation();
          download(new Blob([item.markdown], { type: 'text/markdown;charset=utf-8' }),
                   baseName(item.name) + '.md');
        });
        const tick = document.createElement('span');
        tick.className = 'tick';
        tick.textContent = '✓';
        side.append(btn, tick);
      } else if (item.status === 'error') {
        const cross = document.createElement('span');
        cross.className = 'cross';
        cross.textContent = '!';
        side.append(cross);
      }

      li.append(ext, main, side);
      li.addEventListener('click', () => select(item.id));
      li.addEventListener('keydown', e => { if (e.key === 'Enter') select(item.id); });
      list.append(li);
    });
  }

  function select(id) {
    state.selected = id;
    renderQueue();
    renderPreviewIn();
  }

  function renderPreviewIn() {
    const box = $('#previewIn');
    const item = state.items.find(i => i.id === state.selected);
    box.textContent = '';
    box.className = 'preview';

    if (!item || item.status !== 'done') {
      const p = document.createElement('p');
      p.className = 'empty';
      p.textContent = item && item.status === 'error'
        ? item.error
        : item ? 'Converting…' : 'Convert a file to see its Markdown here.';
      box.append(p);
      return;
    }

    // A spreadsheet can convert to megabytes of pipe table. Rendering all of
    // it would lock the tab for seconds; the download still gets everything.
    const { text, truncated } = clip(item.markdown);

    if (state.view === 'source') {
      const pre = document.createElement('pre');
      pre.className = 'raw';
      pre.textContent = text;
      box.append(pre);
    } else {
      box.classList.add('md');
      box.innerHTML = FromMd.renderPreview(text);
    }
    if (truncated) {
      const note = document.createElement('p');
      note.className = 'empty';
      note.textContent = `Preview truncated at ${PREVIEW_LIMIT.toLocaleString()} characters — ` +
                         `“Save .md” writes the full ${item.markdown.length.toLocaleString()}.`;
      box.append(note);
    }
  }

  const PREVIEW_LIMIT = 200000;
  function clip(md) {
    if (md.length <= PREVIEW_LIMIT) return { text: md, truncated: false };
    const cut = md.lastIndexOf('\n', PREVIEW_LIMIT);
    return { text: md.slice(0, cut > 0 ? cut : PREVIEW_LIMIT), truncated: true };
  }

  $$('#panel-to-md .seg-btn').forEach(btn => btn.addEventListener('click', () => {
    $$('#panel-to-md .seg-btn').forEach(b => b.classList.remove('is-active'));
    btn.classList.add('is-active');
    state.view = btn.dataset.view;
    renderPreviewIn();
  }));

  $('#clearIn').addEventListener('click', () => {
    state.items = [];
    state.selected = null;
    $('#queueIn').hidden = true;
    renderQueue();
    renderPreviewIn();
    updateZipButton();
  });

  function updateZipButton() {
    $('#zipIn').disabled = !state.items.some(i => i.status === 'done');
  }

  $('#zipIn').addEventListener('click', async () => {
    const done = state.items.filter(i => i.status === 'done');
    if (!done.length) return;
    const zip = new JSZip();
    const used = new Set();
    done.forEach(item => {
      let name = baseName(item.name) + '.md';
      let n = 2;
      while (used.has(name)) name = baseName(item.name) + '-' + n++ + '.md';
      used.add(name);
      zip.file(name, item.markdown);
    });
    download(await zip.generateAsync({ type: 'blob' }), 'markdown.zip');
    toast(`Saved ${done.length} file${done.length > 1 ? 's' : ''} as markdown.zip`);
  });

  /* ================================================ TAB 2: from Markdown === */

  const source = $('#mdSource');
  const titleInput = $('#docTitle');

  wireDropzone($('#dropOut'), $('#fileOut'), async files => {
    const texts = [];
    for (const f of files) texts.push(await f.text());
    source.value = texts.join('\n\n---\n\n');
    if (!titleInput.value && files[0]) titleInput.value = baseName(files[0].name);
    refreshOut();
  });

  function docOptions() {
    return {
      title: (titleInput.value || '').trim() || inferTitle() || 'Untitled document',
      paper: $('#optPaper').value,
      margin: parseFloat($('#optMargin').value),
      font: $('#optFont').value,
      pageNumbers: $('#optPageNums').checked
    };
  }

  function inferTitle() {
    const m = (source.value || '').match(/^\s*#\s+(.+)$/m);
    return m ? m[1].replace(/[#*`_]/g, '').trim() : '';
  }

  let outTimer = null;
  function refreshOut() {
    clearTimeout(outTimer);
    outTimer = setTimeout(() => {
      const box = $('#previewOut');
      const md = source.value.trim();
      if (!md) {
        box.className = 'preview preview-paper';
        box.innerHTML = '<p class="empty">Your document will render here as you type.</p>';
        return;
      }
      box.className = 'preview preview-paper md';
      box.innerHTML = FromMd.renderPreview(clip(source.value).text);
      if (titleInput.placeholder !== inferTitle()) titleInput.placeholder = inferTitle() || 'Untitled document';
    }, 200);
  }
  source.addEventListener('input', refreshOut);

  function requireMarkdown() {
    if (source.value.trim()) return true;
    toast('Nothing to convert — drop a .md file or type some Markdown.', true);
    return false;
  }

  async function withBusy(button, label, fn) {
    const original = button.textContent;
    button.disabled = true;
    button.textContent = label;
    try { await fn(); }
    catch (err) {
      console.error(err);
      toast(err && err.message ? err.message : String(err), true);
    }
    finally { button.disabled = false; button.textContent = original; }
  }

  $('#btnDocx').addEventListener('click', function () {
    if (!requireMarkdown()) return;
    const o = docOptions();
    withBusy(this, 'Building…', async () => {
      const blob = await FromMd.toDocxBlob(source.value, o);
      download(blob, o.title.replace(/[\\/:*?"<>|]/g, '-') + '.docx');
      toast('Saved ' + o.title + '.docx');
    });
  });

  $('#btnHtml').addEventListener('click', function () {
    if (!requireMarkdown()) return;
    const o = docOptions();
    const html = FromMd.toHtmlDocument(source.value, o);
    download(new Blob([html], { type: 'text/html;charset=utf-8' }),
             o.title.replace(/[\\/:*?"<>|]/g, '-') + '.html');
    toast('Saved ' + o.title + '.html');
  });

  $('#btnPdf').addEventListener('click', function () {
    if (!requireMarkdown()) return;
    const o = docOptions();
    withBusy(this, 'Opening…', async () => {
      await FromMd.printToPdf(source.value, o, $('#printFrame'));
    });
  });

  /* -------------------------------------------------------- startup ---- */

  // A CDN that 404s otherwise degrades silently — tables just quietly vanish.
  // Fail loudly instead.
  const REQUIRED = {
    'marked':                () => window.marked,
    'Turndown':              () => window.TurndownService,
    'Turndown GFM (tables)': () => window.turndownPluginGfm,
    'mammoth (.docx in)':    () => window.mammoth,
    'SheetJS (.xlsx)':       () => window.XLSX,
    'pdf.js (.pdf)':         () => window.pdfjsLib,
    'JSZip (.pptx/.epub)':   () => window.JSZip,
    'docx (.docx out)':      () => window.docx,
    'DOMPurify':             () => window.DOMPurify
  };
  const missing = Object.keys(REQUIRED).filter(k => !REQUIRED[k]());
  if (missing.length) {
    toast('Failed to load: ' + missing.join(', ') + ' — check your connection and reload.', true);
    console.error('Missing libraries:', missing);
  }

  refreshOut();
})();
