/* =============================================================================
 * to-md.js — anything -> Markdown, entirely in the browser.
 *
 * Every converter returns { markdown, warnings[] }. Nothing here touches the
 * DOM of the host page, so the whole module can be swapped for a server call
 * (MarkItDown/Pandoc) later without the UI noticing — see convert() at the end.
 * ========================================================================== */
(function (global) {
  'use strict';

  const PDF_WORKER =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  /* --------------------------------------------------------------- turndown */

  let _td = null;
  function turndown() {
    if (_td) return _td;
    _td = new TurndownService({
      headingStyle: 'atx',
      hr: '---',
      bulletListMarker: '-',
      codeBlockStyle: 'fenced',
      emDelimiter: '*',
      linkStyle: 'inlined'
    });
    if (global.turndownPluginGfm) _td.use(global.turndownPluginGfm.gfm);

    // Word and EPUB both emit <a name="..."></a> bookmarks that turn into
    // stray empty links. Drop anchors that carry neither href nor text.
    _td.addRule('emptyAnchor', {
      filter: n => n.nodeName === 'A' && !n.getAttribute('href') && !n.textContent.trim(),
      replacement: () => ''
    });

    // Turndown's stock list item is "-   item" with a 4-space continuation.
    // Valid, but every other tool writes "- item"; match that, and indent
    // nested content by 2 to stay consistent with the shorter marker.
    _td.addRule('tightListItem', {
      filter: 'li',
      replacement: function (content, node, options) {
        content = content.replace(/^\n+/, '').replace(/\n+$/, '\n').replace(/\n/gm, '\n  ');
        let prefix = options.bulletListMarker + ' ';
        const parent = node.parentNode;
        if (parent.nodeName === 'OL') {
          const start = parent.getAttribute('start');
          const index = Array.prototype.indexOf.call(parent.children, node);
          prefix = (start ? Number(start) + index : index + 1) + '. ';
        }
        return prefix + content + (node.nextSibling && !/\n$/.test(content) ? '\n' : '');
      }
    });

    // A pipe table row must be one physical line. The GFM plugin's own cell
    // rule passes newlines straight through, which shatters any cell holding
    // more than a bare string; addRule takes precedence, so this replaces it.
    _td.addRule('singleLineTableCell', {
      filter: ['th', 'td'],
      replacement: function (content, node) {
        const text = content
          .replace(/\|/g, '\\|')
          .replace(/\s*\n+\s*/g, '<br>')
          .replace(/(<br>)+$/, '')
          .trim();
        return (node.previousElementSibling ? ' ' : '| ') + (text || ' ') + ' |';
      }
    });
    return _td;
  }

  /**
   * "1." -> level 2, "2.1" -> level 3, "2.1.4" -> level 4. Numbered section
   * labels are the most reliable depth signal in documents that never used
   * real heading styles.
   */
  function headingLevelFromNumber(text) {
    const m = String(text).match(/^(\d+(?:\.\d+)*)\s*[.)]?\s+\S/);
    return m ? Math.min(6, 1 + m[1].split('.').length) : 0;
  }

  const CELL_BLOCKS = 'p, div, li, h1, h2, h3, h4, h5, h6, pre, blockquote';

  /**
   * Three things have to be true before turndown will emit a GFM table:
   * the first row must be a heading row, cells must not contain block
   * elements, and rows must not contain stray whitespace text nodes. Word
   * output violates all three — every cell is <td><p>…</p></td> — which is
   * why an unprepared table collapses into a run of loose paragraphs.
   */
  function normalizeTables(doc) {
    doc.querySelectorAll('tr').forEach(tr => {
      Array.from(tr.childNodes).forEach(n => {
        if (n.nodeType === 3 && !n.textContent.trim()) n.remove();
      });
    });

    doc.querySelectorAll('td, th').forEach(cell => {
      const blocks = Array.from(cell.querySelectorAll(CELL_BLOCKS))
        .filter(b => !b.querySelector(CELL_BLOCKS));       // innermost only
      if (!blocks.length) return;
      const html = blocks.map(b => b.innerHTML.trim()).filter(Boolean).join('<br>');
      cell.innerHTML = html || cell.textContent.trim();
    });

    doc.querySelectorAll('table').forEach(table => {
      if (table.querySelector('th')) return;
      const row = table.querySelector('tr');
      if (!row) return;
      Array.from(row.children).forEach(cell => {
        if (cell.nodeName !== 'TD') return;
        const th = doc.createElement('th');
        th.innerHTML = cell.innerHTML;
        cell.replaceWith(th);
      });
      const thead = doc.createElement('thead');
      table.insertBefore(thead, table.firstChild);
      thead.appendChild(row);
    });
  }

  /**
   * Plenty of real documents — CCF's own payment framework among them — never
   * apply Word heading styles and just bold the section titles. A paragraph
   * that is entirely bold, short, and not a sentence is one of those.
   */
  function promoteBoldHeadings(doc) {
    const state = { seen: false };
    Array.from(doc.querySelectorAll('p')).forEach(p => {
      // A bold cell is a table header and a bold list item is emphasis —
      // neither is a section heading, and hoisting them wrecks the structure.
      if (p.closest('td, th, li, blockquote, pre')) return;
      const text = p.textContent.replace(/\s+/g, ' ').trim();
      if (!text || text.length > 100 || /[.:;,!?]$/.test(text)) return;
      if (p.querySelector('img, table, a[href]')) return;
      const bolded = Array.from(p.querySelectorAll('strong, b')).map(n => n.textContent).join('');
      if (!bolded || bolded.replace(/\s+/g, '') !== text.replace(/\s+/g, '')) return;

      const level = headingLevelFromNumber(text) || (state.seen ? 2 : 1);
      state.seen = true;
      const h = doc.createElement('h' + Math.min(level, 6));
      h.textContent = text;
      p.replaceWith(h);
    });
  }

  /**
   * Rejoin the run of one-line <pre> blocks a code paragraph style produces,
   * then give each a <code> child — turndown only fences "pre > code", and a
   * bare <pre> would have its whitespace collapsed away entirely.
   */
  function normalizePres(doc) {
    doc.querySelectorAll('pre').forEach(pre => {
      if (!pre.isConnected) return;
      const lines = [pre.textContent];
      let next = pre.nextElementSibling;
      while (next && next.nodeName === 'PRE') {
        const after = next.nextElementSibling;
        lines.push(next.textContent);
        next.remove();
        next = after;
      }
      const code = doc.createElement('code');
      code.textContent = lines.join('\n').replace(/[ \t]+$/gm, '');
      pre.textContent = '';
      pre.appendChild(code);
    });
  }

  function htmlToMarkdown(html, opts) {
    opts = opts || {};
    const doc = new DOMParser().parseFromString(html, 'text/html');
    doc.querySelectorAll('script, style, noscript, meta, link').forEach(n => n.remove());
    if (!opts.embedImages) doc.querySelectorAll('img').forEach(n => n.remove());
    normalizePres(doc);
    normalizeTables(doc);
    if (opts.detectHeadings !== false) promoteBoldHeadings(doc);
    return tidy(turndown().turndown(doc.body ? doc.body.innerHTML : html));
  }

  /* ----------------------------------------------------------------- shared */

  function tidy(md) {
    return md
      .replace(/\r\n?/g, '\n')
      .replace(/[ \t]+$/gm, '')
      // Turndown escapes "1." so it cannot be read as a list, but inside a
      // heading there is no list to confuse and "## 1\. Overview" just looks wrong.
      .replace(/^#{1,6} .*$/gm, line => line.replace(/\\([.)\-+*_])/g, '$1'))
      .replace(/\n{3,}/g, '\n\n')
      .replace(/^\n+/, '')
      .trimEnd() + '\n';
  }

  const escapeCell = v => String(v == null ? '' : v)
    .replace(/\|/g, '\\|')
    .replace(/\n+/g, ' ')
    .trim();

  /** Rows of raw values -> a GFM pipe table. First row is the header. */
  function toPipeTable(rows) {
    if (!rows.length) return '';
    const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
    if (width < 2) return rows.map(r => escapeCell(r[0])).join('\n\n');

    const pad = r => { const c = r.map(escapeCell); while (c.length < width) c.push(''); return c; };
    let head = pad(rows[0]);
    let body = rows.slice(1);

    // A blank leading row makes an unreadable table; borrow the first row that
    // actually has content as the header instead.
    if (head.every(c => !c) && body.length) { head = pad(body[0]); body = body.slice(1); }
    if (head.every(c => !c)) head = head.map((_, i) => 'Column ' + (i + 1));

    const out = ['| ' + head.join(' | ') + ' |',
                 '| ' + head.map(() => '---').join(' | ') + ' |'];
    body.forEach(r => out.push('| ' + pad(r).join(' | ') + ' |'));
    return out.join('\n');
  }

  /**
   * Drop every column and row that is empty everywhere. Trimming only the
   * trailing columns is not enough: real workbooks use blank spacer columns
   * and have a used range far wider than their data, which on one of CCF's
   * own management accounts meant 225 columns that were 95% empty.
   */
  function trimGrid(rows) {
    const val = v => String(v ?? '').trim();
    const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
    const keep = [];
    for (let c = 0; c < width; c++) {
      if (rows.some(r => val(r[c]))) keep.push(c);
    }
    if (!keep.length) return [];
    return rows.map(r => keep.map(c => r[c] ?? ''))
               .filter(r => r.some(v => val(v)));
  }

  const readAsArrayBuffer = file => file.arrayBuffer();
  const readAsText = file => file.text();

  /* -------------------------------------------------------------------- PDF */

  /**
   * pdf.js hands back positioned text runs, not paragraphs. Everything below
   * rebuilds document structure from geometry: baselines become lines, font
   * size becomes heading level, vertical gaps become paragraph breaks, and
   * horizontal gaps become table columns.
   */
  async function pdfToMarkdown(buf, opts, onProgress) {
    if (!global.pdfjsLib) throw new Error('pdf.js failed to load');
    pdfjsLib.GlobalWorkerOptions.workerSrc = PDF_WORKER;

    const doc = await pdfjsLib.getDocument({
      data: new Uint8Array(buf),
      isEvalSupported: false,
      useSystemFonts: true
    }).promise;

    const warnings = [];
    const perPage = [];

    for (let p = 1; p <= doc.numPages; p++) {
      if (onProgress) onProgress(p, doc.numPages);
      const page = await doc.getPage(p);
      const content = await page.getTextContent();
      perPage.push(normalizeItems(content.items));
      page.cleanup();
    }
    doc.destroy();

    // Body size and body font are decided across the whole document, not per
    // page — a page that happens to be all heading would otherwise redefine
    // what "normal text" means and flatten its own headings away.
    const all = perPage.flat();
    const bodySize = modalBy(all, it => Math.round(it.size * 2) / 2) || 10;
    const bodyFont = modalBy(all, it => it.font);
    const docState = { seenHeading: false };

    const pages = perPage.map(items => pageToMarkdown(items, opts, bodySize, bodyFont, docState));
    const sawText = pages.some(md => md.trim());

    if (!sawText) {
      warnings.push('No selectable text found — this looks like a scanned PDF. ' +
                    'It needs OCR, which this build does not do.');
    }
    const joiner = opts.pageBreaks ? '\n\n---\n\n' : '\n\n';
    return { markdown: tidy(pages.filter(s => s.trim()).join(joiner)), warnings };
  }

  function normalizeItems(rawItems) {
    const items = [];
    for (const it of rawItems) {
      if (!it.str || !it.str.trim()) continue;
      const t = it.transform;
      items.push({
        str: it.str,
        x: t[4],
        y: t[5],
        w: it.width || 0,
        size: Math.abs(t[3]) || it.height || 10,
        font: it.fontName || ''
      });
    }
    return items;
  }

  function pageToMarkdown(items, opts, bodySize, bodyFont, docState) {
    if (!items.length) return '';
    const lines = groupIntoLines(items, bodySize);
    if (!lines.length) return '';

    const gap = typicalBaselineGap(lines);
    const rows = lines.map(l => renderLine(l, bodySize, bodyFont, opts, docState));
    const blocks = opts.detectTables ? extractTables(rows) : rows.map(r => ({ kind: 'line', row: r }));

    return assemble(blocks, gap);
  }

  /** The value most characters in the document carry — body size, or body font. */
  function modalBy(items, keyOf) {
    const weight = new Map();
    for (const it of items) {
      const key = keyOf(it);
      weight.set(key, (weight.get(key) || 0) + it.str.length);
    }
    let best = null, most = -1;
    weight.forEach((n, key) => { if (n > most) { most = n; best = key; } });
    return best;
  }

  function groupIntoLines(items, bodySize) {
    const tol = Math.max(1.2, bodySize * 0.45);
    const sorted = items.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
    const lines = [];
    for (const it of sorted) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(last.y - it.y) <= tol) last.items.push(it);
      else lines.push({ y: it.y, items: [it] });
    }
    lines.forEach(l => {
      l.items.sort((a, b) => a.x - b.x);
      l.size = Math.max(...l.items.map(i => i.size));
      l.x = l.items[0].x;
    });
    return lines;
  }

  /**
   * Stitch one line's runs together. A gap wider than roughly a character
   * becomes a space; a gap wider than a whole em becomes a column boundary,
   * recorded as \t so the table pass can see it.
   */
  const MARKER_ONLY = /^\s*([•●○▪◦‣·⁃]|\d+[.)])\s*$/;

  function renderLine(line, bodySize, bodyFont, opts, docState) {
    const cells = [];
    let cur = { text: '', x: line.items[0].x };

    line.items.forEach((it, i) => {
      if (i > 0) {
        const prev = line.items[i - 1];
        const gap = it.x - (prev.x + prev.w);
        const em = prev.size || bodySize;
        if (gap > em * 1.1) { cells.push(cur); cur = { text: '', x: it.x }; }
        else if (gap > em * 0.18 && !/\s$/.test(cur.text) && !/^\s/.test(it.str)) cur.text += ' ';
      }
      cur.text += it.str;
    });
    cells.push(cur);
    cells.forEach(c => { c.text = c.text.replace(/\s+$/, ''); });

    // Word writes a list bullet as its own text run, far to the left of the
    // item text. That wide gap looks exactly like a column boundary, so every
    // bulleted list was being turned into a two-column table. Re-attach it.
    if (cells.length > 1 && MARKER_ONLY.test(cells[0].text)) {
      cells[1].text = cells[0].text.trim() + ' ' + cells[1].text;
      cells.shift();
    }

    const text = cells.map(c => c.text).join('\t');
    const flat = text.replace(/\t/g, ' ').replace(/\s+/g, ' ').trim();
    const ratio = line.size / bodySize;

    // Size first. When a document sets every heading at body size — common in
    // anything exported from Word — the only remaining signal is that the
    // heading is drawn with a different embedded font resource than the body.
    let heading = 0;
    if (ratio >= 1.55) heading = 1;
    else if (ratio >= 1.3) heading = 2;
    else if (ratio >= 1.13) heading = 3;
    else if (opts.detectHeadings !== false && bodyFont &&
             line.items.every(i => i.font !== bodyFont) &&
             flat.length > 0 && flat.length <= 100 && !/[.;,:]$/.test(flat)) {
      heading = headingLevelFromNumber(flat) || (docState.seenHeading ? 2 : 1);
    }
    if (heading) docState.seenHeading = true;

    return { cells, text, flat, heading, y: line.y, x: line.x, size: line.size, items: line.items };
  }

  function typicalBaselineGap(lines) {
    const deltas = [];
    for (let i = 1; i < lines.length; i++) {
      const d = lines[i - 1].y - lines[i].y;
      if (d > 0.5) deltas.push(d);
    }
    if (!deltas.length) return 12;
    deltas.sort((a, b) => a - b);
    return deltas[Math.floor(deltas.length / 2)];
  }

  /**
   * Conservative table detection: a run of consecutive lines that all split
   * into the same number of columns (2+) is almost always a table. Anything
   * less certain is left as prose — a wrong table is worse than no table.
   */
  function extractTables(rows) {
    const blocks = [];
    let i = 0;
    while (i < rows.length) {
      const width = rows[i].cells.length;
      if (width >= 2) {
        let j = i + 1;
        while (j < rows.length && rows[j].cells.length === width) j++;

        if (j - i >= 3) {                       // header + 2 body rows minimum
          const run = rows.slice(i, j);
          // A leading column of nothing but bullet glyphs is an indented list,
          // not a table column.
          if (!run.every(r => /^\s*[•●○▪◦‣·⁃*-]?\s*$/.test(r.cells[0].text))) {
            const grid = run.map(r => r.cells.map(c => c.text));

            // A header row often spans its columns so widely that the gaps
            // between its labels fall below the column threshold, leaving it as
            // one undivided line above the table. Re-cut it against the columns
            // the data rows establish, rather than losing the first data row to
            // the header position.
            const last = blocks[blocks.length - 1];
            const above = rows[i - 1];
            if (above && above.cells.length < width && last && last.kind === 'line' && last.row === above) {
              const header = splitAtColumns(above.items, columnStarts(run));
              if (header && header.filter(c => c.trim()).length >= 2) {
                grid.unshift(header);
                blocks.pop();                   // un-emit it as a standalone line
              }
            }
            blocks.push({ kind: 'table', rows: grid });
            i = j;
            continue;
          }
        }
      }
      blocks.push({ kind: 'line', row: rows[i] });
      i++;
    }
    return blocks;
  }

  /** Left edge of each column, taken as the leftmost cell start across rows. */
  function columnStarts(run) {
    const starts = [];
    run.forEach(row => row.cells.forEach((c, k) => {
      starts[k] = starts[k] === undefined ? c.x : Math.min(starts[k], c.x);
    }));
    return starts;
  }

  /** Assign each text run of a line to the column its x position falls in. */
  function splitAtColumns(items, starts) {
    if (!items || !starts.length) return null;
    const cells = starts.map(() => '');
    for (const it of items) {
      let k = 0;
      for (let c = starts.length - 1; c >= 0; c--) {
        if (it.x >= starts[c] - 2) { k = c; break; }
      }
      cells[k] += (cells[k] && !/\s$/.test(cells[k]) && !/^\s/.test(it.str) ? ' ' : '') + it.str;
    }
    return cells.map(c => c.trim());
  }

  const BULLET = /^\s*[•●○▪◦‣·⁃*]\s+/;
  const DASH_BULLET = /^\s*[-–—]\s+/;
  const NUMBERED = /^\s*\d+[.)]\s+/;

  function assemble(blocks, gap) {
    const out = [];                 // { list: bool, text: string }
    let para = [];
    let prev = null;

    const push = (text, list) => out.push({ text, list: !!list });
    const flushPara = () => { if (para.length) { push(para.join(' ')); para = []; } };

    for (const block of blocks) {
      if (block.kind === 'table') {
        flushPara();
        push(toPipeTable(block.rows));
        prev = null;
        continue;
      }

      const row = block.row;
      const text = row.flat;
      if (!text) continue;

      if (row.heading) {
        flushPara();
        push('#'.repeat(row.heading) + ' ' + text);
        prev = row;
        continue;
      }

      let listed = null;
      if (BULLET.test(text) || DASH_BULLET.test(text)) listed = '- ' + text.replace(BULLET, '').replace(DASH_BULLET, '');
      else if (NUMBERED.test(text)) listed = text.trim();

      if (listed) { flushPara(); push(listed, true); prev = row; continue; }

      // Prose: a normal baseline step means the sentence wrapped, a bigger one
      // means a new paragraph. Rejoin wrapped lines so the Markdown reflows.
      const newPara = !prev || prev.heading || (prev.y - row.y) > gap * 1.4;
      if (newPara) flushPara();

      if (para.length && /­$|(?:\w)-$/.test(para[para.length - 1])) {
        para[para.length - 1] = para[para.length - 1].replace(/[­-]$/, '') + text;
      } else {
        para.push(text);
      }
      prev = row;
    }
    flushPara();

    // Adjacent list items are a tight list — one newline. Everything else is
    // separated by a blank line.
    return out.reduce((acc, block, i) => {
      if (i === 0) return block.text;
      const tight = block.list && out[i - 1].list;
      return acc + (tight ? '\n' : '\n\n') + block.text;
    }, '');
  }

  /* ------------------------------------------------------------------- DOCX */

  // Word has no code-block or blockquote element — both are paragraph styles.
  // Mapping the usual style names (including the one this app writes on the
  // way out) is what makes a Markdown -> DOCX -> Markdown trip come back whole.
  // mammoth 1.7's browser build rejects the ":separator:" form, so each code
  // line arrives as its own <pre>; normalizePres() stitches them back together.
  const DOCX_STYLE_MAP = [
    "p[style-name='Md Code'] => pre:fresh",
    "p[style-name='Code'] => pre:fresh",
    "p[style-name='Source Code'] => pre:fresh",
    "p[style-name='HTML Preformatted'] => pre:fresh",
    "p[style-name='Quote'] => blockquote > p:fresh",
    "p[style-name='Intense Quote'] => blockquote > p:fresh"
  ];

  async function docxToMarkdown(buf, opts) {
    if (!global.mammoth) throw new Error('mammoth.js failed to load');
    const result = await mammoth.convertToHtml(
      { arrayBuffer: buf },
      { styleMap: DOCX_STYLE_MAP, includeDefaultStyleMap: true });
    const warnings = (result.messages || [])
      .filter(m => m.type === 'warning' || m.type === 'error')
      .map(m => m.message)
      .slice(0, 4);
    return { markdown: htmlToMarkdown(result.value, opts), warnings };
  }

  /* ----------------------------------------------------------- XLSX and CSV */

  function workbookToMarkdown(wb, opts) {
    const parts = [];
    const warnings = [];
    let widest = { name: '', width: 0 };

    wb.SheetNames.forEach(name => {
      const sheet = wb.Sheets[name];
      const rows = trimGrid(
        XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: '', blankrows: false })
      );
      if (!rows.length) return;
      const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
      if (width > widest.width) widest = { name, width };
      if (opts.headingPerSection && wb.SheetNames.length > 1) parts.push('## ' + name);
      parts.push(toPipeTable(rows));
    });

    if (!parts.length) warnings.push('Every sheet was empty.');
    // A pipe table can only represent a grid of values. A sheet laid out as a
    // formatted report — merged cells, spacer columns, sub-totals positioned by
    // hand — survives as data but not as a readable table, and saying so is
    // more useful than quietly producing something unusable.
    else if (widest.width > 40) {
      warnings.push(`Sheet "${widest.name}" is ${widest.width} columns wide — that is a ` +
                    `formatted report layout, not a data table, so its Markdown will be unwieldy.`);
    }
    return { markdown: tidy(parts.join('\n\n')), warnings };
  }

  async function xlsxToMarkdown(buf, opts) {
    if (!global.XLSX) throw new Error('SheetJS failed to load');
    return workbookToMarkdown(XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true }), opts);
  }

  async function csvToMarkdown(text, opts) {
    // Reuse SheetJS rather than a hand-rolled split, so quoted commas,
    // embedded newlines and BOMs behave.
    const wb = XLSX.read(text, { type: 'string', raw: false });
    return workbookToMarkdown(wb, Object.assign({}, opts, { headingPerSection: false }));
  }

  /* ------------------------------------------------------------------- PPTX */

  async function pptxToMarkdown(buf, opts) {
    const zip = await JSZip.loadAsync(buf);
    const slides = Object.keys(zip.files)
      .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
      .sort((a, b) => slideNo(a) - slideNo(b));

    if (!slides.length) return { markdown: '', warnings: ['No slides found in this .pptx.'] };

    const parts = [];
    for (let i = 0; i < slides.length; i++) {
      const xml = await zip.file(slides[i]).async('string');
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const paras = Array.from(doc.getElementsByTagNameNS(A_NS, 'p'))
        .map(p => ({
          text: Array.from(p.getElementsByTagNameNS(A_NS, 't')).map(t => t.textContent).join('').trim(),
          level: indentLevel(p)
        }))
        .filter(p => p.text);

      if (!paras.length) continue;
      const body = [];
      if (opts.headingPerSection) {
        body.push('## ' + paras[0].text);
        paras.slice(1).forEach(p => body.push('  '.repeat(p.level) + '- ' + p.text));
      } else {
        paras.forEach(p => body.push('  '.repeat(p.level) + '- ' + p.text));
      }
      parts.push(body.join('\n'));
    }
    return { markdown: tidy(parts.join('\n\n')), warnings: [] };
  }

  const slideNo = n => parseInt((n.match(/(\d+)\.xml$/) || [0, 0])[1], 10);

  function indentLevel(p) {
    const pr = p.getElementsByTagNameNS(A_NS, 'pPr')[0];
    const lvl = pr && pr.getAttribute('lvl');
    return Math.min(parseInt(lvl || '0', 10) || 0, 4);
  }

  /* ------------------------------------------------------------------- EPUB */

  async function epubToMarkdown(buf, opts) {
    const zip = await JSZip.loadAsync(buf);
    const containerFile = zip.file('META-INF/container.xml');
    if (!containerFile) throw new Error('Not a valid EPUB (no container.xml)');

    const container = new DOMParser()
      .parseFromString(await containerFile.async('string'), 'application/xml');
    const opfPath = container.querySelector('rootfile')?.getAttribute('full-path');
    if (!opfPath) throw new Error('Not a valid EPUB (no OPF rootfile)');

    const opf = new DOMParser()
      .parseFromString(await zip.file(opfPath).async('string'), 'application/xml');
    const base = opfPath.includes('/') ? opfPath.slice(0, opfPath.lastIndexOf('/') + 1) : '';

    const manifest = {};
    opf.querySelectorAll('manifest > item').forEach(item => {
      manifest[item.getAttribute('id')] = item.getAttribute('href');
    });

    const parts = [];
    const refs = Array.from(opf.querySelectorAll('spine > itemref'));
    for (const ref of refs) {
      const href = manifest[ref.getAttribute('idref')];
      if (!href) continue;
      const entry = zip.file(resolvePath(base + href));
      if (!entry) continue;
      const md = htmlToMarkdown(await entry.async('string'), opts);
      if (md.trim()) parts.push(md.trim());
    }
    if (!parts.length) return { markdown: '', warnings: ['No readable chapters in this EPUB.'] };
    return { markdown: tidy(parts.join('\n\n---\n\n')), warnings: [] };
  }

  function resolvePath(p) {
    const out = [];
    p.split('/').forEach(seg => {
      if (seg === '.' || seg === '') return;
      if (seg === '..') out.pop();
      else out.push(seg);
    });
    return out.join('/');
  }

  /* ------------------------------------------------------- text-ish formats */

  function jsonToMarkdown(text) {
    try {
      return { markdown: '```json\n' + JSON.stringify(JSON.parse(text), null, 2) + '\n```\n', warnings: [] };
    } catch (e) {
      return { markdown: '```\n' + text.trim() + '\n```\n', warnings: ['Not valid JSON; wrapped as plain text.'] };
    }
  }

  /* -------------------------------------------------------------- dispatch  */

  const EXTENSIONS = [
    'pdf', 'docx', 'xlsx', 'xls', 'xlsm', 'pptx', 'csv', 'tsv',
    'html', 'htm', 'xhtml', 'epub', 'json', 'xml', 'txt', 'log', 'md', 'markdown'
  ];

  function extOf(name) {
    const i = name.lastIndexOf('.');
    return i < 0 ? '' : name.slice(i + 1).toLowerCase();
  }

  const supports = name => EXTENSIONS.includes(extOf(name));

  /**
   * The single entry point the UI calls. Swapping any branch for a fetch() to
   * a local MarkItDown/Pandoc service is a one-line change here.
   */
  async function convert(file, opts, onProgress) {
    const ext = extOf(file.name);
    opts = Object.assign({
      embedImages: true, pageBreaks: true, detectTables: true,
      headingPerSection: true, detectHeadings: true
    }, opts);

    switch (ext) {
      case 'pdf':   return pdfToMarkdown(await readAsArrayBuffer(file), opts, onProgress);
      case 'docx':  return docxToMarkdown(await readAsArrayBuffer(file), opts);
      case 'xlsx':
      case 'xlsm':
      case 'xls':   return xlsxToMarkdown(await readAsArrayBuffer(file), opts);
      case 'pptx':  return pptxToMarkdown(await readAsArrayBuffer(file), opts);
      case 'epub':  return epubToMarkdown(await readAsArrayBuffer(file), opts);
      case 'csv':
      case 'tsv':   return csvToMarkdown(await readAsText(file), opts);
      case 'html':
      case 'htm':
      case 'xhtml': return { markdown: htmlToMarkdown(await readAsText(file), opts), warnings: [] };
      case 'json':  return jsonToMarkdown(await readAsText(file));
      case 'xml':   return { markdown: '```xml\n' + (await readAsText(file)).trim() + '\n```\n', warnings: [] };
      case 'md':
      case 'markdown': return { markdown: tidy(await readAsText(file)), warnings: ['Already Markdown; passed through unchanged.'] };
      case 'txt':
      case 'log':   return { markdown: tidy(await readAsText(file)), warnings: [] };
      case 'doc':   throw new Error('Legacy .doc is not supported — re-save it as .docx first.');
      default:      throw new Error('Unsupported file type: .' + (ext || '(none)'));
    }
  }

  global.ToMd = { convert, supports, extOf, htmlToMarkdown, toPipeTable, tidy };
})(window);
