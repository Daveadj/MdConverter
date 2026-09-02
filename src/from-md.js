/* =============================================================================
 * from-md.js — Markdown -> DOCX, PDF and standalone HTML, in the browser.
 *
 * DOCX is built with docx.js from marked's token stream (not from HTML), so
 * headings, lists and tables come out as real Word constructs a user can
 * restyle — not as a wall of manually-formatted paragraphs.
 *
 * PDF goes through the browser's own print engine. That is the same layout
 * engine a server-side headless-Chromium pipeline would use, so the output is
 * as good as it would be there, without shipping a server.
 * ========================================================================== */
(function (global) {
  'use strict';

  const PAPER = {
    a4:     { w: 11906, h: 16838, css: 'A4' },
    letter: { w: 12240, h: 15840, css: 'Letter' }
  };

  const DEFAULTS = {
    title: 'Untitled document',
    paper: 'a4',
    margin: 1,
    font: 'Calibri',
    pageNumbers: true
  };

  const opt = o => Object.assign({}, DEFAULTS, o);

  // marked hands back token text with HTML entities still encoded, because it
  // expects to render into HTML. Word is not HTML, so an apostrophe would be
  // written into the document literally as "&#39;". Decode before it gets there.
  const _decoder = document.createElement('textarea');
  function decode(text) {
    const s = String(text ?? '');
    if (s.indexOf('&') === -1) return s;
    _decoder.innerHTML = s;
    return _decoder.value;
  }

  /* ================================================================= DOCX */

  async function toDocxBlob(markdown, options) {
    const o = opt(options);
    if (!global.docx) throw new Error('docx.js failed to load');

    const D = global.docx;
    const tokens = marked.lexer(markdown || '');
    const images = await preloadImages(tokens);

    const ctx = { D, o, images, listInstance: 0 };
    const children = blocksToDocx(tokens, ctx, 0);
    if (!children.length) children.push(new D.Paragraph({ text: '' }));

    const paper = PAPER[o.paper] || PAPER.a4;
    const margin = D.convertInchesToTwip(Number(o.margin) || 1);

    const doc = new D.Document({
      creator: 'Markdown Converter',
      title: o.title,
      styles: docxStyles(D, o.font),
      numbering: {
        config: [{
          reference: 'md-ordered',
          levels: [0, 1, 2, 3].map(level => ({
            level,
            format: ['decimal', 'lowerLetter', 'lowerRoman', 'decimal'][level],
            text: `%${level + 1}.`,
            alignment: D.AlignmentType.START,
            style: { paragraph: { indent: { left: 420 + level * 420, hanging: 300 } } }
          }))
        }]
      },
      sections: [{
        properties: {
          page: {
            size: { width: paper.w, height: paper.h },
            margin: { top: margin, right: margin, bottom: margin, left: margin }
          }
        },
        footers: o.pageNumbers ? { default: pageFooter(D) } : undefined,
        children
      }]
    });

    return D.Packer.toBlob(doc);
  }

  function pageFooter(D) {
    return new D.Footer({
      children: [new D.Paragraph({
        alignment: D.AlignmentType.CENTER,
        children: [new D.TextRun({ children: [D.PageNumber.CURRENT], size: 18, color: '8B949E' })]
      })]
    });
  }

  function docxStyles(D, font) {
    const heading = (size, before) => ({
      run: { font, size, bold: true, color: '1B1F24' },
      paragraph: { spacing: { before, after: 120 }, keepNext: true }
    });
    return {
      default: {
        document: {
          run: { font, size: 22, color: '1B1F24' },
          paragraph: { spacing: { after: 160, line: 288 } }
        },
        heading1: heading(40, 0),
        heading2: heading(30, 320),
        heading3: heading(25, 280),
        heading4: heading(22, 240),
        heading5: heading(21, 220),
        heading6: heading(21, 220)
      },
      // These carry the names the reader half of this app maps back to fenced
      // code and blockquotes. Word has no elements for either; a named
      // paragraph style is the only thing that survives the trip.
      paragraphStyles: [{
        id: 'MdCode',
        name: 'Md Code',
        basedOn: 'Normal',
        quickFormat: false,
        run: { font: 'Consolas', size: 19, color: '24292F' },
        paragraph: { spacing: { before: 120, after: 160, line: 240 }, indent: { left: 240 } }
      }, {
        id: 'MdQuote',
        name: 'Quote',
        basedOn: 'Normal',
        quickFormat: true,
        run: { italics: true, color: '5C6570' },
        paragraph: { indent: { left: 300 }, spacing: { before: 120, after: 160 } }
      }]
    };
  }

  function blocksToDocx(tokens, ctx, depth) {
    const { D } = ctx;
    const out = [];

    for (const token of tokens) {
      switch (token.type) {
        case 'heading': {
          const level = Math.min(Math.max(token.depth, 1), 6);
          out.push(new D.Paragraph({
            heading: D.HeadingLevel['HEADING_' + level],
            children: inlineRuns(token.tokens || [{ type: 'text', text: token.text }], ctx, {})
          }));
          break;
        }

        case 'paragraph': {
          // A paragraph that is nothing but an image should become a picture,
          // not a picture squeezed into a text run.
          const onlyImage = (token.tokens || []).filter(t => t.type !== 'space').length === 1 &&
                            (token.tokens || [])[0]?.type === 'image';
          if (onlyImage) {
            const p = imageParagraph(token.tokens[0], ctx);
            if (p) { out.push(p); break; }
          }
          out.push(new D.Paragraph({ children: inlineRuns(token.tokens || [], ctx, {}) }));
          break;
        }

        case 'text':
          out.push(new D.Paragraph({
            children: inlineRuns(token.tokens || [{ type: 'text', text: token.text }], ctx, {})
          }));
          break;

        case 'list': {
          const instance = token.ordered ? ++ctx.listInstance : 0;
          token.items.forEach(item => out.push(...listItemToDocx(item, token, ctx, depth, instance)));
          break;
        }

        case 'blockquote':
          out.push(...blockquoteParagraphs(token, ctx, depth));
          break;

        case 'code':
          out.push(...codeParagraphs(token.text || '', ctx));
          break;

        case 'table':
          out.push(tableToDocx(token, ctx));
          out.push(new D.Paragraph({ text: '', spacing: { after: 80 } }));
          break;

        case 'hr':
          out.push(new D.Paragraph({
            border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: 'DFE3E8', space: 8 } },
            spacing: { before: 200, after: 200 }
          }));
          break;

        case 'html': {
          const text = decode(String(token.text || '').replace(/<[^>]*>/g, '')).trim();
          if (text) out.push(new D.Paragraph({ children: [new D.TextRun(text)] }));
          break;
        }

        case 'space':
        default:
          break;
      }
    }
    return out;
  }

  /**
   * Word has no blockquote construct, so the quote bar is a left paragraph
   * border. Anything that is not a plain paragraph (a nested list, a code
   * block) falls through to the normal builder rather than being flattened.
   */
  function blockquoteParagraphs(token, ctx, depth) {
    const { D } = ctx;
    const border = { left: { style: D.BorderStyle.SINGLE, size: 12, color: 'C9CFD6', space: 12 } };
    const out = [];
    for (const t of token.tokens || []) {
      if (t.type === 'paragraph' || t.type === 'text') {
        out.push(new D.Paragraph({
          style: 'MdQuote',
          border,
          children: inlineRuns(t.tokens || [{ type: 'text', text: t.text }], ctx, {})
        }));
      } else {
        out.push(...blocksToDocx([t], ctx, depth));
      }
    }
    return out;
  }

  function listItemToDocx(item, list, ctx, depth, instance) {
    const { D } = ctx;
    const out = [];
    const level = Math.min(depth, 3);

    const leading = (item.tokens || []).filter(t => t.type === 'text' || t.type === 'paragraph');
    const runs = leading.length
      ? inlineRuns(leading[0].tokens || [{ type: 'text', text: leading[0].text }], ctx, {})
      : [new D.TextRun(decode(item.text))];

    const props = { children: runs, spacing: { after: 60 } };
    if (list.ordered) props.numbering = { reference: 'md-ordered', level, instance };
    else props.bullet = { level };

    if (item.task) {
      props.children = [new D.TextRun({ text: item.checked ? '☑  ' : '☐  ' })].concat(runs);
      delete props.bullet;
      delete props.numbering;
      props.indent = { left: 420 + level * 420, hanging: 300 };
    }
    out.push(new D.Paragraph(props));

    // Nested lists and any block content that followed the item's first line.
    (item.tokens || []).forEach((t, i) => {
      if (t.type === 'list') {
        const nestedInstance = t.ordered ? ++ctx.listInstance : 0;
        t.items.forEach(sub => out.push(...listItemToDocx(sub, t, ctx, depth + 1, nestedInstance)));
      } else if (t.type === 'code') {
        out.push(...codeParagraphs(t.text || '', ctx));
      } else if ((t.type === 'text' || t.type === 'paragraph') && t !== leading[0]) {
        out.push(new D.Paragraph({
          indent: { left: 420 + level * 420 },
          children: inlineRuns(t.tokens || [{ type: 'text', text: t.text }], ctx, {})
        }));
      }
    });
    return out;
  }

  function codeParagraphs(text, ctx) {
    const { D } = ctx;
    const lines = String(text).replace(/\s+$/, '').split('\n');
    const shading = { type: D.ShadingType.CLEAR, color: 'auto', fill: 'F4F5F7' };
    return lines.map((line, i) => new D.Paragraph({
      style: 'MdCode',
      shading,
      spacing: { before: i === 0 ? 120 : 0, after: i === lines.length - 1 ? 160 : 0, line: 240 },
      children: [new D.TextRun({ text: line || ' ', font: 'Consolas' })]
    }));
  }

  function tableToDocx(token, ctx) {
    const { D } = ctx;
    const align = i => {
      const a = (token.align && token.align[i]) || (token.header[i] && token.header[i].align);
      return a === 'center' ? D.AlignmentType.CENTER
           : a === 'right'  ? D.AlignmentType.RIGHT
           : D.AlignmentType.LEFT;
    };

    const cell = (c, i, header) => new D.TableCell({
      shading: header ? { type: D.ShadingType.CLEAR, color: 'auto', fill: 'F0F2F5' } : undefined,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: [new D.Paragraph({
        alignment: align(i),
        spacing: { after: 0 },
        children: inlineRuns(c.tokens || [{ type: 'text', text: c.text }], ctx, { bold: header })
      })]
    });

    const rows = [new D.TableRow({
      tableHeader: true,
      children: token.header.map((c, i) => cell(c, i, true))
    })];
    (token.rows || []).forEach(r => {
      rows.push(new D.TableRow({ children: r.map((c, i) => cell(c, i, false)) }));
    });

    const edge = { style: D.BorderStyle.SINGLE, size: 4, color: 'DFE3E8' };
    return new D.Table({
      width: { size: 100, type: D.WidthType.PERCENTAGE },
      borders: { top: edge, bottom: edge, left: edge, right: edge, insideHorizontal: edge, insideVertical: edge },
      rows
    });
  }

  /* ------------------------------------------------------------- inline */

  function inlineRuns(tokens, ctx, style) {
    const { D } = ctx;
    const out = [];

    for (const t of tokens || []) {
      switch (t.type) {
        case 'text':
        case 'escape': {
          if (t.tokens && t.tokens.length) { out.push(...inlineRuns(t.tokens, ctx, style)); break; }
          decode(t.text).split('\n').forEach((chunk, i) => {
            out.push(new D.TextRun(Object.assign({ text: chunk }, style, i ? { break: 1 } : {})));
          });
          break;
        }
        case 'strong': out.push(...inlineRuns(t.tokens, ctx, Object.assign({}, style, { bold: true }))); break;
        case 'em':     out.push(...inlineRuns(t.tokens, ctx, Object.assign({}, style, { italics: true }))); break;
        case 'del':    out.push(...inlineRuns(t.tokens, ctx, Object.assign({}, style, { strike: true }))); break;

        case 'codespan':
          out.push(new D.TextRun(Object.assign({}, style, {
            text: decode(t.text), font: 'Consolas', size: 19,
            shading: { type: D.ShadingType.CLEAR, color: 'auto', fill: 'F0F2F5' }
          })));
          break;

        case 'br': out.push(new D.TextRun({ break: 1 })); break;

        case 'link': {
          const inner = inlineRuns(t.tokens && t.tokens.length ? t.tokens : [{ type: 'text', text: t.text }],
                                   ctx, Object.assign({}, style, { color: '2F6FD0', underline: {} }));
          if (/^[a-z][a-z0-9+.-]*:/i.test(t.href || '')) {
            out.push(new D.ExternalHyperlink({ children: inner, link: t.href }));
          } else {
            out.push(...inner);                  // relative link — no target to point at
          }
          break;
        }

        case 'image': {
          const run = imageRun(t, ctx, 420);
          if (run) out.push(run);
          else out.push(new D.TextRun(Object.assign({ text: t.text || '[image]', italics: true }, style)));
          break;
        }

        case 'html': {
          const text = decode(String(t.text || '').replace(/<[^>]*>/g, ''));
          if (text) out.push(new D.TextRun(Object.assign({ text }, style)));
          break;
        }

        default:
          if (t.tokens) out.push(...inlineRuns(t.tokens, ctx, style));
          else if (t.text) out.push(new D.TextRun(Object.assign({ text: decode(t.text) }, style)));
      }
    }
    if (!out.length) out.push(new D.TextRun(''));
    return out;
  }

  function imageRun(token, ctx, maxWidth) {
    const img = ctx.images.get(token.href);
    if (!img) return null;
    const scale = Math.min(1, maxWidth / img.width);
    return new ctx.D.ImageRun({
      data: img.data,
      transformation: { width: Math.round(img.width * scale), height: Math.round(img.height * scale) }
    });
  }

  function imageParagraph(token, ctx) {
    const run = imageRun(token, ctx, 560);
    if (!run) return null;
    return new ctx.D.Paragraph({ alignment: ctx.D.AlignmentType.CENTER, children: [run] });
  }

  /** Fetch every image referenced anywhere in the token tree, once. */
  async function preloadImages(tokens) {
    const srcs = new Set();
    (function walk(list) {
      (list || []).forEach(t => {
        if (t.type === 'image' && t.href) srcs.add(t.href);
        if (t.tokens) walk(t.tokens);
        if (t.items) walk(t.items);
        if (t.rows) t.rows.forEach(r => r.forEach(c => walk(c.tokens)));
        if (t.header) t.header.forEach(c => walk(c.tokens));
      });
    })(tokens);

    const map = new Map();
    await Promise.all(Array.from(srcs).map(async src => {
      try {
        const res = await fetch(src);
        if (!res.ok) return;
        const blob = await res.blob();
        const bitmap = await createImageBitmap(blob);
        map.set(src, {
          data: await blob.arrayBuffer(),
          width: bitmap.width,
          height: bitmap.height
        });
        bitmap.close();
      } catch (e) { /* unreachable or unsupported image — falls back to alt text */ }
    }));
    return map;
  }

  /* ================================================= HTML and print / PDF */

  const PRINT_CSS = font => `
:root { color-scheme: light; }
* { box-sizing: border-box; }
body {
  font-family: "${font}", "Segoe UI", system-ui, sans-serif;
  font-size: 11pt; line-height: 1.55; color: #1b1f24;
  margin: 0; padding: 0; background: #fff;
  -webkit-print-color-adjust: exact; print-color-adjust: exact;
}
.page { max-width: 46em; margin: 0 auto; padding: 2em 1em; }
h1,h2,h3,h4,h5,h6 { line-height: 1.25; margin: 1.4em 0 .5em; font-weight: 650; break-after: avoid; page-break-after: avoid; }
h1 { font-size: 22pt; margin-top: 0; }
h2 { font-size: 16pt; border-bottom: 1px solid #dfe3e8; padding-bottom: .2em; }
h3 { font-size: 13pt; }
h4 { font-size: 11.5pt; color: #40474f; }
p, ul, ol, blockquote, table, pre { margin: 0 0 .85em; orphans: 3; widows: 3; }
li { margin: .2em 0; }
a { color: #1a4f9c; text-decoration: underline; }
code { font-family: Consolas, "Cascadia Mono", monospace; font-size: .88em; background: #f0f2f5; padding: .1em .35em; border-radius: 3px; }
pre { background: #f4f5f7; border: 1px solid #e6e9ed; border-radius: 5px; padding: .8em 1em; overflow-x: auto; break-inside: avoid; page-break-inside: avoid; }
pre code { background: none; padding: 0; font-size: .84em; }
blockquote { border-left: 3px solid #c9cfd6; padding-left: 1em; margin-left: 0; color: #4a525b; }
table { border-collapse: collapse; width: 100%; font-size: .92em; break-inside: avoid; page-break-inside: avoid; }
th, td { border: 1px solid #dfe3e8; padding: .35em .6em; text-align: left; vertical-align: top; }
th { background: #f0f2f5; font-weight: 650; }
img { max-width: 100%; height: auto; break-inside: avoid; }
hr { border: none; border-top: 1px solid #dfe3e8; margin: 1.8em 0; }
ul.contains-task-list { list-style: none; padding-left: 1.1em; }
@media print {
  .page { max-width: none; padding: 0; }
  a { color: #1b1f24; text-decoration: none; }
}
`;

  function toHtmlDocument(markdown, options) {
    const o = opt(options);
    const paper = PAPER[o.paper] || PAPER.a4;
    const body = marked.parse(markdown || '', { gfm: true, breaks: false });
    const safe = global.DOMPurify ? DOMPurify.sanitize(body) : body;
    return `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>${escapeHtml(o.title)}</title>
<style>@page { size: ${paper.css}; margin: ${Number(o.margin) || 1}in; }
${PRINT_CSS(o.font)}</style>
</head><body><div class="page">${safe}</div></body></html>`;
  }

  const escapeHtml = s => String(s).replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  /**
   * Print through a hidden iframe. Resolves once the dialog has been dismissed
   * where the browser reports it, and after a grace period where it does not.
   */
  function printToPdf(markdown, options, frame) {
    return new Promise((resolve, reject) => {
      const html = toHtmlDocument(markdown, options);
      let settled = false;
      const done = () => { if (!settled) { settled = true; resolve(); } };

      frame.onload = () => {
        const win = frame.contentWindow;
        if (!win) return reject(new Error('Could not open the print view'));
        // Give embedded images a moment to decode, or they print blank.
        const go = () => {
          try {
            win.focus();
            if (win.matchMedia) {
              const mql = win.matchMedia('print');
              mql.addEventListener?.('change', e => { if (!e.matches) done(); });
            }
            win.onafterprint = done;
            win.print();
            setTimeout(done, 1500);
          } catch (e) { reject(e); }
        };
        const imgs = Array.from(win.document.images || []);
        if (!imgs.length) return go();
        Promise.all(imgs.map(i => i.complete ? Promise.resolve()
          : new Promise(r => { i.onload = i.onerror = r; }))).then(go);
      };
      frame.srcdoc = html;
    });
  }

  function renderPreview(markdown) {
    const html = marked.parse(markdown || '', { gfm: true, breaks: false });
    return global.DOMPurify ? DOMPurify.sanitize(html) : html;
  }

  global.FromMd = { toDocxBlob, toHtmlDocument, printToPdf, renderPreview };
})(window);
