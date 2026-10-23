/* global XLSX */
(() => {
  'use strict';
  const flagFontURL = new URL('vendor/TwemojiCountryFlags.woff2', document.currentScript?.src || document.baseURI).href;
  // Compact print scale, preserving relative font sizes from the workbook.
  const printScale = 0.77;
  const scaledPoints = points => `${Math.round(points * printScale * 100) / 100}pt`;
  function applyPrintStyle(element, css) {
    if (!css) return;
    Object.assign(element.style, css);
    if (css.fontSize?.endsWith('pt')) element.style.fontSize = scaledPoints(parseFloat(css.fontSize));
  }

  function createDocument(workbook, { sheetName, title, skipHeader, visuals }, xlsx = XLSX) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) throw new Error('Выбранный лист не найден.');
    const doc = document.implementation.createHTMLDocument(title);
    doc.documentElement.lang = 'ru';
    const meta = doc.createElement('meta'); meta.charset = 'utf-8'; doc.head.prepend(meta);
    // Ignore formatting-only cells when finding the printable area.
    const positions = Object.keys(sheet).filter(address => !address.startsWith('!') &&
      String(xlsx.utils.format_cell(sheet[address]) || '').trim()).map(xlsx.utils.decode_cell);
    positions.push(...(visuals?.images || []).map(image => ({ r: image.r, c: image.c })));
    if (!positions.length) throw new Error('Выбранный лист пуст: нет данных для PDF.');
    const rowColumns = new Map();
    for (const pos of positions) {
      if (!rowColumns.has(pos.r)) rowColumns.set(pos.r, new Set());
      rowColumns.get(pos.r).add(pos.c);
    }
    const range = { s: { r: Infinity, c: Infinity }, e: { r: 0, c: 0 } };
    for (const pos of positions) {
      range.s.r = Math.min(range.s.r, pos.r); range.s.c = Math.min(range.s.c, pos.c);
      range.e.r = Math.max(range.e.r, pos.r); range.e.c = Math.max(range.e.c, pos.c);
    }
    const merges = (sheet['!merges'] || []).filter(merge => merge.s.r <= range.e.r &&
      merge.s.c <= range.e.c && merge.e.r >= range.s.r && merge.e.c >= range.s.c);
    for (const merge of merges) {
      range.s.r = Math.min(range.s.r, merge.s.r); range.s.c = Math.min(range.s.c, merge.s.c);
      range.e.r = Math.max(range.e.r, merge.e.r); range.e.c = Math.max(range.e.c, merge.e.c);
    }
    const columns = range.e.c - range.s.c + 1;
    const rowCount = range.e.r - range.s.r + 1;
    if (rowCount * columns > 250000) {
      throw new Error('Область печати слишком велика. Удалите лишние пустые строки и столбцы или разделите лист.');
    }
    // Keep interior blank rows: they separate sections in the source document.
    // The bounds above still exclude the formatting-only tail of the worksheet.
    const retainedRows = Array.from({ length: rowCount }, (_, index) => range.s.r + index);
    const style = doc.createElement('style');
    style.textContent = `
      @font-face { font-family: 'Print Flags'; src: url(${JSON.stringify(flagFontURL)}) format('woff2'); unicode-range: U+1F1E6-1F1FF; font-display: block; }
      /* Zero vertical page margins leave no room for browser date/URL labels.
         Repeated table spacers supply the content insets on every page. */
      @page { size: A4 portrait; margin: 0 10mm; }
      * { box-sizing: border-box; }
      body { margin: 0; color: #111; font: 8.5pt/1.05 'Print Flags', 'Segoe UI', Arial, sans-serif; }
      table { width: 100%; border-collapse: collapse; table-layout: auto; }
      th, td { border: .2mm solid #999; padding: .25mm .5mm; vertical-align: middle; white-space: pre-wrap; overflow-wrap: anywhere; font-weight: normal; }
      th { background: #edf1ec; font-weight: 600; }
      .cell-pictures { display: flex; flex-wrap: wrap; align-items: center; gap: 2px; }
      .cell-pictures img { display: inline-block; max-width: 100%; height: auto; object-fit: contain; vertical-align: middle; }
      thead { display: table-header-group; }
      tfoot { display: table-footer-group; }
      .page-spacer td { height: 5mm; padding: 0; border: 0; background: transparent; }
      .section-separator td { padding: 0; border: 0; background: transparent; }
      tr { break-inside: avoid; }
      body { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
    `;
    doc.head.append(style);
    const table = doc.createElement('table');
    const firstRow = retainedRows[0];
    const head = doc.createElement('thead'), body = doc.createElement('tbody');
    const spacer = () => {
      const row = doc.createElement('tr'); row.className = 'page-spacer'; row.setAttribute('aria-hidden', 'true');
      const cell = doc.createElement('td'); cell.colSpan = columns; row.append(cell); return row;
    };
    head.append(spacer());
    const foot = doc.createElement('tfoot'); foot.append(spacer());
    const covered = new Set(), starts = new Map(), imageCells = new Map();
    for (const image of visuals?.images || []) {
      // A picture anchored in a merged cell belongs to its visible top-left cell.
      const merge = merges.find(m => image.r >= m.s.r && image.r <= m.e.r && image.c >= m.s.c && image.c <= m.e.c);
      const r = merge ? (retainedRows.find(row => row >= merge.s.r && row <= merge.e.r) ?? image.r) : image.r;
      const key = `${r}:${merge ? merge.s.c : image.c}`;
      if (!imageCells.has(key)) imageCells.set(key, []);
      imageCells.get(key).push(image);
    }
    for (const merge of merges) {
      const mergedRows = retainedRows.filter(r => r >= merge.s.r && r <= merge.e.r);
      if (!mergedRows.length) continue;
      starts.set(`${mergedRows[0]}:${merge.s.c}`, { ...merge, rowSpan: mergedRows.length });
      for (const r of mergedRows) {
        for (let c = merge.s.c; c <= merge.e.c; c++) {
          if (r !== mergedRows[0] || c !== merge.s.c) covered.add(`${r}:${c}`);
        }
      }
    }
    for (const r of retainedRows) {
      const row = doc.createElement('tr');
      row.dataset.sourceRow = r + 1;
      if (visuals?.rows[r]?.height) row.style.height = scaledPoints(visuals.rows[r].height);
      const isHeader = skipHeader && r === firstRow;
      const populatedColumns = rowColumns.get(r) || new Set();
      if (!populatedColumns.size) {
        row.classList.add('section-separator');
        row.style.height = scaledPoints(visuals?.rows[r]?.height || sheet['!rows']?.[r]?.hpt || 15);
      }
      for (let c = range.s.c; c <= range.e.c; c++) {
        const key = `${r}:${c}`;
        if (covered.has(key)) continue;
        const td = doc.createElement(isHeader ? 'th' : 'td');
        const address = xlsx.utils.encode_cell({ r, c });
        const cell = sheet[address], appearance = visuals?.cells[address];
        const css = appearance?.css || visuals?.defaultStyle;
        applyPrintStyle(td, css);
        // Text nodes retain Unicode emojis, Cyrillic and literal HTML safely.
        const text = cell ? xlsx.utils.format_cell(cell) : '';
        // Keep short numbers, years and flag codes together while long text wraps.
        if (String(text).length <= 10 && !String(text).includes('\n')) td.style.whiteSpace = 'nowrap';
        if (appearance?.runs.length && text === appearance.raw) {
          for (const run of appearance.runs) {
            const span = doc.createElement('span'); span.textContent = run.text; applyPrintStyle(span, run.css); td.append(span);
          }
        } else td.textContent = text;
        const pictures = imageCells.get(key);
        if (pictures?.length) {
          const gallery = doc.createElement('div'); gallery.className = 'cell-pictures';
          for (const picture of pictures) {
            const img = doc.createElement('img'); img.src = picture.src; img.alt = picture.alt;
            if (picture.width > 0) img.style.width = `${picture.width * printScale}px`;
            gallery.append(img);
          }
          td.append(gallery);
        }
        const merge = starts.get(key);
        if (merge) { td.colSpan = merge.e.c - c + 1; td.rowSpan = merge.rowSpan; }
        else if (String(text).length > 10) {
          // Let section labels extend through adjacent blank cells, stopping
          // before another value, picture or merged region, as Excel does.
          let next = c + 1;
          while (next <= range.e.c && !populatedColumns.has(next) &&
            !covered.has(`${r}:${next}`) && !starts.has(`${r}:${next}`) && !imageCells.has(`${r}:${next}`)) next++;
          if (next > c + 1) { td.colSpan = next - c; c = next - 1; }
        }
        row.append(td);
      }
      // Source headings belong to the document body and print only once.
      body.append(row);
    }
    table.append(head, body, foot); doc.body.append(table);
    return '<!doctype html>\n' + doc.documentElement.outerHTML;
  }

  async function prepare(workbook, options) {
    const html = createDocument(workbook, options);
    // Keep only the current print document. Offscreen (not display:none) so it lays out.
    document.querySelectorAll('.pdf-frame').forEach(frame => frame.remove());
    const frame = document.createElement('iframe');
    frame.className = 'pdf-frame'; frame.title = 'Таблица для сохранения в PDF';
    frame.style.cssText = 'position:fixed;left:-10000px;top:0;width:1100px;height:800px;border:0;';
    const loaded = new Promise((resolve, reject) => {
      frame.onload = resolve;
      frame.onerror = () => reject(new Error('Не удалось подготовить страницу печати.'));
    });
    frame.srcdoc = html; document.body.append(frame);
    try {
      await loaded;
      const fonts = await frame.contentDocument.fonts.load('12px "Print Flags"', '🇩🇪');
      if (!fonts.length) throw new Error('Не удалось загрузить шрифт флагов. Обновите страницу и повторите.');
      await frame.contentDocument.fonts.ready;
      await Promise.all([...frame.contentDocument.images].map(image => new Promise((resolve, reject) => {
        const check = () => image.naturalWidth ? resolve() : reject(new Error('Не удалось отобразить картинку «' + image.alt + '». Сохраните её в исходном Excel в формате PNG или JPEG.'));
        if (image.complete) check();
        else { image.addEventListener('load', check, { once: true }); image.addEventListener('error', check, { once: true }); }
      })));
      frame.contentWindow.addEventListener('afterprint', () => frame.remove(), { once: true });
      return frame;
    } catch (error) { frame.remove(); throw error; }
  }

  async function print(workbook, options) {
    const frame = await prepare(workbook, options);
    frame.contentWindow.focus();
    frame.contentWindow.print();
  }
  globalThis.PdfExport = { createDocument, prepare, print };
})();
