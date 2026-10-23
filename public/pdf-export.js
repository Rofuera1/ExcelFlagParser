/* global XLSX */
(() => {
  'use strict';
  const flagFontURL = new URL('vendor/TwemojiCountryFlags.woff2', document.currentScript?.src || document.baseURI).href;

  function createDocument(workbook, { sheetName, title, skipHeader }, xlsx = XLSX) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) throw new Error('Выбранный лист не найден.');
    const doc = document.implementation.createHTMLDocument(title);
    doc.documentElement.lang = 'ru';
    const meta = doc.createElement('meta'); meta.charset = 'utf-8'; doc.head.prepend(meta);
    // Ignore formatting-only cells when finding the printable area.
    const positions = Object.keys(sheet).filter(address => !address.startsWith('!') &&
      (sheet[address].v != null || sheet[address].f)).map(xlsx.utils.decode_cell);
    if (!positions.length) throw new Error('Выбранный лист пуст: нет данных для PDF.');
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
    if ((range.e.r - range.s.r + 1) * columns > 250000) {
      throw new Error('Область печати слишком велика. Удалите лишние пустые строки и столбцы или разделите лист.');
    }
    const style = doc.createElement('style');
    style.textContent = `
      @font-face { font-family: 'Print Flags'; src: url(${JSON.stringify(flagFontURL)}) format('woff2'); unicode-range: U+1F1E6-1F1FF; font-display: block; }
      /* Zero vertical page margins leave no room for browser date/URL labels.
         Repeated table spacers supply the content insets on every page. */
      @page { size: A4 portrait; margin: 0 10mm; }
      * { box-sizing: border-box; }
      body { margin: 0; color: #111; font: 10pt 'Print Flags', 'Segoe UI', Arial, sans-serif; }
      h1 { font-size: 14pt; font-weight: 600; margin: 8mm 0 0; overflow-wrap: anywhere; }
      table { width: 100%; border-collapse: collapse; table-layout: auto; }
      th, td { border: .2mm solid #999; padding: 1.8mm; vertical-align: middle; white-space: pre-wrap; overflow-wrap: anywhere; font-weight: normal; }
      th { background: #edf1ec; font-weight: 600; }
      thead { display: table-header-group; }
      tfoot { display: table-footer-group; }
      .page-spacer td { height: 8mm; padding: 0; border: 0; background: transparent; }
      tr { break-inside: avoid; }
      body { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
    `;
    doc.head.append(style);
    const heading = doc.createElement('h1'); heading.textContent = sheetName; doc.body.append(heading);
    const table = doc.createElement('table');
    const repeatHeader = skipHeader && !merges.some(m => m.s.r === range.s.r && m.e.r > range.s.r);
    const head = doc.createElement('thead'), body = doc.createElement('tbody');
    const spacer = () => {
      const row = doc.createElement('tr'); row.className = 'page-spacer'; row.setAttribute('aria-hidden', 'true');
      const cell = doc.createElement('td'); cell.colSpan = columns; row.append(cell); return row;
    };
    head.append(spacer());
    const foot = doc.createElement('tfoot'); foot.append(spacer());
    const covered = new Set(), starts = new Map();
    for (const merge of merges) {
      starts.set(`${merge.s.r}:${merge.s.c}`, merge);
      for (let r = merge.s.r; r <= merge.e.r; r++) {
        for (let c = merge.s.c; c <= merge.e.c; c++) {
          if (r !== merge.s.r || c !== merge.s.c) covered.add(`${r}:${c}`);
        }
      }
    }
    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = doc.createElement('tr');
      const isHeader = repeatHeader && r === range.s.r;
      for (let c = range.s.c; c <= range.e.c; c++) {
        const key = `${r}:${c}`;
        if (covered.has(key)) continue;
        const td = doc.createElement(isHeader ? 'th' : 'td');
        const cell = sheet[xlsx.utils.encode_cell({ r, c })];
        // Text nodes retain Unicode emojis, Cyrillic and literal HTML safely.
        td.textContent = cell ? xlsx.utils.format_cell(cell) : '';
        const merge = starts.get(key);
        if (merge) { td.colSpan = merge.e.c - c + 1; td.rowSpan = merge.e.r - r + 1; }
        row.append(td);
      }
      (isHeader ? head : body).append(row);
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
