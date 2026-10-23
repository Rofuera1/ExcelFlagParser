/* global XLSX, JSZip, FlagConverter, PdfExport, WorkbookPackage */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const items = [];
  let outputs = [];
  let busy = false;
  let nextId = 1;
  const mode = () => document.querySelector('input[name="format"]:checked').value;
  const pause = () => new Promise(resolve => setTimeout(resolve, 0));
  function element(tag, className, text) {
    const el = document.createElement(tag);
    if (className) el.className = className;
    if (text !== undefined) el.textContent = text;
    return el;
  }
  function invalidate() {
    outputs = [];
    $('results').replaceChildren();
    $('results-panel').hidden = true;
  }
  function update() {
    $('convert').disabled = busy || !items.length;
    $('file-input').disabled = busy;
    document.querySelectorAll('#files select, #files input, #files button, .formats input').forEach(el => { el.disabled = busy; });
    $('summary').textContent = busy ? 'Обрабатываем файлы на вашем устройстве…' : items.length ? `Файлов в очереди: ${items.length}` : 'Добавьте файлы, чтобы начать';
  }
  function refreshPreview(item) {
    const sheet = item.workbook.Sheets[item.sheetName];
    item.preview.replaceChildren();
    if (!sheet['!ref']) { item.preview.textContent = 'Этот лист пуст.'; return; }
    const range = XLSX.utils.decode_range(sheet['!ref']);
    let shown = 0;
    const start = range.s.r + (item.skipHeader ? 1 : 0);
    for (let row = start; row <= range.e.r && row < start + 100 && shown < 3; row++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: item.column })];
      if (!cell || cell.v == null || !String(cell.v).trim()) continue;
      const result = cell.f || cell.F ? { value: 'формула — без изменений' } : FlagConverter.convertValue(cell.v, mode());
      if (shown++) item.preview.append(' · ');
      item.preview.append(String(cell.v), ' → ', element('b', '', String(result.value)));
    }
    if (!shown) item.preview.textContent = 'В первых 100 строках столбца нет значений.';
  }
  function populateColumns(item) {
    const sheet = item.workbook.Sheets[item.sheetName];
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    item.columnSelect.replaceChildren();
    for (let col = range.s.c; col <= range.e.c; col++) {
      const letter = XLSX.utils.encode_col(col);
      const header = sheet[XLSX.utils.encode_cell({ r: range.s.r, c: col })];
      const label = header?.v == null ? letter : `${letter} — ${String(header.v).slice(0, 70)}`;
      const option = element('option', '', label);
      option.value = col;
      item.columnSelect.append(option);
    }
    if (item.column < range.s.c || item.column > range.e.c) item.column = range.s.c;
    item.columnSelect.value = item.column;
    refreshPreview(item);
  }
  function renderItem(item) {
    const card = element('article', 'file-card');
    const heading = element('div', 'file-heading');
    const remove = element('button', 'remove', '×');
    remove.type = 'button';
    remove.setAttribute('aria-label', `Удалить ${item.file.name}`);
    remove.addEventListener('click', () => {
      items.splice(items.indexOf(item), 1);
      card.remove(); invalidate(); update();
    });
    heading.append(element('strong', '', item.file.name), element('small', '', `${(item.file.size / 1024).toFixed(0)} КБ`), remove);
    const options = element('div', 'file-options');
    const sheetLabel = element('label', 'field', 'Лист');
    const sheetSelect = element('select');
    item.workbook.SheetNames.forEach(name => { const option = element('option', '', name); option.value = name; sheetSelect.append(option); });
    sheetSelect.value = item.sheetName;
    sheetLabel.append(sheetSelect);
    const colLabel = element('label', 'field', 'Столбец с кодом страны');
    item.columnSelect = element('select');
    colLabel.append(item.columnSelect);
    options.append(sheetLabel, colLabel);
    const headerLabel = element('label', 'header-option');
    const checkbox = element('input');
    checkbox.type = 'checkbox'; checkbox.checked = item.skipHeader;
    headerLabel.append(checkbox, 'Первая строка листа — заголовок (пропустить)');
    item.preview = element('div', 'preview');
    card.append(heading, options, headerLabel, item.preview);
    sheetSelect.addEventListener('change', () => { item.sheetName = sheetSelect.value; populateColumns(item); invalidate(); });
    item.columnSelect.addEventListener('change', () => { item.column = Number(item.columnSelect.value); refreshPreview(item); invalidate(); });
    checkbox.addEventListener('change', () => { item.skipHeader = checkbox.checked; refreshPreview(item); invalidate(); });
    populateColumns(item);
    $('files').append(card);
  }
  async function addFiles(files) {
    if (busy || !files.length) return;
    if (!globalThis.XLSX) { $('notice').textContent = 'Не удалось загрузить библиотеку Excel. Обновите страницу.'; return; }
    busy = true; update(); invalidate();
    const errors = [];
    for (const file of files) {
      if (!/\.(xlsx|xls|xlsm)$/i.test(file.name)) { errors.push(`${file.name}: выберите .xlsx, .xls или .xlsm.`); continue; }
      $('notice').textContent = `Читаем ${file.name}…`;
      await pause();
      try {
        const bytes = await file.arrayBuffer();
        // XLSX is a ZIP container; XLS uses OLE or older BIFF. Reject renamed text files.
        const signature = new Uint8Array(bytes, 0, Math.min(bytes.byteLength, 8));
        const zip = signature[0] === 0x50 && signature[1] === 0x4b;
        const ole = signature[0] === 0xd0 && signature[1] === 0xcf && signature[2] === 0x11 && signature[3] === 0xe0;
        const biff = signature[1] === 0x08 && [0x09, 0x02, 0x04].includes(signature[0]);
        if (!zip && !ole && !biff) throw new Error('Файл не похож на книгу Excel.');
        const workbook = XLSX.read(bytes, { type: 'array', cellStyles: true, cellDates: true, bookVBA: false });
        if (!workbook.SheetNames.length) throw new Error('В книге нет листов.');
        const item = { id: nextId++, file, bytes, workbook, sheetName: workbook.SheetNames[0], column: 0, skipHeader: true };
        items.push(item); renderItem(item);
      } catch (error) { errors.push(`${file.name}: не удалось прочитать. ${error.message}`); }
    }
    $('notice').textContent = errors.join('\n');
    busy = false; update();
    $('file-input').value = '';
  }
  function download(blob, name) {
    const url = URL.createObjectURL(blob);
    const link = element('a'); link.href = url; link.download = name;
    document.body.append(link); link.click(); link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  function uniqueName(file, used) {
    const base = file.name.replace(/\.[^.]+$/, '').replace(/[<>:"/\\|?*\x00-\x1f]/g, '_') + '_flags';
    const extension = /\.xlsm$/i.test(file.name) ? 'xlsm' : 'xlsx';
    let name = `${base}.${extension}`, n = 2;
    while (used.has(name.toLowerCase())) name = `${base}_${n++}.${extension}`;
    used.add(name.toLowerCase());
    return name;
  }
  $('file-input').addEventListener('change', event => addFiles(Array.from(event.target.files)));
  ['dragenter', 'dragover'].forEach(type => $('dropzone').addEventListener(type, event => { event.preventDefault(); if (!busy) $('dropzone').classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(type => $('dropzone').addEventListener(type, event => { event.preventDefault(); $('dropzone').classList.remove('dragover'); }));
  $('dropzone').addEventListener('drop', event => addFiles(Array.from(event.dataTransfer.files)));
  // Prevent a drop outside the upload area from navigating away from the app.
  window.addEventListener('dragover', event => event.preventDefault());
  window.addEventListener('drop', event => event.preventDefault());
  document.querySelectorAll('input[name="format"]').forEach(input => input.addEventListener('change', () => { invalidate(); items.forEach(refreshPreview); }));
  $('convert').addEventListener('click', async () => {
    if (busy || !items.length) return;
    invalidate(); busy = true; update();
    const errors = [], used = new Set();
    for (const [index, item] of items.entries()) {
      $('notice').textContent = `Преобразуем ${index + 1} из ${items.length}: ${item.file.name}`;
      await pause();
      try {
        // Always start from the original, including when switching output modes.
        const workbook = XLSX.read(item.bytes, { type: 'array', cellStyles: true, cellDates: true });
        const settings = { sheetName: item.sheetName, column: item.column, skipHeader: item.skipHeader, mode: mode() };
        const packaged = /\.(xlsx|xlsm)$/i.test(item.file.name);
        const data = packaged ? await WorkbookPackage.write(item.bytes, workbook, settings) : null;
        const stats = FlagConverter.transformWorkbook(workbook, settings, XLSX);
        const output = { name: uniqueName(item.file, used), blob: new Blob([data || XLSX.write(workbook, { type: 'array', bookType: 'xlsx', cellStyles: true, compression: true })], { type: /\.xlsm$/i.test(item.file.name) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }) };
        outputs.push(output);
        const row = element('div', 'result-row');
        const details = element('div', 'result-details');
        details.append(element('strong', '', output.name), element('small', '', `Лист «${item.sheetName}», столбец ${XLSX.utils.encode_col(item.column)} · заменено: ${stats.converted}`));
        if (stats.unknown || stats.formulas) {
          const warnings = [];
          if (stats.unknown) warnings.push(`Неизвестных значений: ${stats.unknown} (${stats.unknownCodes.slice(0, 8).map(code => code.slice(0, 30)).join(', ')}${stats.unknownCodes.length > 8 ? ', …' : ''}). Сохранены без изменений.`);
          if (stats.formulas) warnings.push(`Формул пропущено: ${stats.formulas}.`);
          details.append(element('small', 'warning', warnings.join('\n')));
        }
        if (!packaged) details.append(element('small', 'warning', 'Для сохранения картинок и оформления сначала сохраните исходный .xls как .xlsx в Excel.'));
        const button = element('button', 'secondary', /\.xlsm$/i.test(item.file.name) ? 'Скачать .xlsm' : 'Скачать .xlsx');
        button.classList.add('download-xlsx');
        button.addEventListener('click', () => download(output.blob, output.name));
        const pdf = element('button', 'secondary export-pdf', 'Сохранить PDF');
        pdf.addEventListener('click', async () => {
          pdf.disabled = true; pdf.textContent = 'Готовим PDF…';
          try {
            const printBook = XLSX.read(item.bytes, { type: 'array', cellStyles: true, cellDates: true });
            const visuals = packaged ? await WorkbookPackage.visuals(item.bytes, settings.sheetName) : undefined;
            FlagConverter.transformWorkbook(printBook, settings, XLSX);
            await PdfExport.print(printBook, { ...settings, visuals, title: output.name.replace(/\.xls[xm]$/, '') });
          } catch (error) { $('notice').textContent = `Не удалось подготовить PDF: ${error.message}`; }
          finally { pdf.disabled = false; pdf.textContent = 'Сохранить PDF'; }
        });
        const actions = element('div', 'result-actions'); actions.append(button, pdf);
        row.append(details, actions); $('results').append(row);
      } catch (error) { errors.push(`${item.file.name}: ошибка преобразования. ${error.message}`); }
    }
    $('notice').textContent = errors.join('\n');
    $('results-panel').hidden = !outputs.length;
    if (outputs.length) {
      $('results-panel').scrollIntoView({
        behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'instant' : 'smooth',
        block: 'start'
      });
    }
    $('download-all').disabled = !outputs.length || !globalThis.JSZip;
    if (!globalThis.JSZip && outputs.length) $('notice').textContent += '\nZIP недоступен. Скачайте результаты по отдельности.';
    busy = false; update();
  });
  $('download-all').addEventListener('click', async () => {
    const button = $('download-all'); button.disabled = true; button.textContent = 'Собираем ZIP…';
    try {
      const zip = new JSZip();
      // Capture results before yielding; each file gets a unique safe archive name.
      for (const output of outputs.slice()) zip.file(output.name, output.blob);
      download(await zip.generateAsync({ type: 'blob' }), 'excel_flags.zip');
    } catch (error) { $('notice').textContent = `Не удалось создать ZIP: ${error.message}. Скачайте файлы отдельно.`; }
    finally { button.disabled = false; button.textContent = 'Скачать всё ZIP'; }
  });
  if (!globalThis.XLSX) $('notice').textContent = 'Не удалось загрузить библиотеку Excel. Обновите страницу.';
})();
