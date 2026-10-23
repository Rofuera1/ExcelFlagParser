/* global XLSX, JSZip, FlagConverter, PdfExport */
(async () => {
  const log = [];
  let count = 0;
  function equal(actual, expected, name) {
    if (actual !== expected) throw new Error(`${name}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    count++; log.push(`OK ${name}`);
  }
  try {
    const { convertValue, flagFor, transformWorkbook } = FlagConverter;
    equal(flagFor('DE'), '🇩🇪', 'German flag');
    equal(flagFor('JP'), '🇯🇵', 'Japanese flag');
    equal(convertValue(' ru ', 'flag-code').value, '🇷🇺 RU', 'trim and uppercase');
    equal(convertValue('AIN', 'flag').value, '🏳️', 'AIN flag');
    equal(convertValue('ain', 'flag-code').value, '🏳️ AIN', 'AIN with code');
    equal(convertValue('  ', 'flag').value, '', 'whitespace');
    equal(convertValue(undefined, 'flag').value, '', 'absent value');
    equal(convertValue('ZZ', 'flag').kind, 'unknown', 'unknown code');
    equal(convertValue('UK', 'flag').value, 'UK', 'non-ISO alias unchanged');
    equal(convertValue(123, 'flag').value, 123, 'number unchanged');
    const book = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ['Name', 'Country', 'Score'], ['A', ' de ', 7], ['B', 'AIN', 3],
      ['C', null, 4], ['D', 'ZZ', 5], ['E', '  ', 8], ['F', 'JP', 9]
    ]);
    sheet.B8 = { t: 's', f: '"FR"', v: 'FR' };
    sheet.B9 = { t: 's', v: 'US', w: 'STALE', h: 'STALE', r: '<t>STALE</t>' };
    sheet['!ref'] = 'A1:C9';
    XLSX.utils.book_append_sheet(book, sheet, 'Countries');
    XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet([['DE', 42]]), 'Other');
    const original = XLSX.write(book, { type: 'array', bookType: 'xlsx' });
    for (const mode of ['flag', 'flag-code']) {
      const copy = XLSX.read(original, { type: 'array' });
      const stats = transformWorkbook(copy, { sheetName: 'Countries', column: 1, skipHeader: true, mode }, XLSX);
      equal(stats.converted, 4, `${mode} count`);
      equal(stats.unknown, 1, `${mode} unknown count`);
      equal(stats.formulas, 1, `${mode} formula count`);
      const result = XLSX.read(XLSX.write(copy, { type: 'array', bookType: 'xlsx' }), { type: 'array' });
      const out = result.Sheets.Countries;
      equal(out.B1.v, 'Country', `${mode} header retained`);
      equal(out.B2.v, mode === 'flag' ? '🇩🇪' : '🇩🇪 DE', `${mode} roundtrip flag`);
      equal(out.B3.v, mode === 'flag' ? '🏳️' : '🏳️ AIN', `${mode} roundtrip AIN`);
      equal(out.B4?.v ?? '', '', `${mode} empty cell`);
      equal(out.B5.v, 'ZZ', `${mode} unknown retained`);
      equal(out.B6?.v ?? '', '', `${mode} blank text cleared`);
      equal(out.B8.f, '"FR"', `${mode} formula retained`);
      equal(out.C2.v, 7, `${mode} other column retained`);
      equal(result.Sheets.Other.A1.v, 'DE', `${mode} other sheet retained`);
    }
    const cached = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(cached, { A1: { t: 's', v: 'FR', w: 'old', r: '<t>old</t>', h: 'old' }, '!ref': 'A1' }, 'No header');
    transformWorkbook(cached, { sheetName: 'No header', column: 0, skipHeader: false, mode: 'flag' }, XLSX);
    equal(cached.Sheets['No header'].A1.v, '🇫🇷', 'no header first row');
    equal(cached.Sheets['No header'].A1.w, undefined, 'cached display cleared');
    equal(cached.Sheets['No header'].A1.r, undefined, 'rich text cleared');
    const sparse = { SheetNames: ['Sparse'], Sheets: { Sparse: { B100000: { t: 's', v: 'US' }, '!ref': 'A1:B1048576' } } };
    equal(transformWorkbook(sparse, { sheetName: 'Sparse', column: 1, skipHeader: false, mode: 'flag' }, XLSX).converted, 1, 'sparse sheet');
    const empty = { SheetNames: ['Empty'], Sheets: { Empty: {} } };
    equal(transformWorkbook(empty, { sheetName: 'Empty', column: 0 }, XLSX).converted, 0, 'empty sheet');
    const zip = new JSZip(); zip.file('a.xlsx', original); zip.file('a_2.xlsx', original);
    const restored = await JSZip.loadAsync(await zip.generateAsync({ type: 'uint8array' }));
    equal(Object.keys(restored.files).length, 2, 'multiple files in ZIP');
    equal(XLSX.read(await restored.file('a.xlsx').async('uint8array'), { type: 'array' }).SheetNames.length, 2, 'Excel from ZIP readable');
    for (const type of ['xls', 'xlsm']) {
      const input = XLSX.write(book, { type: 'array', bookType: type });
      const parsed = XLSX.read(input, { type: 'array' });
      transformWorkbook(parsed, { sheetName: 'Countries', column: 1, skipHeader: true, mode: 'flag' }, XLSX);
      equal(parsed.Sheets.Countries.B2.v, '🇩🇪', `${type} input`);
    }
    const pdfBook = XLSX.utils.book_new();
    const pdfSheet = XLSX.utils.aoa_to_sheet([['Имя', 'Страна'], ['Иван <script>alert(1)</script>', '🇩🇪 DE'], ['Анна', '🏳️ AIN'], ['Пётр', '']]);
    pdfSheet['!merges'] = [{ s: { r: 4, c: 0 }, e: { r: 4, c: 1 } }];
    pdfSheet.A5 = { t: 's', v: 'Итог' }; pdfSheet['!ref'] = 'A1:B1048576';
    XLSX.utils.book_append_sheet(pdfBook, pdfSheet, 'Протокол');
    const html = PdfExport.createDocument(pdfBook, { sheetName: 'Протокол', skipHeader: true, title: '<title>' });
    const parsedHtml = new DOMParser().parseFromString(html, 'text/html');
    equal(parsedHtml.querySelectorAll('tr').length, 5, 'PDF ignores formatted empty tail');
    equal(parsedHtml.querySelector('thead').textContent, 'ИмяСтрана', 'PDF repeats header');
    equal(parsedHtml.querySelector('td[colspan]').colSpan, 2, 'PDF merged cell');
    equal(parsedHtml.querySelector('script'), null, 'PDF cell HTML escaped');
    equal(parsedHtml.body.textContent.includes('🇩🇪 DE'), true, 'PDF emoji text preserved');
    equal(parsedHtml.body.textContent.includes('Иван <script>'), true, 'PDF Cyrillic and literal markup');
    equal(parsedHtml.querySelector('img, canvas'), null, 'PDF no rasterized cells');
    equal(parsedHtml.querySelector('title').textContent, '<title>', 'PDF title escaped');
    const noHead = new DOMParser().parseFromString(PdfExport.createDocument(pdfBook, { sheetName: 'Протокол', skipHeader: false, title: 'Без заголовка' }), 'text/html');
    equal(noHead.querySelector('thead'), null, 'PDF no header option');
    document.getElementById('output').textContent = `PASS ${count} assertions\n${log.join('\n')}`;
    document.title = 'PASS';
  } catch (error) {
    document.getElementById('output').textContent = `FAIL\n${error.stack}\n${log.join('\n')}`;
    document.title = 'FAIL';
  }
})();
