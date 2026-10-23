/* global XLSX, JSZip, FlagConverter, PdfExport, WorkbookPackage */
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
    equal(parsedHtml.querySelectorAll('tr:not(.page-spacer)').length, 5, 'PDF ignores formatted empty tail');
    equal(parsedHtml.querySelector('thead').textContent, '', 'PDF does not repeat document heading');
    equal(parsedHtml.querySelector('tbody tr').textContent, 'ИмяСтрана', 'PDF keeps heading in first body row');
    equal(parsedHtml.querySelector('tbody td[colspan]').colSpan, 2, 'PDF merged cell');
    equal(parsedHtml.querySelector('script'), null, 'PDF cell HTML escaped');
    equal(parsedHtml.body.textContent.includes('🇩🇪 DE'), true, 'PDF emoji text preserved');
    equal(parsedHtml.body.textContent.includes('Иван <script>'), true, 'PDF Cyrillic and literal markup');
    equal(parsedHtml.querySelector('img, canvas'), null, 'PDF no rasterized cells');
    equal(parsedHtml.querySelector('title').textContent, '<title>', 'PDF title escaped');
    const noHead = new DOMParser().parseFromString(PdfExport.createDocument(pdfBook, { sheetName: 'Протокол', skipHeader: false, title: 'Без заголовка' }), 'text/html');
    equal(noHead.querySelector('th'), null, 'PDF no header option');
    const sourceZip = await JSZip.loadAsync(original);
    const stylesPath = 'xl/styles.xml';
    // A real OOXML drawing with image-only row, cell styles and partial rich text.
    const ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const relNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    sourceZip.file(stylesPath, `<styleSheet xmlns="${ns}"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><i/><u/><sz val="18"/><color rgb="FFFF0000"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border><bottom style="thin"><color rgb="FF0000FF"/></bottom></border></borders><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="1" fillId="0" borderId="0"><alignment horizontal="center"/></xf></cellXfs></styleSheet>`);
    sourceZip.file('xl/worksheets/sheet1.xml', `<worksheet xmlns="${ns}" xmlns:r="${relNs}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Header</t></is></c></row><row r="2"><c r="A2" s="1" t="inlineStr"><is><t>DE</t></is></c><c r="B2" t="inlineStr"><is><r><rPr><b/><sz val="16"/></rPr><t>Жирный</t></r><r><rPr><i/></rPr><t> курсив</t></r></is></c></row><row r="3"><c r="A3" s="1"/></row><row r="4" ht="24"><c r="B4" s="1"/></row><row r="5"><c r="A5" t="inlineStr"><is><t>AIN</t></is></c></row></sheetData><drawing r:id="img"/></worksheet>`);
    sourceZip.file('xl/worksheets/_rels/sheet1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="img" Type="${relNs}/drawing" Target="../drawings/drawing1.xml"/></Relationships>`);
    sourceZip.file('xl/drawings/drawing1.xml', `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${relNs}"><xdr:oneCellAnchor><xdr:from><xdr:col>2</xdr:col><xdr:row>3</xdr:row></xdr:from><xdr:ext cx="152400" cy="152400"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr name="Test picture"/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="pic"/></xdr:blipFill></xdr:pic></xdr:oneCellAnchor></xdr:wsDr>`);
    sourceZip.file('xl/drawings/_rels/drawing1.xml.rels', `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="pic" Type="${relNs}/image" Target="../media/test.png"/></Relationships>`);
    sourceZip.file('xl/media/test.png', 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=', { base64: true });
    const sourceBytes = await sourceZip.generateAsync({ type: 'uint8array' });
    const sourceBook = XLSX.read(sourceBytes, { type: 'array', cellStyles: true });
    const settings = { sheetName: 'Countries', column: 0, skipHeader: true, mode: 'flag-code' };
    const savedBytes = await WorkbookPackage.write(sourceBytes, sourceBook, settings);
    const savedZip = await JSZip.loadAsync(savedBytes);
    for (const [path, file] of Object.entries(sourceZip.files)) {
      if (file.dir || path === 'xl/worksheets/sheet1.xml') continue;
      equal(await savedZip.file(path).async('base64'), await file.async('base64'), `XLSX preserves ${path}`);
    }
    const savedBook = XLSX.read(savedBytes, { type: 'array' });
    equal(savedBook.Sheets.Countries.A2.v, '🇩🇪 DE', 'package replaces country');
    equal(savedBook.Sheets.Countries.A5.v, '🏳️ AIN', 'package replaces AIN');
    const savedXml = new DOMParser().parseFromString(await savedZip.file('xl/worksheets/sheet1.xml').async('string'), 'application/xml');
    equal(savedXml.querySelector('c[r="A2"]').getAttribute('s'), '1', 'package keeps style index');
    const visuals = await WorkbookPackage.visuals(sourceBytes, 'Countries');
    equal(visuals.cells.A2.css.fontWeight, '700', 'original bold style');
    equal(visuals.cells.A2.css.fontSize, '18pt', 'original font size');
    equal(visuals.cells.A2.css.color, '#FF0000', 'original font color');
    equal(visuals.cells.A2.css.backgroundColor, '#FFFF00', 'original fill');
    equal(visuals.cells.A2.css.fontStyle, 'italic', 'original italic');
    equal(visuals.cells.A2.css.textDecoration, 'underline', 'original underline');
    equal(visuals.images.length, 1, 'drawing picture extracted');
    equal(visuals.images[0].c, 2, 'picture outside data columns retained');
    const styledHtml = new DOMParser().parseFromString(PdfExport.createDocument(sourceBook, { sheetName: 'Countries', skipHeader: true, title: 'Styled', visuals }), 'text/html');
    equal(styledHtml.querySelectorAll('tr[data-source-row]').length, 5, 'PDF keeps interior blank and image-only rows');
    equal(styledHtml.querySelector('tr[data-source-row="3"]').classList.contains('section-separator'), true, 'blank row acts as section separator');
    equal(styledHtml.querySelectorAll('img').length, 1, 'PDF includes picture');
    equal(styledHtml.querySelector('tr[data-source-row="2"] td').style.fontSize, '13.86pt', 'PDF scales original font size for compact printing');
    equal(styledHtml.querySelector('tr[data-source-row="2"] td').style.fontWeight, '700', 'PDF applies bold');
    equal(styledHtml.querySelectorAll('tr[data-source-row="2"] span').length, 2, 'PDF rich text runs');
    equal(styledHtml.querySelector('span').style.fontSize, '12.32pt', 'PDF scales rich text size proportionally');
    document.getElementById('output').textContent = `PASS ${count} assertions\n${log.join('\n')}`;
    document.title = 'PASS';
  } catch (error) {
    document.getElementById('output').textContent = `FAIL\n${error.stack}\n${log.join('\n')}`;
    document.title = 'FAIL';
  }
})();
