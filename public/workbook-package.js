/* global JSZip, FlagConverter, XLSX */
(() => {
  'use strict';
  const all = (node, name) => node ? [...node.getElementsByTagNameNS('*', name)] : [];
  const first = (node, name) => all(node, name)[0];
  const children = (node, name) => node ? [...node.children].filter(child => child.localName === name) : [];
  const value = (node, name, fallback = '') => first(node, name)?.getAttribute('val') ?? fallback;
  const enabled = (node, name) => !!first(node, name) && !['0', 'false'].includes(value(node, name));
  const parse = text => {
    const xml = new DOMParser().parseFromString(text, 'application/xml');
    if (all(xml, 'parsererror').length) throw new Error('Не удалось прочитать XML внутри Excel.');
    return xml;
  };
  function resolve(base, target) {
    return new URL(target, 'https://workbook.local/' + base).pathname.slice(1);
  }
  const relPath = path => path.replace(/([^/]+)$/, '_rels/$1.rels');
  async function xmlAt(zip, path) {
    const file = zip.file(path);
    return file ? parse(await file.async('string')) : null;
  }
  async function relationships(zip, path) {
    const xml = await xmlAt(zip, relPath(path));
    return new Map(all(xml, 'Relationship').filter(rel => rel.getAttribute('TargetMode') !== 'External')
      .map(rel => [rel.getAttribute('Id'), { path: resolve(path, rel.getAttribute('Target')), type: rel.getAttribute('Type') }]));
  }
  const relationId = node => node?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');

  async function open(bytes, sheetName) {
    const zip = await JSZip.loadAsync(bytes);
    const workbook = await xmlAt(zip, 'xl/workbook.xml');
    const rels = await relationships(zip, 'xl/workbook.xml');
    const sheet = all(workbook, 'sheet').find(el => el.getAttribute('name') === sheetName);
    const path = rels.get(relationId(sheet))?.path;
    if (!path) throw new Error('Не найден исходный лист Excel.');
    return { zip, path, xml: await xmlAt(zip, path), rels };
  }

  async function write(bytes, workbook, settings) {
    const source = await open(bytes, settings.sheetName);
    const sheet = workbook.Sheets[settings.sheetName];
    const startRow = XLSX.utils.decode_range(sheet['!ref'] || 'A1').s.r + (settings.skipHeader ? 1 : 0);
    for (const cell of all(source.xml, 'c')) {
      const address = cell.getAttribute('r'), pos = XLSX.utils.decode_cell(address);
      const original = sheet[address];
      if (pos.c !== settings.column || pos.r < startRow || !original || original.f || original.F) continue;
      const result = FlagConverter.convertValue(original.v, settings.mode);
      if (result.kind === 'unknown' || (result.kind === 'empty' && original.v == null)) continue;
      // Preserve style index, extensions, drawings and every other ZIP part.
      for (const child of [...cell.children]) if (['v', 'is'].includes(child.localName)) child.remove();
      cell.setAttribute('t', 'inlineStr');
      const inline = source.xml.createElementNS(cell.namespaceURI, 'is');
      const text = source.xml.createElementNS(cell.namespaceURI, 't');
      text.textContent = result.value; inline.append(text); cell.insertBefore(inline, children(cell, 'extLst')[0] || null);
    }
    source.zip.file(source.path, new XMLSerializer().serializeToString(source.xml));
    return source.zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  async function visuals(bytes, sheetName) {
    const { zip, path, xml, rels } = await open(bytes, sheetName);
    const byType = suffix => [...rels.values()].find(rel => rel.type.endsWith('/' + suffix))?.path;
    const stylesXml = await xmlAt(zip, byType('styles') || 'xl/styles.xml');
    const theme = await xmlAt(zip, byType('theme') || 'xl/theme/theme1.xml');
    const scheme = first(theme, 'clrScheme');
    // Spreadsheet theme indexes use light/dark order, unlike clrScheme XML order.
    const themeOrder = ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
    const themeColors = themeOrder.map(name => {
      const el = first(scheme, name)?.firstElementChild;
      return el?.getAttribute('lastClr') || el?.getAttribute('val');
    });
    const indexed = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080'];
    function color(node) {
      if (!node) return '';
      let rgb = node.getAttribute('rgb')?.slice(-6) || themeColors[Number(node.getAttribute('theme') ?? -1)] || indexed[Number(node.getAttribute('indexed') ?? -1)];
      if (!/^[\da-f]{6}$/i.test(rgb || '')) return '';
      const tint = Number(node.getAttribute('tint') || 0);
      if (tint) rgb = rgb.match(/../g).map(hex => {
        const n = parseInt(hex, 16);
        return Math.round(tint < 0 ? n * (1 + tint) : n + (255 - n) * tint).toString(16).padStart(2, '0');
      }).join('');
      return '#' + rgb;
    }
    function font(node) {
      if (!node) return {};
      const css = {};
      const family = value(node, 'name') || value(node, 'rFont');
      if (family) css.fontFamily = `"Print Flags", ${JSON.stringify(family)}, "Segoe UI", Arial, sans-serif`;
      const size = Number(value(node, 'sz'));
      if (size > 0 && size < 500) css.fontSize = size + 'pt';
      css.fontWeight = enabled(node, 'b') ? '700' : '400';
      css.fontStyle = enabled(node, 'i') ? 'italic' : 'normal';
      const decorations = [];
      if (enabled(node, 'u')) decorations.push('underline');
      if (enabled(node, 'strike')) decorations.push('line-through');
      css.textDecoration = decorations.join(' ') || 'none';
      if (color(first(node, 'color'))) css.color = color(first(node, 'color'));
      const vert = value(node, 'vertAlign');
      if (vert === 'superscript' || vert === 'subscript') css.verticalAlign = vert === 'superscript' ? 'super' : 'sub';
      return css;
    }
    const fonts = children(first(stylesXml, 'fonts'), 'font');
    const fills = children(first(stylesXml, 'fills'), 'fill');
    const borders = children(first(stylesXml, 'borders'), 'border');
    const styleList = children(first(stylesXml, 'cellXfs'), 'xf').map(xf => {
      const css = { ...font(fonts[Number(xf.getAttribute('fontId') || 0)]), border: 'none', backgroundColor: 'transparent' };
      const align = first(xf, 'alignment');
      const horizontal = align?.getAttribute('horizontal'), vertical = align?.getAttribute('vertical');
      if (['left','right','center','justify'].includes(horizontal)) css.textAlign = horizontal;
      if (['top','bottom','center'].includes(vertical)) css.verticalAlign = vertical === 'center' ? 'middle' : vertical;
      const indent = Number(align?.getAttribute('indent') || 0);
      if (indent) css.paddingLeft = `${Math.min(indent, 20)}em`;
      const pattern = first(fills[Number(xf.getAttribute('fillId') || 0)], 'patternFill');
      if (pattern?.getAttribute('patternType') === 'solid') css.backgroundColor = color(first(pattern, 'fgColor')) || 'transparent';
      const border = borders[Number(xf.getAttribute('borderId') || 0)];
      for (const side of ['top','right','bottom','left']) {
        const line = children(border, side)[0], kind = line?.getAttribute('style');
        if (!kind) continue;
        const type = /dash|dot/i.test(kind) ? (/dot/i.test(kind) ? 'dotted' : 'dashed') : kind === 'double' ? 'double' : 'solid';
        const width = /thick/i.test(kind) ? 2 : /medium|double/i.test(kind) ? 1.5 : .6;
        css['border' + side[0].toUpperCase() + side.slice(1)] = `${width}pt ${type} ${color(first(line, 'color')) || '#333'}`;
      }
      return css;
    });
    const sharedXml = await xmlAt(zip, byType('sharedStrings') || 'xl/sharedStrings.xml');
    const shared = all(sharedXml, 'si');
    const cells = {}, rows = {}, images = [];
    const columnStyles = all(xml, 'col');
    for (const row of all(xml, 'row')) {
      const r = Number(row.getAttribute('r')) - 1;
      rows[r] = { height: Number(row.getAttribute('ht')) || 0 };
      for (const cell of children(row, 'c')) {
        const address = cell.getAttribute('r'), column = XLSX.utils.decode_cell(address).c + 1;
        const col = columnStyles.find(el => Number(el.getAttribute('min')) <= column && Number(el.getAttribute('max')) >= column);
        const styleIndex = Number(cell.getAttribute('s') ?? row.getAttribute('s') ?? col?.getAttribute('style') ?? 0);
        const string = cell.getAttribute('t') === 's' ? shared[Number(first(cell, 'v')?.textContent)] : first(cell, 'is');
        const runs = children(string, 'r').map(run => ({ text: first(run, 't')?.textContent || '', css: font(first(run, 'rPr')) }));
        cells[address] = { css: styleList[styleIndex] || {}, runs, raw: all(string, 't').map(t => t.textContent).join('') };
      }
    }
    const sheetRels = await relationships(zip, path);
    const mediaCache = new Map();
    async function mediaData(mediaPath) {
      if (mediaCache.has(mediaPath)) return mediaCache.get(mediaPath);
      const file = zip.file(mediaPath);
      if (!file) throw new Error('Не найдено встроенное изображение: ' + mediaPath);
      const ext = mediaPath.split('.').pop().toLowerCase();
      const mime = { png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif', bmp:'image/bmp', webp:'image/webp', svg:'image/svg+xml', tif:'image/tiff', tiff:'image/tiff', emf:'image/emf', wmf:'image/wmf' }[ext];
      if (!mime) throw new Error('Неизвестный формат изображения: ' + ext);
      const data = `data:${mime};base64,${await file.async('base64')}`;
      mediaCache.set(mediaPath, data); return data;
    }
    function coordinate(emu, axis) {
      let pixels = Number(emu) / 9525, index = 0;
      while (pixels > 0 && index < (axis === 'row' ? 1048576 : 16384)) {
        const col = columnStyles.find(el => Number(el.getAttribute('min')) <= index + 1 && Number(el.getAttribute('max')) >= index + 1);
        const size = axis === 'row' ? (rows[index]?.height || 15) * 96 / 72 : (Number(col?.getAttribute('width')) || 8.43) * 7 + 5;
        if (pixels < size) break;
        pixels -= size; index++;
      }
      return index;
    }
    for (const drawing of all(xml, 'drawing')) {
      const drawingPath = sheetRels.get(relationId(drawing))?.path;
      if (!drawingPath) continue;
      const drawingXml = await xmlAt(zip, drawingPath), drawingRels = await relationships(zip, drawingPath);
      for (const anchor of [...drawingXml.documentElement.children]) {
        const from = children(anchor, 'from')[0], position = children(anchor, 'pos')[0];
        const r = from ? Number(first(from, 'row')?.textContent || 0) : coordinate(position?.getAttribute('y') || 0, 'row');
        const c = from ? Number(first(from, 'col')?.textContent || 0) : coordinate(position?.getAttribute('x') || 0, 'col');
        for (const pic of all(anchor, 'pic')) {
          const blip = first(pic, 'blip');
          const id = blip?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed');
          const mediaPath = drawingRels.get(id)?.path;
          if (!mediaPath) throw new Error('PDF не поддерживает внешние картинки. Вставьте их в Excel как встроенные изображения.');
          const ext = first(first(pic, 'xfrm'), 'ext') || children(anchor, 'ext')[0];
          images.push({ r, c, src: await mediaData(mediaPath), width: Number(ext?.getAttribute('cx') || 0) / 9525,
            height: Number(ext?.getAttribute('cy') || 0) / 9525, alt: first(pic, 'cNvPr')?.getAttribute('descr') || first(pic, 'cNvPr')?.getAttribute('name') || 'Изображение из Excel' });
        }
      }
    }
    return { cells, rows, images, defaultStyle: styleList[0] || {} };
  }
  globalThis.WorkbookPackage = { write, visuals };
})();
