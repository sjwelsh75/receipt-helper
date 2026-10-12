// Minimal .xlsx reader/writer that edits the sheet XML in place,
// so AP's formatting, formulas and hidden sheets are left untouched.
(function () {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  const CATEGORIES = ['Hotel', 'Meals', 'Gas', 'Entertainment', 'S&R', 'Airfare', 'Misc.'];
  const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');

  function colToNum(col) {
    let n = 0;
    for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function splitRef(ref) {
    const m = /^([A-Z]+)(\d+)$/.exec(ref);
    return { col: m[1], row: +m[2] };
  }
  function parseXml(text) {
    return new DOMParser().parseFromString(text, 'application/xml');
  }
  function serializeXml(doc) {
    let s = new XMLSerializer().serializeToString(doc);
    if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
    return s;
  }
  function textOf(el) {
    // Concatenate every <t> under el (handles rich text runs)
    const ts = el.getElementsByTagNameNS(NS, 't');
    let s = '';
    for (const t of ts) s += t.textContent;
    return s;
  }

  async function openWorkbook(arrayBuffer, sheetIndex = 0) {
    const zip = await JSZip.loadAsync(arrayBuffer);
    const wbXml = parseXml(await zip.file('xl/workbook.xml').async('string'));
    const sheets = wbXml.getElementsByTagNameNS(NS, 'sheet');
    const rid = sheets[sheetIndex].getAttributeNS(NS_REL, 'id') || sheets[sheetIndex].getAttribute('r:id');
    const relsXml = parseXml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    let target = null;
    for (const rel of relsXml.getElementsByTagName('Relationship')) {
      if (rel.getAttribute('Id') === rid) target = rel.getAttribute('Target');
    }
    const sheetPath = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const doc = parseXml(await zip.file(sheetPath).async('string'));

    let shared = [];
    const ssFile = zip.file('xl/sharedStrings.xml');
    if (ssFile) {
      const ss = parseXml(await ssFile.async('string'));
      shared = Array.from(ss.getElementsByTagNameNS(NS, 'si')).map(textOf);
    }

    // Index cells by reference
    const rows = new Map(); // rowNum -> <row>
    const cells = new Map(); // "B4" -> <c>
    for (const rowEl of doc.getElementsByTagNameNS(NS, 'row')) {
      rows.set(+rowEl.getAttribute('r'), rowEl);
      for (const c of rowEl.getElementsByTagNameNS(NS, 'c')) cells.set(c.getAttribute('r'), c);
    }

    function value(ref) {
      const c = cells.get(ref);
      if (!c) return null;
      const t = c.getAttribute('t');
      if (t === 'inlineStr') return textOf(c);
      const v = c.getElementsByTagNameNS(NS, 'v')[0];
      if (!v) return null;
      if (t === 's') return shared[+v.textContent] ?? '';
      if (t === 'str' || t === 'e') return v.textContent;
      if (t === 'b') return v.textContent === '1';
      const n = parseFloat(v.textContent);
      return isNaN(n) ? v.textContent : n;
    }
    function maxRow() { return Math.max(0, ...rows.keys()); }
    function rowCells(r) {
      const el = rows.get(r);
      return el ? Array.from(el.getElementsByTagNameNS(NS, 'c')) : [];
    }

    return { zip, wbXml, doc, sheetPath, rows, cells, value, maxRow, rowCells };
  }

  // ---------- AP credit-card sheet ----------

  function findLayout(wb) {
    for (let r = 1; r <= Math.min(20, wb.maxRow()); r++) {
      const cols = {};
      const cats = {};
      for (const c of wb.rowCells(r)) {
        const { col } = splitRef(c.getAttribute('r'));
        const h = norm(wb.value(c.getAttribute('r')));
        if (!h) continue;
        if (h === 'date') cols.date = col;
        else if (h.startsWith('place')) cols.place = col;
        else if (h === 'amount') cols.amount = col;
        else if (h.includes('purpose') || h.includes('guest')) cols.note = col;
        else if (h.startsWith('costcode')) cols.code = col;
        else if (h === 'total') cols.total = col;
        else {
          for (const cat of CATEGORIES) if (h === norm(cat)) cats[cat] = col;
        }
      }
      if (cols.code && cols.amount && cols.place) return { headerRow: r, cols, cats };
    }
    throw new Error("Couldn't find the header row (Date / Place / Amount / Cost Code) in this spreadsheet.");
  }

  function excelDateToISO(serial) {
    if (typeof serial !== 'number') return null;
    const d = new Date(Math.round((serial - 25569) * 86400000));
    return d.toISOString().slice(0, 10);
  }

  async function parseApSheet(arrayBuffer) {
    const wb = await openWorkbook(arrayBuffer);
    const layout = findLayout(wb);
    const { cols } = layout;
    const lines = [];
    let totalRow = null, balanceRow = null;
    for (let r = layout.headerRow + 1; r <= wb.maxRow(); r++) {
      const place = wb.value(cols.place + r);
      const placeStr = typeof place === 'string' ? place.trim() : '';
      if (/^total/i.test(placeStr)) { totalRow = r; continue; }
      if (/^balance/i.test(placeStr)) { balanceRow = r; continue; }
      if (totalRow) continue;
      const amount = wb.value(cols.amount + r);
      if (typeof amount !== 'number') continue;
      lines.push({
        row: r,
        date: cols.date ? excelDateToISO(wb.value(cols.date + r)) : null,
        place: placeStr,
        cents: Math.round(amount * 100),
        existingNote: cols.note ? (wb.value(cols.note + r) || '') : '',
        existingCode: cols.code ? (wb.value(cols.code + r) || '') : '',
      });
    }
    const name = wb.value('B2');
    return { layout: { ...layout, totalRow, balanceRow }, lines, employee: typeof name === 'string' ? name : '' };
  }

  // Fill note / cost code / category amount for each line.
  // fills: Map(row -> { note, code, category })
  async function fillApSheet(arrayBuffer, parsed, fills) {
    const wb = await openWorkbook(arrayBuffer);
    const { cols, cats, totalRow, balanceRow } = parsed.layout;
    const doc = wb.doc;

    function getOrCreateCell(ref) {
      let c = wb.cells.get(ref);
      if (c) return c;
      const { col, row } = splitRef(ref);
      let rowEl = wb.rows.get(row);
      if (!rowEl) {
        rowEl = doc.createElementNS(NS, 'row');
        rowEl.setAttribute('r', String(row));
        const sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
        const after = Array.from(wb.rows.keys()).sort((a, b) => a - b).find(k => k > row);
        sheetData.insertBefore(rowEl, after ? wb.rows.get(after) : null);
        wb.rows.set(row, rowEl);
      }
      c = doc.createElementNS(NS, 'c');
      c.setAttribute('r', ref);
      // keep same style as the cell above if there is one
      const above = wb.cells.get(col + (row - 1));
      if (above && above.getAttribute('s')) c.setAttribute('s', above.getAttribute('s'));
      const next = Array.from(rowEl.getElementsByTagNameNS(NS, 'c'))
        .find(x => colToNum(splitRef(x.getAttribute('r')).col) > colToNum(col));
      rowEl.insertBefore(c, next || null);
      wb.cells.set(ref, c);
      return c;
    }
    function clearChildren(c) { while (c.firstChild) c.removeChild(c.firstChild); }
    function setString(ref, text) {
      const c = getOrCreateCell(ref);
      clearChildren(c);
      if (!text) { c.removeAttribute('t'); return; }
      c.setAttribute('t', 'inlineStr');
      const is = doc.createElementNS(NS, 'is');
      const t = doc.createElementNS(NS, 't');
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      t.textContent = text;
      is.appendChild(t);
      c.appendChild(is);
    }
    function setNumber(ref, n) {
      const c = getOrCreateCell(ref);
      clearChildren(c);
      c.removeAttribute('t');
      const v = doc.createElementNS(NS, 'v');
      v.textContent = String(n);
      c.appendChild(v);
    }
    // Update the cached result of an existing formula so previews show correct totals
    function setCachedValue(ref, n) {
      const c = wb.cells.get(ref);
      if (!c || !c.getElementsByTagNameNS(NS, 'f')[0]) return;
      c.removeAttribute('t');
      let v = c.getElementsByTagNameNS(NS, 'v')[0];
      if (!v) { v = doc.createElementNS(NS, 'v'); c.appendChild(v); }
      v.textContent = String(n);
    }
    const money = cents => Math.round(cents) / 100;

    const catTotals = {};
    let grand = 0, amountTotal = 0;
    for (const line of parsed.lines) {
      amountTotal += line.cents;
      const f = fills.get(line.row);
      let rowTotal = 0;
      if (f) {
        if (cols.note) setString(cols.note + line.row, f.note || '');
        if (cols.code) setString(cols.code + line.row, f.code || '');
        // clear every category column, then put the amount in the chosen one
        for (const [cat, col] of Object.entries(cats)) {
          if (cat === f.category) {
            setNumber(col + line.row, money(line.cents));
            catTotals[cat] = (catTotals[cat] || 0) + line.cents;
            rowTotal = line.cents;
          } else if (wb.cells.get(col + line.row)) {
            setString(col + line.row, '');
          }
        }
      }
      grand += rowTotal;
      if (cols.total) setCachedValue(cols.total + line.row, money(rowTotal));
    }
    if (totalRow) {
      setCachedValue(cols.amount + totalRow, money(amountTotal));
      for (const [cat, col] of Object.entries(cats)) setCachedValue(col + totalRow, money(catTotals[cat] || 0));
      if (cols.total) setCachedValue(cols.total + totalRow, money(grand));
    }
    if (balanceRow && cols.total) setCachedValue(cols.total + balanceRow, money(grand));

    wb.zip.file(wb.sheetPath, serializeXml(doc));

    // Ask Excel to recalculate everything when the file is opened
    let wbText = await wb.zip.file('xl/workbook.xml').async('string');
    if (/<calcPr\b/.test(wbText)) {
      wbText = wbText.replace(/<calcPr\b([^>]*?)(\/?)>/, (m, attrs, slash) =>
        '<calcPr' + attrs.replace(/\s*fullCalcOnLoad="[^"]*"/, '') + ' fullCalcOnLoad="1"' + slash + '>');
    } else {
      wbText = wbText.replace(/(<\/definedNames>|<\/sheets>(?![\s\S]*<\/definedNames>))/, '$1<calcPr fullCalcOnLoad="1"/>');
    }
    wb.zip.file('xl/workbook.xml', wbText);

    return wb.zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  // ---------- Cost code list ----------
  // Column A = name, column B = code; extra code-looking values in later columns become "(alt)" entries.
  async function parseCostCodes(arrayBuffer) {
    const wb = await openWorkbook(arrayBuffer);
    const isCode = v => /^\d+(\.\d+)+$/.test(String(v ?? '').trim());
    const out = [];
    for (let r = 1; r <= wb.maxRow(); r++) {
      const name = String(wb.value('A' + r) ?? '').trim();
      const code = String(wb.value('B' + r) ?? '').trim();
      if (!isCode(code)) continue;
      out.push({ name: name || '(no name)', code });
      for (const c of wb.rowCells(r)) {
        const { col } = splitRef(c.getAttribute('r'));
        if (colToNum(col) <= 2) continue;
        const v = String(wb.value(c.getAttribute('r')) ?? '').trim();
        if (isCode(v)) out.push({ name: (name || '(no name)') + ' (alt)', code: v });
      }
    }
    return out;
  }

  window.XLSX_HELPER = { CATEGORIES, parseApSheet, fillApSheet, parseCostCodes };
})();
