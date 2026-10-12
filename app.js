(function () {
  'use strict';
  const { CATEGORIES, parseApSheet, fillApSheet, parseCostCodes } = window.XLSX_HELPER;
  const $ = sel => document.querySelector(sel);
  const el = (tag, attrs = {}, ...kids) => {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) {
      if (k === 'class') e.className = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
      else if (v !== false && v != null) e.setAttribute(k, v === true ? '' : v);
    }
    for (const kid of kids.flat()) if (kid != null) e.append(kid.nodeType ? kid : String(kid));
    return e;
  };

  // ---------- storage (IndexedDB) ----------
  const DB = {
    db: null,
    open() {
      return new Promise((res, rej) => {
        const req = indexedDB.open('receipt-helper', 1);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains('receipts')) db.createObjectStore('receipts', { keyPath: 'id' });
          if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv', { keyPath: 'key' });
        };
        req.onsuccess = () => { this.db = req.result; res(); };
        req.onerror = () => rej(req.error);
      });
    },
    req(store, mode, fn) {
      return new Promise((res, rej) => {
        const tx = this.db.transaction(store, mode);
        const r = fn(tx.objectStore(store));
        tx.oncomplete = () => res(r && r.result);
        tx.onerror = () => rej(tx.error);
        tx.onabort = () => rej(tx.error);
      });
    },
    all(store) { return this.req(store, 'readonly', s => s.getAll()); },
    put(store, v) { return this.req(store, 'readwrite', s => s.put(v)); },
    del(store, k) { return this.req(store, 'readwrite', s => s.delete(k)); },
    async getKv(key, dflt) {
      const r = await this.req('kv', 'readonly', s => s.get(key));
      return r ? r.value : dflt;
    },
    setKv(key, value) { return this.put('kv', { key, value }); },
  };

  // ---------- state ----------
  let receipts = [];
  let codes = [];
  let opts = { caption: true, placeholder: true };
  let lastCode = '';

  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const fmtMoney = cents => '$' + (cents / 100).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const parseMoney = s => {
    const n = parseFloat(String(s).replace(/[$,\s]/g, ''));
    return isNaN(n) ? null : Math.round(n * 100);
  };
  const todayISO = () => {
    const d = new Date();
    return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  };
  const fmtDate = iso => {
    if (!iso) return '';
    const [y, m, d] = iso.split('-');
    return `${+m}/${+d}/${y.slice(2)}`;
  };
  const dayDiff = (a, b) => (a && b) ? Math.abs((Date.parse(a) - Date.parse(b)) / 86400000) : 999;
  const codeLabel = code => {
    const c = codes.find(x => x.code === code);
    return c ? `${c.name} (${c.code})` : code;
  };

  function toast(msg, ms = 2500) {
    const t = $('#toast');
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { t.hidden = true; }, ms);
  }
  function busy(msg) {
    $('#busy-msg').textContent = msg || 'Working…';
    $('#busy').hidden = !msg;
  }
  const readFile = file => file.arrayBuffer();

  // ---------- images ----------
  function loadImage(blob) {
    return new Promise((res, rej) => {
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => { URL.revokeObjectURL(url); res(img); };
      img.onerror = () => { URL.revokeObjectURL(url); rej(new Error("Couldn't read that image.")); };
      img.src = url;
    });
  }
  // Shrink to a reasonable size and convert to JPEG (also handles HEIC/PNG screenshots)
  async function normalizeImage(file) {
    const img = await loadImage(file);
    const max = 2000;
    const scale = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
    const w = Math.round(img.naturalWidth * scale), h = Math.round(img.naturalHeight * scale);
    const canvas = document.createElement('canvas');
    canvas.width = w; canvas.height = h;
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, w, h);
    ctx.drawImage(img, 0, 0, w, h);
    const blob = await new Promise(r => canvas.toBlob(r, 'image/jpeg', 0.82));
    return { type: 'image/jpeg', data: await blob.arrayBuffer(), w, h };
  }
  async function fileToPage(file) {
    if (file.type === 'application/pdf' || /\.pdf$/i.test(file.name)) {
      return { type: 'application/pdf', data: await readFile(file) };
    }
    return normalizeImage(file);
  }

  async function filesToPages(files) {
    busy('Processing…');
    try {
      const pages = [];
      for (const f of files) pages.push(await fileToPage(f));
      return pages;
    } catch (err) {
      alert(err.message);
      return [];
    } finally { busy(); }
  }
  // Drag-and-drop files onto an element (computer only; harmless on phone)
  function enableDrop(target, onFiles) {
    target.addEventListener('dragover', e => {
      if (!e.dataTransfer.types.includes('Files')) return;
      e.preventDefault();
      target.classList.add('dragover');
    });
    target.addEventListener('dragleave', e => {
      if (!target.contains(e.relatedTarget)) target.classList.remove('dragover');
    });
    target.addEventListener('drop', e => {
      e.preventDefault();
      e.stopPropagation();
      target.classList.remove('dragover');
      const files = Array.from(e.dataTransfer.files || []);
      if (files.length) onFiles(files);
    });
  }
  // Dropping a file anywhere else shouldn't navigate away from the app
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => e.preventDefault());

  // thumbnails: object URLs cached per receipt page
  const urlCache = new WeakMap();
  function pageUrl(page) {
    if (page.type !== 'image/jpeg') return null;
    let u = urlCache.get(page);
    if (!u) { u = URL.createObjectURL(new Blob([page.data], { type: page.type })); urlCache.set(page, u); }
    return u;
  }
  function thumbEl(rec, cls = 'thumb') {
    const p = rec.files && rec.files[0];
    const t = el('div', { class: cls });
    if (!p) t.textContent = 'No image';
    else if (p.type === 'application/pdf') t.textContent = 'PDF';
    else t.style.backgroundImage = `url("${pageUrl(p)}")`;
    return t;
  }

  // ---------- tabs ----------
  const titles = { receipts: 'Receipts', monthend: 'Month-end', settings: 'Settings' };
  function showTab(name) {
    document.querySelectorAll('.tab').forEach(s => s.classList.toggle('active', s.id === 'tab-' + name));
    document.querySelectorAll('.tabbar button').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
    $('#title').textContent = titles[name];
    window.scrollTo(0, 0);
  }
  document.querySelectorAll('.tabbar button').forEach(b => b.addEventListener('click', () => showTab(b.dataset.tab)));

  // ---------- receipts list ----------
  function receiptCard(r) {
    return el('button', { class: 'rcpt', onclick: () => openEditor(r) },
      thumbEl(r),
      el('div', { class: 'info' },
        el('div', { class: 'row1' }, el('span', {}, r.vendor || 'Receipt'), el('span', {}, fmtMoney(r.cents))),
        el('div', { class: 'row2' }, `${fmtDate(r.date)} · ${r.category} · ${r.code ? codeLabel(r.code) : 'no cost code'}`),
        r.note ? el('div', { class: 'row3' }, r.note) : null));
  }
  function renderReceipts() {
    const byDate = (a, b) => (b.date || '').localeCompare(a.date || '') || b.created - a.created;
    const pending = receipts.filter(r => !r.submittedAt).sort(byDate);
    const submitted = receipts.filter(r => r.submittedAt).sort(byDate);
    const list = $('#pending-list');
    list.replaceChildren();
    if (!pending.length) {
      list.append(el('div', { class: 'empty' }, 'No receipts yet.', el('br'), 'Tap + to snap your first one.'));
    }
    pending.forEach(r => list.append(receiptCard(r)));
    const total = pending.reduce((s, r) => s + r.cents, 0);
    $('#pending-summary').textContent = pending.length
      ? `${pending.length} receipt${pending.length === 1 ? '' : 's'} waiting · ${fmtMoney(total)}` : '';
    $('#btn-submit-all').hidden = !pending.length;
    $('#submitted-wrap').hidden = !submitted.length;
    $('#submitted-count').textContent = `(${submitted.length})`;
    $('#submitted-list').replaceChildren(...submitted.map(receiptCard));
  }
  $('#btn-submit-all').addEventListener('click', async () => {
    const pending = receipts.filter(r => !r.submittedAt);
    if (!confirm(`Mark all ${pending.length} waiting receipt(s) as submitted?

Use this after you've done month-end on your computer.`)) return;
    const now = Date.now();
    for (const r of pending) { r.submittedAt = now; await DB.put('receipts', r); }
    renderReceipts();
    toast('Marked as submitted');
  });
  $('#btn-clear-submitted').addEventListener('click', async () => {
    const sub = receipts.filter(r => r.submittedAt);
    if (!confirm(`Permanently delete ${sub.length} submitted receipt(s) from this phone?`)) return;
    for (const r of sub) await DB.del('receipts', r.id);
    receipts = receipts.filter(r => !r.submittedAt);
    renderReceipts();
  });

  // ---------- cost code <select> ----------
  function fillCodeSelect(select, current) {
    select.replaceChildren(el('option', { value: '' }, '— Select cost code —'));
    const sorted = [...codes].sort((a, b) => a.name.localeCompare(b.name));
    if (current && !codes.some(c => c.code === current)) select.append(el('option', { value: current }, current));
    sorted.forEach(c => select.append(el('option', { value: c.code }, `${c.name} — ${c.code}`)));
    select.append(el('option', { value: '__new' }, '+ Add new cost code…'));
    select.value = current || '';
  }
  async function saveCodes() {
    await DB.setKv('codes', codes);
    renderCodes();
  }

  // ---------- receipt editor ----------
  const ed = { id: null, pages: [], onSaved: null };
  const editor = $('#editor');
  $('#ed-category').replaceChildren(...CATEGORIES.map(c => el('option', { value: c }, c)));

  function renderPages() {
    const wrap = $('#ed-pages');
    wrap.replaceChildren(...ed.pages.map((p, i) => {
      const d = el('div', { class: 'page' });
      if (p.type === 'application/pdf') d.textContent = 'PDF';
      else d.style.backgroundImage = `url("${pageUrl(p)}")`;
      d.append(el('button', { type: 'button', class: 'rm', 'aria-label': 'Remove page', onclick: () => { ed.pages.splice(i, 1); renderPages(); } }, '×'));
      return d;
    }));
    $('#ed-add-page-label').textContent = ed.pages.length ? '➕ Add another page' : '📷 Take photo / choose file';
  }

  function openEditor(rec, prefill = {}, onSaved = null) {
    ed.id = rec ? rec.id : null;
    ed.pages = rec ? [...rec.files] : (prefill.pages || []);
    ed.onSaved = onSaved;
    const r = rec || {};
    $('#ed-title').textContent = rec ? 'Edit receipt' : 'New receipt';
    $('#ed-amount').value = r.cents != null ? (r.cents / 100).toFixed(2) : (prefill.cents != null ? (prefill.cents / 100).toFixed(2) : '');
    $('#ed-date').value = r.date || prefill.date || todayISO();
    $('#ed-vendor').value = r.vendor || prefill.vendor || '';
    $('#ed-category').value = r.category || 'Misc.';
    fillCodeSelect($('#ed-code'), r.code != null ? r.code : lastCode);
    $('#ed-newcode').hidden = true;
    $('#ed-newcode-name').value = '';
    $('#ed-newcode-code').value = '';
    $('#ed-note').value = r.note || '';
    $('#ed-delete').hidden = !rec;
    renderPages();
    editor.showModal();
    editor.querySelector('.dlg-body').scrollTop = 0;
  }

  $('#btn-add').addEventListener('click', () => {
    openEditor(null);
    // jump straight to the camera / file picker
    $('#ed-file').click();
  });
  $('#ed-cancel').addEventListener('click', () => editor.close());
  async function addEditorFiles(files) {
    if (!files.length) return;
    ed.pages.push(...await filesToPages(files));
    renderPages();
  }
  $('#ed-file').addEventListener('change', e => {
    const files = Array.from(e.target.files || []);
    e.target.value = '';
    addEditorFiles(files);
  });
  enableDrop(editor, addEditorFiles);
  $('#ed-category').addEventListener('change', () => {
    const cat = $('#ed-category').value;
    const pat = cat === 'Meals' ? /meal/i : (cat === 'Hotel' || cat === 'Airfare') ? /travel/i : null;
    const c = pat && codes.find(x => pat.test(x.name));
    if (c) $('#ed-code').value = c.code;
  });
  $('#ed-code').addEventListener('change', () => {
    const isNew = $('#ed-code').value === '__new';
    $('#ed-newcode').hidden = !isNew;
    if (isNew) $('#ed-newcode-name').focus();
  });
  $('#ed-delete').addEventListener('click', async () => {
    if (!confirm('Delete this receipt?')) return;
    await DB.del('receipts', ed.id);
    receipts = receipts.filter(r => r.id !== ed.id);
    editor.close();
    renderReceipts();
    if (ap) renderMatch();
  });
  $('#editor-form').addEventListener('submit', async e => {
    e.preventDefault();
    const cents = parseMoney($('#ed-amount').value);
    if (cents == null) { alert('Enter the amount.'); return; }
    let code = $('#ed-code').value;
    if (code === '__new') {
      const name = $('#ed-newcode-name').value.trim();
      code = $('#ed-newcode-code').value.trim();
      if (!name || !code) { alert('Enter a name and a code for the new cost code.'); return; }
      if (!codes.some(c => c.code === code && c.name === name)) {
        codes.push({ name, code });
        await saveCodes();
      }
    }
    if (!ed.pages.length && !confirm('Save without a picture?')) return;
    const existing = receipts.find(r => r.id === ed.id);
    const rec = {
      id: ed.id || uid(),
      created: existing ? existing.created : Date.now(),
      submittedAt: existing ? existing.submittedAt : null,
      cents,
      date: $('#ed-date').value,
      vendor: $('#ed-vendor').value.trim(),
      category: $('#ed-category').value,
      code,
      note: $('#ed-note').value.trim(),
      files: ed.pages,
    };
    try {
      await DB.put('receipts', rec);
    } catch (err) {
      alert('Could not save: ' + err.message);
      return;
    }
    receipts = receipts.filter(r => r.id !== rec.id).concat(rec);
    if (code) { lastCode = code; DB.setKv('lastCode', code); }
    editor.close();
    renderReceipts();
    toast('Saved');
    if (ed.onSaved) ed.onSaved(rec);
    else if (ap) renderMatch();
  });

  // ---------- settings: cost codes ----------
  function renderCodes() {
    const list = $('#code-list');
    const sorted = [...codes].sort((a, b) => a.name.localeCompare(b.name));
    list.replaceChildren(...sorted.map(c =>
      el('button', { class: 'code-row', onclick: () => openCodeEditor(c) }, el('span', {}, c.name), el('span', { class: 'c' }, c.code))));
    if (!codes.length) list.append(el('p', { class: 'muted' }, 'No cost codes yet. Import your Cost Codes spreadsheet or add them one at a time.'));
  }
  let editingCode = null;
  const codeEditor = $('#code-editor');
  function openCodeEditor(c) {
    editingCode = c;
    $('#ce-title').textContent = c ? 'Edit cost code' : 'New cost code';
    $('#ce-name').value = c ? c.name : '';
    $('#ce-code').value = c ? c.code : '';
    $('#ce-delete').hidden = !c;
    codeEditor.showModal();
  }
  $('#btn-add-code').addEventListener('click', () => openCodeEditor(null));
  $('#ce-cancel').addEventListener('click', () => codeEditor.close());
  $('#ce-delete').addEventListener('click', async () => {
    if (!confirm(`Delete ${editingCode.name}?`)) return;
    codes = codes.filter(c => c !== editingCode);
    await saveCodes();
    codeEditor.close();
  });
  $('#code-form').addEventListener('submit', async e => {
    e.preventDefault();
    const name = $('#ce-name').value.trim(), code = $('#ce-code').value.trim();
    if (!name || !code) return;
    if (editingCode) Object.assign(editingCode, { name, code });
    else codes.push({ name, code });
    await saveCodes();
    codeEditor.close();
  });
  $('#codes-file').addEventListener('change', async e => {
    const f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    try {
      const imported = await parseCostCodes(await readFile(f));
      if (!imported.length) { alert('No cost codes found in that file.'); return; }
      let added = 0;
      for (const c of imported) {
        if (!codes.some(x => x.code === c.code && x.name === c.name)) { codes.push(c); added++; }
      }
      await saveCodes();
      toast(`Imported ${added} cost code${added === 1 ? '' : 's'}`);
    } catch (err) {
      alert("Couldn't read that file: " + err.message);
    }
  });
  for (const [id, key] of [['#opt-caption', 'caption'], ['#opt-placeholder', 'placeholder']]) {
    $(id).addEventListener('change', e => { opts[key] = e.target.checked; DB.setKv('opts', opts); });
  }

  // ---------- backup / restore ----------
  $('#btn-backup').addEventListener('click', async () => {
    busy('Building backup…');
    try {
      const zip = new JSZip();
      const meta = receipts.map(r => ({
        ...r,
        files: r.files.map((p, i) => {
          const name = `files/${r.id}_${i}.${p.type === 'application/pdf' ? 'pdf' : 'jpg'}`;
          zip.file(name, p.data);
          return { type: p.type, name };
        }),
      }));
      zip.file('data.json', JSON.stringify({ version: 1, receipts: meta, codes, opts, lastCode }, null, 1));
      const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/zip' });
      busy();
      await shareOrDownload([new File([blob], `receipt-backup-${todayISO()}.zip`, { type: 'application/zip' })], 'Receipt backup');
    } catch (err) {
      busy();
      if (err.name !== 'AbortError') alert('Backup failed: ' + err.message);
    }
  });
  $('#restore-file').addEventListener('change', async e => {
    const f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    busy('Restoring…');
    try {
      const zip = await JSZip.loadAsync(await readFile(f));
      const data = JSON.parse(await zip.file('data.json').async('string'));
      for (const r of data.receipts) {
        r.files = await Promise.all(r.files.map(async p => ({ type: p.type, data: await zip.file(p.name).async('arraybuffer') })));
        const here = receipts.find(x => x.id === r.id);
        if (here && here.submittedAt && !r.submittedAt) r.submittedAt = here.submittedAt;
        await DB.put('receipts', r);
      }
      for (const c of data.codes || []) if (!codes.some(x => x.code === c.code && x.name === c.name)) codes.push(c);
      await saveCodes();
      receipts = await DB.all('receipts');
      renderReceipts();
      toast(`Restored ${data.receipts.length} receipts`);
    } catch (err) {
      alert('Restore failed: ' + err.message);
    } finally { busy(); }
  });

  async function shareOrDownload(files, title) {
    if (navigator.canShare && navigator.canShare({ files })) {
      await navigator.share({ files, title });
      return;
    }
    for (const f of files) {
      const a = el('a', { href: URL.createObjectURL(f), download: f.name });
      document.body.append(a);
      a.click();
      a.remove();
    }
  }

  // ---------- month-end ----------
  let ap = null; // { buffer, fileName, parsed, assign: Map(row -> receiptId|null), status: Map(row -> 'ok'|'check'|'missing'|'manual'), output }

  $('#ap-file').addEventListener('change', async e => {
    const f = e.target.files[0];
    e.target.value = '';
    if (!f) return;
    try {
      const buffer = await readFile(f);
      const parsed = await parseApSheet(buffer.slice(0));
      ap = { buffer, fileName: f.name, parsed, assign: new Map(), status: new Map(), output: null };
      autoMatch();
      renderMatch();
    } catch (err) {
      alert("Couldn't read that spreadsheet: " + err.message);
    }
  });

  function autoMatch() {
    const pending = receipts.filter(r => !r.submittedAt);
    const used = new Set();
    const order = ap.parsed.lines
      .map(line => ({ line, cands: pending.filter(r => r.cents === line.cents) }))
      .sort((a, b) => a.cands.length - b.cands.length);
    for (const { line, cands } of order) {
      const avail = cands.filter(r => !used.has(r.id))
        .sort((a, b) => dayDiff(a.date, line.date) - dayDiff(b.date, line.date));
      const pick = avail[0];
      if (!pick) {
        ap.assign.set(line.row, null);
        ap.status.set(line.row, 'missing');
        continue;
      }
      used.add(pick.id);
      ap.assign.set(line.row, pick.id);
      const d0 = dayDiff(pick.date, line.date);
      const ambiguous = avail.length > 1 && dayDiff(avail[1].date, line.date) === d0;
      ap.status.set(line.row, d0 <= 7 && !ambiguous ? 'ok' : 'check');
    }
  }

  const STATUS_TEXT = { ok: 'Matched', check: 'Check', missing: 'No receipt', manual: 'Picked' };

  function renderMatch() {
    const area = $('#match-area');
    area.replaceChildren();
    if (!ap) return;
    ap.output = null;
    const { lines } = ap.parsed;
    const byId = id => receipts.find(r => r.id === id);
    // drop assignments to receipts that were deleted
    for (const [row, id] of ap.assign) if (id && !byId(id)) { ap.assign.set(row, null); ap.status.set(row, 'missing'); }

    const counts = { ok: 0, check: 0, missing: 0 };
    for (const line of lines) {
      const s = ap.status.get(line.row);
      counts[s === 'manual' ? 'ok' : s]++;
    }
    const total = lines.reduce((s, l) => s + l.cents, 0);
    area.append(el('div', { class: 'card' },
      el('div', { class: 'match-head' },
        el('h2', {}, ap.parsed.employee || ap.fileName),
        el('span', { class: 'muted' }, `${lines.length} lines · ${fmtMoney(total)}`)),
      el('div', {},
        el('span', { class: 'pill ok' }, `${counts.ok} matched`), ' ',
        counts.check ? el('span', { class: 'pill check' }, `${counts.check} to check`) : null, ' ',
        counts.missing ? el('span', { class: 'pill missing' }, `${counts.missing} missing`) : null)));

    const pending = receipts.filter(r => !r.submittedAt);
    lines.forEach((line, i) => {
      const status = ap.status.get(line.row);
      const rec = byId(ap.assign.get(line.row));
      const sel = el('select', {
        onchange: e => {
          const v = e.target.value;
          if (v === '__add') {
            e.target.value = rec ? rec.id : '';
            addForLine([]);
            return;
          }
          ap.assign.set(line.row, v || null);
          ap.status.set(line.row, v ? 'manual' : 'missing');
          renderMatch();
        },
      });
      sel.append(el('option', { value: '' }, '— No receipt —'));
      const same = pending.filter(r => r.cents === line.cents);
      const other = pending.filter(r => r.cents !== line.cents)
        .sort((a, b) => Math.abs(a.cents - line.cents) - Math.abs(b.cents - line.cents));
      const opt = r => el('option', { value: r.id }, `${fmtMoney(r.cents)} · ${fmtDate(r.date)} · ${r.vendor || r.note || 'receipt'}`);
      if (same.length) sel.append(el('optgroup', { label: 'Same amount' }, same.map(opt)));
      if (other.length) sel.append(el('optgroup', { label: 'Other receipts' }, other.map(opt)));
      if (rec && rec.submittedAt) sel.append(opt(rec));
      sel.append(el('option', { value: '__add' }, '📷 Add a receipt for this line…'));
      sel.value = rec ? rec.id : '';

      let detail = null;
      if (rec) {
        const warn = rec.cents !== line.cents ? `Amount differs (${fmtMoney(rec.cents)}). ` : '';
        detail = el('div', { class: 'detail' + (warn ? ' warn' : '') },
          warn, `${rec.category} · ${rec.code ? codeLabel(rec.code) : 'NO COST CODE'} · ${rec.note || 'no note'} `,
          el('button', { class: 'link', onclick: () => openEditor(rec, {}, () => renderMatch()) }, 'Edit'));
      }
      const addForLine = pages => openEditor(null,
        { cents: line.cents, date: line.date, vendor: shortPlace(line.place), pages }, saved => {
          ap.assign.set(line.row, saved.id);
          ap.status.set(line.row, 'manual');
          renderMatch();
        });
      const row = el('div', { class: 'mrow' + (rec ? '' : ' needs') },
        el('div', { class: 'top' },
          el('div', {},
            el('div', {}, el('strong', {}, `${i + 1}. ${fmtDate(line.date)}`), ' ',
              el('span', { class: 'pill ' + (status === 'manual' ? 'ok' : status) }, STATUS_TEXT[status])),
            el('div', { class: 'place' }, line.place)),
          el('div', { class: 'amt' }, fmtMoney(line.cents))),
        sel, detail,
        rec ? null : el('div', { class: 'drop-hint' }, 'Drop an invoice PDF or image here'));
      enableDrop(row, async files => addForLine(await filesToPages(files)));
      area.append(row);
    });

    const assigned = new Set(ap.assign.values());
    const unused = pending.filter(r => !assigned.has(r.id));
    if (unused.length) {
      area.append(el('h2', { style: { marginTop: '20px' } }, 'Receipts not on this statement'),
        el('p', { class: 'muted small' }, 'These stay saved for next month.'));
      unused.forEach(r => { const c = receiptCard(r); c.classList.add('unused'); area.append(c); });
    }

    const out = el('div', { class: 'card', id: 'output-card' },
      el('h2', {}, 'Create files'),
      el('p', { class: 'muted small' }, 'Fills in Guest/Purpose, Cost Code and the category amount on each matched line, and builds one PDF of the receipts in spreadsheet order.'),
      el('button', { class: 'btn primary block', onclick: buildOutput }, 'Create spreadsheet + PDF'),
      el('div', { id: 'output-files', class: 'result-files' }));
    area.append(out);
  }

  function shortPlace(place) {
    return place.replace(/\s{2,}.*$/, '').replace(/\*.*$/, '').trim();
  }
  const ascii = s => String(s || '').replace(/[^\x20-\x7E]/g, '').trim();

  async function buildPdf() {
    const { PDFDocument, StandardFonts, rgb } = PDFLib;
    const pdf = await PDFDocument.create();
    const font = await pdf.embedFont(StandardFonts.Helvetica);
    const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
    const W = 612, H = 792, M = 36;
    const caption = (page, text) => {
      const size = 9;
      let t = ascii(text);
      while (t.length > 10 && font.widthOfTextAtSize(t, size) > W - 2 * M) t = t.slice(0, -2);
      page.drawText(t, { x: M, y: H - M + 4, size, font, color: rgb(0.35, 0.35, 0.35) });
    };
    for (const [i, line] of ap.parsed.lines.entries()) {
      const label = `Line ${i + 1}  |  ${fmtDate(line.date)}  |  ${fmtMoney(line.cents)}  |  ${line.place}`;
      const rec = receipts.find(r => r.id === ap.assign.get(line.row));
      if (!rec || !rec.files.length) {
        if (!opts.placeholder) continue;
        const page = pdf.addPage([W, H]);
        page.drawText('No receipt', { x: M, y: H - 120, size: 28, font: bold });
        page.drawText(ascii(label), { x: M, y: H - 150, size: 11, font, maxWidth: W - 2 * M });
        if (rec && rec.note) page.drawText(ascii('Note: ' + rec.note), { x: M, y: H - 170, size: 11, font, maxWidth: W - 2 * M });
        continue;
      }
      for (const p of rec.files) {
        if (p.type === 'application/pdf') {
          const src = await PDFDocument.load(p.data, { ignoreEncryption: true });
          const copied = await pdf.copyPages(src, src.getPageIndices());
          copied.forEach(pg => pdf.addPage(pg));
        } else {
          const img = await pdf.embedJpg(p.data);
          const page = pdf.addPage([W, H]);
          const top = opts.caption ? H - M - 10 : H - M;
          const boxW = W - 2 * M, boxH = top - M;
          const s = Math.min(boxW / img.width, boxH / img.height);
          const w = img.width * s, h = img.height * s;
          page.drawImage(img, { x: (W - w) / 2, y: top - h, width: w, height: h });
          if (opts.caption) caption(page, label);
        }
      }
    }
    if (!pdf.getPageCount()) pdf.addPage([W, H]);
    return new Blob([await pdf.save()], { type: 'application/pdf' });
  }

  async function buildOutput() {
    const fills = new Map();
    for (const line of ap.parsed.lines) {
      const rec = receipts.find(r => r.id === ap.assign.get(line.row));
      if (rec) fills.set(line.row, { note: rec.note, code: rec.code, category: rec.category });
    }
    const missingCode = [...fills.values()].filter(f => !f.code).length;
    const missing = ap.parsed.lines.length - fills.size;
    const warnings = [];
    if (missing) warnings.push(`${missing} line(s) have no receipt`);
    if (missingCode) warnings.push(`${missingCode} receipt(s) have no cost code`);
    if (warnings.length && !confirm(warnings.join(' and ') + '. Create the files anyway?')) return;

    busy('Building spreadsheet…');
    try {
      const xlsxBlob = await fillApSheet(ap.buffer.slice(0), ap.parsed, fills);
      busy('Building PDF…');
      const pdfBlob = await buildPdf();
      const base = ap.fileName.replace(/\.xlsx$/i, '');
      const xlsxFile = new File([xlsxBlob], `${base}.xlsx`, { type: xlsxBlob.type });
      const pdfFile = new File([pdfBlob], `${base} receipts.pdf`, { type: 'application/pdf' });
      ap.output = { xlsxFile, pdfFile, ids: [...new Set([...ap.assign.values()].filter(Boolean))] };
      showOutput();
    } catch (err) {
      alert('Something went wrong: ' + err.message);
      console.error(err);
    } finally { busy(); }
  }

  function showOutput() {
    const { xlsxFile, pdfFile, ids } = ap.output;
    const mb = (pdfFile.size / 1048576).toFixed(1);
    const wrap = $('#output-files');
    const share = async files => {
      try { await shareOrDownload(files, 'Credit card report'); }
      catch (err) { if (err.name !== 'AbortError') alert(err.message); }
    };
    wrap.replaceChildren(
      el('p', { class: 'small' }, `✅ Ready: ${xlsxFile.name} and ${pdfFile.name} (${mb} MB).`),
      el('button', { class: 'btn primary', onclick: () => share([xlsxFile, pdfFile]) }, 'Share both (Mail, Files…)'),
      el('button', { class: 'btn', onclick: () => share([xlsxFile]) }, 'Share spreadsheet only'),
      el('button', { class: 'btn', onclick: () => share([pdfFile]) }, 'Share PDF only'),
      el('button', { class: 'btn', onclick: () => window.open(URL.createObjectURL(pdfFile), '_blank') }, 'Preview PDF'),
      el('button', {
        class: 'btn', onclick: async () => {
          if (!confirm(`Mark ${ids.length} receipt(s) as submitted? They'll move to the "Submitted" list.`)) return;
          const now = Date.now();
          for (const id of ids) {
            const r = receipts.find(x => x.id === id);
            if (r) { r.submittedAt = now; await DB.put('receipts', r); }
          }
          renderReceipts();
          ap = null;
          renderMatch();
          showTab('receipts');
          toast('Marked as submitted');
        },
      }, 'Done — mark receipts as submitted'));
    wrap.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ---------- startup ----------
  async function init() {
    await DB.open();
    receipts = await DB.all('receipts');
    codes = await DB.getKv('codes', []);
    opts = Object.assign(opts, await DB.getKv('opts', {}));
    lastCode = await DB.getKv('lastCode', '');
    $('#opt-caption').checked = opts.caption;
    $('#opt-placeholder').checked = opts.placeholder;
    renderReceipts();
    renderCodes();
    if (!codes.length) {
      showTab('settings');
      toast('Start by importing your Cost Codes spreadsheet', 4000);
    }
    if (navigator.storage && navigator.storage.persist) {
      const persisted = await navigator.storage.persist().catch(() => false);
      const est = navigator.storage.estimate ? await navigator.storage.estimate().catch(() => null) : null;
      $('#storage-info').textContent =
        (est ? `Using ${(est.usage / 1048576).toFixed(1)} MB. ` : '') +
        (persisted ? 'Storage is protected.' : 'Tip: add this app to your Home Screen so iOS keeps your receipts.');
    }
    if ('serviceWorker' in navigator && location.protocol === 'https:') {
      navigator.serviceWorker.register('sw.js').catch(() => {});
    }
  }
  init().catch(err => alert('Startup failed: ' + err.message));

  // expose for testing
  window.__app = { get ap() { return ap; }, get receipts() { return receipts; }, DB, buildPdf, autoMatch, renderMatch };
})();
