/* app.js – giao diện dashboard: nạp file, bộ lọc, biểu đồ, bảng, chi tiết, nhắc nhở, xuất file. */
(function () {
  'use strict';
  const { XLSXIO, CRMConvert, CRMModel, CRMReport } = window;
  const { U, NONE, STATUSES, STATUS_COLORS, STAGES, BANDS, BAND_LABEL, STALE, STALE_LABEL, CLOSE, CLOSE_LABEL } = CRMModel;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = U.esc;
  const LS = {
    get(k, d) { try { const v = localStorage.getItem('ntsCRM.' + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem('ntsCRM.' + k, JSON.stringify(v)); return true; } catch (e) { return false; } },
    del(k) { try { localStorage.removeItem('ntsCRM.' + k); } catch (e) { /* bỏ qua */ } },
  };

  // ================================================================ ICONS
  const I = {
    x: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>',
    chev: '<svg class="chev" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>',
    flag: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M5 21V4h11l-1.5 4L16 12H7v9z"/></svg>',
    bell: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9"/><path d="M10.3 21a1.94 1.94 0 0 0 3.4 0"/></svg>',
    clock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
    upload: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m17 8-5-5-5 5"/><path d="M12 3v12"/></svg>',
    download: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/><path d="M12 15V3"/></svg>',
    excel: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="m9 12 5 6M14 12l-5 6"/></svg>',
    file: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>',
    cal: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 10h18"/></svg>',
    cols: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9 4v16M15 4v16"/></svg>',
    rows: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M4 6h16M4 12h16M4 18h16"/></svg>',
    copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>',
    filter: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 5h18l-7 8v6l-4 2v-8z"/></svg>',
    mail: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>',
    list: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/></svg>',
    reset: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7L3 8"/><path d="M3 3v5h5"/></svg>',
    gear: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
  };

  // ================================================================ STATE
  const state = {
    loaded: false,
    source: null,          // 'raw' | 'template' | 'saved'
    fileName: '',
    rawObjs: null,         // giữ dữ liệu raw để convert lại khi đổi ngày chốt
    conv: null,            // kết quả convert: { records, nRaw, nTest, refDate, warnings }
    refDate: null,
    refAuto: true,
    deals: [],
    filtered: [],
    filters: newFilters(),
    sort: { key: 'updated', dir: -1 },
    page: 1,
    pageSize: LS.get('pageSize', 50),
    tab: 'deals',
    pivotSort: { key: 'n', dir: -1 },
    compact: LS.get('compact', false),
    staleActiveOnly: LS.get('staleActiveOnly', true),
    modes: { sale: 'count', vendor: 'count', stage: 'count', segment: 'count', reseller: 'count', trend: 'count', trendBy: 'created' },
  };
  function newFilters() {
    return { q: '', facets: {}, quick: new Set(), dateField: 'close', from: '', to: '' };
  }

  // ================================================================ FACETS
  const ordered = (order) => (a, b) => {
    const ia = order.indexOf(a), ib = order.indexOf(b);
    return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
  };
  const FACETS = {
    status: { label: 'Trạng thái', get: (d) => d.status, order: STATUSES, color: (k) => STATUS_COLORS[k] },
    stage: { label: 'Giai đoạn', get: (d) => d.stage, order: STAGES },
    sale: { label: 'Sale', get: (d) => d.sale, search: true },
    band: { label: 'Giá trị', get: (d) => d.band, order: [...BANDS.map((b) => b.k), NONE], fmt: (k) => BAND_LABEL[k] },
    segment: { label: 'Nhóm KH', get: (d) => d.segment, order: CRMConvert.SEGMENTS, search: true },
    ctype: { label: 'Loại HĐ', get: (d) => d.ctype, fmt: (k) => (k === NONE ? '(Chưa phân loại)' : k) },
    vendor: { label: 'Hãng', get: (d) => d.vendors, search: true, fmt: (k) => (k === NONE ? '(Chưa xác định)' : k) },
    reseller: { label: 'Reseller', get: (d) => d.reseller, search: true, fmt: (k) => (k === NONE ? '(Chưa có reseller)' : k) },
    stale: { label: 'Cờ nhắc', get: (d) => d.stale, order: STALE.map((s) => s.k), fmt: (k) => STALE_LABEL[k], color: (k) => (STALE.find((s) => s.k === k) || {}).color },
    closeState: { label: 'Tình trạng timeline', get: (d) => d.closeState, order: CLOSE.map((c) => c.k), fmt: (k) => CLOSE_LABEL[k], color: (k) => (CLOSE.find((c) => c.k === k) || {}).color },
    customer: { label: 'Khách hàng', get: (d) => d.customer, search: true },
    quality: { label: 'Dữ liệu', get: (d) => (d.issues.length ? 'issue' : 'ok'), order: ['issue', 'ok'], fmt: (k) => (k === 'issue' ? 'Có vấn đề cần kiểm tra' : 'Không có vấn đề') },
  };
  const fmtFacet = (key, k) => (FACETS[key].fmt ? FACETS[key].fmt(k) : k === NONE ? '(Trống)' : k);
  const PILLS = ['status', 'stage', 'sale', 'band', 'DATE', 'segment', 'ctype', 'vendor', 'reseller', 'stale', 'customer', 'closeState', 'quality'];

  const DATE_FIELDS = { close: 'Close date (Timeline dự án)', created: 'Ngày tạo', updated: 'Cập nhật lần cuối' };

  const QUICK = {
    overdue: { label: 'Quá timeline – vẫn Active', color: '#e34948', test: (d) => d.overdue },
    soon30: { label: 'Close ≤ 30 ngày tới', color: '#eb6834', test: (d) => d.status === 'Active' && d.closeState === 'd30' },
    remind: { label: 'Có cờ nhắc', color: '#eda100', test: (d) => d.needRemind },
    stale2: { label: 'Cờ ≥ Mức 2', color: '#c2410c', test: (d) => d.needRemind && +d.stale >= 2 },
    big: { label: 'Deal ≥ 5 tỷ', color: '#0959a2', test: (d) => d.value >= 5e9 },
    novalue: { label: 'Chưa có giá trị', color: '#9aa8b8', test: (d) => !d.value },
    multi: { label: 'Nhiều hãng', color: '#19c3c9', test: (d) => d.vendors.length > 1 },
    issue: { label: 'Có vấn đề dữ liệu', color: '#f06723', test: (d) => d.issues.length > 0 },
  };

  function facetValues(d, key) {
    let v = FACETS[key].get(d);
    if (Array.isArray(v)) return v.length ? v : [NONE];
    return [v == null || v === '' ? NONE : v];
  }
  function passes(d, f, skip) {
    if (f.q) {
      const toks = U.fold(f.q).split(/\s+/).filter(Boolean);
      for (const t of toks) if (!d._s.includes(t)) return false;
    }
    for (const key in f.facets) {
      if (key === skip) continue;
      const set = f.facets[key];
      if (!set || !set.size) continue;
      if (!facetValues(d, key).some((v) => set.has(v))) return false;
    }
    if (skip !== 'DATE' && (f.from || f.to)) {
      const v = d[f.dateField];
      if (!v) return false;
      const day = U.iso(v);
      if (f.from && day < f.from) return false;
      if (f.to && day > f.to) return false;
    }
    if (skip !== 'QUICK') for (const q of f.quick) if (!QUICK[q].test(d)) return false;
    return true;
  }
  function applyFilters() {
    state.filtered = state.deals.filter((d) => passes(d, state.filters));
  }
  function facetCounts(key) {
    const m = new Map();
    for (const d of state.deals) {
      if (!passes(d, state.filters, key)) continue;
      for (const v of facetValues(d, key)) m.set(v, (m.get(v) || 0) + 1);
    }
    return m;
  }
  function allFacetKeys(key) {
    const s = new Set();
    for (const d of state.deals) facetValues(d, key).forEach((v) => s.add(v));
    let arr = [...s];
    const F = FACETS[key];
    if (F.order) {
      arr.sort(ordered(F.order));
    } else {
      const cnt = new Map();
      state.deals.forEach((d) => facetValues(d, key).forEach((v) => cnt.set(v, (cnt.get(v) || 0) + 1)));
      arr.sort((a, b) => (a === NONE) - (b === NONE) || cnt.get(b) - cnt.get(a) || String(a).localeCompare(String(b), 'vi'));
    }
    return arr;
  }
  function toggleFacet(key, val, only) {
    const f = state.filters.facets;
    if (!f[key]) f[key] = new Set();
    if (only) {
      const was = f[key].size === 1 && f[key].has(val);
      f[key] = new Set(was ? [] : [val]);
    } else if (f[key].has(val)) f[key].delete(val);
    else f[key].add(val);
    if (!f[key].size) delete f[key];
    refresh();
  }
  function hasAnyFilter() {
    const f = state.filters;
    return !!(f.q || Object.keys(f.facets).length || f.quick.size || f.from || f.to);
  }

  // ================================================================ TOAST / DOWNLOAD / POPOVER
  function toast(msg, type) {
    const t = document.createElement('div');
    t.className = 'toast ' + (type || '');
    t.innerHTML = msg;
    $('#toasts').appendChild(t);
    setTimeout(() => { t.style.opacity = '0'; t.style.transition = 'opacity .3s'; }, type === 'err' ? 6000 : 2800);
    setTimeout(() => t.remove(), type === 'err' ? 6400 : 3200);
  }
  function download(bytes, name) {
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    if (window.navigator && window.navigator.msSaveOrOpenBlob) { window.navigator.msSaveOrOpenBlob(blob, name); return; }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); return true; } catch (e) {
      const ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      let ok = false;
      try { ok = document.execCommand('copy'); } catch (e2) { ok = false; }
      ta.remove();
      return ok;
    }
  }

  let popEl = null, popAnchor = null;
  function closePop() {
    if (popEl) { popEl.remove(); popEl = null; popAnchor = null; }
  }
  function openPop(anchor, html, width) {
    if (popAnchor === anchor) { closePop(); return null; }
    closePop();
    popEl = document.createElement('div');
    popEl.className = 'popover';
    if (width) popEl.style.width = width + 'px';
    popEl.innerHTML = html;
    document.body.appendChild(popEl);
    popAnchor = anchor;
    placePop();
    return popEl;
  }
  function placePop() {
    if (!popEl || !popAnchor) return;
    const r = popAnchor.getBoundingClientRect();
    const w = popEl.offsetWidth, h = popEl.offsetHeight;
    let left = Math.min(r.left, window.innerWidth - w - 12);
    left = Math.max(12, left);
    let top = r.bottom + 6;
    if (top + h > window.innerHeight - 8 && r.top - h - 6 > 8) top = r.top - h - 6;
    popEl.style.left = left + 'px';
    popEl.style.top = Math.max(8, top) + 'px';
  }
  window.addEventListener('resize', placePop);
  window.addEventListener('scroll', () => { if (popEl) placePop(); }, true);
  document.addEventListener('mousedown', (e) => {
    if (popEl && !popEl.contains(e.target) && !(popAnchor && popAnchor.contains(e.target))) closePop();
    if (!$('#searchBox').contains(e.target)) hideSuggest();
  });

  // ================================================================ LOAD FILE
  async function handleFile(file) {
    if (!file) return;
    if (!/\.xls[xm]$/i.test(file.name)) { toast('Vui lòng chọn file Excel định dạng .xlsx', 'err'); return; }
    $('#dzProgress').classList.remove('hidden');
    try {
      const buf = await file.arrayBuffer();
      const wb = await XLSXIO.read(buf);
      if (!wb.sheets.length) throw new Error('File không có sheet dữ liệu.');
      let sheet = wb.sheets.find((s) => s.name === 'Data') || wb.sheets[0];
      let objs = CRMConvert.rowsToObjects(sheet.rows);
      if (!CRMConvert.isTemplate(objs) && !CRMConvert.isRaw(objs)) {
        const alt = wb.sheets.find((s) => { const o = CRMConvert.rowsToObjects(s.rows); return CRMConvert.isTemplate(o) || CRMConvert.isRaw(o); });
        if (alt) { sheet = alt; objs = CRMConvert.rowsToObjects(alt.rows); }
      }
      const inputRef = U.fromIso($('#refDateInput').value);
      if (CRMConvert.isTemplate(objs)) {
        const records = CRMConvert.normalizeTemplate(objs);
        const meta = readTemplateMeta(wb);
        state.source = 'template';
        state.rawObjs = null;
        state.refAuto = !inputRef;
        state.refDate = inputRef || meta.ref || CRMConvert.refFromName(file.name) || U.today();
        state.conv = { records, nRaw: meta.nRaw, nTest: meta.nTest, refDate: state.refDate, warnings: [] };
        state.fileName = file.name;
        setData();
        toast(`Đã nạp file template chuẩn: <b>${U.int(state.deals.length)}</b> thương vụ`, 'ok');
      } else if (CRMConvert.isRaw(objs)) {
        state.source = 'raw';
        state.rawObjs = objs;
        state.fileName = file.name;
        state.refAuto = !inputRef;
        state.refDate = inputRef || CRMConvert.refFromName(file.name) || U.today();
        runConvert();
        const w = state.conv.warnings;
        toast(`Đã convert <b>${U.int(state.conv.nRaw)}</b> dòng raw → <b>${U.int(state.deals.length)}</b> thương vụ (${U.int(state.conv.records.length)} dòng template), loại ${state.conv.nTest} deal test` + (w.length ? ` · ${w.length} cảnh báo` : ''), 'ok');
        w.forEach((x) => toast('Cảnh báo: ' + esc(x), 'err'));
      } else {
        throw new Error('Không nhận diện được file. Cần file export Deals từ CRM (có cột "Tên thương vụ", "Chủ sở hữu thương vụ"…) hoặc file template chuẩn (sheet "Data").');
      }
      saveSession();
    } catch (err) {
      console.error(err);
      toast(esc(err.message || String(err)), 'err');
    } finally {
      $('#dzProgress').classList.add('hidden');
      $('#fileInput').value = '';
    }
  }
  function readTemplateMeta(wb) {
    const meta = { ref: null, nRaw: null, nTest: null };
    const g = wb.sheets.find((s) => s.name === 'Huong_dan');
    if (g) {
      for (const r of g.rows) {
        if (!r) continue;
        if (r[0] === 'Ngày chốt số liệu') meta.ref = CRMConvert.parseDate(r[1]);
        if (r[0] === 'Số deal trong file raw') meta.nRaw = r[1];
        if (r[0] === 'Deal test đã loại') meta.nTest = r[1];
      }
    }
    return meta;
  }
  function runConvert() {
    state.conv = CRMConvert.convert(state.rawObjs, state.refDate);
    setData();
  }
  function setData() {
    state.deals = CRMModel.buildDeals(state.conv.records);
    CRMModel.derive(state.deals, state.refDate, { staleActiveOnly: state.staleActiveOnly });
    state.filters = newFilters();
    state.page = 1;
    state.loaded = true;
    $('#landing').classList.add('hidden');
    $('#dashboard').classList.remove('hidden');
    $('#q').value = '';
    renderHeader();
    refresh();
    window.scrollTo({ top: 0 });
  }
  function setRefDate(d) {
    if (!d) return;
    state.refDate = d;
    state.refAuto = false;
    if (state.source === 'raw' && state.rawObjs) state.conv = CRMConvert.convert(state.rawObjs, d);
    else if (state.conv) state.conv.refDate = d;
    const keep = state.filters;
    state.deals = CRMModel.buildDeals(state.conv.records);
    CRMModel.derive(state.deals, state.refDate, { staleActiveOnly: state.staleActiveOnly });
    state.filters = keep;
    renderHeader();
    refresh();
    saveSession();
    toast('Đã tính lại cờ nhắc & timeline theo ngày chốt ' + U.date(d));
  }

  // ---- lưu phiên làm việc (chỉ trên máy người dùng)
  function encodeRecords(records) {
    return records.map((r) => CRMConvert.TEMPLATE_COLS.map((c) => (r[c] instanceof Date ? 'D:' + r[c].toISOString() : r[c])));
  }
  function decodeRecords(rows) {
    return rows.map((a) => Object.fromEntries(CRMConvert.TEMPLATE_COLS.map((c, i) => [c, typeof a[i] === 'string' && a[i].startsWith('D:') ? new Date(a[i].slice(2)) : a[i]])));
  }
  function saveSession() {
    const ok = LS.set('session', {
      v: 1, fileName: state.fileName, source: state.source, ref: state.refDate.toISOString(), savedAt: new Date().toISOString(),
      nRaw: state.conv.nRaw, nTest: state.conv.nTest, records: encodeRecords(state.conv.records),
    });
    if (!ok) LS.del('session');
  }
  function restoreSession() {
    const s = LS.get('session', null);
    if (!s || !s.records) return;
    state.source = 'saved';
    state.fileName = s.fileName;
    state.rawObjs = null;
    state.refDate = new Date(s.ref);
    state.conv = { records: decodeRecords(s.records), nRaw: s.nRaw, nTest: s.nTest, refDate: state.refDate, warnings: [] };
    setData();
    toast('Đã mở lại dữ liệu: ' + esc(s.fileName));
  }

  // ================================================================ HEADER
  function renderHeader() {
    const h = $('#headerMeta');
    if (!state.loaded) { h.innerHTML = ''; return; }
    const src = state.source === 'raw' ? 'Raw CRM → đã convert' : state.source === 'template' ? 'File template chuẩn' : 'Dữ liệu đã lưu';
    h.innerHTML = `
      <span class="meta-chip" title="${esc(state.fileName)}">${I.file.replace('<svg', '<svg width="15" height="15"')}<span class="ellipsis">${esc(state.fileName)}</span></span>
      <span class="meta-chip" title="${esc(src)}"><b>${U.int(state.deals.length)}</b> thương vụ</span>
      <label class="meta-chip" title="Ngày chốt số liệu – dùng tính cờ nhắc và timeline">${I.cal.replace('<svg', '<svg width="15" height="15"')} Chốt: <input type="date" id="hdrRef" value="${U.iso(state.refDate)}"></label>
      <button class="btn btn-glass btn-sm" data-act="reload" title="Nạp file khác">${I.upload}Nạp file</button>
      <button class="btn btn-glass btn-sm" data-act="dlTemplate" title="Tải file dữ liệu sau convert (template chuẩn)">${I.download}File chuẩn</button>
      <button class="btn btn-orange btn-sm" data-act="exportReport">${I.excel}Xuất Excel tổng hợp</button>`;
    $('#hdrRef').addEventListener('change', (e) => { const d = U.fromIso(e.target.value); if (d) setRefDate(d); });
  }

  // ================================================================ REFRESH
  function refresh() {
    applyFilters();
    const maxPage = Math.max(1, Math.ceil(state.filtered.length / (state.pageSize || 1e9)));
    if (state.page > maxPage) state.page = maxPage;
    renderPills();
    renderQuick();
    renderActive();
    renderAlert();
    renderKpis();
    renderInsights();
    renderCharts();
    renderTabs();
    renderTable();
  }

  // ---------------------------------------------------------------- pills
  function renderPills() {
    const f = state.filters;
    $('#pills').innerHTML = PILLS.map((key) => {
      if (key === 'DATE') {
        const on = f.from || f.to;
        const txt = on ? `${f.from ? U.date(U.fromIso(f.from)) : '…'} → ${f.to ? U.date(U.fromIso(f.to)) : '…'}` : 'Từ ngày – đến ngày';
        return `<button class="fpill ${on ? 'on' : ''}" data-pill="DATE">${I.cal}${esc(txt)}${I.chev}</button>`;
      }
      const n = f.facets[key] ? f.facets[key].size : 0;
      return `<button class="fpill ${n ? 'on' : ''}" data-pill="${key}">${esc(FACETS[key].label)}${n ? `<span class="badge">${n}</span>` : ''}${I.chev}</button>`;
    }).join('') + `<span class="fb-sep"></span><button class="fpill" data-pill="SETTINGS" title="Thiết lập cờ nhắc">${I.gear}</button>`;
    const total = state.deals.length, n = state.filtered.length;
    $('#fbCount').innerHTML = `Hiển thị <b>${U.int(n)}</b> / ${U.int(total)} thương vụ`;
  }

  function openFacetPop(anchor, key) {
    const F = FACETS[key];
    const counts = facetCounts(key);
    const sel = state.filters.facets[key] || new Set();
    const keys = allFacetKeys(key);
    const pop = openPop(anchor, `
      <div class="pop-head">
        <div class="pop-title">${esc(F.label)} <small>${keys.length} lựa chọn</small></div>
        ${F.search || keys.length > 10 ? `<input class="pop-search" placeholder="Tìm ${esc(F.label.toLowerCase())}…">` : ''}
      </div>
      <div class="pop-list"></div>
      <div class="pop-foot">
        <button class="btn btn-ghost btn-sm" data-pa="clear">Bỏ chọn</button>
        <div style="display:flex;gap:6px"><button class="btn btn-sm" data-pa="all">Chọn tất cả đang hiện</button></div>
      </div>`, key === 'customer' ? 380 : 320);
    if (!pop) return;
    const list = $('.pop-list', pop);
    const draw = (q) => {
      const fq = U.fold(q || '');
      const shown = keys.filter((k) => !fq || U.fold(fmtFacet(key, k)).includes(fq));
      list.innerHTML = shown.slice(0, 400).map((k) => {
        const c = counts.get(k) || 0;
        const color = F.color ? F.color(k) : null;
        return `<label class="opt ${c ? '' : 'zero'}"><input type="checkbox" value="${esc(k)}" ${sel.has(k) ? 'checked' : ''}>` +
          (color ? `<span class="o-dot" style="background:${color}"></span>` : '') +
          `<span class="o-label" title="${esc(fmtFacet(key, k))}">${esc(fmtFacet(key, k))}</span><span class="o-count">${U.int(c)}</span></label>`;
      }).join('') + (shown.length > 400 ? `<div class="muted" style="padding:6px 8px;font-size:12px">… còn ${shown.length - 400} mục, hãy gõ để tìm</div>` : '') +
        (!shown.length ? '<div class="muted" style="padding:10px">Không có kết quả</div>' : '');
      list._shown = shown;
    };
    draw('');
    const s = $('.pop-search', pop);
    if (s) { s.addEventListener('input', () => draw(s.value)); setTimeout(() => s.focus(), 30); }
    list.addEventListener('change', (e) => {
      if (e.target.type !== 'checkbox') return;
      const f = state.filters.facets;
      if (!f[key]) f[key] = new Set();
      e.target.checked ? f[key].add(e.target.value) : f[key].delete(e.target.value);
      if (!f[key].size) delete f[key];
      state.page = 1;
      refresh();
    });
    pop.addEventListener('click', (e) => {
      const a = e.target.closest('[data-pa]');
      if (!a) return;
      const f = state.filters.facets;
      if (a.dataset.pa === 'clear') delete f[key];
      else f[key] = new Set([...(f[key] || []), ...(list._shown || keys)]);
      if (f[key] && !f[key].size) delete f[key];
      state.page = 1;
      refresh();
      $$('input[type=checkbox]', list).forEach((cb) => { cb.checked = !!(f[key] && f[key].has(cb.value)); });
    });
  }

  function openDatePop(anchor) {
    const f = state.filters;
    const pop = openPop(anchor, `
      <div class="pop-head"><div class="pop-title">Lọc theo khoảng thời gian</div>
        <div class="seg" id="dfSeg">${Object.entries(DATE_FIELDS).map(([k, v]) => `<button data-df="${k}" class="${f.dateField === k ? 'on' : ''}">${esc(v.split(' (')[0])}</button>`).join('')}</div>
      </div>
      <div class="pop-body">
        <div class="date-grid">
          <label>Từ ngày<input type="date" id="dFrom" value="${f.from}"></label>
          <label>Đến ngày<input type="date" id="dTo" value="${f.to}"></label>
        </div>
        <div class="presets">
          <button data-pr="m0">Tháng này</button><button data-pr="q0">Quý này</button><button data-pr="y0">Năm nay</button>
          <button data-pr="n30">30 ngày tới</button><button data-pr="n90">90 ngày tới</button><button data-pr="p30">30 ngày qua</button>
          <button data-pr="p90">90 ngày qua</button><button data-pr="past">Trước ngày chốt</button><button data-pr="m1">Tháng sau</button>
        </div>
        <div class="muted" style="font-size:12px;margin-top:10px">Mốc tính theo ngày chốt số liệu ${U.date(state.refDate)}. Thương vụ không có ngày ở trường đã chọn sẽ bị ẩn khi đang lọc.</div>
      </div>
      <div class="pop-foot"><button class="btn btn-ghost btn-sm" data-pa="clear">Xoá lọc thời gian</button><button class="btn btn-primary btn-sm" data-pa="ok">Áp dụng</button></div>`, 360);
    if (!pop) return;
    const apply = () => {
      f.from = $('#dFrom', pop).value;
      f.to = $('#dTo', pop).value;
      if (f.from && f.to && f.from > f.to) [f.from, f.to] = [f.to, f.from];
      state.page = 1;
      refresh();
    };
    $('#dFrom', pop).addEventListener('change', apply);
    $('#dTo', pop).addEventListener('change', apply);
    pop.addEventListener('click', (e) => {
      const df = e.target.closest('[data-df]');
      if (df) { f.dateField = df.dataset.df; $$('#dfSeg button', pop).forEach((b) => b.classList.toggle('on', b === df)); if (f.from || f.to) refresh(); return; }
      const pr = e.target.closest('[data-pr]');
      if (pr) {
        const r = state.refDate, y = r.getUTCFullYear(), m = r.getUTCMonth();
        const D = (yy, mm, dd) => U.iso(new Date(Date.UTC(yy, mm, dd)));
        const plus = (n) => U.iso(new Date(r.getTime() + n * U.DAY));
        const q0 = Math.floor(m / 3) * 3;
        const map = {
          m0: [D(y, m, 1), D(y, m + 1, 0)], q0: [D(y, q0, 1), D(y, q0 + 3, 0)], y0: [D(y, 0, 1), D(y, 11, 31)],
          n30: [plus(0), plus(30)], n90: [plus(0), plus(90)], p30: [plus(-30), plus(0)], p90: [plus(-90), plus(0)],
          past: ['', plus(-1)], m1: [D(y, m + 1, 1), D(y, m + 2, 0)],
        };
        [$('#dFrom', pop).value, $('#dTo', pop).value] = map[pr.dataset.pr];
        apply();
        return;
      }
      const a = e.target.closest('[data-pa]');
      if (!a) return;
      if (a.dataset.pa === 'clear') { f.from = ''; f.to = ''; refresh(); closePop(); }
      else { apply(); closePop(); }
    });
  }

  function openSettingsPop(anchor) {
    const pop = openPop(anchor, `
      <div class="pop-head"><div class="pop-title">Thiết lập cờ nhắc cập nhật</div></div>
      <div class="pop-body" style="font-size:13px;line-height:1.55">
        <label class="opt" style="padding:4px 0"><input type="checkbox" id="optActiveOnly" ${state.staleActiveOnly ? 'checked' : ''}><span>Chỉ gắn cờ nhắc cho thương vụ <b>Active</b> (bỏ qua deal đã Won/Lost)</span></label>
        <div class="muted" style="margin-top:8px">Mức cờ tính từ <b>Cập nhật lần cuối</b> đến ngày chốt ${U.date(state.refDate)}:<br>
        Mức 1: &gt; 1 tháng · Mức 2: &gt; 3 tháng · Mức 3: &gt; 6 tháng · Mức 4: &gt; 1 năm.</div>
      </div>`, 340);
    if (!pop) return;
    $('#optActiveOnly', pop).addEventListener('change', (e) => {
      state.staleActiveOnly = e.target.checked;
      LS.set('staleActiveOnly', state.staleActiveOnly);
      CRMModel.derive(state.deals, state.refDate, { staleActiveOnly: state.staleActiveOnly });
      refresh();
    });
  }

  // ---------------------------------------------------------------- quick chips
  function renderQuick() {
    const base = state.deals.filter((d) => passes(d, state.filters, 'QUICK'));
    const st = state.filters.facets.status || new Set();
    const sc = {};
    state.deals.filter((d) => passes(d, state.filters, 'status')).forEach((d) => { sc[d.status] = (sc[d.status] || 0) + 1; });
    let html = '<span class="qlabel">Lọc nhanh</span>';
    html += STATUSES.map((s) => `<button class="qchip ${st.has(s) ? 'on' : ''}" data-qs="${s}"><i style="background:${STATUS_COLORS[s]}"></i>${s} <b>${U.int(sc[s] || 0)}</b></button>`).join('');
    html += Object.entries(QUICK).map(([k, q]) => {
      const n = base.filter(q.test).length;
      return `<button class="qchip ${state.filters.quick.has(k) ? 'on' : ''}" data-qk="${k}"><i style="background:${q.color}"></i>${esc(q.label)} <b>${U.int(n)}</b></button>`;
    }).join('');
    $('#qchips').innerHTML = html;
  }

  // ---------------------------------------------------------------- active filter chips
  function renderActive() {
    const f = state.filters;
    const chips = [];
    if (f.q) chips.push(`<span class="af-chip"><span><em>Tìm:</em> "${esc(f.q)}"</span><button data-rm="q">${I.x}</button></span>`);
    for (const key of Object.keys(f.facets)) {
      for (const v of f.facets[key]) {
        chips.push(`<span class="af-chip" title="${esc(fmtFacet(key, v))}"><span><em>${esc(FACETS[key].label)}:</em> ${esc(fmtFacet(key, v))}</span><button data-rm="f" data-k="${esc(key)}" data-v="${esc(v)}">${I.x}</button></span>`);
      }
    }
    if (f.from || f.to) chips.push(`<span class="af-chip"><span><em>${esc(DATE_FIELDS[f.dateField].split(' (')[0])}:</em> ${f.from ? U.date(U.fromIso(f.from)) : '…'} → ${f.to ? U.date(U.fromIso(f.to)) : '…'}</span><button data-rm="date">${I.x}</button></span>`);
    for (const k of f.quick) chips.push(`<span class="af-chip"><span><em>Nhanh:</em> ${esc(QUICK[k].label)}</span><button data-rm="quick" data-k="${k}">${I.x}</button></span>`);
    $('#afRow').classList.toggle('hidden', !chips.length);
    $('#activeFilters').innerHTML = chips.join('') + (chips.length ? `<button class="btn btn-ghost btn-sm" data-act="clearAll">${I.reset}Xoá tất cả bộ lọc</button>` : '');
  }

  // ---------------------------------------------------------------- alert
  function renderAlert() {
    const list = state.filtered.filter((d) => d.needRemind);
    const el = $('#staleAlert');
    if (!list.length) { el.classList.add('hidden'); return; }
    el.classList.remove('hidden');
    const oldest = list.reduce((a, b) => (b.staleDays > a.staleDays ? b : a));
    const lv = {};
    list.forEach((d) => { lv[d.stale] = (lv[d.stale] || 0) + 1; });
    el.innerHTML = `
      <div class="alert-ico">${I.bell}</div>
      <div class="alert-body">
        <div class="alert-title"><b>${U.int(list.length)}</b> thương vụ${state.staleActiveOnly ? ' Active' : ''} đã <b>hơn 1 tháng</b> chưa được cập nhật (tính đến ${U.date(state.refDate)}).
          Lâu nhất: <a href="#" data-open="${esc(oldest.id)}">#${esc(oldest.id)} ${esc(oldest.name)}</a> — <b>${U.int(oldest.staleDays)}</b> ngày${oldest.sale ? ` (${esc(oldest.sale)})` : ''}.</div>
        <div class="alert-levels">${STALE.filter((s) => lv[s.k]).map((s) => `<button class="qchip" data-fk="stale" data-fv="${s.k}"><i style="background:${s.color}"></i>${esc(s.label)} <b>${lv[s.k]}</b></button>`).join('')}</div>
      </div>
      <div class="alert-actions">
        <button class="btn btn-primary" data-act="showStale">${I.list}Xem danh sách</button>
        <button class="btn btn-orange" data-act="remind">${I.mail}Soạn nội dung nhắc</button>
      </div>`;
  }

  // ---------------------------------------------------------------- KPI
  function renderKpis() {
    const s = CRMModel.summarize(state.filtered);
    const sales = new Set(state.filtered.map((d) => d.sale).filter(Boolean)).size;
    const vend = new Set(state.filtered.flatMap((d) => d.vendors)).size;
    const cust = new Set(state.filtered.map((d) => d.customer)).size;
    const k = (label, value, sub, color, act) => `<div class="kpi" style="--k:${color}" ${act ? `data-kpi="${act}" title="Bấm để lọc" role="button"` : ''}><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div><div class="kpi-sub">${sub}</div></div>`;
    const money = (v) => { const t = U.money(v); const m = t.match(/^(.*?)\s(tỷ|tr|đ)$/); return m ? `${m[1]}<small> ${m[2]}</small>` : t; };
    $('#kpis').innerHTML = [
      k('Tổng thương vụ', U.int(s.n), `${sales} sale · ${vend} hãng · ${U.int(cust)} KH`, '#0959a2'),
      k('Tổng giá trị', money(s.value), `${U.int(s.withValue)} deal có giá trị`, '#073f7a'),
      k('Pipeline Active', money(s.valueActive), `${U.int(s.active)} deal đang Active`, STATUS_COLORS.Active, 'Active'),
      k('Won', U.int(s.won), `Giá trị ${U.money(s.valueWon)}`, STATUS_COLORS.Won, 'Won'),
      k('Lost', U.int(s.lost), `Giá trị ${U.money(s.valueLost)}`, STATUS_COLORS.Lost, 'Lost'),
      k('Tỷ lệ thắng', U.pct(s.winRate), `Won / (Won + Lost) – ${U.int(s.won + s.lost)} deal`, '#19a0a8'),
      k('Giá trị TB / deal', money(s.avg), `Trung vị: ${U.money(s.median)}`, '#5b4bb7'),
      k('Cần nhắc update', U.int(s.remind), `> 1 tháng chưa cập nhật`, '#eb6834', 'remind'),
      k('Quá timeline', U.int(s.overdue), 'Active nhưng đã qua close date', '#e34948', 'overdue'),
      k('Vấn đề dữ liệu', U.int(s.issues), 'deal có cờ [VẤN ĐỀ] cần rà soát', '#c2551c', 'issue'),
    ].join('');
  }

  // ---------------------------------------------------------------- insight panels
  const charts = {};
  function donut(id, canvasParent, labels, data, colors, onClick) {
    let ch = charts[id];
    if (!ch) {
      const cv = document.createElement('canvas');
      canvasParent.appendChild(cv);
      ch = charts[id] = new Chart(cv, {
        type: 'doughnut',
        data: { labels, datasets: [{ data, backgroundColor: colors, borderColor: 'rgba(255,255,255,.95)', borderWidth: 2, hoverOffset: 6 }] },
        options: {
          cutout: '68%', responsive: true, maintainAspectRatio: false, animation: { duration: 350 },
          plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` ${c.label}: ${U.int(c.parsed)}` } } },
          onClick: (e, els) => { if (els.length && onClick) onClick(els[0].index); },
          onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
        },
      });
    } else {
      ch.data.labels = labels;
      ch.data.datasets[0].data = data;
      ch.data.datasets[0].backgroundColor = colors;
      ch.update();
    }
    return ch;
  }
  function renderInsights() {
    const L = state.filtered;
    const n = L.length;
    // --- trạng thái + giai đoạn
    const st = STATUSES.map((s) => L.filter((d) => d.status === s));
    const stSel = state.filters.facets.status || new Set();
    const active = L.filter((d) => d.status === 'Active');
    const stages = CRMModel.groupBy(active, (d) => d.stage);
    const stageKeys = [...stages.keys()].sort(ordered(STAGES));
    const stageColors = ['#9ec5f4', '#6da7ec', '#3987e5', '#256abf', '#104281'];
    let el = $('#insStatus');
    if (!el.dataset.init) {
      el.innerHTML = `<div class="donut-wrap"><div class="donut-center"><div><b id="insN"></b><span>thương vụ</span></div></div></div><div id="insStatusBody"></div>`;
      el.dataset.init = 1;
    }
    $('#insN').textContent = U.int(n);
    donut('dStatus', $('.donut-wrap', el), STATUSES, st.map((x) => x.length), STATUSES.map((s) => STATUS_COLORS[s]), (i) => toggleFacet('status', STATUSES[i]));
    const sumV = (arr) => arr.reduce((a, d) => a + (d.value || 0), 0);
    $('#insStatusBody').innerHTML = `
      <div class="ins-title">Kết quả thương vụ</div>
      <ul class="ins-list">${STATUSES.map((s, i) => `<li data-fk="status" data-fv="${s}" class="${stSel.has(s) ? 'sel' : ''}"><span class="sw" style="background:${STATUS_COLORS[s]}"></span><span class="lab">${s} <span style="opacity:.75;font-size:12px">· ${U.money(sumV(st[i]))}</span></span><span class="val">${U.int(st[i].length)}</span><span class="pct">${n ? U.pct(st[i].length / n) : '–'}</span></li>`).join('')}</ul>
      <div class="ins-sep"></div>
      <div class="ins-title" style="font-size:11px">Thương vụ Active theo giai đoạn</div>
      <ul class="ins-list small">${stageKeys.map((k) => `<li data-fk="stage" data-fv="${esc(k)}"><span class="sw" style="background:${stageColors[STAGES.indexOf(k)] || '#cfd9e6'}"></span><span class="lab">${esc(k === NONE ? '(Chưa có)' : k)}</span><span class="val">${stages.get(k).length}</span><span class="pct">${U.money(sumV(stages.get(k)))}</span></li>`).join('') || '<li class="muted">Không có deal Active</li>'}</ul>`;

    // --- cờ nhắc + timeline
    const stl = STALE.filter((s) => s.k !== 'na').map((s) => ({ ...s, n: L.filter((d) => d.stale === s.k).length }));
    const na = L.filter((d) => d.stale === 'na').length;
    const considered = stl.reduce((a, s) => a + s.n, 0);
    const remind = stl.filter((s) => s.k !== '0').reduce((a, s) => a + s.n, 0);
    el = $('#insStale');
    if (!el.dataset.init) {
      el.innerHTML = `<div class="donut-wrap"><div class="donut-center"><div><b id="insR"></b><span>cần nhắc</span></div></div></div><div id="insStaleBody"></div>`;
      el.dataset.init = 1;
    }
    $('#insR').textContent = U.int(remind);
    donut('dStale', $('.donut-wrap', el), stl.map((s) => s.short), stl.map((s) => s.n), stl.map((s) => s.color), (i) => toggleFacet('stale', stl[i].k));
    const selS = state.filters.facets.stale || new Set();
    const selC = state.filters.facets.closeState || new Set();
    const cl = CLOSE.map((c) => ({ ...c, n: L.filter((d) => d.closeState === c.k && (d.status === 'Active')).length }));
    $('#insStaleBody').innerHTML = `
      <div class="ins-title">Tình trạng cập nhật (cờ nhắc)${state.staleActiveOnly ? ' · deal Active' : ''}</div>
      <ul class="ins-list">${stl.map((s) => `<li data-fk="stale" data-fv="${s.k}" class="${selS.has(s.k) ? 'sel' : ''}"><span class="sw" style="background:${s.color}"></span><span class="lab">${esc(s.label)}</span><span class="val">${U.int(s.n)}</span><span class="pct">${considered ? U.pct(s.n / considered) : '–'}</span></li>`).join('')}</ul>
      ${na ? `<div style="font-size:11.5px;opacity:.7;margin:2px 6px">+ ${U.int(na)} deal Won/Lost không xét cờ nhắc</div>` : ''}
      <div class="ins-sep"></div>
      <div class="ins-title" style="font-size:11px">Timeline dự án (deal Active)</div>
      <ul class="ins-list small">${cl.map((c) => `<li data-fk="closeState" data-fv="${c.k}" class="${selC.has(c.k) ? 'sel' : ''}"><span class="sw" style="background:${c.color}"></span><span class="lab">${esc(c.label)}</span><span class="val">${U.int(c.n)}</span></li>`).join('')}</ul>`;
  }

  // ---------------------------------------------------------------- charts
  const GRID = '#edf1f6', TICK = '#5b6b80';
  function baseOpts(horizontal, stacked, valueMode) {
    const valueAxis = {
      beginAtZero: true, stacked, grid: { color: GRID, drawTicks: false }, border: { display: false },
      ticks: { color: TICK, padding: 6, callback: (v) => (valueMode ? U.money(v) : U.int(v)), maxTicksLimit: 6 },
    };
    const catAxis = {
      stacked, grid: { display: false }, border: { color: '#cfd9e6' },
      ticks: { color: '#33445a', autoSkip: false, font: { size: 11.5 }, callback: function (v) { const l = this.getLabelForValue(v); return l.length > 26 ? l.slice(0, 25) + '…' : l; } },
    };
    if (!horizontal) catAxis.ticks.maxRotation = 50;
    return {
      indexAxis: horizontal ? 'y' : 'x',
      responsive: true, maintainAspectRatio: false, animation: { duration: 300 },
      layout: { padding: { top: horizontal ? 0 : 18, right: horizontal ? 60 : 6 } },
      scales: horizontal ? { x: valueAxis, y: catAxis } : { x: catAxis, y: valueAxis },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#0f1f33', padding: 10, cornerRadius: 8, titleFont: { weight: '700' },
          callbacks: { label: (c) => ` ${c.dataset.label ? c.dataset.label + ': ' : ''}${valueMode ? U.money(c.raw) : U.int(c.raw)}` },
        },
      },
      onHover: (e, els) => { e.native.target.style.cursor = els.length ? 'pointer' : 'default'; },
    };
  }
  // nhãn tổng ở đầu thanh (chỉ vẽ cho tổng, không gắn số lên mọi đoạn)
  const totalsPlugin = {
    id: 'totals',
    afterDatasetsDraw(chart, args, opts) {
      if (!opts || !opts.enabled) return;
      const { ctx } = chart;
      const horizontal = chart.options.indexAxis === 'y';
      const n = chart.data.labels.length;
      ctx.save();
      ctx.font = '700 11px "Segoe UI", Roboto, Arial';
      ctx.fillStyle = '#0f1f33';
      for (let i = 0; i < n; i++) {
        let sum = 0, edge = null;
        chart.data.datasets.forEach((ds, di) => {
          const meta = chart.getDatasetMeta(di);
          if (meta.hidden) return;
          const v = ds.data[i] || 0;
          sum += v;
          const el = meta.data[i];
          if (!el || !v) return;
          const e = horizontal ? el.x : el.y;
          edge = edge == null ? e : horizontal ? Math.max(edge, e) : Math.min(edge, e);
        });
        if (!sum || edge == null) continue;
        const txt = opts.money ? U.money(sum) : U.int(sum);
        const el0 = chart.getDatasetMeta(0).data[i];
        if (horizontal) { ctx.textAlign = 'left'; ctx.textBaseline = 'middle'; ctx.fillText(txt, edge + 6, el0.y); }
        else { ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(txt, el0.x, edge - 4); }
      }
      ctx.restore();
    },
  };
  Chart.register(totalsPlugin);
  Chart.defaults.font.family = '"Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif';
  Chart.defaults.color = TICK;

  function barChart(id, cfg) {
    const { labels, datasets, horizontal, stacked, valueMode, onClick, totals = true } = cfg;
    const ds = datasets.map((d) => Object.assign({
      borderRadius: 4, borderSkipped: 'start', maxBarThickness: horizontal ? 18 : 34, categoryPercentage: 0.78, barPercentage: 0.9,
      borderColor: '#fff', borderWidth: stacked ? { top: horizontal ? 0 : 1, right: horizontal ? 1 : 0 } : 0,
    }, d));
    let ch = charts[id];
    const opts = baseOpts(horizontal, stacked, valueMode);
    opts.plugins.totals = { enabled: totals, money: valueMode };
    opts.onClick = (e, els) => { if (els.length && onClick) onClick(els[0].index, els[0].datasetIndex); };
    if (ch) {
      ch.data.labels = labels;
      ch.data.datasets = ds;
      ch.options = opts;
      ch.update();
    } else {
      ch = charts[id] = new Chart(document.getElementById(id), { type: 'bar', data: { labels, datasets: ds }, options: opts });
    }
    return ch;
  }
  const valOf = (arr, mode) => (mode === 'value' ? arr.reduce((a, d) => a + (d.value || 0), 0) : arr.length);

  function renderCharts() {
    const L = state.filtered;
    const M = state.modes;
    $$('.seg[data-mode]').forEach((s) => $$('button', s).forEach((b) => b.classList.toggle('on', M[s.dataset.mode] === b.dataset.v)));

    // Sale (stacked theo trạng thái)
    const bySale = CRMModel.groupBy(L, (d) => d.sale);
    const saleKeys = [...bySale.keys()].sort((a, b) => valOf(bySale.get(b), M.sale) - valOf(bySale.get(a), M.sale));
    barChart('chSale', {
      labels: saleKeys.map((k) => (k === NONE ? '(Chưa có sale)' : k)), stacked: true, valueMode: M.sale === 'value',
      datasets: STATUSES.map((s) => ({ label: s, backgroundColor: STATUS_COLORS[s], data: saleKeys.map((k) => valOf(bySale.get(k).filter((d) => d.status === s), M.sale)) })),
      onClick: (i, di) => toggleFacet('sale', saleKeys[i]),
    });

    // Hãng (ngang) – giá trị tính theo hãng chính
    const byV = CRMModel.groupBy(L, (d) => d.vendors);
    const vVal = (k) => (M.vendor === 'value' ? byV.get(k).filter((d) => d.mainVendor === k).reduce((a, d) => a + (d.value || 0), 0) : byV.get(k).length);
    const vKeys = [...byV.keys()].sort((a, b) => vVal(b) - vVal(a));
    $('#boxVendor').style.height = Math.max(120, vKeys.length * 24 + 40) + 'px';
    barChart('chVendor', {
      labels: vKeys.map((k) => fmtFacet('vendor', k)), horizontal: true, valueMode: M.vendor === 'value',
      datasets: [{ label: M.vendor === 'value' ? 'Giá trị (hãng chính)' : 'Số deal', backgroundColor: '#0959a2', data: vKeys.map(vVal) }],
      onClick: (i) => toggleFacet('vendor', vKeys[i]),
    });

    // Giai đoạn (phễu – ramp thứ bậc một màu)
    const byS = CRMModel.groupBy(L, (d) => d.stage);
    const sKeys = [...byS.keys()].sort(ordered(STAGES));
    const ramp = ['#86b6ef', '#5598e7', '#2a78d6', '#1c5cab', '#104281'];
    barChart('chStage', {
      labels: sKeys.map((k) => (k === NONE ? '(Chưa có)' : k)), horizontal: true, valueMode: M.stage === 'value',
      datasets: [{ label: M.stage === 'value' ? 'Giá trị' : 'Số deal', backgroundColor: sKeys.map((k) => ramp[STAGES.indexOf(k)] || '#9aa8b8'), data: sKeys.map((k) => valOf(byS.get(k), M.stage)) }],
      onClick: (i) => toggleFacet('stage', sKeys[i]),
    });

    // Khoảng giá trị
    const byB = CRMModel.groupBy(L, (d) => d.band);
    const bKeys = [...BANDS.map((b) => b.k), NONE];
    barChart('chBand', {
      labels: bKeys.map((k) => (k === NONE ? 'Chưa có' : BANDS.find((b) => b.k === k).short)), valueMode: false,
      datasets: [{ label: 'Số deal', backgroundColor: bKeys.map((k) => (k === NONE ? '#b8c4d2' : '#0959a2')), data: bKeys.map((k) => (byB.get(k) || []).length) }],
      onClick: (i) => toggleFacet('band', bKeys[i]),
    });
    if (charts.chBand) {
      charts.chBand.options.plugins.tooltip.callbacks.afterLabel = (c) => ' Giá trị: ' + U.money((byB.get(bKeys[c.dataIndex]) || []).reduce((a, d) => a + (d.value || 0), 0));
    }

    // Loại HĐ
    const byC = CRMModel.groupBy(L, (d) => d.ctype);
    const cKeys = [...byC.keys()].sort((a, b) => (a === NONE) - (b === NONE) || byC.get(b).length - byC.get(a).length);
    const cColors = { New: '#2a78d6', Renew: '#17a673', 'Up - Cross sales': '#eb6834' };
    barChart('chCtype', {
      labels: cKeys.map((k) => fmtFacet('ctype', k)), valueMode: false,
      datasets: [{ label: 'Số deal', backgroundColor: cKeys.map((k) => cColors[k] || (k === NONE ? '#b8c4d2' : '#5b4bb7')), data: cKeys.map((k) => byC.get(k).length) }],
      onClick: (i) => toggleFacet('ctype', cKeys[i]),
    });
    if (charts.chCtype) charts.chCtype.options.plugins.tooltip.callbacks.afterLabel = (c) => ' Giá trị: ' + U.money(byC.get(cKeys[c.dataIndex]).reduce((a, d) => a + (d.value || 0), 0));

    // Nhóm KH
    const byG = CRMModel.groupBy(L, (d) => d.segment);
    const gKeys = [...byG.keys()].sort((a, b) => valOf(byG.get(b), M.segment) - valOf(byG.get(a), M.segment));
    barChart('chSegment', {
      labels: gKeys, horizontal: true, valueMode: M.segment === 'value',
      datasets: [{ label: M.segment === 'value' ? 'Giá trị' : 'Số deal', backgroundColor: '#1a8fd0', data: gKeys.map((k) => valOf(byG.get(k), M.segment)) }],
      onClick: (i) => toggleFacet('segment', gKeys[i]),
    });

    // Reseller
    const byR = CRMModel.groupBy(L, (d) => d.reseller);
    const rKeys = [...byR.keys()].sort((a, b) => (a === NONE) - (b === NONE) || valOf(byR.get(b), M.reseller) - valOf(byR.get(a), M.reseller));
    $('#boxReseller').style.height = Math.max(120, rKeys.length * 24 + 40) + 'px';
    barChart('chReseller', {
      labels: rKeys.map((k) => fmtFacet('reseller', k)), horizontal: true, valueMode: M.reseller === 'value',
      datasets: [{ label: M.reseller === 'value' ? 'Giá trị' : 'Số deal', backgroundColor: rKeys.map((k) => (k === NONE ? '#b8c4d2' : '#15aab3')), data: rKeys.map((k) => valOf(byR.get(k), M.reseller)) }],
      onClick: (i) => toggleFacet('reseller', rKeys[i]),
    });

    // Xu hướng theo tháng
    const field = M.trendBy === 'close' ? 'close' : 'created';
    $('#trendHint').textContent = field === 'close'
      ? 'số thương vụ theo tháng dự kiến close (Timeline dự án), chia theo trạng thái – bấm cột để lọc theo tháng'
      : 'số thương vụ tạo mới mỗi tháng, chia theo trạng thái hiện tại – bấm cột để lọc theo tháng';
    const withDate = L.filter((d) => d[field]);
    let months = [...new Set(withDate.map((d) => U.monthKey(d[field])))].sort();
    if (months.length) {
      const all = [];
      let [y, m] = months[0].split('-').map(Number);
      const [y2, m2] = months[months.length - 1].split('-').map(Number);
      while (y < y2 || (y === y2 && m <= m2)) { all.push(`${y}-${String(m).padStart(2, '0')}`); m++; if (m > 12) { m = 1; y++; } }
      months = all.length > 36 ? all.slice(field === 'close' ? 0 : -36, field === 'close' ? 36 : undefined) : all;
    }
    const byM = CRMModel.groupBy(withDate, (d) => U.monthKey(d[field]));
    barChart('chTrend', {
      labels: months.map(U.monthLabel), stacked: true, valueMode: M.trend === 'value',
      datasets: STATUSES.map((s) => ({ label: s, backgroundColor: STATUS_COLORS[s], data: months.map((k) => valOf((byM.get(k) || []).filter((d) => d.status === s), M.trend)) })),
      onClick: (i) => {
        const [y, m] = months[i].split('-').map(Number);
        const f = state.filters;
        f.dateField = field;
        f.from = U.iso(new Date(Date.UTC(y, m - 1, 1)));
        f.to = U.iso(new Date(Date.UTC(y, m, 0)));
        state.page = 1;
        refresh();
      },
    });
  }

  // ================================================================ TABLE
  const fmtClose = (d) => {
    if (!d.close) return '<span class="cs cs-far">Chưa có</span>';
    const cls = { overdue: 'cs-overdue', d30: 'cs-soon', d90: 'cs-mid', later: 'cs-far' }[d.closeState];
    const t = d.closeDays < 0 ? `quá ${-d.closeDays} ngày` : d.closeDays === 0 ? 'hôm nay' : `còn ${d.closeDays} ngày`;
    return `<span class="cs ${cls}">${t}</span>`;
  };
  const flagHtml = (d) => {
    if (d.stale === 'na') return '<span class="flag fl-na">Không xét</span>';
    if (d.stale === '0') return '<span class="flag fl-0">Bình thường</span>';
    return `<span class="flag fl-${d.stale}">${I.flag}Mức ${d.stale}</span>`;
  };
  const stBadge = (s) => (s ? `<span class="badge-st st-${esc(s)}">${esc(s)}</span>` : '');
  const vendorChips = (d) => d.vendors.map((v, i) => `<span class="vchip ${i === 0 && d.vendors.length > 1 ? 'main' : ''}">${esc(v)}</span>`).join('') || '<span class="muted">–</span>';

  const COLS = [
    { k: 'id', l: 'ID', def: true, grp: 'Thông tin chính', html: (d) => `<span class="cell-id">${esc(d.id)}</span>`, v: (d) => d.id },
    { k: 'name', l: 'Tên thương vụ', def: true, grp: 'Thông tin chính', html: (d) => `<div class="cell-clip cell-name" title="${esc(d.name)}">${esc(d.name)}</div>`, v: (d) => d.name, sticky: true },
    { k: 'customer', l: 'Khách hàng', def: true, grp: 'Thông tin chính', html: (d) => `<div class="cell-clip" title="${esc(d.customer)}">${esc(d.customer)}</div>`, v: (d) => d.customer },
    { k: 'segment', l: 'Nhóm KH', def: false, grp: 'Thông tin chính', html: (d) => esc(d.segment), v: (d) => d.segment },
    { k: 'sale', l: 'Sale', def: true, grp: 'Thông tin chính', html: (d) => `<span class="nowrap">${esc(d.sale || '–')}</span>`, v: (d) => d.sale },
    { k: 'reseller', l: 'Reseller', def: true, grp: 'Thông tin chính', html: (d) => esc(d.reseller || '–'), v: (d) => d.reseller },
    { k: 'stage', l: 'Giai đoạn', def: true, grp: 'Tiến độ', html: (d) => `<span class="nowrap">${esc(d.stage || '–')}</span>`, v: (d) => STAGES.indexOf(d.stage) },
    { k: 'status', l: 'Trạng thái', def: true, grp: 'Tiến độ', html: (d) => stBadge(d.status), v: (d) => STATUSES.indexOf(d.status) },
    { k: 'value', l: 'Giá trị (VND)', def: true, grp: 'Giá trị', r: true, html: (d) => (d.value ? `<span class="num nowrap" title="${U.money(d.value)}">${U.vnd(d.value)}</span>` : '<span class="muted">–</span>'), v: (d) => d.value },
    { k: 'band', l: 'Khoảng giá trị', def: false, grp: 'Giá trị', html: (d) => `<span class="nowrap">${esc(BAND_LABEL[d.band])}</span>`, v: (d) => [...BANDS.map((b) => b.k), NONE].indexOf(d.band) },
    { k: 'sizeCRM', l: 'Deal size', def: false, grp: 'Giá trị', html: (d) => esc(d.sizeCRM || '–'), v: (d) => d.value },
    { k: 'vendors', l: 'Hãng', def: true, grp: 'Sản phẩm', html: vendorChips, v: (d) => d.mainVendor },
    { k: 'ctype', l: 'Loại HĐ', def: true, grp: 'Sản phẩm', html: (d) => esc(d.ctype || '–'), v: (d) => d.ctype },
    { k: 'bom', l: 'Thông tin BOM', def: false, grp: 'Sản phẩm', html: (d) => `<div class="cell-clip" title="${esc(d.bom || '')}">${esc(d.bom || '–')}</div>`, v: (d) => d.bom },
    { k: 'close', l: 'Close date', def: true, grp: 'Thời gian', html: (d) => `<span class="nowrap">${U.date(d.close) || '–'}</span>`, v: (d) => d.close && d.close.getTime() },
    { k: 'closeState', l: 'Timeline', def: true, grp: 'Thời gian', html: fmtClose, v: (d) => d.closeDays },
    { k: 'closeQuarter', l: 'Quý dự kiến', def: false, grp: 'Thời gian', html: (d) => esc(d.closeQuarter || '–'), v: (d) => d.close && d.close.getTime() },
    { k: 'created', l: 'Ngày tạo', def: false, grp: 'Thời gian', html: (d) => `<span class="nowrap">${U.date(d.created)}</span>`, v: (d) => d.created && d.created.getTime() },
    { k: 'ageDays', l: 'Tuổi deal (ngày)', def: false, grp: 'Thời gian', r: true, html: (d) => U.int(d.ageDays), v: (d) => d.ageDays },
    { k: 'updated', l: 'Cập nhật cuối', def: true, grp: 'Thời gian', html: (d) => `<span class="nowrap" title="${U.dateTime(d.updated)}">${U.date(d.updated)}</span>`, v: (d) => d.updated && d.updated.getTime() },
    { k: 'staleDays', l: 'Số ngày chưa cập nhật', def: false, grp: 'Thời gian', r: true, html: (d) => U.int(d.staleDays), v: (d) => d.staleDays },
    { k: 'stale', l: 'Cờ nhắc', def: true, grp: 'Thời gian', html: flagHtml, v: (d) => (d.stale === 'na' ? -1 : +d.stale) },
    { k: 'failed', l: 'Lý do Failed', def: false, grp: 'Khác', html: (d) => `<div class="cell-clip" title="${esc(d.failed || '')}">${esc(d.failed || '–')}</div>`, v: (d) => d.failed },
    { k: 'issues', l: 'Vấn đề DL', def: true, grp: 'Khác', r: true, html: (d) => (d.issues.length ? `<span class="issue-dot" title="${esc(d.issues.join('\n'))}">${d.issues.length}</span>` : ''), v: (d) => d.issues.length },
    { k: 'quality', l: 'Cờ chất lượng', def: false, grp: 'Khác', html: (d) => `<div class="cell-clip" title="${esc(d.quality || '')}" style="max-width:480px">${esc(d.quality || '–')}</div>`, v: (d) => d.quality },
  ];
  const COLMAP = Object.fromEntries(COLS.map((c) => [c.k, c]));
  let visibleCols = new Set(LS.get('cols', COLS.filter((c) => c.def).map((c) => c.k)).filter((k) => COLMAP[k]));

  const PIVOTS = {
    sale: { label: 'Theo Sale', facet: 'sale', key: (d) => d.sale, name: 'Sale' },
    vendor: { label: 'Theo Hãng', facet: 'vendor', key: (d) => d.vendors, name: 'Hãng', mainOnly: true },
    reseller: { label: 'Theo Reseller', facet: 'reseller', key: (d) => d.reseller, name: 'Reseller' },
    customer: { label: 'Theo Khách hàng', facet: 'customer', key: (d) => d.customer, name: 'Khách hàng' },
    segment: { label: 'Theo Nhóm KH', facet: 'segment', key: (d) => d.segment, name: 'Nhóm khách hàng' },
  };

  function renderTabs() {
    const L = state.filtered;
    const cnt = (fn) => CRMModel.groupBy(L, fn).size;
    const tabs = [['deals', 'Danh sách thương vụ', L.length], ...Object.entries(PIVOTS).map(([k, p]) => [k, p.label, cnt(p.key)]), ['matrix', 'Sale × Giai đoạn', null]];
    $('#tabs').innerHTML = tabs.map(([k, l, n]) => `<button class="tab ${state.tab === k ? 'on' : ''}" data-tab="${k}">${esc(l)}${n != null ? `<span class="badge">${U.int(n)}</span>` : ''}</button>`).join('');
  }

  function sortList(list, key, dir, getter) {
    return [...list].sort((a, b) => {
      const va = getter(a, key), vb = getter(b, key);
      const na = va == null || va === '' || (typeof va === 'number' && isNaN(va)), nb = vb == null || vb === '' || (typeof vb === 'number' && isNaN(vb));
      if (na && nb) return 0;
      if (na) return 1;
      if (nb) return -1;
      if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
      return String(va).localeCompare(String(vb), 'vi', { numeric: true }) * dir;
    });
  }

  function renderTable() {
    const tb = $('#tableToolbar');
    if (state.tab === 'deals') renderDeals(tb);
    else if (state.tab === 'matrix') renderMatrix(tb);
    else renderPivot(tb, state.tab);
  }

  function renderDeals(tb) {
    tb.innerHTML = `
      <span class="muted" style="font-size:13px">Bấm vào dòng để xem chi tiết · bấm tiêu đề cột để sắp xếp</span><span class="grow"></span>
      <button class="btn btn-sm" data-act="cols">${I.cols}Ẩn / hiện cột <span class="badge" style="background:#e8f0fa;color:#073f7a;border-radius:99px;padding:0 6px;font-size:11px">${visibleCols.size}/${COLS.length}</span></button>
      <button class="btn btn-sm" data-act="density" title="Đổi mật độ hiển thị">${I.rows}${state.compact ? 'Thoáng' : 'Gọn'}</button>
      <button class="btn btn-sm" data-act="exportView">${I.excel}Xuất bảng đang xem</button>`;
    const cols = COLS.filter((c) => visibleCols.has(c.k));
    const sorted = sortList(state.filtered, state.sort.key, state.sort.dir, (d, k) => COLMAP[k].v(d));
    const ps = state.pageSize || sorted.length || 1;
    const pages = Math.max(1, Math.ceil(sorted.length / ps));
    const start = (state.page - 1) * ps;
    const rows = sorted.slice(start, start + ps);
    const s = CRMModel.summarize(state.filtered);
    $('#tblWrap').innerHTML = `
      <table class="data ${state.compact ? 'compact' : ''}">
        <thead><tr>${cols.map((c) => `<th data-sort="${c.k}" class="${c.r ? 'r' : ''} ${state.sort.key === c.k ? 'sorted' : ''}">${esc(c.l)}<span class="sort">${state.sort.key === c.k ? (state.sort.dir > 0 ? '▲' : '▼') : '⇅'}</span></th>`).join('')}</tr></thead>
        <tbody>${rows.map((d) => `<tr data-open="${esc(d.id)}">${cols.map((c) => `<td class="${c.r ? 'r' : ''}">${c.html(d)}</td>`).join('')}</tr>`).join('') ||
          `<tr><td colspan="${cols.length}" style="text-align:center;padding:40px" class="muted">Không có thương vụ nào khớp bộ lọc</td></tr>`}</tbody>
        <tfoot class="tfoot"><tr>${cols.map((c, i) => `<td class="${c.r ? 'r' : ''}">${i === 0 ? `Tổng: ${U.int(s.n)}` : c.k === 'value' ? `<span class="num nowrap">${U.vnd(s.value)}</span>` : c.k === 'issues' ? U.int(s.issues) : ''}</td>`).join('')}</tr></tfoot>
      </table>`;
    renderPager(sorted.length, pages);
  }

  function renderPager(total, pages) {
    const p = state.page;
    const btns = [];
    const add = (i) => btns.push(`<button data-page="${i}" class="${i === p ? 'on' : ''}">${i}</button>`);
    if (pages <= 7) for (let i = 1; i <= pages; i++) add(i);
    else {
      add(1);
      if (p > 3) btns.push('<span>…</span>');
      for (let i = Math.max(2, p - 1); i <= Math.min(pages - 1, p + 1); i++) add(i);
      if (p < pages - 2) btns.push('<span>…</span>');
      add(pages);
    }
    const ps = state.pageSize;
    const from = total ? (p - 1) * (ps || total) + 1 : 0, to = Math.min(total, p * (ps || total));
    $('#pager').innerHTML = `<span>Hiển thị ${U.int(from)}–${U.int(to)} / ${U.int(total)}</span><span class="grow"></span>
      <label>Số dòng / trang <select id="pageSize">${[25, 50, 100, 200, 0].map((n) => `<option value="${n}" ${n === ps ? 'selected' : ''}>${n || 'Tất cả'}</option>`).join('')}</select></label>
      <button data-page="${p - 1}" ${p <= 1 ? 'disabled' : ''}>‹</button>${btns.join('')}<button data-page="${p + 1}" ${p >= pages ? 'disabled' : ''}>›</button>`;
    $('#pager').classList.remove('hidden');
  }

  function pivotRows(list, P) {
    const g = CRMModel.groupBy(list, P.key);
    return [...g.entries()].map(([k, arr]) => {
      const s = CRMModel.summarize(arr);
      let value = s.value, valueActive = s.valueActive, valueWon = s.valueWon;
      if (P.mainOnly) {
        const main = arr.filter((d) => d.mainVendor === k);
        const sm = CRMModel.summarize(main);
        value = sm.value; valueActive = sm.valueActive; valueWon = sm.valueWon;
      }
      return { k, name: P.facet === 'vendor' || P.facet === 'reseller' ? fmtFacet(P.facet, k) : k === NONE ? '(Trống)' : k, n: s.n, active: s.active, won: s.won, lost: s.lost,
        value, valueActive, valueWon, winRate: s.winRate, avg: s.withValue ? value / s.withValue : null, remind: s.remind, overdue: s.overdue, soon30: s.soon30, issues: s.issues };
    });
  }
  const PIVOT_COLS = [
    ['name', null, false], ['n', 'Số deal', true], ['active', 'Active', true], ['won', 'Won', true], ['lost', 'Lost', true],
    ['value', 'Tổng giá trị', true], ['valueActive', 'Pipeline Active', true], ['valueWon', 'Giá trị Won', true], ['winRate', 'Tỷ lệ thắng', true],
    ['avg', 'Giá trị TB', true], ['remind', 'Cần nhắc', true], ['overdue', 'Quá timeline', true], ['soon30', 'Close ≤ 30 ngày', true], ['issues', 'Vấn đề DL', true],
  ];
  function renderPivot(tb, key) {
    const P = PIVOTS[key];
    tb.innerHTML = `<span class="muted" style="font-size:13px">Thống kê theo kết quả lọc hiện tại · bấm vào dòng để lọc theo ${esc(P.name.toLowerCase())}${P.mainOnly ? ' · Giá trị tính cho deal có hãng này là <b>hãng chính</b>' : ''}</span><span class="grow"></span>
      <button class="btn btn-sm" data-act="exportReport">${I.excel}Xuất Excel tổng hợp</button>`;
    const rows = sortList(pivotRows(state.filtered, P), state.pivotSort.key, state.pivotSort.dir, (r, k) => r[k]);
    const tot = CRMModel.summarize(state.filtered);
    const fmt = (k, v) => (['value', 'valueActive', 'valueWon', 'avg'].includes(k) ? (v ? `<span class="num nowrap" title="${U.vnd(v)}">${U.money(v)}</span>` : '<span class="muted">–</span>') : k === 'winRate' ? U.pct(v) : U.int(v));
    const sel = state.filters.facets[P.facet] || new Set();
    $('#tblWrap').innerHTML = `<table class="data ${state.compact ? 'compact' : ''}"><thead><tr>${PIVOT_COLS.map(([k, l, r]) => `<th data-psort="${k}" class="${r ? 'r' : ''} ${state.pivotSort.key === k ? 'sorted' : ''}">${esc(l || P.name)}<span class="sort">${state.pivotSort.key === k ? (state.pivotSort.dir > 0 ? '▲' : '▼') : '⇅'}</span></th>`).join('')}</tr></thead>
      <tbody>${rows.map((r) => `<tr data-pivot="${esc(P.facet)}" data-pv="${esc(r.k)}" style="${sel.has(r.k) ? 'outline:2px solid #f06723;outline-offset:-2px' : ''}">${PIVOT_COLS.map(([k, , rr]) => `<td class="${rr ? 'r' : ''}">${k === 'name' ? `<b>${esc(r.name)}</b>` : fmt(k, r[k])}</td>`).join('')}</tr>`).join('')}</tbody>
      <tfoot class="tfoot"><tr><td>Tổng (${U.int(rows.length)})</td><td class="r">${U.int(tot.n)}</td><td class="r">${U.int(tot.active)}</td><td class="r">${U.int(tot.won)}</td><td class="r">${U.int(tot.lost)}</td>
        <td class="r">${U.money(tot.value)}</td><td class="r">${U.money(tot.valueActive)}</td><td class="r">${U.money(tot.valueWon)}</td><td class="r">${U.pct(tot.winRate)}</td><td class="r">${U.money(tot.avg)}</td>
        <td class="r">${U.int(tot.remind)}</td><td class="r">${U.int(tot.overdue)}</td><td class="r">${U.int(tot.soon30)}</td><td class="r">${U.int(tot.issues)}</td></tr></tfoot></table>`;
    $('#pager').classList.add('hidden');
  }

  function renderMatrix(tb) {
    const mode = state.modes.matrix || 'count';
    tb.innerHTML = `<span class="muted" style="font-size:13px">Số thương vụ của từng sale ở mỗi giai đoạn · bấm ô để lọc</span><span class="grow"></span>
      <div class="seg" data-mode="matrix"><button data-v="count" class="${mode === 'count' ? 'on' : ''}">Số deal</button><button data-v="value" class="${mode === 'value' ? 'on' : ''}">Giá trị</button></div>`;
    const bySale = CRMModel.groupBy(state.filtered, (d) => d.sale);
    const stages = [...new Set(state.filtered.map((d) => d.stage || NONE))].sort(ordered(STAGES));
    const saleKeys = [...bySale.keys()].sort((a, b) => bySale.get(b).length - bySale.get(a).length);
    let max = 0;
    const cell = (arr) => valOf(arr, mode);
    saleKeys.forEach((s) => stages.forEach((st) => { max = Math.max(max, cell(bySale.get(s).filter((d) => (d.stage || NONE) === st))); }));
    const fmt = (v) => (mode === 'value' ? (v ? U.money(v) : '') : v ? U.int(v) : '');
    const shade = (v) => {
      if (!v || !max) return '';
      const t = v / max;
      const steps = ['#eaf3fd', '#cde2fb', '#9ec5f4', '#6da7ec', '#3987e5', '#256abf'];
      const c = steps[Math.min(steps.length - 1, Math.floor(t * steps.length))];
      return `background:${c};color:${t > 0.55 ? '#fff' : '#0f1f33'};font-weight:700`;
    };
    $('#tblWrap').innerHTML = `<table class="data"><thead><tr><th>Sale</th>${stages.map((s) => `<th class="r">${esc(s === NONE ? '(Chưa có)' : s)}</th>`).join('')}<th class="r">Tổng</th></tr></thead>
      <tbody>${saleKeys.map((s) => {
        const arr = bySale.get(s);
        return `<tr><td><b>${esc(s === NONE ? '(Chưa có sale)' : s)}</b></td>${stages.map((st) => { const v = cell(arr.filter((d) => (d.stage || NONE) === st)); return `<td class="r" data-mx="${esc(s)}" data-ms="${esc(st)}" style="${shade(v)}">${fmt(v)}</td>`; }).join('')}<td class="r"><b>${fmt(cell(arr))}</b></td></tr>`;
      }).join('')}</tbody>
      <tfoot class="tfoot"><tr><td>Tổng</td>${stages.map((st) => `<td class="r">${fmt(cell(state.filtered.filter((d) => (d.stage || NONE) === st)))}</td>`).join('')}<td class="r">${fmt(cell(state.filtered))}</td></tr></tfoot></table>`;
    $('#pager').classList.add('hidden');
  }

  function openColsPop(anchor) {
    const grps = [...new Set(COLS.map((c) => c.grp))];
    const pop = openPop(anchor, `
      <div class="pop-head"><div class="pop-title">Ẩn / hiện cột <small>${COLS.length} cột</small></div></div>
      <div class="pop-list colmenu">${grps.map((g) => `<div class="grp">${esc(g)}</div>` + COLS.filter((c) => c.grp === g).map((c) => `<label class="opt"><input type="checkbox" value="${c.k}" ${visibleCols.has(c.k) ? 'checked' : ''}><span class="o-label">${esc(c.l)}</span></label>`).join('')).join('')}</div>
      <div class="pop-foot"><button class="btn btn-ghost btn-sm" data-pa="def">Mặc định</button><button class="btn btn-sm" data-pa="all">Hiện tất cả</button></div>`, 300);
    if (!pop) return;
    const save = () => { LS.set('cols', [...visibleCols]); renderTable(); };
    pop.addEventListener('change', (e) => {
      if (e.target.type !== 'checkbox') return;
      e.target.checked ? visibleCols.add(e.target.value) : visibleCols.delete(e.target.value);
      if (!visibleCols.size) { visibleCols.add('name'); e.target.checked = e.target.value === 'name'; }
      save();
    });
    pop.addEventListener('click', (e) => {
      const a = e.target.closest('[data-pa]');
      if (!a) return;
      visibleCols = new Set(a.dataset.pa === 'all' ? COLS.map((c) => c.k) : COLS.filter((c) => c.def).map((c) => c.k));
      $$('input', pop).forEach((cb) => { cb.checked = visibleCols.has(cb.value); });
      save();
    });
  }

  // ================================================================ SEARCH + SUGGEST
  let sgItems = [], sgActive = -1, qTimer = null;
  function highlight(text, fq) {
    if (!fq) return esc(text);
    let folded = '';
    const map = [];
    for (let i = 0; i < text.length; i++) {
      const f = U.fold(text[i]);
      for (let j = 0; j < f.length; j++) { folded += f[j]; map.push(i); }
    }
    const idx = folded.indexOf(fq);
    if (idx < 0) return esc(text);
    const a = map[idx], b = map[idx + fq.length - 1] + 1;
    return esc(text.slice(0, a)) + '<mark>' + esc(text.slice(a, b)) + '</mark>' + esc(text.slice(b));
  }
  function buildSuggest(q) {
    const fq = U.fold(q.trim());
    if (!fq) return [];
    const groups = [];
    const add = (title, type, entries, limit) => {
      const hits = entries.filter((e) => e.f.includes(fq)).sort((a, b) => (b.f.startsWith(fq) - a.f.startsWith(fq)) || b.n - a.n).slice(0, limit);
      if (hits.length) groups.push({ title, items: hits.map((h) => ({ ...h, type })) });
    };
    const cnt = (fn) => {
      const m = new Map();
      state.deals.forEach((d) => { let v = fn(d); (Array.isArray(v) ? v : [v]).forEach((x) => { if (x) m.set(x, (m.get(x) || 0) + 1); }); });
      return [...m.entries()].map(([label, n]) => ({ label, n, f: U.fold(label) }));
    };
    add('Khách hàng', 'customer', cnt((d) => d.customer), 6);
    add('Thương vụ', 'deal', state.deals.map((d) => ({ label: d.name, n: 0, f: U.fold(d.name + ' ' + d.id), deal: d })), 6);
    add('Sale', 'sale', cnt((d) => d.sale), 4);
    add('Hãng', 'vendor', cnt((d) => d.vendors), 4);
    add('Reseller', 'reseller', cnt((d) => d.reseller), 4);
    add('Nhóm khách hàng', 'segment', cnt((d) => d.segment), 3);
    add('Giai đoạn', 'stage', cnt((d) => d.stage), 3);
    return groups;
  }
  const SG_ICON = { customer: 'KH', deal: '#', sale: 'S', vendor: 'H', reseller: 'R', segment: 'N', stage: 'GĐ', all: '⌕' };
  function showSuggest() {
    const q = $('#q').value;
    const box = $('#suggest');
    const groups = buildSuggest(q);
    if (!q.trim()) { hideSuggest(); return; }
    const fq = U.fold(q.trim());
    sgItems = [{ type: 'all', label: q.trim() }];
    let html = `<div class="sg-item" data-i="0"><span class="sg-ico">${SG_ICON.all}</span><div class="sg-main"><div class="sg-label">Lọc mọi trường chứa “<b>${esc(q.trim())}</b>”</div><div class="sg-sub">${U.int(state.deals.filter((d) => U.fold(q).split(/\s+/).filter(Boolean).every((t) => d._s.includes(t))).length)} thương vụ khớp · Enter</div></div></div>`;
    for (const g of groups) {
      html += `<div class="sg-group">${esc(g.title)}</div>`;
      for (const it of g.items) {
        const i = sgItems.push(it) - 1;
        const sub = it.type === 'deal' ? `#${esc(it.deal.id)} · ${esc(it.deal.customer)} · ${esc(it.deal.sale || '')} · ${esc(it.deal.status || '')}` : '';
        html += `<div class="sg-item" data-i="${i}"><span class="sg-ico">${SG_ICON[it.type]}</span><div class="sg-main"><div class="sg-label">${highlight(it.label, fq)}</div>${sub ? `<div class="sg-sub">${sub}</div>` : ''}</div>${it.type !== 'deal' ? `<span class="sg-count">${U.int(it.n)} deal</span>` : ''}</div>`;
      }
    }
    html += '<div class="sg-foot">↑ ↓ để chọn · Enter để áp dụng · Esc để đóng · Không phân biệt dấu / hoa thường</div>';
    box.innerHTML = html;
    box.classList.remove('hidden');
    sgActive = -1;
  }
  function hideSuggest() { $('#suggest').classList.add('hidden'); sgActive = -1; }
  function pickSuggest(i) {
    const it = sgItems[i];
    if (!it) return;
    hideSuggest();
    if (it.type === 'all') { state.filters.q = it.label; }
    else if (it.type === 'deal') { openDrawer(it.deal.id); return; }
    else {
      const f = state.filters;
      if (!f.facets[it.type]) f.facets[it.type] = new Set();
      f.facets[it.type].add(it.label);
      f.q = '';
      $('#q').value = '';
    }
    $('#searchBox').classList.toggle('has-q', !!$('#q').value);
    state.page = 1;
    refresh();
  }
  function bindSearch() {
    const q = $('#q');
    q.addEventListener('input', () => {
      $('#searchBox').classList.toggle('has-q', !!q.value);
      showSuggest();
      clearTimeout(qTimer);
      qTimer = setTimeout(() => { state.filters.q = q.value.trim(); state.page = 1; refresh(); }, 220);
    });
    q.addEventListener('focus', () => { if (q.value.trim()) showSuggest(); });
    q.addEventListener('keydown', (e) => {
      const items = $$('.sg-item', $('#suggest'));
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        if ($('#suggest').classList.contains('hidden')) showSuggest();
        e.preventDefault();
        sgActive = (sgActive + (e.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length;
        items.forEach((x, i) => x.classList.toggle('active', i === sgActive));
        if (items[sgActive]) items[sgActive].scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (sgActive >= 0 && items[sgActive]) pickSuggest(+items[sgActive].dataset.i);
        else { clearTimeout(qTimer); state.filters.q = q.value.trim(); hideSuggest(); state.page = 1; refresh(); }
      } else if (e.key === 'Escape') { hideSuggest(); q.blur(); }
    });
    $('#suggest').addEventListener('mousedown', (e) => {
      const it = e.target.closest('.sg-item');
      if (it) { e.preventDefault(); pickSuggest(+it.dataset.i); }
    });
    $('#clearQ').addEventListener('click', () => { q.value = ''; state.filters.q = ''; $('#searchBox').classList.remove('has-q'); hideSuggest(); refresh(); });
    document.addEventListener('keydown', (e) => {
      if (e.key === '/' && state.loaded && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) { e.preventDefault(); q.focus(); q.select(); }
      if (e.key === 'Escape') { closePop(); closeOverlay(); }
    });
  }

  // ================================================================ DRAWER / MODALS
  let overlayEls = [];
  function closeOverlay() { overlayEls.forEach((e) => e.remove()); overlayEls = []; }
  function showOverlay(panel) {
    closeOverlay();
    const ov = document.createElement('div');
    ov.className = 'overlay';
    ov.addEventListener('click', closeOverlay);
    document.body.append(ov, panel);
    overlayEls = [ov, panel];
  }
  function openDrawer(id) {
    const d = state.deals.find((x) => String(x.id) === String(id));
    if (!d) return;
    const el = document.createElement('aside');
    el.className = 'drawer';
    const kv = (k, v) => `<dt>${k}</dt><dd>${v || '<span class="muted">–</span>'}</dd>`;
    el.innerHTML = `
      <div class="drawer-head">
        <div class="dh-meta"><span>#${esc(d.id)}</span>${stBadge(d.status)}${flagHtml(d)}</div>
        <h3>${esc(d.name)}</h3>
        <div class="dh-meta">${esc(d.customer)} · ${esc(d.segment)}</div>
        <button class="close-x" data-close>${I.x.replace('<svg', '<svg width="16" height="16"')}</button>
      </div>
      <div class="drawer-body">
        <div class="dsec"><h4>Mốc thời gian</h4>
          <div class="timeline">
            <div class="tl-pt"><i></i><b>${U.date(d.created) || '–'}</b><span class="muted">Ngày tạo${d.ageDays != null ? ` · ${d.ageDays} ngày trước` : ''}</span></div>
            <div class="tl-pt"><i></i><b>${U.date(d.updated) || '–'}</b><span class="muted">Cập nhật cuối${d.staleDays != null ? ` · ${d.staleDays} ngày` : ''}</span></div>
            <div class="tl-pt"><i></i><b>${U.date(d.close) || '–'}</b><span class="muted">Close date · ${fmtClose(d)}</span></div>
          </div>
        </div>
        <div class="dsec"><h4>Thông tin thương vụ</h4><dl class="kv">
          ${kv('Giá trị', d.value ? `${U.vnd(d.value)} VND <span class="muted">(${U.money(d.value)})</span>` : '')}
          ${kv('Khoảng giá trị', esc(BAND_LABEL[d.band]))}
          ${kv('Deal size', esc(d.sizeCRM))}
          ${kv('Sale', esc(d.sale))}
          ${kv('Reseller', esc(d.reseller))}
          ${kv('Giai đoạn', esc(d.stage))}
          ${kv('Trạng thái', stBadge(d.status))}
          ${kv('Hãng', vendorChips(d))}
          ${kv('Loại HĐ', esc(d.ctype))}
          ${kv('Khách hàng', esc(d.customer))}
          ${kv('Nhóm KH', esc(d.segment))}
          ${kv('Lý do Failed', esc(d.failed))}
        </dl></div>
        ${d.bom ? `<div class="dsec"><h4>Thông tin BOM</h4><div style="font-size:13px;line-height:1.55;white-space:pre-wrap">${esc(d.bom)}</div></div>` : ''}
        ${d.issues.length ? `<div class="dsec" style="border-color:#fbc8ab"><h4 style="color:#c2551c">Vấn đề dữ liệu cần kiểm tra (${d.issues.length})</h4><ul class="issues-list warn">${d.issues.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''}
        ${d.handled.length ? `<div class="dsec"><h4>Đã xử lý khi convert</h4><ul class="issues-list ok">${d.handled.map((x) => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''}
      </div>
      <div class="drawer-actions">
        <button class="btn btn-sm" data-da="cust">${I.filter}Lọc theo KH này</button>
        ${d.sale ? `<button class="btn btn-sm" data-da="sale">${I.filter}Lọc theo sale</button>` : ''}
        <span style="flex:1"></span>
        <button class="btn btn-primary btn-sm" data-da="copy">${I.copy}Sao chép thông tin</button>
      </div>`;
    el.addEventListener('click', async (e) => {
      if (e.target.closest('[data-close]')) { closeOverlay(); return; }
      const a = e.target.closest('[data-da]');
      if (!a) return;
      if (a.dataset.da === 'cust') { closeOverlay(); toggleFacet('customer', d.customer, true); }
      if (a.dataset.da === 'sale') { closeOverlay(); toggleFacet('sale', d.sale, true); }
      if (a.dataset.da === 'copy') {
        const t = [`[#${d.id}] ${d.name}`, `Khách hàng: ${d.customer} (${d.segment})`, `Sale: ${d.sale || '–'} · Reseller: ${d.reseller || '–'}`,
          `Giai đoạn: ${d.stage || '–'} · Trạng thái: ${d.status || '–'}`, `Giá trị: ${d.value ? U.vnd(d.value) + ' VND' : '–'} · Hãng: ${d.vendors.join(', ') || '–'}`,
          `Close date: ${U.date(d.close) || '–'} · Cập nhật cuối: ${U.date(d.updated) || '–'}`].join('\n');
        toast((await copyText(t)) ? 'Đã sao chép thông tin thương vụ' : 'Không sao chép được', 'ok');
      }
    });
    showOverlay(el);
  }

  // ---- soạn nội dung nhắc
  function openRemind() {
    const list = state.filtered.filter((d) => d.needRemind);
    if (!list.length) { toast('Không có thương vụ cần nhắc trong kết quả lọc'); return; }
    const by = CRMModel.groupBy(list, (d) => d.sale);
    const sales = [...by.keys()].sort((a, b) => by.get(b).length - by.get(a).length);
    let cur = sales[0], minLv = 1;
    const el = document.createElement('div');
    el.className = 'modal';
    el.innerHTML = `
      <div class="modal-head">${I.mail.replace('<svg', '<svg width="20" height="20" style="color:#f06723"')}<h3>Soạn nội dung nhắc cập nhật CRM</h3><button class="close-x" data-close>${I.x.replace('<svg', '<svg width="15" height="15"')}</button></div>
      <div class="modal-body">
        <div class="form-row"><label>Mức cờ tối thiểu <select id="rmLv">${[1, 2, 3, 4].map((l) => `<option value="${l}">Mức ${l} trở lên</option>`).join('')}</select></label>
          <label>Hạn cập nhật <input type="text" id="rmDue" value="${U.date(new Date(state.refDate.getTime() + 7 * U.DAY))}" style="width:110px"></label>
          <span class="muted" style="font-size:12.5px">Danh sách theo kết quả lọc hiện tại, nhóm theo sale.</span></div>
        <div class="remind-layout"><div class="remind-sales" id="rmSales"></div><textarea class="remind-text" id="rmText"></textarea></div>
      </div>
      <div class="modal-foot"><span class="muted" style="font-size:12.5px;margin-right:auto">Nội dung có thể chỉnh sửa trước khi sao chép</span>
        <button class="btn" data-ra="all">${I.copy}Sao chép tất cả sale</button><button class="btn btn-primary" data-ra="one">${I.copy}Sao chép nội dung này</button></div>`;
    const textFor = (sale) => {
      const arr = by.get(sale).filter((d) => +d.stale >= minLv).sort((a, b) => b.staleDays - a.staleDays);
      if (!arr.length) return '';
      const due = $('#rmDue', el).value;
      return `Kính gửi anh/chị ${sale === NONE ? '(chưa có sale phụ trách)' : sale},\n\n` +
        `Phòng Quản lý sản phẩm rà soát dữ liệu CRM (tính đến ${U.date(state.refDate)}) thấy ${arr.length} thương vụ anh/chị phụ trách đã lâu chưa được cập nhật:\n\n` +
        arr.map((d, i) => `${i + 1}. [#${d.id}] ${d.name}\n   KH: ${d.customer} · Giai đoạn: ${d.stage || '–'} · Giá trị: ${d.value ? U.money(d.value) : 'chưa có'}\n   Cập nhật lần cuối: ${U.date(d.updated)} (${d.staleDays} ngày – Mức ${d.stale})${d.overdue ? ` · Close date ${U.date(d.close)} đã qua` : ''}`).join('\n') +
        `\n\nĐề nghị anh/chị cập nhật giai đoạn, giá trị, timeline dự án và trạng thái (Won/Lost nếu đã có kết quả) trên CRM trước ngày ${due}.\n\nTrân trọng,\nPhòng Quản lý sản phẩm – Công ty Nam Trường Sơn Hà Nội`;
    };
    const draw = () => {
      $('#rmSales', el).innerHTML = sales.map((s) => {
        const n = by.get(s).filter((d) => +d.stale >= minLv).length;
        return `<button data-s="${esc(s)}" class="${s === cur ? 'on' : ''}" ${n ? '' : 'style="opacity:.45"'}><span>${esc(s === NONE ? '(Chưa có sale)' : s)}</span><span>${n}</span></button>`;
      }).join('');
      $('#rmText', el).value = textFor(cur) || 'Không có thương vụ ở mức cờ này.';
    };
    el.addEventListener('click', async (e) => {
      if (e.target.closest('[data-close]')) { closeOverlay(); return; }
      const s = e.target.closest('[data-s]');
      if (s) { cur = s.dataset.s; draw(); return; }
      const a = e.target.closest('[data-ra]');
      if (a) {
        const t = a.dataset.ra === 'one' ? $('#rmText', el).value : sales.map(textFor).filter(Boolean).join('\n\n' + '-'.repeat(60) + '\n\n');
        toast((await copyText(t)) ? 'Đã sao chép nội dung nhắc' : 'Không sao chép được – hãy bôi đen và Ctrl+C', 'ok');
      }
    });
    showOverlay(el);
    $('#rmLv', el).addEventListener('change', (e) => { minLv = +e.target.value; draw(); });
    $('#rmDue', el).addEventListener('input', draw);
    draw();
  }

  // ---- xuất báo cáo
  function defaultReportMonth() {
    let max = null;
    state.deals.forEach((d) => { if (d.updated && (!max || d.updated > max)) max = d.updated; });
    const base = max || state.refDate;
    return { m: base.getUTCMonth() + 1, y: base.getUTCFullYear() };
  }
  function openExport() {
    const { m, y } = defaultReportMonth();
    const filtered = hasAnyFilter();
    const el = document.createElement('div');
    el.className = 'modal';
    el.style.width = 'min(640px, calc(100vw - 24px))';
    el.innerHTML = `
      <div class="modal-head">${I.excel.replace('<svg', '<svg width="20" height="20" style="color:#17a673"')}<h3>Xuất file Excel tổng hợp</h3><button class="close-x" data-close>${I.x.replace('<svg', '<svg width="15" height="15"')}</button></div>
      <div class="modal-body">
        <div class="form-row">
          <label>Tháng <select id="exM">${Array.from({ length: 12 }, (_, i) => `<option ${i + 1 === m ? 'selected' : ''}>${i + 1}</option>`).join('')}</select></label>
          <label>Năm <input type="number" id="exY" value="${y}" style="width:90px"></label>
        </div>
        <div class="form-row">
          <label><input type="radio" name="exScope" value="filtered" ${filtered ? 'checked' : ''} ${filtered ? '' : 'disabled'}> Theo bộ lọc hiện tại (${U.int(state.filtered.length)} thương vụ)</label>
          <label><input type="radio" name="exScope" value="all" ${filtered ? '' : 'checked'}> Toàn bộ dữ liệu (${U.int(state.deals.length)} thương vụ)</label>
        </div>
        <div class="muted" style="font-size:12.5px;margin-bottom:6px">Tên file:</div>
        <div class="fname-preview" id="exName"></div>
        <div style="margin-top:14px;font-weight:700;font-size:13px;color:#073f7a">File gồm các sheet:</div>
        <ul class="sheet-list">
          <li>Tổng quan – KPI &amp; các bảng phân bổ</li><li>Danh sách thương vụ – đầy đủ cột + cột phân tích bổ sung</li>
          <li>Theo Sale · Theo Hãng · Theo Reseller</li><li>Theo Nhóm KH · Theo Khách hàng</li>
          <li>Sale × Giai đoạn (số deal &amp; giá trị)</li><li>Forecast theo quý close (pipeline Active)</li>
          <li>Cần nhắc cập nhật – theo cờ mức 1→4</li><li>Vấn đề dữ liệu – cờ [VẤN ĐỀ] cần rà soát</li>
          <li>Data chuẩn – dữ liệu template sau convert</li><li>Ghi chú – định nghĩa các chỉ số</li>
        </ul>
      </div>
      <div class="modal-foot"><button class="btn" data-close>Huỷ</button><button class="btn btn-orange" data-ex>${I.download}Tải file Excel</button></div>`;
    const name = () => `Tổng hợp CRM_ Tháng ${$('#exM', el).value}.${$('#exY', el).value}.xlsx`;
    const upd = () => { $('#exName', el).textContent = name(); };
    el.addEventListener('input', upd);
    el.addEventListener('click', async (e) => {
      if (e.target.closest('[data-close]')) { closeOverlay(); return; }
      if (!e.target.closest('[data-ex]')) return;
      const scope = (el.querySelector('input[name=exScope]:checked') || {}).value || 'all';
      const deals = scope === 'filtered' ? state.filtered : state.deals;
      const ids = new Set(deals.map((d) => d.id));
      try {
        const spec = CRMReport.build({
          deals, allCount: state.deals.length, refDate: state.refDate, fileName: state.fileName,
          month: +$('#exM', el).value, year: +$('#exY', el).value,
          filterText: scope === 'filtered' ? describeFilters() : 'Toàn bộ dữ liệu',
          records: state.conv.records.filter((r) => ids.has(r['ID'])),
          staleActiveOnly: state.staleActiveOnly,
        });
        const bytes = await XLSXIO.write(spec);
        download(bytes, name());
        closeOverlay();
        toast('Đã xuất <b>' + esc(name()) + '</b>', 'ok');
      } catch (err) { console.error(err); toast('Lỗi xuất file: ' + esc(err.message), 'err'); }
    });
    showOverlay(el);
    upd();
  }
  function describeFilters() {
    const f = state.filters, parts = [];
    if (f.q) parts.push(`Tìm "${f.q}"`);
    for (const k of Object.keys(f.facets)) parts.push(`${FACETS[k].label}: ${[...f.facets[k]].map((v) => fmtFacet(k, v)).join(', ')}`);
    if (f.from || f.to) parts.push(`${DATE_FIELDS[f.dateField].split(' (')[0]}: ${f.from ? U.date(U.fromIso(f.from)) : '…'} → ${f.to ? U.date(U.fromIso(f.to)) : '…'}`);
    for (const q of f.quick) parts.push(QUICK[q].label);
    return parts.join(' | ') || 'Toàn bộ dữ liệu';
  }
  async function exportView() {
    const cols = COLS.filter((c) => visibleCols.has(c.k));
    const sorted = sortList(state.filtered, state.sort.key, state.sort.dir, (d, k) => COLMAP[k].v(d));
    const spec = CRMReport.viewSheet(sorted, cols.map((c) => c.k), state.refDate, describeFilters());
    download(await XLSXIO.write(spec), `Danh sách thương vụ CRM_${U.date(state.refDate).replace(/\//g, '.')}.xlsx`);
  }
  async function downloadTemplate() {
    const r = state.conv;
    const spec = CRMConvert.templateWorkbook({ records: r.records, nRaw: r.nRaw ?? '', nTest: r.nTest ?? '', refDate: state.refDate });
    const d = state.refDate;
    const name = `CRM_deals_template_${String(d.getUTCDate()).padStart(2, '0')}_${String(d.getUTCMonth() + 1).padStart(2, '0')}_${String(d.getUTCFullYear()).slice(2)}.xlsx`;
    download(await XLSXIO.write(spec), name);
    toast('Đã tải file chuẩn <b>' + esc(name) + '</b>', 'ok');
  }

  // ================================================================ EVENTS
  function bind() {
    const dz = $('#dropzone');
    ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); if (!state.loaded) dz.classList.add('drag'); }));
    ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, () => dz.classList.remove('drag')));
    document.addEventListener('drop', (e) => { e.preventDefault(); dz.classList.remove('drag'); const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
    $('#btnPick').addEventListener('click', () => $('#fileInput').click());
    $('#fileInput').addEventListener('change', (e) => handleFile(e.target.files[0]));
    const sess = LS.get('session', null);
    if (sess && sess.records) {
      const b = $('#btnRestore');
      b.classList.remove('hidden');
      $('span', b).textContent = `Mở lại: ${sess.fileName} (${new Date(sess.savedAt).toLocaleString('vi-VN')})`;
      b.addEventListener('click', restoreSession);
    }
    bindSearch();

    document.addEventListener('click', (e) => {
      const t = e.target;
      let a;
      if ((a = t.closest('[data-pill]'))) {
        const k = a.dataset.pill;
        if (k === 'DATE') openDatePop(a); else if (k === 'SETTINGS') openSettingsPop(a); else openFacetPop(a, k);
        return;
      }
      if ((a = t.closest('[data-act]'))) {
        e.preventDefault();
        const act = a.dataset.act;
        if (act === 'reload') $('#fileInput').click();
        else if (act === 'dlTemplate') downloadTemplate();
        else if (act === 'exportReport') openExport();
        else if (act === 'clearAll') { state.filters = newFilters(); $('#q').value = ''; $('#searchBox').classList.remove('has-q'); state.page = 1; refresh(); }
        else if (act === 'showStale') {
          state.filters.facets.stale = new Set(['1', '2', '3', '4']);
          state.tab = 'deals'; state.sort = { key: 'staleDays', dir: -1 }; state.page = 1;
          refresh();
          $('#tableCard').scrollIntoView({ behavior: 'smooth' });
        } else if (act === 'remind') openRemind();
        else if (act === 'cols') openColsPop(a);
        else if (act === 'density') { state.compact = !state.compact; LS.set('compact', state.compact); renderTable(); }
        else if (act === 'exportView') exportView();
        return;
      }
      if ((a = t.closest('[data-qs]'))) { toggleFacet('status', a.dataset.qs); return; }
      if ((a = t.closest('[data-qk]'))) {
        const q = state.filters.quick;
        q.has(a.dataset.qk) ? q.delete(a.dataset.qk) : q.add(a.dataset.qk);
        state.page = 1;
        refresh();
        return;
      }
      if ((a = t.closest('[data-fk]'))) { toggleFacet(a.dataset.fk, a.dataset.fv); return; }
      if ((a = t.closest('[data-kpi]'))) {
        const k = a.dataset.kpi;
        if (STATUSES.includes(k)) toggleFacet('status', k, true);
        else { const q = state.filters.quick; q.has(k) ? q.delete(k) : q.add(k); refresh(); }
        return;
      }
      if ((a = t.closest('[data-rm]'))) {
        const f = state.filters, r = a.dataset.rm;
        if (r === 'q') { f.q = ''; $('#q').value = ''; $('#searchBox').classList.remove('has-q'); }
        else if (r === 'f') { f.facets[a.dataset.k].delete(a.dataset.v); if (!f.facets[a.dataset.k].size) delete f.facets[a.dataset.k]; }
        else if (r === 'date') { f.from = ''; f.to = ''; }
        else if (r === 'quick') f.quick.delete(a.dataset.k);
        state.page = 1;
        refresh();
        return;
      }
      if ((a = t.closest('.seg[data-mode] button'))) {
        const mode = a.parentElement.dataset.mode;
        state.modes[mode] = a.dataset.v;
        if (mode === 'matrix') renderTable(); else renderCharts();
        return;
      }
      if ((a = t.closest('[data-tab]'))) { state.tab = a.dataset.tab; state.pivotSort = { key: 'n', dir: -1 }; renderTabs(); renderTable(); return; }
      if ((a = t.closest('th[data-sort]'))) {
        const k = a.dataset.sort;
        state.sort = state.sort.key === k ? { key: k, dir: -state.sort.dir } : { key: k, dir: ['name', 'customer', 'sale', 'reseller', 'segment', 'stage', 'status', 'ctype', 'vendors'].includes(k) ? 1 : -1 };
        renderTable();
        return;
      }
      if ((a = t.closest('th[data-psort]'))) {
        const k = a.dataset.psort;
        state.pivotSort = state.pivotSort.key === k ? { key: k, dir: -state.pivotSort.dir } : { key: k, dir: k === 'name' ? 1 : -1 };
        renderTable();
        return;
      }
      if ((a = t.closest('[data-pivot]'))) { toggleFacet(a.dataset.pivot, a.dataset.pv); return; }
      if ((a = t.closest('[data-mx]'))) {
        state.filters.facets.sale = new Set([a.dataset.mx]);
        state.filters.facets.stage = new Set([a.dataset.ms]);
        state.tab = 'deals'; state.page = 1;
        refresh();
        return;
      }
      if ((a = t.closest('[data-page]'))) { if (!a.disabled) { state.page = +a.dataset.page; renderTable(); $('#tableCard').scrollIntoView({ block: 'start' }); } return; }
      if ((a = t.closest('[data-open]'))) { e.preventDefault(); openDrawer(a.dataset.open); }
    });
    document.addEventListener('change', (e) => {
      if (e.target.id === 'pageSize') { state.pageSize = +e.target.value; LS.set('pageSize', state.pageSize); state.page = 1; renderTable(); }
    });
  }

  bind();
})();
