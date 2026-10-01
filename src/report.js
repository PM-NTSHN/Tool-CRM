/* report.js – dựng file Excel "Tổng hợp CRM_ Tháng M.YYYY.xlsx" (nhiều sheet thống kê) và file "bảng đang xem". */
(function (root) {
  'use strict';
  const { CRMModel, CRMConvert } = root;
  const { U, NONE, STATUSES, STAGES, BANDS, BAND_LABEL, STALE, STALE_LABEL, CLOSE, CLOSE_LABEL } = CRMModel;

  // ---------------------------------------------------------------- styles
  const F = 'Arial';
  const NAVY = '073F7A', BLUE = '0959A2', ORANGE = 'F06723', LINE = 'D5DEEA';
  const S = {
    title: { font: { name: F, sz: 16, b: true, color: 'FFFFFF' }, fill: NAVY, align: { v: 'center', indent: 1 } },
    subtitle: { font: { name: F, sz: 10, i: true, color: '42546B' }, align: { indent: 1 } },
    section: { font: { name: F, sz: 12, b: true, color: NAVY }, border: { color: ORANGE, style: 'medium', sides: ['bottom'] } },
    th: { font: { name: F, sz: 10, b: true, color: 'FFFFFF' }, fill: BLUE, border: LINE, align: { h: 'center', v: 'center', wrap: true } },
    td: { font: { name: F, sz: 10 }, border: LINE, align: { v: 'top' } },
    tdWrap: { font: { name: F, sz: 10 }, border: LINE, align: { v: 'top', wrap: true } },
    tdB: { font: { name: F, sz: 10, b: true }, border: LINE, align: { v: 'top' } },
    num: { font: { name: F, sz: 10 }, border: LINE, numFmt: '#,##0', align: { v: 'top' } },
    pct: { font: { name: F, sz: 10 }, border: LINE, numFmt: '0.0%', align: { v: 'top' } },
    date: { font: { name: F, sz: 10 }, border: LINE, numFmt: 'dd/mm/yyyy', align: { v: 'top' } },
    dateTime: { font: { name: F, sz: 10 }, border: LINE, numFmt: 'dd/mm/yyyy hh:mm', align: { v: 'top' } },
    totLabel: { font: { name: F, sz: 10, b: true, color: NAVY }, fill: 'E8F0FA', border: LINE },
    totNum: { font: { name: F, sz: 10, b: true, color: NAVY }, fill: 'E8F0FA', border: LINE, numFmt: '#,##0' },
    totPct: { font: { name: F, sz: 10, b: true, color: NAVY }, fill: 'E8F0FA', border: LINE, numFmt: '0.0%' },
    kpiLabel: { font: { name: F, sz: 10, b: true, color: '42546B' }, fill: 'F3F7FC', border: LINE },
    kpiVal: { font: { name: F, sz: 11, b: true, color: NAVY }, border: LINE, numFmt: '#,##0' },
    kpiPct: { font: { name: F, sz: 11, b: true, color: NAVY }, border: LINE, numFmt: '0.0%' },
    note: { font: { name: F, sz: 9, i: true, color: '6B7C92' } },
  };
  const ZEBRA = 'F6F9FD';
  const STATUS_FONT = { Active: '1C62B6', Won: '0F7F57', Lost: 'BE2D2C' };
  const STALE_FILL = { 1: 'FFF4D6', 2: 'FFE9DC', 3: 'FDE3E3', 4: '8E1B3E' };
  const with_ = (base, extra) => Object.assign({}, base, extra, extra && extra.font ? { font: Object.assign({}, base.font, extra.font) } : {});
  const c = (v, s) => ({ v: v === undefined ? null : v, s });

  const quarterKey = (d) => `${d.getUTCFullYear()}-Q${Math.floor(d.getUTCMonth() / 3) + 1}`;
  const staleText = (d) => (d.stale === 'na' ? 'Không xét' : d.stale === '0' ? 'Bình thường' : 'Mức ' + d.stale);
  const closeText = (d) => CLOSE_LABEL[d.closeState];

  // ---------------------------------------------------------------- bảng thống kê dùng chung
  function statTable(rows, title, list, keyFn, keys, labelFn) {
    const total = list.length, totalV = list.reduce((a, d) => a + (d.value || 0), 0);
    rows.push([]);
    rows.push([c(title, S.section), c(null, S.section), c(null, S.section), c(null, S.section), c(null, S.section)]);
    rows.push(['Nhóm', 'Số deal', '% số deal', 'Tổng giá trị (VND)', '% giá trị'].map((h) => c(h, S.th)));
    const g = CRMModel.groupBy(list, keyFn);
    const ks = keys || [...g.keys()].sort((a, b) => g.get(b).length - g.get(a).length);
    ks.forEach((k, i) => {
      const arr = g.get(k) || [];
      const v = arr.reduce((a, d) => a + (d.value || 0), 0);
      const z = i % 2 ? { fill: ZEBRA } : null;
      rows.push([c(labelFn ? labelFn(k) : k === NONE ? '(Trống)' : k, with_(S.td, z)), c(arr.length, with_(S.num, z)), c(total ? arr.length / total : 0, with_(S.pct, z)),
        c(v, with_(S.num, z)), c(totalV ? v / totalV : 0, with_(S.pct, z))]);
    });
    rows.push([c('Tổng', S.totLabel), c(total, S.totNum), c(total ? 1 : 0, S.totPct), c(totalV, S.totNum), c(totalV ? 1 : 0, S.totPct)]);
  }

  // ---------------------------------------------------------------- pivot
  const PIVOT_HEAD = ['Số deal', 'Active', 'Won', 'Lost', 'Tổng giá trị (VND)', 'Pipeline Active (VND)', 'Giá trị Won (VND)', 'Tỷ lệ thắng',
    'Giá trị TB / deal (VND)', 'Cần nhắc update', 'Cờ ≥ Mức 2', 'Quá timeline', 'Close ≤ 30 ngày', 'Vấn đề dữ liệu'];
  function pivotSheet(name, list, keyFn, label, opts) {
    opts = opts || {};
    const rows = [[c(name.toUpperCase() + (opts.subtitle ? ' – ' + opts.subtitle : ''), S.title)], [c(opts.note || '', S.subtitle)], []];
    rows.push([label, ...PIVOT_HEAD].map((h) => c(h, S.th)));
    const g = CRMModel.groupBy(list, keyFn);
    const data = [...g.entries()].map(([k, arr]) => {
      const s = CRMModel.summarize(arr);
      let vs = s;
      if (opts.mainOnly) vs = CRMModel.summarize(arr.filter((d) => d.mainVendor === k));
      return { k, arr, s, value: vs.value, valueActive: vs.valueActive, valueWon: vs.valueWon, avg: s.withValue ? vs.value / s.withValue : null };
    }).sort((a, b) => b.s.n - a.s.n || b.value - a.value);
    data.forEach((r, i) => {
      const z = i % 2 ? { fill: ZEBRA } : null;
      const s = r.s;
      rows.push([
        c(opts.fmt ? opts.fmt(r.k) : r.k === NONE ? '(Trống)' : r.k, with_(S.tdB, z)),
        ...[s.n, s.active, s.won, s.lost, r.value, r.valueActive, r.valueWon].map((v) => c(v, with_(S.num, z))),
        c(s.winRate, with_(S.pct, z)), c(r.avg, with_(S.num, z)),
        ...[s.remind, s.stale2, s.overdue, s.soon30, s.issues].map((v) => c(v, with_(S.num, z))),
      ]);
    });
    const t = CRMModel.summarize(list);
    rows.push([c(`Tổng (${data.length})`, S.totLabel), ...[t.n, t.active, t.won, t.lost, t.value, t.valueActive, t.valueWon].map((v) => c(v, S.totNum)),
      c(t.winRate, S.totPct), c(t.avg, S.totNum), ...[t.remind, t.stale2, t.overdue, t.soon30, t.issues].map((v) => c(v, S.totNum))]);
    return { name, rows, cols: [opts.width || 30, 9, 8, 8, 8, 18, 18, 18, 10, 17, 10, 10, 10, 10, 10], freeze: { row: 4, col: 1 },
      autoFilter: `A4:O${rows.length - 1}`, merges: ['A1:O1', 'A2:O2'], rowHeights: { 0: 30, 3: 32 }, showGrid: false };
  }

  // ---------------------------------------------------------------- danh sách deal (đầy đủ)
  const LIST_COLS = [
    ['STT', 6, (d, i) => i + 1, 'num'],
    ['ID', 10, (d) => d.id, 'td'],
    ['Tên thương vụ', 44, (d) => d.name, 'tdWrap'],
    ['Khách hàng', 34, (d) => d.customer, 'tdWrap'],
    ['Phân nhóm KH', 24, (d) => d.segment, 'td'],
    ['Sale', 20, (d) => d.sale, 'td'],
    ['Reseller', 18, (d) => d.reseller, 'td'],
    ['Giai đoạn', 20, (d) => d.stage, 'td'],
    ['Trạng thái', 11, (d) => d.status, 'status'],
    ['Giá trị (VND)', 17, (d) => d.value, 'num'],
    ['Khoảng giá trị', 14, (d) => BAND_LABEL[d.band], 'td'],
    ['Deal size (CRM)', 11, (d) => d.sizeCRM, 'td'],
    ['Hãng chính', 15, (d) => d.mainVendor, 'td'],
    ['Các hãng', 26, (d) => d.vendors.join(', '), 'td'],
    ['Số hãng', 8, (d) => d.vendors.length || null, 'num'],
    ['Phân loại HĐ', 12, (d) => d.ctype, 'td'],
    ['Close date', 12, (d) => d.close, 'date'],
    ['Tình trạng timeline', 18, (d) => closeText(d), 'close'],
    ['Số ngày đến close', 11, (d) => d.closeDays, 'num'],
    ['Quý dự kiến close', 11, (d) => d.closeQuarter, 'td'],
    ['Ngày tạo', 16, (d) => d.created, 'dateTime'],
    ['Tuổi deal (ngày)', 10, (d) => d.ageDays, 'num'],
    ['Cập nhật lần cuối', 16, (d) => d.updated, 'dateTime'],
    ['Số ngày chưa cập nhật', 11, (d) => d.staleDays, 'num'],
    ['Cờ nhắc', 11, (d) => staleText(d), 'stale'],
    ['Lý do Failed', 34, (d) => d.failed, 'tdWrap'],
    ['Thông tin BOM', 40, (d) => d.bom, 'tdWrap'],
    ['Số vấn đề dữ liệu', 10, (d) => d.issues.length || null, 'num'],
    ['Vấn đề dữ liệu', 60, (d) => d.issues.join('; ') || null, 'tdWrap'],
    ['Ghi chú xử lý (convert)', 60, (d) => d.handled.join('; ') || null, 'tdWrap'],
  ];
  function cellFor(kind, v, d, z) {
    if (kind === 'status') return c(v, with_(S.tdB, Object.assign({}, z, { font: { color: STATUS_FONT[v] || '0F1F33' } })));
    if (kind === 'stale') {
      if (d.stale === 'na' || d.stale === '0') return c(v, with_(S.td, z));
      return c(v, with_(S.tdB, { fill: STALE_FILL[d.stale], font: d.stale === '4' ? { color: 'FFFFFF' } : {} }));
    }
    if (kind === 'close') return c(v, with_(S.td, Object.assign({}, z, d.overdue ? { font: { color: 'BE2D2C', b: true } } : {})));
    return c(v, with_(S[kind], z));
  }
  function listSheet(name, deals, colDefs, title) {
    const rows = [[c(title, S.title)], [], colDefs.map((x) => c(x[0], S.th))];
    deals.forEach((d, i) => {
      const z = i % 2 ? { fill: ZEBRA } : null;
      rows.push(colDefs.map(([, , fn, kind]) => cellFor(kind, fn(d, i), d, z)));
    });
    const colL = root.XLSXIO.colName(colDefs.length - 1);
    return { name, rows, cols: colDefs.map((x) => x[1]), freeze: { row: 3, col: 3 }, autoFilter: `A3:${colL}${rows.length}`,
      merges: [`A1:${root.XLSXIO.colName(Math.min(colDefs.length, 12) - 1)}1`], rowHeights: { 0: 28, 2: 32 } };
  }

  // ---------------------------------------------------------------- build
  function build(o) {
    const { deals, refDate, month, year } = o;
    const s = CRMModel.summarize(deals);
    const sheets = [];

    // ===== Tổng quan
    const ov = [];
    ov.push([c(`BÁO CÁO TỔNG HỢP CRM – THÁNG ${month}/${year}`, S.title)]);
    ov.push([c('Công ty Nam Trường Sơn Hà Nội (NTS Hanoi Corp.) – Phòng Quản lý sản phẩm', S.subtitle)]);
    ov.push([c(`Ngày chốt số liệu: ${U.date(refDate)}   ·   Nguồn: ${o.fileName}   ·   Xuất lúc: ${new Date().toLocaleString('vi-VN')}`, S.subtitle)]);
    ov.push([c(`Phạm vi: ${o.filterText}   ·   ${U.int(deals.length)} / ${U.int(o.allCount)} thương vụ`, S.subtitle)]);
    ov.push([]);
    ov.push([c('CHỈ SỐ CHÍNH', S.section), c(null, S.section), c(null, S.section), c(null, S.section), c(null, S.section)]);
    const kpi = (l, v, st, note) => ov.push([c(l, S.kpiLabel), c(v, st || S.kpiVal), c(note || null, S.note)]);
    kpi('Tổng số thương vụ', s.n, null, `${new Set(deals.map((d) => d.sale).filter(Boolean)).size} sale · ${new Set(deals.flatMap((d) => d.vendors)).size} hãng · ${new Set(deals.map((d) => d.customer)).size} khách hàng`);
    kpi('Tổng giá trị (VND)', s.value, null, `${U.int(s.withValue)} deal có giá trị · ${U.money(s.value)}`);
    kpi('Pipeline Active (VND)', s.valueActive, null, `${U.int(s.active)} deal Active · ${U.money(s.valueActive)}`);
    kpi('Số deal Won', s.won, null, `Giá trị Won: ${U.money(s.valueWon)}`);
    kpi('Số deal Lost', s.lost, null, `Giá trị Lost: ${U.money(s.valueLost)}`);
    kpi('Tỷ lệ thắng (Won / (Won + Lost))', s.winRate, S.kpiPct, `${U.int(s.won + s.lost)} deal đã có kết quả`);
    kpi('Giá trị trung bình / deal (VND)', s.avg, null, `Trung vị: ${U.money(s.median)}`);
    kpi('Cần nhắc cập nhật (> 1 tháng)', s.remind, null, o.staleActiveOnly ? 'Chỉ xét deal Active' : 'Xét mọi trạng thái');
    kpi('Cờ nhắc từ Mức 2 (> 3 tháng)', s.stale2);
    kpi('Quá timeline – vẫn Active', s.overdue, null, 'Close date đã qua so với ngày chốt');
    kpi('Close trong 30 ngày tới (Active)', s.soon30);
    kpi('Deal có vấn đề dữ liệu', s.issues, null, 'Có cờ [VẤN ĐỀ] – xem sheet "Vấn đề dữ liệu"');
    statTable(ov, 'THEO TRẠNG THÁI', deals, (d) => d.status, STATUSES.filter((k) => deals.some((d) => d.status === k)));
    statTable(ov, 'THEO GIAI ĐOẠN', deals, (d) => d.stage, [...new Set(deals.map((d) => d.stage || NONE))].sort((a, b) => (STAGES.indexOf(a) + 1 || 99) - (STAGES.indexOf(b) + 1 || 99)), (k) => (k === NONE ? '(Chưa có)' : k));
    statTable(ov, 'THEO KHOẢNG GIÁ TRỊ', deals, (d) => d.band, [...BANDS.map((b) => b.k), NONE], (k) => BAND_LABEL[k]);
    statTable(ov, 'THEO PHÂN LOẠI HỢP ĐỒNG', deals, (d) => d.ctype, null, (k) => (k === NONE ? '(Chưa phân loại)' : k));
    statTable(ov, 'THEO PHÂN NHÓM KHÁCH HÀNG', deals, (d) => d.segment);
    statTable(ov, 'CỜ NHẮC CẬP NHẬT', deals, (d) => d.stale, STALE.map((x) => x.k).filter((k) => deals.some((d) => d.stale === k)), (k) => STALE_LABEL[k]);
    statTable(ov, 'TIMELINE DỰ ÁN (deal Active)', deals.filter((d) => d.status === 'Active'), (d) => d.closeState, CLOSE.map((x) => x.k), (k) => CLOSE_LABEL[k]);
    sheets.push({ name: 'Tổng quan', rows: ov, cols: [38, 20, 14, 22, 12], merges: ['A1:E1', 'A2:E2', 'A3:E3', 'A4:E4'], rowHeights: { 0: 32 }, showGrid: false, tabColor: ORANGE });

    // ===== Danh sách thương vụ
    const sorted = [...deals].sort((a, b) => (STATUSES.indexOf(a.status) - STATUSES.indexOf(b.status)) || ((b.value || 0) - (a.value || 0)));
    sheets.push(listSheet('Danh sách thương vụ', sorted, LIST_COLS, `DANH SÁCH THƯƠNG VỤ – THÁNG ${month}/${year} (${U.int(deals.length)} thương vụ · chốt ${U.date(refDate)})`));

    // ===== Pivot
    sheets.push(pivotSheet('Theo Sale', deals, (d) => d.sale, 'Sale', { note: 'Thống kê theo sale chủ sở hữu thương vụ.' }));
    sheets.push(pivotSheet('Theo Hãng', deals, (d) => d.vendors, 'Hãng', { mainOnly: true, fmt: (k) => (k === NONE ? '(Chưa xác định)' : k),
      note: 'Số deal = mọi deal có hãng này (1 deal nhiều hãng được đếm ở mỗi hãng). Giá trị chỉ tính cho deal có hãng này là hãng chính – tránh cộng trùng.' }));
    sheets.push(pivotSheet('Theo Reseller', deals, (d) => d.reseller, 'Reseller', { fmt: (k) => (k === NONE ? '(Chưa có reseller)' : k), note: 'Thống kê theo reseller (đại lý/đối tác) trên CRM.' }));
    sheets.push(pivotSheet('Theo Nhóm KH', deals, (d) => d.segment, 'Phân nhóm khách hàng', { note: 'Phân nhóm khách hàng tự động khi convert.' }));
    sheets.push(pivotSheet('Theo Khách hàng', deals, (d) => d.customer, 'Khách hàng', { width: 44, note: 'Khách hàng đã chuẩn hoá tên khi convert.' }));

    // ===== Sale x Giai đoạn
    const stages = [...new Set(deals.map((d) => d.stage || NONE))].sort((a, b) => (STAGES.indexOf(a) + 1 || 99) - (STAGES.indexOf(b) + 1 || 99));
    const bySale = CRMModel.groupBy(deals, (d) => d.sale);
    const saleKeys = [...bySale.keys()].sort((a, b) => bySale.get(b).length - bySale.get(a).length);
    const mx = [[c('SALE × GIAI ĐOẠN', S.title)], []];
    const block = (title, fn) => {
      mx.push([c(title, S.section), ...stages.map(() => c(null, S.section)), c(null, S.section)]);
      mx.push([c('Sale', S.th), ...stages.map((st) => c(st === NONE ? '(Chưa có)' : st, S.th)), c('Tổng', S.th)]);
      saleKeys.forEach((sk, i) => {
        const z = i % 2 ? { fill: ZEBRA } : null;
        const arr = bySale.get(sk);
        mx.push([c(sk === NONE ? '(Chưa có sale)' : sk, with_(S.tdB, z)), ...stages.map((st) => c(fn(arr.filter((d) => (d.stage || NONE) === st)) || null, with_(S.num, z))), c(fn(arr), with_(S.totNum, {}))]);
      });
      mx.push([c('Tổng', S.totLabel), ...stages.map((st) => c(fn(deals.filter((d) => (d.stage || NONE) === st)), S.totNum)), c(fn(deals), S.totNum)]);
      mx.push([]);
    };
    block('SỐ THƯƠNG VỤ', (a) => a.length);
    block('GIÁ TRỊ (VND)', (a) => a.reduce((x, d) => x + (d.value || 0), 0));
    block('SỐ THƯƠNG VỤ ACTIVE', (a) => a.filter((d) => d.status === 'Active').length);
    sheets.push({ name: 'Sale x Giai đoạn', rows: mx, cols: [26, ...stages.map(() => 18), 18], merges: [`A1:${root.XLSXIO.colName(stages.length + 1)}1`], rowHeights: { 0: 30 }, showGrid: false, freeze: { col: 1 } });

    // ===== Forecast theo quý
    const act = deals.filter((d) => d.status === 'Active');
    const byQ = CRMModel.groupBy(act, (d) => (d.close ? quarterKey(d.close) : NONE));
    const qKeys = [...byQ.keys()].sort((a, b) => (a === NONE) - (b === NONE) || (a < b ? -1 : 1));
    const fc = [[c('FORECAST PIPELINE ACTIVE THEO QUÝ DỰ KIẾN CLOSE', S.title)], [c('Dựa trên Timeline dự án (close date) của các thương vụ đang Active.', S.subtitle)], []];
    const fcHead = ['Quý dự kiến close', 'Số deal', 'Tổng giá trị (VND)', ...STAGES.map((x) => 'GT ' + x), 'Đã quá timeline'];
    fc.push(fcHead.map((h) => c(h, S.th)));
    qKeys.forEach((q, i) => {
      const arr = byQ.get(q), z = i % 2 ? { fill: ZEBRA } : null;
      const sum = (a) => a.reduce((x, d) => x + (d.value || 0), 0);
      fc.push([c(q === NONE ? '(Chưa có timeline)' : q.replace('-', ' '), with_(S.tdB, z)), c(arr.length, with_(S.num, z)), c(sum(arr), with_(S.num, z)),
        ...STAGES.map((st) => c(sum(arr.filter((d) => d.stage === st)) || null, with_(S.num, z))), c(arr.filter((d) => d.overdue).length || null, with_(S.num, z))]);
    });
    fc.push([c('Tổng', S.totLabel), c(act.length, S.totNum), c(act.reduce((x, d) => x + (d.value || 0), 0), S.totNum),
      ...STAGES.map((st) => c(act.filter((d) => d.stage === st).reduce((x, d) => x + (d.value || 0), 0), S.totNum)), c(act.filter((d) => d.overdue).length, S.totNum)]);
    sheets.push({ name: 'Forecast theo quý', rows: fc, cols: [22, 9, 20, 18, 18, 20, 18, 20, 12], merges: ['A1:I1', 'A2:I2'], rowHeights: { 0: 30, 3: 32 }, showGrid: false });

    // ===== Cần nhắc cập nhật
    const remind = deals.filter((d) => d.needRemind).sort((a, b) => b.staleDays - a.staleDays);
    const RCOLS = [LIST_COLS[0], LIST_COLS[1], LIST_COLS[2], LIST_COLS[3], LIST_COLS[5], LIST_COLS[7], LIST_COLS[8], LIST_COLS[9], LIST_COLS[16], LIST_COLS[17], LIST_COLS[22], LIST_COLS[23], LIST_COLS[24]];
    sheets.push(listSheet('Cần nhắc cập nhật', remind, RCOLS, `THƯƠNG VỤ CẦN NHẮC CẬP NHẬT (${U.int(remind.length)}) – tính đến ${U.date(refDate)} · Mức 1 > 1 tháng · Mức 2 > 3 tháng · Mức 3 > 6 tháng · Mức 4 > 1 năm`));

    // ===== Vấn đề dữ liệu
    const iss = deals.filter((d) => d.issues.length).sort((a, b) => b.issues.length - a.issues.length);
    const ICOLS = [LIST_COLS[0], LIST_COLS[1], LIST_COLS[2], LIST_COLS[3], LIST_COLS[5], LIST_COLS[8], LIST_COLS[9], LIST_COLS[27], LIST_COLS[28]];
    sheets.push(listSheet('Vấn đề dữ liệu', iss, ICOLS, `THƯƠNG VỤ CÓ VẤN ĐỀ DỮ LIỆU CẦN RÀ SOÁT (${U.int(iss.length)})`));

    // ===== Data chuẩn
    const tpl = CRMConvert.templateWorkbook({ records: o.records, nRaw: '', nTest: '', refDate }).sheets[0];
    tpl.name = 'Data chuẩn';
    sheets.push(tpl);

    // ===== Ghi chú
    const notes = [
      ['Mục', 'Giải thích'],
      ['Đơn vị đếm', 'Mỗi thương vụ (ID CRM duy nhất) được đếm 1 lần. Sheet "Data chuẩn" giữ cấu trúc template: deal nhiều hãng tách nhiều dòng cùng ID.'],
      ['Giá trị', 'VND, lấy từ file template chuẩn (đã quy đổi / lấy từ BOM khi CRM = 0). Ô trống = chưa có giá trị.'],
      ['Khoảng giá trị', 'Dưới 1 tỷ · 1–2 tỷ · 2–5 tỷ · 5–10 tỷ · Trên 10 tỷ (cận dưới tính vào khoảng trên, VD 2 tỷ thuộc 2–5 tỷ).'],
      ['Pipeline Active', 'Tổng giá trị các thương vụ đang Active.'],
      ['Tỷ lệ thắng', 'Won / (Won + Lost), không tính deal Active.'],
      ['Cờ nhắc cập nhật', `Tính từ "Cập nhật lần cuối" đến ngày chốt ${U.date(refDate)}: Mức 1 > 1 tháng, Mức 2 > 3 tháng, Mức 3 > 6 tháng, Mức 4 > 1 năm. ${o.staleActiveOnly ? 'Chỉ xét deal Active (Won/Lost = "Không xét").' : 'Xét mọi trạng thái.'}`],
      ['Tình trạng timeline', 'So sánh Close date (Timeline dự án) với ngày chốt: Đã qua timeline / Trong 30 ngày tới / 31–90 ngày / Sau 90 ngày / Chưa có.'],
      ['Quá timeline', 'Deal Active có Close date trước ngày chốt.'],
      ['Tuổi deal', 'Số ngày từ Ngày tạo đến ngày chốt.'],
      ['Hãng chính', 'Hãng có giá trị BOM lớn nhất (theo quy tắc convert). "Theo Hãng": số deal đếm mọi hãng của deal, giá trị chỉ tính ở hãng chính.'],
      ['Vấn đề dữ liệu', 'Các mục [VẤN ĐỀ] trong cột "Cờ chất lượng" của file template – dữ liệu thiếu / mâu thuẫn cần sale cập nhật.'],
      ['Nguồn', `${o.fileName} · ngày chốt ${U.date(refDate)} · Phạm vi: ${o.filterText}`],
    ];
    sheets.push({ name: 'Ghi chú', rows: [[c('GHI CHÚ – ĐỊNH NGHĨA CHỈ SỐ', S.title)], [], ...notes.map((r, i) => (i === 0 ? r.map((x) => c(x, S.th)) : [c(r[0], S.tdB), c(r[1], S.tdWrap)]))],
      cols: [24, 120], merges: ['A1:B1'], rowHeights: { 0: 30 }, showGrid: false });

    return { title: `Tổng hợp CRM tháng ${month}/${year}`, creator: 'NTS Hanoi – Phòng Quản lý sản phẩm', sheets };
  }

  // ---------------------------------------------------------------- xuất bảng đang xem
  const VIEW_MAP = {
    id: 1, name: 2, customer: 3, segment: 4, sale: 5, reseller: 6, stage: 7, status: 8, value: 9, band: 10, sizeCRM: 11,
    vendors: 13, ctype: 15, bom: 26, close: 16, closeState: 17, closeQuarter: 19, created: 20, ageDays: 21, updated: 22, staleDays: 23,
    stale: 24, failed: 25, issues: 28, quality: null,
  };
  function viewSheet(deals, keys, refDate, filterText) {
    const cols = [LIST_COLS[0]];
    keys.forEach((k) => {
      if (k === 'quality') cols.push(['Cờ chất lượng', 80, (d) => d.quality, 'tdWrap']);
      else if (VIEW_MAP[k] != null) cols.push(LIST_COLS[VIEW_MAP[k]]);
    });
    const sh = listSheet('Danh sách', deals, cols, `DANH SÁCH THƯƠNG VỤ CRM – chốt ${U.date(refDate)} · ${U.int(deals.length)} thương vụ · ${filterText}`);
    return { title: 'Danh sách thương vụ CRM', sheets: [sh] };
  }

  root.CRMReport = { build, viewSheet };
})(typeof globalThis !== 'undefined' ? globalThis : this);
