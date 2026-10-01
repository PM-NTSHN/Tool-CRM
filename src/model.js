/* model.js – mô hình dữ liệu phân tích (1 thương vụ = 1 deal theo ID duy nhất) + các tiện ích định dạng.
 * Dùng chung cho dashboard (app.js) và báo cáo Excel (report.js). */
(function (root) {
  'use strict';
  const { fold } = root.CRMConvert;
  const DAY = 86400000;

  // ---------------------------------------------------------------- định dạng
  const nf0 = new Intl.NumberFormat('vi-VN', { maximumFractionDigits: 0 });
  const pad2 = (n) => String(n).padStart(2, '0');
  const U = {
    DAY,
    fold,
    esc(s) {
      return String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    },
    int: (n) => nf0.format(n || 0),
    vnd: (v) => (v == null ? '' : nf0.format(Math.round(v))),
    money(v, unit) {
      if (v == null || isNaN(v)) return '–';
      const a = Math.abs(v);
      if (unit === 'ty' || a >= 1e9) {
        const x = v / 1e9;
        const d = Math.abs(x) >= 100 ? 0 : Math.abs(x) >= 10 ? 1 : 2;
        return x.toLocaleString('vi-VN', { maximumFractionDigits: d }) + ' tỷ';
      }
      if (a >= 1e6) return Math.round(v / 1e6).toLocaleString('vi-VN') + ' tr';
      return nf0.format(Math.round(v)) + ' đ';
    },
    pct(x, d = 1) {
      if (x == null || !isFinite(x)) return '–';
      return (x * 100).toLocaleString('vi-VN', { maximumFractionDigits: d, minimumFractionDigits: d }) + '%';
    },
    date(d) {
      if (!(d instanceof Date) || isNaN(d)) return '';
      return `${pad2(d.getUTCDate())}/${pad2(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;
    },
    dateTime(d) {
      if (!(d instanceof Date) || isNaN(d)) return '';
      return U.date(d) + ` ${pad2(d.getUTCHours())}:${pad2(d.getUTCMinutes())}`;
    },
    iso(d) {
      if (!(d instanceof Date) || isNaN(d)) return '';
      return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
    },
    fromIso(s) {
      const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
      return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
    },
    today() {
      const n = new Date();
      return new Date(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate()));
    },
    addMonths(d, n) {
      const y = d.getUTCFullYear(), m = d.getUTCMonth() + n, day = d.getUTCDate();
      const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
      return new Date(Date.UTC(y, m, Math.min(day, last), d.getUTCHours(), d.getUTCMinutes()));
    },
    startOfDay: (d) => new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())),
    monthKey: (d) => `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}`,
    monthLabel(k) { const [y, m] = k.split('-'); return `T${+m}/${y}`; },
    median(arr) {
      if (!arr.length) return null;
      const s = [...arr].sort((a, b) => a - b), h = s.length >> 1;
      return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2;
    },
  };

  // ---------------------------------------------------------------- danh mục
  const NONE = '∅';
  const STATUSES = ['Active', 'Won', 'Lost'];
  const STATUS_COLORS = { Active: '#2a78d6', Won: '#17a673', Lost: '#e34948' };
  const STAGES = ['Đăng kí cơ hội', 'Tư vấn', 'Phương án kinh doanh', 'Xây dựng hồ sơ', 'Thực hiện hợp đồng'];
  const BANDS = [
    { k: 'lt1', label: 'Dưới 1 tỷ', short: '< 1 tỷ', test: (v) => v < 1e9 },
    { k: '1-2', label: 'Từ 1 – 2 tỷ', short: '1 – 2 tỷ', test: (v) => v < 2e9 },
    { k: '2-5', label: 'Từ 2 – 5 tỷ', short: '2 – 5 tỷ', test: (v) => v < 5e9 },
    { k: '5-10', label: 'Từ 5 – 10 tỷ', short: '5 – 10 tỷ', test: (v) => v < 1e10 },
    { k: 'gt10', label: 'Trên 10 tỷ', short: '> 10 tỷ', test: () => true },
  ];
  const BAND_LABEL = Object.fromEntries(BANDS.map((b) => [b.k, b.label]));
  BAND_LABEL[NONE] = 'Chưa có giá trị';
  const STALE = [
    { k: '0', label: 'Bình thường (≤ 1 tháng)', short: 'Bình thường', color: '#17a673' },
    { k: '1', label: 'Mức 1 · > 1 tháng', short: 'Mức 1', color: '#eda100' },
    { k: '2', label: 'Mức 2 · > 3 tháng', short: 'Mức 2', color: '#eb6834' },
    { k: '3', label: 'Mức 3 · > 6 tháng', short: 'Mức 3', color: '#e34948' },
    { k: '4', label: 'Mức 4 · > 1 năm', short: 'Mức 4', color: '#8e1b3e' },
    { k: 'na', label: 'Không xét (đã Won/Lost)', short: 'Không xét', color: '#9aa8b8' },
  ];
  const STALE_LABEL = Object.fromEntries(STALE.map((s) => [s.k, s.label]));
  const CLOSE = [
    { k: 'overdue', label: 'Đã qua timeline', color: '#e34948' },
    { k: 'd30', label: 'Trong 30 ngày tới', color: '#eb6834' },
    { k: 'd90', label: '31 – 90 ngày tới', color: '#2a78d6' },
    { k: 'later', label: 'Sau 90 ngày', color: '#17a673' },
    { k: 'none', label: 'Chưa có timeline', color: '#9aa8b8' },
  ];
  const CLOSE_LABEL = Object.fromEntries(CLOSE.map((s) => [s.k, s.label]));

  function bandOf(v) {
    if (v == null || !(v > 0)) return NONE;
    return BANDS.find((b) => b.test(v)).k;
  }

  // ---------------------------------------------------------------- dựng deal
  function splitQuality(q) {
    const issues = [], handled = [];
    if (typeof q === 'string') {
      for (const part of q.split(' | ')) {
        if (part.startsWith('[VẤN ĐỀ] ')) issues.push(...part.slice(9).split('; ').filter(Boolean));
        else if (part.startsWith('[ĐÃ XỬ LÝ] ')) handled.push(...part.slice(11).split('; ').filter(Boolean));
      }
    }
    return { issues, handled };
  }

  /** records: các dòng template (1 deal có thể nhiều dòng do tách hãng) -> danh sách deal. */
  function buildDeals(records) {
    const byId = new Map();
    for (const r of records) {
      const id = r['ID'];
      if (id == null) continue;
      if (!byId.has(id)) byId.set(id, []);
      byId.get(id).push(r);
    }
    const deals = [];
    for (const [id, rows] of byId) {
      const f = rows[0];
      const vendors = [];
      rows.forEach((r) => { if (r['Hãng'] && !vendors.includes(r['Hãng'])) vendors.push(r['Hãng']); });
      const value = rows.map((r) => r['Giá trị thương vụ']).find((v) => typeof v === 'number' && !isNaN(v));
      const { issues, handled } = splitQuality(f['Cờ chất lượng']);
      const d = {
        id,
        name: f['Tên thương vụ'] || '',
        customer: f['Tên khách hàng'] || '',
        segment: f['Phân nhóm khách hàng'] || 'Khác',
        sale: f['Sale'] || null,
        reseller: f['Reseller'] || null,
        stage: f['Giai đoạn'] || null,
        status: f['Trạng thái'] || null,
        value: value == null ? null : value,
        vendors,
        mainVendor: vendors[0] || null,
        bom: f['Thông tin BOM (model, số lượng, năm support...)'] || null,
        ctype: f['Phân loại HĐ'] || null,
        sizeCRM: f['Phân loại deal size'] || null,
        close: f['Close date'] || null,
        failed: f['Lý do Failed'] || null,
        created: f['Ngày tạo'] || null,
        updated: f['Cập nhật lần cuối'] || null,
        quality: f['Cờ chất lượng'] || null,
        issues, handled,
        rows: rows.length,
      };
      d.band = bandOf(d.value);
      d._s = fold([id, d.name, d.customer, d.segment, d.sale, d.reseller, d.stage, d.status, vendors.join(' '),
        d.ctype, d.bom, d.failed].filter(Boolean).join(' | '));
      deals.push(d);
    }
    return deals;
  }

  /** Các trường phụ thuộc ngày chốt số liệu & thiết lập cờ nhắc. */
  function derive(deals, ref, opts) {
    const activeOnly = !opts || opts.staleActiveOnly !== false;
    const ms = [1, 3, 6, 12];
    for (const d of deals) {
      if (d.updated) {
        d.staleDays = Math.max(0, Math.floor((ref - U.startOfDay(d.updated)) / DAY));
        let lv = 0;
        ms.forEach((m, i) => { if (U.addMonths(d.updated, m) < ref) lv = i + 1; });
        d.staleLevel = lv;
      } else {
        d.staleDays = null;
        d.staleLevel = 4;
      }
      d.stale = activeOnly && d.status !== 'Active' ? 'na' : String(d.staleLevel);
      if (d.close) {
        d.closeDays = Math.round((U.startOfDay(d.close) - ref) / DAY);
        d.closeState = d.closeDays < 0 ? 'overdue' : d.closeDays <= 30 ? 'd30' : d.closeDays <= 90 ? 'd90' : 'later';
      } else {
        d.closeDays = null;
        d.closeState = 'none';
      }
      d.ageDays = d.created ? Math.max(0, Math.floor((ref - U.startOfDay(d.created)) / DAY)) : null;
      d.closeQuarter = d.close ? `Q${Math.floor(d.close.getUTCMonth() / 3) + 1}/${d.close.getUTCFullYear()}` : null;
      d.overdue = d.status === 'Active' && d.closeState === 'overdue';
      d.needRemind = d.stale !== 'na' && d.stale !== '0';
    }
    return deals;
  }

  // ---------------------------------------------------------------- tổng hợp
  function summarize(list) {
    const s = { n: list.length, active: 0, won: 0, lost: 0, value: 0, valueActive: 0, valueWon: 0, valueLost: 0, withValue: 0,
      remind: 0, stale2: 0, overdue: 0, soon30: 0, issues: 0, values: [] };
    for (const d of list) {
      const v = d.value || 0;
      s.value += v;
      if (d.value) { s.withValue++; s.values.push(d.value); }
      if (d.status === 'Active') { s.active++; s.valueActive += v; }
      else if (d.status === 'Won') { s.won++; s.valueWon += v; }
      else if (d.status === 'Lost') { s.lost++; s.valueLost += v; }
      if (d.needRemind) s.remind++;
      if (d.needRemind && +d.stale >= 2) s.stale2++;
      if (d.overdue) s.overdue++;
      if (d.status === 'Active' && d.closeState === 'd30') s.soon30++;
      if (d.issues.length) s.issues++;
    }
    s.winRate = s.won + s.lost ? s.won / (s.won + s.lost) : null;
    s.avg = s.withValue ? s.value / s.withValue : null;
    s.median = U.median(s.values);
    return s;
  }

  /** Gom nhóm theo khoá (fn trả về 1 giá trị hoặc mảng). */
  function groupBy(list, fn) {
    const m = new Map();
    for (const d of list) {
      let ks = fn(d);
      if (!Array.isArray(ks)) ks = [ks];
      if (!ks.length) ks = [NONE];
      for (let k of ks) {
        if (k == null || k === '') k = NONE;
        if (!m.has(k)) m.set(k, []);
        m.get(k).push(d);
      }
    }
    return m;
  }

  root.CRMModel = {
    U, NONE, STATUSES, STATUS_COLORS, STAGES, BANDS, BAND_LABEL, STALE, STALE_LABEL, CLOSE, CLOSE_LABEL,
    bandOf, buildDeals, derive, summarize, groupBy, splitQuality,
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
