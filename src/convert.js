/* convert.js – Chuyển file export Deals từ CRM (raw) sang dữ liệu template phân tích.
 * Đây là bản chuyển thể 1:1 sang JavaScript của reference/convert_deals.py (cùng quy tắc, cùng tham số,
 * cùng thứ tự ghi chú) để chạy offline trong trình duyệt. Khi sửa quy tắc, sửa đồng thời cả 2 file.
 *
 *   CRMConvert.convert(rawObjects, refDate)  -> { records, nRaw, nTest, refDate, warnings }
 *   CRMConvert.templateWorkbook(result)      -> spec cho XLSXIO.write (file template chuẩn)
 *   CRMConvert.refFromName(fileName)         -> Date | null
 */
(function (root) {
  'use strict';

  // ============================== THAM SỐ ==============================
  const USD_MISLABEL_THRESHOLD = 1000000; // deal gắn USD có giá trị >= ngưỡng này coi là VND nhập nhầm nhãn
  const USD_RATE = 26000;                 // GIẢ ĐỊNH, chỉ dùng cho deal USD < ngưỡng
  const BOM_TOLERANCE = 0.05;             // lệch giá trị so với tổng BOM > 5% thì ghi chú
  const DEAL_SIZE = [[1e9, '< 1B'], [2e9, '1B - 2B'], [5e9, '2B - 5B']]; // >= 5B -> '> 5B'
  const TEST_NAMES = /(^|[^a-z])test([^a-z]|$)/;

  const REQUIRED = ['ID', 'Tên thương vụ', 'Giá trị thương vụ', 'Tiền tệ thương vụ', 'Chủ sở hữu thương vụ',
    'Giai đoạn', 'Trạng thái thương vụ', 'Tên công ty', 'Bảng danh mục sản phẩm dự án',
    'Phân loại HĐ', 'Reseller', 'Tên hãng', 'Thông tin BOM (model, số lượng, năm support...)',
    'Tên dự án', 'Timeline dự án', 'Ngày ký hợp đồng', 'Phân loại nhóm deal',
    'Lý do thất bại', 'Failed additional comment', 'Từ lúc', 'Cập nhật lần cuối'];

  const TEMPLATE_COLS = ['ID', 'Tên thương vụ', 'Tên khách hàng', 'Phân nhóm khách hàng', 'Sale', 'Reseller', 'Giai đoạn',
    'Trạng thái', 'Giá trị thương vụ', 'Hãng', 'Thông tin BOM (model, số lượng, năm support...)', 'Phân loại HĐ',
    'Phân loại deal size', 'Close date', 'Lý do Failed', 'Ngày tạo', 'Cập nhật lần cuối', 'Cờ chất lượng'];

  const VENDOR_MAP = {
    'KASPERSKY': 'Kaspersky', 'KAPERSKY': 'Kaspersky', 'GTB': 'GTB', 'SOPHOS': 'Sophos', 'OPSWAT': 'OPSWAT',
    'NETSCOUT': 'NetScout', 'DELINEA': 'Delinea', 'CLOUDFLARE': 'Cloudflare', 'PROGRESS': 'Progress',
    'INFOEXPRESS': 'Infoexpress', 'OPENTEXT': 'OpenText', 'BARRACUDA': 'Barracuda', 'RADWARE': 'Radware',
    'SAFETICA': 'Safetica', 'SAFEBREACH': 'SafeBreach', 'HCL SOFTWARE': 'HCL Software', 'HCL': 'HCL Software',
    'ZECURION': 'Zecurion', 'ACRONIS': 'Acronis', 'QUALYS': 'Qualys', 'STELLAR CYBER': 'Stellar Cyber',
    'PENTA SECURITY': 'Penta Security', 'ARCON': 'Arcon', 'PENTERA': 'Pentera', 'PENTARA': 'Pentera',
    'TRUSTWAVE': 'Trustwave', 'NETGEAR': 'Netgear', 'THALES': 'Thales', 'FORCEPOINT': 'Forcepoint',
    'SECPOD': 'SecPod', 'PROOFPOINT': 'Proofpoint', 'CYBLE': 'Cyble', 'FORTRA': 'Fortra', 'VIAVI': 'Viavi',
    'TRELLIX': 'Trellix', 'HÃNG KHÁC': 'Khác', 'SẢN PHẨM KHÁC': 'Khác', 'KHÁC': 'Khác',
  };
  const VENDOR_KEYWORDS = [
    ['Kaspersky', /kaspersky|kapersky|\bkas\b/], ['Sophos', /sophos/], ['GTB', /\bgtb\b/],
    ['OPSWAT', /opswat|metadefender/], ['NetScout', /netscout|arbor/], ['Delinea', /delinea/],
    ['Cloudflare', /cloudflare/], ['Progress', /progress|whatsup|\bwug\b|loadmaster|kemp/],
    ['Infoexpress', /infoexpress|easynac/], ['OpenText', /opentext|arcsight|fortify|\bsmax\b|\bopds\b|voltage|encase/],
    ['Barracuda', /barracuda/], ['Radware', /radware/], ['Safetica', /safetica/], ['SafeBreach', /safebreach/],
    ['HCL Software', /\bhcl\b|appscan|\bhlc\b/], ['Zecurion', /zecurion/], ['Acronis', /acronis/],
    ['Qualys', /qualys/], ['Stellar Cyber', /stellar/], ['Penta Security', /penta/], ['Arcon', /arcon/],
    ['Pentera', /pentera|pentara/], ['Trellix', /trellix/], ['SecPod', /secpod/], ['Thales', /thales/],
    ['Proofpoint', /proofpoint/], ['Forcepoint', /forcepoint/], ['Trustwave', /trustwave/],
    ['Netgear', /netgear/], ['Viavi', /viavi/],
  ];
  const SOLUTION_KW = new Set(['waf', 'ddos', 'dr', 'fw', 'firewall', 'sophos', 'kaspersky', 'kapersky', 'kas', 'gtb', 'safetica',
    'safebreach', 'smax', 'loadmaster', 'opswat', 'qualys', 'progress', 'pam', 'secpod', 'renew',
    'foundations', 'optimum', 'appscan', 'hlc', 'endpoint', 'dlp', 'nac', 'easynac', 'trellix', 'delinea',
    'netscout', 'opds', 'cloudflare', 'opentext', 'barracuda', 'radware', 'infoexpress', 'acronis',
    'zecurion', 'stellar', 'pentera', 'fortify', 'arcsight', 'wug', 'whatsup', 'hcl', 'kemp', 'edr',
    'siem', 'mfa', 'ztna', 'appsec', 'db', 'security']);
  const FILLER = new Set(['va', 'and', 'plus', 'cap', 'do', 'bo', 'sung']);
  const PROJECT_PREFIX = /^(mua sam|giai phap|du an|trien khai|xay dung|nang cap|thay the)/;

  // ============================== HÀM TIỆN ÍCH ==============================
  function fold(s) {
    return String(s).toLowerCase().normalize('NFD').replace(/đ/g, 'd').replace(/\p{Mn}/gu, '');
  }
  const isStr = (x) => typeof x === 'string';
  const isNum = (x) => typeof x === 'number' && !isNaN(x);
  const isDate = (x) => x instanceof Date && !isNaN(x);

  const ENTITIES = { nbsp: '\u00a0', amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", ndash: '–', mdash: '—',
    hellip: '…', rsquo: '’', lsquo: '‘', rdquo: '”', ldquo: '“', laquo: '«', raquo: '»', bull: '•', middot: '·',
    times: '×', deg: '°', copy: '©', reg: '®', trade: '™', euro: '€', shy: '\u00ad', zwnj: '\u200c', zwj: '\u200d' };
  let decoderEl = null;
  function unescapeHtml(x) {
    return x.replace(/&(#[0-9]+|#[xX][0-9a-fA-F]+|[A-Za-z][A-Za-z0-9]*);/g, (m, e) => {
      if (e[0] === '#') {
        const n = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        try { return String.fromCodePoint(n); } catch (err) { return m; }
      }
      if (e in ENTITIES) return ENTITIES[e];
      if (typeof document !== 'undefined') {
        decoderEl = decoderEl || document.createElement('textarea');
        decoderEl.innerHTML = m;
        return decoderEl.value;
      }
      return m;
    });
  }
  function cleanText(x) {
    if (!isStr(x)) return x;
    x = unescapeHtml(x);
    x = x.replace(/<[^>]+>/g, ' ');
    x = x.replace(/\s+/g, ' ').trim();
    return x || null;
  }

  function parseDate(s) {
    if (isDate(s)) return s;
    if (!isStr(s) || !s.trim()) return null;
    s = s.trim().replace(/-/g, '/');
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{1,2}))?$/);
    if (!m) return null;
    const [d, mo, y, h, mi] = [+m[1], +m[2], +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0];
    if (mo < 1 || mo > 12 || d < 1 || h > 23 || mi > 59) return null;
    const dt = new Date(Date.UTC(y, mo - 1, d, h, mi));
    if (dt.getUTCDate() !== d) return null;
    return dt;
  }

  function roundHalfEven(x) {
    const f = Math.floor(x), d = x - f;
    if (d > 0.5) return f + 1;
    if (d < 0.5) return f;
    return f % 2 === 0 ? f : f + 1;
  }
  function groupInt(x, sep) {
    const r = roundHalfEven(Math.abs(x));
    const s = BigInt(r).toString().replace(/\B(?=(\d{3})+(?!\d))/g, sep);
    return (x < 0 && r !== 0 ? '-' : '') + s;
  }
  const vnd = (x) => groupInt(x, '.');
  function pctSigned(x) {
    const r = roundHalfEven(x * 100);
    return (r < 0 || (r === 0 && (x < 0 || Object.is(x, -0))) ? '-' + Math.abs(r) : '+' + r) + '%';
  }
  const pad2 = (n) => String(n).padStart(2, '0');
  const fmtDMY = (d) => `${pad2(d.getUTCDate())}/${pad2(d.getUTCMonth() + 1)}/${d.getUTCFullYear()}`;

  function toNum(s) {
    if (!isStr(s)) return null;
    const t = s.replace(/,/g, '');
    if (!/^(\d+\.?\d*|\.\d+)$/.test(t)) return null;
    return parseFloat(t);
  }

  const BOM_SPLIT = /\s*\|?\s*,\s*(?=(?:Hãng:|Số lượng:|Tên giải pháp:))/;
  function parseBom(s) {
    if (!isStr(s)) return [];
    const items = [];
    for (let part of s.split(BOM_SPLIT)) {
      part = part.trim();
      if (!part) continue;
      const m = part.match(/Hãng:\s*([^|]*?)\s*(?:\||$)/);
      const vendor = m ? (m[1].trim() || null) : null;
      const t = part.match(/Thành tiền:\s*([\d,.]+)/);
      let amt = t ? toNum(t[1]) : null;
      if (amt === null) {
        const q = part.match(/Số lượng:\s*([\d,.]+)/);
        const p = part.match(/Đơn giá:\s*([\d,.]+)/);
        if (q && p && toNum(q[1]) !== null && toNum(p[1]) !== null) amt = toNum(q[1]) * toNum(p[1]);
      }
      items.push({ vendor, amount: amt });
    }
    return items;
  }

  function pyTitle(s) {
    return s.replace(/\p{L}+/gu, (w) => w[0].toUpperCase() + w.slice(1).toLowerCase());
  }
  function canonVendor(v) {
    if (!v) return null;
    const k = v.trim().toUpperCase();
    return k in VENDOR_MAP ? VENDOR_MAP[k] : pyTitle(v.trim());
  }

  function isSolution(seg) {
    const toks = fold(seg).split(/[^a-z0-9]+/).filter(Boolean);
    if (!toks.length) return false;
    const hit = toks.filter((t) => SOLUTION_KW.has(t) || FILLER.has(t) || /^\d+$/.test(t)).length;
    return hit >= Math.max(1, toks.length * 0.6);
  }

  function customerFromDealName(name) {
    const n = String(name || '').trim().replace(/;+$/, '').trim();
    const f = fold(n);
    if (TEST_NAMES.test(f)) return [n, 'test'];
    if (/^C12-BCA 175 - /.test(n)) return ['C12-BCA 175', 'tach'];
    const mc = n.match(/(?<![\p{L}\p{N}_])cho\s+(.+)$/u);
    if (PROJECT_PREFIX.test(f) && mc) return [mc[1].trim(), 'tach'];
    if (n.includes('_')) {
      const parts = n.split('_').map((p) => p.trim()).filter(Boolean);
      const keep = parts.filter((p) => !isSolution(p));
      if (keep.length && keep.length < parts.length) return [keep.join(' - '), 'tach'];
      if (keep.length === parts.length) return [parts.join(' - '), 'giu'];
      return [n, 'kiem tra'];
    }
    const m = n.match(/^(.*\S)\s*-\s*([^-]+)$/);
    if (m && isSolution(m[2])) return [m[1].trim(), 'tach'];
    if (PROJECT_PREFIX.test(f)) return [n, 'kiem tra'];
    return [n, 'giu'];
  }

  function dealSize(v) {
    if (v === null || v === undefined || isNaN(v) || v <= 0) return null;
    for (const [lim, lab] of DEAL_SIZE) if (v < lim) return lab;
    return '> 5B';
  }

  function vendorInfo(row) {
    const items = parseBom(row['Bảng danh mục sản phẩm dự án']);
    let last = null;
    const amt = {}, cnt = {}, order = [];
    for (const it of items) {
      const v = it.vendor ? canonVendor(it.vendor) : last;
      if (v) {
        last = v;
        if (!order.includes(v)) order.push(v);
        cnt[v] = (cnt[v] || 0) + 1;
        amt[v] = (amt[v] || 0) + (it.amount || 0);
      }
    }
    const bomTotal = items.reduce((s, it) => s + (it.amount || 0), 0);
    const unpriced = items.filter((it) => it.amount === null).length;
    let pool = order.filter((v) => v !== 'Khác');
    if (!pool.length) pool = order;
    if (pool.length) {
      let main = pool[0];
      for (const v of pool) {
        if ((amt[v] || 0) > (amt[main] || 0) || ((amt[v] || 0) === (amt[main] || 0) && (cnt[v] || 0) > (cnt[main] || 0))) main = v;
      }
      const how = (amt[main] || 0) > 0 ? 'theo giá trị BOM lớn nhất' : 'theo hãng có nhiều dòng BOM nhất (BOM không có giá)';
      const rest = order.filter((v) => v !== main).sort((a, b) => -(amt[a] || 0) - -(amt[b] || 0));
      return [main, [main, ...rest], how, bomTotal, unpriced];
    }
    const col = isStr(row['Tên hãng']) ? canonVendor(row['Tên hãng']) : null;
    if (col && col !== 'Khác') return [col, [col], 'từ cột "Tên hãng" của CRM (BOM không ghi hãng)', bomTotal, unpriced];
    const text = ['Tên thương vụ', 'Thông tin BOM (model, số lượng, năm support...)', 'Tên dự án']
      .filter((c) => isStr(row[c])).map((c) => String(row[c])).join(' ');
    const ft = fold(text);
    const found = [];
    for (const [v, rx] of VENDOR_KEYWORDS) {
      const m = rx.exec(ft);
      if (m) found.push([m.index, v]);
    }
    found.sort((a, b) => a[0] - b[0] || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0));
    if (found.length) {
      const vs = found.map((x) => x[1]);
      return [vs[0], vs, 'suy ra từ từ khóa trong tên thương vụ/BOM text', bomTotal, unpriced];
    }
    if (col === 'Khác') return ['Khác', ['Khác'], 'từ cột "Tên hãng" của CRM', bomTotal, unpriced];
    return [null, [], null, bomTotal, unpriced];
  }

  // ============================== DANH MỤC SALE ==============================
  const SALE_NAMES = {
    'thuydinh': 'Đinh Văn Thủy', 'hangnguyen': 'Nguyễn Thu Hằng', 'trinhluong': 'Lương Thị Tuyết Trinh',
    'phamquyet': 'Phạm Văn Quyết', 'theduong': 'Dương Văn Thế', 'quyvy': 'Vy Công Quý',
    'kienkhong': 'Khổng Đức Kiên', 'nhungnguyen': 'Nguyễn Hồng Nhung', 'hungle': 'Lê Huy Hùng',
    'anhle': 'Lê Tuấn Anh', 'thangnguyen': 'Nguyễn Duy Thắng', 'nghiemphan': 'Phan Quế Nghiệm',
    'hungpham': 'Phạm Quốc Hùng', 'tupham': 'Phạm Cao Anh Tú', 'hieudinh': 'Đinh Văn Hiệu',
    'hiendoan': 'Đoàn Thị Hiền', 'hungtran': 'Trần Quang Hưng',
  };
  let SALE_TAGS = null;
  function saleTags() {
    const tags = new Set();
    for (const [user, full] of Object.entries(SALE_NAMES)) {
      const p = fold(full).split(/\s+/).filter(Boolean);
      const fam = p[0], mid = p.slice(1, -1), giv = p[p.length - 1];
      [user, fam + giv, giv + fam, giv, p.join(''), mid.slice(-1).join('') + giv].forEach((t) => tags.add(t));
      if (p.length >= 2) tags.add(p.slice(-2).join(''));
    }
    return new Set([...tags].filter((t) => t.length >= 4));
  }
  function stripSaleTag(name) {
    const m = name.match(/^(.*\S)\s*_\s*([^_]+)$/);
    SALE_TAGS = SALE_TAGS || saleTags();
    if (m && SALE_TAGS.has(fold(m[2]).replace(/[^a-z]/g, ''))) return [m[1].trim(), m[2].trim()];
    return [name, null];
  }

  // ============================== PHÂN NHÓM KHÁCH HÀNG ==============================
  const SEGMENTS = ['Cơ quan Đảng, Nhà nước', 'Bộ, Ngành, Cơ quan TW', 'UBND / Sở ban ngành địa phương',
    'An ninh - Quốc phòng', 'Ngân hàng - Tài chính - Bảo hiểm', 'Viễn thông - CNTT',
    'Năng lượng - Điện - Dầu khí', 'Y tế - Giáo dục', 'Sản xuất - Thương mại - Dịch vụ',
    'Giao thông - Vận tải - Logistics', 'DNNN khác', 'Khác'];
  const SEGMENT_RULES = [
    ['Viễn thông - CNTT', /viettel|\bvtg\b|natcom/],
    ['An ninh - Quốc phòng', new RegExp('cong an|\\bbca\\b|\\bbo ca\\b|^ca |\\bca (tp|tinh|thanh pho)\\b|\\bpa0\\d|' +
      '\\b(a0[1-9]|c0[1-9]|c1[0-2]|c500|h0[1-9]|v0[1-9]|b01|x01)\\b|quoc phong|\\bbqp\\b|quan doi|' +
      'quan khu|bo tu lenh|binh chung|bien phong|canh sat|quan y|hai quan|military|' +
      'tong cuc (2|ii|tinh bao|ky thuat|chinh tri)|co yeu|si quan|tong tham muu|tac chien|' +
      'an ninh nhan dan|bo tham muu|bo doi')],
    ['Cơ quan Đảng, Nhà nước', new RegExp('tinh uy|thanh uy|quan uy|huyen uy|\\bvptu\\b|trung uong dang|tw dang|vptw|' +
      'tap chi cong san|phu chu tich|van phong chu tich nuoc|quoc hoi|toa an|vien kiem sat|' +
      'kiem toan nha nuoc')],
    ['UBND / Sở ban ngành địa phương', new RegExp('\\bubnd\\b|uy ban nhan dan|^so |\\bso (khoa|kh|thong tin|tttt|4t|xay dung|' +
      'dan toc|tai chinh|y te|giao duc|noi vu|tu phap|ke hoach|cong thuong|van hoa)\\b|' +
      'bqlda|ban quan ly khu kinh te|ban duy tu|chuyen doi so|' +
      'phuc vu hanh chinh cong|dai phat thanh|bao va phat thanh')],
    ['Bộ, Ngành, Cơ quan TW', new RegExp('^bo |\\bbo (noi vu|tai chinh|tai nguyen|ngoai giao|dan toc|xay dung|khoa hoc|' +
      'cong thuong|y te|giao duc|tu phap|ke hoach|nong nghiep|giao thong|van hoa|thong tin)\\b|' +
      '\\bcuc\\b|tong cuc|kho bac|bao hiem xa hoi|\\bbhxh\\b|ngan hang nha nuoc|' +
      '(uy ban|ub) chung khoan|uy ban tieu chuan|quy phat trien|thong tin tin dung|' +
      '\\bcic\\b|vnnic|ttxvn|thong tan xa|vien han lam')],
    ['Năng lượng - Điện - Dầu khí', new RegExp('\\bevn|dien luc|thuy dien|nhiet dien|truyen tai dien|he thong dien|' +
      'thi truong dien|dieu do|\\bnpt\\b|\\bptc ?\\d|^pc |genco|\\bpvn\\b|pv ?gas|petro|' +
      'dau khi|khi viet nam|^than |\\btkv\\b|vimico|khoang san|nhien lieu|xang dau|nsrp|' +
      'npcit|\\bnpc\\b|\\bspc\\b')],
    ['Y tế - Giáo dục', new RegExp('benh vien|^bv|\\bbvdk\\b|y te|huyet hoc|^truong |dai hoc|hoc vien|giao duc|educa|' +
      '\\bttyt\\b|hospital')],
    ['Ngân hàng - Tài chính - Bảo hiểm', new RegExp('ngan hang|bank|chung khoan|securities|bao hiem|insurance|tai chinh vi mo|' +
      'finance|napas|vnpay|thanh toan|\\b(bidv|vietcombank|vietinbank|agribank|' +
      'mbbank|tpbank|vib|scb|shb|shbfc|lpb|lpbs|exim|eximbank|tcbs|hsc|kbsv|' +
      'vndirect|mbs|vps|shs|dnse|dsc|agriseco|opes|pjico|pvi|insmart|mfinance|' +
      'timo|oceanbank|publicbank)\\b')],
    ['Giao thông - Vận tải - Logistics', new RegExp('hang khong|airline|airport|san bay|\\bacv\\b|vaeco|viags|sasco|\\bvna\\b|' +
      'truc thang|\\bvnh\\b|quan ly bay|vnaic|hang hai|giao hang|\\bghtk\\b|' +
      'logistics|van tai|duong sat|\\bvetc\\b|\\bitl\\b|vnpost|buu chinh')],
    ['Viễn thông - CNTT', new RegExp('vnpt|mobifone|\\bcmc\\b|telecom|vien thong|\\bfpt\\b|\\bfis\\b|\\bctin\\b|cong nghe|technology|' +
      'tecapro|gosu|vtvcab|software|vietbay|tntech')],
    ['DNNN khác', /cap nuoc|sawaco|vietlott|xo so|trac dia|dong tau/],
    ['Sản xuất - Thương mại - Dịch vụ', new RegExp('cong ty|\\bcty\\b|\\bctcp\\b|tnhh|tap doan|group|corporation|\\bcorp\\b|' +
      'limited|\\bjsc\\b|holdings')],
  ];
  const SXTMDV = 'Sản xuất - Thương mại - Dịch vụ';
  const SEGMENT_OVERRIDES = {
    'tong cong ty cp buu chinh viettel': 'Giao thông - Vận tải - Logistics',
    'viettel post': 'Giao thông - Vận tải - Logistics',
    'cong ty thuong mai & xuat nhap khau/ thuong mai & xnk viettel (viettel commerce)': SXTMDV,
    'cong ty lien doanh thap ngan hang dau tu va phat trien viet nam': SXTMDV,
    'vietinbank gold & jewellery': 'Ngân hàng - Tài chính - Bảo hiểm',
    'pvi': 'Ngân hàng - Tài chính - Bảo hiểm',
    'cong ty tnhh mot thanh vien dong tau hong ha': 'DNNN khác',
    'cong ty tnhh mtv trac dia ban do': 'DNNN khác',
    'tap doan phenikaa': SXTMDV,
    'access trade': SXTMDV, 'career viet': SXTMDV,
    'diana unicharm': SXTMDV, 'everland van don': SXTMDV,
    'fujikin': SXTMDV, 'freshmart (c.p)': SXTMDV,
    'fujimart': SXTMDV, 'jarllytec': SXTMDV,
    'maison': SXTMDV, 'masterise': SXTMDV,
    'messer viet nam': SXTMDV, 'nutifood': SXTMDV,
    'panasonic r&d': SXTMDV, 'rang dong': SXTMDV,
    'shemar power vn': SXTMDV, 'sun group': SXTMDV,
    'yakult viet nam': SXTMDV, 'yokohama tyre viet nam': SXTMDV,
    'berjaya gia thinh': SXTMDV,
    'viet duc co so 3': 'Y tế - Giáo dục',
    'cong ty tnhh mtv van hanh htd va ttd quoc gia': 'Năng lượng - Điện - Dầu khí',
    'gtb endpoint protector': 'Khác', 'giai phap attt cap do 3': 'Khác',
  };
  const SEGMENT_UNSURE = new Set(['hal vn', 'mcst', 'itdb', 'lpex', 'plc', 'intech',
    'trung tam nghien cuu va ung dung cong nghe truyen thong (r&d)']);

  function classifyCustomer(name) {
    const f = fold(name).replace(/\s+/g, ' ').trim();
    if (SEGMENT_UNSURE.has(f)) return [f.includes('cong nghe') ? 'Viễn thông - CNTT' : 'Khác', false];
    if (f in SEGMENT_OVERRIDES) return [SEGMENT_OVERRIDES[f], true];
    for (const [seg, rx] of SEGMENT_RULES) {
      if (rx.test(f)) {
        if (seg === 'Viễn thông - CNTT' && /viettel/.test(f)) {
          if (/buu chinh|post/.test(f)) return ['Giao thông - Vận tải - Logistics', true];
          if (/commerce|thuong mai/.test(f)) return [SXTMDV, true];
        }
        return [seg, true];
      }
    }
    return ['Khác', false];
  }

  // ============================== XỬ LÝ CHÍNH ==============================
  /* Mô phỏng cách pandas.read_excel suy kiểu dữ liệu: các chuỗi rỗng / "NA" / "null"… thành ô trống,
   * cột mà mọi giá trị đều là số (kể cả số lưu dạng chuỗi) thì đổi thành số. File export CRM lưu
   * toàn bộ ô dưới dạng chuỗi nên bước này cần để giá trị, ID… được hiểu là số như bản Python. */
  const NA_VALUES = new Set(['', '#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', '-NaN', '-nan', '1.#IND', '1.#QNAN',
    '<NA>', 'N/A', 'NA', 'NULL', 'NaN', 'None', 'n/a', 'nan', 'null']);
  const NUMERIC = /^\s*[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\s*$/;
  function pandasCoerce(rows, cols) {
    const out = rows.map((r) => {
      const o = {};
      for (const c of cols) {
        let v = r[c] === undefined ? null : r[c];
        if (isStr(v) && NA_VALUES.has(v)) v = null;
        o[c] = v;
      }
      return o;
    });
    for (const c of cols) {
      const vals = out.map((o) => o[c]).filter((v) => v !== null);
      if (vals.length && vals.every((v) => isNum(v) || (isStr(v) && NUMERIC.test(v)))) {
        for (const o of out) if (o[c] !== null) o[c] = Number(o[c]);
      }
    }
    return out;
  }

  function convert(raw, refDate) {
    const warnings = [];
    if (!raw.length) throw new Error('File raw không có dòng dữ liệu nào.');
    const cols = new Set(Object.keys(raw.columnsFrom || raw[0]));
    raw.forEach((r) => Object.keys(r).forEach((k) => cols.add(k)));
    const missing = REQUIRED.filter((c) => !cols.has(c));
    if (missing.length) throw new Error('File raw thiếu cột: ' + missing.join(', ') + '. Kiểm tra lại cấu hình export CRM.');
    const seen = new Set();
    for (const r of raw) {
      if (seen.has(r['ID'])) throw new Error('File raw có ID trùng (' + r['ID'] + ') – dừng để kiểm tra.');
      seen.add(r['ID']);
    }
    const nRaw = raw.length;

    let df = pandasCoerce(raw, [...cols]).map((r) => {
      const o = {};
      for (const k of Object.keys(r)) o[k] = cleanText(r[k]);
      return o;
    });
    for (const c of ['Timeline dự án', 'Ngày ký hợp đồng', 'Từ lúc', 'Cập nhật lần cuối']) {
      const bad = [];
      for (const r of df) {
        const v = r[c];
        const p = parseDate(v);
        if (v !== null && v !== undefined && p === null) bad.push(v);
        r[c] = p;
      }
      if (bad.length) warnings.push(`${bad.length} giá trị ngày không đọc được ở cột "${c}": ${bad.slice(0, 5).join(', ')}`);
    }

    // ---- Bước 1: Tên khách hàng
    const testRows = new Set();
    for (const r of df) {
      if (isStr(r['Tên công ty']) && r['Tên công ty']) {
        r._cust = r['Tên công ty']; r._cust_note = null;
        if (TEST_NAMES.test(fold(r['Tên thương vụ'] || ''))) testRows.add(r['ID']);
        continue;
      }
      let [c, how] = customerFromDealName(r['Tên thương vụ']);
      if (how === 'giu' && isSolution(c)) how = 'kiem tra';
      r._cust = c;
      if (how === 'test') { testRows.add(r['ID']); r._cust_note = null; }
      else if (how === 'tach') r._cust_note = 'Tên KH lấy từ tên thương vụ (CRM trống "Tên công ty"), đã bỏ phần hãng/giải pháp/reseller';
      else if (how === 'kiem tra') r._cust_note = 'Tên KH lấy nguyên tên thương vụ nhưng tên giống tên dự án/giải pháp – cần kiểm tra lại';
      else r._cust_note = 'Tên KH lấy từ tên thương vụ (CRM trống "Tên công ty")';
    }
    for (const r of df) {
      const [c, t] = stripSaleTag(r._cust);
      r._cust = c; r._sale_tag = t;
    }
    df = df.filter((r) => !testRows.has(r['ID']));

    // gộp biến thể tên KH chỉ khác hoa/thường, dấu, khoảng trắng
    const groups = new Map();
    for (const r of df) {
      r._key = fold(r._cust).replace(/[^a-z0-9]+/g, ' ').trim();
      if (!groups.has(r._key)) groups.set(r._key, new Map());
      const g = groups.get(r._key);
      g.set(r._cust, (g.get(r._cust) || 0) + 1);
    }
    const isUpper = (s) => s === s.toUpperCase() && s !== s.toLowerCase();
    const canon = new Map();
    for (const [k, g] of groups) {
      const best = [...g.keys()].sort((a, b) => {
        const ka = [fold(a) === a.toLowerCase(), isUpper(a), -g.get(a)], kb = [fold(b) === b.toLowerCase(), isUpper(b), -g.get(b)];
        for (let i = 0; i < 3; i++) if (ka[i] !== kb[i]) return ka[i] < kb[i] ? -1 : 1;
        return a < b ? -1 : a > b ? 1 : 0;
      })[0];
      canon.set(k, best);
    }
    for (const r of df) r._cust_std = canon.get(r._key);

    // ---- Bước 2: hãng, giá trị, cờ chất lượng
    for (const r of df) {
      const [main, all, how, bom, unpriced] = vendorInfo(r);
      Object.assign(r, { _main: main, _all: all, _vhow: how, _bom: bom, _unpriced: unpriced });
    }
    const nameCounts = new Map();
    for (const r of df) {
      const k = r['Tên thương vụ'];
      if (!nameCounts.has(k)) nameCounts.set(k, []);
      nameCounts.get(k).push(r['ID']);
    }

    const out = [];
    for (const r of df) {
      const v0 = isNum(r['Giá trị thương vụ']) ? r['Giá trị thương vụ'] : null;
      const cur = r['Tiền tệ thương vụ'], bom = r._bom;
      const issues = [], notes = [];
      let val;
      if ((v0 === null || v0 === 0) && bom > 0) {
        val = bom; notes.push(`Giá trị lấy từ tổng BOM (${vnd(bom)}) vì giá trị CRM = 0`);
      } else if (v0 === null || v0 === 0) {
        val = null; issues.push('Thiếu giá trị thương vụ (CRM = 0, BOM không có giá)');
      } else if (cur === 'USD' && v0 < USD_MISLABEL_THRESHOLD) {
        val = v0 * USD_RATE; notes.push(`Quy đổi ${groupInt(v0, ',')} USD x ${groupInt(USD_RATE, ',')} (tỷ giá giả định)`);
      } else {
        val = v0;
        if (cur === 'USD') notes.push('CRM ghi tiền tệ USD nhưng số tiền là VND – giữ số gốc, coi là VND');
      }
      if (val && bom > 0 && Math.abs(val - bom) > BOM_TOLERANCE * bom) {
        issues.push(`Giá trị ${vnd(val)} lệch tổng BOM ${vnd(bom)} (${pctSigned((val - bom) / bom)})`);
      }
      if (r._unpriced > 0) issues.push(`BOM có ${r._unpriced} dòng chưa có giá`);
      const st = r['Trạng thái thương vụ'], gd = r['Giai đoạn'];
      if (st === 'Active' && gd === 'Thực hiện hợp đồng') issues.push('Trạng thái Active nhưng giai đoạn CRM là "Thực hiện hợp đồng" (có thể đã Won)');
      if (st === 'Won' && gd !== 'Thực hiện hợp đồng') issues.push(`Trạng thái Won nhưng giai đoạn CRM là "${gd == null ? 'None' : gd}"`);
      if (st === 'Won' && !r['Ngày ký hợp đồng']) issues.push('Won nhưng thiếu ngày ký hợp đồng');
      if (st === 'Lost' && !isStr(r['Lý do thất bại'])) issues.push('Lost nhưng thiếu lý do thất bại');
      const cd = r['Timeline dự án'];
      if (cd && st === 'Active' && cd < refDate) issues.push(`Close date ${fmtDMY(cd)} đã qua nhưng deal vẫn Active`);
      const same = nameCounts.get(r['Tên thương vụ']).filter((i) => i !== r['ID']);
      if (same.length) {
        const ids = same.slice(0, 5).map(String).join(', ') + (same.length > 5 ? ' …' : '');
        issues.push(`Trùng tên thương vụ với ${same.length} deal khác (ID ${ids})`);
      }
      if (!isStr(r['Chủ sở hữu thương vụ'])) issues.push('Thiếu Sale (owner) trong CRM');
      const miss = [['Close date', !!cd], ['Reseller', isStr(r['Reseller'])], ['Phân loại HĐ', isStr(r['Phân loại HĐ'])]]
        .filter((x) => !x[1]).map((x) => x[0]);
      if (miss.length) issues.push('Thiếu thông tin: ' + miss.join(', '));
      if (!isStr(r._main)) issues.push('Chưa xác định được hãng (BOM, cột Tên hãng và tên thương vụ đều không có)');
      else notes.push(`Hãng chính xác định ${r._vhow}`);
      if (isStr(r._cust_note)) (r._cust_note.includes('cần kiểm tra') ? issues : notes).push(r._cust_note);
      if (isStr(r._sale_tag)) notes.push(`Đã bỏ hậu tố tên sale "_${r._sale_tag}" trong tên KH`);
      if (r._cust_std !== r._cust) notes.push(`Tên KH chuẩn hóa từ "${r._cust}"`);
      const ds = dealSize(val);
      const ds0 = r['Phân loại nhóm deal'];
      if (ds && !isStr(ds0)) notes.push('Deal size tính từ giá trị (CRM trống)');
      else if (ds && ds0 !== ds) notes.push(`Deal size CRM "${ds0}" không khớp giá trị – đã tính lại`);
      const u = r['Chủ sở hữu thương vụ'];
      let sale;
      if (isStr(u) && u.trim().toLowerCase() in SALE_NAMES) sale = SALE_NAMES[u.trim().toLowerCase()];
      else {
        sale = isStr(u) ? u : null;
        if (isStr(u)) issues.push(`Sale "${u}" chưa có trong bảng tên sale – giữ tên tài khoản CRM`);
      }
      const [seg, sure] = classifyCustomer(r._cust_std);
      if (!sure) issues.push(`Phân nhóm KH "${seg}" chưa chắc chắn – cần kiểm tra`);
      const reasons = [r['Lý do thất bại'], r['Failed additional comment']].filter(isStr);
      const failed = reasons.length ? reasons.join(' – ') : null;

      const vendors = r._all.length ? r._all : [null];
      const n = vendors.length;
      vendors.forEach((ven, k) => {
        let extra = k === 0 ? [] : [`Dòng tách hãng ${k + 1}/${n} của deal ID ${r['ID']}; giá trị thương vụ chỉ ghi ở dòng hãng chính (${vendors[0]})`];
        if (k === 0 && n > 1) extra = [`Deal có ${n} hãng: ${vendors.join(', ')} – đã tách thành ${n} dòng, giá trị ghi ở dòng này`];
        const text = [];
        if (issues.length) text.push('[VẤN ĐỀ] ' + issues.join('; '));
        if (notes.length || extra.length) text.push('[ĐÃ XỬ LÝ] ' + [...extra, ...notes].join('; '));
        out.push({
          'ID': r['ID'], 'Tên thương vụ': r['Tên thương vụ'], 'Tên khách hàng': r._cust_std,
          'Phân nhóm khách hàng': seg, 'Sale': sale, 'Reseller': r['Reseller'] ?? null, 'Giai đoạn': gd ?? null, 'Trạng thái': st ?? null,
          'Giá trị thương vụ': k === 0 ? val : null, 'Hãng': ven,
          'Thông tin BOM (model, số lượng, năm support...)': r['Thông tin BOM (model, số lượng, năm support...)'] ?? null,
          'Phân loại HĐ': r['Phân loại HĐ'] ?? null, 'Phân loại deal size': ds,
          'Close date': cd, 'Lý do Failed': failed, 'Ngày tạo': r['Từ lúc'],
          'Cập nhật lần cuối': r['Cập nhật lần cuối'], 'Cờ chất lượng': text.join(' | ') || null,
        });
      });
    }
    return { records: out, nRaw, nTest: testRows.size, refDate, warnings };
  }

  // ============================== GHI FILE TEMPLATE ==============================
  function templateWorkbook(res) {
    const F = 'Arial';
    const bd = 'D9D9D9';
    const widths = { 'ID': 10, 'Tên thương vụ': 42, 'Tên khách hàng': 38, 'Phân nhóm khách hàng': 26, 'Sale': 20, 'Reseller': 18, 'Giai đoạn': 20,
      'Trạng thái': 11, 'Thông tin BOM (model, số lượng, năm support...)': 50,
      'Giá trị thương vụ': 18, 'Hãng': 15, 'Phân loại HĐ': 13, 'Phân loại deal size': 12, 'Close date': 12,
      'Lý do Failed': 40, 'Ngày tạo': 16, 'Cập nhật lần cuối': 16, 'Cờ chất lượng': 90 };
    const hdr = { font: { name: F, b: true, color: 'FFFFFF', sz: 10 }, fill: '1F4E78', border: bd, align: { h: 'center', v: 'center', wrap: true } };
    const rows = [TEMPLATE_COLS.map((h) => ({ v: h, s: hdr }))];
    for (const rec of res.records) {
      const isSplit = isStr(rec['Cờ chất lượng']) && rec['Cờ chất lượng'].includes('Dòng tách hãng');
      rows.push(TEMPLATE_COLS.map((h) => {
        const s = { font: { name: F, sz: 10 }, border: bd };
        if (h === 'Giá trị thương vụ') s.numFmt = '#,##0';
        else if (h === 'Close date') s.numFmt = 'dd/mm/yyyy';
        else if (h === 'Ngày tạo' || h === 'Cập nhật lần cuối') s.numFmt = 'dd/mm/yyyy hh:mm';
        else if (h === 'ID') s.numFmt = '0';
        if (isSplit) s.fill = 'F2F2F2';
        return { v: rec[h] ?? null, s };
      }));
    }
    const nDeals = new Set(res.records.map((r) => r['ID'])).size;
    const lines = [
      ['HƯỚNG DẪN ĐỌC FILE', null],
      ['Ngày chốt số liệu', fmtDMY(res.refDate)],
      ['Số deal trong file raw', res.nRaw],
      ['Deal test đã loại', res.nTest],
      ['Số deal (ID duy nhất)', nDeals],
      ['Số dòng trong sheet Data', res.records.length],
      ['Quy tắc đếm', 'Một deal nhiều hãng được tách thành nhiều dòng cùng ID. Đếm số deal phải dùng ID duy nhất. ' +
        'Giá trị thương vụ chỉ ghi ở dòng hãng chính (dòng tách tô nền xám, giá trị để trống), nên cộng ' +
        'cột Giá trị thương vụ không bị đếm trùng.'],
      ['Giá trị thương vụ', 'Đơn vị VND. Để trống = chưa có dữ liệu (không phải 0). Khi CRM = 0 thì lấy từ tổng BOM.'],
      ['Hãng', 'Dòng đầu của mỗi deal là hãng chính (hãng có giá trị BOM lớn nhất). Để trống = chưa xác định được.'],
      ['Phân loại deal size', '< 1B: dưới 1 tỷ · 1B - 2B: 1 đến dưới 2 tỷ · 2B - 5B: 2 đến dưới 5 tỷ · > 5B: từ 5 tỷ. ' +
        'Tính lại từ Giá trị thương vụ, không dùng phân loại CRM vì CRM ghi không nhất quán.'],
      ['Sale', 'Đổi tài khoản CRM (cột "Chủ sở hữu thương vụ") sang họ tên theo bảng ở sheet Danh_muc. ' +
        'Tài khoản chưa có trong bảng thì giữ nguyên và gắn cờ.'],
      ['Phân nhóm khách hàng', 'Phân tự động theo tên KH với 12 khối ở sheet Danh_muc. Cơ quan nhà nước phân theo cấp ' +
        '(Đảng/NN, Bộ ngành TW, địa phương); mọi đơn vị thuộc Bộ Công an, Bộ Quốc phòng (kể cả bệnh viện, ' +
        'trường) vào An ninh - Quốc phòng; doanh nghiệp phân theo ngành kinh doanh chính, công ty con theo ' +
        'ngành của tập đoàn (EVN → Năng lượng, Viettel → Viễn thông, trừ Viettel Post → Logistics); ' +
        'bệnh viện/trường dân sự vào Y tế - Giáo dục. KH chưa phân được hoặc chưa chắc chắn có cờ cần kiểm tra.'],
      ['Giai đoạn', 'Giai đoạn dự án theo CRM: Đăng kí cơ hội → Tư vấn → Phương án kinh doanh → Xây dựng hồ sơ → ' +
        'Thực hiện hợp đồng. Giữ nguyên giá trị CRM; mâu thuẫn với Trạng thái được ghi ở Cờ chất lượng.'],
      ['Thông tin BOM', 'Lấy nguyên cột "Thông tin BOM (model, số lượng, năm support...)" của CRM (đã bỏ ký tự HTML, ' +
        'khoảng trắng thừa). Cột này do sale nhập tay, phần lớn deal để trống.'],
      ['Close date', 'Lấy từ cột "Timeline dự án" của CRM.'],
      ['Lý do Failed', 'Gộp "Lý do thất bại" và "Failed additional comment": <lý do> – <ghi chú>.'],
      ['Cờ chất lượng', '[VẤN ĐỀ] = dữ liệu thiếu/mâu thuẫn cần người phụ trách kiểm tra. ' +
        '[ĐÃ XỬ LÝ] = ghi lại thay đổi so với CRM (tên KH, hãng, giá trị, deal size). Trống = không có ghi chú.'],
      ['Tham số', `Ngưỡng USD nhập nhầm: ${groupInt(USD_MISLABEL_THRESHOLD, ',')}; tỷ giá giả định: ${groupInt(USD_RATE, ',')}; ` +
        `ngưỡng lệch BOM: ${Math.round(BOM_TOLERANCE * 100)}%.`],
    ];
    const g = lines.map(([a, b], i) => [
      { v: a, s: { font: { name: F, sz: i === 0 ? 12 : 10, b: true }, align: { v: 'top' } } },
      { v: b, s: { font: { name: F, sz: 10 }, align: { wrap: true, v: 'top' } } },
    ]);
    const hdr2 = { font: { name: F, b: true, color: 'FFFFFF', sz: 10 }, fill: '1F4E78' };
    const dm = [[{ v: 'Tài khoản CRM', s: hdr2 }, { v: 'Tên Sale', s: hdr2 }, undefined, { v: 'Phân nhóm khách hàng', s: hdr2 }, { v: 'Số deal', s: hdr2 }, { v: 'Ví dụ khách hàng', s: hdr2 }]];
    const fs = { font: { name: F, sz: 10 } };
    Object.entries(SALE_NAMES).forEach(([u, n], i) => { dm[i + 1] = dm[i + 1] || []; dm[i + 1][0] = { v: u, s: fs }; dm[i + 1][1] = { v: n, s: fs }; });
    const firstById = new Map();
    for (const r of res.records) if (!firstById.has(r['ID'])) firstById.set(r['ID'], r);
    const first = [...firstById.values()];
    SEGMENTS.forEach((seg, i) => {
      const sub = first.filter((r) => r['Phân nhóm khách hàng'] === seg);
      const cnt = new Map();
      sub.forEach((r) => cnt.set(r['Tên khách hàng'], (cnt.get(r['Tên khách hàng']) || 0) + 1));
      const ex = [...cnt.entries()].sort((a, b) => b[1] - a[1]).slice(0, 4).map((x) => x[0]).join(', ');
      dm[i + 1] = dm[i + 1] || [];
      dm[i + 1][3] = { v: seg, s: fs }; dm[i + 1][4] = { v: sub.length, s: fs }; dm[i + 1][5] = { v: ex, s: fs };
    });
    return {
      title: 'CRM deals template',
      sheets: [
        { name: 'Data', rows, cols: TEMPLATE_COLS.map((h) => widths[h] || 15), freeze: { row: 1, col: 2 },
          autoFilter: `A1:${root.XLSXIO.colName(TEMPLATE_COLS.length - 1)}${res.records.length + 1}` },
        { name: 'Huong_dan', rows: g, cols: [26, 110] },
        { name: 'Danh_muc', rows: dm, cols: [16, 24, 3, 32, 9, 90] },
      ],
    };
  }

  // ============================== ĐỌC FILE ==============================
  /** Chuyển mảng hàng (row[0] là tiêu đề) thành mảng object theo tên cột. */
  function rowsToObjects(rows) {
    let h = 0;
    while (h < rows.length && !(rows[h] || []).some((x) => x !== null && x !== undefined && x !== '')) h++;
    const header = (rows[h] || []).map((x) => (x == null ? '' : String(x).trim()));
    const out = [];
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (!r.some((x) => x !== null && x !== undefined && x !== '')) continue;
      const o = {};
      header.forEach((k, j) => { if (k) o[k] = r[j] === undefined ? null : r[j]; });
      out.push(o);
    }
    out.columnsFrom = Object.fromEntries(header.filter(Boolean).map((k) => [k, 1]));
    return out;
  }
  function isTemplate(objs) {
    const c = objs.columnsFrom || objs[0] || {};
    return ['ID', 'Tên khách hàng', 'Phân nhóm khách hàng', 'Hãng', 'Cờ chất lượng'].every((k) => k in c);
  }
  function isRaw(objs) {
    const c = objs.columnsFrom || objs[0] || {};
    return ['Tên thương vụ', 'Chủ sở hữu thương vụ', 'Trạng thái thương vụ'].every((k) => k in c);
  }
  /** Chuẩn hoá các bản ghi đọc từ file template có sẵn (ngày dạng chuỗi -> Date). */
  function normalizeTemplate(objs) {
    return objs.map((o) => {
      const r = {};
      for (const k of TEMPLATE_COLS) {
        let v = o[k] === undefined ? null : o[k];
        if (isStr(v)) v = v.trim() || null;
        if (['Close date', 'Ngày tạo', 'Cập nhật lần cuối'].includes(k) && v !== null && !isDate(v)) v = parseDate(String(v));
        if (k === 'Giá trị thương vụ' && isStr(v)) v = toNum(v.replace(/\./g, ''));
        r[k] = v;
      }
      return r;
    });
  }

  function refFromName(path) {
    const m = String(path).match(/[_.](\d{2})[_.](\d{2})[_.](\d{2})\.xlsx$/i);
    if (m) {
      const d = +m[1], mo = +m[2], y = 2000 + +m[3];
      const dt = new Date(Date.UTC(y, mo - 1, d));
      if (mo >= 1 && mo <= 12 && dt.getUTCDate() === d) return dt;
    }
    return null;
  }

  root.CRMConvert = {
    convert, templateWorkbook, rowsToObjects, isTemplate, isRaw, normalizeTemplate, refFromName, parseDate,
    fold, fmtDMY, SEGMENTS, SALE_NAMES, TEMPLATE_COLS, REQUIRED,
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
