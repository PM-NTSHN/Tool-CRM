// Kiểm tra bản JS của bộ convert cho kết quả giống hệt reference/convert_deals.py.
// Chạy:  node test/convert.test.mjs <file_raw.xlsx> <file_template_tu_python.xlsx> [dd/mm/yyyy]
// (dữ liệu CRM thật không commit vào repo; truyền đường dẫn file khi chạy)
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
for (const f of ['xlsx-io.js', 'convert.js']) (0, eval)(fs.readFileSync(path.join(here, '..', 'src', f), 'utf8'));
const { XLSXIO, CRMConvert } = globalThis;

const [rawPath, expPath, refArg] = process.argv.slice(2);
if (!rawPath || !expPath) {
  console.error('Usage: node test/convert.test.mjs <raw.xlsx> <expected_template.xlsx> [dd/mm/yyyy]');
  process.exit(2);
}
let ref = CRMConvert.refFromName(rawPath);
if (refArg) { const [d, m, y] = refArg.split('/').map(Number); ref = new Date(Date.UTC(y, m - 1, d)); }

const raw = await XLSXIO.read(fs.readFileSync(rawPath));
const objs = CRMConvert.rowsToObjects(raw.sheets[0].rows);
const res = CRMConvert.convert(objs, ref);
console.log(`JS: ${new Set(res.records.map((r) => r.ID)).size} deal, ${res.records.length} dòng, loại ${res.nTest} deal test, warnings=${res.warnings.length}`);

const exp = await XLSXIO.read(fs.readFileSync(expPath));
const expObjs = CRMConvert.normalizeTemplate(CRMConvert.rowsToObjects(exp.sheets[0].rows));

const norm = (v) => (v instanceof Date ? v.toISOString() : v === undefined ? null : v);
let bad = 0;
if (expObjs.length !== res.records.length) { console.log('Số dòng khác nhau', expObjs.length, res.records.length); bad++; }
for (let i = 0; i < Math.min(expObjs.length, res.records.length); i++) {
  for (const c of CRMConvert.TEMPLATE_COLS) {
    const a = norm(res.records[i][c]), b = norm(expObjs[i][c]);
    if (a !== b && !(typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) < 1e-6)) {
      if (bad < 15) console.log(`Dòng ${i + 2} cột "${c}":\n  JS : ${a}\n  PY : ${b}`);
      bad++;
    }
  }
}

// ghi file template rồi đọc lại để kiểm tra writer
const out = await XLSXIO.write(CRMConvert.templateWorkbook(res));
const back = await XLSXIO.read(out);
const backObjs = CRMConvert.normalizeTemplate(CRMConvert.rowsToObjects(back.sheets[0].rows));
for (let i = 0; i < backObjs.length; i++) {
  for (const c of CRMConvert.TEMPLATE_COLS) {
    const a = norm(res.records[i][c]), b = norm(backObjs[i][c]);
    if (a !== b && !(typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) < 1e-3)) {
      if (bad < 30) console.log(`Round-trip dòng ${i + 2} cột "${c}": ${a} != ${b}`);
      bad++;
    }
  }
}
if (process.env.OUT) fs.writeFileSync(process.env.OUT, out);
console.log(bad ? `FAIL: ${bad} khác biệt` : 'OK: JS khớp 100% với Python, ghi/đọc lại file khớp');
process.exit(bad ? 1 : 0);
