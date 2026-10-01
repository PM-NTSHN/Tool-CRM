// Đóng gói toàn bộ tool thành 1 file HTML duy nhất chạy offline: node build.mjs
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');
// chặn chuỗi "</script" làm vỡ thẻ script khi nhúng inline
const safe = (js) => js.replace(/<\/script/gi, '<\\/script').replace(/<!--/g, '<\\!--');

const logo = 'data:image/png;base64,' + fs.readFileSync(path.join(root, 'src/logo.png')).toString('base64');
const parts = {
  CSS: read('src/styles.css'),
  CHARTJS: safe(read('vendor/chart.umd.min.js')),
  XLSXIO: safe(read('src/xlsx-io.js')),
  CONVERT: safe(read('src/convert.js') + '\n' + read('src/model.js') + '\n' + read('src/report.js')),
  APP: safe(read('src/app.js')),
};
let html = read('src/index.html');
for (const [k, v] of Object.entries(parts)) html = html.replace(`/*{{${k}}}*/`, () => v);
html = html.replaceAll('{{LOGO}}', logo);
const out = path.join(root, 'Tool tổng hợp CRM.html');
fs.writeFileSync(out, html);
console.log(`Đã tạo ${path.basename(out)} (${(html.length / 1024).toFixed(0)} KB)`);
