// ब-१२: एका स्रोतावरून (tools/help_content.html) — (१) ॲपमधील "❓ मदत" पान (index.html मध्ये) आणि (२) छापता येणारी A4 PDF बनवतो.
// वापर: node tools/build_help.js            → index.html मध्ये मदत-पान/मेनू/कार्ड घालतो (पुन्हा चालवल्यास बदलतो, डुप्लिकेट नाही)
//       node tools/build_help.js --pdf <out.pdf> [--font-dir <Noto Sans Devanagari TTF चा फोल्डर>]   → PDF (Chromium; फॉन्ट embed)
const fs = require('fs'), path = require('path');
const ROOT = path.join(__dirname, '..'), WEB = process.env.SGS_WEB || path.join(ROOT, 'github_pages');
const src = fs.readFileSync(path.join(__dirname, 'help_content.html'), 'utf8').replace(/<!--[\s\S]*?-->/g, '');
const ROLE_LABEL = { teacher: 'वर्ग शिक्षक', deo: 'DEO (डेटा एंट्री)', cert: 'प्रमाणपत्र-वापरकर्ता (Cert)', peon: 'शिपाई', master: 'मुख्याध्यापक / Master', super: 'Super Admin' };

function matchingClose(html, openIdx) {   // openIdx वरील <div …> चा संतुलित </div> — कार्डच्या आत घालण्याची चूक (V19.64) टाळण्यासाठी
  const re = /<div\b|<\/div>/g; re.lastIndex = openIdx; let depth = 0, m;
  while ((m = re.exec(html))) { if (m[0] === '</div>') { depth--; if (depth === 0) return m.index + m[0].length; } else depth++; }
  throw new Error('संतुलित </div> सापडला नाही');
}
function mark(html, name, block, anchorRe, where, balanced) {   // name_BEGIN … name_END दरम्यान बदला; नसल्यास anchor नुसार घाला
  const b = `<!--${name}_BEGIN-->`, e = `<!--${name}_END-->`, full = b + '\n' + block + '\n' + e;
  const i = html.indexOf(b), j = html.indexOf(e);
  if (i >= 0 && j > i) return html.slice(0, i) + full + html.slice(j + e.length);
  const m = anchorRe.exec(html); if (!m) throw new Error('anchor सापडला नाही: ' + name);
  const at = balanced ? matchingClose(html, m.index) : (where === 'before' ? m.index : m.index + m[0].length);
  return html.slice(0, at) + '\n' + full + html.slice(at);
}
function buildApp() {
  const p = path.join(WEB, 'index.html'); let h = fs.readFileSync(p, 'utf8');
  h = mark(h, 'HELP_NAV', '<button class="navbtn" data-page="help" onclick="showPage(\'help\',this)">❓ मदत</button>', /<button class="navbtn" data-page="teacher" onclick="showPage\('teacher',this\)">🏫 वर्ग शिक्षक<\/button>/, 'after');
  h = mark(h, 'HELP_DCARD', '<div class="dcard" data-page="help" onclick="showPage(\'help\',this)">\n<div class="icon">❓</div>\n<h3>मदत</h3>\n<p>भूमिकेनुसार वापर-मार्गदर्शिका, अडचणी व उपाय</p>\n</div>', /<div class="dcard super-only" data-page="dataqc"/, 'after', true);
  const page = '<div id="pg-help" class="page">\n<div class="card"><div class="card-header"><div><h2>❓ मदत — वापरकर्ता-मार्गदर्शिका</h2><p id="help_roleLine"></p></div></div>\n<div class="card-body" id="help_body"></div></div>\n</div>\n<template id="help_tpl">\n' + src.trim() + '\n</template>';
  h = mark(h, 'HELP_PAGE', page, /<div id="pg-dataqc" class="page">/, 'before');
  fs.writeFileSync(p, h); console.log('index.html अद्ययावत (मदत-पान ' + Math.round(page.length / 1024) + ' KB)');
}
function sections() { const out = []; const re = /<section class="hp-sec" data-roles="([^"]+)"[^>]*>([\s\S]*?)<\/section>/g; let m; while ((m = re.exec(src))) out.push({ roles: m[1].split(','), html: m[0] }); return out; }
function printHtml(fontDir) {
  const secs = sections(), common = secs.filter(s => s.roles.includes('all')), order = ['teacher', 'deo', 'cert', 'peon', 'master', 'super'];
  const imgBase = 'file://' + path.join(WEB, '') + '/';
  const fix = h => h.replace(/ loading="lazy"/g, '').replace(/src="help\//g, 'src="' + imgBase + 'help/');
  let body = `<div class="cover"><div class="cover-in"><div class="cover-t">SGS School Management System</div><div class="cover-s">वापरकर्ता-मार्गदर्शिका</div><div class="cover-d">श्री. गोविंदराव सेक्सरिया हायस्कूल, पाचोरा<br>ॲप आवृत्ती V19.64</div><div class="cover-n">प्रत्येक भूमिकेसाठी स्वतंत्र प्रकरण — तुमच्या भूमिकेचे प्रकरण वाचा. सर्वांसाठी असलेले भाग (Login, ऑफलाइन, Pending Saves, अडचणी) आधी आहेत.</div></div></div>`;
  body += `<div class="chap"><h2 class="ch">सर्वांसाठी — सामान्य माहिती</h2>${common.map(s => fix(s.html)).join('\n')}</div>`;
  for (const r of order) { const list = secs.filter(s => s.roles.includes(r) || (r === 'super' && false)); if (!list.length) continue; body += `<div class="chap"><h2 class="ch">भूमिका: ${ROLE_LABEL[r]}</h2>${list.map(s => fix(s.html)).join('\n')}</div>`; }
  const ff = f => 'file://' + path.resolve(fontDir, f);
  return `<!DOCTYPE html><html lang="mr"><head><meta charset="utf-8"><title>SGS वापरकर्ता-मार्गदर्शिका</title><style>
@font-face{font-family:'NotoDev';font-weight:400;src:url('${ff('NotoSansDevanagari_400Regular.ttf')}')}
@font-face{font-family:'NotoDev';font-weight:700;src:url('${ff('NotoSansDevanagari_700Bold.ttf')}')}
@page{size:A4;margin:16mm 14mm 16mm 14mm}
*{box-sizing:border-box}body{font-family:'NotoDev',sans-serif;font-size:10.5pt;line-height:1.65;color:#1a1a1a}
.cover{height:262mm;display:flex;align-items:center;justify-content:center;text-align:center;page-break-after:always;background:#fff}
.cover-in{border:2mm solid #1a2a4a;border-radius:6mm;padding:22mm 14mm;width:100%}.cover-t{font-size:22pt;font-weight:700;color:#1a2a4a;font-family:Arial,sans-serif}.cover-s{font-size:26pt;font-weight:700;color:#b9770e;margin:6mm 0}.cover-d{font-size:12pt;margin-top:6mm}.cover-n{font-size:10.5pt;margin-top:12mm;color:#444}
.chap{page-break-before:always}.chap:first-of-type{page-break-before:auto}
h2.ch{font-size:17pt;color:#fff;background:#1a2a4a;padding:3mm 5mm;border-radius:2mm;margin:0 0 5mm}
section.hp-sec{margin-bottom:6mm}h3{font-size:13.5pt;color:#1a2a4a;border-bottom:0.6mm solid #d4902a;padding-bottom:1mm;margin:5mm 0 2mm;page-break-after:avoid}h4{font-size:11.5pt;color:#7a4a00;margin:4mm 0 1mm;page-break-after:avoid}
p,li{margin:1mm 0}ul,ol{padding-left:6mm;margin:1mm 0}li{page-break-inside:avoid}
figure{margin:3mm 0;text-align:center;page-break-inside:avoid}figure img{width:58mm;border:0.3mm solid #999;border-radius:2mm}figcaption{font-size:9pt;color:#555;margin-top:1mm}
.hp-tip{background:#fff4d6;border:0.4mm solid #b8860b;border-radius:2mm;padding:2mm 3mm;margin:3mm 0;font-size:10pt;page-break-inside:avoid}.hp-small{font-size:9pt;color:#555}
table.hp-tbl{border-collapse:collapse;width:100%;font-size:10pt;page-break-inside:auto}.hp-tbl th,.hp-tbl td{border:0.3mm solid #888;padding:1.5mm 2.5mm;vertical-align:top;text-align:left}.hp-tbl th{background:#e6ecf5}tr{page-break-inside:avoid}
</style></head><body>${body}</body></html>`;
}
async function buildPdf(out, fontDir) {
  const { chromium } = require('playwright'); const html = printHtml(fontDir), tmp = path.join(__dirname, 'guide_print.html'); fs.writeFileSync(tmp, html);
  const b = await chromium.launch(), page = await b.newPage(); await page.goto('file://' + tmp); await page.evaluate(() => document.fonts.ready); await page.waitForTimeout(800);
  await page.pdf({ path: out, format: 'A4', printBackground: true, displayHeaderFooter: true, margin: { top: '16mm', bottom: '18mm', left: '14mm', right: '14mm' },
    headerTemplate: '<div></div>', footerTemplate: '<div style="width:100%;font-size:8px;font-family:Arial,sans-serif;color:#666;text-align:center">SGS School Management System &mdash; User Guide V19.64 &mdash; Page <span class="pageNumber"></span> / <span class="totalPages"></span></div>' });
  await b.close(); console.log('PDF तयार: ' + out);
}
const a = process.argv.slice(2);
if (a.includes('--pdf')) { const out = a[a.indexOf('--pdf') + 1], fd = a.includes('--font-dir') ? a[a.indexOf('--font-dir') + 1] : path.join(__dirname, 'fonts'); buildPdf(out, fd).catch(e => { console.error(e); process.exit(1); }); }
else buildApp();
