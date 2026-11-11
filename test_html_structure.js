// index.html की रचना-तपासणी: <div> संतुलन, सर्व .page एकाच पालकाखाली, .dcard एकमेकांत नाहीत व सर्व एकाच .dash मध्ये, डॅशबोर्ड-कार्ड (widgets) .dash च्या बाहेर.
// (V19.64 मध्ये "मदत" कार्ड "डेटा तपासणी" कार्डाच्या आत घुसले होते — हा वर्ग-दोष आता पकडला जातो.)
// वापर: SGS_WEB=<index.html चा फोल्डर> [SGS_INDEX=<file>] node tests/test_html_structure.js
const path = require('path'), fs = require('fs'); const { chromium } = require('playwright');
const FILE = process.env.SGS_INDEX || path.join(process.env.SGS_WEB || path.join(__dirname, '..', 'github_pages'), 'index.html');
let fails = 0, passes = 0; function ok(c, m) { if (c) passes++; else { fails++; console.log('  FAIL: ' + m); } }
// १) टॅग-संतुलन (script/style/comment/template वगळून नाही — template आतील HTML देखील संतुलित असावे)
function balance(html) {
  const VOID = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr']);
  const cleaned = html.replace(/<!--[\s\S]*?-->/g, m => m.replace(/[^\n]/g, ' ')).replace(/<script\b[\s\S]*?<\/script>/gi, m => m.replace(/[^\n]/g, ' ')).replace(/<style\b[\s\S]*?<\/style>/gi, m => m.replace(/[^\n]/g, ' '));
  /* quoted attribute मूल्यांतील '>' / '<' टॅग मानले जात नाहीत */
  const re = /<\/?([a-zA-Z][a-zA-Z0-9-]*)(?:\s+(?:[^\s"'>\/=]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s"'>]+))?|\/))*\s*(\/?)>/g; const stack = []; const errs = []; let m;
  const line = i => cleaned.slice(0, i).split('\n').length;
  while ((m = re.exec(cleaned))) {
    const tag = m[1].toLowerCase(), closing = m[0][1] === '/'; if (VOID.has(tag) || m[2] === '/') continue;
    if (!closing) stack.push({ tag, line: line(m.index) });
    else { const top = stack.pop(); if (!top || top.tag !== tag) { errs.push(`</${tag}> (ओळ ${line(m.index)}) ला जुळणारा उघडा टॅग नाही; अपेक्षित: ${top ? '</' + top.tag + '> (उघडला ओळ ' + top.line + ')' : 'काही नाही'}`); if (top && stack.length) {} break; } }
  }
  if (!errs.length && stack.length) errs.push('बंद न केलेले टॅग: ' + stack.slice(-3).map(s => `<${s.tag}> ओळ ${s.line}`).join(', '));
  return errs;
}
(async () => {
  const html = fs.readFileSync(FILE, 'utf8'); const be = balance(html); ok(be.length === 0, 'टॅग-संतुलन: ' + be.join(' | '));
  const b = await chromium.launch(); const page = await b.newPage({ viewport: { width: 1080, height: 900 } }); await page.route(/^(?!file:)/, r => r.abort()); await page.goto('file://' + FILE); await page.waitForTimeout(300);
  const r = await page.evaluate(() => {
    const pages = [...document.querySelectorAll('.page')], parents = new Set(pages.map(p => p.parentElement && (p.parentElement.id || p.parentElement.tagName)));
    const dcards = [...document.querySelectorAll('.dcard')], dparents = new Set(dcards.map(d => d.parentElement.className)), nested = dcards.filter(d => d.parentElement.closest('.dcard')).map(d => d.dataset.page);
    const dash = document.querySelector('#pg-dashboard .dash'), widgetsInGrid = dash ? [...dash.children].filter(c => !c.classList.contains('dcard')).map(c => c.className.slice(0, 30)) : ['.dash नाही'];
    const widgets = [...document.querySelectorAll('#pg-dashboard > .card')].length;
    const dup = {}; document.querySelectorAll('[id]').forEach(e => { dup[e.id] = (dup[e.id] || 0) + 1; });
    return { pages: pages.length, pageParents: [...parents], dcards: dcards.length, dparents: [...dparents], nested, widgetsInGrid, widgets, dupIds: Object.keys(dup).filter(k => dup[k] > 1) };
  });
  ok(r.pages >= 20, '.page संख्या: ' + r.pages); ok(r.pageParents.length === 1, 'सर्व .page एकाच पालकाखाली (नाहीतर एक पान दुसऱ्यात घुसले): ' + JSON.stringify(r.pageParents));
  ok(r.dcards >= 19, '.dcard संख्या: ' + r.dcards); ok(r.nested.length === 0, '.dcard दुसऱ्या .dcard च्या आत नको: ' + r.nested.join(',')); ok(r.dparents.length === 1, 'सर्व .dcard एकाच पालकाखाली: ' + JSON.stringify(r.dparents));
  ok(r.widgetsInGrid.length === 0, 'डॅशबोर्ड-कार्ड (widgets) .dash ग्रिडमध्ये घुसलेले नाहीत: ' + r.widgetsInGrid.join(',')); ok(r.widgets >= 3, 'डॅशबोर्ड widgets पानावर थेट: ' + r.widgets); ok(r.dupIds.length === 0, 'डुप्लिकेट id: ' + r.dupIds.join(','));
  await b.close(); console.log(`\n${fails ? 'FAIL' : 'PASS'}: test_html_structure (${passes} तपासण्या पास, ${fails} अपयशी)`); process.exit(fails ? 1 : 0);
})();
