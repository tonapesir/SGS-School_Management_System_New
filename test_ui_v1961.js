// V19.61 UI चाचण्या — अ-४ (त्रुटी-नोंद: क्लायंट) आणि अ-५ (Pending Saves पाहणे). बनावट backend (jsonpRequest/smartSaveRaw बदलून); खऱ्या सर्व्हर/फोनवर नाही.
// वापर: SGS_WEB=<index.html चा फोल्डर> node tests/test_ui_v1961.js
const path = require('path');
const { chromium } = require('playwright');
const WEB = process.env.SGS_WEB || path.join(__dirname, '..', 'github_pages');
let fails = 0, passes = 0;
function ok(c, m) { if (c) passes++; else { fails++; console.log('  FAIL: ' + m); } }
function eq(a, b, m) { ok(JSON.stringify(a) === JSON.stringify(b), m + ' → मिळाले ' + JSON.stringify(a) + ' अपेक्षित ' + JSON.stringify(b)); }

async function boot(browser, user) {
  const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, acceptDownloads: true });
  const page = await ctx.newPage(); const pageErrors = [];
  page.on('pageerror', e => pageErrors.push(e.message));
  await page.route(/^(?!file:)/, r => r.abort());
  await page.goto('file://' + path.join(WEB, 'index.html')); await page.waitForTimeout(300);
  await page.evaluate(u => {
    window.__reqs = []; window.__resp = function () { return { status: 'ok', saved: 1 }; };
    window.jsonpRequest = function (params, cb) { window.__reqs.push(JSON.parse(JSON.stringify(params))); var r = window.__resp(params); if (r && r.__throw) throw new Error('stub throw'); setTimeout(function () { cb(r); }, 0); };
    var ls = document.getElementById('loginScreen'); if (ls) ls.style.display = 'none';
    window.currentUser = Object.assign({ username: 'u1', role: 'teacher', label: 'T', assignedClass: '5th A', token: 'tok', exp: Date.now() + 3600e3 }, u || {});
    currentUser = window.currentUser; document.body.classList.remove('auth-locked'); applyRoleUI();
    localStorage.removeItem('sgs_errq_v1'); localStorage.removeItem('sgs_pending_saves_v1');
  }, user);
  return { ctx, page, pageErrors };
}
const sleep = (p, ms) => p.waitForTimeout(ms);
const errReqs = p => p.evaluate(() => window.__reqs.filter(r => r.action === 'logClientErrors'));
const queue = p => p.evaluate(() => JSON.parse(localStorage.getItem('sgs_errq_v1') || '[]'));
const T = {};

// ---------------- अ-४ ----------------
T.capture_and_send_scrubbed = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { cerrCapture('error', 'Boom for 9822011111 mail a@b.com regNo=R101 at https://script.google.com/macros/s/AKfyc123/exec?token=SECRET&regNo=R9', 'Error: x\n at f (https://x.github.io/script.js?v=1:10:5)\n    at 123456789012'); });
  let q = await queue(page); eq(q.length, 1, 'रांगेत १ नोंद'); const blob = JSON.stringify(q);
  ok(!/9822011111|a@b\.com|R101|SECRET|123456789012|AKfyc123/.test(blob), 'scrub: संवेदनशील मजकूर रांगेतही नाही — ' + blob);
  ok(q[0].by === 'u1' && q[0].c === 1 && q[0].ua.length <= 60 && q[0].s.length <= 500 && q[0].m.length <= 300, 'by/c/ua/stack/msg लांबी');
  await page.evaluate(() => cerrFlush()); await sleep(page, 150);
  const r = await errReqs(page); eq(r.length, 1, 'एक logClientErrors विनंती'); const sent = JSON.parse(r[0].errors); eq(sent.length, 1, 'batch १');
  ok(!/9822011111|R101|SECRET/.test(r[0].errors), 'पाठवलेल्या मजकुरातही scrub'); eq((await queue(page)).length, 0, 'यशस्वी पाठवल्यावर रांग रिकामी');
  await ctx.close();
};
T.window_error_and_promise_hooks = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { setTimeout(function () { throw new Error('async fail A'); }, 0); Promise.reject(new Error('promise fail B')); });
  await sleep(page, 300); const q = await queue(page);
  ok(q.some(x => x.t === 'error' && /async fail A/.test(x.m)), 'window error पकडला'); ok(q.some(x => x.t === 'promise' && /promise fail B/.test(x.m)), 'unhandledrejection पकडला');
  await ctx.close();
};
T.dedupe_rate_session_limits = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { for (let i = 0; i < 20; i++) cerrCapture('error', 'same message', 'at f'); });
  let q = await queue(page); eq(q.length, 1, 'समान संदेश = १ नोंद'); eq(q[0].c, 20, 'संख्या २०');
  await page.evaluate(() => { for (let i = 0; i < 12; i++) cerrCapture('error', 'distinct ' + i, 'at g' + i); });
  q = await queue(page); eq(q.length, 5, 'प्रति मिनिट मर्यादा (५ नवीन प्रकार)'); eq(await page.evaluate(() => _cerr.dropped), 8, 'मर्यादेमुळे सोडलेले ८ मोजले');
  // सत्र-मर्यादा: मिनिट-खिडकी पुढे सरकवून
  await page.evaluate(() => { for (let k = 0; k < 10; k++) { _cerr.minStart = 0; for (let i = 0; i < 5; i++) cerrCapture('error', 'bulk' + k + '_' + i, 'at h'); } });
  eq(await page.evaluate(() => _cerr.total), 30, 'प्रति सत्र कमाल ३०');
  ok((await queue(page)).length <= 50, 'रांग ≤ ५०');
  await ctx.close();
};
T.queue_cap_50 = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { var q = []; for (let i = 0; i < 80; i++) q.push({ id: 'x' + i, by: 'u1', t: 'error', m: 'm' + i, s: '', c: 1, v: '', pg: '', ua: '', ts: 1 }); cerrStore(q); cerrEnqueue({ id: 'new', t: 'error', m: 'n', s: '', c: 1, v: '', pg: '', ua: '', ts: 2 }); });
  const q = await queue(page); eq(q.length, 50, 'रांग कमाल ५०'); eq(q[q.length - 1].id, 'new', 'नवीन टिकते; जुनी जातात');
  await ctx.close();
};
T.offline_and_noToken_stay_queued = async (b) => {
  let h = await boot(b, { offline: true });
  await h.page.evaluate(() => { cerrCapture('error', 'e1', 's'); cerrFlush(); }); await sleep(h.page, 150);
  eq((await errReqs(h.page)).length, 0, 'ऑफलाइन सत्रात पाठवत नाही'); eq((await queue(h.page)).length, 1, 'रांगेत राहते'); await h.ctx.close();
  h = await boot(b, { token: '' });
  await h.page.evaluate(() => { cerrCapture('error', 'e2', 's'); cerrFlush(); }); await sleep(h.page, 150);
  eq((await errReqs(h.page)).length, 0, 'Token नसताना पाठवत नाही'); eq((await queue(h.page)).length, 1, 'रांगेत राहते'); await h.ctx.close();
  h = await boot(b);
  await h.page.evaluate(() => { cerrCapture('error', 'e3', 's'); Object.defineProperty(navigator, 'onLine', { get: () => false, configurable: true }); cerrFlush(); }); await sleep(h.page, 150);
  eq((await errReqs(h.page)).length, 0, 'navigator.onLine=false असताना पाठवत नाही'); await h.ctx.close();
};
T.other_user_entries_not_sent = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { cerrStore([{ id: 'o1', by: 'someone_else', t: 'error', m: 'theirs', s: '', c: 1, v: '', pg: '', ua: '', ts: 1 }]); cerrCapture('error', 'mine', 's'); cerrFlush(); }); await sleep(page, 150);
  const r = await errReqs(page); eq(r.length, 1, '१ विनंती'); ok(!/theirs/.test(r[0].errors) && /mine/.test(r[0].errors), 'फक्त स्वतःच्या नोंदी पाठवल्या');
  ok((await queue(page)).some(x => x.id === 'o1'), 'दुसऱ्याची नोंद रांगेत शिल्लक');
  await ctx.close();
};
T.failure_backoff_no_loop_no_hang = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { window.__resp = function () { return { status: 'timeout', message: 'Timeout' }; }; cerrCapture('error', 'e1', 's'); cerrFlush(); }); await sleep(page, 200);
  eq((await errReqs(page)).length, 1, 'पहिला प्रयत्न'); eq((await queue(page)).length, 1, 'अपयशानंतर नोंद रांगेत');
  await page.evaluate(() => { cerrFlush(); cerrFlush(); }); await sleep(page, 100); eq((await errReqs(page)).length, 1, 'backoff: लगेच पुन्हा प्रयत्न नाही');
  ok((await page.evaluate(() => _cerr.nextTry)) > Date.now(), 'nextTry भविष्यात');
  // पाठवणे स्वतः throw झाले तरी ॲप अडत नाही व नवीन त्रुटी रांगेत भरत नाही
  await page.evaluate(() => { _cerr.nextTry = 0; _cerr.sending = false; window.__resp = function () { return { __throw: true }; }; });
  await page.evaluate(() => { try { cerrFlush(); } catch (e) { window.__escaped = true; } }); await sleep(page, 100);
  ok(!(await page.evaluate(() => window.__escaped)), 'cerrFlush अपवाद बाहेर फेकत नाही'); eq(await page.evaluate(() => _cerr.sending), false, 'sending ध्वज अडकला नाही');
  const qn = (await queue(page)).length; ok(qn === 1, 'अनंत-फेरी नाही: रांगेत फक्त मूळ १ नोंद (' + qn + ')');
  // server 'busy' / 'needauth' = रांगेत राहते
  await page.evaluate(() => { _cerr.nextTry = 0; _cerr.sending = false; window.__resp = function () { return { status: 'error', code: 'busy' }; }; cerrFlush(); }); await sleep(page, 100);
  eq((await queue(page)).length, 1, 'busy = रांगेत राहते');
  // पुन्हा यश
  await page.evaluate(() => { _cerr.nextTry = 0; _cerr.sending = false; window.__resp = function () { return { status: 'ok', saved: 1 }; }; cerrFlush(); }); await sleep(page, 100);
  eq((await queue(page)).length, 0, 'यशानंतर रांग रिकामी');
  await ctx.close();
};
T.counters_timeouts_and_savefail = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { cerrNoteResult('getAll', { status: 'timeout' }); cerrNoteResult('getAll', { status: 'timeout' }); cerrNoteResult('logClientErrors', { status: 'timeout' }); cerrNoteResult('getAll', { status: 'ok' }); cerrCount('savefail', 'saveFees'); cerrFlush(); }); await sleep(page, 200);
  const r = await errReqs(page); eq(r.length, 1, 'एक विनंती'); const sent = JSON.parse(r[0].errors);
  const to = sent.find(x => x.m === 'timeout:getAll'), sf = sent.find(x => x.m === 'savefail:saveFees');
  ok(to && to.c === 2, 'timeout:getAll संख्या २ — ' + JSON.stringify(sent)); ok(sf && sf.c === 1, 'savefail संख्या १'); ok(!sent.some(x => /logClientErrors/.test(x.m)), 'त्रुटी-नोंदीचे स्वतःचे timeout मोजले जात नाहीत');
  await ctx.close();
};
T.batch_fits_url_budget = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { var q = []; for (let i = 0; i < 40; i++) q.push({ id: 'b' + i, by: 'u1', t: 'error', m: 'M'.repeat(300) + i, s: 'S'.repeat(500), c: 1, v: 'V19.61', pg: 'lc', ua: 'UA', ts: 1 }); cerrStore(q); cerrFlush(); }); await sleep(page, 150);
  const r = await errReqs(page); const len = encodeURIComponent(r[0].errors).length; ok(len <= 5400 && JSON.parse(r[0].errors).length >= 1, 'URL अंदाजपत्रकात (' + len + ')'); ok(JSON.parse(r[0].errors).length < 10, 'मोठ्या नोंदी = कमी प्रति batch');
  await ctx.close();
};
T.dq_errors_tab_super = async (b) => {
  const { ctx, page } = await boot(b, { username: 'boss', role: 'super', assignedClass: '' });
  await page.evaluate(() => { window.__resp = function (p) { if (p.action === 'getClientErrors') return { status: 'ok', total: 7, rowsRead: 5, data: [{ type: 'error', message: 'Boom <b>x</b>', stack: 'at f', count: 5, users: 2, first: '2026-11-01 10:00:00', last: '2026-11-02 11:00:00', vers: 'V19.60, V19.61', pages: 'lc, search' }, { type: 'count', message: 'timeout:getAll', stack: '', count: 2, users: 1, first: '2026-11-01 10:00:00', last: '2026-11-01 12:00:00', vers: 'V19.61', pages: '' }] }; return { status: 'ok' }; }; showPage('dataqc'); dqTab('e'); });
  await sleep(page, 300);
  const html = await page.evaluate(() => document.getElementById('dq_body').innerHTML); ok(/ce_tbl/.test(html) && /timeout:getAll/.test(html), 'तक्ता आला'); ok(!/<b>x<\/b>/.test(html) && /&lt;b&gt;x/.test(html), 'संदेश escape (XSS नाही)');
  eq(await page.evaluate(() => document.querySelectorAll('#ce_tbl tbody tr').length), 2, '२ ओळी');
  await page.fill('#arq_ce_tbl', 'timeout'); await sleep(page, 100); eq(await page.evaluate(() => [...document.querySelectorAll('#ce_tbl tbody tr')].filter(t => t.style.display !== 'none').length), 1, 'filter');
  const reqs = await page.evaluate(() => window.__reqs.filter(r => r.action === 'getClientErrors').length); eq(reqs, 1, 'getClientErrors एकदा');
  await ctx.close();
};

// ---------------- अ-५ ----------------
const NOW = () => Date.now();
const seedPending = (page, items) => page.evaluate(i => { localStorage.setItem('sgs_pending_saves_v1', JSON.stringify(i)); pendingUpdateBadge(); }, items);
const pend = p => p.evaluate(() => JSON.parse(localStorage.getItem('sgs_pending_saves_v1') || '[]'));
const mk = (action, extra, by, ageMs, more) => Object.assign({ data: Object.assign({ action, reqId: 'r_' + action + '_' + (extra.k || '') }, extra), queuedAt: NOW() - (ageMs || 60000), by }, more || {});
T.viewer_list_fields = async (b) => {
  const { ctx, page } = await boot(b);
  await seedPending(page, [
    mk('saveAttendance', { iyatta: '5th', tukdi: 'A', date: '2026-10-29' }, 'u1', 3600e3 * 2),
    mk('saveFees', { fullName: 'राम पाटील', regNo: 'R101' }, 'u1', 4 * 86400e3, { failed: true, lastError: 'अधिकार नाही', tries: 2 }),
    mk('saveNotice', { title: '<img src=x onerror="window.__xss=1">', k: 'n' }, 'other', 60e3),
  ]);
  await page.click('#pendingSavesBadge'); await sleep(page, 100);
  const txt = await page.evaluate(() => document.getElementById('pendingModal').innerText);
  ok(/हजेरी/.test(txt) && /29\/10\/2026/.test(txt), 'हजेरी — वर्ग — तारीख: ' + txt.slice(0, 120)); ok(/फी — राम पाटील/.test(txt), 'फी प्रकार'); ok(/👤 u1/.test(txt) && /👤 other/.test(txt), 'कोणाचे (by)');
  ok(/३ दिवसांपेक्षा जुनी/.test(txt), 'वय-इशारा (>३ दिवस)'); eq((txt.match(/३ दिवसांपेक्षा जुनी/g) || []).length, 1, 'फक्त जुन्या नोंदीला इशारा');
  ok(/अयशस्वी/.test(txt) && /अधिकार नाही/.test(txt) && /प्रयत्न: 2/.test(txt), 'स्थिती अयशस्वी + शेवटचा संदेश + प्रयत्न'); ok(/वाट पाहत/.test(txt), 'वाट पाहत स्थिती');
  ok(/other च्या Login नंतर जाईल/.test(txt), 'दुसऱ्या user ची नोंद लॉक');
  eq(await page.evaluate(() => document.querySelectorAll('#pendingModal .pv-cancel').length), 2, 'रद्द बटण फक्त स्वतःच्या २ नोंदींना');
  eq(await page.evaluate(() => !!window.__xss || !!document.querySelector('#pendingModal img')), false, 'नोटीस शीर्षक escape (XSS नाही)');
  ok(/Pending Saves \(3\)/.test(txt), 'शीर्षकात संख्या'); await ctx.close();
};
T.badge_shows_failed = async (b) => {
  const { ctx, page } = await boot(b);
  await seedPending(page, [mk('saveFees', { k: '1' }, 'u1', 1000, { failed: true }), mk('saveFees', { k: '2' }, 'u1', 1000)]);
  const t = await page.evaluate(() => document.getElementById('pendingSavesBadge').textContent); ok(/2/.test(t) && /1 अयशस्वी/.test(t), 'badge: ' + t);
  ok(await page.evaluate(() => document.getElementById('pendingSavesBadge').classList.contains('pv-badge-bad')), 'अयशस्वी असल्यास लाल');
  await ctx.close();
};
T.failed_not_deleted_and_not_autoresent = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { window.__sent = []; window.smartSaveRaw = function (d, cb) { window.__sent.push(d.reqId); cb(window.__srv(d)); }; window.__srv = () => ({ status: 'error', message: 'अधिकार नाही' }); });
  await seedPending(page, [mk('saveFees', { k: '1' }, 'u1'), mk('saveDiary', { k: '2' }, 'u1')]);
  await page.evaluate(() => { window.__srv = d => d.action === 'saveFees' ? { status: 'error', message: 'अधिकार नाही' } : { status: 'ok' }; pendingFlush(); }); await sleep(page, 200);
  let l = await pend(page); eq(l.length, 1, 'यशस्वी गेली, अयशस्वी राहिली (पुसली नाही)'); ok(l[0].failed === true && l[0].lastError === 'अधिकार नाही' && l[0].tries === 1, 'failed/lastError/tries नोंद: ' + JSON.stringify(l[0]));
  const n1 = await page.evaluate(() => window.__sent.length); await page.evaluate(() => pendingFlush()); await sleep(page, 100); eq(await page.evaluate(() => window.__sent.length), n1, 'आपोआप flush अयशस्वी नोंद पुन्हा पाठवत नाही');
  await page.evaluate(() => { window.__srv = () => ({ status: 'ok' }); pendingFlush(true); }); await sleep(page, 200);
  eq((await pend(page)).length, 0, 'हाताने "आत्ता पाठवा" = अयशस्वीचा पुन्हा प्रयत्न → यश → नोंद गेली'); await ctx.close();
};
T.network_failure_keeps_item = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { window.smartSaveRaw = function (d, cb) { cb({ status: 'timeout', message: 'Timeout' }); }; });
  await seedPending(page, [mk('saveFees', { k: '1' }, 'u1')]); await page.evaluate(() => pendingFlush()); await sleep(page, 150);
  const l = await pend(page); eq(l.length, 1, 'नेट अपयश: नोंद राहते'); ok(!l[0].failed && /नेट/.test(l[0].lastError || ''), 'failed नाही; lastError नेट'); await ctx.close();
};
T.other_user_locked_not_sent = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { window.__sent = []; window.smartSaveRaw = function (d, cb) { window.__sent.push(d.reqId); cb({ status: 'ok' }); }; });
  await seedPending(page, [mk('saveFees', { k: 'o' }, 'other'), mk('saveFees', { k: 'm' }, 'u1')]); await page.evaluate(() => pendingFlush(true)); await sleep(page, 200);
  eq(await page.evaluate(() => window.__sent), ['r_saveFees_m'], 'फक्त स्वतःची नोंद गेली'); const l = await pend(page); ok(l.length === 1 && l[0].by === 'other', 'दुसऱ्याची नोंद शिल्लक'); await ctx.close();
};
T.viewer_cancel_one = async (b) => {
  const { ctx, page } = await boot(b); let dialogs = [];
  page.on('dialog', d => { dialogs.push(d.message()); d.dismiss(); });
  await seedPending(page, [mk('saveFees', { k: '1' }, 'u1'), mk('saveDiary', { k: '2' }, 'u1')]); await page.evaluate(() => pendingOpenViewer());
  await page.click('#pendingModal .pv-cancel >> nth=0'); await sleep(page, 100); eq((await pend(page)).length, 2, 'पुष्टी नाकारल्यास नोंद राहते'); ok(dialogs.length === 1 && /JSON/.test(dialogs[0]), 'पुष्टीत JSON जतन करण्याची सूचना');
  page.removeAllListeners('dialog'); page.on('dialog', d => d.accept());
  await page.click('#pendingModal .pv-cancel >> nth=0'); await sleep(page, 100); const l = await pend(page); eq(l.length, 1, 'पुष्टीनंतर एक नोंद रद्द'); eq(l[0].data.action, 'saveDiary', 'योग्य नोंद रद्द झाली'); await ctx.close();
};
T.viewer_clear_double_confirm_only_mine = async (b) => {
  const { ctx, page } = await boot(b); let n = 0, answers = [false];
  page.on('dialog', d => { n++; (answers[n - 1] === undefined ? answers[answers.length - 1] : answers[n - 1]) ? d.accept() : d.dismiss(); });
  await seedPending(page, [mk('saveFees', { k: '1' }, 'u1'), mk('saveDiary', { k: '2' }, 'u1'), mk('saveFees', { k: 'o' }, 'other')]); await page.evaluate(() => pendingOpenViewer());
  answers = [false]; n = 0; await page.click('#pv_clear'); await sleep(page, 100); eq((await pend(page)).length, 3, 'पहिली पुष्टी नाकारली = काहीही पुसले नाही'); eq(n, 1, '१ dialog');
  answers = [true, false]; n = 0; await page.click('#pv_clear'); await sleep(page, 100); eq((await pend(page)).length, 3, 'दुसरी पुष्टी नाकारली = काहीही पुसले नाही'); eq(n, 2, '२ dialogs');
  answers = [true, true]; n = 0; await page.click('#pv_clear'); await sleep(page, 100); const l = await pend(page); ok(l.length === 1 && l[0].by === 'other', 'दुहेरी पुष्टीनंतर फक्त स्वतःच्या नोंदी पुसल्या; दुसऱ्याची राहिली'); await ctx.close();
};
T.viewer_export_json = async (b) => {
  const { ctx, page } = await boot(b);
  await seedPending(page, [mk('saveFees', { k: '1', fullName: 'राम' }, 'u1'), mk('saveFees', { k: 'o' }, 'other')]); await page.evaluate(() => pendingOpenViewer());
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#pv_export')]);
  const name = dl.suggestedFilename(); ok(/^pending_saves_\d{8}_\d{4}\.json$/.test(name), 'फाइलनाव ' + name);
  const txt = require('fs').readFileSync(await dl.path(), 'utf8'); const j = JSON.parse(txt);
  eq(j.items.length, 2, 'सर्व नोंदी JSON मध्ये'); ok(!/"token"/.test(txt), 'Token JSON मध्ये नाही'); ok(j.items[0].data.fullName === 'राम' && j.by === 'u1', 'डेटा पूर्ण');
  eq((await pend(page)).length, 2, 'export मुळे नोंदी पुसल्या जात नाहीत'); await ctx.close();
};
T.viewer_retry_button = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { window.__sent = []; window.smartSaveRaw = function (d, cb) { window.__sent.push(d.reqId); cb({ status: 'ok' }); }; });
  await seedPending(page, [mk('saveFees', { k: '1' }, 'u1', 1000, { failed: true, lastError: 'x' })]); await page.evaluate(() => pendingOpenViewer());
  await page.click('#pv_retry'); await sleep(page, 700); eq((await pend(page)).length, 0, 'पुन्हा पाठवा → गेली'); ok(await page.evaluate(() => /कोणतीही Pending नोंद नाही/.test(document.getElementById('pendingModal').innerText)), 'यादी रिकामी संदेश'); await ctx.close();
};
T.viewer_offline_session_hint = async (b) => {
  const { ctx, page } = await boot(b, { offline: true });
  await page.evaluate(() => { window.__sent = []; window.smartSaveRaw = function (d, cb) { window.__sent.push(1); cb({ status: 'ok' }); }; });
  await seedPending(page, [mk('saveFees', { k: '1' }, 'u1')]); await page.evaluate(() => pendingOpenViewer());
  ok(await page.evaluate(() => /पुन्हा Login करा/.test(document.getElementById('pendingModal').innerText)), '"पुन्हा Login करा" सूचना');
  await page.click('#pv_retry'); await sleep(page, 300); eq(await page.evaluate(() => window.__sent.length), 0, 'ऑफलाइन सत्रात पाठवत नाही'); eq((await pend(page)).length, 1, 'नोंद राहते'); await ctx.close();
};
T.reqId_preserved = async (b) => {
  const { ctx, page } = await boot(b);
  await page.evaluate(() => { window.__sent = []; window.smartSaveRaw = function (d, cb) { window.__sent.push(d.reqId); cb({ status: 'error', message: 'no' }); }; });
  await seedPending(page, [mk('saveFees', { k: 'Z' }, 'u1')]); await page.evaluate(() => pendingFlush(true)); await sleep(page, 150); await page.evaluate(() => pendingFlush(true)); await sleep(page, 150);
  eq(await page.evaluate(() => window.__sent), ['r_saveFees_Z', 'r_saveFees_Z'], 'पुन्हा पाठवताना तोच reqId (सर्व्हरवर idempotency)'); await ctx.close();
};

(async () => {
  const browser = await chromium.launch();
  for (const k of Object.keys(T)) { const before = fails; try { await T[k](browser); } catch (e) { fails++; console.log('  EXCEPTION in ' + k + ': ' + String(e && e.message || e).split('\n').slice(0, 2).join(' | ')); } console.log((fails === before ? 'ok   ' : 'FAIL ') + k); }
  await browser.close();
  console.log(`\n${fails ? 'FAIL' : 'PASS'}: test_ui_v1961 (${passes} तपासण्या पास, ${fails} अपयशी)`);
  process.exit(fails ? 1 : 0);
})();
