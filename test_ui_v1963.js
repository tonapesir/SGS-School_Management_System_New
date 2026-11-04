// V19.63 UI चाचण्या — ब-९ (🎓 वर्ष-अखेर मार्गदर्शक) आणि ब-१० (📋 रोजचा सारांश). बनावट backend (jsonpRequest/smartSave/smartSaveRaw बदलून); खऱ्या सर्व्हर/फोनवर नाही.
// वापर: SGS_WEB=<index.html चा फोल्डर> node tests/test_ui_v1963.js
const path = require('path');
const { chromium } = require('playwright');
const WEB = process.env.SGS_WEB || path.join(__dirname, '..', 'github_pages');
let fails = 0, passes = 0;
function ok(c, m) { if (c) passes++; else { fails++; console.log('  FAIL: ' + m); } }
function eq(a, b, m) { ok(JSON.stringify(a) === JSON.stringify(b), m + ' → मिळाले ' + JSON.stringify(a) + ' अपेक्षित ' + JSON.stringify(b)); }
const sleep = (p, ms) => p.waitForTimeout(ms);

// ---- बनावट वर्ष-अखेर सर्व्हर (पेजमध्ये): सर्व्हरच्या नियमांची प्रत — स्थिती, क्रम, पूर्वअटी ----
const FAKE_SERVER = `(function(){
  var S = window.__S = { closing: '2026-27', newYear: '2027-28', current: '2026-27', last: '10th', rec: {}, f: { backup: false, dq: 0, lcPending: 0, lastCount: 3, promo: 0, newyear: false, archived: false }, fail: {}, log: [], startedAt: null };
  function det(n) { var f = S.f; return n === 1 ? f.backup : n === 2 ? f.dq === 0 : n === 4 ? (f.lastCount === 0 || f.lcPending === 0) : n === 5 ? (f.lastCount === 0 || f.promo > 0) : n === 6 ? f.promo > 0 : n === 7 ? f.newyear : n === 8 ? f.archived : false; }
  function state() {
    var steps = [], active = 0, i;
    for (i = 1; i <= 9; i++) { var r = S.rec[i]; steps.push({ n: i, status: r ? r.s : (det(i) ? 'auto' : 'pending'), by: r ? r.by : (det(i) ? 'आपोआप' : ''), time: r ? r.t : null, reason: r ? r.r : '', info: {} }); }
    steps[0].info = { hoursAgo: S.f.backup ? 3 : null, fresh: S.f.backup }; steps[1].info = { issues: S.f.dq }; steps[2].info = { count: 12, total: 4500, lastClassCount: 2 };
    steps[3].info = { lastClass: '10th', lastCount: S.f.lastCount, lcPending: S.f.lcPending }; steps[4].info = { lastClass: '10th', count: S.f.lastCount, label: '2026-27' };
    steps[5].info = { promoted: S.f.promo }; steps[6].info = { newYear: '2027-28', cur: S.current }; steps[7].info = { year: '2026-27', startedAt: S.startedAt };
    for (i = 0; i < 9; i++) if (!/done|auto|skipped/.test(steps[i].status)) { active = i + 1; steps[i].status = 'active'; break; }
    if (active) for (i = active; i < 9; i++) if (!/done|auto|skipped/.test(steps[i].status)) steps[i].status = 'locked';
    var verify = S.f.newyear ? { curOk: true, current: S.current, newYear: '2027-28', noSnapshot: false, classes: [{ cls: '9th', before: 2, expected: 0, actual: 0, ok: true }, { cls: '10th', before: 3, expected: 2, actual: S.f.verifyActual === undefined ? 2 : S.f.verifyActual, ok: (S.f.verifyActual === undefined ? 2 : S.f.verifyActual) === 2 }], passedOut: { sheet: 'PassedOut_2026-27', expected: 3, actual: 3 }, notNewYear: 0, allOk: (S.f.verifyActual === undefined ? 2 : S.f.verifyActual) === 2, archived: S.f.archived } : null;
    return { status: 'ok', closing: S.closing, newYear: S.newYear, current: S.current, lastClass: S.last, active: active, complete: active === 0, steps: steps, verify: verify, hasSnapshot: true, _steps: steps };
  }
  function setStep(p) {
    var n = parseInt(p.step, 10), st = state();
    if (p.status === 'skipped' && String(p.reason || '').trim().length < 5) return { status: 'error', message: 'कारण (किमान ५ अक्षरे) लिहा' };
    for (var i = 1; i < n; i++) if (!/done|auto|skipped/.test(st.steps[i - 1].status)) return { status: 'error', message: 'आधी पायरी ' + i + ' पूर्ण करा' };
    if (/^(done|skipped)$/.test(st.steps[n - 1].status)) return { status: 'error', message: 'आधीच नोंदवलेली' };
    if (p.status === 'done') { var okd = n === 3 ? true : n === 9 ? (st.verify && st.verify.allOk) : det(n); if (!okd) return { status: 'error', message: 'पूर्वअट पूर्ण नाही' }; }
    S.rec[n] = { s: p.status, by: 'boss', t: new Date().toISOString(), r: p.reason || '' }; if (n === 7 && !S.startedAt) S.startedAt = S.rec[n].t; S.log.push(['set', n, p.status]);
    return state();
  }
  window.__serve = function (p) {
    S.log.push([p.action]);
    switch (p.action) {
      case 'getYearEndWizard': return state();
      case 'setYearEndStep': return setStep(p);
      case 'runBackup': if (S.fail.backup) return { status: 'error', message: 'Drive quota' }; S.f.backup = true; return { status: 'ok' };
      case 'getBackupStatus': return { status: 'ok', last: { ok: true, time: new Date().toISOString() } };
      case 'getYearEndCheck': return { status: 'ok', dues: [['R1', 'राम', '9th', 'A', 1, 500], ['R2', '<b>x</b>', '10th', 'A', 2, 300]], total: { dueSum: 800 } };
      case 'archivePassedOut': if (S.fail.passout) return { status: 'error', message: 'Backup अयशस्वी' }; S.f.lastCount = 0; return { status: 'ok', count: 3, extra: { batchId: 'B1' } };
      case 'previewPromotion': return { status: 'ok', classes: [{ iyatta: '9th', count: 2, suggestedNext: '10th' }, { iyatta: '10th', count: 0, suggestedNext: '' }, { iyatta: '5th', count: 1, suggestedNext: '6th' }] };
      case 'promoteStudents': if (S.fail.promo) return { status: 'error', message: 'Students लॉक' }; S.f.promo = 3; S.f.lastCount = 2; return { status: 'ok', count: 3 };
      case 'startNewYear': S.f.newyear = true; S.current = '2027-28'; return { status: 'ok', extra: { newYear: '2027-28', carried: 5 } };
      case 'archiveYear': S.f.archived = true; return { status: 'ok', extra: { moved: ['Fees_2026-27', 'Attendance_2026_06'], failed: [] } };
      case 'getArchiveHealth': return { status: 'ok', ok: true, problems: [], warnings: ['एक <script>window.__xssAH=1</script> इशारा'] };
      default: return { status: 'ok' };
    }
  };
  window.jsonpRequest = function (p, cb) { var r = window.__serve(JSON.parse(JSON.stringify(p))); setTimeout(function () { cb(r); }, 0); };
  window.smartSaveRaw = function (d, cb) { var r = window.__serve(JSON.parse(JSON.stringify(d))); setTimeout(function () { cb(r); }, 0); };
  window.smartSave = window.smartSaveRaw;
})()`;

async function boot(browser, user, preStub) {
  const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, permissions: ['clipboard-read', 'clipboard-write'] });
  const page = await ctx.newPage(); const errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route(/^(?!file:)/, r => r.abort());
  await page.goto('file://' + path.join(WEB, 'index.html')); await sleep(page, 300);
  if (preStub) await page.evaluate(preStub);
  await page.evaluate(u => {
    var ls = document.getElementById('loginScreen'); if (ls) ls.style.display = 'none';
    window.currentUser = Object.assign({ username: 'boss', role: 'super', label: 'B', assignedClass: '', token: 'tok', exp: Date.now() + 3600e3 }, u || {});
    currentUser = window.currentUser; document.body.classList.remove('auth-locked'); applyRoleUI();
    window.__dialogs = []; window.clearStudentCache = function () {};
  }, user);
  return { ctx, page, errors };
}
const serverLog = p => p.evaluate(() => window.__S.log);
const stepStatus = p => p.evaluate(() => window.__S && yeS.data ? yeS.data.steps.map(s => s.status) : []);
const openWizard = async (page) => { await page.evaluate(FAKE_SERVER); await page.evaluate(() => { showPage('dataqc'); dqTab('g'); }); await sleep(page, 300); };
const T = {};

// ---------------- ब-९ ----------------
T.wizard_initial_render_order_and_texts = async (b) => {
  const { ctx, page } = await boot(b); await openWizard(page);
  const st = await stepStatus(page); eq(st, ['active', 'auto', 'locked', 'auto', 'locked', 'locked', 'locked', 'locked', 'locked'], 'पायरी १ सक्रिय; २ व ४ आपोआप (त्रुटी/LC बाकी शून्य); बाकी लॉक');
  const txt = await page.evaluate(() => document.getElementById('dq_body').innerText);
  ok(/2026-27 → 2027-28/.test(txt) && /Pass-out आधी, बढती नंतर/.test(txt), 'शीर्षक व क्रम-सूचना'); YE9: for (let i = 1; i <= 9; i++) ok(new RegExp('पायरी ' + i + ' —').test(txt), 'पायरी ' + i + ' दिसते');
  ok(/काय होईल:/.test(txt) && /परत करता येते का:/.test(txt), '"काय होईल" व "परत करता येते का" मराठीत'); ok(/परत आणा/.test(txt) && /↩️ Undo/.test(txt) && /सहज परत करता येत नाही/.test(txt), 'परत-करण्याची माहिती (Pass-out/बढती/नवे वर्ष) — खऱ्या बटणांच्या नावासह');
  eq(await page.evaluate(() => document.querySelectorAll('#dq_body .ye-step').length), 9, '९ पायऱ्या');
  eq(await page.evaluate(() => [...document.querySelectorAll('#dq_body .ye-step')].map((e, i) => [i + 1, !!e.querySelector('button')]).filter(x => x[1]).map(x => x[0])), [1], 'फक्त सक्रिय पायरीला (१) बटणे; लॉक पायऱ्यांना नाहीत');
  await ctx.close();
};
T.wizard_flow_end_to_end = async (b) => {
  const { ctx, page } = await boot(b); await openWizard(page);
  page.on('dialog', d => d.accept());
  // १ Backup
  await page.evaluate(() => yeBackupNow()); await sleep(page, 300); ok((await stepStatus(page))[0] === 'auto', 'Backup नंतर पायरी १ पूर्ण'); eq((await stepStatus(page))[2], 'active', 'पुढे: पायरी ३ (२ आपोआप)');
  // ३ आढावा (CSV) व done
  const [dl] = await Promise.all([page.waitForEvent('download'), page.evaluate(() => yeDuesCsv())]); ok(/Thakbaki_2026-27\.csv/.test(dl.suggestedFilename()), 'थकबाकी CSV: ' + dl.suggestedFilename());
  const csv = require('fs').readFileSync(await dl.path(), 'utf8'); ok(/R1/.test(csv) && /500/.test(csv), 'CSV मध्ये डेटा');
  await page.evaluate(() => yeMarkDone(3)); await sleep(page, 250); eq((await stepStatus(page))[2], 'done', 'पायरी ३ आढावा घेतला');
  // ४ LC शून्य → आपोआप; ५ सक्रिय
  eq((await stepStatus(page))[3], 'auto', 'LC बाकी शून्य = आपोआप'); eq((await stepStatus(page))[4], 'active', 'पायरी ५ सक्रिय');
  // ५ चुकीची पुष्टी → काहीही होत नाही
  await page.evaluate(() => { document.getElementById('ye_c5').value = 'archive'; yePassOut(); }); await sleep(page, 150); ok(!(await serverLog(page)).some(x => x[0] === 'archivePassedOut'), 'चुकीचा शब्द (archive) = विनंती नाही'); ok(await page.evaluate(() => /ARCHIVE टाइप करा/.test(document.getElementById('ye_msg5').textContent)), 'सूचना दिसते');
  await page.evaluate(() => { document.getElementById('ye_c5').value = 'ARCHIVE'; yePassOut(); }); await sleep(page, 350);
  const logs = await serverLog(page); ok(logs.some(x => x[0] === 'archivePassedOut'), 'archivePassedOut बोलावले'); eq((await stepStatus(page))[4], 'done', 'पायरी ५ पूर्ण (नोंद)'); eq((await stepStatus(page))[5], 'active', 'पायरी ६ सक्रिय');
  // ६ बढती
  await sleep(page, 200); ok(await page.evaluate(() => /9th → 10th/.test(document.getElementById('ye_pv6').innerText) && !/10th → /.test(document.getElementById('ye_pv6').innerText)), 'आराखडा: 9th→10th, शेवटचा वर्ग वगळला');
  await page.evaluate(() => { document.getElementById('ye_c6').value = 'PROMOTE'; yePromote(); }); await sleep(page, 350);
  const pcall = await page.evaluate(() => window.__S.log.filter(x => x[0] === 'promoteStudents').length); eq(pcall, 1, 'promoteStudents एकदा'); eq((await stepStatus(page))[5], 'done', 'पायरी ६ पूर्ण');
  // ७ नवे वर्ष
  await page.evaluate(() => { document.getElementById('ye_c7').value = 'NEWYEAR'; yeNewYear(); }); await sleep(page, 400); eq((await stepStatus(page))[6], 'done', 'पायरी ७ पूर्ण'); ok(await page.evaluate(() => !!window.__S.startedAt), 'वेळ नोंद (पायरी ८ च्या इशाऱ्यासाठी)');
  // ८ — २ दिवसांपेक्षा कमी → अतिरिक्त पुष्टी
  const dialogs = []; page.removeAllListeners('dialog'); page.on('dialog', d => { dialogs.push(d.message()); d.accept(); });
  ok(await page.evaluate(() => /२–३ दिवस तपासणी/.test(document.getElementById('dq_body').innerText)), 'पायरी ८: "२–३ दिवस तपासणी" इशारा');
  await page.evaluate(() => { document.getElementById('ye_c8').value = 'ARCHIVEYEAR'; yeArchiveYear(); }); await sleep(page, 400);
  ok(dialogs.some(d => /फक्त .* दिवस झाले/.test(d)), 'घाईचा अतिरिक्त इशारा-पुष्टी'); eq((await stepStatus(page))[7], 'done', 'पायरी ८ पूर्ण');
  // ९ पडताळणी
  const t9 = await page.evaluate(() => document.getElementById('dq_body').innerText); ok(/वर्ग\s+आधी\s+अपेक्षित\s+प्रत्यक्ष/.test(t9.replace(/\n/g, ' ').replace(/\s+/g, ' ')) || /अपेक्षित/.test(t9), 'पडताळणी तक्ता'); ok(/PassedOut_2026-27/.test(t9) && /अपेक्षित 3, प्रत्यक्ष 3/.test(t9), 'Pass-out संख्या');
  await page.evaluate(() => yeArchiveHealth()); await sleep(page, 200); ok(await page.evaluate(() => /Archive 2026-27: ✅ ठीक/.test(document.getElementById('ye_ah').innerText)), 'Archive आरोग्य तपासणी'); eq(await page.evaluate(() => !!window.__xssAH || !!document.querySelector('#ye_ah script')), false, 'आरोग्य इशारा escape (XSS नाही)');
  await page.evaluate(() => yeMarkDone(9)); await sleep(page, 250); ok(await page.evaluate(() => /पूर्ण झाले आहे/.test(document.getElementById('dq_body').innerText)), 'wizard पूर्ण संदेश');
  eq(await page.evaluate(() => window.__S.log.filter(x => x[0] === 'setYearEndStep' || x[0] === 'set').length > 0), true, 'स्थिती सर्व्हरवर नोंदली');
  await ctx.close();
};
T.wizard_failed_step_stops = async (b) => {
  const { ctx, page } = await boot(b); await openWizard(page); page.on('dialog', d => d.accept());
  await page.evaluate(() => { __S.f.backup = true; __S.rec[3] = { s: 'done', by: 'boss', t: new Date().toISOString(), r: '' }; __S.fail.passout = true; yeRefresh(); }); await sleep(page, 300);
  eq((await stepStatus(page))[4], 'active', 'पायरी ५ सक्रिय');
  await page.evaluate(() => { document.getElementById('ye_c5').value = 'ARCHIVE'; yePassOut(); }); await sleep(page, 300);
  ok(await page.evaluate(() => /❌ .*Backup अयशस्वी.*पुढे जात नाही/.test(document.getElementById('ye_msg5').textContent)), 'त्रुटी संदेश + "पुढे जात नाही"');
  ok(!(await serverLog(page)).some(x => x[0] === 'set' && x[1] === 5), 'अयशस्वी = पायरी ५ ची नोंद नाही'); eq((await stepStatus(page))[4], 'active', 'पायरी ५ अजून सक्रिय'); eq((await stepStatus(page))[5], 'locked', 'पायरी ६ लॉकच');
  // बढतीही अयशस्वी: पायरी ६ पूर्ण होत नाही
  await page.evaluate(() => { __S.fail.passout = false; __S.fail.promo = true; document.getElementById('ye_c5').value = 'ARCHIVE'; yePassOut(); }); await sleep(page, 400);
  await page.evaluate(() => { document.getElementById('ye_c6').value = 'PROMOTE'; yePromote(); }); await sleep(page, 300);
  eq((await stepStatus(page))[5], 'active', 'बढती अयशस्वी = पायरी ६ सक्रियच'); ok(await page.evaluate(() => /Students लॉक/.test(document.getElementById('ye_msg6').textContent)), 'सर्व्हर संदेश दिसतो');
  await ctx.close();
};
T.wizard_skip_with_reason_and_state_after_refresh = async (b) => {
  const { ctx, page } = await boot(b); await openWizard(page);
  let promptAnswers = ['abc']; page.on('dialog', d => { if (d.type() === 'prompt') d.accept(promptAnswers.shift() || ''); else d.accept(); });
  await page.evaluate(() => yeSkip(1)); await sleep(page, 250); eq((await stepStatus(page))[0], 'active', 'छोटे कारण (३ अक्षरे) = वगळले जात नाही'); ok(await page.evaluate(() => /किमान ५ अक्षरे/.test(document.getElementById('ye_msg1').textContent)), 'कारण-त्रुटी दिसते');
  promptAnswers = ['Backup हाताने घेतला <b>x</b>']; await page.evaluate(() => yeSkip(1)); await sleep(page, 300);
  eq((await stepStatus(page))[0], 'skipped', 'कारणासह वगळले'); let txt = await page.evaluate(() => document.getElementById('dq_body').innerText); ok(/वगळले\/स्वीकारले — कारण: Backup हाताने घेतला <b>x<\/b>/.test(txt), 'कारण दिसते (escape केलेले)'); eq(await page.evaluate(() => !!document.querySelector('#dq_body .ye-state b')), false, 'कारणातील HTML चालत नाही');
  // पान Refresh सिम्युलेट: नवा लोड → स्थिती सर्व्हरवरून तीच
  const before = await stepStatus(page); await page.evaluate(() => { yeS.data = null; document.getElementById('dq_body').innerHTML = ''; dqTab('g'); }); await sleep(page, 300);
  eq(await stepStatus(page), before, 'Refresh नंतर स्थिती तीच (सर्व्हर-आधारित)'); eq((await stepStatus(page))[2], 'active', 'पायरी १ वगळल्यावर २ आपोआप, ३ सक्रिय');
  // क्रम मोडणे: सर्व्हर नाकारतो (UI बटणे नसली तरी थेट कॉल)
  const r = await page.evaluate(() => new Promise(res => yeReq({ action: 'setYearEndStep', step: 6, status: 'skipped', reason: 'घाईत वगळतो' }, res))); ok(r.status === 'error' && /आधी पायरी/.test(r.message), 'क्रम मोडल्यास सर्व्हर नाकारतो');
  await ctx.close();
};
T.wizard_verify_mismatch_and_accept = async (b) => {
  const { ctx, page } = await boot(b); await openWizard(page); let ans = ['फरक तपासून स्वीकारला'];
  page.on('dialog', d => { if (d.type() === 'prompt') d.accept(ans.shift() || ''); else d.accept(); });
  await page.evaluate(() => { var now = new Date(Date.now() - 3 * 86400000).toISOString(); [1, 3, 5, 6, 7, 8].forEach(n => { __S.rec[n] = { s: 'done', by: 'boss', t: now, r: '' }; }); __S.f.backup = true; __S.f.lastCount = 0; __S.f.promo = 2; __S.f.newyear = true; __S.current = '2027-28'; __S.f.archived = true; __S.f.verifyActual = 5; yeRefresh(); }); await sleep(page, 350);
  eq((await stepStatus(page))[8], 'active', 'पायरी ९ सक्रिय'); const t = await page.evaluate(() => document.getElementById('dq_body').innerText); ok(/❌/.test(t), 'जुळत नसलेल्या वर्गावर ❌');
  eq(await page.evaluate(() => [...document.querySelectorAll('#dq_body .ye-step')[8].querySelectorAll('button')].map(x => x.innerText.trim()).filter(x => /पूर्ण करा/.test(x)).length), 0, '"सर्व ठीक — पूर्ण करा" बटण नाही (जुळत नाही)');
  await page.evaluate(() => yeSkip(9)); await sleep(page, 300); ok(await page.evaluate(() => /पूर्ण झाले आहे/.test(document.getElementById('dq_body').innerText)), 'कारणासह स्वीकारून पूर्ण'); await ctx.close();
};
T.wizard_not_for_non_super = async (b) => {
  const { ctx, page } = await boot(b, { username: 'm', role: 'master' }); eq(await page.evaluate(() => canUsePage('dataqc')), false, 'Master ला डेटा तपासणी पान (व wizard) नाही');
  const { ctx: c2, page: p2 } = await boot(b, { username: 't', role: 'teacher', assignedClass: '5th A' }); eq(await p2.evaluate(() => canUsePage('dataqc')), false, 'शिक्षकाला नाही'); await ctx.close(); await c2.close();
};

// ---------------- ब-१० ----------------
const SUM = (extra) => Object.assign({
  status: 'ok', date: '2026-11-02', weekday: 1, isSunday: false, today: false, future: false, dayType: 'normal',
  attendance: { hasMarks: true, filled: 2, classes: [{ iyatta: '8th', tukdi: 'A', total: 3, absent: 1, present: 2, pct: 66.7, by: 'tA', time: '9:10' }, { iyatta: '9th', tukdi: 'A', total: 2, absent: 0, present: 2, pct: 100, by: 'tC', time: '9:20' }], totals: { total: 5, absent: 1, present: 4, pct: 80 }, pending: [{ iyatta: '8th', tukdi: 'B', teacher: 'श्री. ब <i>x</i>', mobile: '9822222222' }, { iyatta: '7th', tukdi: 'A', teacher: 'सौ. क', mobile: '' }], noTeacher: [{ iyatta: '6th', tukdi: 'A' }] },
  earlyLeave: { permitted: 4, gone: 2, notGone: 1, other: 1, total: 5 }, fees: { total: 900, count: 4, byClass: [{ cls: '8th', amount: 600 }, { cls: '9th', amount: 300 }], year: '2026-27' },
  notices: [{ title: 'सुट्टी & सूचना #1', postedBy: 'boss', targetClass: '' }, { title: 'पालक सभा', postedBy: 'mas', targetClass: '8th' }], changes: { total: 5, students: 3, byAction: [{ action: 'saveStudentContact', count: 2 }] }
}, extra || {});
const openDash = async (page, data) => { await page.evaluate(d => { window.__reqs = []; window.__sum = d; window.jsonpRequest = function (p, cb) { window.__reqs.push(JSON.parse(JSON.stringify(p))); setTimeout(function () { cb(p.action === 'getDailySummary' ? window.__sum : { status: 'ok' }); }, 0); }; document.getElementById('ds_date').value = d.date; dsLoad(); }, data); await sleep(page, 250); };
T.summary_message_and_render = async (b) => {
  const { ctx, page } = await boot(b); await openDash(page, SUM());
  const msg = await page.evaluate(() => document.getElementById('ds_msg').value);
  eq(msg.split('\n'), ['📋 श्री. गोविंदराव सेक्सरिया हायस्कूल, पाचोरा — 02/11/2026', 'उपस्थिती: 4/5 (80%) — गैरहजर 1', 'हजेरी बाकी (2): 8 वी B (श्री. ब <i>x</i>), 7 वी A (सौ. क)', 'घरी गेले: 2 (परवानगी 4, न गेलेले 1)', 'फी जमा: ₹900 (4 नोंदी)', 'नोटीस: 2 — सुट्टी & सूचना #1, पालक सभा', 'विद्यार्थी-बदल: 5 (3 विद्यार्थी)'], 'मराठी संदेश (स्वरूप)');
  const txt = await page.evaluate(() => document.getElementById('ds_body').innerText); ok(/4\/5/.test(txt) && /गैरहजर/.test(txt) && /हजेरी बाकी/.test(txt) && /₹900/.test(txt), 'कार्डवर आकडे'); ok(/8 वी A/.test(txt) && /66\.7/.test(txt), 'वर्गनिहाय तक्ता'); ok(/वर्ग-शिक्षक नेमलेला नाही: 6 वी A/.test(txt), 'वर्ग-शिक्षक नसलेले वर्ग');
  eq(await page.evaluate(() => !!document.querySelector('#ds_body i') || !!document.querySelector('#ds_body b i')), false, 'शिक्षकाचे नाव escape (HTML चालत नाही)');
  eq(await page.evaluate(() => [...document.querySelectorAll('#ds_body a')].map(a => a.href)), ['https://wa.me/919822222222'], 'शिक्षकाला WhatsApp दुवा (मोबाईल असल्यास, 91 सह)');
  eq((await page.evaluate(() => window.__reqs.filter(r => r.action === 'getDailySummary')))[0].date, '2026-11-02', 'तारीख पाठवली'); await ctx.close();
};
T.summary_day_types_messages = async (b) => {
  const { ctx, page } = await boot(b);
  await openDash(page, { status: 'ok', date: '2026-11-01', dayType: 'sunday', message: 'रविवार — सुट्टी', attendance: { classes: [], totals: { total: 0, absent: 0, present: 0, pct: null }, pending: [], noTeacher: [] }, earlyLeave: { total: 0 }, fees: { total: 0, count: 0, byClass: [] }, notices: [], changes: { total: 0, students: 0 } });
  let msg = await page.evaluate(() => document.getElementById('ds_msg').value); ok(/ℹ️ रविवार — सुट्टी/.test(msg) && !/उपस्थिती:/.test(msg), 'रविवार संदेश (उपस्थिती ओळ नाही)'); ok(/फी जमा: ₹0/.test(msg), 'फी जमा ₹0 दाखवतो');
  await openDash(page, { status: 'ok', date: '2026-11-10', dayType: 'future', message: 'भविष्यातील तारीख — माहिती उपलब्ध नाही.' });
  msg = await page.evaluate(() => document.getElementById('ds_msg').value); eq(msg.split('\n').length, 2, 'भविष्यातील तारीख: २ ओळी'); ok(/भविष्यातील/.test(msg), 'संदेश'); ok(await page.evaluate(() => /भविष्यातील/.test(document.getElementById('ds_body').innerText) && !/उपस्थिती/.test(document.getElementById('ds_body').innerText)), 'कार्डवर फक्त संदेश');
  await openDash(page, SUM({ date: '2026-11-03', dayType: 'nodata_today', message: 'आज अजून कोणत्याही वर्गाची हजेरी भरलेली नाही.', attendance: { hasMarks: true, filled: 0, classes: [], totals: { total: 0, absent: 0, present: 0, pct: null }, pending: [{ iyatta: '8th', tukdi: 'A', teacher: 'सौ. अ', mobile: '' }], noTeacher: [] } }));
  msg = await page.evaluate(() => document.getElementById('ds_msg').value); ok(/आज अजून कोणत्याही वर्गाची हजेरी भरलेली नाही/.test(msg), 'आज हजेरी नाही'); await ctx.close();
};
T.summary_actions_copy_whatsapp_pdf = async (b) => {
  const { ctx, page } = await boot(b); await openDash(page, SUM());
  await page.evaluate(() => { window.__opened = []; window.open = function (u) { window.__opened.push(u); return null; }; }); await page.click('button:has-text("💬 WhatsApp")');
  const url = (await page.evaluate(() => window.__opened))[0]; ok(/^https:\/\/wa\.me\/\?text=/.test(url), 'wa.me दुवा'); const dec = decodeURIComponent(url.split('?text=')[1]);
  eq(dec, await page.evaluate(() => document.getElementById('ds_msg').value), 'संदेश बरोबर encode (& # आणि नवीन ओळ सह)'); ok(url.includes('%26') && url.includes('%23') && url.includes('%0A') && !/[&#]/.test(url.split('?text=')[1]), '& → %26, # → %23, नवीन ओळ → %0A');
  await page.click('button:has-text("📋 कॉपी")'); await sleep(page, 200); eq(await page.evaluate(() => navigator.clipboard.readText()), await page.evaluate(() => document.getElementById('ds_msg').value), 'कॉपी: क्लिपबोर्डवर तोच संदेश');
  // PDF: html2canvas/jsPDF बनावट (CDN बंद) — फाइलनाव व HTML तपासतो
  await page.evaluate(() => { window.__saved = null; window.html2canvas = function (el) { window.__pdfHtml = el.innerHTML; return Promise.resolve({ width: 800, height: 1100, toDataURL: function () { return 'data:image/jpeg;base64,AAAA'; } }); }; window.jspdf = { jsPDF: function () { return { addImage: function () {}, save: function (n) { window.__saved = n; } }; } }; });
  await page.click('#ds_pdfBtn'); await sleep(page, 700); eq(await page.evaluate(() => window.__saved), 'Daily_Summary_2026-11-02.pdf', 'PDF फाइलनाव'); const ph = await page.evaluate(() => window.__pdfHtml); ok(/रोजचा सारांश — 02\/11\/2026/.test(ph) && /8 वी A/.test(ph) && /₹900/.test(ph), 'PDF मजकूर'); ok(!/<i>x<\/i>/.test(ph) && /&lt;i&gt;x&lt;\/i&gt;/.test(ph), 'PDF मध्येही escape');
  ok(await page.evaluate(() => !document.getElementById('ds_pdfBtn').disabled), 'PDF नंतर बटण पुन्हा सक्षम'); await ctx.close();
};
T.summary_date_change_and_error = async (b) => {
  const { ctx, page } = await boot(b); await openDash(page, SUM());
  await page.evaluate(() => { document.getElementById('ds_date').value = '2026-10-15'; document.getElementById('ds_date').dispatchEvent(new Event('change')); }); await sleep(page, 250);
  const reqs = await page.evaluate(() => window.__reqs.filter(r => r.action === 'getDailySummary')); eq(reqs.map(r => r.date), ['2026-11-02', '2026-10-15'], 'तारीख बदलल्यावर नवी विनंती');
  await page.evaluate(() => { window.__sum = { status: 'error', message: 'अधिकार नाही' }; dsLoad(); }); await sleep(page, 200); ok(await page.evaluate(() => /❌/.test(document.getElementById('ds_body').innerText)), 'त्रुटी संदेश दिसतो'); await ctx.close();
};
T.summary_card_roles_and_dashboard_hook = async (b) => {
  const { ctx, page } = await boot(b, { username: 'm', role: 'master' }); ok(await page.evaluate(() => getComputedStyle(document.getElementById('ds_body').closest('.card')).display !== 'none'), 'Master ला कार्ड दिसते');
  const { ctx: c2, page: p2 } = await boot(b, { username: 't', role: 'teacher', assignedClass: '5th A' }); ok(await p2.evaluate(() => getComputedStyle(document.getElementById('ds_body').closest('.card')).display === 'none'), 'शिक्षकाला कार्ड लपलेले');
  const { ctx: c3, page: p3 } = await boot(b, { username: 'd', role: 'deo' }); ok(await p3.evaluate(() => getComputedStyle(document.getElementById('ds_body').closest('.card')).display === 'none'), 'DEO ला कार्ड लपलेले');
  const { ctx: c4, page: p4 } = await boot(b, { username: 'boss', role: 'super' }, () => { window.__reqs = []; window.jsonpRequest = function (p, cb) { window.__reqs.push(p.action); setTimeout(function () { cb({ status: 'ok' }); }, 0); }; }); await p4.evaluate(() => showPage('dashboard')); await sleep(p4, 300); const before = await p4.evaluate(() => window.__reqs.includes('getDailySummary')); await p4.evaluate(() => document.getElementById('ds_body').scrollIntoView()); await sleep(p4, 1200); ok(before === false, 'कार्ड दिसण्यापूर्वी लोड होत नाही (lazy)');
  ok(await p4.evaluate(() => window.__reqs.includes('getDailySummary')), 'कार्ड दिसू लागल्यावर सारांश आपोआप लोड (lazy)'); await ctx.close(); await c2.close(); await c3.close(); await c4.close();
};

(async () => {
  const browser = await chromium.launch();
  for (const k of Object.keys(T)) { const before = fails; try { await T[k](browser); } catch (e) { fails++; console.log('  EXCEPTION in ' + k + ': ' + String(e && e.message || e).split('\n').slice(0, 2).join(' | ')); } console.log((fails === before ? 'ok   ' : 'FAIL ') + k); }
  await browser.close();
  console.log(`\n${fails ? 'FAIL' : 'PASS'}: test_ui_v1963 (${passes} तपासण्या पास, ${fails} अपयशी)`);
  process.exit(fails ? 1 : 0);
})();
