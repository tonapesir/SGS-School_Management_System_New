// V19.63 सर्व्हर चाचण्या — ब-९ (वर्ष-अखेर मार्गदर्शक: फक्त स्थिती; विद्यमान action बदललेले नाहीत) आणि ब-१० (रोजचा सारांश). बनावट Apps Script; खऱ्या Sheet/Drive वर नाही.
// वापर: SGS_GS=<Code.gs> [SGS_GS_BASE=<V19.62 Code.gs>] node tests/test_server_v1963.js
const path = require('path');
const { makeEnv } = require('./lib/gas_mock');
const GS = process.env.SGS_GS || path.join(__dirname, '..', 'apps_script', 'Code.gs');
const GS_BASE = process.env.SGS_GS_BASE || '';
let fails = 0, passes = 0;
function ok(c, m) { if (c) passes++; else { fails++; console.log('  FAIL: ' + m); } }
function eq(a, b, m) { ok(JSON.stringify(a) === JSON.stringify(b), m + ' → मिळाले ' + JSON.stringify(a) + ' अपेक्षित ' + JSON.stringify(b)); }
const T = {};
const bookWrites = env => Object.values(env.books).reduce((a, b) => a + b.writes, 0);
const bookReads = env => Object.values(env.books).reduce((a, b) => a + b.reads, 0);
const sup = env => env.token('boss', 'super');
const stu = (reg, cls, tukdi, name, yr) => ['t', yr || '2026-27', reg, 'S' + reg, '', '', cls, tukdi, '', '1', name || ('नाव ' + reg), '', 'Male', '', '', '', '2012-05-04', '', '', '', '', '', '', '', '', '', '', '9822011111'];

// ---------------- ब-९ ----------------
function yeEnv(file) {
  const env = makeEnv(); env.load(file || GS); const c = env.ctx;
  env.seed('Students', c.MAIN_HEADERS, [stu('R1', '5th', 'A'), stu('R2', '9th', 'A'), stu('R3', '9th', 'A'), stu('R4', '10th', 'A'), stu('R5', '10th', 'A'), stu('R6', '10th', 'A')]);
  env.seed('Users', c.USER_HEADERS, [['boss', 'x', 'super', 'B', '', ''], ['mas', 'x', 'master', 'M', '', '']]);
  env.seed('LC', Array.from({ length: 19 }, (_, i) => 'c' + i), [Array.from({ length: 19 }, (_, i) => i === 2 ? 'R4' : '')]);   // फक्त R4 ची LC
  env.seed('PromotionLog', c.PROMO_LOG_HEADERS, []);
  env.props['CURRENT_AC_YEAR'] = '2026-27'; env.props['LAST_BACKUP'] = JSON.stringify({ ok: true, time: new Date().toISOString() });
  env.run('dqCompute_ = function(r, t) { return { byStudent: {} }; };');   // डेटा तपासणी: बाय-डीफॉल्ट शून्य त्रुटी (वेगळ्या चाचणीत बदलतो)
  return env;
}
const wz = (env, extra) => env.get(Object.assign({ action: 'getYearEndWizard', token: sup(env) }, extra || {}));
const step = (env, n, status, reason, year) => env.get({ action: 'setYearEndStep', token: sup(env), year: year === undefined ? '2026-27' : year, step: n, status, reason: reason || '' });
const st = (r, n) => r.steps[n - 1].status;

T.wizard_initial_state_and_order = () => {
  const e = yeEnv(); const r = wz(e); ok(r.status === 'ok', 'ok'); eq([r.closing, r.newYear, r.current, r.lastClass], ['2026-27', '2027-28', '2026-27', '10th'], 'वर्षे/शेवटचा वर्ग');
  eq(st(r, 1), 'auto', 'पायरी १: ताजा Backup = आपोआप'); eq(st(r, 2), 'auto', 'पायरी २: शून्य त्रुटी = आपोआप'); eq(st(r, 3), 'active', 'पायरी ३ सक्रिय (थकबाकी आढावा — नोंद आवश्यक)');
  eq(r.steps.slice(3).map(x => x.status), ['locked', 'locked', 'locked', 'locked', 'locked', 'locked'], 'पुढच्या पायऱ्या लॉक'); eq(r.active, 3, 'active = 3');
  const bad = step(e, 5, 'skipped', 'काहीतरी कारण'); ok(bad.status === 'error' && /पायरी 3/.test(bad.message), 'क्रम मोडता येत नाही: ' + bad.message);
  ok(step(e, 3, 'skipped', 'abc').status === 'error', 'वगळण्यास कारण ≥ ५ अक्षरे');
  const d3 = step(e, 3, 'done'); ok(d3.status === 'ok' && st(d3, 3) === 'done' && d3.steps[2].by === 'boss', 'पायरी ३ आढावा घेतला'); eq(d3.active, 4, 'पुढची सक्रिय');
  ok(step(e, 3, 'done').status === 'error', 'आधीच पूर्ण पायरी पुन्हा नाही');
};
T.wizard_backup_and_dq_rules = () => {
  const e = yeEnv(); e.props['LAST_BACKUP'] = JSON.stringify({ ok: true, time: new Date(Date.now() - 80 * 3600000).toISOString() });
  let r = wz(e); eq(st(r, 1), 'active', 'जुनी Backup (८० तास) = पायरी १ बाकी'); ok(step(e, 1, 'done').status === 'error', 'ताजी Backup नसताना done नाही');
  const s1 = step(e, 1, 'skipped', 'Backup हाताने घेतली'); ok(s1.status === 'ok' && st(s1, 1) === 'skipped', 'कारणासह वगळले'); eq(s1.steps[0].reason, 'Backup हाताने घेतली', 'कारण नोंद');
  const e2 = yeEnv(); e2.run('dqCompute_ = function(r, t) { return { byStudent: { r1: [1], r2: [2] } }; };'); r = wz(e2); eq(st(r, 2), 'active', 'डेटा-त्रुटी असताना पायरी २ बाकी'); eq(r.steps[1].info.issues, 2, 'त्रुटी-संख्या दिसते');
  ok(step(e2, 2, 'done').status === 'error', 'त्रुटी असताना done नाही — स्वीकारायचे तर कारण'); const a = step(e2, 2, 'skipped', '2 जुन्या नोंदी स्वीकारल्या'); ok(a.status === 'ok' && st(a, 2) === 'skipped', 'त्रुटी "स्वीकारले + कारण"');
};
T.wizard_lc_pending_rule = () => {
  const e = yeEnv(); step(e, 3, 'done'); let r = wz(e); eq(st(r, 4), 'active', 'पायरी ४ सक्रिय'); eq(r.steps[3].info.lcPending, 2, '10th मध्ये ३ पैकी २ ची LC बाकी (R5, R6)');
  ok(step(e, 4, 'done').status === 'error', 'LC बाकी असताना done नाही'); const s = step(e, 4, 'skipped', 'LC पुढच्या आठवड्यात देणार'); ok(s.status === 'ok' && st(s, 4) === 'skipped', 'कारणासह स्वीकारले'); eq(s.active, 5, 'पुढे: पायरी ५');
};
T.wizard_snapshot_logaudit_and_isolation = () => {
  const e = yeEnv(); const before = e.snapshot(['Log_2026_11', 'Log_' + e.run('Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy_MM")')]);
  wz(e); const snapProp = JSON.parse(e.props['YEAREND_2026-27']); eq(snapProp.snap.counts, { '5th': 1, '9th': 2, '10th': 3 }, 'बढतीपूर्वीची वर्गनिहाय संख्या जतन');
  step(e, 3, 'done'); const logSheets = e.ss.sheets.filter(s => /^Log_/.test(s.name)); ok(logSheets.length >= 1 && JSON.stringify(logSheets.map(s => s.rows)).indexOf('yearEndStep') >= 0, 'Activity Log मध्ये yearEndStep नोंद');
  const after = e.snapshot(logSheets.map(s => s.name)); eq(after, before, 'Students/Users/LC/PromotionLog इ. कोणतीही Sheet बदलली नाही (फक्त Log)');
  const w = bookWrites(e); wz(e); wz(e); eq(bookWrites(e), w, 'getYearEndWizard = Sheet लेखन शून्य');
  eq(Object.keys(e.props).filter(k => /^YEAREND_/.test(k)), ['YEAREND_2026-27'], 'फक्त YEAREND_<वर्ष> property');
};
T.wizard_full_flow_detection = () => {
  const e = yeEnv(); const stS = e.sheet('Students'); step(e, 3, 'done'); step(e, 4, 'skipped', 'LC नंतर देणार');
  let r = wz(e); eq(r.active, 5, 'पायरी ५ सक्रिय'); ok(step(e, 5, 'done').status === 'error', 'Pass-out न झाल्यास done नाही (आढळत नाही)');
  // विद्यमान archivePassedOut चे परिणाम अनुकरण: 10th च्या ओळी Students मधून काढा, PassedOut_2026-27 मध्ये ठेवा
  const po = e.seed('PassedOut_2026-27', e.ctx.MAIN_HEADERS.concat(['A', 'B', 'C']), []); [4, 5, 6].forEach(i => po.appendRow(stS.rows[i].slice()));
  stS.rows = stS.rows.filter((row, i) => i === 0 || row[6] !== '10th'); e.run('invalidateRosterCache_()'); r = wz(e); eq(st(r, 5), 'auto', 'पायरी ५ आपोआप आढळली (10th रिकामी)'); eq(r.active, 6, 'पुढे बढती');
  const d5 = step(e, 5, 'done'); ok(d5.status === 'ok' && st(d5, 5) === 'done', 'आढळल्यावर done नोंदता येते'); eq(d5.steps[4].by, 'boss', 'कोणी');
  // बढती: PromotionLog + वर्ग बदल
  const pl = e.sheet('PromotionLog'); const mkLog = (reg, o, n, un) => pl.appendRow(['ts', 'P1', reg, 'x', o, 'A', '2026-27', n, '2027-28', 'boss', un || '']);
  mkLog('R2', '9th', '10th', 'Y'); r = wz(e); eq(st(r, 6), 'active', 'सर्व Undo झालेली बढती = बाकी'); pl.rows = [pl.rows[0]];
  mkLog('R2', '9th', '10th'); mkLog('R3', '9th', '10th'); mkLog('R1', '5th', '6th');
  stS.rows.forEach(row => { if (row[2] === 'R2' || row[2] === 'R3') { row[6] = '10th'; row[1] = '2027-28'; } if (row[2] === 'R1') { row[6] = '6th'; row[1] = '2027-28'; } });
  e.run('invalidateRosterCache_()'); r = wz(e); eq(st(r, 6), 'auto', 'बढती आपोआप आढळली'); eq(r.steps[5].info.promoted, 3, '३ विद्यार्थी'); ok(step(e, 6, 'done').status === 'ok', 'पायरी ६ नोंद');
  // नवे वर्ष
  r = wz(e); eq(st(r, 7), 'active', 'पायरी ७ सक्रिय'); ok(step(e, 7, 'done').status === 'error', 'CURRENT_AC_YEAR बदलले नसताना done नाही');
  e.props['CURRENT_AC_YEAR'] = '2027-28'; r = wz(e); eq(r.closing, '2026-27', 'नवे वर्ष सुरू झाल्यावरही wizard मागील वर्षासाठी चालू (पायरी ९ बाकी)'); eq(st(r, 7), 'auto', 'पायरी ७ आढळली');
  const d7 = step(e, 7, 'done'); ok(d7.status === 'ok' && d7.steps[6].time, 'पायरी ७ वेळेसह नोंद (पायरी ८ च्या २-३ दिवस इशाऱ्यासाठी)');
  r = wz(e); eq(r.active, 8, 'पायरी ८ सक्रिय'); eq(r.steps[7].info.startedAt, d7.steps[6].time, 'नवे वर्ष सुरू केल्याची वेळ');
  e.props['ARCHIVED_YEARS'] = '2026-27'; r = wz(e); eq(st(r, 8), 'auto', 'जुने वर्ष Archive आढळले'); step(e, 8, 'done');
  // अंतिम पडताळणी: ५th→6th(१), 9th→10th(२), PassedOut(३)
  r = wz(e); eq(r.active, 9, 'पायरी ९ सक्रिय'); const v = r.verify; ok(v && v.curOk && v.allOk, 'पडताळणी सर्व ठीक: ' + JSON.stringify(v));
  eq(v.classes.find(c => c.cls === '10th').expected, 2, '10th अपेक्षित = मागील 9th (२)'); eq(v.classes.find(c => c.cls === '10th').actual, 2, '10th प्रत्यक्ष २'); eq(v.classes.find(c => c.cls === '6th').expected, 1, '6th अपेक्षित = मागील 5th (१)'); eq(v.classes.find(c => c.cls === '5th').expected, null, '5th: नवे प्रवेश — तुलना नाही');
  eq([v.passedOut.expected, v.passedOut.actual], [3, 3], 'Pass-out अपेक्षित/प्रत्यक्ष'); eq(v.notNewYear, 0, 'सर्व विद्यार्थी 2027-28');
  const d9 = step(e, 9, 'done'); ok(d9.status === 'ok' && d9.complete && d9.active === 0, 'wizard पूर्ण'); eq(d9.steps.every(x => ['done', 'auto', 'skipped'].includes(x.status)), true, 'सर्व पायऱ्या पूर्ण');
  const r2 = wz(e); eq(r2.closing, '2027-28', 'पूर्ण झाल्यानंतर नवा चक्र चालू वर्षासाठी (2027-28)');
};
T.wizard_step5_auto_after_promotion_even_if_10th_refilled = () => {
  const e = yeEnv(); const stS = e.sheet('Students');
  stS.rows = stS.rows.filter((row, i) => i === 0 || row[6] !== '10th'); stS.rows.forEach(row => { if (row[2] === 'R2' || row[2] === 'R3') { row[6] = '10th'; row[1] = '2027-28'; } });   // 10th रिकामी होऊन बढतीने पुन्हा भरली
  e.sheet('PromotionLog').appendRow(['ts', 'P1', 'R2', 'x', '9th', 'A', '2026-27', '10th', '2027-28', 'boss', '']); e.run('invalidateRosterCache_()');
  const r = wz(e); eq(st(r, 5), 'auto', 'बढती झाली असल्यास पायरी ५ (Pass-out) पूर्ण मानली जाते — 10th पुन्हा भरली तरी'); eq(st(r, 6), 'auto', 'पायरी ६ आढळली');
};
T.wizard_verify_mismatch_blocks_done = () => {
  const e = yeEnv(); const stS = e.sheet('Students'); wz(e);   // snapshot
  e.props['CURRENT_AC_YEAR'] = '2027-28'; stS.rows.forEach(row => { if (row[2] === 'R2') { row[6] = '10th'; row[1] = '2027-28'; } }); e.run('invalidateRosterCache_()');   // फक्त एक बढती; 10th अजून आहे
  const pl = e.sheet('PromotionLog'); pl.appendRow(['ts', 'P1', 'R2', 'x', '9th', 'A', '2026-27', '10th', '2027-28', 'boss', '']);
  ['3', '4', '5', '6', '7', '8'].forEach(n => { e.props['YEAREND_2026-27'] = e.props['YEAREND_2026-27']; });
  let rec = JSON.parse(e.props['YEAREND_2026-27']); [3, 4, 5, 6, 7, 8].forEach(n => { rec.steps[n] = { s: 'skipped', by: 'boss', t: new Date().toISOString(), r: 'चाचणीसाठी वगळले' }; }); e.props['YEAREND_2026-27'] = JSON.stringify(rec);
  const r = wz(e); eq(r.active, 9, 'पायरी ९ सक्रिय'); ok(r.verify && r.verify.allOk === false, 'पडताळणी जुळत नाही (allOk=false)'); ok(r.verify.classes.some(c => !c.ok), 'काही वर्गांची संख्या जुळत नाही'); ok(r.verify.notNewYear > 0, 'काही विद्यार्थ्यांचे वर्ष जुने');
  const bad = step(e, 9, 'done'); ok(bad.status === 'error', 'जुळत नसताना done नाही — ' + bad.message); const ok9 = step(e, 9, 'skipped', 'फरक तपासून स्वीकारला'); ok(ok9.status === 'ok' && ok9.complete, 'कारणासह स्वीकारता येते');
};
T.wizard_roles_and_validation = () => {
  const e = yeEnv(); ['master', 'deo', 'cert', 'teacher', 'peon'].forEach(role => { const tk = e.token('z', role, role === 'teacher' ? { c: '8th|A' } : {}); ok(e.get({ action: 'getYearEndWizard', token: tk }).status === 'error', role + ': getYearEndWizard नाकारला'); ok(e.get({ action: 'setYearEndStep', token: tk, year: '2026-27', step: 3, status: 'done' }).status === 'error', role + ': setYearEndStep नाकारला'); });
  ok(e.get({ action: 'getYearEndWizard' }).status === 'error', 'Token नसताना नाकारला');
  ['x', '2026', ''].forEach(y => ok(step(e, 3, 'done', '', y).status === 'error', 'अवैध वर्ष ' + y)); [0, 10, 'a'].forEach(n => ok(step(e, n, 'done').status === 'error', 'अवैध पायरी ' + n)); ok(step(e, 3, 'maybe').status === 'error', 'अवैध स्थिती');
};
T.wizard_property_size = () => {
  const e = yeEnv(); const rec = { steps: {} }; for (let i = 1; i <= 9; i++) rec.steps[i] = { s: 'skipped', by: 'boss', t: new Date().toISOString(), r: 'क'.repeat(150) }; rec.snap = { t: 'x', counts: { '5th': 1 }, total: 1 };
  e.run('yeSave_("2026-27", ' + JSON.stringify(rec) + ')'); ok(e.props['YEAREND_2026-27'].length < 9000, 'Property < ९ KB (' + e.props['YEAREND_2026-27'].length + ')');
};
T.wizard_existing_actions_unchanged = () => {
  if (!GS_BASE) { console.log('  (वगळले: SGS_GS_BASE दिलेला नाही)'); return; }
  const a = makeEnv(), b = makeEnv(); a.load(GS_BASE); b.load(GS);
  ['doStartNewYear', 'doPromoteStudents', 'doUndoPromotion', 'doArchivePassedOut', 'doRestorePassedOut', 'doArchiveYear', 'doRunBackup', 'backupStaleErr_', 'computeCarryDues_', 'doGetYearEndCheck', 'doGetYearStatus', 'dqCompute_', 'doGetArchiveHealth', 'nextAcYear_', 'currentAcYear_'].forEach(fn => eq(b.run(fn + '.toString()'), a.run(fn + '.toString()'), fn + ' चा कोड V19.62 सारखाच (diff शून्य)'));
};

// ---------------- ब-१० ----------------
const D = '2026-11-02';   // सोमवार
function dsEnv() {
  const env = makeEnv(); env.load(GS); const c = env.ctx; env.run('todayStr = function() { return "2026-11-03"; };');
  env.seed('Students', c.MAIN_HEADERS, [stu('R1', '8th', 'A'), stu('R2', '8th', 'A'), stu('R3', '8th', 'A'), stu('R4', '8th', 'B'), stu('R5', '8th', 'B'), stu('R6', '9th', 'A'), stu('R7', '9th', 'A'), stu('R8', '7th', 'A')]);
  env.seed('Users', c.USER_HEADERS, [['boss', 'x', 'super', 'B', '', ''], ['tA', 'x', 'teacher', 'सौ. अ', '8th|A', '9811111111'], ['tB', 'x', 'teacher', 'श्री. ब', '8th|B', '9822222222'], ['tC', 'x', 'teacher', 'सौ. क', '9th|A', '']]);
  env.seed('AttMarks_2026_11', c.ATT_MARKS_HEADERS, [[D, '8th', 'A', 3, 1, 2, 'tA', '2/11/2026, 9:10:00 am'], [D, '9th', 'A', 2, 0, 2, 'tC', '2/11/2026, 9:20:00 am'], ['2026-11-01', '8th', 'A', 3, 3, 0, 'tA', 'x']]);
  env.seed('EL_2026-11-02', c.EL_HEADERS, [
    ['ts', D, '8th', 'A', 'R1', 'S1', 'a', 'permitted', 3, 'tA', 'teacher', 'gone', 'peon', 'x'], ['ts', D, '8th', 'A', 'R2', 'S2', 'b', 'permitted', 4, 'tA', 'teacher', 'gone', 'peon', 'x'],
    ['ts', D, '9th', 'A', 'R6', 'S6', 'c', 'permitted', 5, 'tC', 'teacher', 'notgone', 'peon', 'x'], ['ts', D, '9th', 'A', 'R7', 'S7', 'd', 'permitted', 5, 'tC', 'teacher', '', '', ''],
    ['ts', D, '8th', 'B', 'R4', 'S4', 'e', 'unpermitted', 2, 'tB', 'teacher', '', '', ''], ['ts', '2026-11-03', '8th', 'A', 'R1', 'S1', 'a', 'permitted', 3, 'tA', 'teacher', 'gone', '', '']]);
  const fe = env.seed('Fees_2026-27', c.FEES_HEADERS, []);
  [['1/11/2026, 10:00:00 am', 'R1', '8th', 500], ['2/11/2026, 9:00:00 am', 'R2', '8th', 300], ['2/11/2026, 11:00:00 am', 'R1', '8th', 800], ['2/11/2026, 12:00:00 pm', 'R6', '9th', 200], ['12/11/2026, 10:00:00 am', 'R3', '8th', 999]].forEach(x => fe.appendRow([x[0], x[1], 'S' + x[1], 'n', x[2], 'A', '2026-27', 5000, x[3], 0, 'u']));
  fe.appendRow(['2/11/2026, 1:00:00 pm', 'R7', 'SR7', 'n', '9th', 'A', '2026-27', 5000, 0, 400, 'OPENING']);
  fe.appendRow(['2/11/2026, 2:00:00 pm', 'R9', 'SR9', 'n', '9th', 'A', '2026-27', 5000, 150, 0, 'OPENING']);   // OPENING (मागील वर्षाची बाकी) — जमा म्हणून मोजू नये
  fe.appendRow([new Date(Date.UTC(2026, 10, 2, 5, 0)), 'R7', 'SR7', 'n', '9th', 'A', '2026-27', 5000, 100, 0, 'u']);   // Sheets ने Date म्हणून साठवलेली (IST 10:30)
  env.seed('Notices', c.NOTICE_HEADERS, [['ts', D, 'सुट्टी सूचना', 'm', 'boss', ''], ['ts', D, 'पालक सभा', 'm', 'mas', '8th'], ['ts', '2026-11-05', 'पुढची', 'm', 'boss', '']]);
  env.seed('Student_Audit', c.AUDIT_HEADERS, [['2/11/2026, 10:00:00 am', 'u', 'r', 'R1', 'saveStudentContact', 'f', 'o', 'n'], ['2/11/2026, 10:05:00 am', 'u', 'r', 'R1', 'updateClassDivision', 'f', 'o', 'n'], ['2/11/2026, 11:00:00 am', 'u', 'r', 'R2', 'saveStudentContact', 'f', 'o', 'n'], ['2/11/2026, 12:00:00 pm', 'u', 'r', '—', 'renumberRolls', '—', '', 'एकूण 700'], ['3/11/2026, 9:00:00 am', 'u', 'r', 'R3', 'saveStudentContact', 'f', 'o', 'n'], [new Date(Date.UTC(2026, 10, 2, 8, 0)), 'u', 'r', 'R5', 'updateClassDivision', 'f', 'o', 'n']]);
  return env;
}
const ds = (env, date, extra) => env.get(Object.assign({ action: 'getDailySummary', token: sup(env), date }, extra || {}));
T.summary_attendance_numbers = () => {
  const e = dsEnv(); const r = ds(e, D); ok(r.status === 'ok' && r.dayType === 'normal', 'ok/normal — ' + JSON.stringify(r).slice(0, 80));
  eq(r.attendance.classes.map(c => [c.iyatta, c.tukdi, c.total, c.absent, c.present]), [['8th', 'A', 3, 1, 2], ['9th', 'A', 2, 0, 2]], 'वर्गनिहाय हजर/गैरहजर'); eq(r.attendance.totals, { total: 5, absent: 1, present: 4, pct: 80 }, 'एकूण'); eq(r.attendance.classes[0].pct, 66.7, '8th A %');
  eq(r.attendance.pending, [{ iyatta: '8th', tukdi: 'B', teacher: 'श्री. ब', mobile: '9822222222' }], 'हजेरी भरली नसलेला वर्ग + वर्ग-शिक्षक + मोबाईल'); eq(r.attendance.noTeacher, [{ iyatta: '7th', tukdi: 'A' }], 'वर्ग-शिक्षक नसलेला वर्ग'); eq(r.attendance.filled, 2, 'भरलेले वर्ग २');
};
T.summary_early_leave = () => { const e = dsEnv(); const r = ds(e, D); eq(r.earlyLeave, { permitted: 4, gone: 2, notGone: 1, other: 1, total: 5 }, 'घरी जाणारे: ४ परवानगी, २ गेले, १ गेला नाही, १ इतर; दुसऱ्या तारखेची ओळ वगळली'); };
T.summary_fees_cumulative_delta = () => {
  const e = dsEnv(); const r = ds(e, D); eq(r.fees.total, 900, 'जमा = R2 ३००(पहिली) + R1 ३००(८००−५००) + R6 २०० + R7 १००(Date मूल्य) = ९००'); eq(r.fees.count, 4, '४ नोंदी');
  eq(r.fees.byClass, [{ cls: '8th', amount: 600 }, { cls: '9th', amount: 300 }], 'वर्गनिहाय'); eq(r.fees.year, '2026-27', 'शैक्षणिक वर्ष');
  eq(ds(e, '2026-11-01').fees.total, 500, 'दुसऱ्या दिवशी स्वतंत्र (500)'); e.run('todayStr = function() { return "2026-11-30"; };'); eq(ds(e, '2026-11-12').fees.total, 999, '१२/११ ची नोंद "2/11" शी गल्लत होत नाही');
};
T.summary_notices_and_changes = () => {
  const e = dsEnv(); const r = ds(e, D); eq(r.notices.map(n => n.title), ['सुट्टी सूचना', 'पालक सभा'], 'आजच्या २ नोटीस'); eq(r.notices[1].targetClass, '8th', 'लक्ष्य वर्ग');
  eq([r.changes.total, r.changes.students], [5, 3], 'बदल: ५ ओळी (Date मूल्यासह; ३/११ वगळून), ३ विद्यार्थी (R1, R2, R5; "—" वगळला)'); eq(r.changes.byAction[0], { action: 'saveStudentContact', count: 2 }, 'सर्वाधिक action');
  ok(!/9822/.test(JSON.stringify(r.changes)), 'सारांशात फोन/माहिती नाही');
};
T.summary_day_types = () => {
  const e = dsEnv(); let r = ds(e, '2026-11-01'); eq([r.dayType, r.message], ['sunday', 'रविवार — सुट्टी'], 'रविवार');
  r = ds(e, '2026-10-20'); eq(r.dayType, 'nodata_past', 'हजेरी नसलेला जुना दिवस'); ok(/सुट्टी असावी/.test(r.message), 'सुट्टी संदेश');
  r = ds(e, '2026-11-03'); eq(r.dayType, 'nodata_today', 'आज अजून हजेरी नाही'); ok(/अजून/.test(r.message), 'आजचा संदेश'); eq(r.attendance.pending.length, 3, 'आज ३ वर्ग-शिक्षकांची हजेरी बाकी');
  r = ds(e, '2026-11-10'); eq(r.dayType, 'future', 'भविष्यातील तारीख'); ok(!r.attendance && /भविष्यातील/.test(r.message), 'कोणताही डेटा नाही');
  r = e.get({ action: 'getDailySummary', token: sup(e), date: 'junk' }); eq(r.date, '2026-11-03', 'अवैध तारीख = आज'); r = e.get({ action: 'getDailySummary', token: sup(e) }); eq(r.date, '2026-11-03', 'तारीख नसल्यास आज');
};
T.summary_roles = () => {
  const e = dsEnv(); ['deo', 'cert', 'teacher', 'peon'].forEach(role => ok(ds(e, D, { token: e.token('z', role, role === 'teacher' ? { c: '8th|A' } : {}) }).status === 'error', role + ': नाकारला'));
  ok(e.get({ action: 'getDailySummary', date: D }).status === 'error', 'Token नसताना नाकारला'); ok(ds(e, D, { token: e.token('mas', 'master') }).status === 'ok', 'Master ला मिळतो'); ok(ds(e, D).status === 'ok', 'Super ला मिळतो');
};
T.summary_readonly_cache_and_read_budget = () => {
  const e = dsEnv(); const w0 = bookWrites(e), r0 = bookReads(e); const a = ds(e, D); const reads = bookReads(e) - r0; eq(bookWrites(e), w0, 'फक्त वाचन — Sheet लेखन शून्य');
  ok(reads <= 14, 'वाचन-संख्या मर्यादित: ' + reads + ' (≤ १४) — पूर्ण Sheet-वाचन नाही, Sheet प्रत्येकी एकदाच'); ok(!a.cached, 'पहिल्यांदा cached नाही');
  const r1 = bookReads(e); const b = ds(e, D); ok(b.cached === true && bookReads(e) === r1, 'दुसऱ्यांदा Cache — शून्य Sheet-वाचन'); eq(b.fees.total, a.fees.total, 'Cache चे उत्तर तेच');
  e.run('attBump_()'); const c = ds(e, D); ok(!c.cached, 'हजेरी Save नंतर (attBump_) Cache रद्द');
};
T.summary_missing_sheets_no_crash = () => {
  const e = makeEnv(); e.load(GS); e.run('todayStr = function() { return "2026-11-03"; };'); e.seed('Students', e.ctx.MAIN_HEADERS, [stu('R1', '8th', 'A')]); e.seed('Users', e.ctx.USER_HEADERS, [['boss', 'x', 'super', 'B', '', '']]);
  const r = ds(e, D); ok(r.status === 'ok', 'सर्व सहाय्यक Sheets नसतानाही ok — ' + JSON.stringify(r).slice(0, 100)); eq([r.fees.total, r.earlyLeave.total, r.notices.length, r.changes.total], [0, 0, 0, 0], 'सर्व शून्य');
};
T.version = () => { const e = makeEnv(); e.load(GS); eq(e.run('SERVER_VERSION'), 'V19.63', 'SERVER_VERSION'); };

for (const k of Object.keys(T)) { const before = fails; try { T[k](); } catch (err) { fails++; console.log('  EXCEPTION in ' + k + ': ' + (err && err.stack ? err.stack.split('\n').slice(0, 3).join(' | ') : err)); } console.log((fails === before ? 'ok   ' : 'FAIL ') + k); }
console.log(`\n${fails ? 'FAIL' : 'PASS'}: test_server_v1963 (${passes} तपासण्या पास, ${fails} अपयशी)`);
process.exit(fails ? 1 : 0);
