// V19.65 सर्व्हर चाचण्या — चालू शैक्षणिक वर्ष (ping), Students column B = चालू वर्ष (जुळवणी + Save), बढती: वर्ष-पडताळणी व Backup-पडताळणी.
// बनावट Apps Script; खऱ्या Sheet/Drive वर नाही. वापर: SGS_GS=<Code.gs> [SGS_GS_BASE=<V19.64 Code.gs>] node tests/test_server_v1965.js
const path = require('path');
const { makeEnv } = require('./lib/gas_mock');
const GS = process.env.SGS_GS || path.join(__dirname, '..', 'apps_script', 'Code.gs'), GS_BASE = process.env.SGS_GS_BASE || '';
let fails = 0, passes = 0;
function ok(c, m) { if (c) passes++; else { fails++; console.log('  FAIL: ' + m); } }
function eq(a, b, m) { ok(JSON.stringify(a) === JSON.stringify(b), m + ' → मिळाले ' + JSON.stringify(a) + ' अपेक्षित ' + JSON.stringify(b)); }
const T = {};
let HDR1 = ''; { const e0 = makeEnv(); e0.load(GS); HDR1 = e0.ctx.MAIN_HEADERS[1]; }   // column B चे शीर्षक (बदलांच्या इतिहासातील Field नाव)
const sup = e => e.token('boss', 'super'), bookWrites = e => Object.values(e.books).reduce((a, b) => a + b.writes, 0);
const stu = (reg, cls, yr, name) => ['t', yr, reg, 'S' + reg, '', '', cls, 'A', '', '1', name || ('नाव ' + reg), 'मातोश्री', 'Male', '', '', '', '2012-05-04', '', '', '', '', '', '', '', '', '', '', '9822011111'];
function mk(extra) {
  const env = makeEnv(); env.load(GS); const c = env.ctx; env.run('todayStr = function() { return "2026-11-03"; };');
  env.seed('Students', c.MAIN_HEADERS, [stu('R1', '8th', '2025-26'), stu('R2', '8th', '2026-27'), stu('R3', '9th', '2024-25'), stu('R4', '9th', ''), stu('R5', '9th', '2026-27'), stu('R6', '5th', '2025-26')]);
  env.seed('Users', c.USER_HEADERS, [['boss', 'x', 'super', 'B', '', ''], ['mas', 'x', 'master', 'M', '', '']]); env.seed('PromotionLog', c.PROMO_LOG_HEADERS, []);
  env.props['CURRENT_AC_YEAR'] = '2026-27'; return env;
}
const info = (e, role) => e.get({ action: 'getAcYearInfo', token: e.token('u', role || 'super') });
const yearsCol = e => e.sheet('Students').rows.slice(1).map(r => r[1]);
const otherCols = e => JSON.stringify(e.sheet('Students').rows.map(r => r.filter((v, i) => i !== 1)));

T.ping_has_current_year = () => {
  const e = mk(); eq(e.get({ action: 'ping' }).acYear, '2026-27', 'ping: चालू शैक्षणिक वर्ष (Property)'); delete e.props['CURRENT_AC_YEAR'];
  eq(e.get({ action: 'ping' }).acYear, '2026-27', 'Property नसल्यास तारखेवरून (Nov 2026 → 2026-27)'); e.run('todayStr = function() { return "2026-04-10"; };'); eq(e.get({ action: 'ping' }).acYear, '2025-26', 'एप्रिल 2026 → 2025-26 (वर्ष जून–मे)');
};
T.info_roles_counts_flags = () => {
  const e = mk(); ['deo', 'cert', 'teacher', 'peon'].forEach(r => ok(info(e, r).status === 'error', r + ': नाकारला')); ok(e.get({ action: 'getAcYearInfo' }).status === 'error', 'Token नसताना नाकारला');
  let r = info(e, 'master'); ok(r.status === 'ok', 'Master ला मिळते'); eq(r.canSync, false, 'Master ला sync नाही'); r = info(e);
  eq([r.current, r.dateYear, r.total, r.mismatch, r.future, r.promoPending], ['2026-27', '2026-27', 6, 4, 0, 0], 'चालू/तारीख-वर्ष/एकूण/जुळत नाही 4/पुढचे 0/बढती-बाकी 0'); eq(r.counts, { '2025-26': 2, '2026-27': 2, '2024-25': 1, '—': 1 }, 'वर्षनिहाय संख्या (रिकामे "—")'); eq(r.canSync, true, 'Super ला sync शक्य'); eq(r.aheadOfDate, false, 'तारखेपुढे नाही'); ok(r.sample.length <= 8 && r.sample[0].regNo, 'नमुना ओळी');
  e.props['CURRENT_AC_YEAR'] = '2027-28'; r = info(e); eq([r.aheadOfDate, r.future, r.canSync], [false || true, 0, true], 'Property 2027-28 > तारीख-वर्ष: aheadOfDate = true (इशारा)'.replace('[false || true', '[true'));
  const e2 = mk(); e2.sheet('Students').rows[2][1] = '2027-28'; e2.run('invalidateRosterCache_()'); r = info(e2); eq([r.future, r.canSync], [1, false], 'पुढच्या वर्षाचा विद्यार्थी असल्यास sync निषिद्ध');
  const e3 = mk(); e3.sheet('PromotionLog').appendRow(['ts', 'P1', 'R1', 'x', '8th', 'A', '2026-27', '9th', '2027-28', 'boss', '']); r = info(e3); eq([r.promoPending, r.canSync], [1, false], 'बढती झाली, नवे वर्ष नाही → sync निषिद्ध'); e3.sheet('PromotionLog').rows[1][10] = 'Y'; eq(info(e3).promoPending, 0, 'Undo झालेली बढती मोजत नाही');
};
T.sync_success_backup_audit = () => {
  const e = mk(); const beforeOther = otherCols(e); const w0 = Object.keys(e.snapshot()).length;
  const r = e.get({ action: 'syncStudentsAcYear', token: sup(e), confirm: 'SYNC' }); ok(r.status === 'ok' && r.count === 4, 'sync ok, 4 बदल — ' + JSON.stringify(r)); eq(yearsCol(e), ['2026-27', '2026-27', '2026-27', '2026-27', '2026-27', '2026-27'], 'column B सर्वत्र चालू वर्ष');
  eq(otherCols(e), beforeOther, 'column B शिवाय कोणताही स्तंभ बदलला नाही');
  ok(/^Students_BeforeAcYearSync_\d{8}_\d{4}$/.test(r.backupSheet) && r.backupRows === 6, 'Backup नाव/ओळी: ' + r.backupSheet); const bk = e.ss.sheets.find(s => s.name === r.backupSheet); ok(!!bk && bk.getLastRow() === 7 && bk.rows[1][1] === '2025-26' && bk.rows[3][1] === '2024-25', 'Backup मध्ये जुने वर्ष (बदलापूर्वीची स्थिती)'); ok(bk.hidden === true && r.backupExternal === false, 'अंतर्गत (लपवलेली) प्रत — DELETED_BACKUP_ID नसताना');
  const a = e.sheet('Student_Audit'); const rows = a.rows.slice(1).filter(x => x[4] === 'syncAcYear'); eq(rows.length, 4, '४ Audit ओळी'); ok(rows.every(x => x[5] === HDR1 && x[2] === 'super' && x[1] === 'boss'), 'Field = column B शीर्षक (' + HDR1 + '); कोणी'); ok(rows.some(x => x[3] === 'R3' && x[6] === '2024-25' && x[7] === '2026-27') && rows.some(x => x[3] === 'R4' && x[6] === '' && x[7] === '2026-27'), 'जुने→नवे (रिकामे सहित)');
  ok(JSON.stringify(e.ss.sheets.filter(s => /^Log_/.test(s.name)).map(s => s.rows)).includes('syncAcYear'), 'Activity Log मध्ये नोंद'); eq(info(e).mismatch, 0, 'roster Cache रद्द — आता जुळत नाही = 0');
  const sheets = e.ss.sheets.length; const r2 = e.get({ action: 'syncStudentsAcYear', token: sup(e), confirm: 'SYNC' }); ok(r2.status === 'ok' && r2.count === 0, 'पुन्हा चालवल्यास 0'); eq(e.ss.sheets.length, sheets, 'नवीन Backup Sheet नाही (बदल नसल्यास)');
};
T.sync_external_backup_when_id_present = () => {
  const e = mk(); const b = e.addSource('delbk', 'SGS_Deleted_Backup'); e.props['DELETED_BACKUP_ID'] = 'delbk';
  const r = e.get({ action: 'syncStudentsAcYear', token: sup(e), confirm: 'SYNC' }); ok(r.status === 'ok' && r.backupExternal === true, 'बाह्य (SGS_Deleted_Backup) Backup'); ok(b.sheets.some(s => s.name === r.backupSheet), 'Backup त्या फाइलमध्ये'); ok(!e.ss.sheets.some(s => s.name === r.backupSheet), 'मुख्य फाइल हलकी (प्रत तिथे नाही)');
};
T.sync_guards_and_roles = () => {
  const e = mk(); ['master', 'deo', 'cert', 'teacher', 'peon'].forEach(r => ok(e.get({ action: 'syncStudentsAcYear', token: e.token('u', r), confirm: 'SYNC' }).status === 'error', r + ': नाकारला')); ok(e.get({ action: 'syncStudentsAcYear', confirm: 'SYNC' }).status === 'error', 'Token नसताना नाकारला');
  const before = JSON.stringify(e.snapshot()); ok(e.get({ action: 'syncStudentsAcYear', token: sup(e) }).status === 'error', 'confirm नसताना नाही'); ok(e.get({ action: 'syncStudentsAcYear', token: sup(e), confirm: 'sync' }).status === 'error', 'confirm केस-संवेदनशील (sync ≠ SYNC)'); eq(JSON.stringify(e.snapshot()), before, 'नकारलेल्या विनंत्यांनी काहीही बदलले नाही');
  const f = mk(); f.sheet('Students').rows[2][1] = '2027-28'; f.run('invalidateRosterCache_()'); const fb = JSON.stringify(f.snapshot()); const r = f.get({ action: 'syncStudentsAcYear', token: sup(f), confirm: 'SYNC' }); ok(r.status === 'error' && /पुढचे/.test(r.message), 'पुढचे वर्ष असल्यास नकार: ' + r.message); eq(JSON.stringify(f.snapshot()), fb, 'काहीही बदलले नाही / Backup Sheet नाही');
  const g = mk(); g.sheet('PromotionLog').appendRow(['ts', 'P1', 'R1', 'x', '8th', 'A', '2026-27', '9th', '2027-28', 'boss', '']); const gb = JSON.stringify(g.snapshot()); const r3 = g.get({ action: 'syncStudentsAcYear', token: sup(g), confirm: 'SYNC' }); ok(r3.status === 'error' && /नवे वर्ष सुरू/.test(r3.message), 'बढती-नंतर-नवे वर्ष-नाही: नकार'); eq(JSON.stringify(g.snapshot()), gb, 'काहीही बदलले नाही');
  const h = mk(); h.lockFail = true; const hb = JSON.stringify(h.snapshot()); const r4 = h.get({ action: 'syncStudentsAcYear', token: sup(h), confirm: 'SYNC' }); ok(r4.status === 'error', 'Lock न मिळाल्यास त्रुटी (अपवाद नाही)'); eq(JSON.stringify(h.snapshot()), hb, 'काहीही बदलले नाही');
};
T.sync_backup_verification_failure_aborts = () => {
  const e = mk(); e.run('makeStudentsBackup_ = function(sh, name, prefix) { var ss = SpreadsheetApp.getActiveSpreadsheet(); var cp = sh.copyTo(ss); cp.setName(name); cp.hideSheet(); cp.rows = cp.rows.slice(0, 3); return {ss: ss, sheet: cp, external: false}; };');   // अपूर्ण प्रत
  const before = otherCols(e), years = yearsCol(e).join(); const r = e.get({ action: 'syncStudentsAcYear', token: sup(e), confirm: 'SYNC' }); ok(r.status === 'error' && /Backup/.test(r.message), 'अपूर्ण Backup = जुळवणी नाही: ' + r.message); eq(yearsCol(e).join(), years, 'column B बदलला नाही'); eq(otherCols(e), before, 'इतर काही बदलले नाही'); ok(!e.ss.sheets.some(s => /BeforeAcYearSync/.test(s.name)), 'अपूर्ण Backup Sheet काढली');
};
T.save_forces_current_year = () => {
  const e = mk(); const up = (extra) => e.post(Object.assign({ action: 'upsert', token: e.token('boss', 'super'), regNo: 'R2', studentId: 'SR2', aadhar: '', iyatta: '8th', tukdi: 'A', firstName: 'नाव R2', whatsappMobile: '9822011111' }, extra));
  let r = up({ acYear: '2020-21' }); ok(r.status === 'ok', 'upsert ok'); eq(e.sheet('Students').rows[2][1], '2026-27', 'चुकीचे वर्ष (2020-21) पाठवले तरी column B = चालू वर्ष'); eq((e.sheet('Student_Audit') ? e.sheet('Student_Audit').rows.slice(1).filter(x => x[5] === HDR1).length : 0), 0, 'आधीच चालू वर्ष असल्यास AcYear चा खोटा बदल नोंद नाही');
  r = up({ regNo: 'R1', studentId: 'SR1', acYear: '2026-27', firstName: 'नाव R1' }); ok(r.status === 'ok', 'R1 upsert'); eq(e.sheet('Students').rows[1][1], '2026-27', 'R1 (2025-26) → चालू वर्ष'); const a = e.sheet('Student_Audit').rows.slice(1).filter(x => x[3] === 'R1' && x[5] === HDR1); ok(a.length === 1 && a[0][6] === '2025-26' && a[0][7] === '2026-27', 'AcYear बदल बदलांच्या इतिहासात: ' + JSON.stringify(a));
  r = up({ regNo: 'R99', studentId: 'SR99', acYear: '', firstName: 'नवा' }); ok(r.status === 'ok', 'नवी नोंद'); eq(e.sheet('Students').rows.slice(-1)[0][1], '2026-27', 'नवी नोंद (वर्ष रिकामे) → चालू वर्ष'); r = up({ regNo: 'R98', studentId: 'SR98', acYear: '1999-00', firstName: 'नवा२' }); eq(e.sheet('Students').rows.slice(-1)[0][1], '2026-27', 'नवी नोंद (1999-00) → चालू वर्ष');
  e.props['CURRENT_AC_YEAR'] = '2027-28'; up({ regNo: 'R5', studentId: 'SR5', acYear: '2026-27', firstName: 'नाव R5' }); eq(e.sheet('Students').rows[5][1], '2027-28', 'चालू वर्ष बदलल्यावर नवे चालू वर्ष वापरले');
};
const promo = (e, map, ny, extra) => e.post(Object.assign({ action: 'promoteStudents', token: sup(e), mapJson: JSON.stringify(map), newAcYear: ny, confirm: 'PROMOTE' }, extra || {}));
T.promotion_year_validation = () => {
  const e = mk(); const before = JSON.stringify(e.snapshot());
  ['2026-26', '2027-29', 'abcd', '27-28', '2029-30', '2024-25'].forEach(y => { const r = promo(e, { '8th': '9th' }, y); ok(r.status === 'error', 'अवैध/चुकीचे वर्ष "' + y + '" नाकारले: ' + (r.message || '').slice(0, 60)); }); eq(JSON.stringify(e.snapshot()), before, 'नकारलेल्या विनंत्यांनी काहीही बदलले नाही');
  const r = promo(e, { '8th': '9th' }, '2027-28'); ok(r.status === 'ok', '2027-28 (चालू+१) स्वीकारले — ' + JSON.stringify(r).slice(0, 100)); const e2 = mk(); ok(promo(e2, { '8th': '9th' }, '2026-27').status === 'ok', 'चालू वर्षच (नवे वर्ष आधी सुरू झाले असल्यास) स्वीकारले');
};
T.promotion_sets_column_b_and_verifies_backup = () => {
  const e = mk(); const r = promo(e, { '8th': '9th', '5th': '6th' }, '2027-28'); ok(r.status === 'ok' && r.count === 3, 'बढती ३ विद्यार्थी — ' + JSON.stringify(r).slice(0, 120));
  const S = e.sheet('Students').rows; eq([S[1][6], S[1][1]], ['9th', '2027-28'], 'R1: वर्ग 9th, column B = 2027-28'); eq([S[2][6], S[2][1]], ['9th', '2027-28'], 'R2'); eq([S[6][6], S[6][1]], ['6th', '2027-28'], 'R6: 5th→6th'); eq([S[3][6], S[3][1]], ['9th', '2024-25'], 'R3 (9th, नकाशात नाही): वर्ग/वर्ष बदलले नाही'); eq(S[5][1], '2026-27', 'R5 बदलले नाही');
  ok(/^Students_BeforePromotion_/.test(r.extra.backupSheet) && r.extra.backupRows === 6 && r.extra.backupExternal === false, 'Backup नाव/ओळी/ठिकाण अहवालात: ' + JSON.stringify([r.extra.backupSheet, r.extra.backupRows, r.extra.backupExternal])); const bk = e.ss.sheets.find(s => s.name === r.extra.backupSheet); ok(bk && bk.getLastRow() === 7 && bk.rows[1][6] === '8th' && bk.rows[1][1] === '2025-26', 'Backup = बदलापूर्वीची स्थिती (8th, 2025-26)');
  const pl = e.sheet('PromotionLog'); eq(pl.getLastRow() - 1, 3, 'PromotionLog ३ ओळी'); ok(pl.rows.slice(1).every(x => x[8] === '2027-28'), 'PromotionLog मध्ये नवे वर्ष');
  const e2 = mk(); const b = e2.addSource('delbk', 'SGS_Deleted_Backup'); e2.props['DELETED_BACKUP_ID'] = 'delbk'; const r2 = promo(e2, { '8th': '9th' }, '2027-28'); ok(r2.status === 'ok' && r2.extra.backupExternal === true && b.sheets.some(s => s.name === r2.extra.backupSheet), 'बाह्य Backup फाइलमध्ये');
};
T.promotion_aborts_when_backup_incomplete = () => {
  const e = mk(); e.run('makeStudentsBackup_ = function(sh, name, prefix) { var ss = SpreadsheetApp.getActiveSpreadsheet(); var cp = sh.copyTo(ss); cp.setName(name); cp.hideSheet(); cp.rows = [cp.rows[0]]; return {ss: ss, sheet: cp, external: false}; };');
  const before = JSON.stringify(e.sheet('Students').rows), plBefore = JSON.stringify(e.sheet('PromotionLog').rows); const r = promo(e, { '8th': '9th' }, '2027-28'); ok(r.status === 'error' && /Backup.*पडताळणीत अयशस्वी/.test(r.message) && /बढती केलेली नाही/.test(r.message), 'अपूर्ण Backup = बढती नाही: ' + r.message);
  eq(JSON.stringify(e.sheet('Students').rows), before, 'Students बदलले नाही'); eq(JSON.stringify(e.sheet('PromotionLog').rows), plBefore, 'PromotionLog बदलले नाही'); ok(!e.ss.sheets.some(s => /BeforePromotion/.test(s.name)), 'अपूर्ण Backup Sheet काढली');
  const e2 = mk(); e2.run('makeStudentsBackup_ = function() { throw new Error("Drive बंद"); };'); const r2 = promo(e2, { '8th': '9th' }, '2027-28'); ok(r2.status === 'error', 'Backup बनवताना अपवाद = बढती नाही'); eq(JSON.stringify(e2.sheet('Students').rows.map(x => x[1])), JSON.stringify(mk().sheet('Students').rows.map(x => x[1])), 'column B बदलला नाही');
};
T.only_intended_functions_changed = () => {
  if (!GS_BASE) { console.log('  (वगळले: SGS_GS_BASE दिलेला नाही)'); return; }
  const fs = require('fs'); const funcs = f => { const s = fs.readFileSync(f, 'utf8'), o = {}; for (const m of s.matchAll(/^function (\w+)\(/gm)) { const i = m.index, j = s.indexOf('\n}\n', i); o[m[1]] = s.slice(i, j + 3); } return o; };
  const a = funcs(GS_BASE), b = funcs(GS); const changed = Object.keys(a).filter(k => b[k] && a[k] !== b[k]).sort(), added = Object.keys(b).filter(k => !a[k]).sort();
  eq(changed, ['doGetCore_', 'doPromoteStudents', 'handleAction'], 'बदललेले विद्यमान फंक्शन्स फक्त हेच (ping/route, बढती, Save)'); eq(added, ['acYearInfo_', 'doGetAcYearInfo', 'doSyncStudentsAcYear'], 'नवे फंक्शन्स');
  ['doStartNewYear', 'doArchiveYear', 'doArchivePassedOut', 'doUndoPromotion', 'currentAcYear_', 'makeStudentsBackup_', 'dqCompute_'].forEach(k => ok(a[k] === b[k], k + ' अपरिवर्तित'));
};
T.version = () => { const e = makeEnv(); e.load(GS); eq(e.run('SERVER_VERSION'), 'V19.65', 'SERVER_VERSION'); };
for (const k of Object.keys(T)) { const before = fails; try { T[k](); } catch (err) { fails++; console.log('  EXCEPTION in ' + k + ': ' + (err && err.stack ? err.stack.split('\n').slice(0, 3).join(' | ') : err)); } console.log((fails === before ? 'ok   ' : 'FAIL ') + k); }
console.log(`\n${fails ? 'FAIL' : 'PASS'}: test_server_v1965 (${passes} तपासण्या पास, ${fails} अपयशी)`); process.exit(fails ? 1 : 0);
