// V19.61 सर्व्हर चाचण्या — अ-४ (ClientErrors) आणि अ-६ (Audit व्याप्ती). बनावट (mock) Apps Script वर; खऱ्या Sheet/Drive वर नाही.
// वापर: SGS_GS=<Code.gs चा path> node tests/test_server_v1961.js
const path = require('path');
const { makeEnv } = require('./lib/gas_mock');
const GS = process.env.SGS_GS || path.join(__dirname, '..', 'apps_script', 'Code.gs');
let fails = 0, passes = 0;
function ok(c, m) { if (c) passes++; else { fails++; console.log('  FAIL: ' + m); } }
function eq(a, b, m) { ok(JSON.stringify(a) === JSON.stringify(b), m + ' → मिळाले ' + JSON.stringify(a) + ' अपेक्षित ' + JSON.stringify(b)); }

function fresh() {
  const env = makeEnv(); env.load(GS); const c = env.ctx;
  env.seed('Students', c.MAIN_HEADERS, [
    ['t', '2026-27', 'R101', 'S1', '', '123456789012', '8th', 'A', '', '1', 'राम पाटील', 'सीता', 'Male', '', '', '', '2012-05-04', '', '', '', '', '', '', '', '', '', '', '9822011111', '9822022222', '', ''],
    ['t', '2026-27', 'R102', 'S2', '', '', '8th', 'A', '', '2', 'श्याम देशमुख', 'गीता', 'Male', '', '', '', '2012-06-04', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['t', '2026-27', 'R103', 'S3', '', '', '8th', 'A', '', '3', 'अनिता जाधव', 'रेखा', 'Female', '', '', '', '2012-07-04', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
  ]);
  env.seed('Users', c.USER_HEADERS, [['boss', 'Boss@1234', 'super', 'Boss', '', ''], ['tch1', 'Tch@12345', 'teacher', 'T1', '8th|A', '9811111111']]);
  env.run('makeStudentsBackup_ = function() { return {}; };');   // Drive प्रत mock मध्ये नाही
  return env;
}
const T = {};
const E = (o) => JSON.stringify([Object.assign({ t: 'error', m: 'x', s: 'at f (a.js:1:1)', c: 1, pg: 'lc', v: 'V19.61', ua: 'UA' }, o)]);

// ---------- अ-४ ----------
T.noToken_notLoggedOut = () => {
  const e = fresh(); const r = e.get({ action: 'logClientErrors', errors: E({}) });
  ok(r.status === 'error' && r.code !== 'auth', 'Token नसताना त्रुटी हवी पण code "auth" नको (क्लायंट Logout होऊ नये) — ' + r.code);
  ok(!e.sheet('ClientErrors'), 'Token नसताना Sheet बनू नये');
};
T.allRoles_canLog = () => {
  ['super', 'master', 'deo', 'cert', 'teacher', 'peon'].forEach(role => {
    const e = fresh(); const r = e.get({ action: 'logClientErrors', token: e.token('u_' + role, role, role === 'teacher' ? { c: '8th|A' } : {}), errors: E({ m: 'm_' + role }) });
    ok(r.status === 'ok' && r.saved === 1, role + ': त्रुटी नोंदवता आली पाहिजे — ' + JSON.stringify(r));
    const s = e.sheet('ClientErrors'); ok(s && s.rows[1][1] === 'u_' + role && s.rows[1][2] === role, role + ': User/Role Token वरून आले पाहिजेत');
  });
};
T.scrub = () => {
  const e = fresh(); const tk = e.token('boss', 'super');
  const r = e.get({ action: 'logClientErrors', token: tk, errors: E({ m: 'Fail for राम 9822011111 and 123456789012 mail a.b@x.com regNo=R101 studentId: S9', s: 'at x?regNo=R102&mobile=9822022222 (a.js:1)' }) });
  ok(r.status === 'ok', 'scrub: ok'); const row = e.sheet('ClientErrors').rows[1], blob = JSON.stringify(row);
  ok(!/9822011111|123456789012|9822022222|a\.b@x\.com|R101|R102/.test(blob), 'scrub: फोन/आधार/ईमेल/Reg No नोंदीत येऊ नयेत — ' + blob);
  const long = 'A'.repeat(2000); e.get({ action: 'logClientErrors', token: tk, errors: E({ m: long, s: long, ua: long, pg: long, v: long }) });
  const r2 = e.sheet('ClientErrors').rows[2]; ok(r2[7].length <= 300 && r2[8].length <= 500 && r2[5].length <= 60 && r2[4].length <= 40 && r2[3].length <= 20, 'लांबी-मर्यादा');
  e.get({ action: 'logClientErrors', token: tk, errors: E({ c: 999999 }) }); eq(e.sheet('ClientErrors').rows[3][9], 9999, 'Count मर्यादा');
  e.get({ action: 'logClientErrors', token: tk, errors: E({ c: 'abc' }) }); eq(e.sheet('ClientErrors').rows[4][9], 1, 'Count अवैध = 1');
};
T.batchAndHourlyLimit = () => {
  const e = fresh(); const tk = e.token('u1', 'teacher', { c: '8th|A' });
  const many = JSON.parse(E({})); const arr = []; for (let i = 0; i < 25; i++) arr.push(Object.assign({}, many[0], { m: 'm' + i }));
  const r = e.get({ action: 'logClientErrors', token: tk, errors: JSON.stringify(arr) }); eq(r.saved, 10, 'एका विनंतीत फक्त १० नोंदी');
  for (let k = 0; k < 15; k++) e.get({ action: 'logClientErrors', token: tk, errors: JSON.stringify(arr) });
  const n = e.sheet('ClientErrors').getLastRow() - 1; ok(n <= e.ctx.CLIENT_ERR_HOURLY_MAX, 'तासाची दर-मर्यादा (' + n + ' ≤ ' + e.ctx.CLIENT_ERR_HOURLY_MAX + ')');
  const r2 = e.get({ action: 'logClientErrors', token: tk, errors: E({}) }); ok(r2.status === 'ok' && r2.saved === 0 && r2.limited, 'मर्यादेनंतर शांतपणे ok (क्लायंट पुन्हा पाठवू नये)');
  const e2 = fresh(); const r3 = e2.get({ action: 'logClientErrors', token: e2.token('x', 'deo'), errors: 'not json' }); ok(r3.status === 'error', 'खराब JSON = त्रुटी, अपवाद नाही');
};
T.writesOnlyClientErrors = () => {
  const e = fresh(); const before = e.snapshot(['ClientErrors']);
  e.get({ action: 'logClientErrors', token: e.token('boss', 'super'), errors: E({}) });
  eq(e.snapshot(['ClientErrors']), before, 'ClientErrors शिवाय कोणतीही Sheet बदलू नये (Log_ सह)');
  ok(!e.ss.sheets.some(s => /^Log_/.test(s.name)), 'Activity Log मध्ये नोंद नको (अनंत-फेरी/गर्दी टाळण्यासाठी)');
};
T.lockBusy_noThrow = () => {
  const e = fresh(); e.lockFail = true;
  const r = e.get({ action: 'logClientErrors', token: e.token('boss', 'super'), errors: E({}) });
  ok(r.status === 'error' && r.code === 'busy', 'Lock न मिळाल्यास code=busy (क्लायंट रांगेत ठेवतो) — ' + JSON.stringify(r));
  ok(!e.sheet('ClientErrors') || e.sheet('ClientErrors').getLastRow() < 2, 'Lock नसताना काही लिहू नये');
};
T.getClientErrors_roles_grouping = () => {
  const e = fresh(); const a = e.token('u1', 'teacher', { c: '8th|A' }), b = e.token('u2', 'deo');
  e.get({ action: 'logClientErrors', token: a, errors: E({ m: 'Boom', c: 2 }) });
  e.get({ action: 'logClientErrors', token: b, errors: E({ m: 'Boom', c: 3, v: 'V19.60' }) });
  e.get({ action: 'logClientErrors', token: b, errors: E({ m: 'Other', t: 'promise' }) });
  ['master', 'deo', 'cert', 'teacher', 'peon'].forEach(role => { const r = e.get({ action: 'getClientErrors', token: e.token('z', role, role === 'teacher' ? { c: '8th|A' } : {}) }); ok(r.status === 'error', role + ': getClientErrors नाकारला पाहिजे'); });
  ok(e.get({ action: 'getClientErrors' }).status === 'error', 'Token नसताना नाकारला');
  const r = e.get({ action: 'getClientErrors', token: e.token('boss', 'super') });
  ok(r.status === 'ok', 'super ला मिळतो'); const g = r.data.find(x => x.message === 'Boom');
  ok(g && g.count === 5 && g.users === 2 && /V19\.60/.test(g.vers) && /V19\.61/.test(g.vers), 'गट: संख्या ५, user २, आवृत्त्या — ' + JSON.stringify(g));
  eq(r.data.length, 2, '२ गट'); eq(r.total, 6, 'एकूण संख्या');
};
T.growthCapAndPurge = () => {
  const e = fresh(); e.run('CLIENT_ERR_MAX_ROWS = 5;'); const tk = e.token('boss', 'super');
  for (let i = 0; i < 4; i++) { e.get({ action: 'logClientErrors', token: tk, errors: JSON.stringify([{ t: 'e', m: 'g' + i, c: 1 }, { t: 'e', m: 'h' + i, c: 1 }]) }); }
  eq(e.sheet('ClientErrors').getLastRow() - 1, 5, 'वाढ-मर्यादा: ५ ओळींवर थांबते (जुन्या पुसल्या)');
  ok(JSON.stringify(e.sheet('ClientErrors').rows).indexOf('"g0"') === -1, 'सर्वात जुनी ओळ पुसली गेली');
  const e2 = fresh(); const sh = e2.seed('ClientErrors', e2.ctx.CLIENT_ERR_HEADERS, []);
  const old = e2.ctx.Utilities.formatDate(new Date(Date.now() - 100 * 86400000), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'), old2 = e2.ctx.Utilities.formatDate(new Date(Date.now() - 91 * 86400000), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'), recent = e2.ctx.Utilities.formatDate(new Date(Date.now() - 80 * 86400000), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'), nw = e2.ctx.Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
  [old, old2, recent, nw].forEach((t, i) => sh.appendRow([t, 'u', 'r', 'v', 'p', 'ua', 'e', 'm' + i, '', 1]));
  const k = e2.run('purgeClientErrors()'); eq(k, 2, '९० दिवसांपेक्षा जुन्या २ ओळी पुसल्या'); eq(sh.getLastRow() - 1, 2, 'नव्या २ ओळी शिल्लक');
  eq(e2.run('purgeClientErrors()'), 0, 'पुन्हा चालवल्यास काहीही पुसले जात नाही');
};

// ---------- अ-६ ----------
const adm = (e, role) => ({ token: e.token('boss', role || 'super') });
T.audit_contact_fullPhone = () => {
  const e = fresh(); const tk = e.token('t1', 'teacher', { c: '8th|A' });
  let r = e.post({ action: 'saveStudentContact', token: tk, regNo: 'R101', fullName: 'राम', iyatta: '8th', tukdi: 'A', whatsappMobile: '9822011111', otherMobile: '9822022222' });
  ok(r.status === 'ok', 'saveStudentContact ok');
  let a = e.sheet('Student_Audit'); ok(a && a.getLastRow() === 3, 'नवीन संपर्क = २ ओळी (WhatsApp + Alternate) — ' + (a && a.getLastRow()));
  ok(a.rows[1][3] === 'R101' && a.rows[1][4] === 'saveStudentContact' && a.rows[1][1] === 't1' && a.rows[1][2] === 'teacher', 'User/Role/Reg/Action');
  ok(a.rows[1][5] === 'WhatsApp Mobile No' && a.rows[1][7] === '9822011111', 'फोन पूर्ण — मास्क नाही: ' + a.rows[1][7]);
  ok(a.rows[2][7] === '9822022222', 'पर्यायी फोन पूर्ण');
  r = e.post({ action: 'saveStudentContact', token: tk, regNo: 'R101', fullName: 'राम', iyatta: '8th', tukdi: 'A', whatsappMobile: '9822011111', otherMobile: '9822022222' });
  eq(e.sheet('Student_Audit').getLastRow(), 3, 'समान मूल्याने खोटा बदल नोंदवू नये');
  e.post({ action: 'saveStudentContact', token: tk, regNo: 'R101', fullName: 'राम', iyatta: '8th', tukdi: 'A', whatsappMobile: '9000000001', otherMobile: '9822022222' });
  const last = e.sheet('Student_Audit').rows[3]; ok(last[6] === '9822011111' && last[7] === '9000000001', 'जुना → नवा पूर्ण क्रमांक: ' + last[6] + ' → ' + last[7]);
};
T.audit_failureDoesNotBlockSave = () => {
  const e = fresh(); e.run('auditAppendRows_ = function() { throw new Error("Audit Sheet बंद"); };');
  const r = e.post({ action: 'saveStudentContact', token: e.token('t1', 'teacher', { c: '8th|A' }), regNo: 'R101', fullName: 'राम', iyatta: '8th', tukdi: 'A', whatsappMobile: '9822011111', otherMobile: '' });
  ok(r.status === 'ok', 'Audit अपयशी असताना मूळ Save ok — ' + JSON.stringify(r));
  ok(e.sheet('StudentContacts') && e.sheet('StudentContacts').getLastRow() === 2, 'संपर्क प्रत्यक्ष Save झाला');
};
T.audit_classDivision = () => {
  const e = fresh(); const r = e.post({ action: 'updateClassDivision', token: e.token('boss', 'super'), regNo: 'R102', iyatta: '9th', tukdi: 'B' });
  ok(r.status === 'ok', 'ok'); const a = e.sheet('Student_Audit'); eq(a.getLastRow(), 3, '२ ओळी');
  ok(a.rows[1][5] === 'Class' && a.rows[1][6] === '8th' && a.rows[1][7] === '9th' && a.rows[2][5] === 'Division' && a.rows[2][7] === 'B', 'वर्ग/तुकडी जुना→नवा');
  e.post({ action: 'updateClassDivision', token: e.token('boss', 'super'), regNo: 'R102', iyatta: '9th', tukdi: 'B' }); eq(e.sheet('Student_Audit').getLastRow(), 3, 'समान मूल्य = नोंद नाही');
};
T.audit_deleteStudent = () => {
  const e = fresh(); const r = e.get({ action: 'deleteStudent', token: e.token('boss', 'master'), regNo: 'R103' });
  ok(r.status === 'ok', 'delete ok — ' + JSON.stringify(r)); const a = e.sheet('Student_Audit');
  ok(a && a.rows[1][3] === 'R103' && a.rows[1][4] === 'deleteStudent' && /अनिता/.test(a.rows[1][6]), 'Delete घटना नोंदली: ' + (a && JSON.stringify(a.rows[1])));
};
T.audit_photoLink_userFromToken = () => {
  const e = fresh(); e.get({ action: 'ping', token: e.token('ph', 'master') });   // applyAuth_ → AUDIT_CTX_ (ping public — Token तपासत नाही)
  e.run('applyAuth_({action:"getAll", token:"' + e.token('ph', 'master') + '"})');
  const res = e.run('updateStudentPhotoUrl("R101","https://drive.google.com/uc?id=AAA")'); ok(res.updated, 'लिंक Save');
  e.run('updateStudentPhotoUrl("R101","https://drive.google.com/uc?id=BBB")');
  const a = e.sheet('Student_Audit'); ok(a.rows[1][1] === 'ph' && a.rows[1][2] === 'master' && a.rows[1][4] === 'photoLinkSet' && a.rows[1][5] === 'Photo URL', 'पहिली = photoLinkSet, user Token वरून: ' + JSON.stringify(a.rows[1]));
  ok(a.rows[2][4] === 'photoLinkChange' && /AAA/.test(a.rows[2][6]) && /BBB/.test(a.rows[2][7]), 'बदल = जुनी→नवी');
  const n = a.getLastRow(); e.run('updateStudentPhotoUrl("R101","https://drive.google.com/uc?id=BBB")'); eq(e.sheet('Student_Audit').getLastRow(), n, 'तीच लिंक = नोंद नाही');
};
T.audit_renumber_archive_restore = () => {
  const e = fresh(); const tk = e.token('boss', 'super');
  // Renumber: क्रम नावानुसार — "अनिता", "राम", "श्याम" → रोल बदलतात
  let r = e.post({ action: 'renumberRolls', token: tk, confirm: 'RENUMBER', classesJson: JSON.stringify(['8th']), order: 'name' });
  ok(r.status === 'ok', 'renumber ok — ' + JSON.stringify(r)); let a = e.sheet('Student_Audit');
  ok(a && a.rows.slice(1).every(x => x[5] === 'Roll No' && x[4] === 'renumberRolls'), 'सर्व ओळी Roll No — ' + (a && JSON.stringify(a.rows.slice(1))));
  const changed = a.getLastRow() - 1; eq(changed, r.count, 'ओळी = बदललेले रोल (' + r.count + ')');
  const before = a.getLastRow(); e.post({ action: 'renumberRolls', token: tk, confirm: 'RENUMBER', classesJson: JSON.stringify(['8th']), order: 'name' });
  eq(e.sheet('Student_Audit').getLastRow(), before, 'पुन्हा तेच क्रमांक = नवीन नोंद नाही');
  // Archive
  r = e.post({ action: 'archivePassedOut', token: tk, confirm: 'ARCHIVE', iyatta: '8th', label: '2026-27' }); ok(r.status === 'ok', 'archive ok — ' + JSON.stringify(r));
  a = e.sheet('Student_Audit'); const arch = a.rows.filter(x => x[4] === 'archivePassedOut'); eq(arch.length, 3, '३ विद्यार्थ्यांच्या Archive ओळी'); ok(arch.every(x => x[5] === 'Status' && /PassedOut_/.test(x[7])), 'Status Students→PassedOut_');
  const batch = r.extra.batchId;
  r = e.post({ action: 'restorePassedOut', token: tk, batchId: batch }); ok(r.status === 'ok', 'restore ok — ' + JSON.stringify(r));
  const res = e.sheet('Student_Audit').rows.filter(x => x[4] === 'restorePassedOut'); eq(res.length, 3, '३ Restore ओळी');
};
T.audit_bulkCap = () => {
  const e = fresh(); const ch = []; for (let i = 0; i < 700; i++) ch.push({ reg: 'R' + i, field: 'Roll No', old: '1', nw: '2' });
  const n = e.run('auditApply_({auditUser:"u",auditRole:"super"}, "renumberRolls", ' + JSON.stringify(ch) + ')');
  eq(n, 601, '६०० + १ सारांश-ओळ'); const a = e.sheet('Student_Audit'); ok(/एकूण 700/.test(a.rows[601][7]), 'सारांश-ओळीत एकूण संख्या');
};
T.audit_aadhar_masked_still = () => {
  const e = fresh(); e.post({ action: 'upsert', token: e.token('boss', 'super'), regNo: 'R101', acYear: '2026-27', studentId: 'S1', aadhar: '999988887777', iyatta: '8th', tukdi: 'A', firstName: 'राम पाटील', whatsappMobile: '9822011111', alternateMobile: '9822022222' });
  const t = e.allText('Student_Audit'); ok(/XXXX-XXXX-7777/.test(t) && !/999988887777/.test(t), 'आधार फक्त शेवटचे ४ अंक (जुना नियम कायम)');
};
T.users_audit_noPassword = () => {
  [false, true].forEach(hard => {
    const e = fresh(); if (hard) { e.props['AUTH_ENFORCE'] = '1'; e.props['PW_HASHED'] = '1'; }
    const tk = e.token('boss', 'super'); const SECRET = 'Zz!SecretPw99';
    let r = e.get({ action: 'saveUser', token: tk, username: 'newu', password: SECRET, role: 'deo', label: 'New U', assignedClass: '', mobile: '9876543210' });
    ok(r.status === 'ok', 'saveUser(नवा) ok — ' + JSON.stringify(r));
    r = e.get({ action: 'saveUser', token: tk, username: 'tch1', password: hard ? '' : 'Tch@12345', role: 'teacher', label: 'T1 renamed', assignedClass: '9th|B', mobile: '9811111111' }); ok(r.status === 'ok', 'saveUser(बदल)');
    r = e.get({ action: 'saveUser', token: tk, username: 'tch1', password: SECRET + 'x', role: 'teacher', label: 'T1 renamed', assignedClass: '9th|B', mobile: '9811111111' }); ok(r.status === 'ok', 'Password बदल');
    r = e.get({ action: 'deleteUser', token: tk, username: 'newu' }); ok(r.status === 'ok', 'deleteUser');
    const u = e.sheet('User_Audit'); ok(!!u, 'User_Audit बनली'); const t = JSON.stringify(u.rows);
    ok(t.indexOf(SECRET) === -1 && t.indexOf('Tch@12345') === -1 && !/h1\$/.test(t), (hard ? '[hash] ' : '[plain] ') + 'User_Audit मध्ये Password/hash कुठेही नाही');
    ok(u.rows.some(x => x[3] === 'newu' && x[4] === 'addUser'), 'नवा User घटना');
    ok(u.rows.some(x => x[3] === 'tch1' && x[5] === 'Label' && x[7] === 'T1 renamed'), 'Label बदल');
    ok(u.rows.some(x => x[3] === 'tch1' && x[5] === 'AssignedClass' && x[6] === '8th|A' && x[7] === '9th|B'), 'वर्ग-नेमणूक बदल');
    ok(u.rows.some(x => x[3] === 'tch1' && x[5] === 'Password' && x[7] === 'बदलला'), '"Password बदलला" घटना');
    ok(u.rows.some(x => x[3] === 'newu' && x[4] === 'deleteUser'), 'User हटवला घटना');
    ok(u.rows.slice(1).every(x => x[1] === 'boss' && x[2] === 'super'), 'कोणी (By/ByRole) नोंद');
  });
};
T.users_changePassword_event_only = () => {
  const e = fresh(); const r = e.get({ action: 'changePassword', token: e.token('tch1', 'teacher', { c: '8th|A' }), username: 'tch1', oldPassword: 'Tch@12345', newPassword: 'Brand-New-Pw-77' });
  ok(r.status === 'ok', 'changePassword ok — ' + JSON.stringify(r)); const t = e.allText('User_Audit');
  ok(/बदलला/.test(t) && !/Brand-New-Pw-77|Tch@12345/.test(t), 'फक्त घटना; Password नाही');
};
T.userAudit_guard_blocksPasswordValue = () => {
  // सहाय्यकाला चुकून Password मूल्य दिले तरी तो कधीही साठवला जाऊ नये (defense-in-depth)
  const e = fresh(); e.run('userAuditApply_({requesterUser:"boss",requesterRole:"super"}, "saveUser", "x", [{field:"Password", old:"OldSecret1", nw:"NewSecret2"}, {field:"passwd", old:"a", nw:"b"}])');
  const t = e.allText('User_Audit'); ok(t && !/OldSecret1|NewSecret2/.test(t) && /बदलला/.test(t), 'Password मूल्य नोंदवले जाऊ नये — ' + t);
};
T.getUserAudit_roles = () => {
  const e = fresh(); e.get({ action: 'deleteUser', token: e.token('boss', 'super'), username: 'tch1' });
  ['master', 'deo', 'teacher'].forEach(role => ok(e.get({ action: 'getUserAudit', token: e.token('z', role, role === 'teacher' ? { c: '8th|A' } : {}) }).status === 'error', role + ' नाकारला'));
  const r = e.get({ action: 'getUserAudit', token: e.token('boss', 'super') }); ok(r.status === 'ok' && r.data.length >= 1 && r.data[0].username === 'tch1', 'super ला मिळतो');
};
T.profileHistory_includesNewSources = () => {
  const e = fresh(); e.post({ action: 'updateClassDivision', token: e.token('boss', 'super'), regNo: 'R102', iyatta: '9th', tukdi: 'B' });
  e.post({ action: 'saveStudentContact', token: e.token('boss', 'super'), regNo: 'R102', fullName: 'श्याम', iyatta: '9th', tukdi: 'B', whatsappMobile: '9000000002', otherMobile: '' });
  const r = e.get({ action: 'getStudentAudit', token: e.token('boss', 'super'), regNo: 'R102' });
  ok(r.status === 'ok' && r.changes.length >= 3, 'getStudentAudit मध्ये नव्या स्रोतांचे बदल (' + (r.changes && r.changes.length) + ')');
  ok(r.changes.some(c => c.action === 'saveStudentContact' && c.newVal === '9000000002'), 'संपर्क बदल दिसतो'); ok(r.changes.some(c => c.action === 'updateClassDivision' && c.field === 'Class'), 'वर्ग बदल दिसतो');
};
T.version = () => { const e = fresh(); eq(e.run('SERVER_VERSION'), 'V19.61', 'SERVER_VERSION'); };

for (const k of Object.keys(T)) { const before = fails; try { T[k](); } catch (err) { fails++; console.log('  EXCEPTION in ' + k + ': ' + (err && err.stack ? err.stack.split('\n').slice(0, 3).join(' | ') : err)); } console.log((fails === before ? 'ok   ' : 'FAIL ') + k); }
console.log(`\n${fails ? 'FAIL' : 'PASS'}: test_server_v1961 (${passes} तपासण्या पास, ${fails} अपयशी)`);
process.exit(fails ? 1 : 0);
