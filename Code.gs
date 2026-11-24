/**
 * الاختبار العملي لمنصة رواء – يستقبل تسليمات الموظفين ودرجات المصحح من صفحة GitHub Pages ويكتبها في الشيت.
 * التبويبات: الدرجات · التسليمات · ملاحظات التصحيح · الأسئلة · الإعدادات
 */
const SS_ID = '1abxknP1PjkXHuEEq-V1zO_s2iCEP7xRWNEPM-b7jjF4';
const SH = { G: 'الدرجات', T: 'التسليمات', N: 'ملاحظات التصحيح', Q: 'الأسئلة', S: 'الإعدادات' };
const TZ = 'Asia/Riyadh';
const NQ = 85;            // عدد الأسئلة
const QCOL = 7;           // عمود Q1 في تبويب الدرجات (G)
const SECS = [['A', 1, 10], ['B', 11, 28], ['C', 29, 38], ['D', 39, 50], ['E', 51, 59], ['F', 60, 76], ['G', 77, 85]];
const ST = { LIVE: 'يحل الآن', D1: 'أنهى اليوم الأول', WAIT: 'بانتظار التصحيح', PART: 'قيد التصحيح', DONE: 'تم التصحيح' };
const TC = 18; // أعمدة تبويب التسليمات (O بدء 1 · P نهاية 1 · Q بدء 2 · R المدة)
const FMT = 'yyyy-MM-dd HH:mm:ss';

function ss_() { try { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SS_ID); } catch (e) { return SpreadsheetApp.openById(SS_ID); } }
function sh_(k) { return ss_().getSheetByName(SH[k]); }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function txt_(v) { v = String(v == null ? '' : v).trim(); return v === '' ? '' : "'" + v; }
function col_(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

function settings_() {
  const v = sh_('S').getRange('A1:B10').getDisplayValues();
  const m = {}; v.forEach(r => { if (r[0]) m[r[0].trim()] = r[1]; });
  return {
    title: m['عنوان الاختبار'] || 'Rewaa Platform Practical Assessment',
    deadline: m['الموعد النهائي'] || '',
    pass: m['نسبة النجاح'] || '80%',
    code: String(m['رمز المصحح'] || '').trim(),
    open: (m['التسليم مفتوح'] || 'نعم').trim() !== 'لا',
    day1Last: Number(m['آخر سؤال في اليوم الأول']) || 50
  };
}

function questions_() {
  const sh = sh_('Q'); const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 7).getDisplayValues()
    .filter(r => r[0])
    .map(r => ({ n: Number(r[0]), s: r[1], sa: r[2], t: r[3], ta: r[4], c: r[5], ca: r[6] }));
}

function doGet() {
  const s = settings_();
  return out_({ ok: true, title: s.title, deadline: s.deadline, pass: s.pass, open: s.open, day1Last: s.day1Last, questions: questions_() });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (b.action === 'start') return out_(start_(b));
    if (b.action === 'endDay1') return out_(endDay1_(b));
    if (b.action === 'startDay2') return out_(startDay2_(b));
    if (b.action === 'resume') return out_(resume_(b));
    if (b.action === 'submit') return out_(submit_(b));
    const s = settings_();
    if (!s.code || String(b.code || '').trim() !== s.code) return out_({ ok: false, error: 'code' });
    if (b.action === 'list') return out_(list_());
    if (b.action === 'load') return out_(load_(b.id));
    if (b.action === 'save') return out_(save_(b));
    if (b.action === 'add') return out_(add_(b));
    return out_({ ok: false, error: 'unknown' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/* ---------- بدء الاختبار واليومين والتسليم ---------- */
function newId_() { return 'PA-' + Utilities.formatDate(new Date(), TZ, 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 90 + 10); }
function stamp_(d) { return Utilities.formatDate(d, TZ, FMT); }
function ms_(t) { return t ? Utilities.parseDate(t, TZ, FMT).getTime() : 0; }

function start_(b) {
  const s = settings_();
  if (!s.open) return { ok: false, error: 'closed' };
  const name = String(b.name || '').trim();
  if (name.length < 3) return { ok: false, error: 'missing' };
  const id = newId_(), t = new Date();
  sh_('T').appendRow([id, '', name, b.team || '', txt_(b.email), '', '', '', '', '', '', '', '', ST.LIVE, txt_(stamp_(t)), '', '', '']);
  return { ok: true, id: id, day: 1, startedAt: t.getTime() };
}

function findT_(id) {
  const sh = sh_('T'); const n = sh.getLastRow() - 1;
  if (n < 1 || !id) return 0;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0] === String(id).trim()) return i + 2;
  return 0;
}
function rowT_(r) { return sh_('T').getRange(r, 1, 1, TC).getDisplayValues()[0]; }

function state_(c) {
  const day = c[13] === ST.LIVE ? (c[16] ? 2 : 1) : c[13] === ST.D1 ? 'between' : 'submitted';
  return { ok: true, id: c[0], name: c[2], team: c[3], email: c[4], status: c[13], day: day,
    startedAt: day === 2 ? ms_(c[16]) : day === 1 ? ms_(c[14]) : 0,
    day1Ms: c[15] ? ms_(c[15]) - ms_(c[14]) : 0 };
}

function resume_(b) {
  const r = findT_(b.id);
  if (!r) return { ok: false, error: 'not_found' };
  return state_(rowT_(r));
}

function endDay1_(b) {
  const r = findT_(b.id); if (!r) return { ok: false, error: 'not_found' };
  const c = rowT_(r);
  if (c[13] === ST.LIVE && !c[15] && !c[16]) {
    sh_('T').getRange(r, 14).setValue(ST.D1);
    sh_('T').getRange(r, 16).setValue(txt_(stamp_(new Date())));
  }
  return state_(rowT_(r));
}

function startDay2_(b) {
  const r = findT_(b.id); if (!r) return { ok: false, error: 'not_found' };
  const c = rowT_(r);
  if (c[13] === ST.D1) {
    sh_('T').getRange(r, 14).setValue(ST.LIVE);
    sh_('T').getRange(r, 17).setValue(txt_(stamp_(new Date())));
  }
  return state_(rowT_(r));
}

function submit_(b) {
  const s = settings_();
  if (!s.open) return { ok: false, error: 'closed' };
  if (!b.login || !b.password || !b.email) return { ok: false, error: 'missing' };
  const sh = sh_('T'); const now = new Date(); const whenS = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm');
  let r = findT_(b.id), id = b.id, c = null, mins = '';
  if (r) {
    c = rowT_(r);
    if (c[13] !== ST.LIVE && c[13] !== ST.D1) return { ok: true, id: id, already: true };
    const d1s = ms_(c[14]), d1e = ms_(c[15]), d2s = ms_(c[16]);
    let tot = 0;
    if (d1e) tot += d1e - d1s; else tot += now.getTime() - d1s;
    if (d2s) tot += now.getTime() - d2s;
    mins = Math.round(tot / 6000) / 10;
  } else {
    id = newId_(); r = sh.getLastRow() + 1;
    c = [id, '', String(b.name || '').trim(), b.team || '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''];
  }
  const name = c[2] || String(b.name || '').trim(), team = c[3] || b.team || '';
  sh.getRange(r, 1, 1, TC).setValues([[id, whenS, name, team, txt_(b.email), txt_(b.login), txt_(b.password), txt_(b.finished),
    txt_(b.skipped), txt_(b.r82), txt_(b.r83), txt_(b.r84), txt_(b.r85), ST.WAIT, txt_(c[14]), txt_(c[15]), txt_(c[16]), mins]]);
  addGradeRow_(id, name, team, whenS);
  return { ok: true, id: id, minutes: mins };
}

/* إضافة موظف يدوياً من شاشة المصحح (ما حل من خلال الرابط) */
function add_(b) {
  const name = String(b.name || '').trim();
  if (name.length < 3) return { ok: false, error: 'missing' };
  const id = newId_(); const whenS = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
  const sh = sh_('T'); const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, TC).setValues([[id, whenS, name, b.team || '', txt_(b.email), txt_(b.login), txt_(b.password), txt_(b.finished),
    txt_(b.skipped), txt_(b.r82), txt_(b.r83), txt_(b.r84), txt_(b.r85), ST.WAIT, '', '', '', txt_(b.minutes)]]);
  addGradeRow_(id, name, b.team || '', whenS);
  return { ok: true, id: id };
}

function addGradeRow_(id, name, team, when) {
  const sh = sh_('G');
  const r = Math.max(sh.getLastRow(), 2) + 1;
  const row = [id, name, team, when, '', ''];
  for (let i = 0; i < NQ; i++) row.push('');
  SECS.forEach(x => row.push('=SUM(' + col_(QCOL + x[1] - 1) + r + ':' + col_(QCOL + x[2] - 1) + r + ')'));
  const first = col_(QCOL), last = col_(QCOL + NQ - 1), sumA = col_(QCOL + NQ), sumG = col_(QCOL + NQ + 6);
  const tot = col_(QCOL + NQ + 7), pct = col_(QCOL + NQ + 8);
  row.push('=SUM(' + sumA + r + ':' + sumG + r + ')');
  row.push('=IF(COUNT(' + first + r + ':' + last + r + ')=0,"",' + tot + r + '/' + NQ + ')');
  row.push('=IF(COUNT(' + first + r + ':' + last + r + ')<' + NQ + ',"' + ST.PART + '",IF(' + pct + r + '>=\'' + SH.S + '\'!$B$4,"ناجح","لم يجتز"))');
  row.push('');
  sh.getRange(r, 1, 1, row.length).setValues([row]);
}

/* ---------- المصحح ---------- */
function gradeIndex_() {
  const sh = sh_('G'); const n = sh.getLastRow() - 2;
  const m = {};
  if (n < 1) return m;
  sh.getRange(3, 1, n, QCOL + NQ + 10).getDisplayValues().forEach((r, i) => {
    if (r[0]) m[r[0]] = { row: i + 3, r: r };
  });
  return m;
}

function list_() {
  const sh = sh_('T'); const n = sh.getLastRow() - 1;
  const g = gradeIndex_();
  const items = [];
  if (n > 0) sh.getRange(2, 1, n, TC).getDisplayValues().forEach(r => {
    if (!r[0]) return;
    const gr = g[r[0]] ? g[r[0]].r : null;
    items.push({
      id: r[0], at: r[1], name: r[2], team: r[3], email: r[4], login: r[5], password: r[6], finished: r[7],
      skipped: r[8], r82: r[9], r83: r[10], r84: r[11], r85: r[12], status: r[13], started: r[14], minutes: r[17],
      total: gr ? gr[QCOL + NQ + 6] : '', pct: gr ? gr[QCOL + NQ + 7] : '', result: gr ? gr[QCOL + NQ + 8] : '',
      graded: gr ? gr.slice(QCOL - 1, QCOL - 1 + NQ).filter(x => x !== '').length : 0
    });
  });
  return { ok: true, items: items.reverse() };
}

function load_(id) {
  const g = gradeIndex_()[id];
  if (!g) return { ok: false, error: 'not_found' };
  const marks = {};
  for (let i = 0; i < NQ; i++) { const v = g.r[QCOL - 1 + i]; if (v !== '') marks[i + 1] = Number(v); }
  const notes = {};
  const ns = sh_('N'); const n = ns.getLastRow() - 1;
  if (n > 0) ns.getRange(2, 1, n, 5).getDisplayValues().forEach(r => {
    if (r[0] === id && r[4]) notes[String(r[2]).replace('Q', '')] = r[4];
  });
  return { ok: true, marks: marks, notes: notes, grader: g.r[4], general: g.r[QCOL + NQ + 9] };
}

function save_(b) {
  const idx = gradeIndex_(); const g = idx[b.id];
  if (!g) return { ok: false, error: 'not_found' };
  const sh = sh_('G'); const r = g.row;
  const marks = b.marks || {}; const vals = [];
  let graded = 0;
  for (let i = 1; i <= NQ; i++) {
    const v = marks[i];
    if (v === 1 || v === 0 || v === '1' || v === '0') { vals.push(Number(v)); graded++; } else vals.push('');
  }
  sh.getRange(r, QCOL, 1, NQ).setValues([vals]);
  sh.getRange(r, 5, 1, 2).setValues([[String(b.grader || '').trim(), now_()]]);
  sh.getRange(r, QCOL + NQ + 10).setValue(String(b.general || '').trim());

  // ملاحظات الأسئلة: نحذف القديمة لهذا التسليم ونكتب الجديدة
  const ns = sh_('N'); const n = ns.getLastRow() - 1;
  if (n > 0) {
    const ids = ns.getRange(2, 1, n, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (ids[i][0] === b.id) ns.deleteRow(i + 2);
  }
  const name = g.r[1]; const when = now_(); const rows = [];
  Object.keys(b.notes || {}).map(Number).sort((a, c) => a - c).forEach(q => {
    const t = String(b.notes[q] || '').trim();
    if (t) rows.push([b.id, name, 'Q' + q, vals[q - 1] === '' ? '' : vals[q - 1], t, when]);
  });
  if (rows.length) ns.getRange(ns.getLastRow() + 1, 1, rows.length, 6).setValues(rows);

  // حالة التسليم
  const ts = sh_('T'); const tn = ts.getLastRow() - 1;
  if (tn > 0) {
    const ids = ts.getRange(2, 1, tn, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (ids[i][0] === b.id) {
      ts.getRange(i + 2, 14).setValue(graded === NQ ? ST.DONE : graded ? ST.PART : ST.WAIT);
      if (b.skipped !== undefined) ts.getRange(i + 2, 9, 1, 5).setValues([[txt_(b.skipped), txt_(b.r82), txt_(b.r83), txt_(b.r84), txt_(b.r85)]]);
      if (b.minutes !== undefined) { const mv = String(b.minutes).trim(); ts.getRange(i + 2, 18).setValue(mv !== '' && !isNaN(Number(mv)) ? Number(mv) : txt_(mv)); }
      break;
    }
  }
  SpreadsheetApp.flush();
  const after = sh.getRange(r, QCOL + NQ + 7, 1, 3).getDisplayValues()[0];
  return { ok: true, graded: graded, total: after[0], pct: after[1], result: after[2] };
}
