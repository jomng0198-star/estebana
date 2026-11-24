/**
 * أفق STEM | منصة العلوم والابتكار — ربط Google Sheets + Gmail + Drive
 * ---------------------------------------------------------------------
 * البيانات تُحفظ في جدول «أفق STEM - بيانات المنصة» (SHEET_ID بالأسفل)
 * وكل ما تفعله الطالبات يصل إلى ADMIN_EMAIL في رسالة مجمّعة كل ساعة
 * (حتى لا يتجاوز حدّ Gmail اليومي)، مع ملخص يومي الساعة ٨ مساءً.
 * بعد أي تعديل: شغّلي setup مرة، ثم نشر ← إدارة عمليات النشر ← تعديل ← إصدار جديد.
 */

const ADMIN_EMAIL = 'asmaalmaliky1400@gmail.com';
const SHEET_ID = '1fAe_tRvWPZFolgfKFECxgZmJ0m3FpjzGzXWmq7hZHvU'; // جدول «أفق STEM - بيانات المنصة»
const APP_NAME = 'أفق STEM | منصة العلوم والابتكار';
const TEACHER_PIN = '2026';          // نفس رمز دخول المعلمة في الصفحة
const DRIVE_FOLDER = 'أفق STEM - ملفات المعلمة';
const DIGEST_EVERY_HOURS = 1;        // كل كم ساعة تصل الرسالة المجمّعة

const SHEETS = {
  view:        ['الزيارات',          ['التاريخ', 'الوقت']],
  login:       ['الدخول',            ['التاريخ', 'الوقت', 'الصفة', 'الاسم', 'الصف', 'الفصل', 'نوع الزائرة']],
  comment:     ['التعليقات',         ['المعرف', 'التاريخ', 'الاسم', 'الصفة', 'التعليق']],
  opinion:     ['آراء أولياء الأمور', ['المعرف', 'التاريخ', 'ولية الأمر', 'الطالبة', 'التقييم', 'الرأي']],
  result:      ['نتائج النماذج',     ['المعرف', 'التاريخ', 'الطالبة', 'الصفة', 'الصف', 'النموذج', 'العنوان', 'الدرجة', 'من', 'النسبة']],
  game:        ['اللعبة',            ['التاريخ', 'الاسم', 'المستوى', 'اسم المستوى', 'النقاط', 'النجوم']],
  certificate: ['الشهادات',          ['التاريخ', 'الاسم', 'النوع', 'أصدرتها', 'الصفة']],
  watch:       ['المشاهدات',         ['التاريخ', 'الاسم', 'الصفة', 'الشرح']],
  plan:        ['الخطط',             ['المعرف', 'الفصل', 'الأسابيع', 'الوحدة', 'المجال', 'النشاط']],
  video:       ['الفيديوهات',        ['المعرف', 'العنوان', 'الرابط', 'المجال', 'التاريخ']],
  worksheet:   ['أوراق العمل',       ['المعرف', 'التاريخ', 'الطالبة', 'الصف', 'المادة', 'المجال', 'الدرجة', 'من', 'إجابة الشرح', 'إجابة الإثراء']],
  activity:    ['نشاط عالم الفيزياء', ['المعرف', 'التاريخ', 'الطالبة', 'الصف', 'المجموعة', 'القسم', 'العنصر', 'النقاط', 'التفاصيل']],
  content:     ['محتوى المعلمة',     ['المعرف', 'النوع', 'البيانات', 'آخر تعديل']],
  mail:        ['سجل البريد',        ['التاريخ', 'النوع', 'النص']]
};

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function sheet_(type) {
  const ss = ss_();
  const [name, head] = SHEETS[type];
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#0d2149').setFontColor('#ffd98a');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

/** شغّليها مرة بعد كل تحديث للكود: تنشئ الأوراق والمجلد وتضبط الرسائل المجمّعة والملخص اليومي */
function setup() {
  Object.keys(SHEETS).forEach(sheet_);
  folder_();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['dailySummary', 'hourlyDigest'].indexOf(t.getHandlerFunction()) > -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(20).create();
  ScriptApp.newTrigger('hourlyDigest').timeBased().everyHours(DIGEST_EVERY_HOURS).create();
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('mailRow')) p.setProperty('mailRow', String(sheet_('mail').getLastRow()));
}

function folder_() {
  const it = DriveApp.getFoldersByName(DRIVE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_FOLDER);
}

/* ---------- استقبال البيانات من المنصة ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const now = new Date();
    const date = Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy-MM-dd');
    const time = Utilities.formatDate(now, 'Asia/Riyadh', 'HH:mm');
    const s = v => String(v == null ? '' : v).slice(0, 2000);
    const teacherOnly = ['content_upsert', 'content_delete', 'upload', 'plan', 'plan_delete', 'video', 'video_delete'];
    if (teacherOnly.indexOf(d.type) > -1 && String(d.pin) !== TEACHER_PIN) return json_({ ok: false, error: 'رمز المعلمة غير صحيح' });

    switch (d.type) {
      case 'view':        sheet_('view').appendRow([date, time]); break;
      case 'login':       sheet_('login').appendRow([date, time, roleAr_(d.role), s(d.name), s(d.grade), s(d.section), s(d.kind)]); break;
      case 'comment':     sheet_('comment').appendRow([s(d.id), now, s(d.name), roleAr_(d.role), s(d.text)]); break;
      case 'opinion':     sheet_('opinion').appendRow([s(d.id), now, s(d.name), s(d.child), Number(d.rating) || 0, s(d.text)]); break;
      case 'result':      sheet_('result').appendRow([s(d.id), now, s(d.name), roleAr_(d.role), s(d.grade), d.model, s(d.title), d.score, d.total, d.pct]); break;
      case 'game':        sheet_('game').appendRow([now, s(d.name), d.level, s(d.levelName), d.points, d.stars]); break;
      case 'certificate': sheet_('certificate').appendRow([now, s(d.name), certAr_(d.certType), s(d.by), roleAr_(d.role)]); break;
      case 'watch':       sheet_('watch').appendRow([now, s(d.name), roleAr_(d.role), s(d.video)]); break;
      case 'plan':        sheet_('plan').appendRow([s(d.id), d.term, s(d.week), s(d.unit), s(d.field), s(d.activity)]); break;
      case 'video':       sheet_('video').appendRow([s(d.id), s(d.title), s(d.url), s(d.field), now]); break;
      case 'worksheet':   sheet_('worksheet').appendRow([s(d.id), now, s(d.name), s(d.grade), s(d.subject), s(d.domain), d.score, d.total, s(d.explain), s(d.enrich)]); break;
      case 'activity':    sheet_('activity').appendRow([s(d.id), now, s(d.name), s(d.grade), s(d.group), s(d.section), s(d.ref), Number(d.points) || 0, s(d.details)]); break;
      case 'plan_delete': deleteById_('plan', d.id); break;
      case 'video_delete':deleteById_('video', d.id); break;
      case 'content_upsert': upsertContent_(s(d.id), s(d.kind), d.data || {}, now); break;
      case 'content_delete': deleteById_('content', d.id); break;
      case 'upload':      return json_(upload_(d));
    }
    if (d.type !== 'view' && !d.quiet && String(d.type).indexOf('content_') !== 0 && String(d.type).indexOf('_delete') < 0) logMail_(d, now);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function upsertContent_(id, kind, data, now) {
  const sh = sheet_('content'), v = sh.getDataRange().getValues(), txt = JSON.stringify(data);
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === id) { sh.getRange(i + 1, 2, 1, 3).setValues([[kind, txt, now]]); return; }
  sh.appendRow([id, kind, txt, now]);
}

function upload_(d) {
  const bytes = Utilities.base64Decode(d.data);
  const blob = Utilities.newBlob(bytes, d.mime || 'application/octet-stream', d.name || 'ملف');
  const f = folder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: f.getId(), url: 'https://drive.google.com/file/d/' + f.getId() + '/view', preview: 'https://drive.google.com/file/d/' + f.getId() + '/preview' };
}

/* ---------- إرسال البيانات للمنصة ---------- */
function doGet() {
  const rows = type => { const v = sheet_(type).getDataRange().getValues(); v.shift(); return v; };
  const day = v => v instanceof Date ? Utilities.formatDate(v, 'Asia/Riyadh', 'yyyy-MM-dd') : String(v).slice(0, 10);
  const ts = v => v instanceof Date ? v.getTime() : Date.parse(v) || 0;
  const roleEn = { 'طالبة': 'student', 'معلمة': 'teacher', 'زائرة': 'visitor' };

  const stats = { views: 0, student: 0, teacher: 0, visitor: 0 }, daily = {};
  rows('view').forEach(r => { stats.views++; const k = day(r[0]); (daily[k] = daily[k] || {}).view = (daily[k].view || 0) + 1; });
  rows('login').forEach(r => { const role = roleEn[r[2]] || 'visitor'; stats[role]++; const k = day(r[0]); daily[k] = daily[k] || {}; daily[k][role] = (daily[k][role] || 0) + 1; });

  const comments = rows('comment').map(r => ({ id: r[0], ts: ts(r[1]), name: r[2], role: roleEn[r[3]] || 'visitor', text: r[4] })).reverse().slice(0, 200);
  const opinions = rows('opinion').map(r => ({ id: r[0], ts: ts(r[1]), name: r[2], child: r[3], rating: r[4], text: r[5] })).reverse().slice(0, 200);
  const results  = rows('result').map(r => ({ id: r[0], ts: ts(r[1]), name: r[2], role: roleEn[r[3]] || 'student', grade: r[4], model: r[5], title: r[6], score: r[7], total: r[8], pct: r[9] })).reverse().slice(0, 1000);
  const planRows = rows('plan');
  const plans    = planRows.length ? planRows.map(r => ({ id: r[0], term: r[1], week: r[2], unit: r[3], field: r[4], activity: r[5] })) : null;
  const videos   = rows('video').map(r => ({ id: r[0], title: r[1], url: r[2], field: r[3], ts: ts(r[4]) })).reverse();
  const worksheets = rows('worksheet').map(r => ({ id: r[0], ts: ts(r[1]), name: r[2], grade: r[3], subject: r[4], domain: r[5], score: r[6], total: r[7], explain: r[8], enrich: r[9] })).reverse().slice(0, 500);

  // عالم الفيزياء: المحتوى، النشاط، المتصدرات، المجموعات
  const content = rows('content').map(r => { let data = {}; try { data = JSON.parse(r[2]); } catch (e) {} return { id: String(r[0]), kind: r[1], data, ts: ts(r[3]) }; });
  const act = rows('activity');
  const board = {}, groups = {};
  act.forEach(r => {
    const name = String(r[2]), pts = Number(r[7]) || 0, g = String(r[4] || '');
    if (!name) return;
    const b = board[name] = board[name] || { name, grade: r[3], group: g, points: 0, count: 0 };
    b.points += pts; b.count++; if (g) b.group = g; if (r[3]) b.grade = r[3];
    if (g) { groups[g] = groups[g] || { name: g, points: 0, members: {} }; groups[g].points += pts; groups[g].members[name] = 1; }
  });
  const leaderboard = Object.keys(board).map(k => board[k]).sort((a, b) => b.points - a.points).slice(0, 30);
  const groupList = Object.keys(groups).map(k => ({ name: k, points: groups[k].points, members: Object.keys(groups[k].members).length })).sort((a, b) => b.points - a.points);
  const activity = act.slice(-400).reverse().map(r => ({ id: r[0], ts: ts(r[1]), name: r[2], grade: r[3], group: r[4], section: r[5], ref: r[6], points: r[7], details: r[8] }));

  const out = { ok: true, stats, daily, comments, opinions, results, videos, worksheets, content, leaderboard, groups: groupList, activity };
  if (plans) out.plans = plans;
  return json_(out);
}

/* ---------- البريد: سطر لكل حدث في «سجل البريد»، ورسالة مجمّعة كل ساعة ---------- */
const TITLES = { login: '🔑 دخول', comment: '💬 تعليق', opinion: '⭐ رأي ولية أمر', result: '📝 نتيجة اختبار', certificate: '🏅 شهادة', game: '🎮 اللعبة', worksheet: '📄 ورقة عمل', watch: '🎬 مشاهدة', activity: '⚛️ عالم الفيزياء' };
function logMail_(d, now) {
  let line = '';
  const e = esc_;
  switch (d.type) {
    case 'login':       line = `<b>${e(d.name)}</b> — ${roleAr_(d.role)} ${d.grade ? '· ' + e(d.grade) : ''} ${d.kind ? '· ' + e(d.kind) : ''}`; break;
    case 'comment':     line = `<b>${e(d.name)}</b>: ${e(d.text)}`; break;
    case 'opinion':     line = `<b>${e(d.name)}</b> ${'★'.repeat(Number(d.rating) || 0)}: ${e(d.text)}`; break;
    case 'result':      line = `<b>${e(d.name)}</b> ${d.grade ? '(' + e(d.grade) + ')' : ''} — ${e(d.title)}: ${d.score}/${d.total} (${d.pct}%)`; break;
    case 'certificate': line = `«${certAr_(d.certType)}» باسم <b>${e(d.name)}</b>`; break;
    case 'game':        line = `<b>${e(d.name)}</b> أنهت ${e(d.levelName)} — ${d.points} نقطة`; break;
    case 'worksheet':   line = `<b>${e(d.name)}</b> — ${e(d.subject)} · ${e(d.domain)}: ${d.score}/${d.total}<br><small>الشرح: ${e(d.explain) || '—'}</small>`; break;
    case 'watch':       line = `<b>${e(d.name)}</b> فتحت: ${e(d.video)}`; break;
    case 'activity':    line = `<b>${e(d.name)}</b> ${d.grade ? '(' + e(d.grade) + ')' : ''} — ${e(d.section)} · ${e(d.ref)}${Number(d.points) ? ' · +' + d.points + ' نقطة' : ''}${d.details ? '<br><small>' + e(d.details) + '</small>' : ''}`; break;
    default: return;
  }
  sheet_('mail').appendRow([now, d.type, line]);
}

function hourlyDigest() {
  const p = PropertiesService.getScriptProperties(), sh = sheet_('mail');
  const from = Number(p.getProperty('mailRow') || 1), last = sh.getLastRow();
  if (last <= from) return;
  const rows = sh.getRange(from + 1, 1, last - from, 3).getValues();
  const byType = {};
  rows.forEach(r => { (byType[r[1]] = byType[r[1]] || []).push(`<li><span style="color:#6b7394">${Utilities.formatDate(new Date(r[0]), 'Asia/Riyadh', 'HH:mm')}</span> — ${r[2]}</li>`); });
  const html = Object.keys(byType).map(t => `<h3 style="margin:16px 0 6px">${TITLES[t] || t} (${byType[t].length})</h3><ul style="padding-right:18px;margin:0">${byType[t].join('')}</ul>`).join('');
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: `${APP_NAME} — ${rows.length} نشاط جديد`, htmlBody: wrapMail_('ما فعلته الطالبات خلال الساعة الماضية', html) });
  p.setProperty('mailRow', String(last));
}

function dailySummary() {
  const today = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd');
  const data = JSON.parse(doGet().getContent());
  const d = data.daily[today] || {};
  const isToday = t => Utilities.formatDate(new Date(t), 'Asia/Riyadh', 'yyyy-MM-dd') === today;
  const todayRes = data.results.filter(r => isToday(r.ts));
  const todayAct = data.activity.filter(r => isToday(r.ts));
  const avg = todayRes.length ? Math.round(todayRes.reduce((a, r) => a + Number(r.pct), 0) / todayRes.length) : 0;
  const top = data.leaderboard.slice(0, 5).map((b, i) => `${i + 1}. ${esc_(b.name)} — ${b.points} نقطة`).join('<br>') || '—';
  const html = `
    <table cellpadding="8" style="border-collapse:collapse">
      <tr><td>زيارات اليوم</td><td><b>${d.view || 0}</b></td></tr>
      <tr><td>دخول الطالبات</td><td><b>${d.student || 0}</b></td></tr>
      <tr><td>دخول الزائرات</td><td><b>${d.visitor || 0}</b></td></tr>
      <tr><td>نتائج اختبارات اليوم</td><td><b>${todayRes.length}</b> (متوسط ${avg}%)</td></tr>
      <tr><td>أنشطة عالم الفيزياء اليوم</td><td><b>${todayAct.length}</b></td></tr>
      <tr><td>المتصدرات</td><td>${top}</td></tr>
    </table>`;
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: `${APP_NAME} — الملخص اليومي ${today}`, htmlBody: wrapMail_('الملخص اليومي لقياس الأثر', html) });
}

/* ---------- أدوات ---------- */
function deleteById_(type, id) {
  const sh = sheet_(type), v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id)) sh.deleteRow(i + 1);
}
function roleAr_(r) { return { student: 'طالبة', teacher: 'معلمة', visitor: 'زائرة' }[r] || 'زائرة'; }
function certAr_(t) { return { excellence: 'شهادة تميّز', complete: 'شهادة إتمام', thanks: 'شهادة شكر وتقدير', tahsili: 'شهادة إتمام خطة قادرون', course: 'شهادة إتمام مقرر' }[t] || 'شهادة'; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function wrapMail_(title, body) {
  return `<div dir="rtl" style="font-family:Tahoma,Arial;background:#07112a;padding:24px">
    <div style="max-width:600px;margin:auto;background:#fff;border-radius:16px;overflow:hidden">
      <div style="background:linear-gradient(#0d2149,#1a6a88);color:#ffd98a;padding:18px 22px;font-size:18px;font-weight:bold">🌅 ${APP_NAME}</div>
      <div style="padding:22px;color:#1d2450;font-size:15px;line-height:1.9"><h2 style="margin:0 0 10px;font-size:18px">${title}</h2>${body}</div>
      <div style="padding:12px 22px;background:#f4f6fb;color:#6b7394;font-size:12px">متوسطة وثانوية العطف · معلمة الفيزياء: أ. سماء أحمد المالكي</div>
    </div></div>`;
}
