/**
 * أفق STEM | منصة العلوم والابتكار — ربط Google Sheets + Gmail
 * --------------------------------------------------------------
 * البيانات تُحفظ في جدول «أفق STEM - بيانات المنصة» (SHEET_ID بالأسفل)
 * والإشعارات تصل إلى ADMIN_EMAIL. بعد أي تعديل: نشر ← إدارة عمليات النشر ← تعديل ← إصدار جديد.
 */

const ADMIN_EMAIL = 'asmaalmaliky1400@gmail.com';
const SHEET_ID = '1fAe_tRvWPZFolgfKFECxgZmJ0m3FpjzGzXWmq7hZHvU'; // جدول «أفق STEM - بيانات المنصة»
const APP_NAME = 'أفق STEM | منصة العلوم والابتكار';
// الأنواع التي يصل عنها بريد فوري (الزيارات العادية تُجمع في الملخص اليومي فقط)
const EMAIL_FOR = ['login', 'comment', 'opinion', 'result', 'certificate', 'game', 'worksheet'];

const SHEETS = {
  view:        ['الزيارات',       ['التاريخ', 'الوقت']],
  login:       ['الدخول',          ['التاريخ', 'الوقت', 'الصفة', 'الاسم', 'الصف', 'الفصل', 'نوع الزائرة']],
  comment:     ['التعليقات',       ['المعرف', 'التاريخ', 'الاسم', 'الصفة', 'التعليق']],
  opinion:     ['آراء أولياء الأمور', ['المعرف', 'التاريخ', 'ولية الأمر', 'الطالبة', 'التقييم', 'الرأي']],
  result:      ['نتائج النماذج',   ['المعرف', 'التاريخ', 'الطالبة', 'الصفة', 'الصف', 'النموذج', 'العنوان', 'الدرجة', 'من', 'النسبة']],
  game:        ['اللعبة',          ['التاريخ', 'الاسم', 'المستوى', 'اسم المستوى', 'النقاط', 'النجوم']],
  certificate: ['الشهادات',        ['التاريخ', 'الاسم', 'النوع', 'أصدرتها', 'الصفة']],
  watch:       ['المشاهدات',       ['التاريخ', 'الاسم', 'الصفة', 'الشرح']],
  plan:        ['الخطط',           ['المعرف', 'الفصل', 'الأسابيع', 'الوحدة', 'المجال', 'النشاط']],
  video:       ['الفيديوهات',      ['المعرف', 'العنوان', 'الرابط', 'المجال', 'التاريخ']],
  worksheet:   ['أوراق العمل',     ['المعرف', 'التاريخ', 'الطالبة', 'الصف', 'المادة', 'المجال', 'الدرجة', 'من', 'إجابة الشرح', 'إجابة الإثراء']]
};

function sheet_(type) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
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

function setup() {
  Object.keys(SHEETS).forEach(sheet_);
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'dailySummary') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(20).create();
}

/* ---------- استقبال البيانات من المنصة ---------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const now = new Date();
    const date = Utilities.formatDate(now, 'Asia/Riyadh', 'yyyy-MM-dd');
    const time = Utilities.formatDate(now, 'Asia/Riyadh', 'HH:mm');
    const s = v => String(v == null ? '' : v).slice(0, 1000);

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
      case 'plan_delete': deleteById_('plan', d.id); break;
      case 'video_delete':deleteById_('video', d.id); break;
    }
    if (EMAIL_FOR.indexOf(d.type) > -1) notify_(d);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- إرسال البيانات للمنصة (العدادات والآراء والخطط...) ---------- */
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
  const out = { ok: true, stats, daily, comments, opinions, results, videos, worksheets };
  if (plans) out.plans = plans; // إن لم تُضف المعلمة خطة بعد تبقى الخطة المقترحة
  return json_(out);
}

/* ---------- البريد ---------- */
function notify_(d) {
  const titles = { login: 'دخول جديد', comment: 'تعليق جديد', opinion: 'رأي ولية أمر', result: 'نتيجة نموذج', certificate: 'إصدار شهادة', game: 'إنجاز في اللعبة', worksheet: 'ورقة عمل مسلّمة' };
  let body = '';
  if (d.type === 'login')       body = `<b>${esc_(d.name)}</b> — ${roleAr_(d.role)} ${d.grade ? '· ' + esc_(d.grade) : ''} ${d.kind ? '· ' + esc_(d.kind) : ''}`;
  if (d.type === 'comment')     body = `<b>${esc_(d.name)}</b> (${roleAr_(d.role)}):<br>${esc_(d.text)}`;
  if (d.type === 'opinion')     body = `<b>${esc_(d.name)}</b> ${d.child ? '— ولية أمر ' + esc_(d.child) : ''}<br>التقييم: ${'★'.repeat(Number(d.rating) || 0)}<br>${esc_(d.text)}`;
  if (d.type === 'result')      body = `<b>${esc_(d.name)}</b> ${d.grade ? '(' + esc_(d.grade) + ')' : ''}<br>النموذج ${d.model}: ${esc_(d.title)}<br>الدرجة: ${d.score}/${d.total} — ${d.pct}%`;
  if (d.type === 'certificate') body = `شهادة «${certAr_(d.certType)}» باسم <b>${esc_(d.name)}</b> — أصدرتها ${esc_(d.by)}`;
  if (d.type === 'worksheet')   body = `<b>${esc_(d.name)}</b> ${d.grade ? '(' + esc_(d.grade) + ')' : ''}<br>${esc_(d.subject)} — ${esc_(d.domain)}<br>الدرجة: ${d.score}/${d.total}<br><br><b>الشرح:</b> ${esc_(d.explain) || '—'}<br><b>الإثراء:</b> ${esc_(d.enrich) || '—'}`;
  if (d.type === 'game')        body = `<b>${esc_(d.name)}</b> أنهت ${esc_(d.levelName)} — ${d.points} نقطة، ${d.stars} نجوم`;
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: `${APP_NAME} — ${titles[d.type] || d.type}`,
    htmlBody: wrapMail_(titles[d.type] || '', body)
  });
}

function dailySummary() {
  const today = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd');
  const data = JSON.parse(doGet().getContent());
  const d = data.daily[today] || {};
  const todayRes = data.results.filter(r => Utilities.formatDate(new Date(r.ts), 'Asia/Riyadh', 'yyyy-MM-dd') === today);
  const avg = todayRes.length ? Math.round(todayRes.reduce((a, r) => a + Number(r.pct), 0) / todayRes.length) : 0;
  const html = `
    <table cellpadding="8" style="border-collapse:collapse">
      <tr><td>زيارات اليوم</td><td><b>${d.view || 0}</b></td></tr>
      <tr><td>دخول الطالبات</td><td><b>${d.student || 0}</b></td></tr>
      <tr><td>دخول الزائرات</td><td><b>${d.visitor || 0}</b></td></tr>
      <tr><td>دخول المعلمات</td><td><b>${d.teacher || 0}</b></td></tr>
      <tr><td>نماذج أُنجزت اليوم</td><td><b>${todayRes.length}</b> (متوسط ${avg}%)</td></tr>
      <tr><td colspan="2" style="border-top:1px solid #ddd">الإجمالي منذ الإطلاق: ${data.stats.views} زيارة · ${data.stats.student} دخول طالبة · ${data.stats.visitor} دخول زائرة</td></tr>
    </table>`;
  MailApp.sendEmail({ to: ADMIN_EMAIL, subject: `${APP_NAME} — الملخص اليومي ${today}`, htmlBody: wrapMail_('الملخص اليومي لقياس الأثر', html) });
}

/* ---------- أدوات ---------- */
function deleteById_(type, id) {
  const sh = sheet_(type), v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]) === String(id)) sh.deleteRow(i + 1);
}
function roleAr_(r) { return { student: 'طالبة', teacher: 'معلمة', visitor: 'زائرة' }[r] || 'زائرة'; }
function certAr_(t) { return { excellence: 'شهادة تميّز', complete: 'شهادة إتمام', thanks: 'شهادة شكر وتقدير', tahsili: 'شهادة إتمام خطة قادرون' }[t] || 'شهادة'; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function wrapMail_(title, body) {
  return `<div dir="rtl" style="font-family:Tahoma,Arial;background:#07112a;padding:24px">
    <div style="max-width:560px;margin:auto;background:#fff;border-radius:16px;overflow:hidden">
      <div style="background:linear-gradient(#0d2149,#1a6a88);color:#ffd98a;padding:18px 22px;font-size:18px;font-weight:bold">🌅 ${APP_NAME}</div>
      <div style="padding:22px;color:#1d2450;font-size:15px;line-height:1.9"><h3 style="margin:0 0 10px">${title}</h3>${body}</div>
      <div style="padding:12px 22px;background:#f4f6fb;color:#6b7394;font-size:12px">متوسطة وثانوية العطف · معلمة الفيزياء: أ. سماء أحمد المالكي</div>
    </div></div>`;
}
