/**
 * لغتي السعادة | منصة لغتي للصف الثاني ابتدائي
 * ربط المنصة مع Google: جدول بيانات + بريد المعلمة
 *
 * طريقة التشغيل:
 * 1) افتحي Google Sheets جديد بحساب fatimaa700fff@gmail.com
 * 2) من القائمة: الإضافات ← Apps Script ، واحذفي الموجود والصقي هذا الكود واحفظي
 * 3) نشر ← نشر جديد ← النوع: تطبيق ويب
 *      تنفيذ كـ: أنا   |   من يمكنه الوصول: أي شخص
 * 4) وافقي على الأذونات، وانسخي رابط تطبيق الويب (ينتهي بـ /exec)
 * 5) ادخلي المنصة كمعلمة ← لوحة المعلمة ← الصقي الرابط واضغطي حفظ
 */

const TEACHER_EMAIL = 'fatimaa700fff@gmail.com';
// الأنواع التي تصل إلى البريد فوراً (الدخول يُسجَّل في الجدول فقط حتى لا يمتلئ البريد)
const EMAIL_TYPES = ['رأي ولي أمر', 'تعليق', 'نتيجة الأسئلة', 'شهادة', 'لعبة', 'اختبار'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName('السجل');
    if (!sh) {
      sh = ss.insertSheet('السجل');
      sh.appendRow(['التاريخ', 'النوع', 'الاسم', 'الفئة', 'التفاصيل', 'النتيجة/النقاط', 'التقييم', 'ولي الأمر']);
      sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#ffd84d');
      sh.setRightToLeft(true);
    }
    sh.appendRow([new Date(), d.type || '', d.name || '', d.role || '', d.detail || '',
                  d.score || '', d.rating || '', d.parent || '']);

    if (EMAIL_TYPES.indexOf(d.type) > -1) {
      const body =
        'منصة لغتي السعادة 📖\n\n' +
        'النوع: ' + (d.type || '') + '\n' +
        'الاسم: ' + (d.name || '') + ' (' + (d.role || '') + ')\n' +
        (d.detail ? 'التفاصيل: ' + d.detail + '\n' : '') +
        (d.score !== undefined && d.score !== '' ? 'النتيجة/النقاط: ' + d.score + '\n' : '') +
        (d.rating ? 'التقييم: ' + d.rating + ' / 5\n' : '') +
        (d.parent ? 'ولي الأمر: ' + d.parent + '\n' : '') +
        'الوقت: ' + (d.time || new Date()) + '\n\n' +
        'جدول البيانات: ' + ss.getUrl();
      MailApp.sendEmail(TEACHER_EMAIL, 'لغتي السعادة | ' + d.type + ' - ' + (d.name || ''), body);
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// إحصائيات تظهر تلقائياً في عدّادات المنصة لكل الأجهزة
function doGet(e) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('السجل');
  const s = { students: 0, visitors: 0, teachers: 0, games: 0, certs: 0, comments: 0, opinions: 0 };
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 2, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
      const type = r[0], role = r[2];
      if (type === 'دخول') {
        if (role === 'طالبة') s.students++;
        else if (role === 'زائرة') s.visitors++;
        else if (role === 'معلمة') s.teachers++;
      }
      if (type === 'لعبة') s.games++;
      if (type === 'شهادة') s.certs++;
      if (type === 'تعليق') s.comments++;
      if (type === 'رأي ولي أمر') s.opinions++;
    });
  }
  return json(s);
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// اختياري: ملخص يومي للدخول إلى البريد — شغّليه مرة واحدة من المحرر
function createDailySummary() {
  ScriptApp.newTrigger('dailySummary').timeBased().everyDays(1).atHour(14).create();
}
function dailySummary() {
  const s = JSON.parse(doGet().getContent());
  MailApp.sendEmail(TEACHER_EMAIL, 'لغتي السعادة | الملخص اليومي',
    'إجمالي دخول الطالبات: ' + s.students + '\nالزائرات: ' + s.visitors +
    '\nجولات اللعب: ' + s.games + '\nالشهادات: ' + s.certs +
    '\nالتعليقات: ' + s.comments + '\nآراء أولياء الأمور: ' + s.opinions);
}
