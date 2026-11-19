/**
 * لغتي السعادة | منصة لغتي للصف الثاني ابتدائي
 * الربط مع Google: جدول بيانات + بريد المعلمة fatimaa700fff@gmail.com
 *
 * التشغيل:
 * 1) افتحي Google Sheets جديد بحساب fatimaa700fff@gmail.com
 * 2) الإضافات ← Apps Script ، الصقي هذا الكود واحفظي
 * 3) نشر ← نشر جديد ← تطبيق ويب ← التنفيذ: أنا ← الوصول: أي شخص
 * 4) انسخي الرابط (ينتهي بـ /exec) وادخلي المنصة كمعلمة (الرمز 2026) ← لوحتي ← الصقيه واحفظي
 */
const TEACHER_EMAIL = 'fatimaa700fff@gmail.com';
const EMAIL_TYPES = ['رأي ولي أمر', 'تعليق', 'شهادة', 'نتيجة'];
const HEAD = ['التاريخ','النوع','المجموعة','الاسم','الفئة','الصف','التفاصيل','النتيجة','التقييم','at'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('السجل');
  if (!sh) { sh = ss.insertSheet('السجل'); sh.appendRow(HEAD); sh.getRange(1,1,1,HEAD.length).setFontWeight('bold').setBackground('#ffd84d'); sh.setRightToLeft(true); }
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    sheet_().appendRow([new Date(), d.type||'', d.col||'', d.name||'', d.role||'', d.grade||'', d.detail||'', d.score||'', d.rating||'', d.at||Date.now()]);
    if (EMAIL_TYPES.indexOf(d.type) > -1) {
      MailApp.sendEmail(TEACHER_EMAIL, 'لغتي السعادة | ' + d.type + ' - ' + (d.name||''),
        'منصة لغتي السعادة 📖\n\nالنوع: ' + d.type + '\nالاسم: ' + (d.name||'') + '\n' +
        (d.detail ? 'التفاصيل: ' + d.detail + '\n' : '') + (d.score ? 'النتيجة/النقاط: ' + d.score + '\n' : '') +
        (d.rating ? 'التقييم: ' + d.rating + ' / 5\n' : '') + 'الوقت: ' + (d.time || new Date()) + '\n\nالجدول: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
    }
    return json_({ok:true});
  } catch (err) { return json_({ok:false,error:String(err)}); } finally { lock.releaseLock(); }
}

// يعرض السجلات لكل الأجهزة: ?action=list&col=visits | comments | opinions
function doGet(e) {
  const col = (e && e.parameter && e.parameter.col) || 'visits';
  const sh = sheet_(); const n = sh.getLastRow(); const rows = [];
  if (n > 1) sh.getRange(2,1,n-1,HEAD.length).getValues().forEach(function(r){
    if (r[2] === col) rows.push({name:r[3], role:r[4], grade:r[5], text:r[6], score:r[7], rating:r[8], at:Number(r[9])||new Date(r[0]).getTime()});
  });
  return json_({rows: rows.slice(-500)});
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
