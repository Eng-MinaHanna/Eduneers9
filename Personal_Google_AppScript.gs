/**
 * Eduneers Personal Sheet Integration Script
 *
 * تعليمات التركيب:
 * 1. افتح شيت جوجل التجميعي الخاص بك.
 * 2. من القائمة العلوية اضغط Extensions (الإضافات) ثم Apps Script.
 * 3. امسح أي كود موجود، والصق هذا الكود بالكامل.
 * 4. اضغط Deploy (نشر) ثم New Deployment (نشر جديد).
 * 5. اختر النوع: Web App.
 * 6. اجعل Execute as: Me ، و Who has access: Anyone.
 * 7. انسخ الرابط الذي سيظهر في النهاية والصقه في النظام عندك في (إعداد الربط الفردي).
 */

function doGet(e) { return handleRequest(e); }
function doPost(e) {
  if (e.postData && e.postData.contents) { e.parameter = JSON.parse(e.postData.contents); }
  return handleRequest(e);
}

function handleRequest(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    var action = e.parameter.action;
    var qrCode = e.parameter.qrCode;
    if (!qrCode) throw new Error("Missing qrCode");

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ==========================================
    // تسجيل الدرجات والتقييمات (شيت التجميع)
    // ==========================================
    if (action === 'update') {
      var taskName = e.parameter.taskName; // مثال: "TASK 1"
      var category = e.parameter.category; // مثال: "Main task", "attendance .15"
      var val = e.parameter.val;

      // ابحث عن ورقة "Grade Sheet" أو استخدم أول ورقة كبديل
      var sheet = ss.getSheetByName('Grade Sheet') || ss.getSheetByName('شيت التقييم العام') || ss.getSheets()[0];

      // 1. تحديد عمود التاسك (بالبحث في الصف الثاني من عمود C إلى عمود AF)
      var headerRow = sheet.getRange(2, 1, 1, 40).getValues()[0];
      var taskColIndex = -1;
      for (var i = 0; i < headerRow.length; i++) {
        if (String(headerRow[i]).trim() === String(taskName).trim()) {
          taskColIndex = i + 1;
          break;
        }
      }
      if (taskColIndex === -1) {
        return output.setContent(JSON.stringify({status: "error", message: "عمود التقييم غير موجود: " + taskName}));
      }

      // 2. البحث عن الطالب عبر الكود في العمود B
      var lastRow = sheet.getLastRow();
      if (lastRow < 3) lastRow = 100;
      var codeValues = sheet.getRange(1, 2, lastRow, 1).getValues(); // Column B
      var studentBaseRow = -1;
      for (var i = 0; i < codeValues.length; i++) {
        if (String(codeValues[i][0]).trim() === String(qrCode).trim()) {
          studentBaseRow = i + 1;
          break;
        }
      }
      if (studentBaseRow === -1) {
         return output.setContent(JSON.stringify({status: "error", message: "كود الطالب غير موجود"}));
      }

      // 3. تحديد صف التقييم في البلوك الخاص بالطالب بناء على التسميات في العمود F (من F3 لأسفل 7 صفوف)
      var catValues = sheet.getRange(studentBaseRow, 6, 7, 1).getValues(); // Column F
      var targetRow = -1;
      for (var i = 0; i < 7; i++) {
        var cellVal = String(catValues[i][0]).toLowerCase();
        var searchCat = String(category).toLowerCase();
        if (cellVal.indexOf(searchCat) !== -1 || searchCat.indexOf(searchCat) !== -1) {
           targetRow = studentBaseRow + i;
           break;
        }
      }

      // في حال وجود خطأ مطبعي في المسميات يتم الاعتماد على الترتيب الثابت
      if (targetRow === -1) {
         var catStr = category.toLowerCase();
         if (catStr.indexOf('main') !== -1) targetRow = studentBaseRow;
         else if (catStr.indexOf('attendance') !== -1) targetRow = studentBaseRow + 1;
         else if (catStr.indexOf('feedback') !== -1) targetRow = studentBaseRow + 2;
         else if (catStr.indexOf('attitude') !== -1) targetRow = studentBaseRow + 3;
         else if (catStr.indexOf('quiz') !== -1) targetRow = studentBaseRow + 4;
         else if (catStr.indexOf('bonus') !== -1) targetRow = studentBaseRow + 5;
         else targetRow = studentBaseRow + 6;
      }

      // 4. وضع الدرجة في المكان المناسب
      sheet.getRange(targetRow, taskColIndex).setValue(val);
      return output.setContent(JSON.stringify({status: "success", targetRow: targetRow, col: taskColIndex}));
    }

    // ==========================================
    // تسجيل الحضور (شيت الحضور بنسبة 80%)
    // ==========================================
    else if (action === 'attend') {
      var lectureNum = parseInt(e.parameter.lectureNum);

      // ابحث عن ورقة الحضور
      var sheet = ss.getSheetByName('Attendance Sheet') || ss.getSheetByName('شيت الحضور') || ss.getSheets()[1] || ss.getSheets()[0];

      // أكواد الطلاب في العمود B والصفوف تبدأ من 5 إلى 60 (النطاق C5:S60)
      var codeValues = sheet.getRange(1, 2, 65, 1).getValues(); // فحص حتى 65
      var studentRow = -1;
      for (var i = 3; i < 65; i++) { // من الصف الرابع عشان نغطي من C5 واسماء الطلبة
         if (String(codeValues[i][0]).trim() === String(qrCode).trim()) {
           studentRow = i + 1;
           break;
         }
      }
      if (studentRow === -1) {
         return output.setContent(JSON.stringify({status: "error", message: "كود الطالب غير موجود في شيت الحضور"}));
      }

      // عمود المحاضرة (محاضرة 1 تبدأ في العمود C وهو العمود رقم 3)
      var targetCol = 2 + lectureNum; 

      // رصد الحضور إما بعلامة معينة أو الرقم 1
      sheet.getRange(studentRow, targetCol).setValue("1"); 

      return output.setContent(JSON.stringify({status: "success", targetRow: studentRow, col: targetCol}));
    }

    else {
      return output.setContent(JSON.stringify({status: "error", message: "Unknown action"}));
    }

  } catch (err) {
    return output.setContent(JSON.stringify({status: "error", message: err.toString()}));
  }
}
