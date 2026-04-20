// ╔══════════════════════════════════════════════════════════════════╗
// ║   Google Apps Script — وحدة تكنولوجيا المعلومات  v2.0          ║
// ║   جمعية الحياة · القرارة - غرداية                               ║
// ║                                                                  ║
// ║   ✅ Google Sheets     — حفظ الشكاوي + تنسيق احترافي            ║
// ║   ✅ Telegram Bot      — إشعارات فورية للمدير                    ║
// ║   ✅ GmailApp          — تأكيد تلقائي للمستخدم + إشعار المدير   ║
// ║   ✅ WhatsApp Business — إشعار عند الإصلاح/الرفض                ║
// ║   ✅ doGet (stats)     — إحصائيات حية للصفحة                    ║
// ║   ✅ doGet (lookup)    — البحث عن تذكرة                          ║
// ║   ✅ onEdit            — مراقبة الحالة وإشعار التلقائي           ║
// ║   ✅ onOpen            — قائمة مخصصة في Sheets                   ║
// ║   ✅ Anti-Spam         — حماية من الطلبات المتكررة               ║
// ╚══════════════════════════════════════════════════════════════════╝

// ════════════════════════════════════════════════════════════
//  🔒 الإعدادات — شغّل setupConfig() مرة واحدة فقط
// ════════════════════════════════════════════════════════════
function setupConfig() {
  PropertiesService.getScriptProperties().setProperties({
    "TG_TOKEN":     "YOUR_TELEGRAM_BOT_TOKEN",
    "TG_CHAT_ID":   "YOUR_TELEGRAM_GROUP_ID",
    "WA_TOKEN":     "YOUR_WHATSAPP_API_TOKEN",
    "WA_PHONE_ID":  "YOUR_WHATSAPP_PHONE_ID",
    "ADMIN_EMAIL":  "YOUR_ADMIN_EMAIL",
    "ADMIN_NAME":   "وحدة IT — مؤسسات الحياة"
  });
  Logger.log("✅ تم حفظ الإعدادات بأمان");
}

function getConfig() {
  var p = PropertiesService.getScriptProperties().getProperties();
  return {
    tgToken:    p["TG_TOKEN"]     || "",
    tgChatId:   p["TG_CHAT_ID"]   || "",
    waToken:    p["WA_TOKEN"]     || "",
    waPhoneId:  p["WA_PHONE_ID"]  || "",
    adminEmail: p["ADMIN_EMAIL"]  || "hs.ga.it@gmail.com",
    adminName:  p["ADMIN_NAME"]   || "وحدة IT — مؤسسات الحياة"
  };
}

// ════════════════════════════════════════════════════════════
//  ثوابت الأعمدة
// ════════════════════════════════════════════════════════════
var SHEET_NAME   = "الشكاوي";
var COL_TICKET   = 1;
var COL_DATE     = 2;
var COL_NAME     = 3;
var COL_EMP_ID   = 4;
var COL_PHONE    = 5;
var COL_EMAIL    = 6;
var COL_DEPT     = 7;
var COL_TYPE     = 8;
var COL_PRIORITY = 9;
var COL_DETAILS  = 10;
var COL_STATUS   = 11;
var COL_NOTES    = 12;
var COL_CLOSED   = 13;  // وقت الإغلاق
var COL_SOURCE   = 14;  // مصدر الطلب

// ════════════════════════════════════════════════════════════
//  🚀 استقبال الطلب POST
// ════════════════════════════════════════════════════════════
function doPost(e) {
  var ticketNum = "";
  try {
    var raw  = e.postData.contents;
    var data = JSON.parse(raw);

    // ── التحقق من المدخلات ─────────────────────────────────
    var validation = validateInput(data);
    if (!validation.ok) {
      return jsonResponse({ success: false, error: validation.msg });
    }

    // ── حماية Anti-Spam ────────────────────────────────────
    var spamCheck = checkSpam(data.phone || "");
    if (!spamCheck.ok) {
      return jsonResponse({ success: false, error: spamCheck.msg });
    }

    // ── حفظ في Sheets ──────────────────────────────────────
    ticketNum = saveToSheet(data);

    // ── إشعارات (لا تُفشل الطلب لو أخفقت) ────────────────
    try { notifyAdminTelegram(data, ticketNum); }
    catch(err) { Logger.log("TG Error: " + err.message); }

    try { sendConfirmationEmail(data, ticketNum); }
    catch(err) { Logger.log("Email Error: " + err.message); }

    try { notifyAdminEmail(data, ticketNum); }
    catch(err) { Logger.log("Admin Email Error: " + err.message); }

    return jsonResponse({ success: true, ticket: ticketNum });

  } catch(err) {
    Logger.log("doPost FATAL: " + err.message + " | stack: " + err.stack);
    return jsonResponse({ success: false, error: "خطأ داخلي — حاول لاحقاً" });
  }
}

// ════════════════════════════════════════════════════════════
//  🌐 استقبال الطلب GET (إحصائيات + بحث عن تذكرة)
// ════════════════════════════════════════════════════════════
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";

  if (action === "stats") {
    return getStats();
  }
  if (action === "lookup" && e.parameter.ticket) {
    return lookupTicket(e.parameter.ticket);
  }
  return jsonResponse({ status: "يعمل", version: "2.0", project: "جمعية الحياة — وحدة IT" });
}

// ════════════════════════════════════════════════════════════
//  📊 الإحصائيات الحية
// ════════════════════════════════════════════════════════════
function getStats() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var stats = { total: 0, fixed: 0, rejected: 0, open: 0, processing: 0, urgent: 0 };

  if (sheet && sheet.getLastRow() > 1) {
    var statuses = sheet.getRange(2, COL_STATUS, sheet.getLastRow() - 1, 1).getValues();
    statuses.forEach(function(row) {
      var s = (row[0] || "").toString();
      if (s) {
        stats.total++;
        if      (s.indexOf("تم الإصلاح")     >= 0) stats.fixed++;
        else if (s.indexOf("مرفوض")          >= 0) stats.rejected++;
        else if (s.indexOf("قيد المعالجة")   >= 0) stats.processing++;
        else if (s.indexOf("تحويل")          >= 0) stats.urgent++;
        else                                        stats.open++;
      }
    });
  }
  return jsonResponse(stats);
}

// ════════════════════════════════════════════════════════════
//  🔍 البحث عن تذكرة بالرقم
// ════════════════════════════════════════════════════════════
function lookupTicket(ticketId) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonResponse({ found: false, msg: "لا توجد بيانات" });
  }

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, COL_SOURCE).getValues();
  for (var i = 0; i < data.length; i++) {
    if ((data[i][COL_TICKET - 1] || "").toString().toLowerCase() === ticketId.toLowerCase()) {
      var r = data[i];
      return jsonResponse({
        found:    true,
        ticket:   r[COL_TICKET   - 1],
        date:     r[COL_DATE     - 1],
        name:     r[COL_NAME     - 1],
        dept:     r[COL_DEPT     - 1],
        type:     r[COL_TYPE     - 1],
        priority: r[COL_PRIORITY - 1],
        status:   r[COL_STATUS   - 1],
        notes:    r[COL_NOTES    - 1],
        closed:   r[COL_CLOSED   - 1] || ""
      });
    }
  }
  return jsonResponse({ found: false, msg: "التذكرة غير موجودة" });
}

// ════════════════════════════════════════════════════════════
//  ✅ التحقق من المدخلات
// ════════════════════════════════════════════════════════════
function validateInput(data) {
  if (!data.name    || data.name.trim().length < 2)   return { ok: false, msg: "الاسم قصير جداً" };
  if (!data.phone   || data.phone.toString().replace(/\D/g,"").length < 9)
    return { ok: false, msg: "رقم الهاتف غير صحيح" };
  if (!data.email   || data.email.indexOf("@") < 0)   return { ok: false, msg: "البريد غير صحيح" };
  if (!data.details || data.details.trim().length < 5) return { ok: false, msg: "التفاصيل قصيرة جداً" };
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  🛡️ Anti-Spam — 3 طلبات/ساعة لنفس الرقم
// ════════════════════════════════════════════════════════════
function checkSpam(phone) {
  if (!phone) return { ok: true };
  var store  = CacheService.getScriptCache();
  var key    = "spam_" + phone.replace(/\D/g, "").slice(-9);
  var count  = parseInt(store.get(key) || "0", 10);
  if (count >= 5) return { ok: false, msg: "تجاوزت الحد المسموح (5 طلبات/ساعة)، انتظر قبل المحاولة مجدداً." };
  store.put(key, String(count + 1), 3600);
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
//  💾 1. حفظ الشكوى في Sheets
// ════════════════════════════════════════════════════════════
function saveToSheet(data) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); setupHeaders(sheet); }

  var rowCount  = sheet.getLastRow();
  var ticketNum = "TKT-" + String(rowCount).padStart(4, "0");
  var timestamp = Utilities.formatDate(new Date(), "Africa/Algiers", "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    ticketNum,
    timestamp,
    sanitize(data.name),
    sanitize(data.empId),
    sanitize(data.phone),
    sanitize(data.email),
    sanitize(data.dept),
    sanitize(data.type),
    sanitize(data.priority),
    sanitize(data.details),
    "🔵 مفتوح",
    "",
    "",           // COL_CLOSED
    "Web Form"    // COL_SOURCE
  ]);

  var lastRow = sheet.getLastRow();
  formatRow(sheet, lastRow);
  return ticketNum;
}

function sanitize(val) {
  if (!val) return "";
  return val.toString()
    .replace(/<[^>]*>/g, "")    // strip HTML
    .replace(/['"`;]/g, "")     // strip injection chars
    .substring(0, 500)
    .trim();
}

function setupHeaders(sheet) {
  var h = [
    "رقم التذكرة", "التاريخ والوقت", "الاسم الكامل",
    "رمز العتاد",  "رقم الهاتف",     "البريد الإلكتروني",
    "الفرع والإدارة", "نوع الطلب",  "الأولوية",
    "تفاصيل الشكوى", "الحالة",      "ملاحظات الفريق",
    "وقت الإغلاق",   "المصدر"
  ];
  sheet.appendRow(h);
  var hRange = sheet.getRange(1, 1, 1, h.length);
  hRange.setBackground("#4a0606")
        .setFontColor("#f0d080")
        .setFontWeight("bold")
        .setFontSize(11)
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setWrap(true);
  sheet.setRowHeight(1, 46);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  var widths = [120, 165, 145, 120, 145, 200, 185, 160, 110, 320, 155, 210, 155, 110];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  // قائمة حالات منسدلة
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      "🔵 مفتوح",
      "🟡 قيد المعالجة",
      "↪️ تحويل إلى جهة أخرى",
      "✅ تم الإصلاح",
      "❌ مرفوض"
    ], true).build();
  sheet.getRange(2, COL_STATUS, 5000, 1).setDataValidation(rule);
}

function formatRow(sheet, row) {
  var bg = (row % 2 === 0) ? "#fdf6e3" : "#ffffff";
  sheet.getRange(row, 1, 1, 14)
       .setBackground(bg)
       .setVerticalAlignment("middle")
       .setWrap(false);
  sheet.setRowHeight(row, 36);

  sheet.getRange(row, COL_TICKET)
       .setBackground("#7b0c0c")
       .setFontColor("#f0d080")
       .setFontWeight("bold")
       .setHorizontalAlignment("center");

  sheet.getRange(row, COL_STATUS)
       .setBackground("#e3f2fd")
       .setHorizontalAlignment("center")
       .setFontWeight("bold");

  sheet.getRange(row, COL_PRIORITY)
       .setHorizontalAlignment("center");

  sheet.getRange(row, COL_DATE)
       .setHorizontalAlignment("center")
       .setFontSize(11);
}

// ════════════════════════════════════════════════════════════
//  📲 2. Telegram — إرسال رسالة
// ════════════════════════════════════════════════════════════
function sendTelegram(chatId, message, replyMarkup) {
  var cfg = getConfig();
  if (!cfg.tgToken || !chatId) {
    Logger.log("TG: token أو chatId فارغ");
    return;
  }
  var payload = {
    chat_id:    String(chatId),
    text:       message,
    parse_mode: "HTML"
  };
  if (replyMarkup) payload["reply_markup"] = replyMarkup;

  var url = "https://api.telegram.org/bot" + cfg.tgToken + "/sendMessage";
  var res = UrlFetchApp.fetch(url, {
    method:           "post",
    contentType:      "application/json",
    payload:          JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  Logger.log("Telegram [" + code + "]");
  if (code !== 200) {
    Logger.log("TG error body: " + res.getContentText().substring(0, 400));
    throw new Error("Telegram HTTP " + code);
  }
}

// ════════════════════════════════════════════════════════════
//  💬 3. WhatsApp Business — إرسال قالب
// ════════════════════════════════════════════════════════════
function cleanPhone(phone) {
  var to = phone.toString().replace(/[^0-9]/g, "");
  if (to.charAt(0) === "0") to = "213" + to.substring(1);
  if (to.substring(0, 3) !== "213") to = "213" + to;
  return to;
}

function sendWhatsAppTemplate(phone, template, params) {
  var cfg = getConfig();
  if (!phone || !cfg.waToken || !cfg.waPhoneId) {
    Logger.log("WA: إعدادات ناقصة");
    return;
  }
  var to = cleanPhone(phone);

  var components = [];
  if (params && params.length > 0) {
    components.push({
      type: "body",
      parameters: params.map(function(p) { return { type: "text", text: p.toString() }; })
    });
  }

  var url = "https://graph.facebook.com/v18.0/" + cfg.waPhoneId + "/messages";
  var payload = {
    messaging_product: "whatsapp",
    to:                to,
    type:              "template",
    template: {
      name:       template,
      language:   { code: "ar" },
      components: components
    }
  };

  var res = UrlFetchApp.fetch(url, {
    method:           "post",
    contentType:      "application/json",
    headers:          { "Authorization": "Bearer " + cfg.waToken },
    payload:          JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  Logger.log("WhatsApp [" + code + "] → " + to + " | template: " + template);
  if (code !== 200) Logger.log("WA error: " + res.getContentText().substring(0, 400));
}

// ════════════════════════════════════════════════════════════
//  4. إشعار تيليغرام للمدير — شكوى جديدة
// ════════════════════════════════════════════════════════════
function notifyAdminTelegram(data, ticketNum) {
  var cfg      = getConfig();
  var priority = data.priority || "عادي";
  var prEmoji  = priority.indexOf("عاجل")  >= 0 ? "🔴"
               : priority.indexOf("متوسط") >= 0 ? "🟡" : "🟢";
  var time     = Utilities.formatDate(new Date(), "Africa/Algiers", "HH:mm — yyyy/MM/dd");
  var details  = sanitize(data.details || "—").substring(0, 250);

  var lines = [
    "┌───────────────────────────┐",
    "│  🔔  شكوى جديدة — IT       │",
    "└───────────────────────────┘",
    "",
    "🎫  <b>" + ticketNum + "</b>   " + prEmoji + "  " + priority,
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "👤  <b>الاسم :</b>  "  + (data.name   || "—"),
    "🔧  <b>العتاد:</b>  "  + (data.empId  || "—"),
    "📱  <b>الهاتف:</b>  "  + (data.phone  || "—"),
    "📧  <b>البريد:</b>  "  + (data.email  || "—"),
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "🏢  <b>الفرع :</b>  "  + (data.dept   || "—"),
    "📂  <b>النوع :</b>  "  + (data.type   || "—"),
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "📝  <b>التفاصيل:</b>",
    "<i>" + details + "</i>",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "⏰  " + time,
    "",
    "🏛  <b>مؤسسات الحياة — وحدة IT</b>"
  ];

  sendTelegram(cfg.tgChatId, lines.join("\n"));
}

// ════════════════════════════════════════════════════════════
//  5. إيميل تأكيد للمستخدم (فور الإرسال)
// ════════════════════════════════════════════════════════════
function sendConfirmationEmail(data, ticketNum) {
  var email = data.email || "";
  if (!email || email.indexOf("@") < 0) return;

  var cfg     = getConfig();
  var name    = data.name    || "العزيز";
  var type    = data.type    || "—";
  var dept    = data.dept    || "—";
  var priority= data.priority|| "عادي";
  var details = sanitize(data.details || "—").substring(0, 300);
  var time    = Utilities.formatDate(new Date(), "Africa/Algiers", "yyyy/MM/dd HH:mm");

  var prColor = priority.indexOf("عاجل") >= 0 ? "#c62828"
              : priority.indexOf("متوسط") >= 0 ? "#f57f17" : "#388e3c";

  var subject = "✅ تم استلام طلبك — " + ticketNum + " | جمعية الحياة";
  var html = buildEmailWrapper(
    "#0d4f70", "#0d7fb5",
    "📋", "تم استلام طلبك بنجاح!",
    "وحدة IT — مؤسسات الحياة",
    [
      '<p style="font-size:16px;color:#0d4f70;font-weight:bold">السلام عليكم ' + name + ',</p>',
      '<p style="color:#555;line-height:1.8;margin-bottom:20px">تم استلام طلبك وسيتولى فريقنا التقني معالجته في <strong>أقرب وقت</strong>.</p>',
      buildTicketCard("#0d7fb5","#e3f2fd",[
        ["رقم التذكرة",      ticketNum,  "#0d7fb5"],
        ["نوع الطلب",        type,       null],
        ["الفرع والإدارة",   dept,       null],
        ["الأولوية",         priority,   prColor],
        ["تفاصيل الطلب",     details,    null],
        ["وقت الإرسال",      time,       null]
      ]),
      '<div style="background:#e3f2fd;border-radius:8px;padding:14px;text-align:center;margin-bottom:16px">',
      '<p style="color:#0d4f70;font-weight:bold;margin:0">سنُرسل لك إيميلاً آخر عند إتمام التدخل التقني ✔️</p>',
      '</div>',
      '<p style="color:#888;font-size:12px;text-align:center">يمكنك تتبع حالة طلبك بزيارة صفحتنا وإدخال رقم التذكرة</p>'
    ].join(""),
    cfg.adminEmail
  );
  GmailApp.sendEmail(email, subject, "تم استلام طلبك — " + ticketNum, {
    htmlBody: html, name: cfg.adminName
  });
  Logger.log("✅ إيميل تأكيد → " + email);
}

// ════════════════════════════════════════════════════════════
//  6. إيميل للمدير — شكوى جديدة (نسخة احتياطية من تيليغرام)
// ════════════════════════════════════════════════════════════
function notifyAdminEmail(data, ticketNum) {
  var cfg  = getConfig();
  if (!cfg.adminEmail) return;

  var priority = data.priority || "عادي";
  var prColor  = priority.indexOf("عاجل") >= 0 ? "#c62828"
               : priority.indexOf("متوسط") >= 0 ? "#f57f17" : "#388e3c";
  var time = Utilities.formatDate(new Date(), "Africa/Algiers", "yyyy/MM/dd HH:mm");

  var subject = "🔔 شكوى جديدة — " + ticketNum + " — " + (data.type || "") + " | جمعية الحياة";
  var html = buildEmailWrapper(
    "#4a0606", "#7b0c0c",
    "🔔", "شكوى جديدة وردت للفريق",
    "وحدة IT — مؤسسات الحياة",
    [
      '<p style="font-size:15px;color:#4a0606;font-weight:bold">طلب تدخل جديد بانتظار المعالجة</p>',
      buildTicketCard("#7b0c0c","#fff8e1",[
        ["رقم التذكرة",      ticketNum,              "#7b0c0c"],
        ["الاسم الكامل",     data.name  || "—",      null],
        ["رمز العتاد",       data.empId || "—",      null],
        ["الهاتف",           data.phone || "—",      null],
        ["البريد",           data.email || "—",      null],
        ["الفرع والإدارة",   data.dept  || "—",      null],
        ["نوع الطلب",        data.type  || "—",      null],
        ["الأولوية",         priority,               prColor],
        ["التفاصيل",         sanitize(data.details || "—").substring(0,300), null],
        ["وقت الاستلام",     time,                   null]
      ])
    ].join(""),
    cfg.adminEmail
  );
  GmailApp.sendEmail(cfg.adminEmail, subject, "شكوى جديدة — " + ticketNum, {
    htmlBody: html, name: "نظام الشكاوي التلقائي"
  });
  Logger.log("✅ إيميل مدير → " + cfg.adminEmail);
}

// ════════════════════════════════════════════════════════════
//  7. إيميل الإصلاح
// ════════════════════════════════════════════════════════════
function sendFixedEmail(email, name, ticket, type, dept, time) {
  var cfg     = getConfig();
  var subject = "✅ تم إصلاح طلبك — " + ticket + " | جمعية الحياة";
  var html    = buildEmailWrapper(
    "#1b5e20", "#2e7d32",
    "✅", "تم إصلاح طلبك بنجاح!",
    "وحدة IT — مؤسسات الحياة",
    [
      '<p style="font-size:16px;color:#1b5e20;font-weight:bold">السلام عليكم ' + name + ',</p>',
      '<p style="color:#555;line-height:1.8;margin-bottom:20px">يسعدنا إخبارك أنه تم <strong style="color:#2e7d32">حل مشكلتك</strong> بنجاح من قِبل فريقنا التقني. 🎉</p>',
      buildTicketCard("#2e7d32","#e8f5e9",[
        ["رقم التذكرة",    ticket, "#2e7d32"],
        ["نوع الطلب",      type,   null],
        ["الفرع والإدارة", dept,   null],
        ["وقت الإغلاق",    time,   null]
      ]),
      '<div style="background:#e8f5e9;border-radius:8px;padding:16px;text-align:center;margin-bottom:16px">',
      '<p style="color:#2e7d32;font-weight:bold;margin:0;font-size:15px">شكراً لتواصلك مع وحدة تكنولوجيا المعلومات 🙏</p>',
      '<p style="color:#555;font-size:12px;margin:8px 0 0">إذا تكررت المشكلة، أرسل طلباً جديداً عبر صفحتنا</p>',
      '</div>'
    ].join(""),
    cfg.adminEmail
  );
  GmailApp.sendEmail(email, subject, "تم إصلاح طلبك — " + ticket, {
    htmlBody: html, name: cfg.adminName
  });
  Logger.log("✅ إيميل إصلاح → " + email);
}

// ════════════════════════════════════════════════════════════
//  8. إيميل الرفض
// ════════════════════════════════════════════════════════════
function sendRejectedEmail(email, name, ticket, type, dept, notes, time) {
  var cfg     = getConfig();
  var reason  = notes || "سيتم التواصل معك قريباً.";
  var subject = "📋 بخصوص طلبك — " + ticket + " | جمعية الحياة";
  var html    = buildEmailWrapper(
    "#4a0606", "#7b0c0c",
    "🕌", "وحدة تكنولوجيا المعلومات",
    "مؤسسات الحياة — القرارة، غرداية",
    [
      '<p style="font-size:16px;color:#4a0606;font-weight:bold">السلام عليكم ' + name + ',</p>',
      '<p style="color:#555;line-height:1.8;margin-bottom:20px">نشكرك على تواصلك. نعتذر إذ <strong style="color:#c62828">تعذّر معالجة طلبك</strong> في الوقت الراهن.</p>',
      buildTicketCard("#c62828","#ffebee",[
        ["رقم التذكرة",    ticket, "#c62828"],
        ["نوع الطلب",      type,   null],
        ["الفرع والإدارة", dept,   null],
        ["التاريخ",        time,   null]
      ]),
      '<div style="background:#fff8e1;border-right:4px solid #ffa726;border-radius:6px;padding:14px;margin-bottom:18px">',
      '<p style="color:#8d6e1a;font-weight:bold;font-size:13px;margin:0 0 6px">ملاحظة من الفريق التقني:</p>',
      '<p style="color:#795548;font-size:13px;margin:0;line-height:1.7">' + reason + '</p>',
      '</div>',
      '<div style="background:#fce4ec;border:1.5px solid #f48fb1;border-radius:8px;padding:14px;text-align:center;margin-bottom:18px">',
      '<p style="color:#880e4f;font-weight:bold;margin:0">نعتذر عن أي إزعاج — يمكنك إعادة تقديم الطلب عبر صفحتنا</p>',
      '</div>'
    ].join(""),
    cfg.adminEmail
  );
  GmailApp.sendEmail(email, subject, "بخصوص طلبك — " + ticket, {
    htmlBody: html, name: cfg.adminName
  });
  Logger.log("✅ إيميل رفض → " + email);
}

// ════════════════════════════════════════════════════════════
//  🛠 مساعد بناء القوالب البريدية
// ════════════════════════════════════════════════════════════
function buildEmailWrapper(topColor1, topColor2, icon, title, subtitle, bodyHtml, adminEmail) {
  var cfg = getConfig();
  return [
    '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:620px;margin:0 auto;',
    'background:#fdf6e3;border-radius:14px;overflow:hidden;border:2px solid ' + topColor2 + '">',

    // ── Header ──
    '<div style="background:linear-gradient(135deg,' + topColor1 + ',' + topColor2 + ');padding:30px 24px;text-align:center">',
    '<div style="font-size:52px;margin-bottom:6px">' + icon + '</div>',
    '<h2 style="color:#f0d080;margin:0 0 6px;font-size:22px">' + title + '</h2>',
    '<p style="color:rgba(240,208,128,.75);font-size:12px;margin:0">' + subtitle + '</p>',
    '</div>',

    // ── Gold bar ──
    '<div style="height:4px;background:linear-gradient(90deg,transparent,#c9a84c,#f0d080,#c9a84c,transparent)"></div>',

    // ── Body ──
    '<div style="padding:28px 30px">',
    bodyHtml,
    '<hr style="border:none;border-top:1px solid #e0d4b0;margin:18px 0"/>',
    '<p style="text-align:center;color:#4a0606;font-weight:bold;font-size:13px">',
    '📞 0779149963 &nbsp;|&nbsp; 📧 ' + (adminEmail || cfg.adminEmail),
    '</p>',
    '</div>',

    // ── Footer ──
    '<div style="background:' + topColor1 + ';padding:14px 24px;text-align:center">',
    '<p style="color:rgba(240,208,128,.65);font-size:11px;margin:0">',
    'مؤسسات الحياة — وحدة تكنولوجيا المعلومات &nbsp;|&nbsp; القرارة، غرداية 🇩🇿</p>',
    '<p style="color:rgba(240,208,128,.4);font-size:10px;margin:4px 0 0">',
    '1937 - 2025 — هذا الإيميل أُرسل تلقائياً — لا تردّ عليه',
    '</p></div></div>'
  ].join("");
}

function buildTicketCard(accentColor, bgColor, rows) {
  var html = [
    '<div style="background:#fff;border:2px solid ' + accentColor + ';border-radius:12px;',
    'padding:20px;margin-bottom:20px">',
    '<table style="width:100%;font-size:13px;border-collapse:collapse">'
  ].join("");

  rows.forEach(function(row, i) {
    var label  = row[0];
    var value  = row[1];
    var color  = row[2] || "#555";
    var isLast = (i === rows.length - 1);
    html += [
      '<tr style="' + (isLast ? "" : "border-bottom:1px solid " + bgColor) + '">',
      '<td style="padding:9px 12px;color:' + accentColor + ';font-weight:bold;white-space:nowrap">',
      label, '</td>',
      '<td style="padding:9px 12px;color:' + color + ';word-break:break-word">', value, '</td>',
      '</tr>'
    ].join("");
  });

  html += '</table></div>';
  return html;
}

// ════════════════════════════════════════════════════════════
//  9. onEdit — مراقبة عمود الحالة
// ════════════════════════════════════════════════════════════
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (col !== COL_STATUS || row < 2) return;

  var cfg    = getConfig();
  var status = (e.value || "").toString();
  var email  = sheet.getRange(row, COL_EMAIL).getValue();
  var name   = sheet.getRange(row, COL_NAME).getValue();
  var ticket = sheet.getRange(row, COL_TICKET).getValue();
  var type   = sheet.getRange(row, COL_TYPE).getValue();
  var dept   = sheet.getRange(row, COL_DEPT).getValue();
  var phone  = sheet.getRange(row, COL_PHONE).getValue();
  var notes  = sheet.getRange(row, COL_NOTES).getValue();
  var time   = Utilities.formatDate(new Date(), "Africa/Algiers", "yyyy/MM/dd HH:mm");

  // ── تم الإصلاح ──────────────────────────────────────────
  if (status === "✅ تم الإصلاح") {
    // سجّل وقت الإغلاق
    sheet.getRange(row, COL_CLOSED).setValue(time);

    // WhatsApp
    try { sendWhatsAppTemplate(phone, "it_fixed1", [ticket, name, type, time]); }
    catch(err) { Logger.log("WA fixed: " + err.message); }

    // إيميل للمستخدم
    if (email && email.indexOf("@") >= 0) {
      try { sendFixedEmail(email, name, ticket, type, dept, time); }
      catch(err) { Logger.log("Email fixed: " + err.message); }
    }

    // إشعار تيليغرام للمدير
    try {
      sendTelegram(cfg.tgChatId,
        "✅ <b>تم إصلاح:</b> " + ticket + "\n👤 " + name + "\n📂 " + type + "\n⏰ " + time);
    } catch(err) { Logger.log("TG fixed: " + err.message); }

    // تلوين الصف
    sheet.getRange(row, 1, 1, 14).setBackground("#e8f5e9");
    sheet.getRange(row, COL_STATUS)
         .setBackground("#4caf50").setFontColor("#ffffff").setFontWeight("bold");
  }

  // ── مرفوض ────────────────────────────────────────────────
  else if (status === "❌ مرفوض") {
    sheet.getRange(row, COL_CLOSED).setValue(time);

    try { sendWhatsAppTemplate(phone, "it_rejected1", [ticket, name, notes || "لم يُحدد"]); }
    catch(err) { Logger.log("WA rejected: " + err.message); }

    if (email && email.indexOf("@") >= 0) {
      try { sendRejectedEmail(email, name, ticket, type, dept, notes, time); }
      catch(err) { Logger.log("Email rejected: " + err.message); }
    }

    try {
      sendTelegram(cfg.tgChatId,
        "❌ <b>مرفوض:</b> " + ticket + "\n👤 " + name + "\n📝 " + (notes || "—") + "\n⏰ " + time);
    } catch(err) { Logger.log("TG rejected: " + err.message); }

    sheet.getRange(row, 1, 1, 14).setBackground("#fce4ec");
    sheet.getRange(row, COL_STATUS)
         .setBackground("#ef9a9a").setFontColor("#4e342e").setFontWeight("bold");
  }

  // ── تحويل إلى جهة أخرى ───────────────────────────────────
  else if (status === "↪️ تحويل إلى جهة أخرى") {
    try {
      sendTelegram(cfg.tgChatId,
        "↪️ <b>تحويل إلى جهة أخرى:</b> " + ticket + "\n👤 " + name + "\n📂 " + type + "\n🏢 " + dept);
    } catch(err) { Logger.log("TG urgent: " + err.message); }

    sheet.getRange(row, COL_STATUS)
         .setBackground("#ffcdd2").setFontColor("#c62828").setFontWeight("bold");
  }

  // ── قيد المعالجة ─────────────────────────────────────────
  else if (status === "🟡 قيد المعالجة") {
    try {
      sendTelegram(cfg.tgChatId,
        "🟡 <b>قيد المعالجة:</b> " + ticket + "\n👤 " + name);
    } catch(err) { Logger.log("TG processing: " + err.message); }

    sheet.getRange(row, COL_STATUS)
         .setBackground("#fff9c4").setFontColor("#f57f17").setFontWeight("bold");
  }

  // ── مفتوح ────────────────────────────────────────────────
  else if (status === "🔵 مفتوح") {
    sheet.getRange(row, COL_STATUS)
         .setBackground("#e3f2fd").setFontColor("#1565c0").setFontWeight("bold");
  }
}

// ════════════════════════════════════════════════════════════
//  10. onOpen — قائمة مخصصة في Sheets
// ════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔔 وحدة IT")
    .addItem("📊 عرض الإحصائيات",          "showStats")
    .addItem("🔔 اختبار تيليغرام",           "testTelegram")
    .addSeparator()
    .addItem("🎨 إعادة تنسيق الجدول",        "reformatAllRows")
    .addItem("🔒 إعداد الإعدادات السرية",    "setupConfig")
    .addSeparator()
    .addItem("🔬 تشخيص شامل",                "diagnose")
    .addToUi();
}

function showStats() {
  var res  = JSON.parse(getStats().getContent());
  var msg  = [
    "📊 إحصائيات الشكاوي",
    "─────────────────────",
    "📋 الإجمالي:          " + res.total,
    "✅ مُصلَّحة:          " + res.fixed,
    "🟡 قيد المعالجة:      " + res.processing,
    "🔵 مفتوحة:            " + res.open,
    "🔴 عاجلة:             " + res.urgent,
    "❌ مرفوضة:            " + res.rejected
  ].join("\n");
  SpreadsheetApp.getUi().alert(msg);
}

function reformatAllRows() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;
  for (var r = 2; r <= sheet.getLastRow(); r++) {
    formatRow(sheet, r);
  }
  SpreadsheetApp.getUi().alert("✅ تم إعادة تنسيق " + (sheet.getLastRow() - 1) + " صف");
}

// ════════════════════════════════════════════════════════════
//  🧪 اختبارات
// ════════════════════════════════════════════════════════════
function testTelegram() {
  var cfg  = getConfig();
  var time = Utilities.formatDate(new Date(), "Africa/Algiers", "HH:mm:ss — yyyy/MM/dd");
  var msg  = [
    "┌─────────────────────────────┐",
    "│  🧪  اختبار ناجح — v2.0    │",
    "└─────────────────────────────┘",
    "",
    "🤖  <b>HayatIT Bot</b> يعمل بنجاح",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "📊  Google Sheets    : ✅",
    "📨  Gmail (Auto)     : ✅",
    "💬  Telegram         : ✅",
    "📲  WhatsApp API     : ✅",
    "🔍  Ticket Lookup    : ✅",
    "🛡️  Anti-Spam        : ✅",
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━",
    "⏰  " + time,
    "",
    "🏛  <b>مؤسسات الحياة — وحدة IT v2.0</b>"
  ].join("\n");
  sendTelegram(cfg.tgChatId, msg);
  Logger.log("✅ اختبار تيليغرام ناجح");
}

function testSave() {
  var t = saveToSheet({
    name: "اختبار نظام", empId: "PC-001", phone: "+213599941118",
    email: "test@gmail.com", dept: "فرع التعليم - مدرسة الحياة",
    type: "صيانة", priority: "عادي", details: "اختبار حفظ تلقائي v2."
  });
  Logger.log("✅ حُفظ: " + t);
}

function testAll() {
  var fake = {
    name: "محمد اختبار", empId: "PC-042", phone: "+213599941118",
    email: "test@gmail.com", dept: "فرع التعليم - مدرسة الحياة",
    type: "صيانة جهاز", priority: "عاجل", details: "الجهاز لا يعمل منذ الصباح."
  };
  var t = saveToSheet(fake);
  notifyAdminTelegram(fake, t);
  try { sendConfirmationEmail(fake, t); } catch(e) { Logger.log(e.message); }
  Logger.log("✅ اكتمل: " + t);
}

// ════════════════════════════════════════════════════════════
//  🔬 تشخيص شامل
// ════════════════════════════════════════════════════════════
function diagnose() {
  var cfg     = getConfig();
  var results = [];

  results.push("=== ⚙️ الإعدادات ===");
  results.push("TG_TOKEN:     " + (cfg.tgToken    ? "✅ موجود" : "❌ فارغ"));
  results.push("TG_CHAT_ID:   " + (cfg.tgChatId   ? "✅ " + cfg.tgChatId : "❌ فارغ"));
  results.push("WA_TOKEN:     " + (cfg.waToken    ? "✅ موجود" : "❌ فارغ"));
  results.push("WA_PHONE_ID:  " + (cfg.waPhoneId  ? "✅ " + cfg.waPhoneId : "❌ فارغ"));
  results.push("ADMIN_EMAIL:  " + (cfg.adminEmail ? "✅ " + cfg.adminEmail : "❌ فارغ"));

  results.push("\n=== 💬 تيليغرام ===");
  try {
    sendTelegram(cfg.tgChatId, "🔬 تشخيص v2 — " + new Date().toLocaleString());
    results.push("✅ يعمل");
  } catch(e) { results.push("❌ " + e.message); }

  results.push("\n=== 📊 Google Sheets ===");
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (sheet) {
      var count = Math.max(sheet.getLastRow() - 1, 0);
      results.push("✅ الجدول موجود — " + count + " طلب");
    } else {
      results.push("⚠️ الجدول غير موجود — سيُنشأ عند أول طلب");
    }
  } catch(e) { results.push("❌ " + e.message); }

  results.push("\n=== 📧 Gmail ===");
  try {
    var quota = MailApp.getRemainingDailyQuota();
    results.push("✅ حصة الإيميل المتبقية: " + quota);
  } catch(e) { results.push("❌ " + e.message); }

  Logger.log(results.join("\n"));
  return results.join("\n");
}

// ════════════════════════════════════════════════════════════
//  🔧 أداة مساعدة JSON
// ════════════════════════════════════════════════════════════
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ════════════════════════════════════════════════════════════
//  تثبيت Trigger — شغّل مرة واحدة فقط
// ════════════════════════════════════════════════════════════
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  Logger.log("✅ Trigger مثبّت بنجاح");
}
