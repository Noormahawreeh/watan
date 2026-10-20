const CONFIG = {
  SPREADSHEET_ID: '1A1suXfnMavXC1RVQmqg_-fF-4F5NxUL1AjwtwD_aLeA',
  TIMEZONE: 'Asia/Amman',
  PREFIX: 'WT-JOIN',
  SHEETS: {
    APPLICATIONS: 'Applications',
    SETTINGS: 'Settings'
  }
};

/**
 * فتح واجهة المنصة.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('وطن تِك | منصة الانتساب')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * الوصول إلى قاعدة البيانات.
 */
function getDatabase_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * إنشاء أوراق النظام وعناوين الأعمدة.
 * شغّل هذه الدالة مرة واحدة بعد لصق الكود.
 */
function setupSystem() {
  const ss = getDatabase_();

  const applicationHeaders = [
    'رقم الطلب',
    'تاريخ التقديم',
    'الاسم الكامل',
    'تاريخ الميلاد',
    'المحافظة',
    'رقم الهاتف',
    'البريد الإلكتروني',
    'المؤهل العلمي',
    'التخصص',
    'المؤسسة التعليمية',
    'المجال الرئيسي',
    'المجالات الفرعية',
    'المهارات',
    'مستوى المهارة',
    'الخبرات السابقة',
    'روابط الأعمال',
    'سبب الانضمام',
    'المساهمة المتوقعة',
    'الساعات الأسبوعية',
    'الالتزام',
    'حالة الطلب',
    'ملاحظات الإدارة'
  ];

  let applicationsSheet = ss.getSheetByName(
    CONFIG.SHEETS.APPLICATIONS
  );

  if (!applicationsSheet) {
    applicationsSheet = ss.insertSheet(
      CONFIG.SHEETS.APPLICATIONS
    );
  }

  if (applicationsSheet.getLastRow() === 0) {
    applicationsSheet
      .getRange(1, 1, 1, applicationHeaders.length)
      .setValues([applicationHeaders]);

    applicationsSheet
      .getRange(1, 1, 1, applicationHeaders.length)
      .setFontWeight('bold')
      .setBackground('#A60002')
      .setFontColor('#FFFFFF');

    applicationsSheet.setFrozenRows(1);

    applicationsSheet.autoResizeColumns(
      1,
      applicationHeaders.length
    );
  }

  let settingsSheet = ss.getSheetByName(
    CONFIG.SHEETS.SETTINGS
  );

  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(
      CONFIG.SHEETS.SETTINGS
    );
  }

  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.getRange(1, 1, 3, 2).setValues([
      ['الإعداد', 'القيمة'],
      ['اسم الفريق', 'فريق وطن تِك | WATAN TECH'],
      ['المنطقة الزمنية', CONFIG.TIMEZONE]
    ]);

    settingsSheet
      .getRange(1, 1, 1, 2)
      .setFontWeight('bold')
      .setBackground('#A60002')
      .setFontColor('#FFFFFF');

    settingsSheet.setFrozenRows(1);
  }

  return {
    success: true,
    message: 'تم تجهيز قاعدة البيانات بنجاح.',
    spreadsheetName: ss.getName(),
    sheets: [
      CONFIG.SHEETS.APPLICATIONS,
      CONFIG.SHEETS.SETTINGS
    ]
  };
}

/**
 * إنشاء رقم طلب جديد مع منع التكرار.
 */
function generateApplicationId_() {
  const properties = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();

  lock.waitLock(10000);

  try {
    const ss = getDatabase_();
    const sheet = ss.getSheetByName(
      CONFIG.SHEETS.APPLICATIONS
    );

    const year = Utilities.formatDate(
      new Date(),
      CONFIG.TIMEZONE,
      'yyyy'
    );

    const counterKey = 'APPLICATION_COUNTER_' + year;

    let counter = Number(
      properties.getProperty(counterKey) || 0
    );

    if (sheet && sheet.getLastRow() > 1) {
      const ids = sheet
        .getRange(2, 1, sheet.getLastRow() - 1, 1)
        .getDisplayValues();

      ids.forEach(function (row) {
        const match = String(row[0]).match(
          new RegExp('^WT-JOIN-' + year + '-(\\d+)$')
        );

        if (match) {
          counter = Math.max(counter, Number(match[1]));
        }
      });
    }

    counter++;

    properties.setProperty(
      counterKey,
      String(counter)
    );

    return CONFIG.PREFIX + '-' + year + '-' +
      String(counter).padStart(5, '0');

  } finally {
    lock.releaseLock();
  }
}

/**
 * تنظيف النصوص والتحقق من طولها.
 */
function cleanText_(value, maxLength) {
  const text = String(value == null ? '' : value)
    .trim()
    .replace(/\s+/g, ' ');

  if (text.length > maxLength) {
    throw new Error('أحد الحقول يتجاوز الحد المسموح.');
  }

  return text;
}

/**
 * تقديم طلب انتساب جديد.
 */
function submitApplication(form) {
  if (
    !form ||
    typeof form !== 'object' ||
    Array.isArray(form)
  ) {
    throw new Error('بيانات الطلب غير صحيحة.');
  }

  const name = cleanText_(form.fullName, 150);
  const email = cleanText_(form.email, 200).toLowerCase();
  const phone = cleanText_(form.phone, 40);
  const birthDate = cleanText_(form.birthDate, 20);
  const governorate = cleanText_(form.governorate, 100);
  const education = cleanText_(form.education, 150);
  const specialization = cleanText_(form.specialization, 150);
  const institution = cleanText_(form.institution, 200);
  const mainField = cleanText_(form.mainField, 150);
  const subFields = cleanText_(form.subFields, 1000);
  const skills = cleanText_(form.skills, 2000);
  const skillLevel = cleanText_(form.skillLevel, 100);
  const experience = cleanText_(form.experience, 3000);
  const portfolio = cleanText_(form.portfolio, 1000);
  const motivation = cleanText_(form.motivation, 3000);
  const contribution = cleanText_(form.contribution, 2000);
  const weeklyHours = cleanText_(form.weeklyHours, 50);

  if (!name || !email || !phone || !mainField || !motivation) {
    throw new Error(
      'يرجى تعبئة الاسم والبريد الإلكتروني والهاتف والمجال وسبب الانضمام.'
    );
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('يرجى إدخال بريد إلكتروني صحيح.');
  }

  // دعم القيم الشائعة التي قد ترسلها واجهة HTML.
  const consentGiven =
    form.consent === true ||
    form.consent === 'true' ||
    form.consent === 'on';

  if (!consentGiven) {
    throw new Error(
      'يجب الموافقة على سياسة استخدام البيانات قبل التقديم.'
    );
  }

  const ss = getDatabase_();
  const sheet = ss.getSheetByName(
    CONFIG.SHEETS.APPLICATIONS
  );

  if (!sheet) {
    throw new Error(
      'قاعدة البيانات غير مجهزة. شغّل الدالة setupSystem أولًا.'
    );
  }

  const applicationId = generateApplicationId_();

  const row = [
    applicationId,
    new Date(),
    name,
    birthDate,
    governorate,
    phone,
    email,
    education,
    specialization,
    institution,
    mainField,
    subFields,
    skills,
    skillLevel,
    experience,
    portfolio,
    motivation,
    contribution,
    weeklyHours,
    'أوافق على الالتزام',
    'قيد المراجعة',
    ''
  ];

  sheet.appendRow(row);

  return {
    success: true,
    applicationId: applicationId,
    status: 'قيد المراجعة',
    message: 'تم استلام طلبك بنجاح.'
  };
}

/**
 * متابعة الطلب باستخدام رقم الطلب والبريد الإلكتروني.
 */
function getApplicationStatus(applicationId, email) {
  applicationId = cleanText_(applicationId, 50);
  email = cleanText_(email, 200).toLowerCase();

  if (!applicationId || !email) {
    throw new Error('أدخل رقم الطلب والبريد الإلكتروني.');
  }

  const sheet = getDatabase_()
    .getSheetByName(CONFIG.SHEETS.APPLICATIONS);

  // معالجة حالة عدم وجود ورقة الطلبات.
  if (!sheet) {
    throw new Error(
      'قاعدة بيانات الطلبات غير موجودة. تحقق من إعداد النظام.'
    );
  }

  if (sheet.getLastRow() < 2) {
    throw new Error('لا توجد طلبات مسجلة.');
  }

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const rowId = String(rows[i][0]);
    const rowEmail = String(rows[i][6]).toLowerCase();

    if (rowId === applicationId && rowEmail === email) {
      const rawDate = rows[i][1];
      const submittedDate = rawDate instanceof Date
        ? rawDate
        : new Date(rawDate);

      const submittedAt = !isNaN(submittedDate.getTime())
        ? Utilities.formatDate(
            submittedDate,
            CONFIG.TIMEZONE,
            'yyyy-MM-dd HH:mm'
          )
        : String(rawDate || '');

      return {
        success: true,
        applicationId: rowId,
        name: String(rows[i][2]),
        status: String(rows[i][20]),
        submittedAt: submittedAt
      };
    }
  }

  throw new Error(
    'لم نعثر على طلب مطابق. تحقق من رقم الطلب والبريد الإلكتروني.'
  );
}

/**
 * إحصائيات عامة للطلبات.
 */
function getApplicationStats() {
  const sheet = getDatabase_()
    .getSheetByName(CONFIG.SHEETS.APPLICATIONS);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      total: 0,
      underReview: 0,
      accepted: 0,
      rejected: 0
    };
  }

  const rows = sheet
    .getRange(2, 21, sheet.getLastRow() - 1, 1)
    .getDisplayValues();

  const stats = {
    total: rows.length,
    underReview: 0,
    accepted: 0,
    rejected: 0
  };

  rows.forEach(function (row) {
    if (row[0] === 'قيد المراجعة') {
      stats.underReview++;
    } else if (row[0] === 'مقبول') {
      stats.accepted++;
    } else if (row[0] === 'مرفوض') {
      stats.rejected++;
    }
  });

  return stats;
}