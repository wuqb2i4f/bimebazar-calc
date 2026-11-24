// ====================== تنظیمات اولیه ======================
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // اگر جداست، آیدی شیت رو بگذارید
const EXPERTS_SHEET = "Experts";
const RECORDS_SHEET = "Records";

// ستون‌های جدول نهایی
const COLUMNS = [
  "شیفت",
  "نام کارشناس",
  "تعداد پردازش اولیه ثالث",
  "تعداد پردازش ثانویه ثالث",
  "تعداد صدور ثالث زیر 15 دقیقه",
  "تعداد صدور ثالث 15 تا 20 دقیقه",
  "تعداد صدور ثالث 20 تا 30 دقیقه",
  "تعداد صدور ثالث بالای 30 دقیقه",
  "تعداد سفارشات لغو شده",
  "مجموع صدور ثالث",
  "فاصله تا پله بعدی حقوق ثابت",
  "حقوق ثابت",
  "حقوق متغیر",
  "حقوق قبل از جریمه خطای صدور",
  "تعداد خطای صدور",
  "نرخ خطای صدور",
  "حقوق بعد از جریمه خطای صدور",
  "تاریخ ثبت"
];

// ====================== توابع کمکی ======================
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function setupSheets() {
  // شیت Experts
  const expSheet = getSheet(EXPERTS_SHEET);
  if (expSheet.getLastRow() === 0) {
    expSheet.appendRow(["نام کارشناس", "شیفت"]);
    // چند نمونه اولیه
    expSheet.appendRow(["علی رضایی", "صبح"]);
    expSheet.appendRow(["مریم احمدی", "عصر"]);
    expSheet.appendRow(["حسین محمدی", "شب"]);
  }

  // شیت Records
  const recSheet = getSheet(RECORDS_SHEET);
  if (recSheet.getLastRow() === 0) {
    recSheet.appendRow(COLUMNS);
  }
}

// ====================== Web App ======================
function doGet() {
  setupSheets();
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("سیستم ثبت عملکرد کارشناسان")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// لیست کارشناس‌ها برای دراپ‌داون
function getExperts() {
  const sheet = getSheet(EXPERTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const experts = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      experts.push({
        name: data[i][0],
        shift: data[i][1] || ""
      });
    }
  }
  return experts;
}

// ذخیره رکورد جدید
function submitRecord(formData) {
  try {
    const sheet = getSheet(RECORDS_SHEET);
    
    // محاسبات پایه
    const primary = Number(formData.primaryProcess) || 0;
    const secondary = Number(formData.secondaryProcess) || 0;
    const under15 = Number(formData.under15) || 0;
    const between15_20 = Number(formData.between15_20) || 0;
    const between20_30 = Number(formData.between20_30) || 0;
    const over30 = Number(formData.over30) || 0;
    const canceled = Number(formData.canceled) || 0;
    const errors = Number(formData.errors) || 0;

    const totalIssuance = under15 + between15_20 + between20_30 + over30;
    
    // نرخ خطای صدور (درصد)
    const errorRate = totalIssuance > 0 ? (errors / totalIssuance) * 100 : 0;

    // ===== محاسبات حقوق (فعلاً placeholder) =====
    // اینجا باید فرمول‌های واقعی حقوق ثابت، متغیر، پله‌ها و جریمه رو بگذارید
    const fixedSalary = 0;                    // حقوق ثابت
    const variableSalary = 0;                 // حقوق متغیر
    const salaryBeforePenalty = fixedSalary + variableSalary;
    const salaryAfterPenalty = salaryBeforePenalty; // فعلاً بدون جریمه
    const distanceToNextStep = 0;             // فاصله تا پله بعدی

    const row = [
      formData.shift,
      formData.expertName,
      primary,
      secondary,
      under15,
      between15_20,
      between20_30,
      over30,
      canceled,
      totalIssuance,
      distanceToNextStep,
      fixedSalary,
      variableSalary,
      salaryBeforePenalty,
      errors,
      Number(errorRate.toFixed(2)),
      salaryAfterPenalty,
      new Date()
    ];

    sheet.appendRow(row);

    return { success: true, message: "اطلاعات با موفقیت ثبت شد." };
  } catch (e) {
    return { success: false, message: "خطا: " + e.message };
  }
}

// برای ادمین: اضافه کردن کارشناس جدید
function addExpert(name, shift) {
  const sheet = getSheet(EXPERTS_SHEET);
  sheet.appendRow([name, shift]);
  return { success: true };
}

// لیست آخرین رکوردها (اختیاری)
function getRecentRecords(limit = 20) {
  const sheet = getSheet(RECORDS_SHEET);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const records = data.slice(1).reverse().slice(0, limit);
  return records;
}