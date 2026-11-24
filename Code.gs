// ====================== تنظیمات اولیه ======================
const EXPERTS_SHEET = "Experts";
const RECORDS_SHEET = "Records";
const FIXED_SALARY_SHEET = "Config_FixedSalary";
const PIECE_RATES_SHEET = "Config_PieceRates";
const ERROR_PENALTIES_SHEET = "Config_ErrorPenalties";

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
  "حقوق ثابت",
  "حقوق متغیر",
  "حقوق قبل از جریمه خطای صدور",
  "تعداد خطای صدور",
  "نرخ خطای صدور",
  "حقوق بعد از جریمه خطای صدور"
];

// ====================== راه‌اندازی شیت‌ها ======================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Experts ---
  let sheet = ss.getSheetByName(EXPERTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EXPERTS_SHEET);
    sheet.appendRow(["نام کارشناس", "شیفت"]);
    sheet.appendRow(["علی رضایی", 1]);
    sheet.appendRow(["مریم احمدی", 1]);
    sheet.appendRow(["حسین محمدی", 2]);
    sheet.appendRow(["سارا کریمی", 2]);
  }

  // --- Records ---
  sheet = ss.getSheetByName(RECORDS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RECORDS_SHEET);
    sheet.appendRow([...COLUMNS, "تاریخ ثبت"]);
  } else {
    // اگر هدر قدیمی است، آن را به‌روز می‌کنیم (فقط اگر خالی باشد)
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (header[0] !== "شیفت" || header.length < 16) {
      sheet.clear();
      sheet.appendRow([...COLUMNS, "تاریخ ثبت"]);
    }
  }

  // --- Config_FixedSalary ---
  sheet = ss.getSheetByName(FIXED_SALARY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(FIXED_SALARY_SHEET);
    sheet.appendRow(["شیفت", "از", "تا", "مبلغ حقوق ثابت"]);
    sheet.appendRow([1, 0, 199, 3000000]);
    sheet.appendRow([1, 200, 349, 5000000]);
    sheet.appendRow([1, 350, 499, 8000000]);
    sheet.appendRow([1, 500, 699, 12000000]);
    sheet.appendRow([1, 700, 999999, 15000000]);
    sheet.appendRow([2, 0, 199, 3500000]);
    sheet.appendRow([2, 200, 349, 6000000]);
    sheet.appendRow([2, 350, 499, 9500000]);
    sheet.appendRow([2, 500, 699, 14500000]);
    sheet.appendRow([2, 700, 999999, 18000000]);
  }

  // --- Config_PieceRates ---
  sheet = ss.getSheetByName(PIECE_RATES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PIECE_RATES_SHEET);
    sheet.appendRow(["کلید", "توضیح", "مبلغ"]);
    sheet.appendRow(["initial_third_party", "پردازش اولیه ثالث", 15000]);
    sheet.appendRow(["secondary_third_party", "پردازش ثانویه ثالث", 15000]);
    sheet.appendRow(["issue_under_15", "صدور زیر ۱۵ دقیقه", 48000]);
    sheet.appendRow(["issue_15_to_20", "صدور ۱۵ تا ۲۰ دقیقه", 43000]);
    sheet.appendRow(["issue_20_to_30", "صدور ۲۰ تا ۳۰ دقیقه", 30000]);
    sheet.appendRow(["issue_over_30", "صدور بالای ۳۰ دقیقه", 25000]);
    sheet.appendRow(["issue_cancelled", "سفارشات لغو شده", 7500]);
  }

  // --- Config_ErrorPenalties ---
  sheet = ss.getSheetByName(ERROR_PENALTIES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(ERROR_PENALTIES_SHEET);
    sheet.appendRow(["حداکثر نرخ خطا (اعشاری)", "درصد جریمه"]);
    sheet.appendRow([0.01, 0.00]);
    sheet.appendRow([0.02, 0.10]);
    sheet.appendRow([0.03, 0.15]);
    sheet.appendRow([999, 0.20]);
  }
}

// ====================== خواندن تنظیمات ======================
function getFixedSalaryRules() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FIXED_SALARY_SHEET);
  const data = sheet.getDataRange().getValues();
  const rules = { 1: [], 2: [] };

  for (let i = 1; i < data.length; i++) {
    const shift = Number(data[i][0]);
    const low = Number(data[i][1]);
    const high = Number(data[i][2]);
    const amount = Number(data[i][3]);
    if (shift && !isNaN(low) && !isNaN(amount)) {
      rules[shift].push({ low, high, amount });
    }
  }
  rules[1].sort((a, b) => a.low - b.low);
  rules[2].sort((a, b) => a.low - b.low);
  return rules;
}

function getPieceRates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PIECE_RATES_SHEET);
  const data = sheet.getDataRange().getValues();
  const rates = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      rates[data[i][0]] = Number(data[i][2]) || 0;
    }
  }
  return rates;
}

function getErrorPenalties() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ERROR_PENALTIES_SHEET);
  const data = sheet.getDataRange().getValues();
  const penalties = [];
  for (let i = 1; i < data.length; i++) {
    penalties.push({
      threshold: Number(data[i][0]),
      penalty: Number(data[i][1])
    });
  }
  penalties.sort((a, b) => a.threshold - b.threshold);
  return penalties;
}

// ====================== محاسبات ======================
function calculateFullIssuance(primary, secondary, under15, b15_20, b20_30, over30) {
  const avg = (Number(primary) + Number(secondary)) / 2;
  return Math.floor(avg) + Number(under15) + Number(b15_20) + Number(b20_30) + Number(over30);
}

function calculateFixedSalary(shift, totalIssuance, rules) {
  const shiftRules = rules[shift] || [];
  for (const r of shiftRules) {
    if (totalIssuance >= r.low && totalIssuance <= r.high) {
      return r.amount;
    }
  }
  return 0;
}

function calculateVariableSalary(primary, secondary, under15, b15_20, b20_30, over30, canceled, rates) {
  return (
    Number(primary) * (rates.initial_third_party || 0) +
    Number(secondary) * (rates.secondary_third_party || 0) +
    Number(under15) * (rates.issue_under_15 || 0) +
    Number(b15_20) * (rates.issue_15_to_20 || 0) +
    Number(b20_30) * (rates.issue_20_to_30 || 0) +
    Number(over30) * (rates.issue_over_30 || 0) +
    Number(canceled) * (rates.issue_cancelled || 0)
  );
}

function applyErrorPenalty(salaryBefore, errorRate, penalties) {
  let rate = Number(errorRate);
  if (rate > 1) rate = rate / 100;

  let penaltyPct = 0.20;
  for (const p of penalties) {
    if (rate <= p.threshold) {
      penaltyPct = p.penalty;
      break;
    }
  }
  return Math.round(salaryBefore * (1 - penaltyPct));
}

// ====================== توابع وب‌اپ ======================
function doGet() {
  setupSheets();
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("سیستم حقوق کارشناسان ثالث")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getExperts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPERTS_SHEET);
  const data = sheet.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      list.push({
        name: String(data[i][0]),
        shift: Number(data[i][1]) || 1
      });
    }
  }
  return list;
}

function getRecords(limit) {
  limit = limit || 200;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECORDS_SHEET);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 17).getValues(); // ۱۶ ستون داده + تاریخ
  
  // جدیدترین‌ها اول
  const records = data.reverse().slice(0, limit);
  
  // تبدیل تاریخ به رشته خوانا و اطمینان از عدد بودن مقادیر
  return records.map(row => {
    return row.map((cell, idx) => {
      if (idx === 16 && cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
      }
      return cell;
    });
  });
}

function submitRecord(formData) {
  try {
    const rules = getFixedSalaryRules();
    const rates = getPieceRates();
    const penalties = getErrorPenalties();

    const primary = Number(formData.primaryProcess) || 0;
    const secondary = Number(formData.secondaryProcess) || 0;
    const under15 = Number(formData.under15) || 0;
    const b15_20 = Number(formData.between15_20) || 0;
    const b20_30 = Number(formData.between20_30) || 0;
    const over30 = Number(formData.over30) || 0;
    const canceled = Number(formData.canceled) || 0;
    const errors = Number(formData.errors) || 0;
    const shift = Number(formData.shift) || 1;

    const totalIssuance = calculateFullIssuance(primary, secondary, under15, b15_20, b20_30, over30);
    const fixedSalary = calculateFixedSalary(shift, totalIssuance, rules);
    const variableSalary = calculateVariableSalary(primary, secondary, under15, b15_20, b20_30, over30, canceled, rates);
    const salaryBefore = fixedSalary + variableSalary;
    const errorRate = totalIssuance > 0 ? (errors / totalIssuance) : 0;
    const salaryAfter = applyErrorPenalty(salaryBefore, errorRate, penalties);

    const row = [
      shift,
      formData.expertName,
      primary,
      secondary,
      under15,
      b15_20,
      b20_30,
      over30,
      canceled,
      totalIssuance,
      fixedSalary,
      variableSalary,
      salaryBefore,
      errors,
      Number(errorRate.toFixed(4)),
      salaryAfter,
      new Date()
    ];

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RECORDS_SHEET);
    sheet.appendRow(row);

    return { success: true, message: "رکورد با موفقیت ثبت شد." };
  } catch (e) {
    return { success: false, message: "خطا: " + e.message };
  }
}