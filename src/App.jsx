import React, { useState, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

const SCHOOL_NAME = "Habib Shining Star School";
const SCHOOL_ADDRESS = "Multan, Pakistan";

const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const CLASSES = ["PG","Nursery","Prep","KG","1","2","3","4","5","6","7","8","9","10","11","12"];
const PAPERS_FUND_MONTHS = ["April","August"];
const ITEM_TYPES = ["Uniform - Summer","Uniform - Winter","Books set","Notebooks","Bag","Shoes","Tie/Belt","Other"];
const PETTY_SUGGESTIONS = ["Director","School Fund","Stationery","Utilities","Maintenance","Transport","Refreshments","Miscellaneous"];
const LEDGER_SUGGESTIONS = ["Director","Bank","Vendor","Utility Company","Landlord","School Fund","Contractor","Supplier"];
const REGULAR_HOURS_PER_DAY = 8;
const OVERTIME_MULTIPLIER = 1.5;

// ---------- Language / translations ----------
// Covers navigation, login, Dashboard, and My Duty in full (the screens
// teachers use every day), plus common actions/status words used
// throughout the app. Screens not yet translated (most admin-only forms
// and tables — Fee Ledger, Staff records, Petty Cash, etc.) simply stay in
// English regardless of language selected; see README for how to extend.
const TRANSLATIONS = {
  en: {
    schoolTagline: "Sign in to continue",
    dashboard: "Dashboard", myDuty: "My duty", staffRecords: "Staff records", admissions: "Admissions",
    feeLedger: "Fee ledger", uniformsBooks: "Uniforms & books", pettyCash: "Petty cash", staffSalary: "Staff salary",
    ledgerReport: "Ledger report", studentAttendance: "Student attendance", staffAttendance: "Staff attendance",
    staffLogins: "Staff logins",
    settingsAccess: "⚙ Settings & access", logOut: "↩ Log out",
    syncedCloud: "☁ Synced across devices", syncedLocal: "💻 Local only (this device)",
    admin: "Admin", staff: "Staff", administrator: "Administrator",
    username: "Username", password: "Password", signIn: "Sign in",
    yourName: "Your name (for record tracking)", adminPassword: "Admin password", staffAccessCode: "Staff access code",
    incorrectPassword: "Incorrect password. Please try again.", incorrectLogin: "Incorrect username or password.",
    save: "Save", cancel: "Cancel", edit: "Edit", delete: "Delete", update: "Update", add: "Add",
    approve: "Approve", reject: "Reject", exportExcel: "⬇ Export Excel",
    active: "Active", inactive: "Inactive", present: "Present", absent: "Absent",
    paid: "Paid", unpaid: "Unpaid", partial: "Partial", pending: "Pending", approved: "Approved", rejected: "Rejected", free: "Free/exempt",
    activeStaff: "Active staff", activeStudents: "Active students", collected: "Collected", overdueStudents: "Overdue students", totalDuesPending: "Total dues pending",
    thisMonthGlance: "This month at a glance", studentsPaidFee: "Students paid fee", studentsDue: "Students due",
    uniformFundCollected: "Uniform fund collected", booksFundCollected: "Books fund collected",
    staffAttendanceToday: "Staff attendance", today: "today", presentToday: "Present today", absentToday: "Absent today", notMarkedYet: "Not marked yet",
    recentAdmissions: "Recent admissions", downloadReport: "⬇ Download full report (Excel)",
    startDuty: "▶ Start Duty", endDuty: "⏹ End Duty",
    notStartedDuty: "You haven't started duty today.", onDutySince: "On duty since", dutyCompleted: "Duty completed",
    hoursToday: "hours today", includesOvertime: "includes overtime",
    awaitingApproval: "Awaiting admin approval — this won't count toward salary until approved.",
    entryRejected: "This entry was rejected by an admin. Contact them if this looks wrong.",
    recentHistory: "Recent history", last10: "last 10", noHistoryYet: "No duty history yet.",
    date: "Date", checkIn: "Check in", checkOut: "Check out", totalHrs: "Total hrs", approval: "Approval",
    language: "Language",
  },
  ur: {
    schoolTagline: "جاری رکھنے کے لیے سائن اِن کریں",
    dashboard: "ڈیش بورڈ", myDuty: "میری ڈیوٹی", staffRecords: "عملے کا ریکارڈ", admissions: "داخلے",
    feeLedger: "فیس لیجر", uniformsBooks: "یونیفارم اور کتابیں", pettyCash: "پیٹی کیش", staffSalary: "عملے کی تنخواہ",
    ledgerReport: "لیجر رپورٹ", studentAttendance: "طلبہ کی حاضری", staffAttendance: "عملے کی حاضری",
    staffLogins: "عملے کے لاگ اِن",
    settingsAccess: "⚙ ترتیبات اور رسائی", logOut: "↩ لاگ آؤٹ",
    syncedCloud: "☁ تمام آلات پر مطابقت پذیر", syncedLocal: "💻 صرف یہی آلہ",
    admin: "ایڈمن", staff: "عملہ", administrator: "منتظم",
    username: "یوزر نیم", password: "پاس ورڈ", signIn: "سائن اِن کریں",
    yourName: "آپ کا نام (ریکارڈ کے لیے)", adminPassword: "ایڈمن پاس ورڈ", staffAccessCode: "عملے کا رسائی کوڈ",
    incorrectPassword: "غلط پاس ورڈ۔ دوبارہ کوشش کریں۔", incorrectLogin: "غلط یوزر نیم یا پاس ورڈ۔",
    save: "محفوظ کریں", cancel: "منسوخ کریں", edit: "ترمیم", delete: "حذف کریں", update: "اپ ڈیٹ کریں", add: "شامل کریں",
    approve: "منظور کریں", reject: "مسترد کریں", exportExcel: "⬇ ایکسل میں ایکسپورٹ کریں",
    active: "فعال", inactive: "غیر فعال", present: "حاضر", absent: "غیر حاضر",
    paid: "ادا شدہ", unpaid: "غیر ادا شدہ", partial: "جزوی", pending: "زیرِ التوا", approved: "منظور شدہ", rejected: "مسترد شدہ", free: "مفت/مستثنیٰ",
    activeStaff: "فعال عملہ", activeStudents: "فعال طلبہ", collected: "وصول شدہ", overdueStudents: "واجب الادا طلبہ", totalDuesPending: "کل بقایا رقم",
    thisMonthGlance: "اس ماہ کا خلاصہ", studentsPaidFee: "فیس ادا کرنے والے طلبہ", studentsDue: "واجب الادا طلبہ",
    uniformFundCollected: "یونیفارم فنڈ وصول شدہ", booksFundCollected: "کتابوں کا فنڈ وصول شدہ",
    staffAttendanceToday: "عملے کی حاضری", today: "آج", presentToday: "آج حاضر", absentToday: "آج غیر حاضر", notMarkedYet: "ابھی درج نہیں ہوا",
    recentAdmissions: "حالیہ داخلے", downloadReport: "⬇ مکمل رپورٹ ڈاؤن لوڈ کریں (ایکسل)",
    startDuty: "▶ ڈیوٹی شروع کریں", endDuty: "⏹ ڈیوٹی ختم کریں",
    notStartedDuty: "آپ نے آج ابھی ڈیوٹی شروع نہیں کی۔", onDutySince: "ڈیوٹی پر ہیں، وقت:", dutyCompleted: "ڈیوٹی مکمل ہوئی",
    hoursToday: "آج کے گھنٹے", includesOvertime: "اووَر ٹائم شامل ہے",
    awaitingApproval: "ایڈمن کی منظوری کا انتظار ہے — منظوری تک یہ تنخواہ میں شمار نہیں ہوگا۔",
    entryRejected: "یہ اندراج ایڈمن نے مسترد کر دیا ہے۔ اگر یہ درست نہیں لگتا تو ان سے رابطہ کریں۔",
    recentHistory: "حالیہ تاریخ", last10: "آخری 10", noHistoryYet: "ابھی تک کوئی ڈیوٹی ریکارڈ نہیں۔",
    date: "تاریخ", checkIn: "حاضری وقت", checkOut: "روانگی وقت", totalHrs: "کل گھنٹے", approval: "منظوری",
    language: "زبان",
  },
};
function useTranslate(lang) {
  const dict = TRANSLATIONS[lang] || TRANSLATIONS.en;
  return (key) => dict[key] || TRANSLATIONS.en[key] || key;
}

const ALL_TABS = [
  { key: "dashboard", label: "Dashboard", labelKey: "dashboard", i: "01" },
  { key: "myduty", label: "My duty", labelKey: "myDuty", i: "02", requiresStaffLink: true },
  { key: "staff", label: "Staff records", labelKey: "staffRecords", i: "03" },
  { key: "admissions", label: "Admissions", labelKey: "admissions", i: "04" },
  { key: "fees", label: "Fee ledger", labelKey: "feeLedger", i: "05" },
  { key: "items", label: "Uniforms & books", labelKey: "uniformsBooks", i: "06" },
  { key: "pettycash", label: "Petty cash", labelKey: "pettyCash", i: "07" },
  { key: "salary", label: "Staff salary", labelKey: "staffSalary", i: "08" },
  { key: "ledger", label: "Ledger report", labelKey: "ledgerReport", i: "09" },
  { key: "studentattendance", label: "Student attendance", labelKey: "studentAttendance", i: "10" },
  { key: "staffattendance", label: "Staff attendance", labelKey: "staffAttendance", i: "11" },
  { key: "useraccounts", label: "Staff logins", labelKey: "staffLogins", i: "12", adminOnly: true },
];

function resolveTabAccess(user, settings, tabKey) {
  if (!user) return false;
  if (user.role === "admin") return true;
  const account = user.accountId ? (settings.userAccounts || []).find((a) => a.id === user.accountId) : null;
  if (account?.customTabAccess) return account.tabAccess?.[tabKey] !== false;
  return settings.staffTabAccess?.[tabKey] !== false;
}
// Finds the first tab a user is allowed to land on — used right after login
// and after a page refresh, so nobody lands on a blank/blocked screen if
// Dashboard (or anything else) has been restricted for their account.
function firstAccessibleTab(user, settings) {
  if (!user) return "dashboard";
  if (user.role === "admin") return "dashboard";
  if (user.staffId) return "myduty";
  for (const t of ALL_TABS) {
    if (t.adminOnly || t.requiresStaffLink) continue;
    if (resolveTabAccess(user, settings, t.key)) return t.key;
  }
  return "noaccess";
}

function uid(prefix) {
  return prefix + "-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}
function currency(n) {
  const v = Number(n) || 0;
  return "Rs " + v.toLocaleString("en-IN", { maximumFractionDigits: 0 });
}
function todayISO() {
  return new Date().toISOString().slice(0, 10);
}
function nowMonthYear() {
  const now = new Date();
  return { monthIdx: now.getMonth(), year: now.getFullYear() };
}
function formatDateTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  return d.toLocaleDateString("en-GB") + " " + d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
}

function computeNetFee(standardFee, discountType, discountValue) {
  const std = Number(standardFee) || 0;
  const dv = Number(discountValue) || 0;
  let discount = 0;
  if (discountType === "percent") discount = std * (dv / 100);
  else if (discountType === "amount") discount = dv;
  discount = Math.min(Math.max(discount, 0), std);
  return { discount, netFee: std - discount };
}
function feeTotalDue(fee) {
  if (fee.status === "free") return 0;
  return (Number(fee.netFee) || 0) + (Number(fee.papersFund) || 0) + (Number(fee.lateFee) || 0);
}
function feeBalance(fee) {
  if (fee.status === "free") return 0;
  return feeTotalDue(fee) - (Number(fee.paidAmount) || 0);
}
// due-ness: "free" | "paid" | "due" (current month, unpaid/partial) | "overdue" (past month, unpaid/partial)
function feeDueState(fee) {
  if (fee.status === "free") return "free";
  if (fee.status === "paid") return "paid";
  const { monthIdx, year } = nowMonthYear();
  const feeMonthIdx = MONTHS.indexOf(fee.month);
  if (fee.year < year || (fee.year === year && feeMonthIdx < monthIdx)) return "overdue";
  return "due";
}
function itemTotal(item) { return (Number(item.quantity) || 0) * (Number(item.unitPrice) || 0); }
function itemBalance(item) { return itemTotal(item) - (Number(item.paidAmount) || 0); }
function computeItemStatus(item) {
  const total = itemTotal(item);
  const paid = Number(item.paidAmount) || 0;
  if (paid <= 0) return "unpaid";
  if (paid >= total) return "paid";
  return "partial";
}
function stampNew(user) { return { enteredBy: user.name, enteredRole: user.role, enteredAt: new Date().toISOString() }; }
function stampEdit(user) { return { editedBy: user.name, editedRole: user.role, editedAt: new Date().toISOString() }; }

function salaryNet(sal) { return (Number(sal.baseSalary) || 0) + (Number(sal.bonus) || 0) - (Number(sal.deduction) || 0); }
function salaryBalance(sal) { return salaryNet(sal) - (Number(sal.paidAmount) || 0); }
function computeSalaryStatus(sal) {
  const total = salaryNet(sal);
  const paid = Number(sal.paidAmount) || 0;
  if (paid <= 0) return "unpaid";
  if (paid >= total) return "paid";
  return "partial";
}
// due-ness for salary rows, mirrors feeDueState
function salaryDueState(sal) {
  if (sal.status === "paid") return "paid";
  const { monthIdx, year } = nowMonthYear();
  const salMonthIdx = MONTHS.indexOf(sal.month);
  if (sal.year < year || (sal.year === year && salMonthIdx < monthIdx)) return "overdue";
  return "due";
}

function dateMonthYear(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  return { month: MONTHS[d.getMonth()], year: d.getFullYear() };
}

// Hours worked between two "HH:MM" (24-hour) time strings on the same day.
function computeHoursFromTimes(checkIn, checkOut) {
  if (!checkIn || !checkOut) return 0;
  const [inH, inM] = checkIn.split(":").map(Number);
  const [outH, outM] = checkOut.split(":").map(Number);
  let mins = (outH * 60 + outM) - (inH * 60 + inM);
  if (mins < 0) mins += 24 * 60; // handles an overnight shift
  return Math.round((mins / 60) * 100) / 100;
}
function splitRegularOvertime(totalHrs) {
  const regular = Math.min(totalHrs, REGULAR_HOURS_PER_DAY);
  const overtime = Math.max(0, totalHrs - REGULAR_HOURS_PER_DAY);
  return { regular, overtime };
}
function nowTimeHHMM() {
  const d = new Date();
  return String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
}
function formatTime12h(hhmm) {
  if (!hhmm) return "—";
  const [h, m] = hhmm.split(":").map(Number);
  const period = h >= 12 ? "PM" : "AM";
  const h12 = h % 12 === 0 ? 12 : h % 12;
  return `${h12}:${String(m).padStart(2, "0")} ${period}`;
}

const STYLES = `
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500&family=Noto+Nastaliq+Urdu:wght@400;600;700&display=swap');

  .sms-root {
    --ink: #1E2A44; --ink-soft: #3B4A6B; --paper: #F6F1E4; --paper-2: #FDFBF5;
    --gold: #B08850; --gold-dark: #8C6529; --green: #3F6B53; --green-bg: #E4EEE6;
    --rust: #A6432E; --rust-bg: #F5E4DF; --amber: #B98418; --amber-bg: #F7EBD3;
    --blue: #3B4A6B; --blue-bg: #E4E9F5; --line: #D9CFB8; --text: #2B2A26; --text-soft: #6B6559;
    font-family: 'Inter', sans-serif; background: var(--paper); color: var(--text);
    min-height: 100vh; display: flex; border-radius: 12px; overflow: hidden;
    border: 1px solid var(--line); box-shadow: 0 1px 3px rgba(30,42,68,0.08); position: relative;
  }
  .sms-serif { font-family: 'Playfair Display', serif; }
  .sms-mono { font-family: 'IBM Plex Mono', monospace; }

  .sms-login-wrap { width: 100%; min-height: 100vh; display: flex; align-items: center; justify-content: center; background: var(--paper); padding: 24px; }
  .sms-login-card { background: var(--paper-2); border: 1px solid var(--line); border-radius: 10px; padding: 32px; width: 100%; max-width: 360px; }
  .sms-login-title { font-size: 20px; color: var(--ink); margin-bottom: 2px; }
  .sms-login-sub { font-size: 11.5px; color: var(--gold-dark); margin-bottom: 22px; }
  .sms-role-toggle { display: flex; gap: 8px; margin-bottom: 16px; }
  .sms-lang-switch { display: flex; gap: 6px; }
  .sms-lang-btn { padding: 5px 12px; border-radius: 20px; border: 1px solid var(--line); background: var(--paper-2); font-size: 12px; font-weight: 600; cursor: pointer; color: var(--text-soft); }
  .sms-lang-btn.active { background: var(--ink); color: #F6F1E4; border-color: var(--ink); }
  .sms-login-lang { display: flex; justify-content: center; margin-bottom: 16px; }

  .sms-root.rtl { direction: rtl; text-align: right; font-family: 'Noto Nastaliq Urdu', 'Inter', sans-serif; }
  .sms-root.rtl .sms-serif { font-family: 'Noto Nastaliq Urdu', 'Playfair Display', serif; }
  .sms-root.rtl .sms-tab { border-left: none; border-right: 3px solid transparent; }
  .sms-root.rtl .sms-tab.active { border-left: none; border-right: 3px solid var(--gold); }
  .sms-root.rtl .sms-tab-index { margin-left: 10px; }
  .sms-root.rtl .sms-name-flag, .sms-root.rtl .sms-checkbox-row, .sms-root.rtl .sms-role-toggle, .sms-root.rtl .sms-toolbar, .sms-root.rtl .sms-header-actions, .sms-root.rtl .sms-field-row { direction: rtl; }
  .sms-root.rtl table.sms-table th, .sms-root.rtl table.sms-table td, .sms-root.rtl table.sms-ledger-table th, .sms-root.rtl table.sms-ledger-table td { text-align: right; }
  .sms-root.rtl .sms-ledger-page { padding-left: 0; padding-right: 28px; }
  .sms-root.rtl .sms-ledger-page::before { left: auto; right: 20px; }
  .sms-role-btn { flex: 1; padding: 9px; border-radius: 6px; border: 1px solid var(--line); background: var(--paper-2); font-size: 13px; font-weight: 500; cursor: pointer; color: var(--text-soft); }
  .sms-role-btn.active { background: var(--ink); color: #F6F1E4; border-color: var(--ink); }
  .sms-login-error { color: var(--rust); font-size: 12px; margin-top: 6px; }
  .sms-login-hint { font-size: 11px; color: var(--text-soft); margin-top: 14px; line-height: 1.5; }

  .sms-duty-card {
    background: var(--paper-2); border: 1px solid var(--line); border-radius: 10px; padding: 28px 24px;
    display: flex; flex-direction: column; align-items: center; gap: 16px; text-align: center; margin-bottom: 24px;
  }
  .sms-duty-status { font-size: 15px; color: var(--ink); display: flex; flex-direction: column; gap: 4px; }
  .sms-duty-btn { padding: 14px 32px; font-size: 15px; border-radius: 8px; }

  .sms-mobile-topbar { display: none; }
  .sms-sidebar-overlay { display: none; }

  .sms-sidebar { width: 224px; background: var(--ink); color: #EFE9D8; display: flex; flex-direction: column; flex-shrink: 0; }
  .sms-brand { padding: 22px 20px 16px 20px; border-bottom: 1px solid rgba(239,233,216,0.15); }
  .sms-brand-title { font-size: 18px; font-weight: 700; line-height: 1.3; color: #F6F1E4; }
  .sms-brand-sub { font-size: 10.5px; letter-spacing: 0.05em; color: var(--gold); margin-top: 6px; line-height: 1.4; }

  .sms-tabs { padding: 14px 0; flex: 1; overflow-y: auto; }
  .sms-tab {
    display: flex; align-items: center; gap: 10px; padding: 11px 20px; font-size: 13.5px; font-weight: 500;
    cursor: pointer; color: #C9C2AA; border-left: 3px solid transparent; position: relative; transition: background 0.12s ease;
  }
  .sms-tab:hover { background: rgba(239,233,216,0.06); }
  .sms-tab.active { background: rgba(176,136,80,0.14); color: #F6F1E4; border-left: 3px solid var(--gold); }
  .sms-tab-index { font-family: 'IBM Plex Mono', monospace; font-size: 10.5px; color: var(--gold); opacity: 0.85; width: 16px; }
  .sms-tab-badge { margin-left: auto; background: var(--rust); color: #fff; font-size: 10px; font-weight: 600; padding: 1px 6px; border-radius: 20px; }

  .sms-sidebar-foot { padding: 14px 20px 16px 20px; border-top: 1px solid rgba(239,233,216,0.15); }
  .sms-sync-badge { font-size: 10.5px; color: #C9C2AA; background: rgba(239,233,216,0.08); border: 1px solid rgba(239,233,216,0.15); border-radius: 5px; padding: 5px 8px; margin-bottom: 10px; }
  .sms-sync-badge.cloud { color: var(--gold); border-color: rgba(176,136,80,0.4); background: rgba(176,136,80,0.1); }
  .sms-user-chip { font-size: 12px; color: #F6F1E4; margin-bottom: 8px; }
  .sms-user-chip .role { color: var(--gold); text-transform: uppercase; font-size: 9.5px; letter-spacing: 0.06em; display: block; margin-top: 2px; }
  .sms-sidebar-link { display: block; width: 100%; text-align: left; background: none; border: none; color: #C9C2AA; font-size: 12px; padding: 5px 0; cursor: pointer; }
  .sms-sidebar-link:hover { color: #F6F1E4; }

  .sms-main { flex: 1; min-width: 0; background: var(--paper-2); display: flex; flex-direction: column; max-height: 92vh; overflow-y: auto; -webkit-overflow-scrolling: touch; }
  .sms-header { padding: 22px 32px 16px 32px; border-bottom: 1px solid var(--line); display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; }
  .sms-header h1 { font-size: 22px; margin: 0; color: var(--ink); }
  .sms-header .sms-datestamp { font-size: 11.5px; color: var(--text-soft); font-family: 'IBM Plex Mono', monospace; }
  .sms-header-actions { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }

  .sms-content { padding: 24px 32px 40px 32px; }

  .sms-toast { position: fixed; top: 18px; right: 18px; z-index: 200; background: var(--green); color: #fff; padding: 12px 18px; border-radius: 8px; font-size: 13px; max-width: 360px; box-shadow: 0 4px 14px rgba(0,0,0,0.18); }

  .sms-alert-banner { display: flex; align-items: center; justify-content: space-between; gap: 12px; background: var(--rust-bg); border: 1px solid var(--rust); color: var(--rust); border-radius: 8px; padding: 12px 16px; font-size: 13px; font-weight: 500; margin-bottom: 22px; }
  .sms-alert-banner .sms-alert-actions { display: flex; gap: 8px; }

  .sms-ledger-page { background: var(--paper-2); border: 1px solid var(--line); border-radius: 6px; position: relative; padding-left: 28px; overflow-x: auto; }
  .sms-ledger-page::before { content: ""; position: absolute; left: 20px; top: 0; bottom: 0; width: 1px; background: var(--rust); opacity: 0.35; }

  .sms-cards-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 24px; }
  .sms-cards-row.three { grid-template-columns: repeat(3, 1fr); }
  .sms-card { background: var(--paper-2); border: 1px solid var(--line); border-radius: 8px; padding: 16px 18px; }
  .sms-card.warn { border-color: var(--rust); background: var(--rust-bg); }
  .sms-card-label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.06em; color: var(--text-soft); margin-bottom: 6px; }
  .sms-card.warn .sms-card-label { color: var(--rust); }
  .sms-card-value { font-family: 'IBM Plex Mono', monospace; font-size: 24px; font-weight: 500; color: var(--ink); }
  .sms-card-value.green { color: var(--green); }
  .sms-card-value.rust { color: var(--rust); }

  .sms-section-title { font-size: 15px; font-weight: 600; color: var(--ink); margin: 28px 0 12px 0; display: flex; align-items: center; gap: 8px; }
  .sms-section-title .sms-tag { font-family: 'IBM Plex Mono', monospace; font-size: 10px; color: var(--gold-dark); background: rgba(176,136,80,0.12); padding: 2px 7px; border-radius: 3px; }

  .sms-toolbar { display: flex; gap: 10px; align-items: center; margin-bottom: 14px; flex-wrap: wrap; }
  .sms-input, .sms-select { border: 1px solid var(--line); background: var(--paper-2); border-radius: 5px; padding: 8px 11px; font-size: 13px; font-family: 'Inter', sans-serif; color: var(--text); outline: none; }
  .sms-input:focus, .sms-select:focus { border-color: var(--gold); }
  .sms-input:disabled, .sms-select:disabled { opacity: 0.5; background: rgba(217,207,184,0.2); }
  .sms-btn { border: 1px solid var(--ink); background: var(--ink); color: #F6F1E4; border-radius: 5px; padding: 8px 14px; font-size: 13px; font-weight: 500; cursor: pointer; white-space: nowrap; }
  .sms-btn:hover { background: var(--ink-soft); }
  .sms-btn:disabled { opacity: 0.45; cursor: not-allowed; }
  .sms-btn.secondary { background: transparent; color: var(--ink); border: 1px solid var(--line); }
  .sms-btn.secondary:hover { background: rgba(30,42,68,0.05); }
  .sms-btn.gold { background: var(--gold); border-color: var(--gold-dark); color: #2B2306; }
  .sms-btn.gold:hover { background: var(--gold-dark); }
  .sms-btn.small { padding: 5px 10px; font-size: 12px; }
  .sms-btn.danger { background: transparent; color: var(--rust); border: 1px solid var(--rust); }
  .sms-btn.danger:hover { background: var(--rust-bg); }

  table.sms-table { width: 100%; border-collapse: collapse; font-size: 13px; min-width: 640px; }
  table.sms-table th { text-align: left; font-size: 10.5px; text-transform: uppercase; letter-spacing: 0.05em; color: var(--text-soft); border-bottom: 1px solid var(--line); padding: 8px 10px; font-weight: 600; white-space: nowrap; }
  table.sms-table td { padding: 10px; border-bottom: 1px solid var(--line); vertical-align: middle; }
  table.sms-table tr:hover td { background: rgba(176,136,80,0.05); }
  table.sms-table tr.sms-row-overdue td { background: var(--rust-bg); }
  table.sms-table tr.sms-row-overdue:hover td { background: #F0D4C9; }
  table.sms-table tr.sms-row-due td { background: var(--amber-bg); }
  table.sms-table tr.sms-row-free td { background: var(--blue-bg); }
  .sms-empty { padding: 40px 20px; text-align: center; color: var(--text-soft); font-size: 13.5px; }
  .sms-subtext { display: block; font-size: 11px; color: var(--text-soft); margin-top: 2px; }
  .sms-locked { font-size: 11px; color: var(--text-soft); font-style: italic; }

  .sms-pill { font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 20px; display: inline-block; white-space: nowrap; }
  .sms-pill.paid { background: var(--green-bg); color: var(--green); }
  .sms-pill.unpaid { background: var(--rust-bg); color: var(--rust); }
  .sms-pill.partial { background: var(--amber-bg); color: var(--amber); }
  .sms-pill.due { background: var(--amber-bg); color: var(--amber); }
  .sms-pill.overdue { background: var(--rust); color: #fff; }
  .sms-pill.free { background: var(--blue-bg); color: var(--blue); }
  .sms-pill.active { background: var(--green-bg); color: var(--green); }
  .sms-pill.inactive { background: rgba(107,101,89,0.12); color: var(--text-soft); }
  .sms-pill.fund { background: rgba(176,136,80,0.16); color: var(--gold-dark); }
  .sms-pill.receipt { background: var(--green-bg); color: var(--green); }
  .sms-pill.expense { background: var(--rust-bg); color: var(--rust); }

  .sms-name-flag { display: inline-flex; align-items: center; gap: 6px; }
  .sms-flag-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--rust); flex-shrink: 0; }

  .sms-modal-overlay { position: fixed; inset: 0; background: rgba(30,42,68,0.45); display: flex; align-items: center; justify-content: center; z-index: 50; padding: 20px; }
  .sms-modal { background: var(--paper-2); border-radius: 8px; width: 100%; max-width: 460px; max-height: 88vh; overflow-y: auto; border: 1px solid var(--line); }
  .sms-modal.wide { max-width: 640px; }
  .sms-modal-head { padding: 18px 22px; border-bottom: 1px solid var(--line); display: flex; justify-content: space-between; align-items: center; }
  .sms-modal-head h3 { margin: 0; font-size: 16px; color: var(--ink); }
  .sms-modal-body { padding: 20px 22px; display: flex; flex-direction: column; gap: 12px; }
  .sms-modal-foot { padding: 16px 22px; border-top: 1px solid var(--line); display: flex; justify-content: flex-end; gap: 10px; }
  .sms-field label { display: block; font-size: 12px; font-weight: 500; color: var(--text-soft); margin-bottom: 5px; }
  .sms-field input, .sms-field select, .sms-field textarea { width: 100%; box-sizing: border-box; }
  .sms-field textarea { border: 1px solid var(--line); background: var(--paper-2); border-radius: 5px; padding: 8px 11px; font-size: 13px; font-family: 'Inter', sans-serif; color: var(--text); outline: none; resize: vertical; min-height: 56px; }
  .sms-field-row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
  .sms-field-row.three { grid-template-columns: 1fr 1fr 1fr; }
  .sms-computed { font-size: 13px; background: rgba(63,107,83,0.08); border: 1px solid var(--green); color: var(--green); border-radius: 5px; padding: 8px 11px; font-family: 'IBM Plex Mono', monospace; }
  .sms-checkbox-row { display: flex; align-items: center; gap: 8px; font-size: 13px; color: var(--text); }
  .sms-checkbox-row input { width: auto; }
  .sms-free-box { background: var(--blue-bg); border: 1px solid var(--blue); border-radius: 6px; padding: 10px 12px; }

  .sms-close-x { cursor: pointer; color: var(--text-soft); font-size: 18px; background: none; border: none; }
  .sms-loading { padding: 60px; text-align: center; color: var(--text-soft); }

  .sms-overdue-list { display: flex; flex-direction: column; gap: 8px; }
  .sms-overdue-item { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; background: var(--rust-bg); border-radius: 6px; border: 1px solid rgba(166,67,46,0.25); }
  .sms-overdue-item .sms-oi-name { font-weight: 600; color: var(--ink); font-size: 13.5px; }
  .sms-overdue-item .sms-oi-meta { font-size: 11.5px; color: var(--text-soft); }
  .sms-overdue-item .sms-oi-amount { font-family: 'IBM Plex Mono', monospace; font-weight: 600; color: var(--rust); }

  @media (max-width: 880px) {
    html, body { overflow-x: hidden; }
    .sms-root { flex-direction: column; border-radius: 0; min-height: 100vh; }

    .sms-mobile-topbar {
      display: flex; align-items: center; gap: 12px; background: var(--ink); color: #F6F1E4;
      padding: 12px 16px; position: sticky; top: 0; z-index: 40; flex-shrink: 0;
    }
    .sms-hamburger-btn {
      background: none; border: 1px solid rgba(239,233,216,0.3); border-radius: 6px; color: #F6F1E4;
      width: 38px; height: 38px; font-size: 18px; cursor: pointer; flex-shrink: 0; display: flex; align-items: center; justify-content: center;
    }
    .sms-mobile-topbar .title { font-size: 14px; font-weight: 600; line-height: 1.3; }
    .sms-mobile-topbar .sub { font-size: 10px; color: var(--gold); }

    .sms-sidebar {
      position: fixed; top: 0; left: 0; bottom: 0; z-index: 60; width: 260px; max-width: 82vw;
      transform: translateX(-100%); transition: transform 0.22s ease; box-shadow: 4px 0 18px rgba(0,0,0,0.25);
    }
    .sms-sidebar.open { transform: translateX(0); }
    .sms-sidebar-overlay {
      display: block; position: fixed; inset: 0; background: rgba(20,24,36,0.5); z-index: 50;
      opacity: 0; pointer-events: none; transition: opacity 0.2s ease;
    }
    .sms-sidebar-overlay.open { opacity: 1; pointer-events: auto; }

    .sms-main { max-height: none; overflow: visible; }
    .sms-header { padding: 16px 16px 12px 16px; }
    .sms-header h1 { font-size: 19px; }
    .sms-content { padding: 16px 16px 32px 16px; }
    .sms-cards-row, .sms-cards-row.three { grid-template-columns: repeat(2, 1fr); gap: 10px; }
    .sms-card { padding: 12px 14px; }
    .sms-card-value { font-size: 19px; }
    .sms-field-row, .sms-field-row.three { grid-template-columns: 1fr; }
    .sms-modal-overlay { padding: 0; align-items: flex-end; }
    .sms-modal { max-width: 100%; width: 100%; max-height: 92vh; border-radius: 14px 14px 0 0; }
    .sms-ledger-page { -webkit-overflow-scrolling: touch; }
    .sms-btn, .sms-input, .sms-select { font-size: 14px; }
  }

  @media (max-width: 460px) {
    .sms-cards-row, .sms-cards-row.three { grid-template-columns: 1fr; }
  }

  .sms-ledger-report-head { margin-bottom: 18px; }
  .sms-ledger-report-head .school { font-size: 20px; color: var(--ink); }
  .sms-ledger-report-head .addr { font-size: 11.5px; color: var(--gold-dark); margin-top: 2px; }
  .sms-ledger-report-head .meta { font-size: 11.5px; color: var(--text-soft); margin-top: 8px; }
  table.sms-ledger-table { width: 100%; border-collapse: collapse; font-size: 11.5px; }
  table.sms-ledger-table th, table.sms-ledger-table td { border: 1px solid var(--line); padding: 6px 8px; text-align: left; }
  table.sms-ledger-table th { background: rgba(30,42,68,0.06); font-size: 10px; text-transform: uppercase; letter-spacing: 0.03em; color: var(--text-soft); }

  @media print {
    .sms-sidebar, .sms-no-print, .sms-mobile-topbar, .sms-sidebar-overlay { display: none !important; }
    .sms-root { border: none !important; box-shadow: none !important; border-radius: 0 !important; display: block !important; }
    .sms-main { max-height: none !important; overflow: visible !important; }
    .sms-content { padding: 0 !important; }
    body { background: #fff !important; }
    @page { size: A4 landscape; margin: 12mm; }
  }
`;

function Modal({ title, onClose, children, footer, wide }) {
  return (
    <div className="sms-modal-overlay" onClick={onClose}>
      <div className={"sms-modal" + (wide ? " wide" : "")} onClick={(e) => e.stopPropagation()}>
        <div className="sms-modal-head">
          <h3 className="sms-serif">{title}</h3>
          <button className="sms-close-x" onClick={onClose}>✕</button>
        </div>
        <div className="sms-modal-body">{children}</div>
        <div className="sms-modal-foot">{footer}</div>
      </div>
    </div>
  );
}
function Field({ label, children }) {
  return (
    <div className="sms-field">
      <label>{label}</label>
      {children}
    </div>
  );
}

function LoginScreen({ settings, onLogin, language, changeLanguage, t }) {
  const hasAccounts = Array.isArray(settings.userAccounts) && settings.userAccounts.length > 0;
  const [role, setRole] = useState("admin");
  const [username, setUsername] = useState("");
  const [name, setName] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  function handleSubmit(e) {
    e.preventDefault();
    setError("");

    if (hasAccounts) {
      // Always allow the master admin password as a safety net, even once
      // individual logins exist, so nobody gets permanently locked out.
      if (username.trim().toLowerCase() === "admin" && password === settings.adminPassword) {
        onLogin({ role: "admin", name: "Admin", accountId: "master-admin", staffId: null });
        return;
      }
      const match = (settings.userAccounts || []).find(
        (a) => a.status !== "inactive" && a.username.trim().toLowerCase() === username.trim().toLowerCase() && a.password === password
      );
      if (match) {
        onLogin({ role: match.role, name: match.name || match.username, accountId: match.id, staffId: match.staffId || null });
      } else {
        setError(t("incorrectLogin"));
      }
      return;
    }

    // Legacy mode: no individual accounts created yet, fall back to the
    // original shared Admin/Staff password.
    const expected = role === "admin" ? settings.adminPassword : settings.staffPassword;
    if (password === expected) {
      onLogin({ role, name: name.trim() || (role === "admin" ? "Admin" : "Staff"), accountId: null, staffId: null });
    } else {
      setError(t("incorrectPassword"));
    }
  }

  return (
    <div className={"sms-root" + (language === "ur" ? " rtl" : "")} dir={language === "ur" ? "rtl" : "ltr"}>
      <style>{STYLES}</style>
      <div className="sms-login-wrap">
        <form className="sms-login-card" onSubmit={handleSubmit}>
          <div className="sms-login-lang">
            <div className="sms-lang-switch">
              <button type="button" className={"sms-lang-btn" + (language === "en" ? " active" : "")} onClick={() => changeLanguage("en")}>English</button>
              <button type="button" className={"sms-lang-btn" + (language === "ur" ? " active" : "")} onClick={() => changeLanguage("ur")}>اردو</button>
            </div>
          </div>
          <div className="sms-login-title sms-serif">{SCHOOL_NAME}</div>
          <div className="sms-login-sub">{SCHOOL_ADDRESS} · {t("schoolTagline")}</div>

          {hasAccounts ? (
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              <Field label={t("username")}>
                <input className="sms-input" value={username} onChange={(e) => setUsername(e.target.value)} placeholder="e.g. sadia.t" autoCapitalize="none" />
              </Field>
              <Field label={t("password")}>
                <input className="sms-input" type="password" value={password} onChange={(e) => setPassword(e.target.value)} />
              </Field>
            </div>
          ) : (
            <>
              <div className="sms-role-toggle">
                <button type="button" className={"sms-role-btn" + (role === "admin" ? " active" : "")} onClick={() => { setRole("admin"); setError(""); }}>{t("admin")}</button>
                <button type="button" className={"sms-role-btn" + (role === "user" ? " active" : "")} onClick={() => { setRole("user"); setError(""); }}>{t("staff")}</button>
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
                <Field label={t("yourName")}>
                  <input className="sms-input" value={name} onChange={(e) => setName(e.target.value)} placeholder={role === "admin" ? "e.g. Principal Habib" : "e.g. Front desk clerk"} />
                </Field>
                <Field label={role === "admin" ? t("adminPassword") : t("staffAccessCode")}>
                  <input className="sms-input" type="password" value={password} onChange={(e) => setPassword(e.target.value)} />
                </Field>
              </div>
            </>
          )}

          {error && <div className="sms-login-error">{error}</div>}
          <button className="sms-btn" type="submit" style={{ width: "100%", marginTop: 14 }}>{t("signIn")}</button>
          <div className="sms-login-hint">
            {hasAccounts
              ? "Ask your admin for your personal username and password if you don't have one yet."
              : "Admins have full edit and delete access. Staff can add new fee, uniform/book, and petty cash entries, but only an admin can edit or delete a saved entry."}
          </div>
        </form>
      </div>
    </div>
  );
}

const SESSION_KEY = "hss_session_user_v1";
function loadSavedSession() {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}
const LANG_KEY = "hss_language_v1";
function loadSavedLanguage() {
  try { return localStorage.getItem(LANG_KEY) || "en"; } catch (e) { return "en"; }
}

export default function SchoolManagementSystem() {
  const [currentUser, setCurrentUser] = useState(() => loadSavedSession());
  const [language, setLanguage] = useState(() => loadSavedLanguage());
  const t = useTranslate(language);
  function changeLanguage(lang) {
    setLanguage(lang);
    try { localStorage.setItem(LANG_KEY, lang); } catch (e) { /* ignore */ }
  }
  const [page, setPage] = useState("dashboard");
  const [loading, setLoading] = useState(true);
  const [staff, setStaff] = useState([]);
  const [students, setStudents] = useState([]);
  const [fees, setFees] = useState([]);
  const [items, setItems] = useState([]);
  const [pettyCash, setPettyCash] = useState([]);
  const [salaries, setSalaries] = useState([]);
  const [ledgerEntries, setLedgerEntries] = useState([]);
  const [studentAttendance, setStudentAttendance] = useState([]);
  const [staffAttendance, setStaffAttendance] = useState([]);
  const [settings, setSettings] = useState({ papersFundAmount: 1000, lateFeeAmount: 200, adminPassword: "admin123", staffPassword: "staff123", ledgerAccountantName: "", ledgerTimePeriod: "", staffTabAccess: {}, userAccounts: [] });
  const [toast, setToast] = useState(null);

  const [staffModal, setStaffModal] = useState(null);
  const [studentModal, setStudentModal] = useState(null);
  const [feeModal, setFeeModal] = useState(null);
  const [itemModal, setItemModal] = useState(null);
  const [pettyModal, setPettyModal] = useState(null);
  const [salaryModal, setSalaryModal] = useState(null);
  const [ledgerModal, setLedgerModal] = useState(null);
  const [staffAttModal, setStaffAttModal] = useState(null);
  const [settingsModal, setSettingsModal] = useState(false);
  const [accountModal, setAccountModal] = useState(null);
  const [showAdmissionLedger, setShowAdmissionLedger] = useState(false);
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [overdueModal, setOverdueModal] = useState(false);
  const [bannerDismissed, setBannerDismissed] = useState(false);
  const [autoPopupShown, setAutoPopupShown] = useState(false);

  const [staffSearch, setStaffSearch] = useState("");
  const [studentSearch, setStudentSearch] = useState("");
  const [itemSearch, setItemSearch] = useState("");
  const [feeMonthFilter, setFeeMonthFilter] = useState(MONTHS[new Date().getMonth()]);
  const [feeYearFilter, setFeeYearFilter] = useState(new Date().getFullYear());
  const [feeStatusFilter, setFeeStatusFilter] = useState("all");
  const [pettyMonthFilter, setPettyMonthFilter] = useState(MONTHS[new Date().getMonth()]);
  const [pettyYearFilter, setPettyYearFilter] = useState(new Date().getFullYear());
  const [salaryMonthFilter, setSalaryMonthFilter] = useState(MONTHS[new Date().getMonth()]);
  const [salaryYearFilter, setSalaryYearFilter] = useState(new Date().getFullYear());
  const [salaryStatusFilter, setSalaryStatusFilter] = useState("all");
  const [attDate, setAttDate] = useState(todayISO());
  const [attClass, setAttClass] = useState(CLASSES[0]);
  const [attDraft, setAttDraft] = useState({});
  const [staffAttStaffId, setStaffAttStaffId] = useState("");
  const [staffAttMonthFilter, setStaffAttMonthFilter] = useState(MONTHS[new Date().getMonth()]);
  const [staffAttYearFilter, setStaffAttYearFilter] = useState(new Date().getFullYear());


  useEffect(() => {
    (async () => {
      try {
        const [s, st, f, it, pc, sal, ledg, satt, satt2, cfg] = await Promise.allSettled([
          window.storage.get("sms-staff"),
          window.storage.get("sms-students"),
          window.storage.get("sms-fees"),
          window.storage.get("sms-items"),
          window.storage.get("sms-pettycash"),
          window.storage.get("sms-salaries"),
          window.storage.get("sms-ledger"),
          window.storage.get("sms-student-attendance"),
          window.storage.get("sms-staff-attendance"),
          window.storage.get("sms-settings"),
        ]);
        if (s.status === "fulfilled" && s.value) setStaff(JSON.parse(s.value.value));
        if (st.status === "fulfilled" && st.value) setStudents(JSON.parse(st.value.value));
        if (f.status === "fulfilled" && f.value) setFees(JSON.parse(f.value.value));
        if (it.status === "fulfilled" && it.value) setItems(JSON.parse(it.value.value));
        if (pc.status === "fulfilled" && pc.value) setPettyCash(JSON.parse(pc.value.value));
        if (sal.status === "fulfilled" && sal.value) setSalaries(JSON.parse(sal.value.value));
        if (ledg.status === "fulfilled" && ledg.value) setLedgerEntries(JSON.parse(ledg.value.value));
        if (satt.status === "fulfilled" && satt.value) setStudentAttendance(JSON.parse(satt.value.value));
        if (satt2.status === "fulfilled" && satt2.value) setStaffAttendance(JSON.parse(satt2.value.value));
        if (cfg.status === "fulfilled" && cfg.value) setSettings((prev) => ({ ...prev, ...JSON.parse(cfg.value.value) }));
      } catch (e) { console.error("Load error", e); }
      setLoading(false);
    })();
  }, []);

  async function persist(key, value) {
    try { await window.storage.set(key, JSON.stringify(value)); } catch (e) { console.error("Save error", e); }
  }
  function saveStaff(next) { setStaff(next); persist("sms-staff", next); }
  function saveStudents(next) { setStudents(next); persist("sms-students", next); }
  function saveFees(next) { setFees(next); persist("sms-fees", next); }
  function saveItems(next) { setItems(next); persist("sms-items", next); }
  function savePettyCash(next) { setPettyCash(next); persist("sms-pettycash", next); }
  function saveSalaries(next) { setSalaries(next); persist("sms-salaries", next); }
  function saveLedger(next) { setLedgerEntries(next); persist("sms-ledger", next); }
  function saveStudentAttendance(next) { setStudentAttendance(next); persist("sms-student-attendance", next); }
  function saveStaffAttendance(next) { setStaffAttendance(next); persist("sms-staff-attendance", next); }
  function saveSettings(next) { setSettings(next); persist("sms-settings", next); }
  function upsertAccount(account) {
    const list = settings.userAccounts || [];
    const exists = list.some((a) => a.id === account.id);
    const next = exists ? list.map((a) => (a.id === account.id ? account : a)) : [...list, account];
    saveSettings({ ...settings, userAccounts: next });
  }
  function deleteAccount(id) {
    saveSettings({ ...settings, userAccounts: (settings.userAccounts || []).filter((a) => a.id !== id) });
  }

  function showToast(msg) { setToast(msg); setTimeout(() => setToast(null), 7000); }

  function handleLogin(user) {
    setCurrentUser(user);
    try { localStorage.setItem(SESSION_KEY, JSON.stringify(user)); } catch (e) { /* ignore */ }
    setPage(firstAccessibleTab(user, settings));
  }
  useEffect(() => {
    if (currentUser) setPage(firstAccessibleTab(currentUser, settings));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);
  function handleLogout() {
    setCurrentUser(null);
    try { localStorage.removeItem(SESSION_KEY); } catch (e) { /* ignore */ }
  }

  const isAdmin = currentUser && currentUser.role === "admin";
  const myAccount = currentUser?.accountId ? (settings.userAccounts || []).find((a) => a.id === currentUser.accountId) : null;
  function hasTabAccess(key) {
    if (myAccount?.customTabAccess) return myAccount.tabAccess?.[key] !== false;
    return settings.staffTabAccess?.[key] !== false;
  }

  useEffect(() => {
    if (!currentUser || isAdmin || page === "noaccess") return;
    const tabDef = ALL_TABS.find((t) => t.key === page);
    const blocked =
      (tabDef?.adminOnly) ||
      (tabDef?.requiresStaffLink && !currentUser.staffId) ||
      (!tabDef?.adminOnly && !tabDef?.requiresStaffLink && !hasTabAccess(page));
    if (blocked) {
      setPage(firstAccessibleTab(currentUser, settings));
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser, isAdmin, page, settings.staffTabAccess, myAccount]);


  function upsertStaff(record) {
    const exists = staff.some((s) => s.id === record.id);
    saveStaff(exists ? staff.map((s) => (s.id === record.id ? record : s)) : [...staff, record]);
    setStaffModal(null);
  }
  function deleteStaff(id) {
    saveStaff(staff.filter((s) => s.id !== id));
    saveSalaries(salaries.filter((sal) => sal.staffId !== id));
  }

  function upsertStudent(record) {
    const exists = students.some((s) => s.id === record.id);
    saveStudents(exists ? students.map((s) => (s.id === record.id ? record : s)) : [...students, record]);
    setStudentModal(null);
  }
  function deleteStudent(id) {
    saveStudents(students.filter((s) => s.id !== id));
    saveFees(fees.filter((f) => f.studentId !== id));
    saveItems(items.filter((i) => i.studentId !== id));
  }

  function upsertFee(record) {
    const prior = fees.find((f) => f.id === record.id);
    let final;
    if (prior) {
      final = { ...prior, ...record, ...stampEdit(currentUser) };
      const priorState = feeDueState(prior);
      if (priorState === "overdue" && final.status === "paid") {
        showToast(`✅ Payment received for ${final.month} ${final.year} (previously overdue) — ${currency(feeTotalDue(final))} recorded by ${currentUser.name} at ${formatDateTime(new Date().toISOString())}.`);
      }
    } else {
      final = { ...record, ...stampNew(currentUser) };
    }
    saveFees(prior ? fees.map((f) => (f.id === final.id ? final : f)) : [...fees, final]);
    setFeeModal(null);
  }
  function deleteFee(id) { saveFees(fees.filter((f) => f.id !== id)); }

  function upsertItem(record) {
    const prior = items.find((i) => i.id === record.id);
    const final = prior ? { ...prior, ...record, ...stampEdit(currentUser) } : { ...record, ...stampNew(currentUser) };
    saveItems(prior ? items.map((i) => (i.id === final.id ? final : i)) : [...items, final]);
    setItemModal(null);
  }
  function deleteItem(id) { saveItems(items.filter((i) => i.id !== id)); }

  function upsertPetty(record) {
    const prior = pettyCash.find((p) => p.id === record.id);
    const final = prior ? { ...prior, ...record, ...stampEdit(currentUser) } : { ...record, ...stampNew(currentUser) };
    savePettyCash(prior ? pettyCash.map((p) => (p.id === final.id ? final : p)) : [...pettyCash, final]);
    setPettyModal(null);
  }
  function deletePetty(id) { savePettyCash(pettyCash.filter((p) => p.id !== id)); }

  function upsertSalary(record) {
    const prior = salaries.find((s) => s.id === record.id);
    let final;
    if (prior) {
      final = { ...prior, ...record, ...stampEdit(currentUser) };
      const priorState = salaryDueState(prior);
      if (priorState === "overdue" && final.status === "paid") {
        showToast(`✅ Salary paid for ${final.month} ${final.year} (was pending) — ${currency(salaryNet(final))} recorded by ${currentUser.name} at ${formatDateTime(new Date().toISOString())}.`);
      }
    } else {
      final = { ...record, ...stampNew(currentUser) };
    }
    saveSalaries(prior ? salaries.map((s) => (s.id === final.id ? final : s)) : [...salaries, final]);
    setSalaryModal(null);
  }
  function deleteSalary(id) { saveSalaries(salaries.filter((s) => s.id !== id)); }

  function generateMonthlySalaries() {
    const month = MONTHS[new Date().getMonth()];
    const year = new Date().getFullYear();
    const existingKeys = new Set(salaries.map((s) => s.staffId + "|" + s.month + "|" + s.year));
    const newRecords = [];
    staff.forEach((st) => {
      if (st.status !== "active") return;
      const key = st.id + "|" + month + "|" + year;
      if (existingKeys.has(key)) return;

      const monthAttendanceAll = staffAttendance.filter((a) => a.staffId === st.id && a.month === month && a.year === year);
      const monthAttendance = monthAttendanceAll.filter((a) => a.approvalStatus !== "pending" && a.approvalStatus !== "rejected");
      const pendingCount = monthAttendanceAll.length - monthAttendance.length;
      if (Number(st.hourlyRate) > 0 && monthAttendance.length > 0) {
        const hourlyRate = Number(st.hourlyRate) || 0;
        const dailyRate = hourlyRate * REGULAR_HOURS_PER_DAY;
        const regularHrs = monthAttendance.reduce((s, a) => s + (Number(a.regularHrs) || 0), 0);
        const overtimeHrs = monthAttendance.reduce((s, a) => s + (Number(a.overtimeHrs) || 0), 0);
        const lwpDays = monthAttendance.filter((a) => a.lwp).length;
        const absentDays = monthAttendance.filter((a) => a.absent).length;
        const advanceSum = monthAttendance.reduce((s, a) => s + (Number(a.advance) || 0), 0);
        const otherSum = monthAttendance.reduce((s, a) => s + (Number(a.other) || 0), 0);
        const regularPay = Math.round(regularHrs * hourlyRate);
        const overtimePay = Math.round(overtimeHrs * hourlyRate * OVERTIME_MULTIPLIER);
        const lwpDeduction = Math.round(lwpDays * dailyRate);
        newRecords.push({
          id: uid("sal"), staffId: st.id, month, year,
          baseSalary: regularPay + overtimePay, bonus: 0, deduction: Math.round(lwpDeduction + advanceSum + otherSum),
          paidAmount: 0, status: "unpaid", paidDate: "",
          comments: `Auto-calculated from attendance: ${regularHrs}h regular + ${overtimeHrs}h overtime, ${lwpDays} LWP day(s), ${absentDays} absent day(s), advance ${currency(advanceSum)}, other ${currency(otherSum)}.` + (pendingCount > 0 ? ` NOTE: ${pendingCount} attendance entr${pendingCount > 1 ? "ies" : "y"} this month are still pending/rejected approval and were NOT included — approve them and regenerate if needed.` : ""),
          ...stampNew(currentUser),
        });
      } else {
        newRecords.push({
          id: uid("sal"), staffId: st.id, month, year,
          baseSalary: Number(st.salary) || 0, bonus: 0, deduction: 0,
          paidAmount: 0, status: "unpaid", paidDate: "", comments: "",
          ...stampNew(currentUser),
        });
      }
    });
    if (newRecords.length) saveSalaries([...salaries, ...newRecords]);
    return newRecords.length;
  }

  function upsertLedger(record) {
    const prior = ledgerEntries.find((l) => l.id === record.id);
    const final = prior ? { ...prior, ...record, ...stampEdit(currentUser) } : { ...record, ...stampNew(currentUser) };
    saveLedger(prior ? ledgerEntries.map((l) => (l.id === final.id ? final : l)) : [...ledgerEntries, final]);
    setLedgerModal(null);
  }
  function deleteLedgerEntry(id) { saveLedger(ledgerEntries.filter((l) => l.id !== id)); }

  function saveStudentAttendanceBulk(date, cls, draft) {
    const activeInClass = students.filter((s) => s.class === cls && s.status === "active");
    let next = [...studentAttendance];
    activeInClass.forEach((st) => {
      const statusVal = draft[st.id] || "present";
      const idx = next.findIndex((a) => a.studentId === st.id && a.date === date);
      if (idx >= 0) {
        next[idx] = { ...next[idx], status: statusVal, ...stampEdit(currentUser) };
      } else {
        next.push({ id: uid("att"), studentId: st.id, date, status: statusVal, comments: "", ...stampNew(currentUser) });
      }
    });
    saveStudentAttendance(next);
  }
  function deleteStudentAttendanceEntry(id) { saveStudentAttendance(studentAttendance.filter((a) => a.id !== id)); }

  function upsertStaffAttendance(record, options) {
    const closeModal = !options || options.closeModal !== false;
    const prior = staffAttendance.find((a) => a.id === record.id);
    // Approval logic: an admin touching a record (creating or editing it
    // through the normal form/table) implies review, so it's auto-approved.
    // A non-admin creating/editing their own record (e.g. via the Duty tab)
    // stays "pending" until an admin reviews it — unless the caller passed
    // an explicit approvalStatus (used by the Approve/Reject buttons).
    const approvalStatus = record.approvalStatus !== undefined ? record.approvalStatus : (isAdmin ? "approved" : "pending");
    const merged = { ...record, approvalStatus };
    const final = prior ? { ...prior, ...merged, ...stampEdit(currentUser) } : { ...merged, ...stampNew(currentUser) };
    saveStaffAttendance(prior ? staffAttendance.map((a) => (a.id === final.id ? final : a)) : [...staffAttendance, final]);
    if (closeModal) setStaffAttModal(null);
  }
  function deleteStaffAttendanceEntry(id) { saveStaffAttendance(staffAttendance.filter((a) => a.id !== id)); }
  function setAttendanceApproval(record, approvalStatus) {
    upsertStaffAttendance({ ...record, approvalStatus }, { closeModal: false });
  }

  // Self-service "Duty" clock for a teacher/staff account linked to a staff record.
  function myTodayAttendance(staffId) {
    const today = todayISO();
    return staffAttendance.find((a) => a.staffId === staffId && a.date === today) || null;
  }
  function startDuty() {
    if (!currentUser.staffId) return;
    const existing = myTodayAttendance(currentUser.staffId);
    if (existing) return;
    const today = todayISO();
    const d = new Date();
    upsertStaffAttendance({
      id: uid("satt"), staffId: currentUser.staffId, date: today, month: MONTHS[d.getMonth()], year: d.getFullYear(),
      checkIn: nowTimeHHMM(), checkOut: "", absent: false, lwp: false, advance: 0, other: 0,
      regularHrs: 0, overtimeHrs: 0, approvalStatus: "pending", comments: "Self check-in via Duty tab",
    }, { closeModal: false });
  }
  function endDuty() {
    if (!currentUser.staffId) return;
    const existing = myTodayAttendance(currentUser.staffId);
    if (!existing || existing.checkOut) return;
    const checkOut = nowTimeHHMM();
    const totalHrs = computeHoursFromTimes(existing.checkIn, checkOut);
    const { regular, overtime } = splitRegularOvertime(totalHrs);
    upsertStaffAttendance({ ...existing, checkOut, regularHrs: regular, overtimeHrs: overtime, approvalStatus: "pending" }, { closeModal: false });
  }

  function generateMonthlyFees() {
    const month = MONTHS[new Date().getMonth()];
    const year = new Date().getFullYear();
    const existingKeys = new Set(fees.map((f) => f.studentId + "|" + f.month + "|" + f.year));
    const newRecords = [];
    students.forEach((st) => {
      if (st.status !== "active") return;
      const key = st.id + "|" + month + "|" + year;
      if (existingKeys.has(key)) return;
      if (st.feeWaived) {
        newRecords.push({
          id: uid("fee"), studentId: st.id, month, year,
          standardFee: Number(st.standardFee) || 0, discountType: "none", discountValue: 0, discount: 0, netFee: 0,
          papersFund: 0, papersFundEnabled: false, lateFee: 0, lateFeeEnabled: false,
          paidAmount: 0, status: "free", paidDate: "", comments: "Fee waived (financial assistance)",
          ...stampNew(currentUser),
        });
        return;
      }
      const { discount, netFee } = computeNetFee(st.standardFee, st.discountType, st.discountValue);
      newRecords.push({
        id: uid("fee"), studentId: st.id, month, year,
        standardFee: Number(st.standardFee) || 0, discountType: st.discountType || "none", discountValue: Number(st.discountValue) || 0,
        discount, netFee,
        papersFund: PAPERS_FUND_MONTHS.includes(month) ? Number(settings.papersFundAmount) || 0 : 0,
        papersFundEnabled: PAPERS_FUND_MONTHS.includes(month),
        lateFee: 0, lateFeeEnabled: false,
        paidAmount: 0, status: "unpaid", paidDate: "", comments: "",
        ...stampNew(currentUser),
      });
    });
    if (newRecords.length) saveFees([...fees, ...newRecords]);
    return newRecords.length;
  }

  function applyLateFeeToOverdue() {
    let count = 0;
    const next = fees.map((f) => {
      if (feeDueState(f) === "overdue" && (Number(f.lateFee) || 0) === 0) {
        count++;
        return { ...f, lateFee: Number(settings.lateFeeAmount) || 0, lateFeeEnabled: true, ...stampEdit(currentUser) };
      }
      return f;
    });
    if (count > 0) saveFees(next);
    return count;
  }

  const activeStaffCount = staff.filter((s) => s.status === "active").length;
  const activeStudentCount = students.filter((s) => s.status === "active").length;
  const currentMonth = MONTHS[new Date().getMonth()];
  const currentYear = new Date().getFullYear();
  const thisMonthFees = fees.filter((f) => f.month === currentMonth && f.year === currentYear);
  const collectedThisMonth = thisMonthFees.reduce((sum, f) => sum + (Number(f.paidAmount) || 0), 0);
  const pendingDues = fees.filter((f) => f.status !== "paid" && f.status !== "free").reduce((sum, f) => sum + feeBalance(f), 0);
  const studentsPaidThisMonth = new Set(thisMonthFees.filter((f) => f.status === "paid").map((f) => f.studentId)).size;
  const studentsDueThisMonth = new Set(thisMonthFees.filter((f) => f.status === "unpaid" || f.status === "partial").map((f) => f.studentId)).size;
  const isSameMonth = (dateStr) => {
    if (!dateStr) return false;
    const d = new Date(dateStr);
    return d.getMonth() === new Date().getMonth() && d.getFullYear() === new Date().getFullYear();
  };
  const itemsThisMonth = items.filter((i) => isSameMonth(i.issueDate));
  const uniformCollectedThisMonth = itemsThisMonth.filter((i) => i.item.startsWith("Uniform")).reduce((s, i) => s + (Number(i.paidAmount) || 0), 0);
  const bookCollectedThisMonth = itemsThisMonth.filter((i) => i.item === "Books set" || i.item === "Notebooks").reduce((s, i) => s + (Number(i.paidAmount) || 0), 0);
  const todayStr = todayISO();
  const staffAttToday = staffAttendance.filter((a) => a.date === todayStr);
  const staffPresentToday = staffAttToday.filter((a) => !a.absent).length;
  const staffAbsentToday = staffAttToday.filter((a) => a.absent).length;

  const studentMap = useMemo(() => Object.fromEntries(students.map((s) => [s.id, s])), [students]);

  const overdueFees = useMemo(() => fees.filter((f) => feeDueState(f) === "overdue"), [fees]);
  const overdueByStudent = useMemo(() => {
    const map = {};
    overdueFees.forEach((f) => {
      if (!map[f.studentId]) map[f.studentId] = { amount: 0, months: [] };
      map[f.studentId].amount += feeBalance(f);
      map[f.studentId].months.push(f.month + " " + f.year);
    });
    return map;
  }, [overdueFees]);
  const overdueStudentIds = Object.keys(overdueByStudent);
  const totalOverdueAmount = Object.values(overdueByStudent).reduce((s, v) => s + v.amount, 0);

  useEffect(() => {
    if (currentUser && !loading && !autoPopupShown && overdueStudentIds.length > 0) {
      setOverdueModal(true);
      setAutoPopupShown(true);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser, loading, overdueStudentIds.length]);

  const filteredStaff = staff.filter((s) => (s.name + s.role + s.subject).toLowerCase().includes(staffSearch.toLowerCase()));
  const filteredStudents = students.filter((s) => (s.name + s.class + s.rollNo).toLowerCase().includes(studentSearch.toLowerCase()));
  const filteredFees = fees.filter((f) => {
    if (feeMonthFilter !== "all" && f.month !== feeMonthFilter) return false;
    if (f.year !== Number(feeYearFilter)) return false;
    if (feeStatusFilter !== "all" && (feeStatusFilter === "free" ? f.status !== "free" : f.status !== feeStatusFilter)) return false;
    return true;
  });
  const filteredItems = items.filter((i) => {
    const st = studentMap[i.studentId];
    return ((st ? st.name : "") + i.item + i.description).toLowerCase().includes(itemSearch.toLowerCase());
  });
  const filteredSalaries = salaries.filter((s) => {
    if (salaryMonthFilter !== "all" && s.month !== salaryMonthFilter) return false;
    if (s.year !== Number(salaryYearFilter)) return false;
    if (salaryStatusFilter !== "all" && s.status !== salaryStatusFilter) return false;
    return true;
  });
  const staffMap = useMemo(() => Object.fromEntries(staff.map((s) => [s.id, s])), [staff]);
  const pettyForMonth = pettyCash
    .filter((p) => p.month === pettyMonthFilter && p.year === Number(pettyYearFilter))
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  let running = 0;
  const pettyForMonthWithBalance = pettyForMonth.map((p) => {
    running += p.type === "receipt" ? Number(p.amount) || 0 : -(Number(p.amount) || 0);
    return { ...p, running };
  });
  const pettyReceivedMonth = pettyForMonth.filter((p) => p.type === "receipt").reduce((s, p) => s + (Number(p.amount) || 0), 0);
  const pettySpentMonth = pettyForMonth.filter((p) => p.type === "expense").reduce((s, p) => s + (Number(p.amount) || 0), 0);

  const sortedLedger = [...ledgerEntries].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  let runDebit = 0, runCredit = 0;
  const ledgerWithRunning = sortedLedger.map((l) => {
    runDebit += Number(l.debit) || 0;
    runCredit += Number(l.credit) || 0;
    return { ...l, totalDebit: runDebit, totalCredit: runCredit };
  });
  const ledgerTotalDebit = ledgerEntries.reduce((s, l) => s + (Number(l.debit) || 0), 0);
  const ledgerTotalCredit = ledgerEntries.reduce((s, l) => s + (Number(l.credit) || 0), 0);
  const ledgerTotalNet = ledgerTotalCredit - ledgerTotalDebit;

  const classStudentsForAttendance = students.filter((s) => s.class === attClass && s.status === "active");
  const attendanceForDate = studentAttendance.filter((a) => a.date === attDate);
  const attendanceForDateByStudent = Object.fromEntries(attendanceForDate.map((a) => [a.studentId, a]));
  const attPresentCount = classStudentsForAttendance.filter((s) => (attendanceForDateByStudent[s.id]?.status || "present") === "present").length;
  const attAbsentCount = classStudentsForAttendance.filter((s) => attendanceForDateByStudent[s.id]?.status === "absent").length;
  const attLeaveCount = classStudentsForAttendance.filter((s) => attendanceForDateByStudent[s.id]?.status === "leave").length;

  const effectiveStaffAttId = staffAttStaffId || (staff[0] ? staff[0].id : "");
  const staffAttForStaff = staffAttendance
    .filter((a) => a.staffId === effectiveStaffAttId && a.month === staffAttMonthFilter && a.year === Number(staffAttYearFilter))
    .sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  const staffAttTotals = staffAttForStaff.reduce((acc, a) => {
    acc.regular += Number(a.regularHrs) || 0;
    acc.overtime += Number(a.overtimeHrs) || 0;
    acc.absent += a.absent ? 1 : 0;
    acc.lwp += a.lwp ? 1 : 0;
    acc.advance += Number(a.advance) || 0;
    acc.other += Number(a.other) || 0;
    return acc;
  }, { regular: 0, overtime: 0, absent: 0, lwp: 0, advance: 0, other: 0 });

  function downloadWorkbook(sheets, filename) {
    const wb = XLSX.utils.book_new();
    sheets.forEach(({ name, rows }) => {
      const ws = XLSX.utils.json_to_sheet(rows);
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    });
    XLSX.writeFile(wb, filename);
  }
  function exportStaff() {
    downloadWorkbook([{ name: "Staff", rows: staff.map((s) => ({
      Name: s.name, Role: s.role, Subject: s.subject, Phone: s.phone, Email: s.email,
      "Joining Date": s.joinDate, "Monthly Salary": s.salary, Status: s.status, "Inactive Date": s.inactiveDate || "", "Inactive Reason": s.inactiveReason || "",
    })) }], "staff-records.xlsx");
  }
  function exportStudents() {
    downloadWorkbook([{ name: "Admissions", rows: students.map((s) => {
      const { discount, netFee } = computeNetFee(s.standardFee, s.discountType, s.discountValue);
      return {
        Name: s.name, Class: s.class, Section: s.section, "Roll No": s.rollNo,
        "Admission Date": s.admissionDate, "Father's Name": s.fatherName || "", "Father's CNIC": s.fatherCnic || "",
        "Parent Name": s.parentName, "Parent Phone": s.parentPhone, "Emergency Contact": s.emergencyContact || "",
        "Registration Fee": s.registrationFee || 0, "Admission Fee": s.admissionFee || 0,
        "Monthly Fee": s.standardFee, "Discount Type": s.discountType, "Discount Value": s.discountValue,
        Discount: discount, "Net Monthly Fee": s.feeWaived ? 0 : netFee, "Fee Waived": s.feeWaived ? "Yes" : "No",
        Status: s.status, "Inactive Date": s.inactiveDate || "", "Inactive Reason": s.inactiveReason || "",
      };
    }) }], "admissions.xlsx");
  }
  function feeRowExport(f) {
    const st = studentMap[f.studentId];
    return {
      Student: st ? st.name : "(removed student)", Class: st ? st.class : "",
      Month: f.month, Year: f.year, "Standard Fee": f.standardFee, Discount: f.discount,
      "Net Fee": f.netFee, "Papers Fund": f.papersFund, "Late Fee": f.lateFee,
      "Total Due": feeTotalDue(f), "Amount Paid": f.paidAmount, Balance: feeBalance(f),
      Status: feeDueState(f), "Paid Date": f.paidDate || "", Comments: f.comments || "",
      "Entered By": f.enteredBy || "", "Entered At": formatDateTime(f.enteredAt),
      "Last Edited By": f.editedBy || "", "Last Edited At": formatDateTime(f.editedAt),
    };
  }
  function exportFees() { downloadWorkbook([{ name: "Fee Ledger", rows: fees.map(feeRowExport) }], "fee-ledger.xlsx"); }
  function itemRowExport(i) {
    const st = studentMap[i.studentId];
    return {
      Student: st ? st.name : "(removed student)", Class: st ? st.class : "", Item: i.item,
      Description: i.description, Quantity: i.quantity, "Unit Price": i.unitPrice,
      "Total Amount": itemTotal(i), "Amount Paid": i.paidAmount, Balance: itemBalance(i),
      Status: computeItemStatus(i), "Issue Date": i.issueDate, Comments: i.comments || "",
      "Entered By": i.enteredBy || "", "Entered At": formatDateTime(i.enteredAt),
    };
  }
  function exportItems() { downloadWorkbook([{ name: "Uniforms and Books", rows: items.map(itemRowExport) }], "uniforms-books.xlsx"); }
  function pettyRowExport(p) {
    return {
      Date: p.date, Month: p.month, Year: p.year, Type: p.type === "receipt" ? "Received" : "Expense",
      "Source / Particulars": p.label, Amount: p.amount, Comments: p.comments || "",
      "Entered By": p.enteredBy || "", "Entered At": formatDateTime(p.enteredAt),
    };
  }
  function exportPetty() { downloadWorkbook([{ name: "Petty Cash", rows: pettyCash.map(pettyRowExport) }], "petty-cash.xlsx"); }
  function salaryRowExport(s) {
    const st = staffMap[s.staffId];
    return {
      Staff: st ? st.name : "(removed staff)", Role: st ? st.role : "",
      Month: s.month, Year: s.year, "Base Salary": s.baseSalary, Bonus: s.bonus, Deduction: s.deduction,
      "Net Salary": salaryNet(s), "Amount Paid": s.paidAmount, Balance: salaryBalance(s),
      Status: salaryDueState(s), "Paid Date": s.paidDate || "", Comments: s.comments || "",
      "Entered By": s.enteredBy || "", "Entered At": formatDateTime(s.enteredAt),
      "Last Edited By": s.editedBy || "", "Last Edited At": formatDateTime(s.editedAt),
    };
  }
  function exportSalaries() { downloadWorkbook([{ name: "Staff Salary", rows: salaries.map(salaryRowExport) }], "staff-salary.xlsx"); }

  function ledgerRowExport(l) {
    return {
      Date: l.date, Description: l.description, "Post Reference": l.postReference,
      Debit: l.debit || 0, Credit: l.credit || 0, Comments: l.comments || "",
      "Entered By": l.enteredBy || "", "Entered At": formatDateTime(l.enteredAt),
      "Last Edited By": l.editedBy || "", "Last Edited At": formatDateTime(l.editedAt),
    };
  }
  function exportLedger() { downloadWorkbook([{ name: "Ledger Report", rows: ledgerWithRunning.map(ledgerRowExport) }], "ledger-report.xlsx"); }

  function studentAttRowExport(a) {
    const st = studentMap[a.studentId];
    return {
      Date: a.date, Student: st ? st.name : "(removed student)", Class: st ? (st.class + (st.section ? "-" + st.section : "")) : "",
      Status: a.status, Comments: a.comments || "", "Entered By": a.enteredBy || "", "Entered At": formatDateTime(a.enteredAt),
    };
  }
  function exportStudentAttendance() { downloadWorkbook([{ name: "Student Attendance", rows: studentAttendance.map(studentAttRowExport) }], "student-attendance.xlsx"); }

  function staffAttRowExport(a) {
    const st = staffMap[a.staffId];
    return {
      Date: a.date, Staff: st ? st.name : "(removed staff)", "Check In": a.checkIn || "", "Check Out": a.checkOut || "",
      "Regular Hrs": a.regularHrs || 0, "Overtime Hrs": a.overtimeHrs || 0, Absent: a.absent ? "Yes" : "No", LWP: a.lwp ? "Yes" : "No",
      Advance: a.advance || 0, Other: a.other || 0, "Total Hours": (Number(a.regularHrs) || 0) + (Number(a.overtimeHrs) || 0),
      Approval: a.approvalStatus || "approved", Comments: a.comments || "", "Entered By": a.enteredBy || "", "Entered At": formatDateTime(a.enteredAt),
    };
  }
  function exportStaffAttendance() { downloadWorkbook([{ name: "Staff Attendance", rows: staffAttendance.map(staffAttRowExport) }], "staff-attendance.xlsx"); }

  function exportFullReport() {
    const overdueRows = overdueStudentIds.map((sid) => {
      const st = studentMap[sid];
      return {
        Student: st ? st.name : "(removed student)", Class: st ? st.class : "",
        "Parent Phone": st ? st.parentPhone : "",
        "Overdue Months": overdueByStudent[sid].months.join(", "),
        "Overdue Amount": overdueByStudent[sid].amount,
      };
    });
    downloadWorkbook([
      { name: "Staff", rows: staff.map((s) => ({ Name: s.name, Role: s.role, Subject: s.subject, Phone: s.phone, Email: s.email, "Joining Date": s.joinDate, "Monthly Salary": s.salary, Status: s.status, "Inactive Date": s.inactiveDate || "", "Inactive Reason": s.inactiveReason || "" })) },
      { name: "Admissions", rows: students.map((s) => {
          const { discount, netFee } = computeNetFee(s.standardFee, s.discountType, s.discountValue);
          return { Name: s.name, Class: s.class, Section: s.section, "Roll No": s.rollNo, "Admission Date": s.admissionDate, "Father's Name": s.fatherName || "", "Father's CNIC": s.fatherCnic || "", "Parent Name": s.parentName, "Parent Phone": s.parentPhone, "Emergency Contact": s.emergencyContact || "", "Registration Fee": s.registrationFee || 0, "Admission Fee": s.admissionFee || 0, "Monthly Fee": s.standardFee, Discount: discount, "Net Monthly Fee": s.feeWaived ? 0 : netFee, "Fee Waived": s.feeWaived ? "Yes" : "No", Status: s.status, "Inactive Date": s.inactiveDate || "", "Inactive Reason": s.inactiveReason || "" };
        }) },
      { name: "Fee Ledger", rows: fees.map(feeRowExport) },
      { name: "Uniforms and Books", rows: items.map(itemRowExport) },
      { name: "Petty Cash", rows: pettyCash.map(pettyRowExport) },
      { name: "Staff Salary", rows: salaries.map(salaryRowExport) },
      { name: "Ledger Report", rows: ledgerWithRunning.map(ledgerRowExport) },
      { name: "Student Attendance", rows: studentAttendance.map(studentAttRowExport) },
      { name: "Staff Attendance", rows: staffAttendance.map(staffAttRowExport) },
      { name: "Overdue Summary", rows: overdueRows },
    ], "habib-shining-star-school-report.xlsx");
  }

  if (loading) {
    return (<div className="sms-root"><style>{STYLES}</style><div className="sms-loading">Loading registry records…</div></div>);
  }
  if (!currentUser) {
    return <LoginScreen settings={settings} onLogin={handleLogin} language={language} changeLanguage={changeLanguage} t={t} />;
  }

  return (
    <div className={"sms-root" + (language === "ur" ? " rtl" : "")} dir={language === "ur" ? "rtl" : "ltr"}>
      <style>{STYLES}</style>
      {toast && <div className="sms-toast">{toast}</div>}

      <div className="sms-mobile-topbar">
        <button className="sms-hamburger-btn" onClick={() => setMobileNavOpen(true)} aria-label="Open menu">☰</button>
        <div>
          <div className="title sms-serif">{SCHOOL_NAME}</div>
          <div className="sub">{t(ALL_TABS.find((tb) => tb.key === page)?.labelKey) || ""}</div>
        </div>
      </div>
      <div className={"sms-sidebar-overlay" + (mobileNavOpen ? " open" : "")} onClick={() => setMobileNavOpen(false)} />

      <aside className={"sms-sidebar" + (mobileNavOpen ? " open" : "")}>
        <div className="sms-brand">
          <div className="sms-brand-title sms-serif">{SCHOOL_NAME}</div>
          <div className="sms-brand-sub">{SCHOOL_ADDRESS}</div>
        </div>
        <nav className="sms-tabs">
          {ALL_TABS.filter((tb) => {
            if (tb.requiresStaffLink) return !!currentUser.staffId;
            if (tb.adminOnly) return isAdmin;
            return isAdmin || hasTabAccess(tb.key);
          }).map((tb) => (
            <div key={tb.key} className={"sms-tab" + (page === tb.key ? " active" : "")} onClick={() => { setPage(tb.key); setMobileNavOpen(false); }}>
              <span className="sms-tab-index">{tb.i}</span>
              <span>{t(tb.labelKey)}</span>
              {tb.key === "fees" && overdueStudentIds.length > 0 && <span className="sms-tab-badge">{overdueStudentIds.length}</span>}
            </div>
          ))}
        </nav>
        <div className="sms-sidebar-foot">
          <div className="sms-lang-switch" style={{ marginBottom: 10 }}>
            <button type="button" className={"sms-lang-btn" + (language === "en" ? " active" : "")} onClick={() => changeLanguage("en")}>English</button>
            <button type="button" className={"sms-lang-btn" + (language === "ur" ? " active" : "")} onClick={() => changeLanguage("ur")}>اردو</button>
          </div>
          <div className="sms-user-chip">{currentUser.name}<span className="role">{currentUser.role === "admin" ? t("administrator") : t("staff")}</span></div>
          <div className={"sms-sync-badge" + (typeof window !== "undefined" && window.__HSS_STORAGE_MODE__ === "cloud" ? " cloud" : "")}>
            {typeof window !== "undefined" && window.__HSS_STORAGE_MODE__ === "cloud" ? t("syncedCloud") : t("syncedLocal")}
          </div>
          {isAdmin && <button className="sms-sidebar-link" onClick={() => { setSettingsModal(true); setMobileNavOpen(false); }}>{t("settingsAccess")}</button>}
          <button className="sms-sidebar-link" onClick={handleLogout}>{t("logOut")}</button>
        </div>
      </aside>

      <main className="sms-main">
        {page === "dashboard" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">{t("dashboard")}</h1>
              <div className="sms-header-actions">
                <span className="sms-datestamp">{todayISO()}</span>
                <button className="sms-btn gold" onClick={exportFullReport}>{t("downloadReport")}</button>
              </div>
            </div>
            <div className="sms-content">
              {overdueStudentIds.length > 0 && !bannerDismissed && (
                <div className="sms-alert-banner">
                  <span>⚠ {overdueStudentIds.length} student{overdueStudentIds.length > 1 ? "s" : ""} have unpaid fees from a previous month — total outstanding {currency(totalOverdueAmount)}.</span>
                  <div className="sms-alert-actions">
                    <button className="sms-btn small" onClick={() => setOverdueModal(true)}>View list</button>
                    <button className="sms-btn secondary small" onClick={() => setBannerDismissed(true)}>Dismiss</button>
                  </div>
                </div>
              )}
              <div className="sms-cards-row">
                <div className="sms-card"><div className="sms-card-label">{t("activeStaff")}</div><div className="sms-card-value">{activeStaffCount}</div></div>
                <div className="sms-card"><div className="sms-card-label">{t("activeStudents")}</div><div className="sms-card-value">{activeStudentCount}</div></div>
                <div className="sms-card"><div className="sms-card-label">{t("collected")} — {currentMonth}</div><div className="sms-card-value green">{currency(collectedThisMonth)}</div></div>
                <div className={"sms-card" + (overdueStudentIds.length > 0 ? " warn" : "")}>
                  <div className="sms-card-label">{overdueStudentIds.length > 0 ? t("overdueStudents") : t("totalDuesPending")}</div>
                  <div className="sms-card-value rust">{overdueStudentIds.length > 0 ? overdueStudentIds.length : currency(pendingDues)}</div>
                </div>
              </div>
              {PAPERS_FUND_MONTHS.includes(currentMonth) && (
                <div className="sms-alert-banner" style={{ background: "var(--amber-bg)", border: "1px solid var(--amber)", color: "var(--amber)" }}>
                  <span>📄 {currentMonth} is a papers fund month — {currency(settings.papersFundAmount)} is suggested when generating fees. Papers fund and late fee are optional per student and can be waived for financial assistance.</span>
                </div>
              )}

              <div className="sms-section-title">{t("thisMonthGlance")} <span className="sms-tag">{currentMonth} {currentYear}</span></div>
              <div className="sms-cards-row">
                <div className="sms-card"><div className="sms-card-label">{t("studentsPaidFee")}</div><div className="sms-card-value green">{studentsPaidThisMonth}</div></div>
                <div className="sms-card"><div className="sms-card-label">{t("studentsDue")}</div><div className="sms-card-value rust">{studentsDueThisMonth}</div></div>
                <div className="sms-card"><div className="sms-card-label">{t("uniformFundCollected")}</div><div className="sms-card-value green">{currency(uniformCollectedThisMonth)}</div></div>
                <div className="sms-card"><div className="sms-card-label">{t("booksFundCollected")}</div><div className="sms-card-value green">{currency(bookCollectedThisMonth)}</div></div>
              </div>

              <div className="sms-section-title">{t("staffAttendanceToday")} <span className="sms-tag">{t("today")} · {todayStr}</span></div>
              <div className="sms-cards-row three">
                <div className="sms-card"><div className="sms-card-label">{t("presentToday")}</div><div className="sms-card-value green">{staffPresentToday}</div></div>
                <div className="sms-card"><div className="sms-card-label">{t("absentToday")}</div><div className="sms-card-value rust">{staffAbsentToday}</div></div>
                <div className="sms-card"><div className="sms-card-label">{t("notMarkedYet")}</div><div className="sms-card-value">{Math.max(activeStaffCount - staffAttToday.length, 0)}</div></div>
              </div>

              <div className="sms-section-title">{t("recentAdmissions")} <span className="sms-tag">last 5</span></div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Name</th><th>Class</th><th>Roll no.</th><th>Admission date</th><th>Status</th></tr></thead>
                  <tbody>
                    {students.slice(-5).reverse().map((s) => (
                      <tr key={s.id}>
                        <td>{overdueByStudent[s.id] ? <span className="sms-name-flag"><span className="sms-flag-dot" />{s.name}</span> : s.name}</td>
                        <td>{s.class}{s.section ? "-" + s.section : ""}</td>
                        <td className="sms-mono">{s.rollNo}</td><td>{s.admissionDate}</td>
                        <td><span className={"sms-pill " + s.status}>{s.status}</span></td>
                      </tr>
                    ))}
                    {students.length === 0 && <tr><td colSpan="5" className="sms-empty">No admissions recorded yet.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "staff" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Staff records</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{staff.length} total entries</span><button className="sms-btn secondary" onClick={exportStaff}>⬇ Export Excel</button></div>
            </div>
            <div className="sms-content">
              <div className="sms-toolbar">
                <input className="sms-input" placeholder="Search by name, role, subject" value={staffSearch} onChange={(e) => setStaffSearch(e.target.value)} style={{ width: 260 }} />
                <button className="sms-btn" onClick={() => setStaffModal({})}>+ Add staff</button>
              </div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Name</th><th>Role</th><th>Subject</th><th>Phone</th><th>Joining date</th><th>Salary</th><th>Status</th><th></th></tr></thead>
                  <tbody>
                    {filteredStaff.map((s) => (
                      <tr key={s.id}>
                        <td>{s.name}</td><td>{s.role}</td><td>{s.subject}</td><td className="sms-mono">{s.phone}</td>
                        <td>{s.joinDate}</td><td className="sms-mono">{currency(s.salary)}</td>
                        <td><span className={"sms-pill " + s.status}>{s.status}</span>{s.status === "inactive" && (s.inactiveReason || s.inactiveDate) && <span className="sms-subtext">{s.inactiveDate ? `Since ${s.inactiveDate}. ` : ""}{s.inactiveReason}</span>}</td>
                        <td style={{ display: "flex", gap: 6 }}>
                          {isAdmin ? (<>
                            <button className="sms-btn secondary small" onClick={() => setStaffModal(s)}>Edit</button>
                            <button className="sms-btn danger small" onClick={() => deleteStaff(s.id)}>Delete</button>
                          </>) : <span className="sms-locked">Admin only</span>}
                        </td>
                      </tr>
                    ))}
                    {filteredStaff.length === 0 && <tr><td colSpan="8" className="sms-empty">No staff records match. Add your first staff member.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "admissions" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Admissions</h1>
              <div className="sms-header-actions sms-no-print">
                <span className="sms-datestamp">{students.length} total students</span>
                <button className="sms-btn secondary" onClick={() => setShowAdmissionLedger((v) => !v)}>{showAdmissionLedger ? "← Back to table" : "🖨 Ledger report"}</button>
                <button className="sms-btn secondary" onClick={exportStudents}>⬇ Export Excel</button>
              </div>
            </div>
            <div className="sms-content">
              {showAdmissionLedger ? (
                <>
                  <div className="sms-ledger-report-head">
                    <div className="school sms-serif">{SCHOOL_NAME}</div>
                    <div className="addr">{SCHOOL_ADDRESS}</div>
                    <div className="meta">Admission Ledger Report · Generated {formatDateTime(new Date().toISOString())} · {students.length} total students</div>
                  </div>
                  <div className="sms-toolbar sms-no-print">
                    <button className="sms-btn gold" onClick={() => window.print()}>🖨 Print this report</button>
                  </div>
                  <div className="sms-ledger-page">
                    <table className="sms-ledger-table">
                      <thead>
                        <tr>
                          <th>S.No</th><th>Name</th><th>Father's name</th><th>Father's CNIC</th><th>Class</th><th>Roll no.</th>
                          <th>Admission date</th><th>Parent phone</th><th>Emergency contact</th>
                          <th>Registration fee</th><th>Admission fee</th><th>Monthly fee</th><th>Discount</th><th>Net fee</th><th>Status</th>
                        </tr>
                      </thead>
                      <tbody>
                        {students.map((s, idx) => {
                          const { discount, netFee } = computeNetFee(s.standardFee, s.discountType, s.discountValue);
                          return (
                            <tr key={s.id}>
                              <td>{idx + 1}</td>
                              <td>{s.name}</td>
                              <td>{s.fatherName || "—"}</td>
                              <td>{s.fatherCnic || "—"}</td>
                              <td>{s.class}{s.section ? "-" + s.section : ""}</td>
                              <td>{s.rollNo}</td>
                              <td>{s.admissionDate}</td>
                              <td>{s.parentPhone || "—"}</td>
                              <td>{s.emergencyContact || "—"}</td>
                              <td>{currency(s.registrationFee || 0)}</td>
                              <td>{currency(s.admissionFee || 0)}</td>
                              <td>{currency(s.standardFee)}</td>
                              <td>{discount > 0 ? currency(discount) : "—"}</td>
                              <td>{s.feeWaived ? "Exempt" : currency(netFee)}</td>
                              <td>{s.status}{s.status === "inactive" ? ` — ${s.inactiveDate ? "since " + s.inactiveDate + ". " : ""}${s.inactiveReason || ""}` : ""}</td>
                            </tr>
                          );
                        })}
                        {students.length === 0 && <tr><td colSpan="15" className="sms-empty">No admission records yet.</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </>
              ) : (
                <>
                  <div className="sms-toolbar">
                    <input className="sms-input" placeholder="Search by name, class, roll no." value={studentSearch} onChange={(e) => setStudentSearch(e.target.value)} style={{ width: 260 }} />
                    <button className="sms-btn" onClick={() => setStudentModal({})}>+ New admission</button>
                  </div>
                  <div className="sms-ledger-page">
                    <table className="sms-table">
                      <thead><tr><th>Name</th><th>Class</th><th>Roll no.</th><th>Father's name</th><th>Parent phone</th><th>Registration + Admission</th><th>Monthly fee</th><th>Net fee</th><th>Status</th><th></th></tr></thead>
                      <tbody>
                        {filteredStudents.map((s) => {
                          const { discount, netFee } = computeNetFee(s.standardFee, s.discountType, s.discountValue);
                          const oneTime = (Number(s.registrationFee) || 0) + (Number(s.admissionFee) || 0);
                          return (
                            <tr key={s.id} className={overdueByStudent[s.id] ? "sms-row-overdue" : ""}>
                              <td>{overdueByStudent[s.id] ? <span className="sms-name-flag"><span className="sms-flag-dot" />{s.name}</span> : s.name}</td>
                              <td>{s.class}{s.section ? "-" + s.section : ""}</td>
                              <td className="sms-mono">{s.rollNo}</td>
                              <td>{s.fatherName || "—"}{s.fatherCnic && <span className="sms-subtext">{s.fatherCnic}</span>}</td>
                              <td className="sms-mono">{s.parentPhone}{s.emergencyContact && <span className="sms-subtext">Emerg: {s.emergencyContact}</span>}</td>
                              <td className="sms-mono">{oneTime > 0 ? currency(oneTime) : "—"}</td>
                              <td className="sms-mono">{currency(s.standardFee)}</td>
                              <td className="sms-mono">
                                {s.feeWaived ? <span className="sms-pill free">exempt</span> : currency(netFee)}
                                {!s.feeWaived && discount > 0 && <span className="sms-subtext">discount {currency(discount)}</span>}
                              </td>
                              <td><span className={"sms-pill " + s.status}>{s.status}</span>{s.status === "inactive" && (s.inactiveReason || s.inactiveDate) && <span className="sms-subtext">{s.inactiveDate ? `Since ${s.inactiveDate}. ` : ""}{s.inactiveReason}</span>}</td>
                              <td style={{ display: "flex", gap: 6 }}>
                                {isAdmin ? (<>
                                  <button className="sms-btn secondary small" onClick={() => setStudentModal(s)}>Edit</button>
                                  <button className="sms-btn danger small" onClick={() => deleteStudent(s.id)}>Delete</button>
                                </>) : <span className="sms-locked">Admin only</span>}
                              </td>
                            </tr>
                          );
                        })}
                        {filteredStudents.length === 0 && <tr><td colSpan="10" className="sms-empty">No admissions match. Record your first admission.</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </>
              )}
            </div>
          </>
        )}

        {page === "fees" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Fee ledger</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{fees.length} records</span><button className="sms-btn secondary" onClick={exportFees}>⬇ Export Excel (all data)</button></div>
            </div>
            <div className="sms-content">
              {overdueStudentIds.length > 0 && (
                <div className="sms-alert-banner">
                  <span>⚠ {overdueStudentIds.length} student{overdueStudentIds.length > 1 ? "s" : ""} overdue — rows highlighted in red below.</span>
                  <div className="sms-alert-actions"><button className="sms-btn small" onClick={() => setOverdueModal(true)}>View list</button></div>
                </div>
              )}
              <div className="sms-toolbar">
                <select className="sms-select" value={feeMonthFilter} onChange={(e) => setFeeMonthFilter(e.target.value)}>
                  <option value="all">All months</option>
                  {MONTHS.map((m) => <option key={m} value={m}>{m}{PAPERS_FUND_MONTHS.includes(m) ? " (papers fund)" : ""}</option>)}
                </select>
                <input className="sms-input" type="number" style={{ width: 90 }} value={feeYearFilter} onChange={(e) => setFeeYearFilter(e.target.value)} />
                <select className="sms-select" value={feeStatusFilter} onChange={(e) => setFeeStatusFilter(e.target.value)}>
                  <option value="all">All statuses</option>
                  <option value="paid">Paid</option>
                  <option value="partial">Partial</option>
                  <option value="unpaid">Unpaid</option>
                  <option value="free">Free / exempt</option>
                </select>
                <button className="sms-btn secondary" onClick={() => {
                  const n = generateMonthlyFees();
                  alert(n > 0 ? `Generated ${n} fee record(s) for ${currentMonth}.` : "All active students already have a record for this month.");
                }}>Generate this month's fees</button>
                {isAdmin && <button className="sms-btn secondary" onClick={() => {
                  const n = applyLateFeeToOverdue();
                  alert(n > 0 ? `Late fee of ${currency(settings.lateFeeAmount)} applied to ${n} overdue record(s).` : "No overdue records needed a late fee.");
                }}>Apply late fee to overdue</button>}
                <button className="sms-btn" onClick={() => setFeeModal({})} disabled={students.length === 0}>+ Add fee record</button>
              </div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Student</th><th>Month</th><th>Net fee</th><th>Papers fund</th><th>Late fee</th><th>Total due</th><th>Paid</th><th>Balance</th><th>Status</th><th>Entered by</th><th></th></tr></thead>
                  <tbody>
                    {filteredFees.map((f) => {
                      const st = studentMap[f.studentId];
                      const dueState = feeDueState(f);
                      const rowClass = dueState === "overdue" ? "sms-row-overdue" : dueState === "due" ? "sms-row-due" : dueState === "free" ? "sms-row-free" : "";
                      return (
                        <tr key={f.id} className={rowClass}>
                          <td>{st ? (dueState === "overdue" ? <span className="sms-name-flag"><span className="sms-flag-dot" />{st.name}</span> : st.name) : "(removed student)"}
                            {st && <span className="sms-subtext">{st.class}{st.section ? "-" + st.section : ""}</span>}
                          </td>
                          <td>{f.month} {f.year}</td>
                          <td className="sms-mono">{f.status === "free" ? "—" : currency(f.netFee)}{f.discount > 0 && f.status !== "free" && <span className="sms-subtext">std {currency(f.standardFee)}</span>}</td>
                          <td className="sms-mono">{f.papersFund > 0 ? <span className="sms-pill fund">{currency(f.papersFund)}</span> : "—"}</td>
                          <td className="sms-mono">{f.lateFee > 0 ? currency(f.lateFee) : "—"}</td>
                          <td className="sms-mono">{currency(feeTotalDue(f))}</td>
                          <td className="sms-mono">{currency(f.paidAmount)}</td>
                          <td className="sms-mono">{currency(feeBalance(f))}</td>
                          <td><span className={"sms-pill " + dueState}>{dueState}</span></td>
                          <td className="sms-subtext" style={{ display: "block", maxWidth: 120 }}>{f.enteredBy || "—"}<br />{formatDateTime(f.enteredAt)}</td>
                          <td style={{ display: "flex", gap: 6 }}>
                            {isAdmin ? (<>
                              <button className="sms-btn secondary small" onClick={() => setFeeModal(f)}>Update</button>
                              <button className="sms-btn danger small" onClick={() => deleteFee(f.id)}>Delete</button>
                            </>) : <span className="sms-locked">Admin only</span>}
                          </td>
                        </tr>
                      );
                    })}
                    {filteredFees.length === 0 && <tr><td colSpan="11" className="sms-empty">No fee records for this filter.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "items" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Uniforms &amp; books</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{items.length} records</span><button className="sms-btn secondary" onClick={exportItems}>⬇ Export Excel (all data)</button></div>
            </div>
            <div className="sms-content">
              <div className="sms-toolbar">
                <input className="sms-input" placeholder="Search by student or item" value={itemSearch} onChange={(e) => setItemSearch(e.target.value)} style={{ width: 260 }} />
                <button className="sms-btn" onClick={() => setItemModal({})} disabled={students.length === 0}>+ Add issue record</button>
              </div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Student</th><th>Item</th><th>Qty</th><th>Unit price</th><th>Total</th><th>Paid</th><th>Balance</th><th>Status</th><th>Entered by</th><th></th></tr></thead>
                  <tbody>
                    {filteredItems.map((i) => {
                      const st = studentMap[i.studentId];
                      const status = computeItemStatus(i);
                      const rowClass = status !== "paid" ? "sms-row-due" : "";
                      return (
                        <tr key={i.id} className={rowClass}>
                          <td>{st ? st.name : "(removed student)"}{st && <span className="sms-subtext">{st.class}{st.section ? "-" + st.section : ""}</span>}</td>
                          <td>{i.item}{i.description ? <span className="sms-subtext">{i.description}</span> : null}</td>
                          <td className="sms-mono">{i.quantity}</td>
                          <td className="sms-mono">{currency(i.unitPrice)}</td>
                          <td className="sms-mono">{currency(itemTotal(i))}</td>
                          <td className="sms-mono">{currency(i.paidAmount)}</td>
                          <td className="sms-mono">{currency(itemBalance(i))}</td>
                          <td><span className={"sms-pill " + status}>{status}</span></td>
                          <td className="sms-subtext" style={{ display: "block", maxWidth: 120 }}>{i.enteredBy || "—"}<br />{formatDateTime(i.enteredAt)}</td>
                          <td style={{ display: "flex", gap: 6 }}>
                            {isAdmin ? (<>
                              <button className="sms-btn secondary small" onClick={() => setItemModal(i)}>Update</button>
                              <button className="sms-btn danger small" onClick={() => deleteItem(i.id)}>Delete</button>
                            </>) : <span className="sms-locked">Admin only</span>}
                          </td>
                        </tr>
                      );
                    })}
                    {filteredItems.length === 0 && <tr><td colSpan="10" className="sms-empty">No uniform or book records yet.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "pettycash" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Petty cash</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{pettyCash.length} records</span><button className="sms-btn secondary" onClick={exportPetty}>⬇ Export Excel (all data)</button></div>
            </div>
            <div className="sms-content">
              <div className="sms-toolbar">
                <select className="sms-select" value={pettyMonthFilter} onChange={(e) => setPettyMonthFilter(e.target.value)}>
                  {MONTHS.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
                <input className="sms-input" type="number" style={{ width: 90 }} value={pettyYearFilter} onChange={(e) => setPettyYearFilter(e.target.value)} />
                <button className="sms-btn" onClick={() => setPettyModal({})}>+ Add entry</button>
              </div>
              <div className="sms-cards-row three">
                <div className="sms-card"><div className="sms-card-label">Received — {pettyMonthFilter}</div><div className="sms-card-value green">{currency(pettyReceivedMonth)}</div></div>
                <div className="sms-card"><div className="sms-card-label">Spent — {pettyMonthFilter}</div><div className="sms-card-value rust">{currency(pettySpentMonth)}</div></div>
                <div className="sms-card"><div className="sms-card-label">Balance — {pettyMonthFilter}</div><div className="sms-card-value">{currency(pettyReceivedMonth - pettySpentMonth)}</div></div>
              </div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Date</th><th>Type</th><th>Source / particulars</th><th>Amount</th><th>Running balance</th><th>Comments</th><th>Entered by</th><th></th></tr></thead>
                  <tbody>
                    {pettyForMonthWithBalance.map((p) => (
                      <tr key={p.id}>
                        <td>{p.date}</td>
                        <td><span className={"sms-pill " + p.type}>{p.type === "receipt" ? "Received" : "Expense"}</span></td>
                        <td>{p.label}</td>
                        <td className="sms-mono">{p.type === "receipt" ? "+" : "-"}{currency(p.amount)}</td>
                        <td className="sms-mono">{currency(p.running)}</td>
                        <td className="sms-subtext" style={{ display: "block", maxWidth: 160 }}>{p.comments || "—"}</td>
                        <td className="sms-subtext" style={{ display: "block", maxWidth: 120 }}>{p.enteredBy || "—"}<br />{formatDateTime(p.enteredAt)}</td>
                        <td style={{ display: "flex", gap: 6 }}>
                          {isAdmin ? (<>
                            <button className="sms-btn secondary small" onClick={() => setPettyModal(p)}>Update</button>
                            <button className="sms-btn danger small" onClick={() => deletePetty(p.id)}>Delete</button>
                          </>) : <span className="sms-locked">Admin only</span>}
                        </td>
                      </tr>
                    ))}
                    {pettyForMonthWithBalance.length === 0 && <tr><td colSpan="8" className="sms-empty">No petty cash entries for {pettyMonthFilter} {pettyYearFilter}.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "salary" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Staff salary</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{salaries.length} records</span><button className="sms-btn secondary" onClick={exportSalaries}>⬇ Export Excel (all data)</button></div>
            </div>
            <div className="sms-content">
              <div className="sms-toolbar">
                <select className="sms-select" value={salaryMonthFilter} onChange={(e) => setSalaryMonthFilter(e.target.value)}>
                  <option value="all">All months</option>
                  {MONTHS.map((m) => <option key={m} value={m}>{m}</option>)}
                </select>
                <input className="sms-input" type="number" style={{ width: 90 }} value={salaryYearFilter} onChange={(e) => setSalaryYearFilter(e.target.value)} />
                <select className="sms-select" value={salaryStatusFilter} onChange={(e) => setSalaryStatusFilter(e.target.value)}>
                  <option value="all">All statuses</option>
                  <option value="paid">Paid</option>
                  <option value="partial">Partial</option>
                  <option value="unpaid">Unpaid</option>
                </select>
                <button className="sms-btn secondary" onClick={() => {
                  const n = generateMonthlySalaries();
                  alert(n > 0 ? `Generated ${n} salary record(s) for ${currentMonth}.` : "All active staff already have a record for this month.");
                }}>Generate this month's salaries</button>
                <button className="sms-btn" onClick={() => setSalaryModal({})} disabled={staff.length === 0}>+ Add salary record</button>
              </div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Staff</th><th>Month</th><th>Base salary</th><th>Bonus</th><th>Deduction</th><th>Net salary</th><th>Paid</th><th>Balance</th><th>Status</th><th>Entered by</th><th></th></tr></thead>
                  <tbody>
                    {filteredSalaries.map((s) => {
                      const st = staffMap[s.staffId];
                      const dueState = salaryDueState(s);
                      const rowClass = dueState === "overdue" ? "sms-row-overdue" : dueState === "due" ? "sms-row-due" : "";
                      return (
                        <tr key={s.id} className={rowClass}>
                          <td>{st ? st.name : "(removed staff)"}{st && <span className="sms-subtext">{st.role}</span>}</td>
                          <td>{s.month} {s.year}</td>
                          <td className="sms-mono">{currency(s.baseSalary)}</td>
                          <td className="sms-mono">{s.bonus > 0 ? currency(s.bonus) : "—"}</td>
                          <td className="sms-mono">{s.deduction > 0 ? currency(s.deduction) : "—"}</td>
                          <td className="sms-mono">{currency(salaryNet(s))}</td>
                          <td className="sms-mono">{currency(s.paidAmount)}</td>
                          <td className="sms-mono">{currency(salaryBalance(s))}</td>
                          <td><span className={"sms-pill " + dueState}>{dueState}</span></td>
                          <td className="sms-subtext" style={{ display: "block", maxWidth: 120 }}>{s.enteredBy || "—"}<br />{formatDateTime(s.enteredAt)}</td>
                          <td style={{ display: "flex", gap: 6 }}>
                            {isAdmin ? (<>
                              <button className="sms-btn secondary small" onClick={() => setSalaryModal(s)}>Update</button>
                              <button className="sms-btn danger small" onClick={() => deleteSalary(s.id)}>Delete</button>
                            </>) : <span className="sms-locked">Admin only</span>}
                          </td>
                        </tr>
                      );
                    })}
                    {filteredSalaries.length === 0 && <tr><td colSpan="11" className="sms-empty">No salary records for this filter.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "ledger" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Ledger report</h1>
              <div className="sms-header-actions sms-no-print">
                <span className="sms-datestamp">{ledgerEntries.length} entries</span>
                <button className="sms-btn secondary" onClick={() => window.print()}>🖨 Print</button>
                <button className="sms-btn secondary" onClick={exportLedger}>⬇ Export Excel (all data)</button>
              </div>
            </div>
            <div className="sms-content">
              <div className="sms-ledger-report-head">
                <div style={{ display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 16 }}>
                  <div>
                    <div className="school sms-serif">{SCHOOL_NAME}</div>
                    <div className="addr">{SCHOOL_ADDRESS}</div>
                  </div>
                  <div className="sms-no-print" style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
                    <Field label="Accountant name">
                      <input className="sms-input" value={settings.ledgerAccountantName} onChange={(e) => saveSettings({ ...settings, ledgerAccountantName: e.target.value })} placeholder="e.g. Muhammad Nauman Habib" />
                    </Field>
                    <Field label="Time period">
                      <input className="sms-input" value={settings.ledgerTimePeriod} onChange={(e) => saveSettings({ ...settings, ledgerTimePeriod: e.target.value })} placeholder="e.g. 2025 - 2026" />
                    </Field>
                  </div>
                </div>
                <div className="meta">
                  Accountant: {settings.ledgerAccountantName || "—"} · Period: {settings.ledgerTimePeriod || "—"} · Generated {formatDateTime(new Date().toISOString())}
                </div>
              </div>

              <div className="sms-cards-row three">
                <div className="sms-card warn"><div className="sms-card-label">Total debit</div><div className="sms-card-value rust">{currency(ledgerTotalDebit)}</div></div>
                <div className="sms-card"><div className="sms-card-label">Total credit</div><div className="sms-card-value green">{currency(ledgerTotalCredit)}</div></div>
                <div className={"sms-card" + (ledgerTotalNet < 0 ? " warn" : "")}><div className="sms-card-label">Total net</div><div className={"sms-card-value" + (ledgerTotalNet < 0 ? " rust" : " green")}>{ledgerTotalNet < 0 ? `(${currency(Math.abs(ledgerTotalNet))})` : currency(ledgerTotalNet)}</div></div>
              </div>

              <div className="sms-toolbar sms-no-print">
                <button className="sms-btn" onClick={() => setLedgerModal({})}>+ Add ledger entry</button>
              </div>

              <div className="sms-ledger-page">
                <table className="sms-ledger-table">
                  <thead>
                    <tr><th>Date</th><th>Description</th><th>Post reference</th><th>Debit</th><th>Credit</th><th>Total debit</th><th>Total credit</th><th>Entered by</th><th className="sms-no-print"></th></tr>
                  </thead>
                  <tbody>
                    {ledgerWithRunning.map((l) => (
                      <tr key={l.id}>
                        <td>{l.date}</td>
                        <td>{l.description}{l.comments && <span className="sms-subtext">{l.comments}</span>}</td>
                        <td>{l.postReference}</td>
                        <td>{l.debit > 0 ? currency(l.debit) : "—"}</td>
                        <td>{l.credit > 0 ? currency(l.credit) : "—"}</td>
                        <td>{currency(l.totalDebit)}</td>
                        <td>{currency(l.totalCredit)}</td>
                        <td className="sms-subtext">{l.enteredBy || "—"}<br />{formatDateTime(l.enteredAt)}</td>
                        <td className="sms-no-print" style={{ display: "flex", gap: 6 }}>
                          {isAdmin ? (<>
                            <button className="sms-btn secondary small" onClick={() => setLedgerModal(l)}>Update</button>
                            <button className="sms-btn danger small" onClick={() => deleteLedgerEntry(l.id)}>Delete</button>
                          </>) : <span className="sms-locked">Admin only</span>}
                        </td>
                      </tr>
                    ))}
                    {ledgerWithRunning.length === 0 && <tr><td colSpan="9" className="sms-empty">No ledger entries yet.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "studentattendance" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Student attendance</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{studentAttendance.length} records</span><button className="sms-btn secondary" onClick={exportStudentAttendance}>⬇ Export Excel (all data)</button></div>
            </div>
            <div className="sms-content">
              <div className="sms-toolbar">
                <Field label="Date"><input className="sms-input" type="date" value={attDate} onChange={(e) => { setAttDate(e.target.value); setAttDraft({}); }} /></Field>
                <Field label="Class"><select className="sms-select" value={attClass} onChange={(e) => { setAttClass(e.target.value); setAttDraft({}); }}>{CLASSES.map((c) => <option key={c} value={c}>{c}</option>)}</select></Field>
              </div>

              <div className="sms-cards-row three">
                <div className="sms-card"><div className="sms-card-label">Present</div><div className="sms-card-value green">{attPresentCount}</div></div>
                <div className="sms-card"><div className="sms-card-label">Absent</div><div className="sms-card-value rust">{attAbsentCount}</div></div>
                <div className="sms-card"><div className="sms-card-label">Leave</div><div className="sms-card-value">{attLeaveCount}</div></div>
              </div>

              <div className="sms-section-title">Mark attendance <span className="sms-tag">{attClass} · {attDate}</span></div>
              <div className="sms-toolbar">
                <button className="sms-btn secondary small" onClick={() => {
                  const all = {}; classStudentsForAttendance.forEach((s) => { all[s.id] = "present"; }); setAttDraft(all);
                }}>Mark all present</button>
              </div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Name</th><th>Roll no.</th><th>Status</th></tr></thead>
                  <tbody>
                    {classStudentsForAttendance.map((s) => {
                      const existing = attendanceForDateByStudent[s.id];
                      const current = attDraft[s.id] || (existing ? existing.status : "present");
                      return (
                        <tr key={s.id}>
                          <td>{s.name}</td>
                          <td className="sms-mono">{s.rollNo}</td>
                          <td>
                            <select className="sms-select" value={current} onChange={(e) => setAttDraft({ ...attDraft, [s.id]: e.target.value })}>
                              <option value="present">Present</option>
                              <option value="absent">Absent</option>
                              <option value="leave">Leave</option>
                            </select>
                          </td>
                        </tr>
                      );
                    })}
                    {classStudentsForAttendance.length === 0 && <tr><td colSpan="3" className="sms-empty">No active students in this class.</td></tr>}
                  </tbody>
                </table>
              </div>
              {classStudentsForAttendance.length > 0 && (
                <div className="sms-toolbar" style={{ marginTop: 12 }}>
                  <button className="sms-btn" onClick={() => { saveStudentAttendanceBulk(attDate, attClass, attDraft); setAttDraft({}); }}>Save attendance for {attDate}</button>
                </div>
              )}
            </div>
          </>
        )}

        {page === "staffattendance" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Staff attendance</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{staffAttendance.length} records</span><button className="sms-btn secondary" onClick={exportStaffAttendance}>⬇ Export Excel (all data)</button></div>
            </div>
            <div className="sms-content">
              <div className="sms-toolbar">
                <Field label="Staff member">
                  <select className="sms-select" value={effectiveStaffAttId} onChange={(e) => setStaffAttStaffId(e.target.value)}>
                    {staff.map((s) => <option key={s.id} value={s.id}>{s.name} — {s.role}</option>)}
                  </select>
                </Field>
                <Field label="Month"><select className="sms-select" value={staffAttMonthFilter} onChange={(e) => setStaffAttMonthFilter(e.target.value)}>{MONTHS.map((m) => <option key={m}>{m}</option>)}</select></Field>
                <Field label="Year"><input className="sms-input" type="number" style={{ width: 90 }} value={staffAttYearFilter} onChange={(e) => setStaffAttYearFilter(e.target.value)} /></Field>
                <button className="sms-btn" onClick={() => setStaffAttModal({ staffId: effectiveStaffAttId, date: todayISO() })} disabled={!effectiveStaffAttId}>+ Add day entry</button>
              </div>

              <div className="sms-cards-row">
                <div className="sms-card"><div className="sms-card-label">Regular hrs</div><div className="sms-card-value">{staffAttTotals.regular}</div></div>
                <div className="sms-card"><div className="sms-card-label">Overtime hrs</div><div className="sms-card-value">{staffAttTotals.overtime}</div></div>
                <div className="sms-card"><div className="sms-card-label">Absent / LWP days</div><div className="sms-card-value rust">{staffAttTotals.absent} / {staffAttTotals.lwp}</div></div>
                <div className="sms-card"><div className="sms-card-label">Advance + other</div><div className="sms-card-value">{currency(staffAttTotals.advance + staffAttTotals.other)}</div></div>
              </div>

              {staffAttForStaff.some((a) => (a.approvalStatus || "approved") === "pending") && (
                <div className="sms-alert-banner" style={{ background: "var(--amber-bg)", border: "1px solid var(--amber)", color: "var(--amber)" }}>
                  <span>⏳ {staffAttForStaff.filter((a) => (a.approvalStatus || "approved") === "pending").length} self check-in/out entr{staffAttForStaff.filter((a) => (a.approvalStatus || "approved") === "pending").length > 1 ? "ies are" : "y is"} awaiting your approval. Unapproved hours are not included when generating salaries.</span>
                </div>
              )}

              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Date</th><th>Check in</th><th>Check out</th><th>Regular hrs</th><th>Overtime hrs</th><th>Absent</th><th>LWP</th><th>Advance</th><th>Other</th><th>Total hrs</th><th>Approval</th><th></th></tr></thead>
                  <tbody>
                    {staffAttForStaff.map((a) => {
                      const approval = a.approvalStatus || "approved";
                      return (
                      <tr key={a.id} className={approval === "pending" ? "sms-row-due" : approval === "rejected" ? "sms-row-overdue" : ""}>
                        <td>{a.date}</td>
                        <td className="sms-mono">{a.checkIn || "—"}</td>
                        <td className="sms-mono">{a.checkOut || "—"}</td>
                        <td className="sms-mono">{a.regularHrs || 0}</td>
                        <td className="sms-mono">{a.overtimeHrs || 0}</td>
                        <td>{a.absent ? <span className="sms-pill unpaid">Yes</span> : "—"}</td>
                        <td>{a.lwp ? <span className="sms-pill unpaid">Yes</span> : "—"}</td>
                        <td className="sms-mono">{a.advance > 0 ? currency(a.advance) : "—"}</td>
                        <td className="sms-mono">{a.other > 0 ? currency(a.other) : "—"}</td>
                        <td className="sms-mono">{(Number(a.regularHrs) || 0) + (Number(a.overtimeHrs) || 0)}</td>
                        <td><span className={"sms-pill " + (approval === "pending" ? "due" : approval === "rejected" ? "overdue" : "paid")}>{approval}</span></td>
                        <td style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                          {isAdmin ? (<>
                            <button className="sms-btn secondary small" onClick={() => setStaffAttModal(a)}>Update</button>
                            <button className="sms-btn danger small" onClick={() => deleteStaffAttendanceEntry(a.id)}>Delete</button>
                            {approval !== "approved" && <button className="sms-btn small" onClick={() => setAttendanceApproval(a, "approved")}>Approve</button>}
                            {approval !== "rejected" && <button className="sms-btn danger small" onClick={() => setAttendanceApproval(a, "rejected")}>Reject</button>}
                          </>) : <span className="sms-locked">Admin only</span>}
                        </td>
                      </tr>
                    );})}
                    {staffAttForStaff.length === 0 && <tr><td colSpan="12" className="sms-empty">No attendance entries for this staff member this month.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "myduty" && (() => {
          const linkedStaff = staffMap[currentUser.staffId];
          const today = myTodayAttendance(currentUser.staffId);
          const myHistory = staffAttendance
            .filter((a) => a.staffId === currentUser.staffId)
            .sort((a, b) => (b.date || "").localeCompare(a.date || ""))
            .slice(0, 10);
          const approval = today ? (today.approvalStatus || "approved") : null;
          return (
            <>
              <div className="sms-header">
                <h1 className="sms-serif">{t("myDuty")}</h1>
                <span className="sms-datestamp">{todayISO()}</span>
              </div>
              <div className="sms-content">
                <div className="sms-section-title">{linkedStaff ? linkedStaff.name : currentUser.name} <span className="sms-tag">{linkedStaff ? linkedStaff.role : ""}</span></div>

                <div className="sms-duty-card">
                  {!today && (
                    <>
                      <div className="sms-duty-status">{t("notStartedDuty")}</div>
                      <button className="sms-btn gold sms-duty-btn" onClick={startDuty}>{t("startDuty")}</button>
                    </>
                  )}
                  {today && !today.checkOut && (
                    <>
                      <div className="sms-duty-status">{t("onDutySince")} <strong>{formatTime12h(today.checkIn)}</strong></div>
                      <button className="sms-btn danger sms-duty-btn" onClick={endDuty}>{t("endDuty")}</button>
                    </>
                  )}
                  {today && today.checkOut && (
                    <>
                      <div className="sms-duty-status">
                        {t("dutyCompleted")}: <strong>{formatTime12h(today.checkIn)} – {formatTime12h(today.checkOut)}</strong>
                        <span className="sms-subtext">{(Number(today.regularHrs) || 0) + (Number(today.overtimeHrs) || 0)} {t("hoursToday")}{today.overtimeHrs > 0 ? ` (${t("includesOvertime")}: ${today.overtimeHrs}h)` : ""}</span>
                      </div>
                      <span className={"sms-pill " + (approval === "pending" ? "due" : approval === "rejected" ? "overdue" : "paid")}>{t(approval)}</span>
                      {approval === "pending" && <div className="sms-subtext" style={{ marginTop: 8 }}>{t("awaitingApproval")}</div>}
                      {approval === "rejected" && <div className="sms-subtext" style={{ marginTop: 8 }}>{t("entryRejected")}</div>}
                    </>
                  )}
                </div>

                <div className="sms-section-title">{t("recentHistory")} <span className="sms-tag">{t("last10")}</span></div>
                <div className="sms-ledger-page">
                  <table className="sms-table">
                    <thead><tr><th>{t("date")}</th><th>{t("checkIn")}</th><th>{t("checkOut")}</th><th>{t("totalHrs")}</th><th>{t("approval")}</th></tr></thead>
                    <tbody>
                      {myHistory.map((a) => {
                        const st = a.approvalStatus || "approved";
                        return (
                          <tr key={a.id}>
                            <td>{a.date}</td>
                            <td className="sms-mono">{formatTime12h(a.checkIn)}</td>
                            <td className="sms-mono">{formatTime12h(a.checkOut)}</td>
                            <td className="sms-mono">{(Number(a.regularHrs) || 0) + (Number(a.overtimeHrs) || 0)}</td>
                            <td><span className={"sms-pill " + (st === "pending" ? "due" : st === "rejected" ? "overdue" : "paid")}>{t(st)}</span></td>
                          </tr>
                        );
                      })}
                      {myHistory.length === 0 && <tr><td colSpan="5" className="sms-empty">{t("noHistoryYet")}</td></tr>}
                    </tbody>
                  </table>
                </div>
              </div>
            </>
          );
        })()}

        {page === "useraccounts" && (
          <>
            <div className="sms-header">
              <h1 className="sms-serif">Staff logins</h1>
              <div className="sms-header-actions"><span className="sms-datestamp">{(settings.userAccounts || []).length} accounts</span></div>
            </div>
            <div className="sms-content">
              <div className="sms-subtext" style={{ marginBottom: 14 }}>
                Give each teacher their own username and password instead of sharing one staff code. Link an account to a staff record to unlock their personal "My duty" check-in/check-out tab.
              </div>
              <div className="sms-toolbar">
                <button className="sms-btn" onClick={() => setAccountModal({})}>+ Add login</button>
              </div>
              <div className="sms-ledger-page">
                <table className="sms-table">
                  <thead><tr><th>Username</th><th>Name</th><th>Role</th><th>Linked staff member</th><th>Tab access</th><th>Status</th><th></th></tr></thead>
                  <tbody>
                    {(settings.userAccounts || []).map((a) => {
                      const linked = staffMap[a.staffId];
                      return (
                        <tr key={a.id}>
                          <td className="sms-mono">{a.username}</td>
                          <td>{a.name}</td>
                          <td>{a.role === "admin" ? "Admin" : "Staff"}</td>
                          <td>{linked ? linked.name : "—"}</td>
                          <td>{a.role === "admin" ? "Full access" : a.customTabAccess ? <span className="sms-pill fund">Custom</span> : <span className="sms-subtext">General default</span>}</td>
                          <td><span className={"sms-pill " + (a.status === "inactive" ? "inactive" : "active")}>{a.status === "inactive" ? "inactive" : "active"}</span></td>
                          <td style={{ display: "flex", gap: 6 }}>
                            <button className="sms-btn secondary small" onClick={() => setAccountModal(a)}>Edit</button>
                            <button className="sms-btn danger small" onClick={() => deleteAccount(a.id)}>Delete</button>
                          </td>
                        </tr>
                      );
                    })}
                    {(settings.userAccounts || []).length === 0 && <tr><td colSpan="7" className="sms-empty">No individual logins yet — everyone is still using the shared Admin/Staff passwords. Add one above to get started.</td></tr>}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {page === "noaccess" && (
          <>
            <div className="sms-header"><h1 className="sms-serif">No access yet</h1></div>
            <div className="sms-content">
              <div className="sms-empty" style={{ padding: 60 }}>
                Your account doesn't currently have access to any section of the app.<br />
                Please contact your admin to have some tabs enabled for your login.
              </div>
            </div>
          </>
        )}
      </main>

      {staffModal && <StaffForm initial={staffModal} onCancel={() => setStaffModal(null)} onSave={upsertStaff} />}
      {studentModal && <StudentForm initial={studentModal} onCancel={() => setStudentModal(null)} onSave={upsertStudent} />}
      {feeModal && <FeeForm initial={feeModal} students={students} settings={settings} onCancel={() => setFeeModal(null)} onSave={upsertFee} />}
      {itemModal && <ItemForm initial={itemModal} students={students} onCancel={() => setItemModal(null)} onSave={upsertItem} />}
      {pettyModal && <PettyForm initial={pettyModal} onCancel={() => setPettyModal(null)} onSave={upsertPetty} />}
      {salaryModal && <SalaryForm initial={salaryModal} staff={staff} onCancel={() => setSalaryModal(null)} onSave={upsertSalary} />}
      {ledgerModal && <LedgerForm initial={ledgerModal} onCancel={() => setLedgerModal(null)} onSave={upsertLedger} />}
      {staffAttModal && <StaffAttendanceForm initial={staffAttModal} staff={staff} onCancel={() => setStaffAttModal(null)} onSave={upsertStaffAttendance} />}
      {settingsModal && isAdmin && (
        <SettingsForm settings={settings} onCancel={() => setSettingsModal(false)} onSave={(s) => { saveSettings(s); setSettingsModal(false); }} />
      )}
      {accountModal && isAdmin && (
        <AccountForm initial={accountModal} staff={staff} existingAccounts={settings.userAccounts || []} globalStaffTabAccess={settings.staffTabAccess}
          onCancel={() => setAccountModal(null)} onSave={(a) => { upsertAccount(a); setAccountModal(null); }} />
      )}

      {overdueModal && (
        <Modal title={`Overdue fee alerts (${overdueStudentIds.length})`} wide onClose={() => setOverdueModal(false)}
          footer={<button className="sms-btn secondary" onClick={() => setOverdueModal(false)}>Close</button>}>
          {overdueStudentIds.length === 0 ? <div className="sms-empty">No overdue students right now.</div> : (
            <div className="sms-overdue-list">
              {overdueStudentIds.map((sid) => {
                const st = studentMap[sid];
                const info = overdueByStudent[sid];
                if (!st) return null;
                return (
                  <div className="sms-overdue-item" key={sid}>
                    <div>
                      <div className="sms-oi-name">{st.name} — {st.class}{st.section ? "-" + st.section : ""}</div>
                      <div className="sms-oi-meta">Unpaid: {info.months.join(", ")} · Parent: {st.parentPhone || "—"}</div>
                    </div>
                    <div className="sms-oi-amount">{currency(info.amount)}</div>
                  </div>
                );
              })}
            </div>
          )}
        </Modal>
      )}
    </div>
  );
}

function AccountForm({ initial, staff, existingAccounts, globalStaffTabAccess, onCancel, onSave }) {
  const isEdit = !!initial.id;
  const restrictableTabs = ALL_TABS.filter((t) => !t.requiresStaffLink && !t.adminOnly);
  const [form, setForm] = useState({
    id: initial.id || uid("acct"),
    username: initial.username || "",
    password: initial.password || "",
    name: initial.name || "",
    role: initial.role || "staff",
    staffId: initial.staffId || "",
    status: initial.status || "active",
    customTabAccess: initial.customTabAccess || false,
    tabAccess: { ...(initial.tabAccess || {}) },
  });
  const [error, setError] = useState("");
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  function onStaffLinkChange(e) {
    const staffId = e.target.value;
    const st = staff.find((s) => s.id === staffId);
    setForm({ ...form, staffId, name: form.name || (st ? st.name : form.name) });
  }
  function toggleCustom(e) {
    const on = e.target.checked;
    if (on) {
      // Seed the per-person checklist from the current general default so
      // the admin starts from what this person already effectively sees.
      const seeded = {};
      restrictableTabs.forEach((t) => { seeded[t.key] = form.tabAccess[t.key] !== undefined ? form.tabAccess[t.key] : (globalStaffTabAccess?.[t.key] !== false); });
      setForm({ ...form, customTabAccess: true, tabAccess: seeded });
    } else {
      setForm({ ...form, customTabAccess: false });
    }
  }
  function toggleTab(key) {
    const current = form.tabAccess[key] !== false;
    setForm({ ...form, tabAccess: { ...form.tabAccess, [key]: !current } });
  }

  function handleSave() {
    if (!form.username.trim() || !form.password.trim() || !form.name.trim()) {
      setError("Username, password, and name are all required.");
      return;
    }
    const dupe = existingAccounts.some((a) => a.id !== form.id && a.username.trim().toLowerCase() === form.username.trim().toLowerCase());
    if (dupe) {
      setError("That username is already taken — pick a different one.");
      return;
    }
    onSave(form);
  }

  return (
    <Modal title={isEdit ? "Edit staff login" : "Add staff login"} wide onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" onClick={handleSave}>Save login</button>
    </>}>
      <Field label="Full name"><input className="sms-input" value={form.name} onChange={set("name")} placeholder="Sadia Tariq" /></Field>
      <div className="sms-field-row">
        <Field label="Username"><input className="sms-input" value={form.username} onChange={set("username")} placeholder="sadia.t" autoCapitalize="none" /></Field>
        <Field label="Password"><input className="sms-input" value={form.password} onChange={set("password")} placeholder="Choose a password" /></Field>
      </div>
      <div className="sms-field-row">
        <Field label="Role"><select className="sms-select" value={form.role} onChange={set("role")}>
          <option value="staff">Staff (limited access)</option>
          <option value="admin">Admin (full access)</option>
        </select></Field>
        <Field label="Linked staff member (optional)">
          <select className="sms-select" value={form.staffId} onChange={onStaffLinkChange}>
            <option value="">None</option>
            {staff.map((s) => <option key={s.id} value={s.id}>{s.name} — {s.role}</option>)}
          </select>
        </Field>
      </div>
      <div className="sms-subtext">Linking a staff member unlocks their personal "My duty" check-in/check-out tab, and ties their self-logged hours to their own attendance and salary records.</div>
      <Field label="Status"><select className="sms-select" value={form.status} onChange={set("status")}><option value="active">Active</option><option value="inactive">Inactive (blocks login)</option></select></Field>

      {form.role === "staff" && (
        <>
          <div className="sms-checkbox-row"><input type="checkbox" checked={form.customTabAccess} onChange={toggleCustom} id="customtabs" /><label htmlFor="customtabs">Give this person custom tab access (by position/designation), instead of the general staff default</label></div>
          {form.customTabAccess ? (
            <Field label="Tabs this person can see">
              <div className="sms-subtext" style={{ marginBottom: 8 }}>Uncheck anything they shouldn't have access to — e.g. a Teacher might only need Student Attendance, while an Accountant needs Fee Ledger and Petty Cash.</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px 16px" }}>
                {restrictableTabs.map((t) => (
                  <label key={t.key} className="sms-checkbox-row">
                    <input type="checkbox" checked={form.tabAccess[t.key] !== false} onChange={() => toggleTab(t.key)} />
                    {t.label}
                  </label>
                ))}
              </div>
            </Field>
          ) : (
            <div className="sms-subtext">This person follows the general staff default set in Settings &amp; access. Check the box above to set specific tabs for just this person instead.</div>
          )}
        </>
      )}

      {error && <div className="sms-login-error">{error}</div>}
    </Modal>
  );
}

function SettingsForm({ settings, onCancel, onSave }) {
  const [form, setForm] = useState({ ...settings, staffTabAccess: { ...(settings.staffTabAccess || {}) } });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  function toggleTab(key) {
    const current = form.staffTabAccess[key] !== false;
    setForm({ ...form, staffTabAccess: { ...form.staffTabAccess, [key]: !current } });
  }
  const restrictableTabs = ALL_TABS.filter((t) => !t.requiresStaffLink && !t.adminOnly);
  return (
    <Modal title="Settings & access" wide onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" onClick={() => onSave(form)}>Save settings</button>
    </>}>
      <div className="sms-field-row">
        <Field label="Papers fund amount (Apr & Aug)"><input className="sms-input" type="number" value={form.papersFundAmount} onChange={set("papersFundAmount")} /></Field>
        <Field label="Late fee amount"><input className="sms-input" type="number" value={form.lateFeeAmount} onChange={set("lateFeeAmount")} /></Field>
      </div>
      <div className="sms-field-row">
        <Field label="Admin password"><input className="sms-input" value={form.adminPassword} onChange={set("adminPassword")} /></Field>
        <Field label="Staff access code"><input className="sms-input" value={form.staffPassword} onChange={set("staffPassword")} /></Field>
      </div>
      <Field label="General default tab access for staff">
        <div className="sms-subtext" style={{ marginBottom: 8 }}>Dashboard is included below like any other tab — uncheck it if staff shouldn't see overall totals and financial summaries. This is the default used by the shared "Staff" login and any individual account without custom access set. To give a specific person a different set of tabs (e.g. by their position), edit their login in "Staff logins" instead.</div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "6px 16px" }}>
          {restrictableTabs.map((t) => (
            <label key={t.key} className="sms-checkbox-row">
              <input type="checkbox" checked={form.staffTabAccess[t.key] !== false} onChange={() => toggleTab(t.key)} />
              {t.label}
            </label>
          ))}
        </div>
      </Field>
    </Modal>
  );
}

function StaffForm({ initial, onCancel, onSave }) {
  const [form, setForm] = useState({
    id: initial.id || uid("staff"), name: initial.name || "", role: initial.role || "Teacher",
    subject: initial.subject || "", phone: initial.phone || "", email: initial.email || "",
    joinDate: initial.joinDate || todayISO(), salary: initial.salary || "", hourlyRate: initial.hourlyRate || "",
    status: initial.status || "active", inactiveReason: initial.inactiveReason || "", inactiveDate: initial.inactiveDate || todayISO(),
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const needsReason = form.status === "inactive";
  const valid = form.name.trim().length > 0 && (!needsReason || (form.inactiveReason.trim().length > 0 && form.inactiveDate));
  return (
    <Modal title={initial.id ? "Edit staff record" : "Add staff record"} onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" disabled={!valid} onClick={() => onSave({ ...form, salary: Number(form.salary) || 0, hourlyRate: Number(form.hourlyRate) || 0, inactiveDate: needsReason ? form.inactiveDate : "" })}>Save record</button>
    </>}>
      <Field label="Full name"><input className="sms-input" value={form.name} onChange={set("name")} placeholder="Anita Sharma" /></Field>
      <div className="sms-field-row">
        <Field label="Role"><select className="sms-select" value={form.role} onChange={set("role")}>
          {["Teacher","Principal","Vice Principal","Admin staff","Accountant","Librarian","Lab assistant","Support staff"].map((r) => <option key={r}>{r}</option>)}
        </select></Field>
        <Field label="Subject / department"><input className="sms-input" value={form.subject} onChange={set("subject")} placeholder="Mathematics" /></Field>
      </div>
      <div className="sms-field-row">
        <Field label="Phone"><input className="sms-input" value={form.phone} onChange={set("phone")} placeholder="+92 300 1234567" /></Field>
        <Field label="Email"><input className="sms-input" value={form.email} onChange={set("email")} placeholder="anita@school.edu.pk" /></Field>
      </div>
      <div className="sms-field-row">
        <Field label="Joining date"><input className="sms-input" type="date" value={form.joinDate} onChange={set("joinDate")} /></Field>
        <Field label="Monthly salary"><input className="sms-input" type="number" value={form.salary} onChange={set("salary")} placeholder="45000" /></Field>
      </div>
      <Field label="Hourly rate (optional)">
        <input className="sms-input" type="number" value={form.hourlyRate} onChange={set("hourlyRate")} placeholder="e.g. 200" />
      </Field>
      <div className="sms-subtext" style={{ marginTop: -6 }}>If set and this staff member has check-in/check-out attendance logged for a month, "Generate this month's salaries" will calculate their pay from actual hours worked instead of the flat monthly salary above.</div>
      <Field label="Status"><select className="sms-select" value={form.status} onChange={set("status")}><option value="active">Active</option><option value="inactive">Inactive</option></select></Field>
      {needsReason && (
        <div className="sms-field-row">
          <Field label="Inactive since (required)">
            <input className="sms-input" type="date" value={form.inactiveDate} onChange={set("inactiveDate")} />
          </Field>
          <Field label="Reason for inactive status (required)">
            <input className="sms-input" value={form.inactiveReason} onChange={set("inactiveReason")} placeholder="e.g. Resigned, termination…" />
          </Field>
        </div>
      )}
    </Modal>
  );
}

function StudentForm({ initial, onCancel, onSave }) {
  const [form, setForm] = useState({
    id: initial.id || uid("stu"), name: initial.name || "", class: initial.class || CLASSES[0], section: initial.section || "",
    rollNo: initial.rollNo || "", admissionDate: initial.admissionDate || todayISO(), parentName: initial.parentName || "",
    parentPhone: initial.parentPhone || "", fatherName: initial.fatherName || "", fatherCnic: initial.fatherCnic || "",
    emergencyContact: initial.emergencyContact || "",
    registrationFee: initial.registrationFee ?? "", admissionFee: initial.admissionFee ?? "",
    standardFee: initial.standardFee ?? initial.monthlyFee ?? "",
    discountType: initial.discountType || "none", discountValue: initial.discountValue || "",
    feeWaived: initial.feeWaived || false, status: initial.status || "active",
    inactiveReason: initial.inactiveReason || "", inactiveDate: initial.inactiveDate || todayISO(),
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const setBool = (k) => (e) => setForm({ ...form, [k]: e.target.checked });
  const needsReason = form.status === "inactive";
  const valid = form.name.trim().length > 0 && form.rollNo.trim().length > 0 && (!needsReason || (form.inactiveReason.trim().length > 0 && form.inactiveDate));
  const { discount, netFee } = computeNetFee(form.standardFee, form.discountType, form.discountValue);
  const oneTimeFees = (Number(form.registrationFee) || 0) + (Number(form.admissionFee) || 0);
  return (
    <Modal title={initial.id ? "Edit admission" : "New admission"} wide onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" disabled={!valid} onClick={() => onSave({
        ...form, standardFee: Number(form.standardFee) || 0, discountValue: Number(form.discountValue) || 0,
        registrationFee: Number(form.registrationFee) || 0, admissionFee: Number(form.admissionFee) || 0,
        inactiveDate: needsReason ? form.inactiveDate : "",
      })}>Save admission</button>
    </>}>
      <Field label="Student name"><input className="sms-input" value={form.name} onChange={set("name")} placeholder="Ahmed Raza" /></Field>
      <div className="sms-field-row">
        <Field label="Class"><select className="sms-select" value={form.class} onChange={set("class")}>{CLASSES.map((c) => <option key={c} value={c}>{c}</option>)}</select></Field>
        <Field label="Section"><input className="sms-input" value={form.section} onChange={set("section")} placeholder="A" /></Field>
      </div>
      <div className="sms-field-row">
        <Field label="Roll number"><input className="sms-input" value={form.rollNo} onChange={set("rollNo")} placeholder="2026-014" /></Field>
        <Field label="Admission date"><input className="sms-input" type="date" value={form.admissionDate} onChange={set("admissionDate")} /></Field>
      </div>
      <div className="sms-field-row">
        <Field label="Father's name"><input className="sms-input" value={form.fatherName} onChange={set("fatherName")} placeholder="Imran Raza" /></Field>
        <Field label="Father's CNIC"><input className="sms-input" value={form.fatherCnic} onChange={set("fatherCnic")} placeholder="36302-1234567-1" /></Field>
      </div>
      <div className="sms-field-row">
        <Field label="Parent / guardian name"><input className="sms-input" value={form.parentName} onChange={set("parentName")} placeholder="Imran Raza" /></Field>
        <Field label="Parent phone"><input className="sms-input" value={form.parentPhone} onChange={set("parentPhone")} placeholder="+92 300 1234567" /></Field>
      </div>
      <Field label="Emergency contact number"><input className="sms-input" value={form.emergencyContact} onChange={set("emergencyContact")} placeholder="+92 300 7654321" /></Field>

      <div className="sms-field-row">
        <Field label="Registration fee (one-time)"><input className="sms-input" type="number" value={form.registrationFee} onChange={set("registrationFee")} placeholder="1000" /></Field>
        <Field label="Admission fee (one-time)"><input className="sms-input" type="number" value={form.admissionFee} onChange={set("admissionFee")} placeholder="2000" /></Field>
      </div>
      {oneTimeFees > 0 && <div className="sms-computed">One-time fees due at admission: {currency(oneTimeFees)}</div>}

      <Field label="Monthly fee"><input className="sms-input" type="number" value={form.standardFee} onChange={set("standardFee")} placeholder="3500" disabled={form.feeWaived} /></Field>
      <div className="sms-field-row">
        <Field label="Discount type"><select className="sms-select" value={form.discountType} onChange={set("discountType")} disabled={form.feeWaived}>
          <option value="none">No discount — pays standard fee</option>
          <option value="percent">Percentage off</option>
          <option value="amount">Fixed amount off</option>
        </select></Field>
        <Field label={form.discountType === "percent" ? "Discount %" : "Discount amount"}>
          <input className="sms-input" type="number" value={form.discountValue} onChange={set("discountValue")} disabled={form.discountType === "none" || form.feeWaived} placeholder={form.discountType === "percent" ? "10" : "500"} />
        </Field>
      </div>
      <div className="sms-checkbox-row"><input type="checkbox" checked={form.feeWaived} onChange={setBool("feeWaived")} id="feewaived" /><label htmlFor="feewaived">Fee exempt (financial assistance) — do not charge this student</label></div>
      <div className="sms-computed">{form.feeWaived ? "Net monthly fee: Rs 0 (exempt)" : `Net monthly fee: ${currency(netFee)}${discount > 0 ? ` (discount ${currency(discount)})` : ""}`}</div>
      <Field label="Status"><select className="sms-select" value={form.status} onChange={set("status")}><option value="active">Active</option><option value="inactive">Inactive</option></select></Field>
      {needsReason && (
        <div className="sms-field-row">
          <Field label="Inactive since (required)">
            <input className="sms-input" type="date" value={form.inactiveDate} onChange={set("inactiveDate")} />
          </Field>
          <Field label="Reason for inactive status (required)">
            <input className="sms-input" value={form.inactiveReason} onChange={set("inactiveReason")} placeholder="e.g. Left the school, transferred…" />
          </Field>
        </div>
      )}
    </Modal>
  );
}

function FeeForm({ initial, students, settings, onCancel, onSave }) {
  const isEdit = !!initial.id;
  const firstStudent = students[0];
  const startMonth = initial.month || MONTHS[new Date().getMonth()];
  const [form, setForm] = useState({
    id: initial.id || uid("fee"),
    studentId: initial.studentId || (firstStudent ? firstStudent.id : ""),
    month: startMonth, year: initial.year || new Date().getFullYear(),
    standardFee: initial.standardFee ?? (firstStudent ? firstStudent.standardFee : 0),
    discountType: initial.discountType ?? (firstStudent ? firstStudent.discountType : "none"),
    discountValue: initial.discountValue ?? (firstStudent ? firstStudent.discountValue : 0),
    papersFundEnabled: initial.papersFundEnabled ?? PAPERS_FUND_MONTHS.includes(startMonth),
    papersFund: initial.papersFund ?? (PAPERS_FUND_MONTHS.includes(startMonth) ? settings.papersFundAmount : 0),
    lateFeeEnabled: initial.lateFeeEnabled ?? false,
    lateFee: initial.lateFee ?? 0,
    paidAmount: initial.paidAmount ?? 0, paidDate: initial.paidDate || "",
    free: initial.status === "free",
    comments: initial.comments || "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  function onStudentChange(e) {
    const sid = e.target.value;
    const st = students.find((s) => s.id === sid);
    setForm({ ...form, studentId: sid, standardFee: st ? st.standardFee : form.standardFee,
      discountType: st ? st.discountType : form.discountType, discountValue: st ? st.discountValue : form.discountValue,
      free: st ? !!st.feeWaived : form.free });
  }
  function onMonthChange(e) {
    const month = e.target.value;
    const enable = PAPERS_FUND_MONTHS.includes(month);
    setForm({ ...form, month, papersFundEnabled: enable, papersFund: enable ? settings.papersFundAmount : 0 });
  }
  function togglePapersFund(e) {
    const on = e.target.checked;
    setForm({ ...form, papersFundEnabled: on, papersFund: on ? settings.papersFundAmount : 0 });
  }
  function toggleLateFee(e) {
    const on = e.target.checked;
    setForm({ ...form, lateFeeEnabled: on, lateFee: on ? settings.lateFeeAmount : 0 });
  }

  const { discount, netFee } = computeNetFee(form.standardFee, form.discountType, form.discountValue);
  const totalDue = form.free ? 0 : netFee + (Number(form.papersFund) || 0) + (Number(form.lateFee) || 0);

  function handleSave() {
    if (form.free) {
      onSave({ ...form, year: Number(form.year), discount: 0, netFee: 0, papersFund: 0, lateFee: 0, paidAmount: 0, status: "free", paidDate: "" });
      return;
    }
    const paid = Number(form.paidAmount) || 0;
    const status = paid <= 0 ? "unpaid" : paid >= totalDue ? "paid" : "partial";
    onSave({
      ...form, year: Number(form.year), standardFee: Number(form.standardFee) || 0, discountValue: Number(form.discountValue) || 0,
      discount, netFee, papersFund: Number(form.papersFund) || 0, lateFee: Number(form.lateFee) || 0,
      paidAmount: paid, status, paidDate: status === "unpaid" ? "" : (form.paidDate || todayISO()),
    });
  }

  return (
    <Modal title={isEdit ? "Update fee record" : "Add fee record"} wide onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" onClick={handleSave} disabled={!form.studentId}>Save record</button>
    </>}>
      <Field label="Student">
        <select className="sms-select" value={form.studentId} onChange={onStudentChange} disabled={isEdit}>
          {students.map((s) => <option key={s.id} value={s.id}>{s.name} — {s.class}{s.section ? "-" + s.section : ""}</option>)}
        </select>
      </Field>
      <div className="sms-field-row">
        <Field label="Month"><select className="sms-select" value={form.month} onChange={onMonthChange}>
          {MONTHS.map((m) => <option key={m} value={m}>{m}{PAPERS_FUND_MONTHS.includes(m) ? " (papers fund)" : ""}</option>)}
        </select></Field>
        <Field label="Year"><input className="sms-input" type="number" value={form.year} onChange={set("year")} /></Field>
      </div>

      <div className="sms-checkbox-row"><input type="checkbox" checked={form.free} onChange={(e) => setForm({ ...form, free: e.target.checked })} id="feefree" /><label htmlFor="feefree">This is a free / exempt month for this student — do not charge anything (financial assistance)</label></div>

      {!form.free ? (
        <>
          <div className="sms-field-row three">
            <Field label="Standard fee"><input className="sms-input" type="number" value={form.standardFee} onChange={set("standardFee")} /></Field>
            <Field label="Discount type"><select className="sms-select" value={form.discountType} onChange={set("discountType")}>
              <option value="none">None</option><option value="percent">Percent</option><option value="amount">Fixed amount</option>
            </select></Field>
            <Field label={form.discountType === "percent" ? "Discount %" : "Discount amount"}>
              <input className="sms-input" type="number" value={form.discountValue} onChange={set("discountValue")} disabled={form.discountType === "none"} />
            </Field>
          </div>

          <div className="sms-checkbox-row"><input type="checkbox" checked={form.papersFundEnabled} onChange={togglePapersFund} id="pf" /><label htmlFor="pf">Charge papers fund this month (optional — waive for poor students)</label></div>
          {form.papersFundEnabled && <Field label="Papers fund amount"><input className="sms-input" type="number" value={form.papersFund} onChange={set("papersFund")} /></Field>}

          <div className="sms-checkbox-row"><input type="checkbox" checked={form.lateFeeEnabled} onChange={toggleLateFee} id="lf" /><label htmlFor="lf">Charge late fee (optional — waive for poor students)</label></div>
          {form.lateFeeEnabled && <Field label="Late fee amount"><input className="sms-input" type="number" value={form.lateFee} onChange={set("lateFee")} /></Field>}

          <div className="sms-computed">
            Net fee {currency(netFee)}{discount > 0 ? ` (discount ${currency(discount)})` : ""} + papers fund {currency(form.papersFund)} + late fee {currency(form.lateFee)} = total due {currency(totalDue)}
          </div>
          <div className="sms-field-row">
            <Field label="Amount paid"><input className="sms-input" type="number" value={form.paidAmount} onChange={set("paidAmount")} /></Field>
            <Field label="Payment date"><input className="sms-input" type="date" value={form.paidDate} onChange={set("paidDate")} /></Field>
          </div>
        </>
      ) : (
        <div className="sms-free-box">This student will not be charged for {form.month} {form.year}. All amounts are set to Rs 0.</div>
      )}

      <Field label="Comments (optional)"><textarea value={form.comments} onChange={set("comments")} placeholder="Any notes about this payment…" /></Field>
    </Modal>
  );
}

function ItemForm({ initial, students, onCancel, onSave }) {
  const isEdit = !!initial.id;
  const firstStudent = students[0];
  const [form, setForm] = useState({
    id: initial.id || uid("item"), studentId: initial.studentId || (firstStudent ? firstStudent.id : ""),
    item: initial.item || ITEM_TYPES[0], description: initial.description || "", quantity: initial.quantity ?? 1,
    unitPrice: initial.unitPrice ?? "", paidAmount: initial.paidAmount ?? 0, issueDate: initial.issueDate || todayISO(),
    comments: initial.comments || "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const total = (Number(form.quantity) || 0) * (Number(form.unitPrice) || 0);
  function handleSave() { onSave({ ...form, quantity: Number(form.quantity) || 0, unitPrice: Number(form.unitPrice) || 0, paidAmount: Number(form.paidAmount) || 0 }); }
  return (
    <Modal title={isEdit ? "Update issue record" : "Add uniform / book issue"} onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" onClick={handleSave} disabled={!form.studentId}>Save record</button>
    </>}>
      <Field label="Student"><select className="sms-select" value={form.studentId} onChange={set("studentId")} disabled={isEdit}>
        {students.map((s) => <option key={s.id} value={s.id}>{s.name} — {s.class}{s.section ? "-" + s.section : ""}</option>)}
      </select></Field>
      <div className="sms-field-row">
        <Field label="Item"><select className="sms-select" value={form.item} onChange={set("item")}>{ITEM_TYPES.map((t) => <option key={t}>{t}</option>)}</select></Field>
        <Field label="Description (optional)"><input className="sms-input" value={form.description} onChange={set("description")} placeholder="Size medium" /></Field>
      </div>
      <div className="sms-field-row three">
        <Field label="Quantity"><input className="sms-input" type="number" value={form.quantity} onChange={set("quantity")} /></Field>
        <Field label="Unit price"><input className="sms-input" type="number" value={form.unitPrice} onChange={set("unitPrice")} /></Field>
        <Field label="Issue date"><input className="sms-input" type="date" value={form.issueDate} onChange={set("issueDate")} /></Field>
      </div>
      <div className="sms-computed">Total amount: {currency(total)}</div>
      <Field label="Amount paid"><input className="sms-input" type="number" value={form.paidAmount} onChange={set("paidAmount")} /></Field>
      <Field label="Comments (optional)"><textarea value={form.comments} onChange={set("comments")} placeholder="Any notes…" /></Field>
    </Modal>
  );
}

function SalaryForm({ initial, staff, onCancel, onSave }) {
  const isEdit = !!initial.id;
  const firstStaff = staff[0];
  const [form, setForm] = useState({
    id: initial.id || uid("sal"),
    staffId: initial.staffId || (firstStaff ? firstStaff.id : ""),
    month: initial.month || MONTHS[new Date().getMonth()],
    year: initial.year || new Date().getFullYear(),
    baseSalary: initial.baseSalary ?? (firstStaff ? firstStaff.salary : 0),
    bonus: initial.bonus ?? 0,
    deduction: initial.deduction ?? 0,
    paidAmount: initial.paidAmount ?? 0,
    paidDate: initial.paidDate || "",
    comments: initial.comments || "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });

  function onStaffChange(e) {
    const sid = e.target.value;
    const st = staff.find((s) => s.id === sid);
    setForm({ ...form, staffId: sid, baseSalary: st ? st.salary : form.baseSalary });
  }

  const net = (Number(form.baseSalary) || 0) + (Number(form.bonus) || 0) - (Number(form.deduction) || 0);

  function handleSave() {
    const paid = Number(form.paidAmount) || 0;
    const status = paid <= 0 ? "unpaid" : paid >= net ? "paid" : "partial";
    onSave({
      ...form, year: Number(form.year), baseSalary: Number(form.baseSalary) || 0,
      bonus: Number(form.bonus) || 0, deduction: Number(form.deduction) || 0,
      paidAmount: paid, status, paidDate: status === "unpaid" ? "" : (form.paidDate || todayISO()),
    });
  }

  return (
    <Modal title={isEdit ? "Update salary record" : "Add salary record"} wide onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" onClick={handleSave} disabled={!form.staffId}>Save record</button>
    </>}>
      <Field label="Staff member">
        <select className="sms-select" value={form.staffId} onChange={onStaffChange} disabled={isEdit}>
          {staff.map((s) => <option key={s.id} value={s.id}>{s.name} — {s.role}</option>)}
        </select>
      </Field>
      <div className="sms-field-row">
        <Field label="Month"><select className="sms-select" value={form.month} onChange={set("month")}>{MONTHS.map((m) => <option key={m}>{m}</option>)}</select></Field>
        <Field label="Year"><input className="sms-input" type="number" value={form.year} onChange={set("year")} /></Field>
      </div>
      <div className="sms-field-row three">
        <Field label="Base salary"><input className="sms-input" type="number" value={form.baseSalary} onChange={set("baseSalary")} /></Field>
        <Field label="Bonus (optional)"><input className="sms-input" type="number" value={form.bonus} onChange={set("bonus")} /></Field>
        <Field label="Deduction (optional)"><input className="sms-input" type="number" value={form.deduction} onChange={set("deduction")} /></Field>
      </div>
      <div className="sms-computed">
        Base {currency(form.baseSalary)} + bonus {currency(form.bonus)} - deduction {currency(form.deduction)} = net salary {currency(net)}
      </div>
      <div className="sms-field-row">
        <Field label="Amount paid"><input className="sms-input" type="number" value={form.paidAmount} onChange={set("paidAmount")} /></Field>
        <Field label="Payment date"><input className="sms-input" type="date" value={form.paidDate} onChange={set("paidDate")} /></Field>
      </div>
      <Field label="Comments (optional)"><textarea value={form.comments} onChange={set("comments")} placeholder="Any notes about this payment…" /></Field>
    </Modal>
  );
}

function LedgerForm({ initial, onCancel, onSave }) {
  const isEdit = !!initial.id;
  const [form, setForm] = useState({
    id: initial.id || uid("ledg"), date: initial.date || todayISO(), type: initial.credit > 0 && !(initial.debit > 0) ? "credit" : "debit",
    description: initial.description || "", postReference: initial.postReference || "",
    amount: initial.debit > 0 ? initial.debit : (initial.credit > 0 ? initial.credit : ""),
    comments: initial.comments || "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const valid = form.description.trim().length > 0 && Number(form.amount) > 0;
  function handleSave() {
    const amt = Number(form.amount) || 0;
    onSave({
      id: form.id, date: form.date, description: form.description, postReference: form.postReference,
      debit: form.type === "debit" ? amt : 0, credit: form.type === "credit" ? amt : 0, comments: form.comments,
    });
  }
  return (
    <Modal title={isEdit ? "Update ledger entry" : "Add ledger entry"} onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" disabled={!valid} onClick={handleSave}>Save entry</button>
    </>}>
      <div className="sms-role-toggle">
        <button type="button" className={"sms-role-btn" + (form.type === "debit" ? " active" : "")} onClick={() => setForm({ ...form, type: "debit" })}>Debit (money out)</button>
        <button type="button" className={"sms-role-btn" + (form.type === "credit" ? " active" : "")} onClick={() => setForm({ ...form, type: "credit" })}>Credit (money in)</button>
      </div>
      <Field label="Date"><input className="sms-input" type="date" value={form.date} onChange={set("date")} /></Field>
      <Field label="Description"><input className="sms-input" value={form.description} onChange={set("description")} placeholder="e.g. Electricity bill Jan to May" /></Field>
      <Field label="Post reference (source / to whom)">
        <input className="sms-input" list="ledger-suggestions" value={form.postReference} onChange={set("postReference")} placeholder="Director" />
        <datalist id="ledger-suggestions">{LEDGER_SUGGESTIONS.map((s) => <option key={s} value={s} />)}</datalist>
      </Field>
      <Field label="Amount"><input className="sms-input" type="number" value={form.amount} onChange={set("amount")} placeholder="5000" /></Field>
      <Field label="Comments (optional)"><textarea value={form.comments} onChange={set("comments")} placeholder="Any notes…" /></Field>
    </Modal>
  );
}

function StaffAttendanceForm({ initial, staff, onCancel, onSave }) {
  const isEdit = !!initial.id;
  const firstStaff = staff[0];
  const [form, setForm] = useState({
    id: initial.id || uid("satt"), staffId: initial.staffId || (firstStaff ? firstStaff.id : ""),
    date: initial.date || todayISO(), checkIn: initial.checkIn || "", checkOut: initial.checkOut || "",
    absent: initial.absent || false, lwp: initial.lwp || false,
    advance: initial.advance ?? 0, other: initial.other ?? 0, comments: initial.comments || "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  const setBool = (k) => (e) => setForm({ ...form, [k]: e.target.checked });

  const totalHrs = form.absent ? 0 : computeHoursFromTimes(form.checkIn, form.checkOut);
  const { regular, overtime } = splitRegularOvertime(totalHrs);

  function handleSave() {
    const d = new Date(form.date);
    onSave({
      ...form, month: MONTHS[d.getMonth()], year: d.getFullYear(),
      regularHrs: regular, overtimeHrs: overtime,
      advance: Number(form.advance) || 0, other: Number(form.other) || 0,
      checkIn: form.absent ? "" : form.checkIn, checkOut: form.absent ? "" : form.checkOut,
    });
  }
  const valid = !!form.staffId && !!form.date;

  return (
    <Modal title={isEdit ? "Update attendance entry" : "Add attendance entry"} wide onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" disabled={!valid} onClick={handleSave}>Save entry</button>
    </>}>
      <Field label="Staff member">
        <select className="sms-select" value={form.staffId} onChange={set("staffId")} disabled={isEdit}>
          {staff.map((s) => <option key={s.id} value={s.id}>{s.name} — {s.role}</option>)}
        </select>
      </Field>
      <Field label="Date"><input className="sms-input" type="date" value={form.date} onChange={set("date")} /></Field>

      <div className="sms-checkbox-row"><input type="checkbox" checked={form.absent} onChange={setBool("absent")} id="absentday" /><label htmlFor="absentday">Absent this day (no check-in/check-out)</label></div>

      {!form.absent && (
        <div className="sms-field-row">
          <Field label="Check-in time"><input className="sms-input" type="time" value={form.checkIn} onChange={set("checkIn")} /></Field>
          <Field label="Check-out time"><input className="sms-input" type="time" value={form.checkOut} onChange={set("checkOut")} /></Field>
        </div>
      )}
      <div className="sms-computed">
        Total hours {totalHrs} = regular {regular}h + overtime {overtime}h (regular day is capped at {REGULAR_HOURS_PER_DAY}h, rest counts as overtime)
      </div>

      <div className="sms-checkbox-row"><input type="checkbox" checked={form.lwp} onChange={setBool("lwp")} id="lwpday" /><label htmlFor="lwpday">Leave without pay (L.W.P) — deduct a day's pay</label></div>
      <div className="sms-field-row">
        <Field label="Advance (optional)"><input className="sms-input" type="number" value={form.advance} onChange={set("advance")} placeholder="0" /></Field>
        <Field label="Other deduction (optional)"><input className="sms-input" type="number" value={form.other} onChange={set("other")} placeholder="0" /></Field>
      </div>
      <Field label="Comments (optional)"><textarea value={form.comments} onChange={set("comments")} placeholder="Any notes…" /></Field>
    </Modal>
  );
}

function PettyForm({ initial, onCancel, onSave }) {
  const isEdit = !!initial.id;
  const initDate = initial.date || todayISO();
  const [form, setForm] = useState({
    id: initial.id || uid("petty"), type: initial.type || "receipt", date: initDate,
    label: initial.label || "", amount: initial.amount ?? "", comments: initial.comments || "",
  });
  const set = (k) => (e) => setForm({ ...form, [k]: e.target.value });
  function handleSave() {
    const d = new Date(form.date);
    onSave({ ...form, amount: Number(form.amount) || 0, month: MONTHS[d.getMonth()], year: d.getFullYear() });
  }
  const valid = form.label.trim().length > 0 && Number(form.amount) > 0;
  return (
    <Modal title={isEdit ? "Update petty cash entry" : "Add petty cash entry"} onClose={onCancel} footer={<>
      <button className="sms-btn secondary" onClick={onCancel}>Cancel</button>
      <button className="sms-btn" disabled={!valid} onClick={handleSave}>Save entry</button>
    </>}>
      <div className="sms-role-toggle">
        <button type="button" className={"sms-role-btn" + (form.type === "receipt" ? " active" : "")} onClick={() => setForm({ ...form, type: "receipt" })}>Received</button>
        <button type="button" className={"sms-role-btn" + (form.type === "expense" ? " active" : "")} onClick={() => setForm({ ...form, type: "expense" })}>Expense</button>
      </div>
      <Field label="Date"><input className="sms-input" type="date" value={form.date} onChange={set("date")} /></Field>
      <Field label={form.type === "receipt" ? "Source (e.g. Director, School Fund)" : "Particulars (what was it spent on)"}>
        <input className="sms-input" list="petty-suggestions" value={form.label} onChange={set("label")} placeholder={form.type === "receipt" ? "Director" : "Stationery"} />
        <datalist id="petty-suggestions">{PETTY_SUGGESTIONS.map((s) => <option key={s} value={s} />)}</datalist>
      </Field>
      <Field label="Amount"><input className="sms-input" type="number" value={form.amount} onChange={set("amount")} placeholder="25000" /></Field>
      <Field label="Comments (optional)"><textarea value={form.comments} onChange={set("comments")} placeholder="Any notes…" /></Field>
    </Modal>
  );
}
