// @ts-nocheck
import React, { useState, useMemo, useRef, useEffect } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx"; 

// ==========================================
// 🔴 CLOUD DATABASE CONNECTION (LIVE) 🔴
const SUPABASE_URL = "https://vmntpwethpuvptczrfft.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InZtbnRwd2V0aHB1dnB0Y3pyZmZ0Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzY0MzQ3ODUsImV4cCI6MjA5MjAxMDc4NX0.eAtuPwCE2WH5ReV1cXaxah6c4hxdo2pjS8d62nWIKCo";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
// ==========================================

const CO = "GATEWAY IT SOLUTIONS";
const AD = "Office#401, 4th Floor, Manjeera Majestic Commercial, JNTU Road, 1st Phase, KPHB Colony, Hyderabad-500072";
const MS = ["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar"];
const CAL_MS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const SEC_QS = ["What city were you born in?", "What is your mother's maiden name?", "What was the name of your first pet?", "What high school did you attend?"];

const getInitState = () => {
  const d = new Date();
  const mo = CAL_MS[d.getMonth()];
  const y = d.getFullYear();
  const fy = ["Jan", "Feb", "Mar"].includes(mo) ? String(y - 1) : String(y);
  return { mo, fy };
};
const { mo: initMo, fy: initFy } = getInitState();

const fyL = (y) => `${parseInt(y || 0)}-${String(parseInt(y || 0) + 1).slice(-2)}`;
const mL = (m, y) => `${m}-${["Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].includes(m) ? parseInt(y || 0) : parseInt(y || 0) + 1}`;
const f$ = (n) => Number(n || 0).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const gr = (r) => (r.basic || 0) + (r.hra || 0) + (r.conv || 0) + (r.med || 0) + (r.inc || 0) + (r.oth || 0);
const dd = (r) => (r.lop || 0) + (r.pt || 0) + (r.tds || 0) + (r.adv || 0) + (r.othD || 0);
const np = (r) => gr(r) - dd(r);
const txInc = (r) => np(r) + (r.pt || 0) + (r.tds || 0);

const fmtDate = (dStr) => {
  if (!dStr) return "-";
  if (dStr.includes("-")) {
    const parts = dStr.split("-");
    if (parts[0].length === 4) return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return dStr;
};

const TYPE_META = {
  receipt: { label: "Receipt", color: "#0F9D6D", bg: "#EAF7F1" },
  payment: { label: "Payment", color: "#B8730C", bg: "#FDF3E4" },
  expense: { label: "Expense", color: "#C23A55", bg: "#FBEAEE" },
};
const deriveFYLabel = (startDate) => {
  const d = new Date(startDate + "T00:00:00");
  const startYear = d.getFullYear();
  return `FY ${startYear}-${String(startYear + 1).slice(-2)}`;
};

const isActiveInMonth = (emp, mStr, yStr) => {
  if (emp.id === "admin") return false;
  const y = ["Jan", "Feb", "Mar"].includes(mStr) ? parseInt(yStr) + 1 : parseInt(yStr);
  const mIdx = CAL_MS.indexOf(mStr);
  const mStart = `${y}-${String(mIdx + 1).padStart(2, '0')}-01`;
  const lastDay = new Date(y, mIdx + 1, 0).getDate();
  const mEnd = `${y}-${String(mIdx + 1).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  if (emp.start && emp.start > mEnd) return false; 
  if (emp.end && emp.end < mStart) return false;   
  return true;
};

const getWD = (mStr, yStr) => {
  if (!yStr) return 0;
  const y = ["Jan", "Feb", "Mar"].includes(mStr) ? parseInt(yStr) + 1 : parseInt(yStr);
  const mIdx = CAL_MS.indexOf(mStr);
  const daysInMonth = new Date(y, mIdx + 1, 0).getDate();
  let count = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dayOfWeek = new Date(y, mIdx, d).getDay();
    if (dayOfWeek !== 0 && dayOfWeek !== 6) count++;
  }
  return count;
};

const getCalendarDays = (mStr, yStr) => {
  if (!yStr) return 30;
  const y = ["Jan", "Feb", "Mar"].includes(mStr) ? parseInt(yStr) + 1 : parseInt(yStr);
  const mIdx = CAL_MS.indexOf(mStr);
  return new Date(y, mIdx + 1, 0).getDate();
};

const getEmptyAtt = () => {
  const a = {};
  MS.forEach((m) => { a[m] = { present: null, leave: null, bal: null, lop: null, holiday: null, comments: "" }; });
  return a;
};

const exportExcel = (rows, fn) => {
  const ws = XLSX.utils.aoa_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Report");
  XLSX.writeFile(wb, `${fn}.xlsx`);
};

// Generic HTML templates for PDF rendering
const buildReportPdf = (title, subtitle, head, rows) => `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title><style>body{font-family:'Segoe UI',sans-serif;background:#fff;padding:20px;color:#333;margin:0}.hdr{text-align:center;border-bottom:2px solid #185FA5;padding-bottom:20px;margin-bottom:25px}.hdr h2{margin:0 0 5px;color:#185FA5;font-size:24px}.hdr p{margin:0;font-size:12px;color:#666}.hdr h3{margin:15px 0 0;font-size:18px}table{width:100%;border-collapse:collapse;font-size:11px}th{background:#1a1a2e;color:#fff;text-align:left;padding:8px;border:1px solid #1a1a2e}td{padding:8px;border:1px solid #e2e8f0}.print-btn{display:block;width:200px;margin:0 auto 20px;padding:12px;text-align:center;background:#1D9E75;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:600}@media print{.print-btn{display:none}}</style></head><body><button class="print-btn" onclick="window.print()">🖨️ Print / Save PDF</button><div class="hdr"><h2>${CO}</h2><p>${AD}</p><h3>${title}</h3><p><b>${subtitle}</b></p></div><table><thead><tr>${head.map(h=>`<th>${h}</th>`).join("")}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(c=>`<td>${c!==undefined&&c!==null?c:""}</td>`).join("")}</tr>`).join("")}</tbody></table></body></html>`;

const generateVoucherPdf = (txn) => {
  const title = txn.type === "receipt" ? "RECEIPT VOUCHER" : txn.type === "payment" ? "PAYMENT VOUCHER" : "EXPENSE VOUCHER";
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${title}</title><style>body{font-family:'Segoe UI',sans-serif;background:#eee;padding:40px;color:#333;margin:0}.box{max-width:800px;margin:auto;background:#fff;border:2px solid #1a1a2e;padding:40px;border-radius:8px}.hdr{text-align:center;border-bottom:2px solid #eee;padding-bottom:20px;margin-bottom:30px}.hdr h2{margin:0 0 5px;color:#185FA5;font-size:26px}.hdr p{margin:0;font-size:13px;color:#666}.v-title{text-align:center;margin-bottom:40px}.v-title span{background:#1a1a2e;color:#fff;padding:10px 25px;font-weight:bold;border-radius:4px;font-size:18px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:40px;font-size:15px}.row{display:flex;margin-bottom:15px;border-bottom:1px dashed #ddd;padding-bottom:8px}.lbl{width:160px;font-weight:700;color:#444}.val{font-weight:600;flex:1}.amt-box{background:#f8f9fa;border:1px solid #ccc;padding:20px;text-align:center;border-radius:6px;margin-bottom:60px}.amt-val{font-size:32px;font-weight:900;color:#1D9E75;margin-top:5px}.sig-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:30px;margin-top:80px;text-align:center}.sig-line{border-top:1px solid #222;padding-top:10px;font-size:13px;font-weight:bold}.print-btn{display:block;width:250px;margin:0 auto 20px;padding:14px;background:#1D9E75;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:600}@media print{body{background:#fff;padding:0}.box{border:2px solid #000;padding:30px}.print-btn{display:none}}</style></head><body><button class="print-btn" onclick="window.print()">🖨️ Print Voucher</button><div class="box"><div class="hdr"><h2>${CO}</h2><p>${AD}</p></div><div class="v-title"><span>${title}</span></div><div class="grid"><div class="row"><div class="lbl">Voucher No.</div><div class="val">: ${txn.voucher_no||"N/A"}</div></div><div class="row"><div class="lbl">Date</div><div class="val">: ${fmtDate(txn.txn_date)}</div></div><div class="row" style="grid-column:1/-1"><div class="lbl">${txn.type==="receipt"?"Received From":"Paid To"}</div><div class="val">: ${txn.party_name}</div></div><div class="row" style="grid-column:1/-1"><div class="lbl">Description</div><div class="val">: ${txn.description||"-"}</div></div></div><div class="amt-box"><div style="font-size:14px;color:#666;font-weight:bold">Total Amount</div><div class="amt-val">₹ ${f$(txn.amount)}</div></div><div class="sig-grid"><div class="sig-line">Prepared By</div><div class="sig-line">Authorized Signatory</div><div class="sig-line">Receiver's Signature</div></div></div></body></html>`;
};

const generatePayslipsHtml = (slipsData, fy) => {
  let htmlContent = slipsData.map((data) => {
    const { emp, mo, entries, att } = data;
    const totals = entries.reduce((acc, curr) => ({ basic: acc.basic + Number(curr.basic||0), hra: acc.hra + Number(curr.hra||0), conv: acc.conv + Number(curr.conv||0), med: acc.med + Number(curr.med||0), inc: acc.inc + Number(curr.inc||0), oth: acc.oth + Number(curr.oth||0), lop: acc.lop + Number(curr.lop||0), adv: acc.adv + Number(curr.adv||0), pt: acc.pt + Number(curr.pt||0), tds: acc.tds + Number(curr.tds||0), othD: acc.othD + Number(curr.othD||0), note: acc.note ? (curr.note ? acc.note + " | " + curr.note : acc.note) : curr.note || "" }), { basic: 0, hra: 0, conv: 0, med: 0, inc: 0, oth: 0, lop: 0, adv: 0, pt: 0, tds: 0, othD: 0, note: "" });
    const a = att || {}; const wd = getWD(mo, fy);
    const present = a.present !== undefined && a.present !== null ? a.present : "-";
    const leave = a.leave !== undefined && a.leave !== null ? a.leave : "-";
    const bal = a.bal !== undefined && a.bal !== null ? a.bal : "-";
    const g = gr(totals); const d = dd(totals); const n = g - d;
    const displayYear = ["Jan", "Feb", "Mar"].includes(mo) ? parseInt(fy) + 1 : parseInt(fy);

    return `<div class="slip-container" style="page-break-after:always;margin-bottom:40px"><div class="box" style="max-width:800px;margin:auto;background:#fff;padding:40px;border-radius:8px;border-top:5px solid #185FA5"><div class="hdr" style="display:flex;justify-content:space-between;border-bottom:2px solid #eee;padding-bottom:20px;margin-bottom:25px"><div class="h-l"><img src="/logo.png" style="max-height:50px;margin-bottom:5px"/><p style="margin:0;font-size:12px;color:#666">${AD}</p></div><div class="h-r" style="text-align:right"><h2 style="margin:0 0 5px;color:#185FA5;font-size:20px">PAYSLIP</h2><p style="margin:0;font-weight:600">${mo.toUpperCase()} ${displayYear}</p></div></div><div class="grid" style="display:grid;grid-template-columns:1fr 1fr;gap:15px;background:#f8fafc;padding:20px;border-radius:6px;border:1px solid #e2e8f0;margin-bottom:25px;font-size:12px"><div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Employee Name</div><div class="val" style="font-weight:600">: ${emp.name}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Employee ID</div><div class="val" style="font-weight:600">: ${emp.id}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Designation</div><div class="val" style="font-weight:600">: ${emp.desig || "-"}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Date of Joining</div><div class="val" style="font-weight:600">: ${fmtDate(emp.start)}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">PAN</div><div class="val" style="font-weight:600">: ${emp.pan || "-"}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Bank A/C</div><div class="val" style="font-weight:600">: ${emp.bank || "-"}</div></div></div><div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Total Working Days</div><div class="val" style="font-weight:600">: ${wd}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Days Worked</div><div class="val" style="font-weight:600">: ${present}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Leave Availed</div><div class="val" style="font-weight:600">: ${leave}</div></div><div class="row" style="display:flex;margin-bottom:6px"><div class="lbl" style="width:135px;font-weight:600;color:#555">Leave Balance</div><div class="val" style="font-weight:600">: ${emp.leavePolicy === 'No' ? 'NA' : bal}</div></div></div></div><table style="width:100%;border-collapse:collapse;margin-bottom:25px;font-size:12px"><thead><tr><th style="background:#1a1a2e;color:#fff;text-align:left;padding:10px;width:35%">EARNINGS</th><th style="background:#1a1a2e;color:#fff;text-align:right;padding:10px;width:15%">AMOUNT</th><th style="background:#1a1a2e;color:#fff;text-align:left;padding:10px;width:35%">DEDUCTIONS</th><th style="background:#1a1a2e;color:#fff;text-align:right;padding:10px;width:15%">AMOUNT</th></tr></thead><tbody><tr><td style="padding:10px;border:1px solid #e2e8f0">Basic Salary</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.basic)}</td><td style="padding:10px;border:1px solid #e2e8f0">Profession Tax</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.pt)}</td></tr><tr><td style="padding:10px;border:1px solid #e2e8f0">House Rent Allowance</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.hra)}</td><td style="padding:10px;border:1px solid #e2e8f0">Tax Deducted at Source (TDS)</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.tds)}</td></tr><tr><td style="padding:10px;border:1px solid #e2e8f0">Conveyance</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.conv)}</td><td style="padding:10px;border:1px solid #e2e8f0">Staff Advance</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.adv)}</td></tr><tr><td style="padding:10px;border:1px solid #e2e8f0">Medical Allowance</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.med)}</td><td style="padding:10px;border:1px solid #e2e8f0">Loss of Pay (LOP)</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.lop)}</td></tr><tr><td style="padding:10px;border:1px solid #e2e8f0">Incentives</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.inc)}</td><td style="padding:10px;border:1px solid #e2e8f0">Other Deductions</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.othD)}</td></tr><tr><td style="padding:10px;border:1px solid #e2e8f0">Arrears & Others</td><td style="padding:10px;border:1px solid #e2e8f0;text-align:right;font-weight:600">${f$(totals.oth)}</td><td style="padding:10px;border:1px solid #e2e8f0"></td><td style="padding:10px;border:1px solid #e2e8f0"></td></tr><tr style="font-weight:700;background:#f8fafc"><td style="padding:10px;border:1px solid #e2e8f0;border-top:2px solid #cbd5e1">Gross Earnings</td><td style="padding:10px;border:1px solid #e2e8f0;border-top:2px solid #cbd5e1;text-align:right">${f$(g)}</td><td style="padding:10px;border:1px solid #e2e8f0;border-top:2px solid #cbd5e1">Total Deductions</td><td style="padding:10px;border:1px solid #e2e8f0;border-top:2px solid #cbd5e1;text-align:right">${f$(d)}</td></tr></tbody></table><div class="net" style="background:#e8f5e9;border-left:5px solid #1D9E75;padding:15px 20px;display:flex;justify-content:space-between;align-items:center;border-radius:4px;margin-bottom:30px"><div style="font-size:14px;font-weight:700;color:#1D9E75;margin:0;text-transform:uppercase">Net Pay<br><span style="font-size:10px;color:#555;text-transform:none;font-weight:normal">(Gross Earnings - Total Deductions)</span></div><div style="font-size:22px;font-weight:800;color:#111;margin:0">₹ ${f$(n)}</div></div>${totals.note ? `<div style="font-size:12px;color:#444;background:#fff8e1;padding:10px 15px;border-left:3px solid #ffc107;border-radius:4px;margin-bottom:20px"><b>Note:</b> ${totals.note}</div>` : ""}<div class="sig" style="display:flex;justify-content:space-between;margin-top:50px"><div style="font-size:12px;font-weight:600;color:#444;border-top:1px solid #ccc;padding-top:10px;width:180px;text-align:center">Employer Signature</div><div style="font-size:12px;font-weight:600;color:#444;border-top:1px solid #ccc;padding-top:10px;width:180px;text-align:center">Employee Signature</div></div><div class="ftr" style="text-align:center;margin-top:30px;font-size:10px;color:#888;border-top:1px solid #eee;padding-top:15px">This is a computer-generated document. No physical signature is required.</div></div></div>`;
  }).join('');

  return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>GITS Bulk Payslips</title><style>body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#eee;padding:20px;color:#333;margin:0}.print-btn{display:block;width:250px;margin:0 auto 20px;padding:12px;text-align:center;background:#1D9E75;color:#fff;border:none;border-radius:4px;cursor:pointer;font-weight:600;font-size:15px}@media print{body{background:#fff;padding:0}.box{box-shadow:none;border:none;padding:0}.print-btn{display:none}.slip-container{margin-bottom:0}}</style></head><body><button class="print-btn" onclick="window.print()">🖨️ Save All as PDF</button>${htmlContent}</body></html>`;
};

const getDiffText = (oldObj, newObj, keys) => {
  let diffs = [];
  keys.forEach(k => {
    if (String(oldObj[k] || "") !== String(newObj[k] || "")) {
      diffs.push(`${k}: '${oldObj[k] || ""}' -> '${newObj[k] || ""}'`);
    }
  });
  return diffs.join(" | ");
};

export default function App() {
  const [dbLoaded, setDbLoaded] = useState(false);
  const [sysWarn, setSysWarn] = useState(null);

  const [emps, setEmps] = useState([]);
  const [pay, setPay] = useState({});
  const [att, setAtt] = useState({});
  const [dailyAtt, setDailyAtt] = useState([]); 
  const [auditLogs, setAuditLogs] = useState([]);

  const [cbFYs, setCbFYs] = useState([]);
  const [cbTxns, setCbTxns] = useState([]);
  const [cbSelFY, setCbSelFY] = useState("");
  const [cbType, setCbType] = useState("receipt");
  const [cbParty, setCbParty] = useState("");
  const [cbDate, setCbDate] = useState(new Date().toISOString().split('T')[0]);
  const [cbAmount, setCbAmount] = useState("");
  const [cbDesc, setCbDesc] = useState("");
  const [cbTabState, setCbTabState] = useState("all");
  const [cbSearch, setCbSearch] = useState("");
  const [showAddCbFy, setShowAddCbFy] = useState(false);
  const [newCbFyStart, setNewCbFyStart] = useState("");
  const [newCbFyEnd, setNewCbFyEnd] = useState("");
  const [editCbTxn, setEditCbTxn] = useState(null); 

  const [ses, setSes] = useState(null);
  const [tab, setTab] = useState("dashboard");
  const [fy, setFy] = useState(initFy);
  const [mo, setMo] = useState(initMo);
  const [syncMo, setSyncMo] = useState(initMo); 
  const [attSubTab, setAttSubTab] = useState("daily"); 
  const [attLedgerMo, setAttLedgerMo] = useState(""); 
  const [slip, setSlip] = useState(null);

  const [lEmp, setLEmp] = useState("");
  const [pEmp, setPEmp] = useState("");
  const [pMo, setPMo] = useState(initMo);
  const [pFy, setPFy] = useState(initFy);

  useEffect(() => {
    const initDb = async () => {
      let errs = [];
      let loadedEmps = []; let loadedLedData = []; let loadedAttData = []; let loadedDailyData = []; let loadedAuditData = [];
      let loadedCbFYs = []; let loadedCbTxns = [];

      try {
        const { data, error } = await supabase.from("gits_employees").select("*");
        if (error) throw error; loadedEmps = data || [];
      } catch (e) { errs.push(`Employees: ${e.message}`); }

      try {
        const { data, error } = await supabase.from("gits_ledger").select("*");
        if (error) throw error; loadedLedData = data || [];
      } catch (e) { errs.push(`Ledger: ${e.message}`); }

      try {
        const { data, error } = await supabase.from("gits_attendance").select("*");
        if (error) throw error; loadedAttData = data || [];
      } catch (e) { errs.push(`Attendance: ${e.message}`); }

      try {
        const { data, error } = await supabase.from("gits_daily_attendance").select("*");
        if (error && error.code !== "42P01") throw error; loadedDailyData = data || [];
      } catch (e) { }
      
      try {
        const { data, error } = await supabase.from("gits_audit_log").select("*").order("created_at", { ascending: false });
        if (error && error.code !== "42P01") throw error; loadedAuditData = data || [];
      } catch (e) { }

      try {
        const { data, error } = await supabase.from("financial_years").select("*").order("start_date", { ascending: false });
        if (error && error.code !== "42P01") throw error; loadedCbFYs = data || [];
      } catch (e) { errs.push(`Cashbook FYs: ${e.message}`); }

      try {
        const { data, error } = await supabase.from("transactions").select("*").order("txn_date", { ascending: false });
        if (error && error.code !== "42P01") throw error; loadedCbTxns = data || [];
      } catch (e) { errs.push(`Cashbook Txns: ${e.message}`); }

      if (loadedEmps.length === 0) {
        loadedEmps.push({ id: "admin", name: "System Admin", pwd: "admin123", status: "Active", leave_policy: "Yes" });
      }

      const formattedEmps = loadedEmps.map((e) => ({
        id: e.id, name: e.name, desig: e.desig, pan: e.pan || "", cat: e.cat, basic: e.basic,
        phone: e.phone, email: e.email, pwd: e.pwd, start: e.start_date, end: e.end_date,
        status: e.status, comments: e.comments, bank: e.bank, driveLink: e.drive_link, sec_q: e.sec_q, sec_a: e.sec_a, leavePolicy: e.leave_policy || "Yes"
      }));

      const formattedPay = {};
      loadedLedData.forEach((r) => {
        if (!formattedPay[r.emp_id]) formattedPay[r.emp_id] = {};
        if (!formattedPay[r.emp_id][r.fy]) formattedPay[r.emp_id][r.fy] = [];
        formattedPay[r.emp_id][r.fy].push({ db_id: r.id, m: r.mo, t: r.t, basic: r.basic, hra: r.hra, conv: r.conv, med: r.med, inc: r.inc, oth: r.oth, lop: r.lop, adv: r.adv, pt: r.pt, tds: r.tds, othD: r.othd, note: r.note });
      });

      const formattedAtt = {};
      loadedAttData.forEach((r) => {
        if (!formattedAtt[r.emp_id]) formattedAtt[r.emp_id] = {};
        if (!formattedAtt[r.emp_id][r.fy]) formattedAtt[r.emp_id][r.fy] = getEmptyAtt();
        formattedAtt[r.emp_id][r.fy][r.mo] = { present: r.present, leave: r.leave, bal: r.bal, lop: r.lop, holiday: r.holiday, comments: r.comments };
      });

      let maxScore = -1;
      let latestMo = null;
      let latestFy = null;

      const d = new Date();
      const curMoStr = CAL_MS[d.getMonth()];
      const curYStr = d.getFullYear();
      const curFyStr = ["Jan", "Feb", "Mar"].includes(curMoStr) ? String(curYStr - 1) : String(curYStr);
      const curScoreLimit = parseInt(curFyStr) * 100 + MS.indexOf(curMoStr) + 1;

      const evalData = (arr, isLedger) => {
        arr.forEach(r => {
          const hasData = isLedger 
            ? (r.basic > 0 || r.inc > 0 || r.oth > 0 || r.lop > 0 || (r.note && r.note.trim() !== ""))
            : (r.present !== null || r.leave !== null || r.holiday !== null || r.lop !== null || (r.comments && r.comments.trim() !== ""));
          if (r.fy && r.mo && hasData) {
            const mIdx = MS.indexOf(r.mo);
            if (mIdx !== -1) {
              const score = parseInt(r.fy) * 100 + mIdx;
              if (score > maxScore && score <= curScoreLimit) {
                maxScore = score;
                latestMo = r.mo;
                latestFy = r.fy;
              }
            }
          }
        });
      };
      
      evalData(loadedLedData, true);
      evalData(loadedAttData, false);

      if (latestMo && latestFy) {
        setMo(latestMo);
        setFy(latestFy);
        setPMo(latestMo);
        setPFy(latestFy);
      }

      setEmps(formattedEmps);
      setPay(formattedPay);
      setAtt(formattedAtt);
      setDailyAtt(loadedDailyData);
      setAuditLogs(loadedAuditData);
      setCbFYs(loadedCbFYs);
      setCbTxns(loadedCbTxns);
      if (loadedCbFYs.length > 0) setCbSelFY(loadedCbFYs[0].id);

      if (errs.length > 0) setSysWarn(errs.join(" | "));
      setDbLoaded(true);
    };
    initDb();
  }, []);

  const [showAddEmp, setShowAddEmp] = useState(false);
  const [nE, setNE] = useState({ id: "", name: "", desig: "", pan: "", cat: "Onshore", basic: "", phone: "", email: "", pwd: "", start: "", end: "", status: "Active", bank: "", comments: "", driveLink: "", leavePolicy: "Yes" });
  const [editEmp, setEditEmp] = useState(null);
  const [editData, setEditData] = useState({});

  const [showAddEntry, setShowAddEntry] = useState(false);
  const [editLedger, setEditLedger] = useState(null);
  const [nEn, setNEn] = useState({ m: "Apr", t: "s", basic: 0, hra: 0, conv: 800, med: 1500, inc: 0, oth: 0, lop: 0, adv: 0, pt: 200, tds: 0, othD: 0, note: "" });

  const [showBulk, setShowBulk] = useState(false);
  const [bulkData, setBulkData] = useState({});
  const [showOffCycle, setShowOffCycle] = useState(false);
  const [offCycleData, setOffCycleData] = useState({ empId: "", basic: 0, hra: 0, conv: 0, med: 0, inc: 0, oth: 0, lop: 0, adv: 0, pt: 0, tds: 0, othD: 0, note: "" });

  const [markDate, setMarkDate] = useState(new Date().toISOString().split('T')[0]); 
  const [dailyForm, setDailyForm] = useState({}); 

  const idR = useRef(""), pwR = useRef("");
  const [forgotStep, setForgotStep] = useState(0);
  const [forgotId, setForgotId] = useState("");
  const [forgotUser, setForgotUser] = useState(null);
  const [secAnsInput, setSecAnsInput] = useState("");
  const [newForgotPwd, setNewForgotPwd] = useState("");

  const [curPwd, setCurPwd] = useState("");
  const [newPwd, setNewPwd] = useState("");
  const [confPwd, setConfPwd] = useState("");
  const [secQ, setSecQ] = useState("");
  const [secA, setSecA] = useState("");

  const myE = ses?.role === "e" ? emps.find((e) => e.id === ses.id) : ses?.role === "a" ? emps.find((e) => e.id === "admin") : null;

  const dTot = useMemo(() => {
    let g = 0, d = 0, n = 0;
    const targetEmps = ses?.role === "a" ? emps : myE ? [myE] : [];
    targetEmps.forEach((e) =>
      (pay[e.id]?.[fy] || []).filter((r) => r.m === mo).forEach((r) => {
        g += gr(r);
        d += dd(r);
        n += np(r);
      })
    );
    return { g, d, n };
  }, [emps, pay, fy, mo, ses, myE]);

  useEffect(() => {
    const form = {};
    emps.forEach(e => {
      if (e.id !== 'admin' && e.status === 'Active') {
        const existing = dailyAtt.find(r => r.emp_id === e.id && r.date === markDate);
        form[e.id] = existing ? { status: existing.status, reason: existing.reason || "" } : { status: "Present", reason: "" };
      }
    });
    setDailyForm(form);
  }, [markDate, dailyAtt, emps]);

  const logAction = async (module, desc) => {
    if (!ses || ses.role !== "a") return;
    const newLog = { admin_id: ses.id, module, description: desc, created_at: new Date().toISOString() };
    setAuditLogs(prev => [newLog, ...prev]);
    try { await supabase.from("gits_audit_log").insert(newLog); } catch (e) { console.error("Audit log failed to save remotely", e); }
  };

  const login = () => {
    const id = idR.current.trim(), pw = pwR.current.trim();
    const dbAdmin = emps.find((x) => x.id === "admin");
    if (id === "admin" && (dbAdmin ? pw === dbAdmin.pwd : pw === "admin123")) {
      setSes({ role: "a", id: "admin" });
      setTab("dashboard");
    } else {
      const e = emps.find((x) => x.id === id && x.pwd === pw);
      if (e) {
        setSes({ role: "e", id: e.id });
        setLEmp(e.id);
        setPEmp(e.id);
        setTab("dashboard");
      } else alert("Invalid Credentials");
    }
  };

  const handleForgotNext = async () => {
    if (forgotStep === 0) {
      const e = emps.find((x) => x.id === forgotId.trim() || x.email === forgotId.trim());
      if (!e) return alert("User not found.");
      if (!e.sec_q || !e.sec_a) return alert("You have not set up a security question. Please contact your Admin to reset your password.");
      setForgotUser(e);
      setForgotStep(1);
    } else if (forgotStep === 1) {
      if (secAnsInput.toLowerCase().trim() !== forgotUser.sec_a.toLowerCase().trim()) return alert("Incorrect answer.");
      setForgotStep(2);
    } else if (forgotStep === 2) {
      if (!newForgotPwd) return alert("Enter a new password");
      await supabase.from("gits_employees").update({ pwd: newForgotPwd }).eq("id", forgotUser.id);
      setEmps((p) => p.map((x) => (x.id === forgotUser.id ? { ...x, pwd: newForgotPwd } : x)));
      alert("Password reset successfully! You can now log in.");
      setForgotStep(0);
      setForgotId("");
      setForgotUser(null);
      setSecAnsInput("");
      setNewForgotPwd("");
    }
  };

  const handleUpdateProfile = async () => {
    const userId = ses.id;
    let pwdToSave = undefined;

    if (curPwd || newPwd || confPwd) {
      if (!curPwd || !newPwd || !confPwd) return alert("To change password, fill all password fields.");
      if (newPwd !== confPwd) return alert("New passwords do not match.");
      if (ses.role === "a") {
        const dbAdmin = emps.find((x) => x.id === "admin");
        if (!dbAdmin && curPwd !== "admin123") return alert("Incorrect current password.");
        if (dbAdmin && curPwd !== dbAdmin.pwd) return alert("Incorrect current password.");
      } else {
        if (curPwd !== myE.pwd) return alert("Incorrect current password.");
      }
      pwdToSave = newPwd;
    }

    const payload = {};
    if (pwdToSave) payload.pwd = pwdToSave;
    if (secQ) payload.sec_q = secQ;
    if (secA) payload.sec_a = secA;

    if (ses.role === "a") {
      const dbAdmin = emps.find((x) => x.id === "admin");
      if (!dbAdmin) {
        await supabase.from("gits_employees").insert({ id: "admin", name: "System Admin", status: "Active", ...payload });
        setEmps([...emps, { id: "admin", name: "System Admin", status: "Active", ...payload }]);
      } else {
        await supabase.from("gits_employees").update(payload).eq("id", "admin");
        setEmps((p) => p.map((x) => (x.id === "admin" ? { ...x, ...payload } : x)));
      }
    } else {
      await supabase.from("gits_employees").update(payload).eq("id", userId);
      setEmps((p) => p.map((x) => (x.id === userId ? { ...x, ...payload } : x)));
    }
    alert("Profile updated successfully!");
    setCurPwd(""); setNewPwd(""); setConfPwd("");
  };

  const addEmployee = async () => {
    if (!nE.id || !nE.name) return alert("ID and Name are required");
    const newEmp = { ...nE, basic: nE.basic === "" ? 0 : Number(nE.basic) };
    await supabase.from("gits_employees").insert({
      id: newEmp.id, name: newEmp.name, desig: newEmp.desig, pan: newEmp.pan, cat: newEmp.cat, basic: newEmp.basic,
      phone: newEmp.phone, email: newEmp.email, pwd: newEmp.pwd, start_date: newEmp.start, end_date: newEmp.end,
      status: newEmp.status, comments: newEmp.comments, bank: newEmp.bank, drive_link: newEmp.driveLink, leave_policy: newEmp.leavePolicy
    });
    setEmps([...emps, newEmp]);
    setPay({ ...pay, [nE.id]: {} });
    setAtt({ ...att, [nE.id]: {} });
    logAction("Employees", `Added new employee ${nE.name} (${nE.id})`);
    setShowAddEmp(false);
    setNE({ id: "", name: "", desig: "", pan: "", cat: "Onshore", basic: "", phone: "", email: "", pwd: "", start: "", end: "", status: "Active", bank: "", comments: "", driveLink: "", leavePolicy: "Yes" });
  };

  const saveEditEmployee = async () => {
    const oldEmp = emps.find(e => e.id === editEmp);
    const combinedComments = editData.reason ? editData.reason : editData.comments;
    const payload = {
      phone: editData.phone, email: editData.email, start_date: editData.start, end_date: editData.end,
      status: editData.status, comments: combinedComments, bank: editData.bank, drive_link: editData.driveLink,
      leave_policy: editData.leavePolicy, pan: editData.pan
    };
    if (editData.adminForcePwd) payload.pwd = editData.adminForcePwd;

    await supabase.from("gits_employees").update(payload).eq("id", editEmp);
    setEmps((p) => p.map((x) => x.id === editEmp ? { ...x, phone: editData.phone, email: editData.email, start: editData.start, end: editData.end, status: editData.status, comments: combinedComments, bank: editData.bank, driveLink: editData.driveLink, leavePolicy: editData.leavePolicy, pan: editData.pan, ...(editData.adminForcePwd ? { pwd: editData.adminForcePwd } : {}) } : x));
    
    const diffKeys = ['basic', 'desig', 'status', 'leavePolicy', 'pan'];
    const diffTxt = getDiffText(oldEmp, editData, diffKeys);
    if(diffTxt) logAction("Employees", `Updated ${oldEmp.name} (${oldEmp.id}): ${diffTxt}`);
    setEditEmp(null);
  };

  const delEmployee = async (id) => {
    if (confirm("Are you sure you want to delete this employee?")) {
      await supabase.from("gits_employees").delete().eq("id", id);
      setEmps((p) => p.filter((x) => x.id !== id));
      logAction("Employees", `Deleted employee ID: ${id}`);
    }
  };

  const addLedgerEntry = async () => {
    if (!pEmp) return alert("Select Employee First");
    const e = { ...nEn, basic: +nEn.basic || 0, hra: +nEn.hra || 0, conv: +nEn.conv || 0, med: +nEn.med || 0, inc: +nEn.inc || 0, oth: +nEn.oth || 0, lop: +nEn.lop || 0, adv: +nEn.adv || 0, pt: +nEn.pt || 0, tds: +nEn.tds || 0, othD: +nEn.othD || 0 };
    if (editLedger) {
      const oldEntry = pay[pEmp][fy][editLedger.idx];
      await supabase.from("gits_ledger").update({ mo: e.m, t: e.t, basic: e.basic, hra: e.hra, conv: e.conv, med: e.med, inc: e.inc, oth: e.oth, lop: e.lop, adv: e.adv, pt: e.pt, tds: e.tds, othd: e.othD, note: e.note }).eq("id", editLedger.db_id);
      const updatedArr = [...(pay[pEmp]?.[fy] || [])];
      updatedArr[editLedger.idx] = { ...e, db_id: editLedger.db_id };
      setPay({ ...pay, [pEmp]: { ...pay[pEmp], [fy]: updatedArr } });
      const diffKeys = ['basic', 'hra', 'conv', 'med', 'inc', 'oth', 'lop', 'adv', 'pt', 'tds', 'othD'];
      const diffTxt = getDiffText(oldEntry, e, diffKeys);
      if(diffTxt) logAction("Ledger", `Edited payroll for ${pEmp} in ${e.m}: ${diffTxt}`);
    } else {
      const { data } = await supabase.from("gits_ledger").insert({ emp_id: pEmp, fy: fy, mo: e.m, t: e.t, basic: e.basic, hra: e.hra, conv: e.conv, med: e.med, inc: e.inc, oth: e.oth, lop: e.lop, adv: e.adv, pt: e.pt, tds: e.tds, othd: e.othD, note: e.note }).select();
      if (data) setPay({ ...pay, [pEmp]: { ...pay[pEmp], [fy]: [...(pay[pEmp]?.[fy] || []), { ...e, db_id: data[0].id }] } });
      logAction("Ledger", `Added manual payroll entry for ${pEmp} in ${e.m} (Net: ${np(e)})`);
    }
    setShowAddEntry(false);
    setEditLedger(null);
  };

  const delLedgerEntry = async (eid, idx, db_id) => {
    if (confirm("Delete this payroll entry?")) {
      if (db_id) await supabase.from("gits_ledger").delete().eq("id", db_id);
      setPay({ ...pay, [eid]: { ...pay[eid], [fy]: pay[eid][fy].filter((_, i) => i !== idx) } });
      logAction("Ledger", `Deleted payroll entry for ${eid}`);
    }
  };

  const getLastPay = (eid) => {
    const ks = Object.keys(pay[eid] || {}).sort();
    for (let i = ks.length - 1; i >= 0; i--) {
      const arr = pay[eid][ks[i]].filter((r) => r.t === "s" && (r.basic > 0 || r.hra > 0));
      if (arr.length) return arr[arr.length - 1]; 
    }
    return null; 
  };

  const getEmpDefaults = (eid) => {
    const e = emps.find((x) => x.id === eid);
    const l = getLastPay(eid);
    if (l) return { basic: l.basic || 0, hra: l.hra || 0, conv: l.conv || 0, med: l.med || 0, inc: 0, oth: 0, lop: 0, adv: 0, pt: l.pt || 0, tds: l.tds || 0, othD: 0, note: "" };
    const gross = Number(e?.basic || 0);
    let basic = 0, hra = 0, conv = 800, med = 1500;
    if (gross > 0) { basic = Math.round(gross * 0.6839); hra = gross - basic - conv - med; if (hra < 0) hra = 0; } else { conv = 0; med = 0; }
    return { basic, hra, conv, med, inc: 0, oth: 0, lop: 0, adv: 0, pt: 200, tds: 0, othD: 0, note: "" };
  };

  const openBulkPayroll = () => {
    const defaults = {};
    emps.filter((e) => isActiveInMonth(e, mo, fy) && e.id !== "admin").forEach((emp) => {
      const d = getEmpDefaults(emp.id);
      const a = att[emp.id]?.[fy]?.[mo];
      const calDays = getCalendarDays(mo, fy);
      const corePay = d.basic + d.hra + d.conv + d.med;
      const lopDays = Number(a?.lop || 0);
      let lopAmt = 0;
      if (lopDays > 0) lopAmt = Math.round((corePay / calDays) * lopDays);
      let encashAmt = 0; let autoNote = "";
      if (a?.comments && a.comments.toLowerCase().includes("encash")) {
          const match = a.comments.match(/(\d+)\s*leave/i);
          if (match) { const eDays = Number(match[1]); encashAmt = Math.round((corePay / 30) * eDays); autoNote = `${eDays} leaves encashed`; }
      }
      if (lopDays > 0) autoNote = autoNote ? `${autoNote} | ${lopDays} days LOP` : `${lopDays} days LOP`;
      defaults[emp.id] = { ...d, lop: lopAmt, oth: encashAmt, note: autoNote };
    });
    setBulkData(defaults); setShowBulk(true); setShowOffCycle(false);
  };

  const saveBulkPayroll = async () => {
    const nextPay = { ...pay };
    const dbInserts = [];
    Object.keys(bulkData).forEach((eid) => {
      const d = bulkData[eid];
      const existing = nextPay[eid]?.[fy] || [];
      if (!existing.find((r) => r.m === mo && r.t === "s")) {
        const entry = { m: mo, t: "s", basic: +d.basic || 0, hra: +d.hra || 0, conv: +d.conv || 0, med: +d.med || 0, inc: +d.inc || 0, oth: +d.oth || 0, lop: +d.lop || 0, adv: +d.adv || 0, pt: +d.pt || 0, tds: +d.tds || 0, othD: +d.othD || 0, note: d.note || "" };
        dbInserts.push({ emp_id: eid, fy: fy, mo: mo, t: "s", basic: entry.basic, hra: entry.hra, conv: entry.conv, med: entry.med, inc: entry.inc, oth: entry.oth, lop: entry.lop, adv: entry.adv, pt: entry.pt, tds: entry.tds, othd: entry.othD, note: entry.note });
      }
    });
    if (dbInserts.length) {
      const { data } = await supabase.from("gits_ledger").insert(dbInserts).select();
      if (data) {
        data.forEach((r) => {
          if (!nextPay[r.emp_id]) nextPay[r.emp_id] = {};
          if (!nextPay[r.emp_id][r.fy]) nextPay[r.emp_id][r.fy] = [];
          nextPay[r.emp_id][r.fy].push({ db_id: r.id, m: r.mo, t: r.t, basic: r.basic, hra: r.hra, conv: r.conv, med: r.med, inc: r.inc, oth: r.oth, lop: r.lop, adv: r.adv, pt: r.pt, tds: r.tds, othD: r.othd, note: r.note });
        });
      }
    }
    setPay(nextPay); setShowBulk(false);
    logAction("Ledger", `Processed Bulk Monthly Payroll for ${mo} FY ${fyL(fy)}`);
    alert(`Bulk Payroll processed for ${mL(mo, fy)}`);
  };

  const saveOffCycle = async () => {
    if (!offCycleData.empId) return alert("Select an employee");
    const entry = { m: mo, t: "s", basic: +offCycleData.basic || 0, hra: +offCycleData.hra || 0, conv: +offCycleData.conv || 0, med: +offCycleData.med || 0, inc: +offCycleData.inc || 0, oth: +offCycleData.oth || 0, lop: +offCycleData.lop || 0, adv: +offCycleData.adv || 0, pt: +offCycleData.pt || 0, tds: +offCycleData.tds || 0, othD: +offCycleData.othD || 0, note: offCycleData.note || "Off-Cycle" };
    const { data } = await supabase.from("gits_ledger").insert({ emp_id: offCycleData.empId, fy: fy, mo: mo, t: "s", basic: entry.basic, hra: entry.hra, conv: entry.conv, med: entry.med, inc: entry.inc, oth: entry.oth, lop: entry.lop, adv: entry.adv, pt: entry.pt, tds: entry.tds, othd: entry.othD, note: entry.note }).select();
    if (data) setPay((prev) => ({ ...prev, [offCycleData.empId]: { ...prev[offCycleData.empId], [fy]: [...(prev[offCycleData.empId]?.[fy] || []), { ...entry, db_id: data[0].id }] } }));
    logAction("Ledger", `Processed Off-Cycle Payroll for ${offCycleData.empId} in ${mo} FY ${fyL(fy)} (Net: ${np(entry)})`);
    setShowOffCycle(false);
    setOffCycleData({ empId: "", basic: 0, hra: 0, conv: 0, med: 0, inc: 0, oth: 0, lop: 0, adv: 0, pt: 0, tds: 0, othD: 0, note: "" });
    alert(`Off-Cycle Payroll saved for ${emps.find((e) => e.id === offCycleData.empId)?.name}`);
  };

  const updAtt = (eid, m, field, val) => {
    setAtt((p) => {
      const empData = p[eid] || {}; const fyData = empData[fy] || getEmptyAtt(); const moData = fyData[m] || {};
      if (field === "CLEAR_ALL") return { ...p, [eid]: { ...empData, [fy]: { ...fyData, [m]: { present: null, leave: null, bal: null, lop: null, holiday: null, comments: "" } } } };
      let newVal = val; if (field !== "comments") newVal = val === "" ? null : Number(val);
      return { ...p, [eid]: { ...empData, [fy]: { ...fyData, [m]: { ...moData, [field]: newVal } } } };
    });
  };

  const getCarryOver = (eid) => {
    if (!["Jan", "Feb", "Mar"].includes(mo)) return 0;
    const decBal = Number(att[eid]?.[fy]?.["Dec"]?.bal || 0);
    if (decBal <= 0) return 0;
    let encashedSoFar = 0;
    if (mo === "Feb" || mo === "Mar") { const jC = att[eid]?.[fy]?.["Jan"]?.comments || ""; const match = jC.match(/(\d+)\s*leave.*encash/i); if (match) encashedSoFar += Number(match[1]); }
    if (mo === "Mar") { const fC = att[eid]?.[fy]?.["Feb"]?.comments || ""; const match = fC.match(/(\d+)\s*leave.*encash/i); if (match) encashedSoFar += Number(match[1]); }
    const remaining = decBal - encashedSoFar;
    return remaining > 0 ? remaining : 0;
  };

  const saveDailyAttendance = async () => {
    const inserts = [];
    Object.keys(dailyForm).forEach(eid => { inserts.push({ emp_id: eid, date: markDate, status: dailyForm[eid].status, reason: dailyForm[eid].reason }); });
    try {
      await supabase.from("gits_daily_attendance").delete().eq("date", markDate);
      if (inserts.length > 0) { const { error } = await supabase.from("gits_daily_attendance").insert(inserts); if (error) throw error; }
      setDailyAtt(prev => [...prev.filter(r => r.date !== markDate), ...inserts]);
      logAction("Attendance", `Saved Daily Calendar records for date: ${markDate}`);
      alert(`Daily Attendance saved for ${markDate}`);
    } catch (err) { alert("Error saving daily attendance: " + err.message); }
  };

  const runLeaveAccrual = () => {
    if (!confirm(`Auto-calculate from Daily Records for ${mo}? This will overwrite manual Leave/Holiday counts.`)) return;
    const yNum = ["Jan", "Feb", "Mar"].includes(mo) ? parseInt(fy) + 1 : parseInt(fy);
    const mIdxStr = String(CAL_MS.indexOf(mo) + 1).padStart(2, '0'); 
    const targetPrefix = `${yNum}-${mIdxStr}`; 
    const newAtt = { ...att };
    emps.forEach(emp => {
        if (!isActiveInMonth(emp, mo, fy)) return;
        if(!newAtt[emp.id]) newAtt[emp.id] = {};
        if(!newAtt[emp.id][fy]) newAtt[emp.id][fy] = getEmptyAtt();
        const empDaily = dailyAtt.filter(r => r.emp_id === emp.id && r.date.startsWith(targetPrefix));
        let aggLeaves = 0; let aggHolidays = 0; let dailyReasons = [];
        empDaily.forEach(d => {
            if (d.status === "Leave" || d.status === "Absent") { aggLeaves++; if (d.reason && d.reason.trim() !== "") dailyReasons.push(d.reason.trim()); }
            if (d.status === "Holiday") { aggHolidays++; }
        });
        let currentAtt = newAtt[emp.id][fy][mo] || {};
        let currentLeaves = aggLeaves; let currentHolidays = aggHolidays; const wDays = getWD(mo, fy);
        let baseComments = currentAtt.comments || ""; let encashStr = "";
        if (baseComments.toLowerCase().includes("encash")) { const match = baseComments.match(/(\d+)\s*leaves?\s*encashed/i); if (match) encashStr = match[0]; }
        if (emp.leavePolicy === "No") {
            let newPresent = wDays - currentHolidays - currentLeaves; if (newPresent < 0) newPresent = 0;
            let finalComment = encashStr; if (dailyReasons.length > 0) { const rStr = dailyReasons.join(", "); finalComment = finalComment ? `${finalComment} | ${rStr}` : rStr; }
            newAtt[emp.id][fy][mo] = { ...currentAtt, present: newPresent, leave: currentLeaves, holiday: currentHolidays, bal: null, lop: null, comments: finalComment };
            return;
        }
        const mIdx = MS.indexOf(mo); let prevMo = mIdx > 0 ? MS[mIdx - 1] : "Mar"; let prevFy = mIdx > 0 ? fy : String(parseInt(fy) - 1);
        let prevBal = 0;
        if (newAtt[emp.id]?.[prevFy]?.[prevMo]?.bal !== undefined && newAtt[emp.id]?.[prevFy]?.[prevMo]?.bal !== null) { prevBal = Number(newAtt[emp.id][prevFy][prevMo].bal); }
        let accrual = 0; if (["Jan", "Apr", "Jul"].includes(mo)) accrual = 1; else if (mo === "Oct") accrual = 2;
        let available = prevBal + accrual; let newBal = 0; let newLop = 0;
        if (currentLeaves > available) { newLop = currentLeaves - available; available = 0; } else { newLop = 0; available -= currentLeaves; }
        if (encashStr) { const match = encashStr.match(/(\d+)\s*leave/i); if (match) { const requestedEncash = Number(match[1]); if (requestedEncash > 0) { const actualEncash = Math.min(requestedEncash, available); available -= actualEncash; if (actualEncash > 0) { encashStr = `${actualEncash} leave${actualEncash > 1 ? 's' : ''} encashed`; } else { encashStr = ""; } } } }
        let newPresent = wDays - currentHolidays - currentLeaves; if (newPresent < 0) newPresent = 0;
        let finalComment = encashStr; if (dailyReasons.length > 0) { const rStr = dailyReasons.join(", "); finalComment = finalComment ? `${finalComment} | ${rStr}` : rStr; }
        newAtt[emp.id][fy][mo] = { ...currentAtt, present: newPresent, leave: currentLeaves, holiday: currentHolidays, bal: available, lop: newLop, comments: finalComment };
    });
    setAtt(newAtt); alert(`Attendance auto-calculated based on Daily Records for ${mo}!`);
  };

  const saveAttendance = async () => {
    try {
      const { error: delErr } = await supabase.from("gits_attendance").delete().eq("fy", fy).eq("mo", mo);
      if (delErr) throw delErr;
      const inserts = [];
      emps.filter((e) => e.id !== "admin").forEach((e) => {
        const a = att[e.id]?.[fy]?.[mo];
        if (a && (a.present !== null || a.leave !== null || a.holiday !== null || a.lop !== null || (a.comments && a.comments.trim() !== ""))) {
          inserts.push({ emp_id: e.id, fy: fy, mo: mo, present: a.present === "" ? null : a.present, leave: a.leave === "" ? null : a.leave, bal: a.bal === "" ? null : a.bal, lop: a.lop === "" ? null : a.lop, holiday: a.holiday === "" ? null : a.holiday, comments: a.comments });
        }
      });
      if (inserts.length > 0) { const { error: insErr } = await supabase.from("gits_attendance").insert(inserts); if (insErr) throw insErr; }
      logAction("Attendance", `Saved bulk Monthly Report grid for ${mo} FY ${fyL(fy)}`);
      alert(`Attendance for ${mo} ${fyL(fy)} saved successfully!`);
    } catch (err) { alert("Error saving attendance: " + err.message); }
  };

  const postPayrollToCashbook = async (targetMo) => {
    if (ses?.role !== "a") return;
    const currentFyLabel = `FY ${fy}-${String(parseInt(fy) + 1).slice(-2)}`;
    const targetCbFy = cbFYs.find(f => f.label === currentFyLabel);
    if (!targetCbFy) return alert(`Please create Cashbook FY '${currentFyLabel}' first!`);

    let totalNet = 0, totalPT = 0, totalTDS = 0;
    emps.forEach(e => { (pay[e.id]?.[fy] || []).filter(r => r.m === targetMo).forEach(r => { totalNet += np(r); totalPT += (r.pt || 0); totalTDS += (r.tds || 0); }); });

    if (totalNet === 0 && totalPT === 0 && totalTDS === 0) return alert(`No payroll data found for ${targetMo}.`);

    let eCount = cbTxns.filter(x => x.financial_year_id === targetCbFy.id && x.type === "expense").length;
    const yr = currentFyLabel.replace("FY 20", "").replace("FY ", ""); 
    const getVNo = () => `E${String(++eCount).padStart(4, '0')}/${yr}`;
    
    const txnYear = ["Jan", "Feb", "Mar"].includes(targetMo) ? parseInt(fy) + 1 : parseInt(fy);
    const txnMonth = String(CAL_MS.indexOf(targetMo) + 1).padStart(2, '0');
    const txnDate = `${txnYear}-${txnMonth}-05`;

    const inserts = [];
    if (totalNet > 0) inserts.push({ financial_year_id: targetCbFy.id, type: "expense", party_name: "Staff Payroll", txn_date: txnDate, amount: totalNet, description: `Net Salary for ${targetMo} ${fy}`, voucher_no: getVNo() });
    if (totalPT > 0) inserts.push({ financial_year_id: targetCbFy.id, type: "expense", party_name: "Professional Tax (PT)", txn_date: txnDate, amount: totalPT, description: `PT Deductions for ${targetMo} ${fy}`, voucher_no: getVNo() });
    if (totalTDS > 0) inserts.push({ financial_year_id: targetCbFy.id, type: "expense", party_name: "TDS Payable", txn_date: txnDate, amount: totalTDS, description: `TDS Deductions for ${targetMo} ${fy}`, voucher_no: getVNo() });

    try {
      const { data, error } = await supabase.from("transactions").insert(inserts).select();
      if (error) throw error;
      setCbTxns(prev => [...data, ...prev]);
      logAction("Cashbook", `Posted Payroll to Cashbook for ${targetMo} ${fy}`);
      alert(`Successfully posted Payroll for ${targetMo} to Cashbook!`);
    } catch (err) { alert("Failed to post: " + err.message); }
  };

  const saveCbFy = async () => {
    if (!newCbFyStart || !newCbFyEnd) return alert("Select both dates");
    const label = deriveFYLabel(newCbFyStart);
    if (cbFYs.some(f => f.label === label)) return alert("FY already exists");
    try {
      const { data, error } = await supabase.from("financial_years").insert({ label, start_date: newCbFyStart, end_date: newCbFyEnd }).select().single();
      if (error) throw error;
      setCbFYs([data, ...cbFYs]); setCbSelFY(data.id); setShowAddCbFy(false);
      logAction("Cashbook", `Created Financial Year: ${label}`);
      alert(`Financial Year ${label} added successfully`);
    } catch (err) { alert("Failed to add FY: " + err.message); }
  };

  const saveCbTxn = async () => {
    if (!cbSelFY || !cbParty || !cbAmount) return alert("Fill required fields");
    if (editCbTxn) {
        const payload = { type: cbType, party_name: cbParty, txn_date: cbDate, amount: Number(cbAmount), description: cbDesc };
        try {
            const { data, error } = await supabase.from("transactions").update(payload).eq("id", editCbTxn.id).select().single();
            if (error) throw error;
            setCbTxns(cbTxns.map(t => t.id === editCbTxn.id ? data : t));
            setCbParty(""); setCbAmount(""); setCbDesc(""); setEditCbTxn(null);
            logAction("Cashbook", `Edited ${cbType} for ${cbParty} (Voucher: ${editCbTxn.voucher_no})`);
            alert("Transaction updated!");
        } catch (err) { alert("Update failed: " + err.message); }
    } else {
        const vNo = (() => {
          const p = cbType === "receipt" ? "R" : cbType === "payment" ? "P" : "E";
          const fyLbl = cbFYs.find(f => f.id === cbSelFY)?.label || "";
          const yr = fyLbl.replace("FY 20", "").replace("FY ", ""); 
          const existing = cbTxns.filter(x => x.financial_year_id === cbSelFY && x.type === cbType);
          return `${p}${String(existing.length + 1).padStart(4, '0')}/${yr}`;
        })();
        const payload = { financial_year_id: cbSelFY, type: cbType, party_name: cbParty, txn_date: cbDate, amount: Number(cbAmount), description: cbDesc, voucher_no: vNo };
        try {
          const { data, error } = await supabase.from("transactions").insert(payload).select().single();
          if (error) throw error;
          setCbTxns([data, ...cbTxns]); setCbParty(""); setCbAmount(""); setCbDesc("");
          logAction("Cashbook", `Added ${cbType} for ${cbParty}: ₹${cbAmount}`);
          alert("Transaction saved! Voucher No: " + vNo);
        } catch (err) { alert("Save failed: " + err.message); }
    }
  };

  const delCbTxn = async (id, vNo) => {
    if (!confirm(`Delete transaction ${vNo || ""} permanently?`)) return;
    try {
      await supabase.from("transactions").delete().eq("id", id);
      setCbTxns(cbTxns.filter(t => t.id !== id));
      logAction("Cashbook", `Deleted transaction ID: ${id} (${vNo || ""})`);
    } catch (err) { alert("Delete failed"); }
  };

  if (!dbLoaded) return <div style={{ padding: 50, textAlign: "center", fontFamily: "sans-serif" }}><h3>Connecting to Live Database...</h3></div>;
  if (slip) return <div style={{ position: "fixed", inset: 0, zIndex: 9999, background: "rgba(0,0,0,0.8)", display: "flex", flexDirection: "column" }}><button style={{ padding: 15, background: "#1a1a2e", color: "#fff", border: "none", cursor: "pointer", fontWeight: "bold", fontSize: 16 }} onClick={() => setSlip(null)}>✕ Close PDF Viewer</button><iframe srcDoc={slip} style={{ flex: 1, border: "none", background: "#fff" }} /></div>;

  if (!ses) return (
    <div style={{ display: "flex", justifyContent: "center", paddingTop: 100, fontFamily: "sans-serif", background: "#f4f7f6", minHeight: "100vh" }}>
      <div style={{ width: 320, padding: 30, background: "#fff", border: "1px solid #ddd", borderRadius: 12, boxShadow: "0 4px 12px rgba(0,0,0,0.1)", height: "fit-content" }}>
        <img src="/logo.png" alt="GITS Logo" style={{ maxHeight: 60, display: "block", margin: "0 auto 20px" }} />
        <h3 style={{ textAlign: "center", marginBottom: 20, color: "#666", marginTop: -10 }}>Portal Login</h3>
        {forgotStep === 0 ? (
          <>
            <input style={sInp} placeholder="Employee ID" onChange={(e) => (idR.current = e.target.value)} />
            <input style={{ ...sInp, marginTop: 15, marginBottom: 5 }} type="password" placeholder="Password" onChange={(e) => (pwR.current = e.target.value)} />
            <div style={{ textAlign: "right", marginBottom: 20 }}><a href="#" style={{ fontSize: 11, color: "#185FA5", textDecoration: "none" }} onClick={(e) => { e.preventDefault(); setForgotStep(1); }}>Forgot Password?</a></div>
            <button style={{ width: "100%", padding: 12, background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer" }} onClick={login}>Login</button>
          </>
        ) : forgotStep === 1 && !forgotUser ? (
          <>
            <h4 style={{ marginTop: 0, marginBottom: 10 }}>Password Recovery</h4>
            <p style={{ fontSize: 13, color: "#666", marginBottom: 15 }}>Enter your Employee ID or Email.</p>
            <input style={{ ...sInp, marginBottom: 20 }} placeholder="Employee ID or Email" value={forgotId} onChange={(e) => setForgotId(e.target.value)} />
            <button style={{ width: "100%", padding: 12, background: "#185FA5", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", marginBottom: 10 }} onClick={handleForgotNext}>Continue</button>
            <button style={{ width: "100%", padding: 12, background: "#fff", color: "#666", border: "1px solid #ddd", borderRadius: 6, cursor: "pointer" }} onClick={() => { setForgotStep(0); setForgotId(""); }}>Cancel</button>
          </>
        ) : forgotStep === 1 && forgotUser ? (
          <>
            <h4 style={{ marginTop: 0, marginBottom: 10 }}>Security Question</h4>
            <p style={{ fontSize: 13, color: "#1a1a2e", fontWeight: "bold", marginBottom: 15 }}>{forgotUser.sec_q}</p>
            <input style={{ ...sInp, marginBottom: 20 }} placeholder="Type your answer" value={secAnsInput} onChange={(e) => setSecAnsInput(e.target.value)} />
            <button style={{ width: "100%", padding: 12, background: "#185FA5", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", marginBottom: 10 }} onClick={handleForgotNext}>Verify Answer</button>
            <button style={{ width: "100%", padding: 12, background: "#fff", color: "#666", border: "1px solid #ddd", borderRadius: 6, cursor: "pointer" }} onClick={() => { setForgotStep(0); setForgotUser(null); setSecAnsInput(""); }}>Cancel</button>
          </>
        ) : (
          <>
            <h4 style={{ marginTop: 0, marginBottom: 10 }}>Reset Password</h4>
            <input style={{ ...sInp, marginBottom: 20 }} type="password" placeholder="Enter New Password" value={newForgotPwd} onChange={(e) => setNewForgotPwd(e.target.value)} />
            <button style={{ width: "100%", padding: 12, background: "#1D9E75", color: "#fff", border: "none", borderRadius: 6, cursor: "pointer", marginBottom: 10 }} onClick={handleForgotNext}>Save & Login</button>
          </>
        )}
      </div>
    </div>
  );

  const TABS = ses?.role === "a" ? ["dashboard", "employees", "attendance", "ledger", "payslips", "cashbook", "audit", "profile"] : ["dashboard", "attendance", "ledger", "payslips", "profile"];
  const ADMIN_DRIVE = "https://drive.google.com/drive/folders/1kq4pVPRpaycQczhgycGz0dGHLUn2LmG6?usp=sharing";
  const DRIVE_LINK = ses.role === "a" ? ADMIN_DRIVE : myE?.driveLink;

  // --- CASHBOOK LEDGER CALCULATIONS ---
  const selectedFyObj = cbFYs.find(f => f.id === cbSelFY);
  let openingBalance = 0; let closingBalance = 0;
  
  if (selectedFyObj) {
      const fyStart = new Date(selectedFyObj.start_date);
      const priorTxns = cbTxns.filter(t => new Date(t.txn_date) < fyStart);
      const priorRec = priorTxns.filter(t => t.type === 'receipt').reduce((s, t) => s + Number(t.amount), 0);
      const priorPay = priorTxns.filter(t => t.type !== 'receipt').reduce((s, t) => s + Number(t.amount), 0);
      openingBalance = 35256.42 + priorRec - priorPay;
  }

  const cbActiveTxns = cbTxns.filter(t => t.financial_year_id === cbSelFY);
  let cbFilteredTxns = [...cbActiveTxns]; 
  if (cbTabState !== "all") cbFilteredTxns = cbFilteredTxns.filter(t => t.type === cbTabState);
  if (cbSearch.trim()) {
    const q = cbSearch.toLowerCase();
    cbFilteredTxns = cbFilteredTxns.filter(t => t.party_name.toLowerCase().includes(q) || (t.description||"").toLowerCase().includes(q) || (t.voucher_no||"").toLowerCase().includes(q));
  }
  
  cbFilteredTxns.sort((a, b) => new Date(a.txn_date).getTime() - new Date(b.txn_date).getTime());
  
  const cbSum = (type) => cbActiveTxns.filter((t) => t.type === type).reduce((s, t) => s + Number(t.amount || 0), 0);
  const cbReceipts = cbSum("receipt"); const cbPayments = cbSum("payment"); const cbExpenses = cbSum("expense");
  const cbNet = cbReceipts - cbPayments - cbExpenses;
  closingBalance = openingBalance + cbNet;

  const isFullLedger = cbTabState === "all" && cbSearch.trim() === "";

  // --- DASHBOARD CASHBOOK CALCULATIONS ---
  const dashCbFy = cbFYs.find(f => f.label === `FY ${fy}-${String(parseInt(fy) + 1).slice(-2)}`);
  const dashCbTxns = dashCbFy ? cbTxns.filter(t => t.financial_year_id === dashCbFy.id) : [];
  
  let dashOpeningBalance = 0;
  if (dashCbFy) {
      const fyStart = new Date(dashCbFy.start_date);
      const priorTxns = cbTxns.filter(t => new Date(t.txn_date) < fyStart);
      const priorRec = priorTxns.filter(t => t.type === 'receipt').reduce((s, t) => s + Number(t.amount), 0);
      const priorPay = priorTxns.filter(t => t.type !== 'receipt').reduce((s, t) => s + Number(t.amount), 0);
      dashOpeningBalance = 35256.42 + priorRec - priorPay;
  }
  
  const dCbRec = dashCbTxns.filter(t => t.type === 'receipt').reduce((s, t) => s + Number(t.amount), 0);
  const dCbPay = dashCbTxns.filter(t => t.type === 'payment').reduce((s, t) => s + Number(t.amount), 0);
  const dCbExp = dashCbTxns.filter(t => t.type === 'expense').reduce((s, t) => s + Number(t.amount), 0);
  const dCbNet = dCbRec - dCbPay - dCbExp;
  const dashClosingBalance = dashOpeningBalance + dCbNet;

  const currentFyLabelCheck = `FY ${fy}-${String(parseInt(fy) + 1).slice(-2)}`;
  const targetCbFyCheck = cbFYs.find(f => f.label === currentFyLabelCheck);
  const isPayrollSynced = targetCbFyCheck ? cbTxns.some(t => t.financial_year_id === targetCbFyCheck.id && t.description && t.description.includes(`for ${syncMo} ${fy}`) && t.party_name === "Staff Payroll") : false;

  // ==========================================
  // SINGLE COMPONENT RENDER BLOCK 
  // (Ensuring no tabs are missed or duplicated)
  // ==========================================
  return (
    <div style={{ padding: 20, fontFamily: "sans-serif", maxWidth: 1200, margin: "auto" }}>
      {sysWarn && (
        <div style={{ background: "#fff3cd", color: "#856404", padding: "10px 15px", borderRadius: 6, marginBottom: 20, border: "1px solid #ffeeba", fontSize: 13 }}>
          <b>Warning: Database issue detected.</b> Some data may not load. Detail: {sysWarn}.
        </div>
      )}
      
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "2px solid #1a1a2e", paddingBottom: 10, marginBottom: 20 }}>
        <div style={{ display: "flex", alignItems: "center", gap: "15px" }}>
          <img src="/logo.png" alt="GITS Logo" style={{ maxHeight: 45 }} />
          <div><h2 style={{ margin: 0, color: "#444" }}>{ses.role === "e" ? "Employee Portal" : "Admin Portal"}</h2><small>{AD}</small></div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          {DRIVE_LINK && <a href={DRIVE_LINK} target="_blank" rel="noreferrer" style={{ ...btn, background: "#4285F4", color: "#fff", textDecoration: "none", fontWeight: "bold", display: "flex", alignItems: "center" }}>📂 {ses.role === "a" ? "Admin Drive" : "My Folder"}</a>}
          {tab !== "cashbook" && <div style={{ display: "flex", alignItems: "center", gap: 5 }}><span style={{ fontSize: 13, fontWeight: "bold" }}>FY</span><input type="number" value={fy} onChange={(e) => setFy(e.target.value)} style={{ ...sInp, width: 80, padding: "6px" }} /><span style={{ fontSize: 13, fontWeight: "bold" }}>{fyL(fy)}</span></div>}
          <button style={btn} onClick={() => { setSes(null); setTab("dashboard"); }}>Sign Out</button>
        </div>
      </div>
      
      <div style={{ display: "flex", gap: 10, marginBottom: 25, borderBottom: "1px solid #eee", flexWrap: "wrap" }}>
        {TABS.map((t) => (
          <button key={t} style={{ padding: "10px 15px", background: "none", border: "none", borderBottom: tab === t ? "3px solid #1a1a2e" : "none", cursor: "pointer", fontWeight: tab === t ? "bold" : "normal" }} onClick={() => { setTab(t); if (t === "profile" && myE) { setSecQ(myE.sec_q || ""); setSecA(myE.sec_a || ""); } }}>{t.toUpperCase()}</button>
        ))}
      </div>

      {tab === "dashboard" && (
        <div>
          <div style={{ display: "flex", gap: 10, marginBottom: 20, alignItems: "center" }}>
            <h3 style={{ margin: 0, color: "#444" }}>Payroll Overview ({mo} {fy})</h3>
            <select value={mo} onChange={(e) => setMo(e.target.value)} style={{ ...sInp, width: 150, marginLeft: 15 }}>{MS.map((m) => (<option key={m} value={m}>{mL(m, fy)}</option>))}</select>
            {ses.role === "a" && (
              <div style={{ marginLeft: "auto", display: "flex", gap: 10, flexWrap: "wrap" }}>
                <button style={{ ...btn, background: "#1a1a2e", color: "#fff" }} onClick={openBulkPayroll}>+ Run Bulk Payroll</button>
                <button style={{ ...btn, background: "#185FA5", color: "#fff", marginRight: 15 }} onClick={() => { setShowOffCycle(true); setShowBulk(false); }}>+ Off-Cycle Payroll</button>
                <button style={{ ...btn, background: "#1D9E75", color: "#fff" }} onClick={() => {
                  const head = ["S.No", "Emp ID", "Employee", "Role", "Basic", "HRA", "Conv", "Med", "Inc", "Oth Earn", "Gross", "LOP", "Advance", "PT", "TDS", "Oth Ded", "Total Deductions", "Taxable Income", "Net", "Note"];
                  const rows = []; let sno = 1;
                  emps.filter((e) => isActiveInMonth(e, mo, fy)).forEach((e) => (pay[e.id]?.[fy] || []).filter((r) => r.m === mo).forEach((r) => rows.push([sno++, e.id, e.name, e.desig, r.basic, r.hra, r.conv, r.med, r.inc, r.oth, gr(r), r.lop, r.adv, r.pt, r.tds, r.othD || 0, dd(r), txInc(r), np(r), r.note || ""])));
                  exportExcel([head, ...rows], `Payroll_${mL(mo, fy)}`);
                }}>📊 Excel</button>
                <button style={{ ...btn, background: "#d32f2f", color: "#fff" }} onClick={() => {
                  const head = ["S.No", "Emp ID", "Employee", "Gross", "LOP", "Adv", "PT", "TDS", "Oth Ded", "Tot Ded", "Net", "Note"];
                  const rows = []; let sno = 1;
                  emps.filter((e) => isActiveInMonth(e, mo, fy)).forEach((e) => (pay[e.id]?.[fy] || []).filter((r) => r.m === mo).forEach((r) => rows.push([sno++, e.id, e.name, f$(gr(r)), f$(r.lop), f$(r.adv), f$(r.pt), f$(r.tds), f$(r.othD || 0), f$(dd(r)), f$(np(r)), r.note || "-"])));
                  setSlip(buildReportPdf("Monthly Payroll Report", `Payroll generated for ${mo} ${fyL(fy)}`, head, rows));
                }}>📄 PDF</button>
              </div>
            )}
          </div>
          
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 20, textAlign: "center", marginBottom: 30 }}>
            <div style={card}><div style={lbl}>{ses.role === "a" ? "Total Staff" : "My Salary"}</div><div style={{ fontSize: 24, fontWeight: "bold" }}>{ses.role === "a" ? emps.filter((e) => isActiveInMonth(e, mo, fy)).length : f$(myE?.basic)}</div></div>
            <div style={card}><div style={lbl}>{ses.role === "a" ? "Gross Month" : "My YTD Gross"}</div><div style={{ fontSize: 24, fontWeight: "bold" }}>{ses.role === "a" ? f$(dTot.g) : f$((pay[myE?.id]?.[fy] || []).reduce((s, r) => s + gr(r), 0))}</div></div>
            <div style={card}><div style={lbl}>Deductions</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#D85A30" }}>{f$(dTot.d)}</div></div>
            <div style={card}><div style={lbl}>Net Amount</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#1D9E75" }}>{f$(dTot.n)}</div></div>
          </div>

          {ses.role === "a" && (
            <>
              <h3 style={{ margin: "0 0 15px 0", color: "#444", borderTop: "2px solid #eee", paddingTop: 20 }}>Cashbook Overview (FY {fyL(fy)})</h3>
              {!dashCbFy ? (
                  <div style={{ padding: 20, background: "#fff8e1", color: "#f57c00", borderRadius: 6, border: "1px solid #ffc107", fontSize: 13 }}>
                      <b>Note:</b> Please create Financial Year <b>FY {fyL(fy)}</b> in the Cashbook tab to view overarching financial statistics here.
                  </div>
              ) : (
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 20, textAlign: "center", marginBottom: 20 }}>
                      <div style={{ ...card, borderTop: "4px solid #185FA5" }}><div style={lbl}>Opening Balance</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#185FA5" }}>{f$(dashOpeningBalance)}</div></div>
                      <div style={{ ...card, borderTop: "4px solid #0F9D6D" }}><div style={lbl}>Total Receipts</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#0F9D6D" }}>{f$(dCbRec)}</div></div>
                      <div style={{ ...card, borderTop: "4px solid #C23A55" }}><div style={lbl}>Payments & Expenses</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#C23A55" }}>{f$(dCbPay + dCbExp)}</div></div>
                      <div style={{ ...card, borderTop: dashClosingBalance >= 0 ? "4px solid #10243E" : "4px solid #C23A55" }}><div style={lbl}>Closing Balance</div><div style={{ fontSize: 24, fontWeight: "bold", color: dashClosingBalance >= 0 ? "#10243E" : "#C23A55" }}>{f$(dashClosingBalance)}</div></div>
                  </div>
              )}
            </>
          )}

          {showBulk && ses.role === "a" && (
            <div style={{ ...card, border: "1px solid #1a1a2e", marginBottom: 20 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <h3 style={{ margin: 0 }}>Bulk Payroll for {mL(mo, fy)}</h3>
                <div>
                  <button style={{ ...btn, background: "#1D9E75", color: "#fff" }} onClick={saveBulkPayroll}>Save All</button>
                  <button style={{ ...btn, marginLeft: 10 }} onClick={() => setShowBulk(false)}>Cancel</button>
                </div>
              </div>
              <div style={{background: "#e8f5e9", color: "#1D9E75", padding: "10px", borderRadius: 4, marginBottom: 15, fontSize: 12, fontWeight: "bold"}}>
                  ✓ Dynamic LOP: LOP auto-calculates purely on (Basic+HRA+Conv+Med). Incentives will not affect LOP. Encashments use flat 30 days.
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", fontSize: 12, borderCollapse: "collapse" }}>
                  <thead><tr style={{ textAlign: "left", whiteSpace: "nowrap" }}>
                    <th style={{ padding: 8 }}>S.No</th><th style={{ padding: 8 }}>Emp ID</th><th style={{ padding: 8 }}>Employee</th><th style={{ padding: 8 }}>Basic</th><th style={{ padding: 8 }}>HRA</th><th style={{ padding: 8 }}>Conv.</th><th style={{ padding: 8 }}>Med.</th><th style={{ padding: 8 }}>Inc.</th><th style={{ padding: 8 }}>Oth Earn</th><th style={{ padding: 8 }}>LOP (₹)</th><th style={{ padding: 8 }}>Advance</th><th style={{ padding: 8 }}>PT</th><th style={{ padding: 8 }}>TDS</th><th style={{ padding: 8 }}>Oth Ded.</th><th style={{ padding: 8 }}>Note</th><th style={{ padding: 8 }}>Taxable</th><th style={{ padding: 8 }}>Net</th>
                  </tr></thead>
                  <tbody>{Object.keys(bulkData).map((eid, idx) => {
                    const d = bulkData[eid]; 
                    const upd = (k, v) => {
                        const nextD = { ...d, [k]: v };
                        if (["basic", "hra", "conv", "med"].includes(k)) {
                            const calDays = getCalendarDays(mo, fy);
                            const corePay = (+nextD.basic || 0) + (+nextD.hra || 0) + (+nextD.conv || 0) + (+nextD.med || 0);
                            const a = att[eid]?.[fy]?.[mo];
                            const lopDays = Number(a?.lop || 0);
                            if (lopDays > 0) nextD.lop = Math.round((corePay / calDays) * lopDays);
                            if (a?.comments && a.comments.toLowerCase().includes("encash")) {
                                const match = a.comments.match(/(\d+)\s*leave/i);
                                if (match) { const eDays = Number(match[1]); nextD.oth = Math.round((corePay / 30) * eDays); }
                            }
                        }
                        setBulkData({ ...bulkData, [eid]: nextD });
                    };
                    const bInp = { padding: "6px", border: "1px solid #ddd", borderRadius: 4, minWidth: 80, fontSize: 12 };
                    const g = (+d.basic || 0) + (+d.hra || 0) + (+d.conv || 0) + (+d.med || 0) + (+d.inc || 0) + (+d.oth || 0);
                    const ded = (+d.lop || 0) + (+d.adv || 0) + (+d.pt || 0) + (+d.tds || 0) + (+d.othD || 0);
                    const n = g - ded; const taxIncVal = n + (+d.pt || 0) + (+d.tds || 0);
                    return (
                      <tr key={eid} style={{ borderBottom: "1px solid #eee" }}>
                        <td style={{ padding: 8, color: "#666" }}>{idx + 1}</td><td style={{ padding: 8, color: "#666" }}>{eid}</td>
                        <td style={{ whiteSpace: "nowrap", paddingRight: 10 }}>{emps.find((e) => e.id === eid)?.name}</td>
                        <td><input style={bInp} type="number" value={d.basic} onChange={(e) => upd("basic", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.hra} onChange={(e) => upd("hra", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.conv} onChange={(e) => upd("conv", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.med} onChange={(e) => upd("med", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.inc} onChange={(e) => upd("inc", e.target.value)} /></td>
                        <td><input style={{...bInp, background: d.oth > 0 ? "#fff8e1" : "#fff"}} type="number" value={d.oth} onChange={(e) => upd("oth", e.target.value)} /></td>
                        <td><input style={{...bInp, background: d.lop > 0 ? "#ffebee" : "#fff"}} type="number" value={d.lop} onChange={(e) => upd("lop", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.adv} onChange={(e) => upd("adv", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.pt} onChange={(e) => upd("pt", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.tds} onChange={(e) => upd("tds", e.target.value)} /></td>
                        <td><input style={bInp} type="number" value={d.othD} onChange={(e) => upd("othD", e.target.value)} /></td>
                        <td><input style={{ ...bInp, minWidth: 120 }} type="text" placeholder="Optional" value={d.note || ""} onChange={(e) => upd("note", e.target.value)} /></td>
                        <td style={{ fontWeight: "bold", paddingLeft: 10 }}>{f$(taxIncVal)}</td>
                        <td style={{ fontWeight: "bold", color: "#1D9E75", paddingLeft: 10 }}>{f$(n)}</td>
                      </tr>
                    );
                  })}</tbody>
                </table>
              </div>
            </div>
          )}
          {showOffCycle && ses.role === "a" && (
            <div style={{ ...card, border: "1px solid #185FA5", marginBottom: 20 }}>
              <h3>Run Off-Cycle Payroll for {mL(mo, fy)}</h3>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 10, marginBottom: 15 }}>
                <div><label style={lbl}>Employee</label><select style={sInp} value={offCycleData.empId} onChange={(e) => {
                  const id = e.target.value;
                  if (id) setOffCycleData({ empId: id, basic: 0, hra: 0, conv: 0, med: 0, inc: 0, oth: 0, lop: 0, adv: 0, pt: 0, tds: 0, othD: 0, note: "Off-Cycle Arrears/Adj" });
                  else setOffCycleData({ empId: "", basic: 0, hra: 0, conv: 0, med: 0, inc: 0, oth: 0, lop: 0, adv: 0, pt: 0, tds: 0, othD: 0, note: "" });
                }}><option value="">Select...</option>{emps.filter((e) => isActiveInMonth(e, mo, fy)).map((e) => (<option key={e.id} value={e.id}>{e.name}</option>))}</select></div>
                {[["Basic", "basic"], ["HRA", "hra"], ["Conv", "conv"], ["Med", "med"], ["Incentive", "inc"], ["Oth Earn", "oth"], ["LOP", "lop"], ["Advance", "adv"], ["PT", "pt"], ["TDS", "tds"], ["Oth Ded", "othD"]].map(([l, k]) => (<div key={k}><label style={lbl}>{l}</label><input style={sInp} type="number" value={offCycleData[k]} onChange={(e) => setOffCycleData({ ...offCycleData, [k]: e.target.value })} /></div>))}
                <div style={{ gridColumn: "1/-1" }}><label style={lbl}>Note / Reason</label><input style={sInp} value={offCycleData.note} placeholder="e.g. Arrears" onChange={(e) => setOffCycleData({ ...offCycleData, note: e.target.value })} /></div>
              </div>
              <button style={{ ...btn, background: "#185FA5", color: "#fff", marginRight: 10 }} onClick={saveOffCycle}>Save Off-Cycle Entry</button>
              <button style={btn} onClick={() => setShowOffCycle(false)}>Cancel</button>
            </div>
          )}
        </div>
      )}

      {tab === "employees" && ses.role === "a" && (
        <div>
          <div style={{ marginBottom: 15, display: "flex", justifyContent: "flex-end", gap: 10 }}>
            <button style={{ ...btn, background: "#1D9E75", color: "#fff" }} onClick={() => {
              const head = ["S.No", "Emp ID", "Name", "Designation", "PAN", "Email", "Phone", "Start Date", "End Date", "Status", "Category", "Basic Salary", "Bank Details", "Comments", "Drive Link", "Leave Policy"];
              const rows = []; let sno = 1;
              emps.filter((e) => e.id !== "admin").forEach((e) => rows.push([sno++, e.id, e.name, e.desig, e.pan || "-", e.email, e.phone, e.start || "", e.end || "", e.status, e.cat, e.basic, e.bank, e.comments || "", e.driveLink || "", e.leavePolicy || "Yes"]));
              exportExcel([head, ...rows], `Employees_${fyL(fy)}`);
            }}>📊 Excel</button>
            <button style={{ ...btn, background: "#d32f2f", color: "#fff", marginRight: 15 }} onClick={() => {
              const head = ["S.No", "Emp ID", "Name", "Designation", "PAN", "Phone", "Start Date", "Status", "Bank Details"];
              const rows = []; let sno = 1;
              emps.filter((e) => e.id !== "admin").forEach((e) => rows.push([sno++, e.id, e.name, e.desig, e.pan || "-", e.phone || "-", e.start || "-", e.status, e.bank || "-"]));
              setSlip(buildReportPdf("Employee Roster", `Generated on ${new Date().toLocaleDateString()}`, head, rows));
            }}>📄 PDF</button>

            <button style={{ ...btn, background: "#1a1a2e", color: "#fff" }} onClick={() => setShowAddEmp(!showAddEmp)}>+ Add Employee</button>
          </div>
          {showAddEmp && (
            <div style={{ ...card, border: "1px solid #1a1a2e", marginBottom: 20 }}>
              <h3 style={{ marginTop: 0 }}>Add New Employee</h3>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(180px,1fr))", gap: 15 }}>
                <div><label style={lbl}>Emp ID</label><input style={sInp} value={nE.id} onChange={(e) => setNE({ ...nE, id: e.target.value })} /></div>
                <div><label style={lbl}>Name</label><input style={sInp} value={nE.name} onChange={(e) => setNE({ ...nE, name: e.target.value })} /></div>
                <div><label style={lbl}>Designation</label><input style={sInp} value={nE.desig} onChange={(e) => setNE({ ...nE, desig: e.target.value })} /></div>
                <div><label style={lbl}>PAN</label><input style={sInp} value={nE.pan} onChange={(e) => setNE({ ...nE, pan: e.target.value })} /></div>
                <div><label style={lbl}>Gross Basic Salary</label><input style={sInp} type="number" value={nE.basic} onChange={(e) => setNE({ ...nE, basic: e.target.value })} /></div>
                <div><label style={lbl}>Leave Policy Apply</label><select style={sInp} value={nE.leavePolicy} onChange={(e) => setNE({ ...nE, leavePolicy: e.target.value })}><option>Yes</option><option>No</option></select></div>
                <div><label style={lbl}>Phone</label><input style={sInp} value={nE.phone} onChange={(e) => setNE({ ...nE, phone: e.target.value })} /></div>
                <div><label style={lbl}>Email</label><input style={sInp} value={nE.email} onChange={(e) => setNE({ ...nE, email: e.target.value })} /></div>
                <div><label style={lbl}>Start Date</label><input style={sInp} type="date" value={nE.start} onChange={(e) => setNE({ ...nE, start: e.target.value })} /></div>
                <div><label style={lbl}>Bank A/C</label><input style={sInp} value={nE.bank} onChange={(e) => setNE({ ...nE, bank: e.target.value })} /></div>
                <div><label style={lbl}>Drive Link</label><input style={sInp} value={nE.driveLink} placeholder="Optional URL" onChange={(e) => setNE({ ...nE, driveLink: e.target.value })} /></div>
              </div>
              <div style={{ marginTop: 15 }}><button style={{ ...btn, background: "green", color: "#fff", marginRight: 10 }} onClick={addEmployee}>Save Employee</button><button style={btn} onClick={() => setShowAddEmp(false)}>Cancel</button></div>
            </div>
          )}
          {editEmp && (
            <div style={{ background: "#f0f7ff", padding: 20, borderRadius: 8, marginBottom: 20, border: "1px solid #007bff" }}>
              <h4 style={{ marginTop: 0 }}>Editing: {emps.find((e) => e.id === editEmp)?.name}</h4>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 15 }}>
                <div><label style={lbl}>Phone</label><input style={sInp} value={editData.phone || ""} onChange={(e) => setEditData({ ...editData, phone: e.target.value })} /></div>
                <div><label style={lbl}>Email</label><input style={sInp} value={editData.email || ""} onChange={(e) => setEditData({ ...editData, email: e.target.value })} /></div>
                <div><label style={lbl}>PAN</label><input style={sInp} value={editData.pan || ""} onChange={(e) => setEditData({ ...editData, pan: e.target.value })} /></div>
                <div><label style={lbl}>Bank Details</label><input style={sInp} value={editData.bank || ""} onChange={(e) => setEditData({ ...editData, bank: e.target.value })} /></div>
                <div><label style={lbl}>Drive Link</label><input style={sInp} value={editData.driveLink || ""} onChange={(e) => setEditData({ ...editData, driveLink: e.target.value })} /></div>
                <div><label style={lbl}>Start Date</label><input style={sInp} type="date" value={editData.start || ""} onChange={(e) => setEditData({ ...editData, start: e.target.value })} /></div>
                <div><label style={lbl}>End Date</label><input style={sInp} type="date" value={editData.end || ""} onChange={(e) => setEditData({ ...editData, end: e.target.value })} /></div>
                <div><label style={lbl}>Leave Policy Apply</label><select style={sInp} value={editData.leavePolicy || "Yes"} onChange={(e) => setEditData({ ...editData, leavePolicy: e.target.value })}><option>Yes</option><option>No</option></select></div>
                <div><label style={lbl}>Status</label><select style={sInp} value={editData.status || "Active"} onChange={(e) => setEditData({ ...editData, status: e.target.value })}><option>Active</option><option>Resigned</option><option>Terminated</option></select></div>
                {editData.end && <div><label style={lbl}>Reason for leaving</label><input style={sInp} value={editData.reason || ""} onChange={(e) => setEditData({ ...editData, reason: e.target.value })} /></div>}
                <div style={{ borderLeft: "3px solid #D85A30", paddingLeft: 10 }}>
                  <label style={{ ...lbl, color: "#D85A30", fontWeight: "bold" }}>Force Reset Pwd</label>
                  <input style={sInp} value={editData.adminForcePwd || ""} onChange={(e) => setEditData({ ...editData, adminForcePwd: e.target.value })} placeholder="Type new password" />
                </div>
                <div style={{ gridColumn: "1/-1" }}><label style={lbl}>Comments</label><input style={sInp} value={editData.comments || ""} onChange={(e) => setEditData({ ...editData, comments: e.target.value })} /></div>
              </div>
              <div style={{ marginTop: 15, display: "flex", gap: 10 }}>
                <button style={{ ...btn, background: "#1D9E75", color: "#fff", border: "none" }} onClick={saveEditEmployee}>Save Update</button>
                <button style={btn} onClick={() => setEditEmp(null)}>Cancel</button>
              </div>
            </div>
          )}
          <div style={{ ...card, padding: 0, overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead>
                <tr style={{ background: "#f4f4f4", textAlign: "left", whiteSpace: "nowrap" }}>
                  <th style={thS}>Sl.No</th><th style={thS}>Employee ID</th><th style={thS}>Employee Name</th><th style={thS}>PAN</th><th style={thS}>Leave Policy</th><th style={thS}>Email ID</th><th style={thS}>Phone Number</th><th style={thS}>Start Date</th><th style={thS}>End Date</th><th style={thS}>Status</th><th style={thS}>Bank Account</th><th style={thS}>Action</th>
                </tr>
              </thead>
              <tbody>
                {emps.filter((e) => e.id !== "admin").map((e, idx) => (
                  <tr key={e.id} style={{ borderBottom: "1px solid #eee" }}>
                    <td style={{ ...tdS, color: "#666" }}>{idx + 1}</td>
                    <td style={tdS}>{e.id}</td>
                    <td style={tdS}><b>{e.name}</b><br /><small style={{ color: "#888" }}>{e.desig}</small></td>
                    <td style={tdS}>{e.pan || "-"}</td>
                    <td style={tdS}><span style={{ padding: "3px 8px", borderRadius: 12, fontSize: 11, fontWeight: "bold", background: e.leavePolicy === "Yes" ? "#e8f5e9" : "#fff8e1", color: e.leavePolicy === "Yes" ? "#1D9E75" : "#f57c00" }}>{e.leavePolicy || "Yes"}</span></td>
                    <td style={tdS}>{e.email || "-"}</td>
                    <td style={tdS}>{e.phone || "-"}</td>
                    <td style={tdS}>{fmtDate(e.start)}</td>
                    <td style={tdS}>{fmtDate(e.end)}</td>
                    <td style={tdS}><span style={{ padding: "3px 8px", borderRadius: 12, fontSize: 11, fontWeight: "bold", background: e.status === "Active" ? "#e8f5e9" : (e.status === "Resigned" ? "#fff8e1" : "#ffebee"), color: e.status === "Active" ? "#1D9E75" : (e.status === "Resigned" ? "#d32f2f" : "#b71c1c") }}>{e.status || "Unknown"}</span></td>
                    <td style={tdS}>{e.bank || "-"}</td>
                    <td style={tdS}>
                      <div style={{ display: "flex", gap: 5 }}>
                        <button style={{ ...btn, padding: "4px 8px" }} onClick={() => { setEditEmp(e.id); setEditData({ ...e }); }}>Edit</button>
                        <button style={{ ...btn, padding: "4px 8px", color: "red" }} onClick={() => delEmployee(e.id)}>Del</button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === "attendance" && (
        <div>
          <div style={{ display: "flex", gap: 10, marginBottom: 20, alignItems: "center" }}>
             <select value={mo} onChange={(e) => setMo(e.target.value)} style={{ ...sInp, width: 150 }}>{MS.map((m) => (<option key={m} value={m}>{mL(m, fy)}</option>))}</select>
             <div style={{ display: "flex", background: "#fff", borderRadius: 6, border: "1px solid #ddd", overflow: "hidden", marginLeft: 20 }}>
                <button style={{ padding: "8px 16px", border: "none", cursor: "pointer", fontWeight: "bold", background: attSubTab === "daily" ? "#1a1a2e" : "#fff", color: attSubTab === "daily" ? "#fff" : "#666" }} onClick={() => setAttSubTab("daily")}>📅 Daily Marking</button>
                <button style={{ padding: "8px 16px", border: "none", cursor: "pointer", fontWeight: "bold", background: attSubTab === "monthly" ? "#1a1a2e" : "#fff", color: attSubTab === "monthly" ? "#fff" : "#666" }} onClick={() => setAttSubTab("monthly")}>📊 Monthly Report</button>
             </div>
          </div>

          {attSubTab === "daily" && (
            <div style={card}>
               <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
                  <h3 style={{ margin: 0 }}>Daily Attendance Calendar</h3>
                  <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                     <label style={{ fontWeight: "bold", fontSize: 13 }}>Select Date:</label>
                     <input type="date" style={{ ...sInp, width: "auto" }} value={markDate} onChange={(e) => setMarkDate(e.target.value)} />
                     {ses.role === "a" && <button style={{ ...btn, background: "#1D9E75", color: "#fff", margin: 0 }} onClick={saveDailyAttendance}>Save Daily Attendance</button>}
                  </div>
               </div>
               
               <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead><tr style={{ background: "#f4f4f4", textAlign: "left" }}><th style={thS}>Emp ID</th><th style={thS}>Name</th><th style={thS}>Status</th><th style={thS}>Reason (Required for Leave/Absent)</th></tr></thead>
                  <tbody>
                     {emps.filter((e) => e.status === "Active" && e.id !== "admin" && (!ses.id || e.id === ses.id || ses.role === "a")).map(e => (
                        <tr key={e.id} style={{ borderBottom: "1px solid #eee" }}>
                           <td style={tdS}>{e.id}</td>
                           <td style={tdS}>{e.name}</td>
                           <td style={tdS}>
                              <select style={{ ...sInp, width: 120, background: dailyForm[e.id]?.status === "Present" ? "#e8f5e9" : dailyForm[e.id]?.status === "Holiday" ? "#e3f2fd" : "#ffebee" }} value={dailyForm[e.id]?.status || "Present"} disabled={ses.role !== "a"} onChange={(evt) => setDailyForm({...dailyForm, [e.id]: { ...dailyForm[e.id], status: evt.target.value }})}>
                                 <option>Present</option><option>WFH</option><option>Leave</option><option>Absent</option><option>Holiday</option>
                              </select>
                           </td>
                           <td style={tdS}>
                              {(dailyForm[e.id]?.status === "Leave" || dailyForm[e.id]?.status === "Absent") ? (
                                 <input style={{ ...sInp, border: "1px solid red" }} placeholder="Enter reason..." value={dailyForm[e.id]?.reason || ""} disabled={ses.role !== "a"} onChange={(evt) => setDailyForm({...dailyForm, [e.id]: { ...dailyForm[e.id], reason: evt.target.value }})} />
                              ) : <span style={{ color: "#aaa", fontSize: 11 }}>N/A</span>}
                           </td>
                        </tr>
                     ))}
                  </tbody>
               </table>

               {ses.role === "a" && (
                 <div style={{ marginTop: 30, borderTop: "2px solid #eee", paddingTop: 20 }}>
                   <h3 style={{ margin: "0 0 15px 0" }}>Daily Attendance Reports</h3>
                   <div style={{ display: "flex", gap: 15, flexWrap: "wrap" }}>
                     <div style={{ background: "#f8f9fa", padding: 15, borderRadius: 6, border: "1px solid #ddd", flex: 1, minWidth: 250 }}>
                       <h4 style={{ margin: "0 0 10px 0", color: "#185FA5" }}>For Selected Date: {fmtDate(markDate)}</h4>
                       <div style={{ display: "flex", gap: 10 }}>
                          <button style={{ ...btn, background: "#1D9E75", color: "#fff", flex: 1 }} onClick={() => {
                             const head = ["S.No", "Emp ID", "Employee", "Date", "Status", "Reason"];
                             const rows = []; let sno = 1;
                             emps.filter(e => e.id !== "admin").forEach(e => {
                               const dLog = dailyAtt.find(r => r.emp_id === e.id && r.date === markDate);
                               if (dLog) rows.push([sno++, e.id, e.name, fmtDate(markDate), dLog.status, dLog.reason || "-"]);
                             });
                             if (!rows.length) return alert("No records found for " + fmtDate(markDate));
                             exportExcel([head, ...rows], `Daily_Attendance_${markDate}`);
                          }}>📊 Excel</button>
                          <button style={{ ...btn, background: "#d32f2f", color: "#fff", flex: 1 }} onClick={() => {
                             const head = ["S.No", "Emp ID", "Employee", "Date", "Status", "Reason"];
                             const rows = []; let sno = 1;
                             emps.filter(e => e.id !== "admin").forEach(e => {
                               const dLog = dailyAtt.find(r => r.emp_id === e.id && r.date === markDate);
                               if (dLog) rows.push([sno++, e.id, e.name, fmtDate(markDate), dLog.status, dLog.reason || "-"]);
                             });
                             if (!rows.length) return alert("No records found for " + fmtDate(markDate));
                             setSlip(buildReportPdf("Daily Attendance Report", `Date: ${fmtDate(markDate)}`, head, rows));
                          }}>📄 PDF</button>
                       </div>
                     </div>
                     <div style={{ background: "#f8f9fa", padding: 15, borderRadius: 6, border: "1px solid #ddd", flex: 1, minWidth: 250 }}>
                       <h4 style={{ margin: "0 0 10px 0", color: "#8e44ad" }}>For Entire Month: {mo} {["Jan", "Feb", "Mar"].includes(mo) ? parseInt(fy) + 1 : fy}</h4>
                       <div style={{ display: "flex", gap: 10 }}>
                          <button style={{ ...btn, background: "#1D9E75", color: "#fff", flex: 1 }} onClick={() => {
                             const head = ["S.No", "Date", "Emp ID", "Employee", "Status", "Reason"];
                             const rows = []; let sno = 1;
                             const yNum = ["Jan", "Feb", "Mar"].includes(mo) ? parseInt(fy) + 1 : parseInt(fy);
                             const mIdxStr = String(CAL_MS.indexOf(mo) + 1).padStart(2, '0');
                             const prefix = `${yNum}-${mIdxStr}`;
                             const mLogs = dailyAtt.filter(r => r.date.startsWith(prefix)).sort((a,b) => a.date.localeCompare(b.date));
                             mLogs.forEach(log => {
                               const emp = emps.find(e => e.id === log.emp_id);
                               if (emp && emp.id !== "admin") rows.push([sno++, fmtDate(log.date), emp.id, emp.name, log.status, log.reason || "-"]);
                             });
                             if (!rows.length) return alert(`No daily records found for ${mo} ${yNum}. Ensure you have saved daily attendance for this month.`);
                             exportExcel([head, ...rows], `Daily_Logs_${mo}_${yNum}`);
                          }}>📊 Excel</button>
                          <button style={{ ...btn, background: "#d32f2f", color: "#fff", flex: 1 }} onClick={() => {
                             const head = ["S.No", "Date", "Emp ID", "Employee", "Status", "Reason"];
                             const rows = []; let sno = 1;
                             const yNum = ["Jan", "Feb", "Mar"].includes(mo) ? parseInt(fy) + 1 : parseInt(fy);
                             const mIdxStr = String(CAL_MS.indexOf(mo) + 1).padStart(2, '0');
                             const prefix = `${yNum}-${mIdxStr}`;
                             const mLogs = dailyAtt.filter(r => r.date.startsWith(prefix)).sort((a,b) => a.date.localeCompare(b.date));
                             mLogs.forEach(log => {
                               const emp = emps.find(e => e.id === log.emp_id);
                               if (emp && emp.id !== "admin") rows.push([sno++, fmtDate(log.date), emp.id, emp.name, log.status, log.reason || "-"]);
                             });
                             if (!rows.length) return alert(`No daily records found for ${mo} ${yNum}. Ensure you have saved daily attendance for this month.`);
                             setSlip(buildReportPdf("Monthly Log of Daily Attendance", `Reporting Period: ${mo} ${yNum}`, head, rows));
                          }}>📄 PDF</button>
                       </div>
                     </div>
                   </div>
                 </div>
               )}
            </div>
          )}

          {attSubTab === "monthly" && (
           <>
            <div style={{ display: "flex", gap: 10, marginBottom: 15 }}>
              {ses.role === "a" && <button style={{ ...btn, background: "#8e44ad", color: "#fff", margin: 0 }} onClick={runLeaveAccrual}>⚡ Auto-Calc (Pull from Daily Logs)</button>}
              {ses.role === "a" && <button style={{ ...btn, background: "#1a1a2e", color: "#fff", margin: 0 }} onClick={saveAttendance}>Save Monthly Report</button>}
            </div>
            <div style={{background: "#f3e5f5", color: "#8e44ad", padding: "10px", borderRadius: 4, marginBottom: 15, fontSize: 12, fontWeight: "bold"}}>
                ✓ Click 'Auto-Calc' to instantly count the Daily Calendar records for {mo} and populate this grid!
            </div>
            <div style={{ ...card, padding: 0, overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                <thead><tr style={{ background: "#f4f4f4", textAlign: "left", whiteSpace: "nowrap" }}><th style={thS}>S.No</th><th style={thS}>Emp ID</th><th style={thS}>Name</th><th style={thS}>Work Days</th><th>Present</th><th>Holidays</th><th>Leave</th><th>Balance</th><th>LOP (Days)</th><th>Comments</th>{ses.role === "a" && <th>Action</th>}</tr></thead>
                <tbody>{emps.filter((e) => ses.role === "a" ? isActiveInMonth(e, mo, fy) : e.id === ses.id).map((e, idx) => {
                  const a = att[e.id]?.[fy]?.[mo] || {};
                  const wDays = getWD(mo, fy);
                  const carryOver = getCarryOver(e.id);
                  const isEncashMonth = ["Jan", "Feb", "Mar"].includes(mo);
                  const hasEncash = a.comments?.toLowerCase().includes("encash");
                  const pureComment = a.comments ? a.comments.replace(/\d+\s*leaves?\s*encashed\s*\|?\s*/i, '').replace(/^\|\s*/, '').trim() : "";

                  return (
                    <tr key={e.id} style={{ borderBottom: "1px solid #eee" }}>
                      <td style={{ ...tdS, color: "#666" }}>{idx + 1}</td><td style={{ ...tdS, color: "#666" }}>{e.id}</td><td style={tdS}>{e.name}</td>
                      <td style={{ ...tdS, fontWeight: "bold", textAlign: "center", color: "#185FA5" }}>{wDays}</td>
                      <td><input style={{ ...sInp, minWidth: 60 }} disabled={ses.role !== "a"} type="number" value={a.present !== undefined && a.present !== null ? a.present : ""} onChange={(x) => updAtt(e.id, mo, "present", x.target.value)} /></td>
                      <td><input style={{ ...sInp, minWidth: 60, background: "#f0f0f0" }} disabled={ses.role !== "a"} type="number" value={a.holiday !== undefined && a.holiday !== null ? a.holiday : ""} onChange={(x) => updAtt(e.id, mo, "holiday", x.target.value)} title="Auto-filled from Daily Logs" /></td>
                      <td><input style={{ ...sInp, minWidth: 60, background: "#f0f0f0" }} disabled={ses.role !== "a"} type="number" value={a.leave !== undefined && a.leave !== null ? a.leave : ""} onChange={(x) => updAtt(e.id, mo, "leave", x.target.value)} title="Auto-filled from Daily Logs" /></td>
                      <td>
                        {e.leavePolicy === "No" ? (
                          <input style={{ ...sInp, minWidth: 60, background: "#f8f9fa", textAlign: "center" }} disabled value="NA" />
                        ) : (
                          <input style={{ ...sInp, minWidth: 60, background: "#f8f9fa" }} disabled={ses.role !== "a"} type="number" value={a.bal !== undefined && a.bal !== null ? a.bal : ""} onChange={(x) => updAtt(e.id, mo, "bal", x.target.value)} />
                        )}
                      </td>
                      <td><input style={{ ...sInp, minWidth: 60, background: "#ffebee" }} disabled={ses.role !== "a"} type="number" value={a.lop !== undefined && a.lop !== null ? a.lop : ""} onChange={(x) => updAtt(e.id, mo, "lop", x.target.value)} /></td>
                      
                      <td style={{minWidth: 200}}>
                        {isEncashMonth && carryOver > 0 && e.leavePolicy !== "No" && ses.role === "a" && (
                          <div style={{ background: "#fff8e1", padding: "4px 8px", borderRadius: 4, marginBottom: 5, border: "1px solid #ffc107" }}>
                            <div style={{ fontSize: 10, fontWeight: "bold", marginBottom: 3, color: "#f57c00" }}>Leave Encash (Max: {carryOver})</div>
                            <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                              <select style={{ padding: 2, fontSize: 11, borderRadius: 3, border: "1px solid #ccc" }} value={hasEncash ? "Yes" : "No"} onChange={(evt) => { if (evt.target.value === "Yes") { updAtt(e.id, mo, "comments", pureComment ? `${carryOver} leaves encashed | ${pureComment}` : `${carryOver} leaves encashed`); } else { updAtt(e.id, mo, "comments", pureComment); } }}>
                                <option>No</option><option>Yes</option>
                              </select>
                              {hasEncash && (
                                <input type="number" style={{ padding: 2, fontSize: 11, borderRadius: 3, border: "1px solid #ccc", width: 40 }} max={carryOver} min={0} value={a.comments?.match(/(\d+)\s*leave/i)?.[1] || ""} onChange={(evt) => { let v = Number(evt.target.value); if (v > carryOver) v = carryOver; if (v < 0) v = 0; updAtt(e.id, mo, "comments", pureComment ? `${v} leaves encashed | ${pureComment}` : `${v} leaves encashed`); }} />
                              )}
                            </div>
                          </div>
                        )}
                        <input style={{ ...sInp, width: "100%" }} disabled={ses.role !== "a"} placeholder="General Comments" value={pureComment} onChange={(x) => {
                          const v = x.target.value; const encashStr = hasEncash ? (a.comments.match(/\d+\s*leaves?\s*encashed/i)?.[0] || "") : "";
                          updAtt(e.id, mo, "comments", encashStr ? (v ? `${encashStr} | ${v}` : encashStr) : v);
                        }} />
                      </td>
                      {ses.role === "a" && <td><button style={{ ...btn, padding: "4px 8px", color: "#d32f2f", fontSize: 11, border: "1px solid #ffcdd2" }} onClick={() => { if(confirm("Clear this entire row?")) updAtt(e.id, mo, "CLEAR_ALL"); }}>Clear</button></td>}
                    </tr>
                  );
                })}</tbody>
              </table>
            </div>

            {ses.role === "a" && (
              <>
                <div style={{ display: "flex", justifyContent: "space-between", marginTop: 30, marginBottom: 15, alignItems: "center" }}><h3 style={{ margin: 0 }}>Leave Register ({fyL(fy)})</h3></div>
                <div style={{ ...card, padding: 0, overflowX: "auto", marginBottom: 40 }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead><tr style={{ background: "#1a1a2e", color: "#fff", textAlign: "left", whiteSpace: "nowrap" }}><th style={thS}>S.No</th><th style={thS}>Emp ID</th><th style={thS}>Employee</th>{MS.map((m) => <th key={m} style={thS}>{mL(m, fy)}</th>)}</tr></thead>
                    <tbody>
                      {emps.filter((e) => e.id !== "admin" && MS.some(m => isActiveInMonth(e, m, fy))).map((e, idx) => (
                        <tr key={e.id} style={{ borderBottom: "1px solid #eee" }}>
                          <td style={{ ...tdS, color: "#666" }}>{idx + 1}</td><td style={{ ...tdS, color: "#666" }}>{e.id}</td><td style={tdS}><b>{e.name}</b></td>
                          {MS.map((m) => { const isActive = isActiveInMonth(e, m, fy); const leaveVal = att[e.id]?.[fy]?.[m]?.leave; return <td key={m} style={tdS}>{(isActive && leaveVal > 0) ? `${leaveVal} L` : "-"}</td>; })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 15, alignItems: "center" }}>
                  <div style={{ display: "flex", gap: 10, alignItems: "center" }}><h3 style={{ margin: 0 }}>Attendance Ledger</h3><select style={{ ...sInp, width: 130, margin: 0, padding: "4px 8px" }} value={attLedgerMo} onChange={(e) => setAttLedgerMo(e.target.value)}><option value="">All Months</option>{MS.map((m) => <option key={m} value={m}>{mL(m, fy)}</option>)}</select></div>
                  <div style={{ display: "flex", gap: 10 }}>
                    <button style={{ ...btn, background: "#1D9E75", color: "#fff", margin: 0 }} onClick={() => {
                      const head = ["S.No", "Emp ID", "Employee", "Month", "Work Days", "Present", "Holidays", "Leave", "Balance", "LOP", "Comments"];
                      const rows = []; let sno = 1;
                      emps.filter((e) => e.id !== "admin").forEach((e) => MS.forEach((m) => {
                        if (attLedgerMo && m !== attLedgerMo) return;
                        const a = att[e.id]?.[fy]?.[m];
                        if (isActiveInMonth(e, m, fy) && a && (a.present !== null || a.leave !== null || a.holiday !== null || a.lop !== null || (a.comments && a.comments.trim() !== ""))) {
                          rows.push([sno++, e.id, e.name, mL(m, fy), getWD(m, fy), a.present !== null ? a.present : "-", a.holiday !== null ? a.holiday : "-", a.leave !== null ? a.leave : "-", e.leavePolicy === "No" ? "NA" : (a.bal !== null ? a.bal : "-"), a.lop !== null ? a.lop : "-", a.comments || ""]);
                        }
                      }));
                      exportExcel([head, ...rows], `Attendance_Ledger_${attLedgerMo || "All"}_${fyL(fy)}`);
                    }}>📊 Excel</button>
                    <button style={{ ...btn, background: "#d32f2f", color: "#fff", margin: 0 }} onClick={() => {
                      const head = ["S.No", "Emp ID", "Employee", "Month", "W.Days", "Present", "Hol", "Leave", "Bal", "LOP", "Comments"];
                      const rows = []; let sno = 1;
                      emps.filter((e) => e.id !== "admin").forEach((e) => MS.forEach((m) => {
                        if (attLedgerMo && m !== attLedgerMo) return;
                        const a = att[e.id]?.[fy]?.[m];
                        if (isActiveInMonth(e, m, fy) && a && (a.present !== null || a.leave !== null || a.holiday !== null || a.lop !== null || (a.comments && a.comments.trim() !== ""))) {
                          rows.push([sno++, e.id, e.name, mL(m, fy), getWD(m, fy), a.present !== null ? a.present : "-", a.holiday !== null ? a.holiday : "-", a.leave !== null ? a.leave : "-", e.leavePolicy === "No" ? "NA" : (a.bal !== null ? a.bal : "-"), a.lop !== null ? a.lop : "-", a.comments || ""]);
                        }
                      }));
                      setSlip(buildReportPdf("Attendance Ledger", `Reporting Period: ${attLedgerMo || "All Months"} FY ${fyL(fy)}`, head, rows));
                    }}>📄 PDF</button>
                  </div>
                </div>
                <div style={{ ...card, padding: 0, overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                    <thead><tr style={{ background: "#f4f6f8", textAlign: "left", whiteSpace: "nowrap" }}><th style={thS}>S.No</th><th style={thS}>Emp ID</th><th style={thS}>Employee</th><th style={thS}>Month</th><th style={thS}>Work Days</th><th style={thS}>Present</th><th style={thS}>Holidays</th><th style={thS}>Leave</th><th style={thS}>Balance</th><th style={thS}>LOP</th><th style={thS}>Comments</th><th style={thS}>Action</th></tr></thead>
                    <tbody>
                      {emps.filter((e) => e.id !== "admin")
                        .flatMap((e) => MS.filter(m => !attLedgerMo || m === attLedgerMo).map((m) => ({ e, m, a: att[e.id]?.[fy]?.[m] })))
                        .filter((x) => isActiveInMonth(x.e, x.m, fy) && x.a && (x.a.present !== null || x.a.leave !== null || x.a.holiday !== null || x.a.lop !== null || (x.a.comments && x.a.comments.trim() !== "")))
                        .map((item, idx) => (
                        <tr key={item.e.id + item.m} style={{ borderBottom: "1px solid #eee" }}>
                          <td style={{ ...tdS, color: "#666" }}>{idx + 1}</td><td style={{ ...tdS, color: "#666" }}>{item.e.id}</td><td style={tdS}><b>{item.e.name}</b></td><td style={tdS}>{mL(item.m, fy)}</td><td style={{ ...tdS, color: "#185FA5", fontWeight: "bold" }}>{getWD(item.m, fy)}</td>
                          <td style={tdS}>{item.a.present !== null ? item.a.present : "-"}</td><td style={tdS}>{item.a.holiday !== null ? item.a.holiday : "-"}</td><td style={tdS}>{item.a.leave !== null ? item.a.leave : "-"}</td><td style={tdS}>{item.e.leavePolicy === "No" ? "NA" : (item.a.bal !== null ? item.a.bal : "-")}</td><td style={tdS}>{item.a.lop !== null ? item.a.lop : "-"}</td><td style={{ ...tdS, fontSize: 11, color: "#666" }}>{item.a.comments || "-"}</td>
                          <td style={tdS}><button style={{ ...btn, padding: "2px 8px", color: "red", fontSize: 11 }} onClick={async () => { if (confirm("Delete this record permanently?")) { const { error } = await supabase.from("gits_attendance").delete().eq("emp_id", item.e.id).eq("fy", fy).eq("mo", item.m); if (error) alert("Database Error: " + error.message); else { updAtt(item.e.id, item.m, "CLEAR_ALL"); logAction("Attendance", `Deleted ledger record for ${item.e.id} in ${item.m}`); } } }}>Del</button></td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            )}
           </>
          )}
        </div>
      )}

      {tab === "ledger" && (
        <div>
          <div style={{ display: "flex", gap: 10, marginBottom: 20 }}>
            <select style={sInp} value={mo} onChange={(e) => setMo(e.target.value)}><option value="">All Months</option>{MS.map((m) => (<option key={m} value={m}>{mL(m, fy)}</option>))}</select>
            {ses.role === "a" && <select style={sInp} value={lEmp} onChange={(e) => setLEmp(e.target.value)}><option value="">All Employees</option>{emps.filter((e) => e.id !== "admin").map((e) => (<option key={e.id} value={e.id}>{e.name}</option>))}</select>}
            <div style={{ marginLeft: "auto", display: "flex", gap: 10 }}>
              {ses.role === "a" && <button style={{ ...btn, background: "#1a1a2e", color: "#fff" }} onClick={() => { setEditLedger(null); setNEn({ m: "Apr", t: "s", basic: 0, hra: 0, conv: 800, med: 1500, inc: 0, oth: 0, lop: 0, adv: 0, pt: 200, tds: 0, othD: 0, note: "" }); setShowAddEntry(true); }}>+ Add Entry</button>}
              <button style={{ ...btn, background: "#1D9E75", color: "#fff" }} onClick={() => {
                const head = ["S.No", "Emp ID", "Employee", "Month", "Type", "Basic", "HRA", "Conv", "Med", "Incentive", "Other Earnings", "Gross Amount", "LOP", "Staff Advance", "Prof Tax", "TDS", "Other Deductions", "Total Deductions", "Taxable Income", "Net Amount", "Note/Comments"];
                const rows = [];
                const tEmps = ses.role === "a" ? (lEmp ? emps.filter((e) => e.id === lEmp) : emps.filter((e) => e.id !== "admin")) : [myE];
                let sno = 1;
                tEmps.forEach((e) => {
                  if (!e) return;
                  (pay[e.id]?.[fy] || []).filter((r) => !mo || r.m === mo).forEach((r) => rows.push([sno++, e.id, e.name, mL(r.m, fy), r.t === "s" ? "Salary" : "Incentive", r.basic || 0, r.hra || 0, r.conv || 0, r.med || 0, r.inc || 0, r.oth || 0, gr(r), r.lop || 0, r.adv || 0, r.pt || 0, r.tds || 0, r.othD || 0, dd(r), txInc(r), np(r), r.note || ""]));
                });
                exportExcel([head, ...rows], `Ledger_${fyL(fy)}`);
              }}>📊 Excel</button>
              <button style={{ ...btn, background: "#d32f2f", color: "#fff" }} onClick={() => {
                const head = ["S.No", "Emp ID", "Employee", "Month", "Basic", "Gross", "LOP", "Adv", "PT", "TDS", "Oth Ded", "Tot Ded", "Net", "Note"];
                const rows = [];
                const tEmps = ses.role === "a" ? (lEmp ? emps.filter((e) => e.id === lEmp) : emps.filter((e) => e.id !== "admin")) : [myE];
                let sno = 1;
                tEmps.forEach((e) => {
                  if (!e) return;
                  (pay[e.id]?.[fy] || []).filter((r) => !mo || r.m === mo).forEach((r) => rows.push([sno++, e.id, e.name, mL(r.m, fy), f$(r.basic), f$(gr(r)), f$(r.lop), f$(r.adv), f$(r.pt), f$(r.tds), f$(r.othD || 0), f$(dd(r)), f$(np(r)), r.note || "-"]));
                });
                setSlip(buildReportPdf("Financial Ledger Report", `Reporting Period: FY ${fyL(fy)}`, head, rows));
              }}>📄 PDF</button>
            </div>
          </div>
          {showAddEntry && ses.role === "a" && (
            <div style={{ ...card, border: "1px solid #1a1a2e", marginBottom: 20 }}>
              <h4 style={{ marginTop: 0 }}>{editLedger ? "Edit Ledger Entry" : "Add Manual Ledger Entry"}</h4>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(120px,1fr))", gap: 10, marginBottom: 15 }}>
                <div><label style={lbl}>Employee</label><select style={sInp} value={pEmp} disabled={!!editLedger} onChange={(e) => {
                  const id = e.target.value; setPEmp(id);
                  if (id) { 
                    const l = getLastPay(id); const a = att[id]?.[fy]?.[nEn.m];
                    const calDays = getCalendarDays(nEn.m, fy);
                    const corePay = (l.basic || 0) + (l.hra || 0) + (l.conv || 0) + (l.med || 0);
                    const lopDays = Number(a?.lop || 0);
                    let lopAmt = lopDays > 0 ? Math.round((corePay / calDays) * lopDays) : 0;
                    let encashAmt = 0; let autoNote = "";
                    if (a?.comments && a.comments.toLowerCase().includes("encash")) {
                        const match = a.comments.match(/(\d+)\s*leave/i);
                        if (match) { const eDays = Number(match[1]); encashAmt = Math.round((corePay / 30) * eDays); autoNote = `${eDays} leaves encashed`; }
                    }
                    if (lopDays > 0) autoNote = autoNote ? `${autoNote} | ${lopDays} days LOP` : `${lopDays} days LOP`;
                    setNEn((prev) => ({ ...prev, basic: l.basic, hra: l.hra, conv: l.conv, med: l.med, pt: l.pt, tds: l.tds, lop: lopAmt, oth: encashAmt, note: autoNote })); 
                  }
                }}><option value="">Select...</option>{emps.filter((e) => e.id !== "admin").map((e) => (<option key={e.id} value={e.id}>{e.name}</option>))}</select></div>
                <div><label style={lbl}>Month</label><select style={sInp} value={nEn.m} onChange={(e) => setNEn({ ...nEn, m: e.target.value })}>{MS.map((m) => (<option key={m} value={m}>{mL(m, fy)}</option>))}</select></div>
                <div><label style={lbl}>Type</label><select style={sInp} value={nEn.t} onChange={(e) => setNEn({ ...nEn, t: e.target.value })}><option value="s">Salary</option><option value="i">Incentive</option></select></div>
                {[["Basic", "basic"], ["HRA", "hra"], ["Conv", "conv"], ["Med", "med"], ["Incentive", "inc"], ["Other Earn", "oth"], ["LOP (₹)", "lop"], ["Advance", "adv"], ["PT", "pt"], ["TDS", "tds"], ["Other Ded", "othD"]].map(([l, k]) => (<div key={k}><label style={lbl}>{l}</label><input style={sInp} type="number" value={nEn[k]} onChange={(e) => {
                    const v = e.target.value; const nextEn = { ...nEn, [k]: v };
                    if (["basic", "hra", "conv", "med"].includes(k) && pEmp) {
                        const calDays = getCalendarDays(nextEn.m, fy);
                        const corePay = (+nextEn.basic || 0) + (+nextEn.hra || 0) + (+nextEn.conv || 0) + (+nextEn.med || 0);
                        const a = att[pEmp]?.[fy]?.[nextEn.m];
                        const lopDays = Number(a?.lop || 0);
                        if (lopDays > 0) nextEn.lop = Math.round((corePay / calDays) * lopDays);
                        if (a?.comments && a.comments.toLowerCase().includes("encash")) {
                            const match = a.comments.match(/(\d+)\s*leave/i);
                            if (match) { const eDays = Number(match[1]); nextEn.oth = Math.round((corePay / 30) * eDays); }
                        }
                    }
                    setNEn(nextEn);
                }} /></div>))}
                <div style={{ gridColumn: "1/-1" }}><label style={lbl}>Note / Comments</label><input style={sInp} value={nEn.note} placeholder="Optional reason for entry" onChange={(e) => setNEn({ ...nEn, note: e.target.value })} /></div>
              </div>
              <button style={{ ...btn, background: "green", color: "#fff" }} onClick={addLedgerEntry}>{editLedger ? "Update Entry" : "Save Entry"}</button><button style={{ ...btn, marginLeft: 10 }} onClick={() => { setShowAddEntry(false); setEditLedger(null); }}>Cancel</button>
            </div>
          )}
          <div style={{ ...card, padding: 0, overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
              <thead style={{ background: "#1a1a2e", color: "#fff" }}><tr style={{ textAlign: "left", whiteSpace: "nowrap" }}>
                <th style={thS}>S.No</th><th style={thS}>Emp ID</th><th style={thS}>Employee</th><th style={thS}>Month</th><th style={thS}>Basic</th><th style={thS}>HRA</th><th style={thS}>Conv</th><th style={thS}>Med</th><th style={thS}>Inc</th><th style={thS}>Oth Earn</th><th style={thS}>Gross</th><th style={thS}>LOP</th><th style={thS}>Adv</th><th style={thS}>PT</th><th style={thS}>TDS</th><th style={thS}>Oth Ded</th><th style={thS}>Tot. Ded</th><th style={thS}>Taxable</th><th style={thS}>Net</th><th style={thS}>Note</th>{ses.role === "a" && <th style={thS}>Action</th>}
              </tr></thead>
              <tbody>
                {emps.filter((e) => ses.role === "a" ? (e.id !== "admin" && (!lEmp || e.id === lEmp)) : e.id === ses.id)
                  .flatMap((e) => (pay[e.id]?.[fy] || []).filter((r) => !mo || r.m === mo).map((r, i) => ({ e, r, i })))
                  .map((item, idx) => (
                    <tr key={item.e.id + item.i} style={{ borderBottom: "1px solid #eee" }}>
                      <td style={{ ...tdS, color: "#666" }}>{idx + 1}</td>
                      <td style={{ ...tdS, color: "#666" }}>{item.e.id}</td>
                      <td style={tdS}><b>{item.e.name}</b></td>
                      <td style={tdS}>{mL(item.r.m, fy)}</td>
                      <td style={tdS}>{f$(item.r.basic)}</td><td style={tdS}>{f$(item.r.hra)}</td><td style={tdS}>{f$(item.r.conv)}</td><td style={tdS}>{f$(item.r.med)}</td><td style={tdS}>{f$(item.r.inc)}</td><td style={tdS}>{f$(item.r.oth)}</td><td style={{ ...tdS, fontWeight: "bold" }}>{f$(gr(item.r))}</td><td style={tdS}>{f$(item.r.lop)}</td><td style={tdS}>{f$(item.r.adv)}</td><td style={tdS}>{f$(item.r.pt)}</td><td style={tdS}>{f$(item.r.tds)}</td><td style={tdS}>{f$(item.r.othD || 0)}</td><td style={{ ...tdS, color: "#D85A30" }}>{f$(dd(item.r))}</td><td style={{ ...tdS, fontWeight: "bold" }}>{f$(txInc(item.r))}</td><td style={{ ...tdS, color: "#1D9E75", fontWeight: "bold", fontSize: 12 }}>{f$(np(item.r))}</td><td style={{ ...tdS, fontSize: 11, color: "#666" }}>{item.r.note || "-"}</td>
                      {ses.role === "a" && <td style={tdS}>
                        <div style={{ display: "flex", gap: 5 }}>
                          <button style={{ ...btn, padding: "4px 8px" }} onClick={() => { setPEmp(item.e.id); setNEn({ ...item.r }); setEditLedger({ eid: item.e.id, idx: item.i, db_id: item.r.db_id }); setShowAddEntry(true); }}>Edit</button>
                          <button style={{ ...btn, padding: "4px 8px", color: "red" }} onClick={() => delLedgerEntry(item.e.id, item.i, item.r.db_id)}>Del</button>
                        </div>
                      </td>}
                    </tr>
                  ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === "payslips" && (
        <div style={card}>
          <h3 style={{ marginTop: 0, marginBottom: 20 }}>Payslip Generator</h3>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 15, marginBottom: 20 }}>
            {ses.role === "a" && <div><label style={lbl}>Select Staff</label><select style={sInp} value={pEmp} onChange={(e) => setPEmp(e.target.value)}><option value="">Select...</option>{emps.filter((e) => e.id !== "admin").map((e) => (<option key={e.id} value={e.id}>{e.name}</option>))}</select></div>}
            <div><label style={lbl}>Start Year</label><input type="number" style={sInp} value={pFy} onChange={(e) => setPFy(e.target.value)} /></div>
            <div><label style={lbl}>Month</label><select style={sInp} value={pMo} onChange={(e) => setPMo(e.target.value)}>{MS.map((m) => (<option key={m} value={m}>{mL(m, pFy)}</option>))}</select></div>
          </div>
          
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <button style={{ padding: "12px 24px", background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 6, fontWeight: "bold", cursor: "pointer" }} onClick={() => {
              const t = ses.role === "a" ? pEmp : ses.id; 
              if (!t) return alert("Select an employee first.");
              const ents = (pay[t]?.[pFy] || []).filter((x) => x.m === pMo);
              if (ents.length) setSlip(generatePayslipsHtml([{ emp: emps.find((e) => e.id === t), mo: pMo, entries: ents, att: att[t]?.[pFy]?.[pMo] }], pFy)); 
              else alert("No record found");
            }}>📄 Generate Single Payslip</button>

            {ses.role === "a" && (
              <>
                <button style={{ padding: "12px 24px", background: "#185FA5", color: "#fff", border: "none", borderRadius: 6, fontWeight: "bold", cursor: "pointer" }} onClick={() => {
                  const slips = [];
                  emps.filter(e => e.id !== "admin").forEach(e => {
                      const ents = (pay[e.id]?.[pFy] || []).filter((x) => x.m === pMo);
                      if(ents.length) slips.push({ emp: e, mo: pMo, entries: ents, att: att[e.id]?.[pFy]?.[pMo] });
                  });
                  if (slips.length) setSlip(generatePayslipsHtml(slips, pFy));
                  else alert("No records found for this month");
                }}>📑 Bulk Generate (All Staff, This Month)</button>

                <button style={{ padding: "12px 24px", background: "#8e44ad", color: "#fff", border: "none", borderRadius: 6, fontWeight: "bold", cursor: "pointer" }} onClick={() => {
                  const slips = [];
                  emps.filter(e => e.id !== "admin").forEach(e => {
                      MS.forEach(m => {
                          const ents = (pay[e.id]?.[pFy] || []).filter((x) => x.m === m);
                          if(ents.length) slips.push({ emp: e, mo: m, entries: ents, att: att[e.id]?.[pFy]?.[m] });
                      });
                  });
                  if (slips.length) setSlip(generatePayslipsHtml(slips, pFy));
                  else alert("No records found for this year");
                }}>📚 Bulk Generate (All Staff, Entire Year)</button>
              </>
            )}
          </div>
        </div>
      )}

      {tab === "cashbook" && ses.role === "a" && (
        <div>
          <div style={{ display: "flex", gap: 10, marginBottom: 20, alignItems: "center" }}>
            <h3 style={{ margin: 0 }}>Company Financials</h3>
            <div style={{ display: "flex", gap: 5, alignItems: "center", background: "#f8f9fa", padding: "4px 8px", borderRadius: 6, border: "1px solid #ddd", marginLeft: 20 }}>
                <span style={{fontSize: 11, fontWeight: "bold", color: "#666"}}>PAYROLL SYNC:</span>
                <select style={{...sInp, width: 70, padding: "4px", minHeight: 0}} value={syncMo} onChange={e => setSyncMo(e.target.value)}>
                    {MS.map(m => <option key={m} value={m}>{m}</option>)}
                </select>
                <button 
                    style={{ ...btn, background: isPayrollSynced ? "#ccc" : "#8e44ad", color: "#fff", padding: "4px 10px", cursor: isPayrollSynced ? "not-allowed" : "pointer" }} 
                    disabled={isPayrollSynced} onClick={() => postPayrollToCashbook(syncMo)}
                >
                    {isPayrollSynced ? `✅ Synced` : `⬇️ Pull Data`}
                </button>
            </div>

            <div style={{ marginLeft: "auto", display: "flex", gap: 10, alignItems: "center" }}>
              <select style={sInp} value={cbSelFY} onChange={(e) => setCbSelFY(e.target.value)}>
                {cbFYs.map(f => <option key={f.id} value={f.id}>{f.label}</option>)}
              </select>
              <button style={{ ...btn, background: "#1a1a2e", color: "#fff" }} onClick={() => setShowAddCbFy(true)}>+ Add FY</button>
            </div>
          </div>

          {showAddCbFy && (
            <div style={{ ...card, border: "1px solid #1a1a2e", marginBottom: 20 }}>
              <h4 style={{ marginTop: 0 }}>Add New Cashbook Financial Year</h4>
              <div style={{ display: "flex", gap: 15, alignItems: "flex-end" }}>
                <div><label style={lbl}>Start Date</label><input type="date" style={sInp} value={newCbFyStart} onChange={e => setNewCbFyStart(e.target.value)} /></div>
                <div><label style={lbl}>End Date</label><input type="date" style={sInp} value={newCbFyEnd} onChange={e => setNewCbFyEnd(e.target.value)} /></div>
                <button style={{ ...btn, background: "green", color: "#fff" }} onClick={saveCbFy}>Save</button>
                <button style={btn} onClick={() => setShowAddCbFy(false)}>Cancel</button>
              </div>
            </div>
          )}

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 20, textAlign: "center", marginBottom: 20 }}>
            <div style={{ ...card, borderTop: "4px solid #185FA5" }}><div style={lbl}>Opening Balance</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#185FA5" }}>{f$(openingBalance)}</div></div>
            <div style={{ ...card, borderTop: "4px solid #0F9D6D" }}><div style={lbl}>Total Receipts</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#0F9D6D" }}>{f$(cbReceipts)}</div></div>
            <div style={{ ...card, borderTop: "4px solid #C23A55" }}><div style={lbl}>Payments & Expenses</div><div style={{ fontSize: 24, fontWeight: "bold", color: "#C23A55" }}>{f$(cbPayments + cbExpenses)}</div></div>
            <div style={{ ...card, borderTop: closingBalance >= 0 ? "4px solid #10243E" : "4px solid #C23A55" }}><div style={lbl}>Closing Balance</div><div style={{ fontSize: 24, fontWeight: "bold", color: closingBalance >= 0 ? "#10243E" : "#C23A55" }}>{f$(closingBalance)}</div></div>
          </div>

          <div style={{ display: "flex", gap: 20, alignItems: "flex-start", flexWrap: "wrap" }}>
            
            <div style={{ ...card, flex: "1 1 300px", position: "sticky", top: 20 }}>
              <h4 style={{ marginTop: 0, borderBottom: "1px solid #eee", paddingBottom: 10 }}>{editCbTxn ? "Edit Transaction" : "New Transaction"}</h4>
              <div style={{ display: "flex", gap: 5, marginBottom: 15, background: "#f1f2ee", padding: 5, borderRadius: 6 }}>
                {Object.entries(TYPE_META).map(([k, m]) => (
                  <button key={k} onClick={() => setCbType(k)} style={{ flex: 1, padding: "8px 0", border: "none", borderRadius: 4, cursor: "pointer", fontSize: 12, fontWeight: "bold", background: cbType === k ? "#fff" : "transparent", color: cbType === k ? m.color : "#666", boxShadow: cbType === k ? "0 1px 3px rgba(0,0,0,0.1)" : "none" }}>{m.label}</button>
                ))}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 15 }}>
                <div><label style={lbl}>Party Name</label><input style={sInp} value={cbParty} onChange={e => setCbParty(e.target.value)} placeholder="e.g. Client or Vendor" /></div>
                <div style={{ display: "flex", gap: 10 }}>
                  <div style={{ flex: 1 }}><label style={lbl}>Date</label><input type="date" style={sInp} value={cbDate} onChange={e => setCbDate(e.target.value)} /></div>
                  <div style={{ flex: 1 }}><label style={lbl}>Amount (₹)</label><input type="number" style={sInp} value={cbAmount} onChange={e => setCbAmount(e.target.value)} placeholder="0.00" /></div>
                </div>
                <div><label style={lbl}>Description / Notes</label><textarea style={{ ...sInp, resize: "none" }} rows={3} value={cbDesc} onChange={e => setCbDesc(e.target.value)} placeholder="Optional" /></div>
                <div style={{display: "flex", gap: 5, marginTop: 5}}>
                    <button style={{ ...btn, flex: 1, background: TYPE_META[cbType].color, color: "#fff", fontWeight: "bold", padding: 12 }} onClick={saveCbTxn}>{editCbTxn ? "Update Entry" : `Save ${TYPE_META[cbType].label}`}</button>
                    {editCbTxn && <button style={{ ...btn, flex: 1, padding: 12 }} onClick={() => { setEditCbTxn(null); setCbParty(""); setCbAmount(""); setCbDesc(""); }}>Cancel</button>}
                </div>
              </div>
            </div>

            <div style={{ ...card, flex: "2 1 600px", padding: 0 }}>
              <div style={{ padding: "15px 20px", display: "flex", gap: 10, background: "#fafafa", borderBottom: "1px solid #eee", flexWrap: "wrap" }}>
                {["all", "receipt", "payment", "expense"].map(t => (
                  <button key={t} onClick={() => setCbTabState(t)} style={{ ...btn, background: cbTabState === t ? "#1a1a2e" : "#fff", color: cbTabState === t ? "#fff" : "#444", borderRadius: 20 }}>{t === "all" ? "All Entries" : TYPE_META[t]?.label + "s"}</button>
                ))}
                <input style={{ ...sInp, width: 200, marginLeft: "auto", borderRadius: 20 }} placeholder="Search party, voucher..." value={cbSearch} onChange={e => setCbSearch(e.target.value)} />
              </div>
              <div style={{ padding: "10px 20px", display: "flex", justifyContent: "flex-end", gap: 10 }}>
                  <button style={{ ...btn, background: "#1D9E75", color: "#fff" }} onClick={() => {
                    if(!cbFilteredTxns.length) return alert("No data");
                    const head = ["Date", "Voucher No", "Type", "Party", "Description", "Amount (INR)"];
                    const rows = [];
                    if (isFullLedger) rows.push([fmtDate(selectedFyObj?.start_date), "OPENING", "-", "-", "Opening Balance", Number(openingBalance)]);
                    cbFilteredTxns.forEach(t => rows.push([fmtDate(t.txn_date), t.voucher_no || "-", TYPE_META[t.type]?.label || t.type, t.party_name, t.description || "", Number(t.amount)]));
                    if (isFullLedger) rows.push([fmtDate(selectedFyObj?.end_date), "CLOSING", "-", "-", "Closing Balance", Number(closingBalance)]);
                    exportExcel([head, ...rows], `Cashbook_${cbFYs.find(f=>f.id===cbSelFY)?.label}`);
                  }}>📊 Excel</button>
                  <button style={{ ...btn, background: "#d32f2f", color: "#fff" }} onClick={() => {
                     if(!cbFilteredTxns.length) return alert("No data");
                     const head = ["Date", "Voucher", "Type", "Party", "Description", "Amount (INR)"];
                     const rows = [];
                     if (isFullLedger) rows.push([fmtDate(selectedFyObj?.start_date), "OPENING", "-", "-", "Opening Balance", f$(openingBalance)]);
                     cbFilteredTxns.forEach(t => rows.push([fmtDate(t.txn_date), t.voucher_no || "-", TYPE_META[t.type]?.label || t.type, t.party_name, t.description || "", f$(t.amount)]));
                     if (isFullLedger) rows.push([fmtDate(selectedFyObj?.end_date), "CLOSING", "-", "-", "Closing Balance", f$(closingBalance)]);
                     setSlip(buildReportPdf("Financial Cashbook Report", `Reporting Period: ${cbFYs.find(f=>f.id===cbSelFY)?.label}`, head, rows));
                  }}>📄 PDF</button>
              </div>

              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                  <thead><tr style={{ textAlign: "left", background: "#f8f9fa" }}>
                    <th style={thS}>Date</th><th style={thS}>Voucher No</th><th style={thS}>Type</th><th style={thS}>Party</th><th style={thS}>Description</th><th style={{ ...thS, textAlign: "right" }}>Amount</th><th style={thS}></th>
                  </tr></thead>
                  <tbody>
                    {isFullLedger && (
                        <tr style={{ background: "#eef2f6" }}>
                           <td style={tdS}>{fmtDate(selectedFyObj?.start_date)}</td>
                           <td style={{ ...tdS, fontWeight: "bold", color: "#185FA5" }}>OPENING</td>
                           <td style={tdS}>-</td><td style={tdS}>-</td>
                           <td style={{ ...tdS, fontWeight: "bold", color: "#185FA5" }}>Opening Balance</td>
                           <td style={{ ...tdS, textAlign: "right", fontWeight: "bold", color: "#185FA5" }}>{f$(openingBalance)}</td>
                           <td style={tdS}></td>
                        </tr>
                    )}
                    
                    {!cbFilteredTxns.length ? <tr><td colSpan={7} style={{ padding: 30, textAlign: "center", color: "#888" }}>No records found.</td></tr> : cbFilteredTxns.map(t => (
                      <tr key={t.id} style={{ borderBottom: "1px solid #eee" }}>
                        <td style={tdS}>{fmtDate(t.txn_date)}</td>
                        <td style={{ ...tdS, fontWeight: "bold", color: "#666" }}>{t.voucher_no || "-"}</td>
                        <td style={tdS}><span style={{ padding: "3px 8px", borderRadius: 12, fontSize: 10, fontWeight: "bold", background: TYPE_META[t.type]?.bg, color: TYPE_META[t.type]?.color }}>{TYPE_META[t.type]?.label || t.type}</span></td>
                        <td style={{ ...tdS, fontWeight: "bold" }}>{t.party_name}</td>
                        <td style={{ ...tdS, color: "#666" }}>{t.description || "-"}</td>
                        <td style={{ ...tdS, textAlign: "right", fontWeight: "bold", color: TYPE_META[t.type]?.color }}>{f$(t.amount)}</td>
                        <td style={tdS}>
                          <div style={{ display: "flex", gap: 5 }}>
                            <button style={{ ...btn, padding: "2px 6px" }} onClick={() => setSlip(generateVoucherPdf(t))}>📄 Print</button>
                            <button style={{ ...btn, padding: "2px 6px" }} onClick={() => {
                                setEditCbTxn(t); setCbType(t.type); setCbParty(t.party_name); setCbDate(t.txn_date); setCbAmount(t.amount); setCbDesc(t.description || "");
                            }}>Edit</button>
                            <button style={{ ...btn, color: "red", padding: "2px 6px" }} onClick={() => delCbTxn(t.id, t.voucher_no)}>Del</button>
                          </div>
                        </td>
                      </tr>
                    ))}

                    {isFullLedger && (
                        <tr style={{ background: "#e8f5e9" }}>
                           <td style={tdS}>{fmtDate(selectedFyObj?.end_date)}</td>
                           <td style={{ ...tdS, fontWeight: "bold", color: "#1D9E75" }}>CLOSING</td>
                           <td style={tdS}>-</td><td style={tdS}>-</td>
                           <td style={{ ...tdS, fontWeight: "bold", color: "#1D9E75" }}>Closing Balance</td>
                           <td style={{ ...tdS, textAlign: "right", fontWeight: "bold", color: "#1D9E75" }}>{f$(closingBalance)}</td>
                           <td style={tdS}></td>
                        </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {tab === "audit" && ses.role === "a" && (
        <div style={{ ...card, padding: 0 }}>
          <div style={{ padding: 20, display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid #eee" }}>
            <h3 style={{ margin: 0 }}>System Audit Logs</h3>
            <button style={{ ...btn, background: "#1D9E75", color: "#fff", border: "none" }} onClick={() => {
              const head = ["Log ID", "Date & Time", "Admin ID", "Module", "Action / Changes"];
              const rows = auditLogs.map(l => [l.id, new Date(l.created_at).toLocaleString('en-IN'), l.admin_id, l.module, l.description]);
              exportExcel([head, ...rows], `Audit_Logs_${new Date().toISOString().split('T')[0]}`);
            }}>📊 Download Logs (Excel)</button>
          </div>
          <div style={{ overflowX: "auto", maxHeight: 600 }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
              <thead style={{ position: "sticky", top: 0, zIndex: 1 }}>
                <tr style={{ background: "#f4f4f4", textAlign: "left", whiteSpace: "nowrap" }}>
                  <th style={thS}>Log ID</th><th style={thS}>Timestamp (IST)</th><th style={thS}>Admin ID</th><th style={thS}>Module</th><th style={thS}>Description of Change</th>
                </tr>
              </thead>
              <tbody>
                {auditLogs.length === 0 ? (
                  <tr><td colSpan="5" style={{ padding: 20, textAlign: "center", color: "#888" }}>No audit records found.</td></tr>
                ) : (
                  auditLogs.map((log) => (
                    <tr key={log.id || log.created_at} style={{ borderBottom: "1px solid #eee", background: log.module === "Ledger" ? "#fbfdff" : log.module === "Employees" ? "#fefdfa" : (log.module === "Cashbook" ? "#f0fdf4" : "#fff") }}>
                      <td style={{ ...tdS, color: "#999" }}>{log.id || "-"}</td>
                      <td style={{ ...tdS, fontWeight: "bold" }}>{new Date(log.created_at).toLocaleString('en-IN')}</td>
                      <td style={tdS}>{log.admin_id}</td>
                      <td style={tdS}>
                        <span style={{ padding: "3px 8px", borderRadius: 12, fontSize: 10, fontWeight: "bold", background: log.module === "Ledger" ? "#e3f2fd" : log.module === "Employees" ? "#fff3e0" : (log.module === "Cashbook" ? "#e8f5e9" : "#f3e5f5"), color: log.module === "Ledger" ? "#1565c0" : log.module === "Employees" ? "#ef6c00" : (log.module === "Cashbook" ? "#2e7d32" : "#7b1fa2") }}>
                          {log.module}
                        </span>
                      </td>
                      <td style={{ ...tdS, whiteSpace: "normal" }}>{log.description}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab === "profile" && (
        <div style={{ ...card, maxWidth: 600, margin: "auto" }}>
          <h3 style={{ marginTop: 0, borderBottom: "1px solid #eee", paddingBottom: 10 }}>My Profile & Security</h3>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20, marginTop: 20 }}>
            <div style={{ gridColumn: "1/-1", padding: 15, background: "#f4f6f8", borderRadius: 6, marginBottom: 10 }}>
              <p style={{ margin: 0, fontSize: 14 }}><b>Name:</b> {myE?.name || "System Admin"}</p>
              <p style={{ margin: "5px 0 0 0", fontSize: 14 }}><b>ID:</b> {myE?.id || "admin"}</p>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 15 }}>
              <h4 style={{ margin: 0 }}>Change Password</h4>
              <div><label style={lbl}>Current Password</label><input type="password" style={sInp} value={curPwd} onChange={(e) => setCurPwd(e.target.value)} /></div>
              <div><label style={lbl}>New Password</label><input type="password" style={sInp} value={newPwd} onChange={(e) => setNewPwd(e.target.value)} /></div>
              <div><label style={lbl}>Confirm New Password</label><input type="password" style={sInp} value={confPwd} onChange={(e) => setConfPwd(e.target.value)} /></div>
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 15 }}>
              <h4 style={{ margin: 0 }}>Recovery Question</h4>
              <p style={{ fontSize: 11, color: "#666", margin: 0 }}>Set this up to reset your password yourself if you forget it.</p>
              <div>
                <label style={lbl}>Security Question</label>
                <select style={sInp} value={secQ} onChange={(e) => setSecQ(e.target.value)}>
                  <option value="">Select a question...</option>
                  {SEC_QS.map((q) => (<option key={q} value={q}>{q}</option>))}
                </select>
              </div>
              <div><label style={lbl}>Your Answer</label><input type="text" style={sInp} value={secA} onChange={(e) => setSecA(e.target.value)} placeholder="Answer is not case sensitive" /></div>
            </div>
            <button style={{ gridColumn: "1/-1", padding: "12px", background: "#1a1a2e", color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", marginTop: 10, fontWeight: "bold" }} onClick={handleUpdateProfile}>Save Security Settings</button>
          </div>
        </div>
      )}
    </div>
  );
}

const card = { padding: 25, border: "1px solid #eee", borderRadius: 8, background: "#fff", boxShadow: "0 2px 8px rgba(0,0,0,0.04)", marginBottom: 20 };
const thS = { padding: 12, fontWeight: 500, whiteSpace: "nowrap", borderBottom: "2px solid #ddd" };
const tdS = { padding: 12, whiteSpace: "nowrap" };
const sInp = { padding: "8px 12px", border: "1px solid #ddd", borderRadius: 4, width: "100%", fontSize: 13, background: "#fff" };
const btn = { padding: "8px 16px", background: "#fff", border: "1px solid #ddd", borderRadius: 4, cursor: "pointer", fontSize: 12 };
const lbl = { display: "block", fontSize: 11, color: "#666", marginBottom: 4, fontWeight: "bold" };