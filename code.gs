var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ'; 
var SHEET_NAME = 'm_actionplan';
var APP_VERSION = '690127-1400'; 

function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  template.appVersion = getVersion(); 
  return template.evaluate()
      .setTitle('AP 2569 MONITORING (v.' + template.appVersion + ')') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getVersion() { return APP_VERSION; }

// 1. DATA LOADER
function getAllMasterDataForClient() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var getIdx = (name) => headers.findIndex(h => String(h).trim() === name);
    
    var iID = getIdx('รหัสโครงการ'); var iOrder = getIdx('ลำดับโครงการ'); var iDept = getIdx('กลุ่มงาน/งาน');
    var iProj = getIdx('โครงการ'); var iAct = getIdx('กิจกรรมหลัก'); var iSub = getIdx('กิจกรรมย่อย');
    var iType = getIdx('ประเภทงบ'); var iSource = getIdx('แหล่งงบประมาณ'); var iAlloc = getIdx('จัดสรร');
    var iBal = getIdx('คงเหลือ (ไม่รวมเงินยืม)'); 
    if(iBal == -1) iBal = getIdx('คงเหลือ');
    var iLoan = getIdx('เงินยืม');

    return data.map(r => ({
      id: r[iID], order: r[iOrder], dept: r[iDept], project: r[iProj], activity: r[iAct], subActivity: r[iSub],
      budgetType: r[iType], budgetSource: r[iSource], allocated: r[iAlloc], balance: r[iBal], loan: (iLoan > -1) ? r[iLoan] : 0
    })).filter(r => r.id && r.project); 
  } catch (e) { return []; }
}

// 2. DASHBOARD LOGIC
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: "ไม่พบชีต " + SHEET_NAME };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { error: "ตารางข้อมูลว่างเปล่า" };

    var headers = data.shift(); 
    var getIdx = function(name) { return headers.findIndex(h => String(h).trim() === name); };
    
    var idxType = getIdx('ประเภทงบ'); var idxApproved = getIdx('อนุมัติตามแผน');
    var idxAllocated = getIdx('จัดสรร'); var idxSpent = getIdx('เบิกจ่าย'); 
    var idxBalance = getIdx('คงเหลือ (ไม่รวมเงินยืม)');
    if (idxBalance == -1) idxBalance = getIdx('คงเหลือ');
    var idxDept = getIdx('กลุ่มงาน/งาน');

    if (idxType == -1 || idxApproved == -1 || idxAllocated == -1 || idxSpent == -1) {
      return { error: "ไม่พบหัวตาราง Dashboard" };
    }

    var summary = { moph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} }, nonMoph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} } };
    var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };

    data.forEach(function(row) {
      var typeVal = String(row[idxType] || "").trim();
      var isMoph = (typeVal.includes('งบประมาณ') || typeVal.includes('สป.สธ') || typeVal === 'PP' || typeVal === 'OP' || typeVal.includes('งบดำเนินงาน')); 
      var target = isMoph ? summary.moph : summary.nonMoph;
      
      target.approved += parseNum(row[idxApproved]);
      var valAlloc = parseNum(row[idxAllocated]);
      var valSpent = parseNum(row[idxSpent]);
      target.allocated += valAlloc;
      target.spent += valSpent;
      target.balance += parseNum(row[idxBalance]);

      var dept = String(row[idxDept] || 'ไม่ระบุ').trim();
      if (dept === '') dept = 'ไม่ระบุ';
      if (!target.deptStats[dept]) target.deptStats[dept] = { allocated: 0, spent: 0 };
      target.deptStats[dept].allocated += valAlloc;
      target.deptStats[dept].spent += valSpent;
    });
    return summary;
  } catch (e) { return { error: e.message }; }
}

// 3. SEARCH & YEARLY
function searchActionPlan(deptName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var getIdx = function(name) { return headers.findIndex(h => String(h).trim() === name); };

    var idxOrder = getIdx('ลำดับโครงการ'); 
    var idxDept = getIdx('กลุ่มงาน/งาน'); 
    var idxProject = getIdx('โครงการ');
    var idxActivity = getIdx('กิจกรรมหลัก'); 
    var idxSub = getIdx('กิจกรรมย่อย'); 
    var idxType = getIdx('ประเภทงบ');
    var idxSource = getIdx('แหล่งงบประมาณ'); 
    var idxApproved = getIdx('อนุมัติตามแผน'); 
    var idxAllocated = getIdx('จัดสรร');
    
    var monthIndices = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'].map(m => getIdx(m));
    var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
    var results = []; var summary = { count: 0, approved: 0, allocated: 0 };

    data.forEach(function(row) {
      var rowDept = (idxDept > -1) ? String(row[idxDept]).trim() : "";
      if (deptName === "" || rowDept === deptName) {
        var actName = (idxActivity > -1) ? row[idxActivity] : "";
        if (idxSub > -1 && row[idxSub]) { actName += " (" + row[idxSub] + ")"; }
        
        var approved = (idxApproved > -1) ? parseNum(row[idxApproved]) : 0;
        var allocated = (idxAllocated > -1) ? parseNum(row[idxAllocated]) : 0;
        
        summary.count++; summary.approved += approved; summary.allocated += allocated;
        var timeline = monthIndices.map(idx => (idx > -1 && String(row[idx]).trim() !== '') ? 1 : 0);
        results.push({ 
            order: (idxOrder > -1) ? row[idxOrder] : "-", dept: rowDept, project: (idxProject > -1) ? row[idxProject] : "-", 
            activity: actName, budgetType: (idxType > -1) ? row[idxType] : "-", budgetSource: (idxSource > -1) ? row[idxSource] : "-", 
            timeline: timeline, approved: approved, allocated: allocated 
        });
      }
    });
    return { summary: summary, list: results };
  } catch(e) { return { error: e.message, summary: {count:0}, list: [] }; }
}

function getYearlyPlanData(deptFilter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var getIdx = (name) => headers.findIndex(h => String(h).trim() === name);
    
    var idxOrder = getIdx('ลำดับโครงการ'); var idxDept = getIdx('กลุ่มงาน/งาน'); var idxProject = getIdx('โครงการ');
    var idxActivity = getIdx('กิจกรรมหลัก'); var idxSub = getIdx('กิจกรรมย่อย'); var idxType = getIdx('ประเภทงบ');
    var idxSource = getIdx('แหล่งงบประมาณ'); var idxApproved = getIdx('อนุมัติตามแผน'); var idxAllocated = getIdx('จัดสรร');
    var idxSpent = getIdx('เบิกจ่าย');
    
    var monthIndices = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'].map(m => getIdx(m));
    var summary = { projects: 0, approved: 0, allocated: 0, spent: 0 };
    var list = [];
    var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };

    data.forEach(row => {
      var rowDept = (idxDept > -1) ? String(row[idxDept]).trim() : "";
      if (deptFilter === "" || rowDept === deptFilter) {
        var actName = (idxActivity > -1) ? row[idxActivity] : "";
        if (idxSub > -1 && row[idxSub]) { actName += " (" + row[idxSub] + ")"; }
        
        var approved = (idxApproved > -1) ? parseNum(row[idxApproved]) : 0;
        var alloc = (idxAllocated > -1) ? parseNum(row[idxAllocated]) : 0;
        var spent = (idxSpent > -1) ? parseNum(row[idxSpent]) : 0;
        
        summary.projects++; summary.approved += approved; summary.allocated += alloc; summary.spent += spent;
        var timeline = monthIndices.map(idx => (idx > -1 && String(row[idx]).trim() !== '') ? 1 : 0);
        
        list.push({ 
            order: (idxOrder > -1) ? row[idxOrder] : "-", dept: rowDept, project: (idxProject > -1) ? row[idxProject] : "-", 
            activity: actName, type: (idxType > -1) ? row[idxType] : "-", budgetSource: (idxSource > -1) ? row[idxSource] : "-", 
            timeline: timeline, allocated: alloc, spent: spent, balance: alloc - spent 
        });
      }
    });
    return { summary: summary, list: list };
  } catch (e) { return { error: e.message }; }
}

// 4. TRANSACTION
function saveTransaction(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('m_actionplan');
    var tSheet = ss.getSheetByName('t_actionplan'); 
    if (!tSheet) { tSheet = ss.insertSheet('t_actionplan'); }

    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData.shift();
    var getIdx = (name) => mHeaders.findIndex(h => String(h).trim() === name);
    var idxID = getIdx('รหัสโครงการ'); var idxSpent = getIdx('เบิกจ่าย'); 
    
    var rowIndex = -1;
    for (var i = 0; i < mData.length; i++) {
      if (String(mData[i][idxID]) === String(form.projectId)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ในฐานข้อมูลหลัก' };

    var rowData = mData[rowIndex];
    var amount = parseFloat(form.amount);
    var currentSpent = parseFloat(String(rowData[idxSpent]).replace(/,/g,'')) || 0;
    mSheet.getRange(rowIndex + 2, idxSpent + 1).setValue(currentSpent + amount);

    tSheet.appendRow([ new Date(), rowData[idxID], rowData[getIdx('ปีงบประมาณ')], rowData[getIdx('หมวด')], rowData[getIdx('ลำดับโครงการ')], rowData[getIdx('กลุ่มงาน/งาน')], rowData[getIdx('แผนงาน')], rowData[getIdx('โครงการ')], rowData[getIdx('กิจกรรมหลัก')], rowData[getIdx('กิจกรรมย่อย')], rowData[getIdx('ประเภทงบ')], rowData[getIdx('แหล่งงบประมาณ')], rowData[getIdx('รหัสงบประมาณ')], rowData[getIdx('รหัสกิจกรรม')], rowData[getIdx('จัดสรร')], amount, 0, form.txDate, form.expenseType, form.desc, "" ]);
    return { status: 'success', message: 'บันทึกการเบิกจ่ายเรียบร้อยแล้ว' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function getTransactionHistory() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_actionplan');
    if (!tSheet) return [];
    var data = tSheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var result = [];
    for (var i = data.length - 1; i >= 1; i--) { 
      var row = data[i];
      if (row[1]) { 
         var d = row[17]; var dateStr = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy") : String(d);
         result.push({ project: row[7], activity: row[8], subActivity: row[9], date: dateStr, type: row[18], desc: row[19], amount: row[15] });
      }
      if (result.length >= 10) break;
    }
    return result;
  } catch(e) { return []; }
}

// 5. LOAN & SEARCH
function saveLoan(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('m_actionplan');
    var tSheet = ss.getSheetByName('t_loan'); 
    if (!tSheet) { tSheet = ss.insertSheet('t_loan'); }

    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData.shift();
    var getIdx = (name) => mHeaders.findIndex(h => String(h).trim() === name);
    var idxID = getIdx('รหัสโครงการ'); var idxLoan = getIdx('เงินยืม'); 
    
    var rowIndex = -1;
    for (var i = 0; i < mData.length; i++) {
      if (String(mData[i][idxID]) === String(form.projectId)) { rowIndex = i; break; }
    }
    if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ในฐานข้อมูลหลัก' };

    var rowData = mData[rowIndex];
    var amount = parseFloat(form.amount);
    
    if (idxLoan > -1) {
        var currentLoan = parseFloat(String(rowData[idxLoan]).replace(/,/g,'')) || 0;
        mSheet.getRange(rowIndex + 2, idxLoan + 1).setValue(currentLoan + amount);
    }

    tSheet.appendRow([ 
      new Date(), rowData[idxID], rowData[getIdx('ปีงบประมาณ')], rowData[getIdx('หมวด')], rowData[getIdx('ลำดับโครงการ')], rowData[getIdx('กลุ่มงาน/งาน')], rowData[getIdx('แผนงาน')], rowData[getIdx('โครงการ')], rowData[getIdx('กิจกรรมหลัก')], rowData[getIdx('กิจกรรมย่อย')], rowData[getIdx('ประเภทงบ')], rowData[getIdx('แหล่งงบประมาณ')], rowData[getIdx('รหัสงบประมาณ')], rowData[getIdx('รหัสกิจกรรม')], rowData[getIdx('จัดสรร')], 
      amount, form.loanDate, form.desc, "", form.loanType,
      "ยังไม่ดำเนินการ", 0, amount, "", "" 
    ]);
    return { status: 'success', message: 'บันทึกการยืมเงินเรียบร้อยแล้ว' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function updateLoanRepayment(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    var data = tSheet.getDataRange().getValues();
    var headers = data[0];
    
    var idxTime = headers.indexOf('ประทับเวลา');
    var idxLoanAmount = headers.indexOf('เงินยืม');
    var idxLoanDate = headers.indexOf('วันที่ยืมเงิน');
    var idxStatus = headers.indexOf('สถานะการเบิกจ่าย');
    var idxPaid = headers.indexOf('จำนวนเบิกจ่าย');
    var idxBal = headers.indexOf('คงเหลือ');
    var idxPayDate = headers.indexOf('วันที่เบิกจ่าย');
    var idxDuration = headers.indexOf('ระยะเวลาที่ยืม');

    if(idxStatus == -1) return {status:'error', message: 'ไม่พบคอลัมน์สถานะ'};

    var targetRow = -1;
    var targetTimestamp = new Date(form.timestamp).getTime();

    for(var i=1; i<data.length; i++) {
       var rowTime = new Date(data[i][idxTime]).getTime();
       if (Math.abs(rowTime - targetTimestamp) < 1000) { targetRow = i + 1; break; }
    }

    if (targetRow == -1) return {status:'error', message: 'ไม่พบรายการที่ต้องการอัปเดต'};

    var loanAmount = parseFloat(data[targetRow-1][idxLoanAmount]) || 0;
    var paidAmount = parseFloat(form.paidAmount) || 0;
    var balance = loanAmount - paidAmount;
    var loanDate = new Date(data[targetRow-1][idxLoanDate]);
    var payDate = new Date(form.payDate);
    var diffTime = Math.abs(payDate - loanDate);
    var duration = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
    var status = (balance <= 0) ? "คืนครบ" : "คืนบางส่วน";

    tSheet.getRange(targetRow, idxStatus + 1).setValue(status);
    tSheet.getRange(targetRow, idxPaid + 1).setValue(paidAmount);
    tSheet.getRange(targetRow, idxBal + 1).setValue(balance);
    tSheet.getRange(targetRow, idxPayDate + 1).setValue(form.payDate);
    tSheet.getRange(targetRow, idxDuration + 1).setValue(duration);

    return { status: 'success', message: 'บันทึกเรียบร้อย' };

  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function getLoanHistory() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];
    var data = tSheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data.shift(); 
    var getIdx = (name) => headers.indexOf(name);

    var idxOrder = getIdx('ลำดับโครงการ');
    var idxProj = getIdx('โครงการ'); var idxAct = getIdx('กิจกรรมหลัก'); var idxSub = getIdx('กิจกรรมย่อย');
    var idxDate = getIdx('วันที่ยืมเงิน'); var idxType = getIdx('ประเภทเงินยืม'); var idxDetail = getIdx('รายละเอียดการยืมเงิน'); 
    var idxAmt = getIdx('เงินยืม'); var idxTime = getIdx('ประทับเวลา');
    var idxStatus = getIdx('สถานะการเบิกจ่าย'); var idxBal = getIdx('คงเหลือ');
    var idxPaid = getIdx('จำนวนเบิกจ่าย'); var idxPayDate = getIdx('วันที่เบิกจ่าย');

    if (idxAmt === -1) return [];

    var result = [];
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (row[idxAmt]) {
         var d = row[idxDate]; var dateStr = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy") : String(d);
         var isoTime = (row[idxTime] instanceof Date) ? row[idxTime].toISOString() : "";
         var payDateStr = (idxPayDate > -1 && row[idxPayDate]) ? Utilities.formatDate(new Date(row[idxPayDate]), "Asia/Bangkok", "dd/MM/yyyy") : '';

         result.push({
           timestamp: isoTime,
           order: (idxOrder > -1) ? row[idxOrder] : '-', 
           project: (idxProj > -1) ? row[idxProj] : '-',
           activity: (idxAct > -1) ? row[idxAct] : '-',
           subActivity: (idxSub > -1) ? row[idxSub] : '',
           date: dateStr,
           type: (idxType > -1) ? row[idxType] : '-',
           details: (idxDetail > -1) ? row[idxDetail] : '-',
           amount: row[idxAmt],
           status: (idxStatus > -1) ? row[idxStatus] : 'ยังไม่ดำเนินการ',
           balance: (idxBal > -1 && row[idxBal] !== "") ? row[idxBal] : row[idxAmt],
           paid: (idxPaid > -1) ? (parseFloat(String(row[idxPaid]).replace(/,/g,'')) || 0) : 0, 
           payDate: payDateStr 
         });
      }
      if (result.length >= 10) break;
    }
    return result;
  } catch(e) { return []; }
}

function searchLoanHistory(criteria) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];
    var data = tSheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data.shift(); 
    var getIdx = (name) => headers.indexOf(name);

    var idxOrder = getIdx('ลำดับโครงการ');
    var idxProj = getIdx('โครงการ'); var idxAct = getIdx('กิจกรรมหลัก'); var idxSub = getIdx('กิจกรรมย่อย');
    var idxDate = getIdx('วันที่ยืมเงิน'); var idxType = getIdx('ประเภทเงินยืม'); var idxDetail = getIdx('รายละเอียดการยืมเงิน'); 
    var idxAmt = getIdx('เงินยืม'); var idxTime = getIdx('ประทับเวลา'); var idxStatus = getIdx('สถานะการเบิกจ่าย'); var idxBal = getIdx('คงเหลือ');
    var idxPaid = getIdx('จำนวนเบิกจ่าย'); var idxPayDate = getIdx('วันที่เบิกจ่าย');

    var result = [];
    data.forEach(function(row) {
        var matchOrder = true; if (criteria.order && String(row[idxOrder]) !== String(criteria.order)) matchOrder = false;
        var matchAct = true; if (criteria.activity && String(row[idxAct]) !== String(criteria.activity)) matchAct = false;
        var matchSub = true; if (criteria.subActivity && String(row[idxSub]).toLowerCase().indexOf(String(criteria.subActivity).toLowerCase()) === -1) matchSub = false;

        if (matchOrder && matchAct && matchSub && row[idxAmt]) {
             var d = row[idxDate]; var dateStr = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy") : String(d);
             var isoTime = (row[idxTime] instanceof Date) ? row[idxTime].toISOString() : "";
             var payDateStr = (idxPayDate > -1 && row[idxPayDate]) ? Utilities.formatDate(new Date(row[idxPayDate]), "Asia/Bangkok", "dd/MM/yyyy") : '';

             result.push({
               timestamp: isoTime,
               order: (idxOrder > -1) ? row[idxOrder] : '-',
               project: (idxProj > -1) ? row[idxProj] : '-',
               activity: (idxAct > -1) ? row[idxAct] : '-',
               subActivity: (idxSub > -1) ? row[idxSub] : '',
               date: dateStr, type: (idxType > -1) ? row[idxType] : '-', details: (idxDetail > -1) ? row[idxDetail] : '-', amount: row[idxAmt],
               status: (idxStatus > -1) ? row[idxStatus] : 'ยังไม่ดำเนินการ', 
               balance: (idxBal > -1 && row[idxBal] !== "") ? row[idxBal] : row[idxAmt],
               paid: (idxPaid > -1) ? (parseFloat(String(row[idxPaid]).replace(/,/g,'')) || 0) : 0,
               payDate: payDateStr
             });
        }
    });
    return result.sort((a,b) => (a.timestamp < b.timestamp) ? 1 : -1);
  } catch(e) { return []; }
}

function getLoanSummaryByProject() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];
    var data = tSheet.getDataRange().getValues();
    var headers = data.shift();
    var idxProj = headers.indexOf('โครงการ'); var idxAmt = headers.indexOf('เงินยืม'); var idxBal = headers.indexOf('คงเหลือ');
    if(idxProj === -1 || idxAmt === -1) return [];

    var summary = {};
    data.forEach(r => {
       var proj = String(r[idxProj]).trim();
       var amt = parseFloat(String(r[idxAmt]).replace(/,/g, '')) || 0;
       var bal = (idxBal > -1 && r[idxBal] !== "") ? (parseFloat(String(r[idxBal]).replace(/,/g, '')) || 0) : amt;
       if(proj && bal > 0) { if(!summary[proj]) summary[proj] = 0; summary[proj] += bal; }
    });
    var result = []; for(var p in summary) { result.push({ project: p, total: summary[p] }); }
    return result.sort((a, b) => b.total - a.total);
  } catch(e) { return []; }
}
