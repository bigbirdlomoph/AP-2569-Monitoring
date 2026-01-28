var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ'; 
var SHEET_NAME = 'm_actionplan';
var APP_VERSION = '690128-1456'; 

function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  template.appVersion = getVersion(); 
  return template.evaluate()
      .setTitle('AP 2569 MONITORING') 
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
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data.shift();
    var getIdx = (name) => headers.findIndex(h => String(h).trim() === name);
    
    var iID = getIdx('รหัสโครงการ'); var iOrder = getIdx('ลำดับโครงการ'); var iDept = getIdx('กลุ่มงาน/งาน');
    var iProj = getIdx('โครงการ'); var iAct = getIdx('กิจกรรมหลัก'); var iSub = getIdx('กิจกรรมย่อย');
    var iType = getIdx('ประเภทงบ'); var iSource = getIdx('แหล่งงบประมาณ'); var iAlloc = getIdx('จัดสรร');
    var iBal = getIdx('คงเหลือ (ไม่รวมเงินยืม)'); if(iBal == -1) iBal = getIdx('คงเหลือ');
    var iLoan = getIdx('เงินยืม');

    return data.map(function(r) {
      return {
        id: (iID > -1) ? r[iID] : "",
        order: (iOrder > -1) ? r[iOrder] : "",
        dept: (iDept > -1) ? r[iDept] : "",
        project: (iProj > -1) ? r[iProj] : "",
        activity: (iAct > -1) ? r[iAct] : "",
        subActivity: (iSub > -1) ? r[iSub] : "",
        budgetType: (iType > -1) ? r[iType] : "",
        budgetSource: (iSource > -1) ? r[iSource] : "",
        allocated: (iAlloc > -1) ? (parseFloat(String(r[iAlloc]).replace(/,/g,'')) || 0) : 0,
        balance: (iBal > -1) ? (parseFloat(String(r[iBal]).replace(/,/g,'')) || 0) : 0,
        loan: (iLoan > -1) ? (parseFloat(String(r[iLoan]).replace(/,/g,'')) || 0) : 0
      };
    }).filter(function(r) { return r.id && r.project; }); 
  } catch (e) { return []; }
}

// 2. DASHBOARD
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: "ไม่พบชีตข้อมูล" };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { error: "ไม่มีข้อมูล" };

    var headers = data.shift(); 
    var getIdx = function(name) { return headers.findIndex(h => String(h).trim() === name); };
    
    var idxType = getIdx('ประเภทงบ'); var idxApproved = getIdx('อนุมัติตามแผน');
    var idxAllocated = getIdx('จัดสรร'); var idxSpent = getIdx('เบิกจ่าย'); 
    var idxBalance = getIdx('คงเหลือ (ไม่รวมเงินยืม)'); if (idxBalance == -1) idxBalance = getIdx('คงเหลือ');
    var idxDept = getIdx('กลุ่มงาน/งาน');

    if (idxAllocated == -1 || idxSpent == -1) return { error: "ไม่พบคอลัมน์สำคัญ" };

    var summary = { moph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} }, nonMoph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} } };
    var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };

    data.forEach(function(row) {
      var typeVal = String(row[idxType] || "").trim();
      var isMoph = (typeVal.includes('งบประมาณ') || typeVal.includes('สป.สธ') || typeVal === 'PP' || typeVal === 'OP' || typeVal.includes('งบดำเนินงาน')); 
      var target = isMoph ? summary.moph : summary.nonMoph;
      
      if(idxApproved > -1) target.approved += parseNum(row[idxApproved]);
      target.allocated += parseNum(row[idxAllocated]);
      target.spent += parseNum(row[idxSpent]);
      if(idxBalance > -1) target.balance += parseNum(row[idxBalance]);

      var dept = String(row[idxDept] || 'ไม่ระบุ').trim();
      if (dept === '') dept = 'ไม่ระบุ';
      if (!target.deptStats[dept]) target.deptStats[dept] = { allocated: 0, spent: 0 };
      target.deptStats[dept].allocated += parseNum(row[idxAllocated]);
      target.deptStats[dept].spent += parseNum(row[idxSpent]);
    });
    return summary;
  } catch (e) { return { error: e.message }; }
}

// 3. SEARCH & YEARLY (DEPT ONLY)
function searchActionPlan(deptName) { 
    var result = getYearlyPlanData(deptName);
    var searchList = result.list.map(r => ({
        order: r.order, dept: r.dept, project: r.project, activity: r.activity,
        budgetType: r.type, budgetSource: r.budgetSource, timeline: r.timeline,
        approved: 0, allocated: r.allocated 
    }));
    return { summary: {count: result.summary.projects, approved: result.summary.approved, allocated: result.summary.allocated}, list: searchList };
}

function getYearlyPlanData(deptFilter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { summary: { projects: 0 }, list: [] };
    
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
            order: (idxOrder > -1) ? row[idxOrder] : "-", dept: rowDept, 
            project: (idxProject > -1) ? row[idxProject] : "-", 
            activity: actName, type: (idxType > -1) ? row[idxType] : "-", 
            budgetSource: (idxSource > -1) ? row[idxSource] : "-", 
            timeline: timeline, allocated: alloc, spent: spent, balance: alloc - spent 
        });
      }
    });
    return { summary: summary, list: list };
  } catch (e) { return { error: e.message }; }
}

// 4. TRANSACTION & LOAN (SAVE)
function saveTransaction(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('m_actionplan');
    var tSheet = ss.getSheetByName('t_actionplan'); 
    if (!tSheet) { tSheet = ss.insertSheet('t_actionplan'); tSheet.appendRow(['Timestamp','ID','Year','Cat','Order','Dept','Plan','Project','Activity','Sub','Type','Source','Code','ActCode','Alloc','Amount','Loan','Date','ExpType','Desc','Note']); }

    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData.shift();
    var getIdx = (name) => mHeaders.findIndex(h => String(h).trim() === name);
    var idxID = getIdx('รหัสโครงการ'); var idxSpent = getIdx('เบิกจ่าย'); 
    
    var rowIndex = -1;
    for (var i = 0; i < mData.length; i++) { if (String(mData[i][idxID]) === String(form.projectId)) { rowIndex = i; break; } }
    if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการ' };

    var rowData = mData[rowIndex];
    var amount = parseFloat(form.amount);
    var currentSpent = (parseFloat(String(rowData[idxSpent]).replace(/,/g,'')) || 0) + amount;
    mSheet.getRange(rowIndex + 2, idxSpent + 1).setValue(currentSpent);

    tSheet.appendRow([ new Date(), rowData[idxID], rowData[getIdx('ปีงบประมาณ')], rowData[getIdx('หมวด')], rowData[getIdx('ลำดับโครงการ')], rowData[getIdx('กลุ่มงาน/งาน')], rowData[getIdx('แผนงาน')], rowData[getIdx('โครงการ')], rowData[getIdx('กิจกรรมหลัก')], rowData[getIdx('กิจกรรมย่อย')], rowData[getIdx('ประเภทงบ')], rowData[getIdx('แหล่งงบประมาณ')], rowData[getIdx('รหัสงบประมาณ')], rowData[getIdx('รหัสกิจกรรม')], rowData[getIdx('จัดสรร')], amount, 0, form.txDate, form.expenseType, form.desc, "" ]);
    return { status: 'success', message: 'บันทึกเรียบร้อย' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function saveLoan(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('m_actionplan');
    var tSheet = ss.getSheetByName('t_loan'); 
    if (!tSheet) { tSheet = ss.insertSheet('t_loan'); tSheet.appendRow(['Timestamp','ID','Year','Cat','Order','Dept','Plan','Project','Activity','Sub','Type','Source','Code','ActCode','Alloc','Loan','Date','Desc','Note','LoanType','Status','Paid','Bal','PayDate','Duration']);}

    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData.shift();
    var getIdx = (name) => mHeaders.findIndex(h => String(h).trim() === name);
    var idxID = getIdx('รหัสโครงการ'); var idxLoan = getIdx('เงินยืม'); 
    
    var rowIndex = -1;
    for (var i = 0; i < mData.length; i++) { if (String(mData[i][idxID]) === String(form.projectId)) { rowIndex = i; break; } }
    if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการ' };

    var rowData = mData[rowIndex];
    var amount = parseFloat(form.amount);
    
    if (idxLoan > -1) {
        var currentLoan = (parseFloat(String(rowData[idxLoan]).replace(/,/g,'')) || 0) + amount;
        mSheet.getRange(rowIndex + 2, idxLoan + 1).setValue(currentLoan);
    }

    tSheet.appendRow([ new Date(), rowData[idxID], rowData[getIdx('ปีงบประมาณ')], rowData[getIdx('หมวด')], rowData[getIdx('ลำดับโครงการ')], rowData[getIdx('กลุ่มงาน/งาน')], rowData[getIdx('แผนงาน')], rowData[getIdx('โครงการ')], rowData[getIdx('กิจกรรมหลัก')], rowData[getIdx('กิจกรรมย่อย')], rowData[getIdx('ประเภทงบ')], rowData[getIdx('แหล่งงบประมาณ')], rowData[getIdx('รหัสงบประมาณ')], rowData[getIdx('รหัสกิจกรรม')], rowData[getIdx('จัดสรร')], amount, form.loanDate, form.desc, "", form.loanType, "ยังไม่ดำเนินการ", 0, amount, "", "" ]);
    return { status: 'success', message: 'บันทึกเงินยืมเรียบร้อย' };
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
    
    var idxTime = headers.indexOf('Timestamp'); if(idxTime == -1) idxTime = 0; 
    var idxLoanAmt = 15; var idxStatus = 20; var idxPaid = 21; var idxBal = 22; var idxPayDate = 23;

    var targetRow = -1;
    var targetTimestamp = new Date(form.timestamp).getTime();

    for(var i=1; i<data.length; i++) {
       var rowTime = new Date(data[i][idxTime]).getTime();
       if (Math.abs(rowTime - targetTimestamp) < 1000) { targetRow = i + 1; break; }
    }
    if (targetRow == -1) return {status:'error', message: 'ไม่พบรายการ'};

    var loanAmount = parseFloat(data[targetRow-1][idxLoanAmt]) || 0;
    var paidAmount = parseFloat(form.paidAmount) || 0;
    var balance = loanAmount - paidAmount;
    var status = (balance <= 0) ? "คืนครบ" : "คืนบางส่วน";

    tSheet.getRange(targetRow, idxStatus + 1).setValue(status);
    tSheet.getRange(targetRow, idxPaid + 1).setValue(paidAmount);
    tSheet.getRange(targetRow, idxBal + 1).setValue(balance);
    tSheet.getRange(targetRow, idxPayDate + 1).setValue(form.payDate);

    return { status: 'success', message: 'บันทึกคืนเงินเรียบร้อย' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

// 6. HISTORY & SEARCH
function getTransactionHistory() { return getHistory('t_actionplan', [4,7,8,9,15,17,18]); } 
function getLoanHistory() { return getHistory('t_loan', [4,7,8,9,15,16,19,20,21,22,23]); }

function searchTransactionHistory(criteria) { return searchHistory('t_actionplan', criteria, [4,7,8,9,15,17,18]); }
function searchLoanHistory(criteria) { return searchHistory('t_loan', criteria, [4,7,8,9,15,16,19,20,21,22,23]); }

function getHistory(sheetName, indices) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName(sheetName);
    if (!tSheet) return [];
    var data = tSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var result = [];
    for (var i = data.length - 1; i >= 1; i--) { 
      var row = data[i];
      if (row[indices[4]]) { 
         var d = row[indices[5]]; var dateStr = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy") : String(d);
         var item = { 
             order: row[indices[0]], project: row[indices[1]], activity: row[indices[2]], subActivity: row[indices[3]], 
             amount: row[indices[4]], date: dateStr, type: row[indices[6]]
         };
         if(sheetName === 't_loan') {
             item.timestamp = (row[0] instanceof Date) ? row[0].toISOString() : "";
             item.status = row[indices[7]] || 'ยังไม่ดำเนินการ';
             item.paid = row[indices[8]] || 0;
             item.balance = (row[indices[9]] !== "") ? row[indices[9]] : item.amount;
             item.payDate = (row[indices[10]] && row[indices[10]] instanceof Date) ? Utilities.formatDate(row[indices[10]], "Asia/Bangkok", "dd/MM/yyyy") : '';
             item.details = row[17]; 
         }
         result.push(item);
      }
      if (result.length >= 10) break;
    }
    return result;
  } catch(e) { return []; }
}

function searchHistory(sheetName, criteria, indices) {
    try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName(sheetName);
    if (!tSheet) return [];
    var data = tSheet.getDataRange().getValues();
    
    var result = [];
    for(var i=1; i<data.length; i++) {
        var row = data[i];
        
        var matchOrder = true; if(criteria.order && String(row[indices[0]]) !== String(criteria.order)) matchOrder = false;
        var matchProj = true; if(criteria.project && String(row[indices[1]]) !== String(criteria.project)) matchProj = false;
        var matchAct = true; if(criteria.activity && String(row[indices[2]]) !== String(criteria.activity)) matchAct = false;
        var matchSub = true; if(criteria.subActivity && String(row[indices[3]]).toLowerCase().indexOf(String(criteria.subActivity).toLowerCase()) === -1) matchSub = false;

        if (matchOrder && matchProj && matchAct && matchSub) {
             var d = row[indices[5]]; var dateStr = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy") : String(d);
             var item = { 
                 order: row[indices[0]], project: row[indices[1]], activity: row[indices[2]], subActivity: row[indices[3]], 
                 amount: row[indices[4]], date: dateStr, type: row[indices[6]]
             };
             if(sheetName === 't_loan') {
                 item.timestamp = (row[0] instanceof Date) ? row[0].toISOString() : "";
                 item.status = row[indices[7]] || 'ยังไม่ดำเนินการ';
                 item.paid = row[indices[8]] || 0;
                 item.balance = (row[indices[9]] !== "") ? row[indices[9]] : item.amount;
                 item.payDate = (row[indices[10]] && row[indices[10]] instanceof Date) ? Utilities.formatDate(row[indices[10]], "Asia/Bangkok", "dd/MM/yyyy") : '';
                 item.details = row[17];
             }
             result.push(item);
        }
    }
    return result.reverse(); 
  } catch(e) { return []; }
}
