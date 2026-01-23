var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ'; 
var SHEET_NAME = 'm_actionplan';

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

function getVersion() {
  var now = new Date();
  var timeZone = 'Asia/Bangkok';
  var year = parseInt(Utilities.formatDate(now, timeZone, 'yyyy')) + 543;
  var dateStr = Utilities.formatDate(now, timeZone, 'MMdd-HHmm');
  return String(year).slice(-2) + dateStr; 
}

// 1. DATA LOADER
function getAllMasterDataForClient() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    
    var getIdx = (name) => headers.findIndex(h => String(h).trim() === name);
    
    var iID = getIdx('รหัสโครงการ');
    var iOrder = getIdx('ลำดับโครงการ');
    var iDept = getIdx('กลุ่มงาน/งาน');
    var iProj = getIdx('โครงการ');
    var iAct = getIdx('กิจกรรมหลัก');
    var iSub = getIdx('กิจกรรมย่อย');
    var iType = getIdx('ประเภทงบ');
    var iSource = getIdx('แหล่งงบประมาณ');
    var iAlloc = getIdx('จัดสรร');
    var iBal = getIdx('คงเหลือ (ไม่รวมเงินยืม)');
    var iLoan = getIdx('เงินยืม');

    return data.map(r => ({
      id: r[iID],
      order: r[iOrder],
      dept: r[iDept],
      project: r[iProj],
      activity: r[iAct],
      subActivity: r[iSub],
      budgetType: r[iType],
      budgetSource: r[iSource],
      allocated: r[iAlloc],
      balance: r[iBal],
      loan: (iLoan > -1) ? r[iLoan] : 0
    })).filter(r => r.id && r.project); 
  } catch (e) {
    Logger.log(e.message);
    return [];
  }
}

// 2. DASHBOARD LOGIC (Fix หน้าขาว)
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: "ไม่พบชีตชื่อ " + SHEET_NAME };

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return { error: "ไม่มีข้อมูลในตาราง" }; // Check empty data

    var headers = data.shift(); 
    var getIdx = function(name) { return headers.findIndex(h => String(h).trim() === name); };
    
    var idxType = getIdx('ประเภทงบ'); 
    var idxApproved = getIdx('อนุมัติตามแผน');
    var idxAllocated = getIdx('จัดสรร');
    var idxSpent = getIdx('เบิกจ่าย'); 
    var idxBalance = getIdx('คงเหลือ (ไม่รวมเงินยืม)');
    var idxDept = getIdx('กลุ่มงาน/งาน');

    if (idxType == -1 || idxApproved == -1 || idxAllocated == -1 || idxSpent == -1 || idxBalance == -1) {
      return { error: "ไม่พบหัวตารางสำคัญสำหรับ Dashboard" };
    }

    var summary = { moph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} }, nonMoph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} } };
    
    // Robust Parse Float
    var parseNum = function(val) {
        if (!val) return 0;
        var num = parseFloat(String(val).replace(/,/g, ''));
        return isNaN(num) ? 0 : num;
    };

    data.forEach(function(row) {
      var typeVal = String(row[idxType] || "").trim();
      var isMoph = false;
      if (typeVal.includes('งบประมาณ') || typeVal.includes('สป.สธ') || typeVal === 'PP' || typeVal === 'OP' || typeVal.includes('งบดำเนินงาน')) { isMoph = true; } 

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
  } catch (e) { return { error: "Dashboard Error: " + e.message }; }
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
    
    var months = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];
    var monthIndices = months.map(m => getIdx(m));

    var parseNum = (val) => parseFloat(String(val || 0).replace(/,/g, '')) || 0;
    var results = [];
    var summary = { count: 0, approved: 0, allocated: 0 };

    data.forEach(function(row) {
      if (deptName === "" || String(row[idxDept]).trim() === deptName) {
        var approved = (idxApproved !== -1) ? parseNum(row[idxApproved]) : 0;
        var allocated = (idxAllocated !== -1) ? parseNum(row[idxAllocated]) : 0;

        summary.count++;
        summary.approved += approved;
        summary.allocated += allocated;

        var timeline = monthIndices.map(idx => (idx > -1 && String(row[idx]).trim() !== '') ? 1 : 0);
        
        var actName = row[idxActivity];
        if (idxSub > -1 && row[idxSub]) { actName += " (" + row[idxSub] + ")"; }

        results.push({
          order: (idxOrder > -1) ? row[idxOrder] : "-",
          dept: row[idxDept],
          project: row[idxProject],
          activity: actName,
          budgetType: (idxType > -1) ? row[idxType] : "-",
          budgetSource: (idxSource > -1) ? row[idxSource] : "-",
          timeline: timeline, 
          approved: approved,
          allocated: allocated
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
    
    var idxOrder = getIdx('ลำดับโครงการ');
    var idxDept = getIdx('กลุ่มงาน/งาน');
    var idxProject = getIdx('โครงการ');
    var idxActivity = getIdx('กิจกรรมหลัก');
    var idxSub = getIdx('กิจกรรมย่อย');
    var idxType = getIdx('ประเภทงบ');
    var idxSource = getIdx('แหล่งงบประมาณ');
    var idxAllocated = getIdx('จัดสรร');
    var idxSpent = getIdx('เบิกจ่าย');

    var monthIndices = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'].map(m => getIdx(m));
    var summary = { projects: 0, allocated: 0, spent: 0 };
    var list = [];
    var parseNum = (v) => parseFloat(String(v || 0).replace(/,/g, '')) || 0;

    data.forEach(row => {
      if (deptFilter === "" || String(row[idxDept]).trim() === deptFilter) {
        var alloc = parseNum(row[idxAllocated]);
        var spent = parseNum(row[idxSpent]);

        summary.projects++;
        summary.allocated += alloc;
        summary.spent += spent;

        var timeline = monthIndices.map(idx => (idx > -1 && String(row[idx]).trim() !== '') ? 1 : 0);
        
        var actName = row[idxActivity];
        if (idxSub > -1 && row[idxSub]) { actName += " (" + row[idxSub] + ")"; }

        list.push({
          order: (idxOrder > -1) ? row[idxOrder] : "-",
          dept: row[idxDept],
          project: row[idxProject],
          activity: actName,
          type: row[idxType],
          budgetSource: row[idxSource],
          timeline: timeline, 
          allocated: alloc,
          spent: spent,
          balance: alloc - spent
        });
      }
    });
    return { summary: summary, list: list };
  } catch (e) { return { error: e.message }; }
}

// 4. TRANSACTION & LOAN
function saveTransaction(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('m_actionplan');
    var tSheet = ss.getSheetByName('t_actionplan'); 
    
    if (!tSheet) {
      tSheet = ss.insertSheet('t_actionplan');
      tSheet.appendRow(['ประทับเวลา', 'รหัสโครงการ', 'ปีงบประมาณ', 'หมวด', 'ลำดับโครงการ', 'กลุ่มงาน/งาน', 'แผนงาน', 'โครงการ', 'กิจกรรมหลัก', 'กิจกรรมย่อย', 'ประเภทงบ', 'แหล่งงบประมาณ', 'รหัสงบประมาณ', 'รหัสกิจกรรม', 'จัดสรร', 'เบิกจ่ายครั้งนี้', 'เงินยืม', 'วันที่เบิกจ่าย', 'ประเภทค่าใช้จ่าย', 'รายละเอียดการเบิกจ่าย', 'หมายเหตุ']);
    }

    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData.shift();
    var getIdx = (name) => mHeaders.findIndex(h => String(h).trim() === name);
    
    var idxID = getIdx('รหัสโครงการ');
    var idxSpent = getIdx('เบิกจ่าย'); 

    var rowIndex = -1;
    for (var i = 0; i < mData.length; i++) {
      if (String(mData[i][idxID]) === String(form.projectId)) { 
        rowIndex = i; break;
      }
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

function saveLoan(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('m_actionplan');
    var tSheet = ss.getSheetByName('t_loan'); 
    
    if (!tSheet) {
      tSheet = ss.insertSheet('t_loan');
      tSheet.appendRow(['ประทับเวลา', 'รหัสโครงการ', 'ปีงบประมาณ', 'หมวด', 'ลำดับโครงการ', 'กลุ่มงาน/งาน', 'แผนงาน', 'โครงการ', 'กิจกรรมหลัก', 'กิจกรรมย่อย', 'ประเภทงบ', 'แหล่งงบประมาณ', 'รหัสงบประมาณ', 'รหัสกิจกรรม', 'จัดสรร', 'เงินยืม', 'วันที่ยืมเงิน', 'รายละเอียดการยืมเงิน', 'หมายเหตุ', 'ประเภทเงินยืม']);
    }

    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData.shift();
    var getIdx = (name) => mHeaders.findIndex(h => String(h).trim() === name);
    
    var idxID = getIdx('รหัสโครงการ');
    var idxLoan = getIdx('เงินยืม'); 

    var rowIndex = -1;
    for (var i = 0; i < mData.length; i++) {
      if (String(mData[i][idxID]) === String(form.projectId)) { 
        rowIndex = i; break;
      }
    }
    if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ในฐานข้อมูลหลัก' };

    var rowData = mData[rowIndex];
    var amount = parseFloat(form.amount);
    
    if (idxLoan > -1) {
        var currentLoan = parseFloat(String(rowData[idxLoan]).replace(/,/g,'')) || 0;
        mSheet.getRange(rowIndex + 2, idxLoan + 1).setValue(currentLoan + amount);
    }

    tSheet.appendRow([ 
      new Date(), 
      rowData[idxID], 
      rowData[getIdx('ปีงบประมาณ')], 
      rowData[getIdx('หมวด')], 
      rowData[getIdx('ลำดับโครงการ')], 
      rowData[getIdx('กลุ่มงาน/งาน')], 
      rowData[getIdx('แผนงาน')], 
      rowData[getIdx('โครงการ')], 
      rowData[getIdx('กิจกรรมหลัก')], 
      rowData[getIdx('กิจกรรมย่อย')], 
      rowData[getIdx('ประเภทงบ')], 
      rowData[getIdx('แหล่งงบประมาณ')], 
      rowData[getIdx('รหัสงบประมาณ')], 
      rowData[getIdx('รหัสกิจกรรม')], 
      rowData[getIdx('จัดสรร')], 
      amount, 
      form.loanDate, 
      form.desc, 
      "", 
      form.loanType 
    ]);
    return { status: 'success', message: 'บันทึกการยืมเงินเรียบร้อยแล้ว' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function getTransactionHistory() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_actionplan');
    if (!tSheet) return [];
    
    var data = tSheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    // Explicit Mapping for t_actionplan
    // 7:Project, 8:Activity, 9:SubActivity, 15:Amount, 17:Date, 18:Type, 19:Desc
    var result = [];
    for (var i = data.length - 1; i >= 1; i--) { // Skip header
      var row = data[i];
      if (row[1]) { // Check ID
         var d = row[17];
         var dateStr = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy") : String(d);
         
         result.push({
           project: row[7],
           activity: row[8],
           subActivity: row[9],
           date: dateStr,
           type: row[18],
           desc: row[19],
           amount: row[15]
         });
      }
      if (result.length >= 10) break;
    }
    return result;
  } catch(e) { return []; }
}

function getLoanHistory() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];
    
    var data = tSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    var headers = data.shift(); 
    var getIdx = function(name) { return headers.indexOf(name); };

    // Find Indices by Name
    var idxProj = getIdx('โครงการ');
    var idxAct = getIdx('กิจกรรมหลัก');
    var idxSub = getIdx('กิจกรรมย่อย');
    var idxDate = getIdx('วันที่ยืมเงิน');
    var idxType = getIdx('ประเภทเงินยืม');
    var idxDetail = getIdx('รายละเอียดการยืมเงิน'); 
    var idxAmt = getIdx('เงินยืม');

    if (idxAmt === -1) return [];

    var result = [];
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      if (row[idxAmt] && String(row[idxAmt]) !== "") {
         var d = row[idxDate];
         var dateStr = (d instanceof Date) ? Utilities.formatDate(d, "Asia/Bangkok", "dd/MM/yyyy") : String(d);

         result.push({
           project: (idxProj > -1) ? row[idxProj] : '-',
           activity: (idxAct > -1) ? row[idxAct] : '-',
           subActivity: (idxSub > -1) ? row[idxSub] : '',
           date: dateStr,
           type: (idxType > -1) ? row[idxType] : '-',
           details: (idxDetail > -1) ? row[idxDetail] : '-',
           amount: row[idxAmt]
         });
      }
      if (result.length >= 10) break;
    }
    return result;
  } catch(e) { return []; }
}

function getLoanSummaryByProject() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];
    
    var data = tSheet.getDataRange().getValues();
    var headers = data.shift();
    var idxProj = headers.indexOf('โครงการ');
    var idxAmt = headers.indexOf('เงินยืม');
    
    if(idxProj === -1 || idxAmt === -1) return [];

    var summary = {};
    data.forEach(r => {
       var proj = String(r[idxProj]).trim();
       var amt = parseFloat(String(r[idxAmt]).replace(/,/g, '')) || 0;
       if(proj && amt > 0) {
          if(!summary[proj]) summary[proj] = 0;
          summary[proj] += amt;
       }
    });

    var result = [];
    for(var p in summary) {
       result.push({ project: p, total: summary[p] });
    }
    return result.sort((a, b) => b.total - a.total);

  } catch(e) { return []; }
}
