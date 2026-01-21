var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ'; 
var SHEET_NAME = 'm_actionplan';

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

function getVersion() {
  var now = new Date();
  var timeZone = 'Asia/Bangkok';
  var year = parseInt(Utilities.formatDate(now, timeZone, 'yyyy')) + 543;
  var dateStr = Utilities.formatDate(now, timeZone, 'MMdd-HHmm');
  return String(year).slice(-2) + dateStr; 
}

// 1. DASHBOARD
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: "ไม่พบชีตชื่อ " + SHEET_NAME };

    var data = sheet.getDataRange().getValues();
    var headers = data.shift(); 
    var getIdx = function(name) { return headers.findIndex(h => String(h).trim() === name); };
    
    var idxType = getIdx('ประเภทงบ'); 
    var idxApproved = getIdx('อนุมัติตามแผน');
    var idxAllocated = getIdx('จัดสรร');
    var idxSpent = getIdx('เบิกจ่าย'); 
    var idxBalance = getIdx('คงเหลือ (ไม่รวมเงินยืม)');
    var idxDept = getIdx('กลุ่มงาน/งาน');

    if (idxType == -1 || idxApproved == -1 || idxAllocated == -1 || idxSpent == -1 || idxBalance == -1) {
      return { error: "ไม่พบหัวตารางสำคัญ (Dashboard)" };
    }

    var summary = { moph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} }, nonMoph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} } };
    var parseNum = (val) => parseFloat(String(val || 0).replace(/,/g, '')) || 0;

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

// 2. SEARCH
function getSearchOptions() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var idxDept = headers.findIndex(h => String(h).trim() === 'กลุ่มงาน/งาน');
    if (idxDept === -1) return [];
    return data.map(r => String(r[idxDept]).trim()).filter(v => v !== '').filter((v, i, a) => a.indexOf(v) === i).sort();
  } catch(e) { return []; }
}

function searchActionPlan(deptName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var getIdx = function(name) { return headers.findIndex(h => String(h).trim() === name); };

    // Columns
    var idxOrder = getIdx('ลำดับโครงการ');
    var idxDept = getIdx('กลุ่มงาน/งาน');
    var idxProject = getIdx('โครงการ');
    var idxActivity = getIdx('กิจกรรมหลัก');
    var idxType = getIdx('ประเภทงบ');
    var idxSource = getIdx('แหล่งงบประมาณ');
    var idxApproved = getIdx('อนุมัติตามแผน');
    var idxAllocated = getIdx('จัดสรร');
    
    // Months
    var months = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];
    var monthIndices = months.map(m => getIdx(m));

    var parseNum = (val) => parseFloat(String(val || 0).replace(/,/g, '')) || 0;

    var results = [];
    var summary = { count: 0, approved: 0, allocated: 0 };

    data.forEach(function(row) {
      var rowDept = String(row[idxDept] || "").trim();
      
      if (deptName === "" || rowDept === deptName) {
        var approved = (idxApproved !== -1) ? parseNum(row[idxApproved]) : 0;
        var allocated = (idxAllocated !== -1) ? parseNum(row[idxAllocated]) : 0;

        summary.count++;
        summary.approved += approved;
        summary.allocated += allocated;

        var timeline = monthIndices.map(idx => (idx > -1 && String(row[idx]).trim() !== '') ? 1 : 0);

        results.push({
          order: row[idxOrder],
          dept: rowDept,
          project: (idxProject !== -1) ? row[idxProject] : "-",
          activity: (idxActivity !== -1) ? row[idxActivity] : "-",
          budgetType: (idxType !== -1) ? row[idxType] : "-",
          budgetSource: (idxSource !== -1) ? row[idxSource] : "-",
          timeline: timeline, 
          approved: approved,
          allocated: allocated
        });
      }
    });
    return { summary: summary, list: results };
  } catch(e) { return { error: e.message, summary: {count:0}, list: [] }; }
}

// 3. TRANSACTION
function getTxOptions(type, parentValue) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var headers = data.shift();
  var getIdx = (name) => headers.findIndex(h => String(h).trim() === name);
  var idxDept = getIdx('กลุ่มงาน/งาน');
  var idxProject = getIdx('โครงการ');
  var idxActivity = getIdx('กิจกรรมหลัก');
  
  var uniqueList = [];
  if (type === 'dept') {
    uniqueList = [...new Set(data.map(r => r[idxDept]).filter(String))].sort();
  } else if (type === 'project') {
    uniqueList = [...new Set(data.filter(r => r[idxDept] === parentValue).map(r => r[idxProject]).filter(String))].sort();
  } else if (type === 'activity') {
    var idxAllocated = getIdx('จัดสรร');
    var idxBalance = getIdx('คงเหลือ (ไม่รวมเงินยืม)');
    uniqueList = data.filter(r => r[idxProject] === parentValue).map(r => {
       return { name: r[idxActivity], allocated: r[idxAllocated], balance: r[idxBalance] };
    });
  }
  return uniqueList;
}

function saveTransaction(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName('m_actionplan');
    var tSheet = ss.getSheetByName('t_actionplan'); 
    if (!tSheet) {
      tSheet = ss.insertSheet('t_actionplan');
      tSheet.appendRow(['Timestamp', 'วันที่เบิกจ่าย', 'กลุ่มงาน', 'โครงการ', 'กิจกรรม', 'ประเภทค่าใช้จ่าย', 'รายละเอียด', 'ยอดเบิกจ่าย']);
    }
    var mData = mSheet.getDataRange().getValues();
    var mHeaders = mData.shift();
    var getIdx = (name) => mHeaders.findIndex(h => String(h).trim() === name);
    var idxDept = getIdx('กลุ่มงาน/งาน');
    var idxProject = getIdx('โครงการ');
    var idxActivity = getIdx('กิจกรรมหลัก');
    var idxSpent = getIdx('เบิกจ่าย'); 

    var rowIndex = -1;
    for (var i = 0; i < mData.length; i++) {
      if (mData[i][idxDept] == form.dept && mData[i][idxProject] == form.project && mData[i][idxActivity] == form.activity) {
        rowIndex = i; break;
      }
    }
    if (rowIndex === -1) return { status: 'error', message: 'ไม่พบข้อมูลโครงการในฐานข้อมูลหลัก' };

    var amount = parseFloat(form.amount);
    var currentSpent = parseFloat(String(mData[rowIndex][idxSpent]).replace(/,/g,'')) || 0;
    var targetRow = rowIndex + 2; 
    mSheet.getRange(targetRow, idxSpent + 1).setValue(currentSpent + amount);
    tSheet.appendRow([ new Date(), form.txDate, form.dept, form.project, form.activity, form.expenseType, form.desc, amount ]);
    return { status: 'success', message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function getTransactionHistory() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var tSheet = ss.getSheetByName('t_actionplan');
  if (!tSheet) return [];
  var data = tSheet.getDataRange().getValues();
  data.shift(); 
  return data.reverse().slice(0, 10);
}

// 4. YEARLY PLAN (Updated: Add Source & Timeline Array)
function getYearlyPlanData(deptFilter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var data = sheet.getDataRange().getValues();
    var headers = data.shift();
    var getIdx = (name) => headers.findIndex(h => String(h).trim() === name);
    
    // Columns
    var idxOrder = getIdx('ลำดับโครงการ');
    var idxDept = getIdx('กลุ่มงาน/งาน');
    var idxProject = getIdx('โครงการ');
    var idxActivity = getIdx('กิจกรรมหลัก');
    var idxType = getIdx('ประเภทงบ');
    var idxSource = getIdx('แหล่งงบประมาณ'); // New
    var idxAllocated = getIdx('จัดสรร');
    var idxSpent = getIdx('เบิกจ่าย');

    var months = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];
    var monthIndices = months.map(m => getIdx(m));

    var summary = { projects: 0, allocated: 0, spent: 0 };
    var list = [];
    var parseNum = (v) => parseFloat(String(v || 0).replace(/,/g, '')) || 0;

    data.forEach(row => {
      var rowDept = String(row[idxDept] || "").trim();
      
      if (deptFilter === "" || deptFilter === rowDept) {
        var alloc = parseNum(row[idxAllocated]);
        var spent = parseNum(row[idxSpent]);

        summary.projects++;
        summary.allocated += alloc;
        summary.spent += spent;

        var timeline = monthIndices.map(idx => (idx > -1 && String(row[idx]).trim() !== '') ? 1 : 0);

        list.push({
          order: row[idxOrder],
          dept: rowDept,
          project: row[idxProject],
          activity: row[idxActivity],
          type: row[idxType],
          budgetSource: (idxSource !== -1) ? row[idxSource] : "-", // Send Source
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
