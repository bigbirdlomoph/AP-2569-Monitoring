var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ'; 
var SHEET_NAME = 'm_actionplan';
var APP_VERSION = '6900206-1545'; 

function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  template.appVersion = APP_VERSION; 
  return template.evaluate()
      .setTitle('AP 2569 MONITORING') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// 1. DATA LOADER (MASTER DATA)
function getAllMasterDataForClient() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    data.shift(); // Remove header

    // 🎯 HARDCODE INDEX ตาม LOG ของท่าน (ชัวร์ 100%)
    // [0]รหัส, [3]ลำดับ, [4]กลุ่มงาน, [6]โครงการ, [7]กิจกรรมหลัก, [8]ย่อย, [9]ประเภทงบ, [10]แหล่ง, [16]จัดสรร, [19]คงเหลือ, [18]เงินยืม
    return data.map(function(r) {
      return {
        id: r[0], order: r[3], dept: r[4], project: r[6], activity: r[7], subActivity: r[8],
        budgetType: r[9], budgetSource: r[10],
        allocated: parseFloat(String(r[16]).replace(/,/g,'')) || 0,
        balance: parseFloat(String(r[19]).replace(/,/g,'')) || 0, // คงเหลือ (ไม่รวมเงินยืม)
        loan: parseFloat(String(r[18]).replace(/,/g,'')) || 0
      };
    }).filter(function(r) { return r.id && r.project; }); 
  } catch (e) { return []; }
}

// 2. DASHBOARD DATA
function getDashboardData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { error: "ไม่พบ Sheet" };
    var data = sheet.getDataRange().getValues();
    data.shift(); // Remove Header

    // 🎯 HARDCODE INDEX
    var I_DEPT=4, I_TYPE=9, I_ALLOC=16, I_SPENT=17, I_BAL=19, I_APPROVE=15;

    var summary = { moph: { approved:0, allocated:0, spent:0, balance:0, deptStats:{} }, nonMoph: { approved:0, allocated:0, spent:0, balance:0, deptStats:{} } };
    var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g,'')); return isNaN(v) ? 0 : v; };

    data.forEach(function(row) {
      var typeVal = String(row[I_TYPE] || "").trim();
      var isMoph = (typeVal.includes('งบประมาณ') || typeVal.includes('สป.สธ') || typeVal === 'PP' || typeVal === 'OP' || typeVal.includes('งบดำเนินงาน')); 
      var target = isMoph ? summary.moph : summary.nonMoph;
      
      target.approved += parseNum(row[I_APPROVE]);
      target.allocated += parseNum(row[I_ALLOC]);
      target.spent += parseNum(row[I_SPENT]);
      target.balance += parseNum(row[I_BAL]);

      var dept = String(row[I_DEPT] || 'ไม่ระบุ').trim();
      if (dept === '') dept = 'ไม่ระบุ';
      if (!target.deptStats[dept]) target.deptStats[dept] = { allocated: 0, spent: 0 };
      target.deptStats[dept].allocated += parseNum(row[I_ALLOC]);
      target.deptStats[dept].spent += parseNum(row[I_SPENT]);
    });
    return summary;
  } catch (e) { return { error: e.message }; }
}

// 3. SEARCH & YEARLY
function searchActionPlan(dept, budgetType, quarter, month) { 
    var result = getYearlyPlanData(dept, budgetType, quarter, month);
    return { summary: {count: result.summary.projects, approved: result.summary.approved, allocated: result.summary.allocated}, list: result.list };
}

function getYearlyPlanData(deptFilter, typeFilter, quarterFilter, monthFilter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { summary: { projects: 0 }, list: [] };
    var data = sheet.getDataRange().getValues();
    data.shift();

    // 🎯 HARDCODE INDEX
    var I_ORDER=3, I_DEPT=4, I_PROJ=6, I_ACT=7, I_SUB=8, I_TYPE=9, I_SOURCE=10, I_ALLOC=16, I_SPENT=17;
    // Index เดือน ต.ค.(26) - ก.ย.(37)
    var I_MONTHS = [26,27,28,29,30,31,32,33,34,35,36,37];
    var quarters = { 'Q1': [0,1,2], 'Q2': [3,4,5], 'Q3': [6,7,8], 'Q4': [9,10,11] };
    
    var summary = { projects: 0, approved: 0, allocated: 0, spent: 0 };
    var list = [];
    var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g,'')); return isNaN(v) ? 0 : v; };

    data.forEach(row => {
      var rowDept = String(row[I_DEPT]).trim();
      var passDept = (deptFilter === "" || deptFilter === null || rowDept === deptFilter);

      var typeVal = String(row[I_TYPE] || "").trim();
      var isMoph = (typeVal.includes('งบประมาณ') || typeVal.includes('สป.สธ') || typeVal === 'PP' || typeVal === 'OP' || typeVal.includes('งบดำเนินงาน')); 
      var passType = true;
      if (typeFilter === 'MOPH') passType = isMoph;
      else if (typeFilter === 'NONMOPH') passType = !isMoph;

      var timeline = I_MONTHS.map(idx => (String(row[idx]).trim() !== '') ? 1 : 0);
      var passTime = true;
      if (quarterFilter && quarters[quarterFilter]) {
          if (!quarters[quarterFilter].some(mIdx => timeline[mIdx] === 1)) passTime = false;
      }
      if (monthFilter) {
          if (timeline[parseInt(monthFilter)] !== 1) passTime = false;
      }

      if (passDept && passType && passTime) {
        var actName = String(row[I_ACT]);
        if (row[I_SUB]) actName += " (" + row[I_SUB] + ")";
        
        var alloc = parseNum(row[I_ALLOC]);
        var spent = parseNum(row[I_SPENT]);
        
        summary.projects++; summary.allocated += alloc; summary.spent += spent;
        
        list.push({ 
            order: row[I_ORDER], dept: rowDept, project: row[I_PROJ], activity: actName, 
            type: row[I_TYPE], budgetSource: row[I_SOURCE], 
            timeline: timeline, allocated: alloc, spent: spent, balance: alloc - spent 
        });
      }
    });
    return { summary: summary, list: list };
  } catch (e) { return { error: e.message }; }
}

// 4. SAVE & UPDATE (Transaction)
function saveTransaction(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var mSheet = ss.getSheetByName(SHEET_NAME);
    var tSheet = ss.getSheetByName('t_actionplan'); 
    
    // หาแถวใน Master เพื่ออัปเดตยอดเบิกจ่าย
    var mData = mSheet.getDataRange().getValues();
    var idxID = 0; // รหัสโครงการอยู่ Col 0
    var idxSpent = 17; // ยอดเบิกจ่ายอยู่ Col 17 (R)
    
    var rowIndex = -1;
    for (var i = 1; i < mData.length; i++) { if (String(mData[i][idxID]) === String(form.projectId)) { rowIndex = i + 1; break; } }
    
    if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการใน Master' };

    // Update Master
    var currentSpent = (parseFloat(String(mSheet.getRange(rowIndex, idxSpent + 1).getValue()).replace(/,/g,'')) || 0) + parseFloat(form.amount);
    mSheet.getRange(rowIndex, idxSpent + 1).setValue(currentSpent);

    // Save Log
    // RowData มาจาก mData[rowIndex-1]
    var r = mData[rowIndex-1];
    tSheet.appendRow([ new Date(), r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[10], r[13], r[14], r[16], form.amount, 0, form.txDate, form.expenseType, form.desc, r[0] ]); // เพิ่ม ID ที่ column สุดท้าย
    
    return { status: 'success', message: 'บันทึกเรียบร้อย' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

function deleteTransaction(rowId, projectId, amount) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var txSheet = ss.getSheetByName('t_actionplan');
    var mSheet = ss.getSheetByName(SHEET_NAME);
    
    // 1. Sync Back (คืนเงินเข้า Master)
    var mData = mSheet.getDataRange().getValues();
    var idxID = 0; var idxSpent = 17;
    var mRow = -1;
    for(var i=1; i<mData.length; i++){ if(String(mData[i][idxID]) === String(projectId)){ mRow = i+1; break; } }
    
    if(mRow !== -1) {
       var cur = parseFloat(String(mSheet.getRange(mRow, idxSpent+1).getValue()).replace(/,/g,'')) || 0;
       mSheet.getRange(mRow, idxSpent+1).setValue(cur - amount);
    }
    
    // 2. Delete Row
    txSheet.deleteRow(rowId);
    return { status: 'success', message: 'ลบรายการและคืนเงินเรียบร้อย' };
  } catch(e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

// 5. LOAN FUNCTIONS (Save & Repay)
function saveLoan(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. เตรียมข้อมูล Master
    var mSheet = ss.getSheetByName(SHEET_NAME); // m_actionplan
    var mData = mSheet.getDataRange().getValues();
    
    // ค้นหาแถวของโครงการใน Master
    var idxID = 0; // Col A
    var idxLoan = 18; // Col S (เงินยืมสะสมใน Master)
    var rowIndex = -1;
    
    for (var i = 1; i < mData.length; i++) { 
      if (String(mData[i][idxID]) === String(form.projectId)) { 
        rowIndex = i + 1; 
        break; 
      } 
    }
    
    // อัปเดตยอดเงินยืมสะสมใน Master (บวกเพิ่ม)
    if (rowIndex !== -1) {
       var cur = (parseFloat(String(mSheet.getRange(rowIndex, idxLoan+1).getValue()).replace(/,/g,'')) || 0) + parseFloat(form.amount);
       mSheet.getRange(rowIndex, idxLoan+1).setValue(cur);
    }
    
    // 2. บันทึกลง Transaction (t_loan)
    var tSheet = ss.getSheetByName('t_loan');
    var r = mData[rowIndex-1]; // ดึงข้อมูลแถวนั้นจาก Master มาใช้

    // แปลงวันที่เริ่ม-สิ้นสุด (ถ้ามี)
    var sDate = form.startDate ? new Date(form.startDate) : "";
    var eDate = form.endDate ? new Date(form.endDate) : "";

    // 🔥 เรียงข้อมูลลงตาราง t_loan (27 คอลัมน์ ตามที่นายท่านระบุ)
    tSheet.appendRow([
       new Date(),       // 1. ประทับเวลา (A)
       r[0],             // 2. รหัสโครงการ (B)
       r[1],             // 3. ปีงบประมาณ (C)
       r[2],             // 4. หมวด (D)
       r[3],             // 5. ลำดับโครงการ (E)
       r[4],             // 6. กลุ่มงาน/งาน (F)
       r[5],             // 7. แผนงาน (G)
       r[6],             // 8. โครงการ (H)
       r[7],             // 9. กิจกรรมหลัก (I)
       r[8],             // 10. กิจกรรมย่อย (J)
       r[9],             // 11. ประเภทงบ (K)
       r[10],            // 12. แหล่งงบประมาณ (L)
       r[13],            // 13. รหัสงบประมาณ (M)
       r[14],            // 14. รหัสกิจกรรม (N)
       r[16],            // 15. จัดสรร (O)
       form.amount,      // 16. เงินยืม (P) -> รับจากฟอร์ม
       form.loanDate,    // 17. วันที่ยืมเงิน (Q)
       form.loanType,    // 18. ประเภทเงินยืม (R)
       form.desc,        // 19. รายละเอียด (S)
       "ยังไม่ดำเนินการ",  // 20. สถานะการเบิกจ่าย (T)
       0,                // 21. จำนวนเบิกจ่าย (U) -> เริ่มต้นเป็น 0
       form.amount,      // 22. คงเหลือ (V) -> เริ่มต้นเท่ากับยอดกู้
       "",               // 23. วันที่เบิกจ่าย (W) -> ว่างไว้
       "",               // 24. ระยะเวลาที่ยืม (X) -> ว่างไว้
       "",               // 25. หมายเหตุ (Y) -> ว่างไว้
       sDate,            // 26. เริ่มดำเนินการ (Z) ✅
       eDate             // 27. สิ้นสุดดำเนินการ (AA) ✅
    ]);
    
    return { status: 'success', message: 'บันทึกเงินยืมเรียบร้อย' };

  } catch (e) { 
    return { status: 'error', message: e.message }; 
  } finally { 
    lock.releaseLock(); 
  }
}

function updateLoanRepayment(form) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    var data = tSheet.getDataRange().getValues();
    var targetRow = -1;
    var targetTime = new Date(form.timestamp).getTime();

    for(var i=1; i<data.length; i++) {
       var rt = new Date(data[i][0]).getTime();
       if(Math.abs(rt - targetTime) < 1000) { targetRow = i+1; break; }
    }
    if (targetRow == -1) return {status:'error', message: 'ไม่พบรายการ'};

    var loanAmt = parseFloat(String(data[targetRow-1][15]).replace(/,/g,'')) || 0;
    var paidAmt = parseFloat(form.paidAmount) || 0;
    var bal = loanAmt - paidAmt;
    var status = (bal <= 0) ? "เบิกจ่ายแล้ว" : "คืนบางส่วน";
    
    // Duration Logic
    var duration = "";
    if (data[targetRow-1][16] && form.payDate) {
        var d1 = new Date(data[targetRow-1][16]); var d2 = new Date(form.payDate);
        if(!isNaN(d1) && !isNaN(d2)) duration = Math.ceil((d2-d1)/(1000*60*60*24));
    }

    // Col 20=Status, 21=Paid, 22=Bal, 23=PayDate, 24=Duration
    tSheet.getRange(targetRow, 20).setValue(status);
    tSheet.getRange(targetRow, 21).setValue(paidAmt);
    tSheet.getRange(targetRow, 22).setValue(bal);
    tSheet.getRange(targetRow, 23).setValue(form.payDate);
    tSheet.getRange(targetRow, 24).setValue(duration);

    return { status: 'success', message: 'บันทึกคืนเงินเรียบร้อย' };
  } catch (e) { return { status: 'error', message: e.message }; } finally { lock.releaseLock(); }
}

// 6. HISTORY GETTERS (Fixed Indices)
function getTransactionHistory() { return getHistory('t_actionplan'); }
function getLoanHistory() { return getHistory('t_loan'); }

function getHistory(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName(sheetName);
    if (!tSheet) return [];
    var data = tSheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var result = [];
    var parseAmount = (v) => parseFloat(String(v).replace(/,/g, '')) || 0;

    for (var i = data.length - 1; i >= 1; i--) { 
      var row = data[i];
      if (!row || !row[0]) continue;
      
      var item = {};
      if(sheetName === 't_actionplan') {
          // [4]Order, [7]Proj, [8]Act, [9]Sub, [15]Amt, [17]Date, [18]Type, [11]Source, [19]Desc, [1]ID
          item = {
             rowId: i+1,
             order: row[4], project: row[7], activity: row[8], subActivity: row[9],
             amount: parseAmount(row[15]),
             date: (row[17] instanceof Date) ? Utilities.formatDate(row[17], "Asia/Bangkok", "dd/MM/yyyy") : row[17],
             type: row[18], source: row[11], desc: row[19], id: row[1]
          };
      } else { // t_loan
          // [4]Order, [7]Proj, [15]Amt, [16]LoanDate, [19]Status, [20]Paid, [21]Bal
          item = {
             timestamp: (row[0] instanceof Date) ? row[0].toISOString() : "",
             order: row[4], project: row[7], activity: row[8], subActivity: row[9],
             amount: parseAmount(row[15]),
             date: (row[16] instanceof Date) ? Utilities.formatDate(row[16], "Asia/Bangkok", "dd/MM/yyyy") : row[16],
             status: row[19], paid: parseAmount(row[20]), balance: parseAmount(row[21])
          };
      }
      if(item.amount > 0 || item.order) result.push(item);
      if (result.length >= 20) break;
    }
    return result;
  } catch(e) { return []; }
}

// ==========================================
// 7. SUMMARY REPORT (HARDCODED INDEX VERSION) 📊
// ==========================================
function getSummaryData(quarterFilter, monthFilter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('m_actionplan');
    if (!sheet) return { error: "ไม่พบ Sheet ข้อมูล" };
    
    var data = sheet.getDataRange().getValues();
    // ไม่ต้อง data.shift() ก็ได้ เดี๋ยวเราเริ่ม loop ที่แถว 1

    // 🎯 HARDCODE INDEX (ตาม Log ที่ท่านส่งมา)
    var I_DEPT = 4;      // กลุ่มงาน [4]
    var I_TYPE = 9;      // ประเภทงบ [9]
    var I_SOURCE = 10;   // แหล่งงบ [10]
    var I_ALLOC = 16;    // จัดสรร [16]
    var I_SPENT = 17;    // เบิกจ่าย [17]
    
    // Index เดือน ต.ค.(26) - ก.ย.(37)
    var I_MONTHS = [26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37];
    var quarters = { 'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11] };
    
    var parseNum = function(val) { 
        var v = parseFloat(String(val).replace(/,/g, '')); 
        return isNaN(v) ? 0 : v; 
    };

    // เตรียมตัวแปรเก็บผลรวม
    var grandTotal = { allocated: 0, spent: 0, count: 0 };
    var bySource = {}; 
    var byDeptAll = {}, byDeptMoph = {}, byDeptNon = {};

    // เริ่มวนลูป (ข้ามหัวตารางแถว 0)
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        
        // --- 1. เช็ค Timeline (ตัวกรองเวลา) ---
        var timeline = I_MONTHS.map(function(idx) { 
            return (String(row[idx] || "").trim() !== '') ? 1 : 0; 
        });

        var passTime = true;
        if (quarterFilter && quarters[quarterFilter]) { 
            if (!quarters[quarterFilter].some(function(mIdx) { return timeline[mIdx] === 1; })) passTime = false; 
        }
        if (monthFilter && String(monthFilter) !== "") { 
            if (timeline[parseInt(monthFilter)] !== 1) passTime = false; 
        }

        if (passTime) {
            var alloc = parseNum(row[I_ALLOC]);
            var spent = parseNum(row[I_SPENT]);
            var typeVal = String(row[I_TYPE] || "").trim();
            var srcVal = String(row[I_SOURCE] || "").trim();
            var deptVal = String(row[I_DEPT] || "ไม่ระบุ").trim(); 
            if(deptVal === "") deptVal = "ไม่ระบุ";

            // Logic แยกประเภทงบ (MOPH vs Non-MOPH)
            var isMoph = (
                typeVal.indexOf('งบประมาณ') > -1 || typeVal.indexOf('สป.') > -1 || 
                srcVal.indexOf('MOPH') > -1 || srcVal.indexOf('งบประมาณ') > -1 || srcVal.indexOf('สป.') > -1 ||
                typeVal === 'PP' || typeVal === 'OP' || typeVal.indexOf('งบดำเนินงาน') > -1
            );
            
            var sourceGroup = isMoph ? "งบประมาณ (สป.สธ.)" : "เงินนอกงบประมาณ (Non-MOPH)";

            // --- 2. บวกยอดเข้ากองกลาง ---
            grandTotal.allocated += alloc; 
            grandTotal.spent += spent; 
            grandTotal.count++;

            // --- 3. บวกยอดแยกตามแหล่งเงิน ---
            if (!bySource[sourceGroup]) bySource[sourceGroup] = { allocated: 0, spent: 0, count: 0 };
            bySource[sourceGroup].allocated += alloc; 
            bySource[sourceGroup].spent += spent; 
            bySource[sourceGroup].count++;

            // --- 4. บวกยอดแยกตามกลุ่มงาน (All) ---
            if (!byDeptAll[deptVal]) byDeptAll[deptVal] = { allocated: 0, spent: 0, count: 0 };
            byDeptAll[deptVal].allocated += alloc; 
            byDeptAll[deptVal].spent += spent; 
            byDeptAll[deptVal].count++;

            // --- 5. บวกยอดแยกตามกลุ่มงาน (MOPH / Non) ---
            if (isMoph) {
                if (!byDeptMoph[deptVal]) byDeptMoph[deptVal] = { allocated: 0, spent: 0, count: 0 };
                byDeptMoph[deptVal].allocated += alloc; 
                byDeptMoph[deptVal].spent += spent; 
                byDeptMoph[deptVal].count++;
            } else {
                if (!byDeptNon[deptVal]) byDeptNon[deptVal] = { allocated: 0, spent: 0, count: 0 };
                byDeptNon[deptVal].allocated += alloc; 
                byDeptNon[deptVal].spent += spent; 
                byDeptNon[deptVal].count++;
            }
        }
    }

    // Helper แปลง Object เป็น Array เพื่อส่งกลับไปวนลูปหน้าบ้าน
    var toList = function(obj) {
        var list = [];
        for (var k in obj) {
            list.push({ name: k, allocated: obj[k].allocated, spent: obj[k].spent, count: obj[k].count });
        }
        return list.sort(function(a, b) { return b.allocated - a.allocated; }); // เรียงตามงบมากไปน้อย
    };

    return {
        grandTotal: grandTotal,
        bySource: toList(bySource),
        byDeptAll: toList(byDeptAll),
        byDeptMoph: toList(byDeptMoph),
        byDeptNon: toList(byDeptNon)
    };

  } catch (e) { return { error: e.message }; }
}

// ==========================================
// 8. DRILL-DOWN DETAILS (SUPER MATCHER - IGNORE SLASH/SPACE) 🛡️✅
// ==========================================
function getDeptDetails(deptName, quarterFilter, monthFilter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('m_actionplan');
    if (!sheet) return { error: "ไม่พบ Sheet m_actionplan" };
    
    var data = sheet.getDataRange().getValues();

    // 🎯 HARDCODE INDEX (ตาม Log เป๊ะๆ)
    var I_DEPT = 4;
    var I_PROJ = 6;
    var I_ACT = 7;
    var I_TYPE = 9;
    var I_SOURCE = 10;
    var I_ALLOC = 16;
    var I_SPENT = 17;
    
    var I_MONTHS = [26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37];
    var quarters = { 'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11] };
    
    var parseNum = function(val) { 
        var v = parseFloat(String(val).replace(/,/g, '')); 
        return isNaN(v) ? 0 : v; 
    };

    // 🧼 ฟังก์ชันเครื่องบดละเอียด: ลบ Space, /, -, _ ออกให้หมด เหลือแต่ตัวหนังสือ
    var cleanName = function(str) {
        return String(str).replace(/[\s\/\-_]+/g, "").trim(); 
    };

    var projectsAll = [], projectsMoph = [], projectsNon = [];
    var sumAll = { allocated: 0, spent: 0 }, sumMoph = { allocated: 0, spent: 0 }, sumNon = { allocated: 0, spent: 0 };
    
    // เตรียมชื่อเป้าหมายแบบ "สะอาด"
    var targetClean = cleanName(deptName);

    // เริ่มวนลูป (ข้ามหัวตาราง)
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        
        // 1. ดึงชื่อจาก Excel แล้วทำความสะอาดก่อนเทียบ
        var rowDeptRaw = String(row[I_DEPT] || "");
        var rowClean = cleanName(rowDeptRaw);

        // 🔥 LOGIC การเทียบ: เทียบแบบเนื้อเน้นๆ (ไม่มี Slash มาขวาง)
        // ถ้าเนื้อในไม่เหมือนกัน -> ข้าม
        if (rowClean.indexOf(targetClean) === -1 && targetClean.indexOf(rowClean) === -1) {
             continue; 
        }

        // 2. เช็ค Timeline
        var timeline = I_MONTHS.map(function(idx) { 
            return (String(row[idx] || "").trim() !== '') ? 1 : 0; 
        });

        var passTime = true;
        if (quarterFilter && quarters[quarterFilter]) { 
            if (!quarters[quarterFilter].some(function(mIdx) { return timeline[mIdx] === 1; })) passTime = false; 
        }
        if (monthFilter && String(monthFilter) !== "") { 
            if (timeline[parseInt(monthFilter)] !== 1) passTime = false; 
        }

        // 3. เก็บข้อมูล
        if (passTime) {
            var alloc = parseNum(row[I_ALLOC]);
            var spent = parseNum(row[I_SPENT]);
            var typeVal = String(row[I_TYPE] || "").trim();
            var srcVal = String(row[I_SOURCE] || "").trim();

            var projObj = {
                project: String(row[I_PROJ] || "-"),
                activity: String(row[I_ACT] || "-"),
                allocated: alloc, 
                spent: spent, 
                balance: alloc - spent,
                progress: (alloc > 0) ? (spent / alloc * 100) : 0
            };

            projectsAll.push(projObj);
            sumAll.allocated += alloc; sumAll.spent += spent;

            // Logic แยกประเภท
            var isMoph = (
                typeVal.indexOf('งบประมาณ') > -1 || typeVal.indexOf('สป.') > -1 || 
                srcVal.indexOf('MOPH') > -1 || srcVal.indexOf('งบประมาณ') > -1 || srcVal.indexOf('สป.') > -1 ||
                typeVal === 'PP' || typeVal === 'OP' || typeVal.indexOf('งบดำเนินงาน') > -1
            );

            if (isMoph) {
                projectsMoph.push(projObj); sumMoph.allocated += alloc; sumMoph.spent += spent;
            } else {
                projectsNon.push(projObj); sumNon.allocated += alloc; sumNon.spent += spent;
            }
        }
    }

    var sortFn = function(a, b) { return b.progress - a.progress; };
    projectsAll.sort(sortFn); 
    projectsMoph.sort(sortFn); 
    projectsNon.sort(sortFn);

    return {
        projectsAll: projectsAll, projectsMoph: projectsMoph, projectsNon: projectsNon,
        sumAll: sumAll, sumMoph: sumMoph, sumNon: sumNon,
        deptName: deptName
    };

  } catch (e) { 
      return { error: "Server Error: " + e.message }; 
  }
}

// ==========================================
// 9. Search Loan 
// ==========================================
function searchLoanHistory(criteria) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('t_loan');
  var data = sheet.getDataRange().getValues();
  var result = [];
  
  // column index: B=Order(1), H=Project(7), I=Activity(8), J=SubActivity(9)
  // ปรับ index ให้ตรงกับ Sheet จริงของนายท่าน
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var match = true;
    
    if (criteria.order && String(row[4]) != String(criteria.order)) match = false; // Col E = Index 4 (ลำดับ)
    if (match && criteria.project && String(row[7]) != String(criteria.project)) match = false; // Col H = Index 7
    if (match && criteria.activity && String(row[8]) != String(criteria.activity)) match = false; // Col I = Index 8
    if (match && criteria.subActivity && String(row[9]) != String(criteria.subActivity)) match = false; // Col J = Index 9
    
    if (match) {
        // ... จัดรูปแบบข้อมูลใส่ array result ...
        // (ใช้ Logic เดียวกับ getLoanHistory)
        // ...
    }
  }
  return result;
}
