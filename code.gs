var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ'; 
var SHEET_NAME = 'm_actionplan';
var APP_VERSION = '6900210-1630'; 

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
    var sheet = ss.getSheetByName('m_actionplan');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    data.shift(); // ตัดหัวตาราง

    // 🎯 MAPPING INDEX (เช็คตามไฟล์จริงของท่าน)
    // [0]=ID, [2]=หมวด, [3]=ลำดับ, [4]=กลุ่มงาน, [5]=แผนงาน
    // [6]=โครงการ, [7]=กิจกรรม, [8]=ย่อย
    // [9]=ประเภท, [10]=แหล่ง
    // [15]=อนุมัติ (Column P) ✅ เพิ่มตัวนี้
    // [16]=จัดสรร (Column Q)
    // [17]=เบิกจ่าย (Column R)
    // [19]=คงเหลือ (Column T)
    
    return data.map(function(r) {
      return {
        id: r[0],
        category: r[2], 
        order: r[3],
        dept: r[4],
        plan: r[5],
        project: r[6],
        activity: r[7],
        subActivity: r[8],
        budgetType: r[9],
        budgetSource: r[10],
        
        // ✅ ดึงยอดอนุมัติจาก Column P (Index 15)
        approved: parseFloat(String(r[15]).replace(/,/g,'')) || 0,
        
        allocated: parseFloat(String(r[16]).replace(/,/g,'')) || 0,
        spent: parseFloat(String(r[17]).replace(/,/g,'')) || 0,
        balance: parseFloat(String(r[19]).replace(/,/g,'')) || 0
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

// 📌 [แทนที่] ฟังก์ชันนี้ในไฟล์ Code.gs ครับ
function updateLoanRepayment(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ====================================================
    // PART 1: อัปเดตสถานะในตาราง "เงินยืม" (t_loan)
    // ====================================================
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) throw new Error("ไม่พบ Sheet: t_loan");

    var tData = tSheet.getDataRange().getValues();
    var tRowIndex = -1;
    var projectId = ""; 
    var loanAmount = 0; 

    var targetDate = new Date(data.timestamp); 

    // วนลูปค้นหา
    for (var i = 1; i < tData.length; i++) {
      var cellValue = tData[i][0]; 
      var isMatch = false;

      if (String(cellValue) == String(data.timestamp)) {
        isMatch = true;
      } else {
        var cellDate = new Date(cellValue);
        if (!isNaN(cellDate.getTime()) && !isNaN(targetDate.getTime())) {
          // ยอมรับความคลาดเคลื่อนได้ 60 วินาที
          if (Math.abs(cellDate.getTime() - targetDate.getTime()) < 60000) { 
             isMatch = true;
          }
        }
      }

      if (isMatch) {
        tRowIndex = i + 1;
        projectId = tData[i][1];     
        loanAmount = parseFloat(tData[i][15] || 0); 
        break;
      }
    }

    if (tRowIndex === -1) throw new Error("ไม่พบรายการกู้ยืม (Timestamp ไม่ตรง)");

    // คำนวณยอดใน t_loan
    var currentPaid = parseFloat(tData[tRowIndex-1][20] || 0); 
    var payAmount = parseFloat(data.paidAmount); 
    var newPaid = currentPaid + payAmount;
    var newBalance = loanAmount - newPaid;

    var status = (newBalance <= 0.01) ? "คืนครบ" : "คืนบางส่วน";
    if (newBalance < 0) newBalance = 0;

    // บันทึกลง t_loan
    tSheet.getRange(tRowIndex, 20).setValue(status);       
    tSheet.getRange(tRowIndex, 21).setValue(newPaid);      
    tSheet.getRange(tRowIndex, 22).setValue(newBalance);   
    tSheet.getRange(tRowIndex, 23).setValue(data.payDate); 

    // ====================================================
    // PART 2: ตัดงบประมาณใน "แผนงาน" (m_actionplan)  <-- ✅ แก้ไขจุดนี้
    // ====================================================
    if (projectId) {
      var mSheet = ss.getSheetByName('m_actionplan');
      if (mSheet) {
        var mData = mSheet.getDataRange().getValues();
        var mRowIndex = -1;

        // ค้นหาบรรทัดโครงการ
        for (var j = 1; j < mData.length; j++) {
          if (String(mData[j][0]) == String(projectId)) { // Col A: ID
            mRowIndex = j + 1;
            break;
          }
        }

        if (mRowIndex !== -1) {
          // 🎯 กำหนดคอลัมน์ใหม่ (ตามที่นายท่านแจ้ง)
          var colAlloc = 17;   // Col Q = 17 (ยอดจัดสรร)
          var colSpent = 18;   // Col R = 18 (เบิกจ่ายสะสม)
          var colBalance = 20; // Col T = 20 (คงเหลือ ไม่รวมเงินยืม)
          // Col U (21) คงเหลือรวมเงินยืม เราจะไม่ยุ่ง ปล่อยให้สูตรใน Sheet ทำงาน หรือคงเดิมไว้

          // 1. อ่านค่ายอดจัดสรร (Allocated)
          var cellAlloc = mSheet.getRange(mRowIndex, colAlloc);
          var allocated = parseFloat(cellAlloc.getValue()) || 0;

          // 2. อ่านค่าเบิกจ่ายเดิม (Current Spent)
          var cellSpent = mSheet.getRange(mRowIndex, colSpent);
          var currentSpent = parseFloat(cellSpent.getValue()) || 0;

          // 3. คำนวณใหม่
          // เบิกจ่ายใหม่ = เบิกจ่ายเดิม + ยอดที่เอามาล้างหนี้ (บิล)
          var updatedSpent = currentSpent + payAmount; 
          
          // คงเหลือใหม่ (Col T) = จัดสรร - เบิกจ่ายใหม่
          var updatedBalance = allocated - updatedSpent;

          // 4. บันทึกกลับ
          cellSpent.setValue(updatedSpent);        // ลงช่อง R
          mSheet.getRange(mRowIndex, colBalance).setValue(updatedBalance); // ลงช่อง T
        }
      }
    }

    return { status: 'success', message: 'บันทึกคืนเงินและตัดงบประมาณเรียบร้อย' };

  } catch (e) {
    return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 6. HISTORY GETTERS (Fixed Indices)
function getTransactionHistory() { return getHistory('t_actionplan'); }

  // function getLoanHistory() { 
  //   return getHistory('t_loan'); 
  //   }
  // 📌 [แทนที่] ฟังก์ชัน getLoanHistory ในไฟล์ Code.gs (เพิ่มวันที่ไทย)
function getLoanHistory() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];

    // 1. เตรียมข้อมูล Master (VLOOKUP)
    var projectMap = {};
    try {
      var mSheet = ss.getSheetByName('m_actionplan');
      if (mSheet) {
        var mData = mSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < mData.length; i++) {
          var pid = String(mData[i][0]).trim();
          if (pid) {
            projectMap[pid] = { type: mData[i][9] || "-", source: mData[i][10] || "-" };
          }
        }
      }
    } catch (e) { console.log("Map Error: " + e); }

    // 2. ดึงข้อมูล Transaction
    var tData = tSheet.getDataRange().getDisplayValues(); 
    var result = [];
    var parseNum = function(val) { return parseFloat(String(val).replace(/,/g, '')) || 0; };
    
    // 🗓️ ฟังก์ชันแปลงวันที่ไทย (เช่น 2026-02-09 -> 9 ก.พ. 2569)
    var toThaiDate = function(val) {
      if (!val) return "-";
      try {
        var d;
        // กรณี 1: เป็น Date Object
        if (Object.prototype.toString.call(val) === '[object Date]') d = val;
        // กรณี 2: เป็น String YYYY-MM-DD
        else if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}$/)) {
          var parts = val.split('-'); d = new Date(parts[0], parts[1]-1, parts[2]);
        }
        // กรณี 3: String อื่นๆ พยายามแปลง
        else { d = new Date(val); }

        if (isNaN(d.getTime())) return String(val); // แปลงไม่ได้ส่งค่าเดิม

        var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
        return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
      } catch (ex) { return String(val); }
    };

    for (var i = tData.length - 1; i >= 1; i--) {
      try {
        var row = tData[i];
        if (!row[0] && !row[1]) continue;

        var pid = String(row[1] || "").trim();
        var meta = projectMap[pid] || { type: '-', source: '-' };

        result.push({
          id: row[0], timestamp: row[0],
          project: row[7], activity: row[8], subActivity: row[9],
          amount: parseNum(row[15]),
          date: toThaiDate(row[16]), // ✅ แปลงเป็นวันที่ไทยตรงนี้เลย
          status: row[19],
          paid: parseNum(row[20]),
          balance: parseNum(row[21]),
          order: row[4],
          type: row[17],
          desc: row[18],
          budgetType: meta.type,     
          budgetSource: meta.source, 
          dept: row[2]
        });

      } catch (e) { console.log("Row Error ("+i+"): " + e); }
    }
    return result;
  }
// จบ function ประวัติการยืมเงิน

//function ดึงรายการเบิกจ่าย
function getHistory(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName(sheetName);
    if (!tSheet) return [];
    
    // ใช้ getDisplayValues เพื่อความชัวร์ (เหมือนตารางเงินยืม)
    var data = tSheet.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    
    var result = [];
    var parseAmount = function(v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };

    // 🗓️ ฟังก์ชันแปลงวันที่ไทย (Reusable)
    var toThaiDate = function(val) {
      if (!val) return "-";
      try {
        var d;
        // กรณี 1: เป็น Date Object
        if (Object.prototype.toString.call(val) === '[object Date]') d = val;
        // กรณี 2: เป็น String YYYY-MM-DD
        else if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}$/)) {
           var parts = val.split('-'); d = new Date(parts[0], parts[1]-1, parts[2]);
        }
        // กรณี 3: String อื่นๆ (เช่น จาก getDisplayValues)
        else { d = new Date(val); }

        if (isNaN(d.getTime())) return String(val); // ถ้าแปลงไม่ได้ ส่งค่าเดิม

        var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
        return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
      } catch (ex) { return String(val); }
    };

    // เริ่มวนลูปจากล่าสุด (ล่างขึ้นบน)
    for (var i = data.length - 1; i >= 1; i--) { 
      var row = data[i];
      if (!row || (!row[0] && !row[1])) continue;
      
      var item = {};
      
      if(sheetName === 't_actionplan') {
          // 📝 โหมดประวัติการเบิกจ่าย
          // [4]Order, [7]Proj, [8]Act, [9]Sub, [15]Amt, [17]Date, [18]Type, [11]Source, [19]Desc, [1]ID
          item = {
             rowId: i+1,
             order: row[4], 
             project: row[7], 
             activity: row[8], 
             subActivity: row[9],
             amount: parseAmount(row[15]),
             date: toThaiDate(row[17]), // ✅ แปลงวันที่เป็นไทย (เช่น 1 ต.ค. 2569)
             type: row[18], 
             source: row[11], 
             desc: row[19], 
             id: row[1]
          };
      } 
      else { 
          // 📝 โหมดอื่นๆ (เผื่อไว้)
          item = {
             timestamp: row[0],
             order: row[4], project: row[7],
             amount: parseAmount(row[15]),
             date: toThaiDate(row[16]),
             status: row[19]
          };
      }

      if(item.amount > 0 || item.order) result.push(item);
      if (result.length >= 50) break; // Limit 50 รายการล่าสุด
    }
    return result;
  } catch(e) { 
    console.log("getHistory Error: " + e);
    return []; 
  }
}
//จบ function ดึงรายการเบิกจ่าย


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
  var tSheet = ss.getSheetByName('t_loan');
  var mSheet = ss.getSheetByName('m_actionplan');

  var projectMap = {};
  if (mSheet) {
    var mData = mSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < mData.length; i++) {
      var pid = String(mData[i][0]).trim();
      projectMap[pid] = { type: mData[i][9], source: mData[i][10] };
    }
  }

  // ใช้ getDisplayValues เหมือนกัน
  var data = tSheet.getDataRange().getDisplayValues();
  var result = [];
  var parseNum = function(v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var match = true;
    
    // Logic กรองข้อมูล
    if (criteria.order && String(row[4]) != String(criteria.order)) match = false;
    if (match && criteria.project && String(row[7]).indexOf(criteria.project) === -1) match = false; // ใช้ indexOf เพื่อให้ค้นหาบางส่วนได้

    if (match) {
        var pid = String(row[1]).trim();
        var meta = projectMap[pid] || { type: '-', source: '-' };
        
        result.push({
          id: row[0], timestamp: row[0], project: row[7], activity: row[8], subActivity: row[9],
          amount: parseNum(row[15]),
          date: row[16], // ✅ Col Q
          status: row[19], paid: parseNum(row[20]), balance: parseNum(row[21]), order: row[4],
          type: row[17], desc: row[18], // ✅ Col R, S
          budgetType: meta.type, budgetSource: meta.source, dept: row[2]
        });
    }
  }
  return result;
}
// จบ function Search Loan 

  //เริ่ม function จัดสรรงบประมาณ
  // ==========================================
  // 9. NEW ALLOCATION SYSTEM (Backend)
  // ==========================================

  function saveAllocation(form) {
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000); // ป้องกันการบันทึกชนกัน
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var mSheet = ss.getSheetByName('m_actionplan');
      var tAllocSheet = ss.getSheetByName('t_allocate');
      
      // 1. หาแถวใน Master Plan (m_actionplan)
      var mData = mSheet.getDataRange().getValues();
      var rowIndex = -1;
      for (var i = 1; i < mData.length; i++) {
        if (String(mData[i][0]) == String(form.id)) { // เทียบ ID
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการในฐานข้อมูล' };

      // 2. คำนวณยอดเงินใหม่ (Accumulate Logic)
      // Col Q (Index 16+1 = 17) คือ ยอดจัดสรร
      var cellAlloc = mSheet.getRange(rowIndex, 17);
      var currentTotal = parseFloat(String(cellAlloc.getValue()).replace(/,/g,'')) || 0;
      var newTotal = currentTotal + parseFloat(form.currentAlloc);

      // 3. อัปเดต Master Plan
      cellAlloc.setValue(newTotal);

      // 4. บันทึก Log ลง t_allocate
      // โครงสร้าง Columns A-S (19 Columns)
      var r = form.fullData; // ข้อมูลเดิมจาก Client (เพื่อความสะดวกในการ Mapping)
      
      // *หมายเหตุ: บาง Field ถ้าไม่มีใน JSON ให้เว้นว่าง หรือดึงจาก mData ก็ได้
      // เพื่อความแม่นยำ ผมจะดึงจาก mData แถวที่เจอ
      var mRow = mData[rowIndex-1]; 
      
      var logRow = [
        new Date(),       // A: Timestamp
        mRow[0],          // B: รหัสโครงการ (ID)
        mRow[1],          // C: ปีงบ
        mRow[2],          // D: หมวด
        mRow[3],          // E: ลำดับ
        mRow[4],          // F: กลุ่มงาน
        mRow[5],          // G: แผนงาน
        mRow[6],          // H: โครงการ
        mRow[7],          // I: กิจกรรมหลัก
        mRow[8],          // J: กิจกรรมย่อย
        mRow[9],          // K: ประเภทงบ
        mRow[10],         // L: แหล่งงบ
        mRow[13],         // M: รหัสงบประมาณ
        mRow[14],         // N: รหัสกิจกรรม
        newTotal,         // O: จัดสรรสะสม (ยอดใหม่) ✅
        form.currentAlloc,// P: จัดสรรครั้งนี้ ✅
        form.date,        // Q: วันที่จัดสรร ✅
        form.letterNo,    // R: เลขหนังสือ ✅
        form.remark       // S: หมายเหตุ ✅
      ];

      tAllocSheet.appendRow(logRow);

      return { status: 'success', message: 'บันทึกเรียบร้อย' };

    } catch (e) {
      return { status: 'error', message: e.message };
    } finally {
      lock.releaseLock();
    }
  }

  function getAllocationHistory() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName('t_allocate');
      if (!sheet) return [];
      
      // อ่านข้อมูลทั้งหมด (ใช้ getDisplayValues เพื่อความชัวร์เรื่องวันที่)
      var data = sheet.getDataRange().getDisplayValues();
      if (data.length < 2) return [];

      var result = [];
      var parseNum = (v) => parseFloat(String(v).replace(/,/g,'')) || 0;
      
      // วนลูปจากล่าสุด (ล่างขึ้นบน)
      for (var i = data.length - 1; i >= 1; i--) {
        var row = data[i];
        if (!row[1]) continue; // ไม่มี ID ข้าม

        // Map Data เพื่อส่งกลับไปแสดงผล
        result.push({
          id: row[1],
          order: row[4],       // E
          project: row[7],     // H
          activity: row[8],    // I
          subActivity: row[9], // J
          type: row[10],       // K
          source: row[11],     // L
          accumulatedAlloc: parseNum(row[14]), // O (สะสม)
          currentAlloc: parseNum(row[15]),     // P (ครั้งนี้)
          date: formatDateThai(row[16]),       // Q (วันที่ - แปลงเป็นไทย)
          letterNo: row[17]    // R
        });
        
        if (result.length >= 100) break; // Limit 100 รายการล่าสุด
      }
      return result;

    } catch (e) { return []; }
  }

  // Helper: แปลงวันที่ไทย (Reused Logic)
  function formatDateThai(dateStr) {
    if(!dateStr) return "-";
    try {
      var d = new Date(dateStr);
      if(isNaN(d.getTime())) return dateStr;
      var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
      return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
    } catch(e) { return dateStr; }
  }
//สิ้นสุด function จัดสรรงบประมาณ
