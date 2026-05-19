// ==========================================
// 1. CONFIGURATION & SETUP
// ==========================================
var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ';
var SHEET_NAME = 'm_actionplan';
var APP_VERSION = 'Version : 690514';

// ==========================================
// 📌 1.1 DYNAMIC COLUMN MAPPING (อัปเดตเพิ่ม 3 คอลัมน์ใหม่)
// ==========================================
var COL_NAME = {
    ID: 'รหัสโครงการ',
    YEAR: 'ปีงบประมาณ',
    CAT: 'หมวด',
    ORDER: 'ลำดับโครงการ',
    DEPT: 'กลุ่มงาน/งาน',
    PLAN: 'แผนงาน',
    PROJ: 'โครงการ',
    ACT: 'กิจกรรมหลัก',
    SUB: 'กิจกรรมย่อย',
    TYPE: 'ประเภทงบ',
    SOURCE: 'แหล่งงบประมาณ',
    RESP: 'ผู้รับผิดชอบ',
    BUDGET_CODE: 'รหัสงบประมาณ',
    ACT_CODE: 'รหัสกิจกรรม',
    APPROVE: 'อนุมัติตามแผน',
    ALLOC: 'จัดสรร',
    SPENT: 'เบิกจ่าย',
    LOAN: 'เงินยืม',
    BAL: 'คงเหลือ (ไม่รวมเงินยืม)',
    STATUS: 'สถานะ',
    REMARK: 'หมายเหตุการปรับ',
    EXPENSE_DETAIL: 'รายละเอียดค่าใช้จ่าย',
    OP_TYPE: 'ประเภทการดำเนินงาน',
    EXP_TYPE: 'ประเภทค่าใช้จ่าย',
    PROJ_REMARK: 'หมายเหตุโครงการ/กิจกรรม',
    M_OCT: 'ต.ค.', M_NOV: 'พ.ย.', M_DEC: 'ธ.ค.',
    M_JAN: 'ม.ค.', M_FEB: 'ก.พ.', M_MAR: 'มี.ค.',
    M_APR: 'เม.ย.', M_MAY: 'พ.ค.', M_JUN: 'มิ.ย.',
    M_JUL: 'ก.ค.', M_AUG: 'ส.ค.', M_SEP: 'ก.ย.'
};

// ==========================================
// 📌 1.2 ฟังก์ชันสร้างแผนที่คอลัมน์ (Header Radar) - ซ่อมบั๊กหน้าเว็บว่างเปล่า
// ==========================================
function getColumnMap(sheet) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var map = {};

    for (var i = 0; i < headers.length; i++) {
        var headerName = String(headers[i]).trim();
        if (headerName === "") continue;

        // 🌟 จุดที่แก้ไข: ให้ระบบจดจำพิกัดด้วย "ชื่อเต็มของหัวตาราง" ด้วย เพื่อให้ส่งข้อมูลไปหน้าเว็บได้
        map[headerName] = i;

        // จดจำพิกัดด้วย "คีย์ย่อ" สำหรับใช้ในฟังก์ชันบันทึกข้อมูล
        for (var key in COL_NAME) {
            if (headerName === COL_NAME[key]) {
                map[key] = i;
                break;
            }
        }
    }

    return map;
}

function doGet() {
    var template = HtmlService.createTemplateFromFile('index');
    template.appVersion = APP_VERSION;
    return template.evaluate()
        .setTitle('LPHO BMD : Loei Provincial Public Health Office Budget Management Dashboard')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 📌 UNIVERSAL MAPPER: ตัวแปลงข้อมูลมาตรฐาน (เพิ่ม 4 คอลัมน์ใหม่)
// ==========================================
function mapProjectRow(row, map) {
    var parseNum = function (val) {
        var v = parseFloat(String(val).replace(/,/g, ''));
        return isNaN(v) ? 0 : v;
    };

    return {
        id: String(row[map[COL_NAME.ID]] || "").trim(),
        order: String(row[map[COL_NAME.ORDER]] || "").trim(),
        dept: String(row[map[COL_NAME.DEPT]] || "").trim(),
        proj: String(row[map[COL_NAME.PROJ]] || "").trim(),
        act: String(row[map[COL_NAME.ACT]] || "").trim(),
        subAct: String(row[map[COL_NAME.SUB]] || "").trim(),
        type: String(row[map[COL_NAME.TYPE]] || "").trim(),
        source: String(row[map[COL_NAME.SOURCE]] || "").trim(),
        resp: String(row[map[COL_NAME.RESP]] || "").trim(),
        budgetCode: String(row[map[COL_NAME.BUDGET_CODE]] || "").trim(),
        actCode: String(row[map[COL_NAME.ACT_CODE]] || "").trim(),
        approved: parseNum(row[map[COL_NAME.APPROVE]]),
        allocated: parseNum(row[map[COL_NAME.ALLOC]]),
        spent: parseNum(row[map[COL_NAME.SPENT]]),
        loan: parseNum(row[map[COL_NAME.LOAN]]),
        balance: parseNum(row[map[COL_NAME.BAL]]),

        status: String(row[map[COL_NAME.STATUS]] || 'ACTIVE').trim().toUpperCase(),
        remark: String(row[map[COL_NAME.REMARK]] || "").trim(),

        // 🌟 4 คอลัมน์ใหม่ที่เพิ่มเข้ามา (เช็คก่อนว่ามีพิกัดคอลัมน์หรือไม่)
        opType: map[COL_NAME.OP_TYPE] !== undefined ? String(row[map[COL_NAME.OP_TYPE]] || "").trim() : "",
        expType: map[COL_NAME.EXP_TYPE] !== undefined ? String(row[map[COL_NAME.EXP_TYPE]] || "").trim() : "",
        projRemark: map[COL_NAME.PROJ_REMARK] !== undefined ? String(row[map[COL_NAME.PROJ_REMARK]] || "").trim() : "",
        expDetail: map[COL_NAME.EXPENSE_DETAIL] !== undefined ? String(row[map[COL_NAME.EXPENSE_DETAIL]] || "").trim() : "",

        timeline: [
            row[map[COL_NAME.M_OCT]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_NOV]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_DEC]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_JAN]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_FEB]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_MAR]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_APR]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_MAY]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_JUN]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_JUL]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_AUG]] == 1 ? 1 : 0,
            row[map[COL_NAME.M_SEP]] == 1 ? 1 : 0
        ]
    };
}

// ==========================================
// 📌 ดึงข้อมูล Master Plan (ใช้ท่อส่งมาตรฐานเดียว)
// ==========================================
function getAllMasterDataForClient() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return [];

        var map = getColumnMap(sheet);
        var data = sheet.getDataRange().getValues();
        var list = [];

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var id = String(row[map[COL_NAME.ID]] || "").trim();
            if (!id) continue; // ข้ามบรรทัดว่าง

            // 🌟 เรียกใช้ Mapper ตัวเดียวจบ ได้ข้อมูลสะอาดและมาตรฐาน 100%
            list.push(mapProjectRow(row, map));
        }
        return list;
    } catch (e) {
        return { error: e.message };
    }
}

// แอดมินและ Client ใช้ข้อมูลมาตรฐานเดียวกัน ไม่ต้องเขียนโค้ดซ้ำ
function getAdminMasterPlan() {
    return getAllMasterDataForClient();
}

// ==================
// 4. DASHBOARD DATA 
// ==================
function getDashboardData() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) return { error: "ไม่พบ Sheet" };

        var map = getColumnMap(sheet);
        var summary = {
            moph: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} },
            loeiFund: { approved: 0, allocated: 0, spent: 0, balance: 0, deptStats: {} },
            hrd1: { approved: 0, allocated: 0, spent: 0, balance: 0 },
            hrd2: { approved: 0, allocated: 0, spent: 0, balance: 0 }
        };

        var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
        var data = sheet.getDataRange().getValues();

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var typeVal = String(row[map[COL_NAME.TYPE]] || "").trim();
            var sourceVal = String(row[map[COL_NAME.SOURCE]] || "").trim();
            var budgetCodeVal = String(row[map[COL_NAME.BUDGET_CODE]] || "").trim();
            var actCodeVal = String(row[map[COL_NAME.ACT_CODE]] || "").trim();
            var statusVal = String(row[map[COL_NAME.STATUS]] || "").trim().toUpperCase();
            var dept = String(row[map[COL_NAME.DEPT]] || 'ไม่ระบุ').trim();
            if (dept === '') dept = 'ไม่ระบุ';

            if (statusVal === 'ACTIVE') {
                var approveAmt = parseNum(row[map[COL_NAME.APPROVE]]);
                var allocAmt = parseNum(row[map[COL_NAME.ALLOC]]);
                var spentAmt = parseNum(row[map[COL_NAME.SPENT]]);
                var balAmt = parseNum(row[map[COL_NAME.BAL]]);

                // 1. สป.สธ.
                var isOldNonMoph = (typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 || typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1);
                if (!isOldNonMoph) {
                    summary.moph.approved += approveAmt;
                    summary.moph.allocated += allocAmt;
                    summary.moph.spent += spentAmt;
                    summary.moph.balance += balAmt;

                    if (!summary.moph.deptStats[dept]) summary.moph.deptStats[dept] = { allocated: 0, spent: 0 };
                    summary.moph.deptStats[dept].allocated += allocAmt;
                    summary.moph.deptStats[dept].spent += spentAmt;
                }

                // 2. เงินนอก สป. (เฉพาะ เงินบำรุง สสจ.เลย)
                var cleanType = typeVal.replace(/\s+/g, '');
                var cleanSource = sourceVal.replace(/\s+/g, '');
                var isLoeiFund = (cleanType === 'เงินนอกสป.' && cleanSource === 'เงินบำรุงสสจ.เลย');

                if (isLoeiFund) {
                    summary.loeiFund.approved += approveAmt;
                    summary.loeiFund.allocated += allocAmt;
                    summary.loeiFund.spent += spentAmt;
                    summary.loeiFund.balance += balAmt;

                    if (!summary.loeiFund.deptStats[dept]) summary.loeiFund.deptStats[dept] = { allocated: 0, spent: 0 };
                    summary.loeiFund.deptStats[dept].allocated += allocAmt;
                    summary.loeiFund.deptStats[dept].spent += spentAmt;
                }

                // 3. หมวด HRD 1 & 2
                if (typeVal === 'เงินนอก สป.' && sourceVal === 'เงินบำรุง สสจ.เลย' && budgetCodeVal === 'เงินบำรุง สสจ.เลย (รพ.เลย)') {
                    if (actCodeVal === 'เงินบำรุง สสจ.เลย (หมวด HRD)') {
                        summary.hrd1.approved += approveAmt;
                        summary.hrd1.allocated += allocAmt;
                        summary.hrd1.spent += spentAmt;
                        summary.hrd1.balance += balAmt;
                    } else if (actCodeVal === 'HRD อบรมหลักสูตรต่างๆ') {
                        summary.hrd2.approved += approveAmt;
                        summary.hrd2.allocated += allocAmt;
                        summary.hrd2.spent += spentAmt;
                        summary.hrd2.balance += balAmt;
                    }
                }
            }
        }
        return summary;
    } catch (e) { return { error: e.message }; }
}

// ==========================================
// 5. SEARCH & YEARLY
// ==========================================
function searchActionPlan(dept, budgetType, quarter, month) {
    var result = getYearlyPlanData(dept, budgetType, quarter, month);
    return { summary: { count: result.summary.projects, approved: result.summary.approved, allocated: result.summary.allocated }, list: result.list };
}

function getYearlyPlanData(deptFilter, typeFilter, quarterFilter, monthFilter) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) return { summary: { projects: 0 }, list: [] };

        var map = getColumnMap(sheet); // 🌟 ใช้ Map แทนตัวเลข
        var data = sheet.getDataRange().getValues();

        var I_MONTHS = [
            map[COL_NAME.M_OCT], map[COL_NAME.M_NOV], map[COL_NAME.M_DEC],
            map[COL_NAME.M_JAN], map[COL_NAME.M_FEB], map[COL_NAME.M_MAR],
            map[COL_NAME.M_APR], map[COL_NAME.M_MAY], map[COL_NAME.M_JUN],
            map[COL_NAME.M_JUL], map[COL_NAME.M_AUG], map[COL_NAME.M_SEP]
        ];

        var quarters = { 'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11] };
        var summary = { projects: 0, approved: 0, allocated: 0, spent: 0 };
        var list = [];
        var parseNum = (val) => { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var rowDept = String(row[map[COL_NAME.DEPT]] || "").trim();
            var passDept = (deptFilter === "" || deptFilter === null || rowDept === deptFilter);

            var typeVal = String(row[map[COL_NAME.TYPE]] || "").trim();
            var isNonMoph = (typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 || typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1);
            var isMoph = !isNonMoph;
            var passType = true;
            if (typeFilter === 'MOPH') passType = isMoph;
            else if (typeFilter === 'NONMOPH') passType = isNonMoph;

            var timeline = I_MONTHS.map(idx => (idx !== undefined && String(row[idx]).trim() !== '') ? 1 : 0);
            var passTime = true;
            if (quarterFilter && quarters[quarterFilter]) {
                if (!quarters[quarterFilter].some(mIdx => timeline[mIdx] === 1)) passTime = false;
            }
            if (monthFilter) {
                if (timeline[parseInt(monthFilter)] !== 1) passTime = false;
            }

            if (passDept && passType && passTime) {
                var actName = String(row[map[COL_NAME.ACT]] || "");
                if (row[map[COL_NAME.SUB]]) actName += " (" + row[map[COL_NAME.SUB]] + ")";
                var alloc = parseNum(row[map[COL_NAME.ALLOC]]);
                var spent = parseNum(row[map[COL_NAME.SPENT]]);

                summary.projects++;
                summary.allocated += alloc;
                summary.spent += spent;

                list.push({
                    order: row[map[COL_NAME.ORDER]],
                    dept: rowDept,
                    project: row[map[COL_NAME.PROJ]],
                    activity: actName,
                    type: row[map[COL_NAME.TYPE]],
                    budgetSource: row[map[COL_NAME.SOURCE]],
                    timeline: timeline,
                    allocated: alloc,
                    spent: spent,
                    balance: alloc - spent
                });
            }
        }
        return { summary: summary, list: list };
    } catch (e) { return { error: e.message }; }
}

// ==========================================
// 6. SAVE & UPDATE (Transaction)
// ==========================================
function saveTransaction(payload) {
    try {
        var lock = LockService.getScriptLock();
        lock.waitLock(10000);

        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetT = ss.getSheetByName('t_actionplan');
        var sheetM = ss.getSheetByName('m_actionplan');

        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);

        // 🔍 1. ค้นหาแถวของโครงการเป้าหมายใน m_actionplan
        var targetRow = -1;
        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(payload.projectId).trim()) {
                targetRow = i;
                break;
            }
        }

        if (targetRow === -1) {
            return { success: false, message: 'ไม่พบรหัสโครงการนี้ในฐานข้อมูลหลัก (ID: ' + payload.projectId + ')' };
        }

        // 📊 2. ดึงค่ายอดเงินปัจจุบัน
        var spentIndex = map[COL_NAME.SPENT];
        var balIndex = map[COL_NAME.BAL];

        var currentSpent = Number(mData[targetRow][spentIndex]) || 0;
        var currentBal = Number(mData[targetRow][balIndex]) || 0;
        var amount = Number(payload.amount) || 0;

        // ดักจับความปลอดภัย ห้ามเบิกเกินงบ
        if (amount > currentBal) {
            return { success: false, message: 'ยอดเงินคงเหลือไม่เพียงพอสำหรับการเบิกจ่าย' };
        }

        // 📋 3. ดึงข้อมูล Year, Cat, Plan จากตาราง Master มาเติมให้สมบูรณ์
        var mYear = map[COL_NAME.YEAR] !== undefined ? mData[targetRow][map[COL_NAME.YEAR]] : "";
        var mCat = map[COL_NAME.CAT] !== undefined ? mData[targetRow][map[COL_NAME.CAT]] : "";
        var mPlan = map[COL_NAME.PLAN] !== undefined ? mData[targetRow][map[COL_NAME.PLAN]] : "";

        // 🎯 4. สร้าง Array แถวใหม่สำหรับตารางประวัติ (t_actionplan) โยนลงช่องให้เป๊ะ!
        var newRowT = [];
        newRowT.length = 30; // จองพื้นที่ว่างไว้ 30 คอลัมน์กันเหนียว
        newRowT.fill('');

        newRowT[0] = new Date();                     // A: ประทับเวลา
        newRowT[1] = payload.projectId;              // B: รหัสโครงการ
        newRowT[2] = mYear;                          // C: ปีงบประมาณ
        newRowT[3] = mCat;                           // D: หมวด
        newRowT[4] = payload.order;                  // E: ลำดับ
        newRowT[5] = payload.dept;                   // F: กลุ่มงาน
        newRowT[6] = mPlan;                          // G: แผนงาน
        newRowT[7] = payload.proj;                   // H: โครงการ
        newRowT[8] = payload.act;                    // I: กิจกรรมหลัก
        newRowT[9] = payload.sub;                    // J: กิจกรรมย่อย
        newRowT[10] = payload.type;                  // K: ประเภทงบ
        newRowT[11] = payload.source;                // L: แหล่งงบ

        newRowT[15] = amount;                        // P: ยอดเบิกจ่าย
        newRowT[16] = currentBal - amount;           // Q: คงเหลือ
        newRowT[17] = formatToStorageDate(payload.txDate); // R: วันที่เบิกจ่าย (แปลงเป็น format ชีต)
        newRowT[18] = payload.expenseType || '';     // S: หมวดรายจ่าย
        newRowT[19] = payload.desc || payload.remark || ''; // T: รายละเอียด/หมายเหตุ

        newRowT[25] = formatToStorageDate(payload.startDate) || ''; // Z: เริ่มดำเนินการ (สำหรับ Gantt)
        newRowT[26] = formatToStorageDate(payload.endDate) || '';   // AA: สิ้นสุดดำเนินการ (สำหรับ Gantt)

        // 🚀 5. บันทึกประวัติก่อนเป็นลำดับแรก (Atomic Transaction Step 1)
        sheetT.appendRow(newRowT);
        SpreadsheetApp.flush();

        // 💾 6. อัปเดตตาราง Master (บวกยอดเบิกจ่าย, ลบยอดคงเหลือ) หลังจากประวัติบันทึกสำเร็จ (Atomic Transaction Step 2)
        sheetM.getRange(targetRow + 1, spentIndex + 1).setValue(currentSpent + amount);
        sheetM.getRange(targetRow + 1, balIndex + 1).setValue(currentBal - amount);
        SpreadsheetApp.flush();
        
        lock.releaseLock();

        return { success: true, count: 1 };
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message };
    }
}

// ==========================================
// 📌 ฟังก์ชันลบประวัติเบิกจ่ายและดึงยอดเงินคืน (Delete Transaction)
// ==========================================
function deleteTransaction(rowIndex, projectId, amount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetT = ss.getSheetByName('t_actionplan');
        var sheetM = ss.getSheetByName('m_actionplan');

        var tData = sheetT.getDataRange().getValues();
        var targetRowT = -1;
        
        var paramProjectId = String(projectId).trim();
        var paramAmount = Number(amount) || 0;

        // 1. ค้นหาแถวที่ต้องการลบในประวัติ จากล่างขึ้นบนเพื่อความแม่นยำ ป้องกันแถวเคลื่อน
        for (var j = tData.length - 1; j >= 1; j--) {
            var recProj = String(tData[j][1]).trim(); // B: รหัสโครงการ
            var recAmt = Number(tData[j][15]) || 0;   // P: ยอดเบิกจ่าย

            if (recProj === paramProjectId && Math.abs(recAmt - paramAmount) < 0.01) {
                targetRowT = j + 1;
                break;
            }
        }

        if (targetRowT === -1) {
            return { success: false, message: 'ข้อมูลไม่ตรงกัน ไม่สามารถลบได้ (ไม่พบประวัติ หรือมีการเปลี่ยนแปลงข้อมูลก่อนหน้า โปรดรีเฟรช)' };
        }

        // 2. หายอดเงินของโครงการนั้น เพื่อดึงเงินคืนกลับเข้ากระเป๋า
        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === paramProjectId) {
                targetRowM = i + 1;
                break;
            }
        }

        if (targetRowM > -1) {
            var spentIdx = map[COL_NAME.SPENT] + 1;
            var balIdx = map[COL_NAME.BAL] + 1;

            var currentSpent = Number(sheetM.getRange(targetRowM, spentIdx).getValue()) || 0;
            var currentBal = Number(sheetM.getRange(targetRowM, balIdx).getValue()) || 0;

            var newSpent = currentSpent - paramAmount;
            if (newSpent < 0) newSpent = 0; // กันติดลบ

            // 🎯 คืนยอดลง Master
            sheetM.getRange(targetRowM, spentIdx).setValue(newSpent);
            sheetM.getRange(targetRowM, balIdx).setValue(currentBal + paramAmount);
        }

        // 3. ลบแถวประวัติทิ้งเมื่อการค้นหาตรงเงื่อนไข
        sheetT.deleteRow(targetRowT);

        SpreadsheetApp.flush();
        return { success: true };
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message };
    } finally {
        lock.releaseLock();
    }
}

// ==========================================
// 📌 แก้ไขรายการเบิกจ่าย (Dynamic Column)
// ==========================================
function editTransaction(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName('t_actionplan');
        var mSheet = ss.getSheetByName('m_actionplan');
        if (!tSheet || !mSheet) return { status: 'error', message: 'ไม่พบ Sheet ข้อมูล' };

        var tRow = parseInt(form.rowId);
        var checkID = tSheet.getRange(tRow, 2).getValue();
        if (String(checkID) != String(form.projectId)) { return { status: 'error', message: 'ข้อมูลไม่ตรงกัน โปรดรีเฟรช' }; }

        var map = getColumnMap(mSheet); // 🌟 เรียกใช้เรดาร์ค้นหาคอลัมน์
        var mData = mSheet.getDataRange().getValues();
        var mRowIndex = -1;
        var targetId = String(form.projectId).trim();

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === targetId) { mRowIndex = i + 1; break; }
        }

        if (mRowIndex !== -1) {
            // 🌟 ชี้เป้าไปที่คอลัมน์ 'เบิกจ่าย' แบบ Dynamic
            var cellSpent = mSheet.getRange(mRowIndex, map[COL_NAME.SPENT] + 1);
            var currentSpent = parseFloat(String(cellSpent.getValue()).replace(/,/g, '')) || 0;

            var oldVal = parseFloat(form.oldAmount) || 0;
            var newVal = parseFloat(form.newAmount) || 0;

            var newSpentTotal = currentSpent - oldVal + newVal;
            cellSpent.setValue(newSpentTotal);
        }

        tSheet.getRange(tRow, 16).setValue(form.newAmount);
        tSheet.getRange(tRow, 18).setValue(formatToStorageDate(form.date));
        tSheet.getRange(tRow, 20).setValue(form.desc);
        tSheet.getRange(tRow, 24).setValue(form.reason);
        return { status: 'success', message: 'แก้ไขเรียบร้อย' };
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

// ==========================================
// 📌 1. บันทึกเงินยืม (ปลอดภัยกับสูตร 100% + ป้องกันข้อมูลเบี้ยว 100%)
// ==========================================
function saveLoan(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var mSheet = ss.getSheetByName('m_actionplan');
        var tSheet = ss.getSheetByName('t_loan');
        if (!mSheet) return { status: 'error', message: 'ไม่พบ Sheet: m_actionplan' };
        if (!tSheet) return { status: 'error', message: 'ไม่พบ Sheet: t_loan' };

        var map = getColumnMap(mSheet);
        var mData = mSheet.getDataRange().getValues();
        var rowIndex = -1;
        var targetId = String(form.projectId).trim();

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === targetId) {
                rowIndex = i + 1;
                break;
            }
        }
        if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ใน Master' };

        var r = mData[rowIndex - 1];

        // 🎯 อ่านค่ามาตรวจสอบเฉยๆ ไม่เขียนทับช่องสูตร
        var cellMasterLoan = mSheet.getRange(rowIndex, map[COL_NAME.LOAN] + 1);
        var cellMasterBal = mSheet.getRange(rowIndex, map[COL_NAME.BAL] + 1);

        var currentMasterLoan = parseFloat(String(cellMasterLoan.getValue()).replace(/,/g, '')) || 0;
        var currentMasterBal = parseFloat(String(cellMasterBal.getValue()).replace(/,/g, '')) || 0;
        var loanAmt = parseFloat(String(form.amount).replace(/,/g, '')) || 0;

        if (loanAmt > currentMasterBal) {
            return { status: 'error', message: 'ยอดคงเหลือไม่เพียงพอสำหรับทำรายการเงินยืม' };
        }

        // 🎯 บันทึกแค่ช่อง "เงินยืม" ใน m_actionplan เท่านั้น!
        cellMasterLoan.setValue(currentMasterLoan + loanAmt);

        var sDate = form.startDate ? formatToStorageDate(form.startDate) : "";
        var eDate = form.endDate ? formatToStorageDate(form.endDate) : "";

        // 🎯 จัดเรียง Array ให้ตรงกับโครงสร้างชีต t_loan 27 คอลัมน์เป๊ะๆ
        tSheet.appendRow([
            new Date(),                                  // 1. ประทับเวลา
            r[map[COL_NAME.ID]],                         // 2. รหัสโครงการ
            r[map[COL_NAME.YEAR]],                       // 3. ปีงบประมาณ
            r[map[COL_NAME.CAT]],                        // 4. หมวด
            r[map[COL_NAME.ORDER]],                      // 5. ลำดับโครงการ
            r[map[COL_NAME.DEPT]],                       // 6. กลุ่มงาน/งาน
            r[map[COL_NAME.PLAN]],                       // 7. แผนงาน
            r[map[COL_NAME.PROJ]],                       // 8. โครงการ
            r[map[COL_NAME.ACT]],                        // 9. กิจกรรมหลัก
            r[map[COL_NAME.SUB]],                        // 10. กิจกรรมย่อย
            r[map[COL_NAME.TYPE]],                       // 11. ประเภทงบ
            r[map[COL_NAME.SOURCE]],                     // 12. แหล่งงบประมาณ
            r[map[COL_NAME.BUDGET_CODE]],                // 13. รหัสงบประมาณ
            r[map[COL_NAME.ACT_CODE]],                   // 14. รหัสกิจกรรม
            r[map[COL_NAME.ALLOC]],                      // 15. จัดสรร
            loanAmt,                                     // 16. เงินยืม
            formatToStorageDate(form.loanDate),          // 17. วันที่ยืมเงิน
            form.loanType,                               // 18. ประเภทเงินยืม
            form.desc,                                   // 19. รายละเอียดการยืมเงิน
            "ยังไม่ดำเนินการ",                              // 20. สถานะการเบิกจ่าย
            0,                                           // 21. จำนวนเบิกจ่าย
            loanAmt,                                     // 22. คงเหลือ
            "",                                          // 23. วันที่เบิกจ่าย
            "",                                          // 24. ระยะเวลาที่ยืม
            "",                                          // 25. หมายเหตุ
            sDate,                                       // 26. เริ่มดำเนินการ
            eDate                                        // 27. สิ้นสุดดำเนินการ
        ]);

        SpreadsheetApp.flush();
        return { status: 'success', message: 'บันทึกเงินยืมเรียบร้อย' };
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

function updateLoanRepayment(data) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName('t_loan');
        if (!tSheet) throw new Error("ไม่พบ Sheet: t_loan");
        var tData = tSheet.getDataRange().getValues();
        var tDisplayData = tSheet.getDataRange().getDisplayValues();
        var tRowIndex = -1; var projectId = ""; var loanAmount = 0;
        var targetDate = new Date(data.timestamp);

        for (var i = 1; i < tData.length; i++) {
            var cellValue = tData[i][0];
            var cellDisplayValue = tDisplayData[i][0];
            var isMatch = false;
            if (String(cellDisplayValue).trim() === String(data.timestamp).trim() || String(cellValue).trim() === String(data.timestamp).trim()) {
                isMatch = true;
            } else {
                var cellDate = new Date(cellValue);
                if (!isNaN(cellDate.getTime()) && !isNaN(targetDate.getTime())) {
                    if (Math.abs(cellDate.getTime() - targetDate.getTime()) < 60000) { isMatch = true; }
                }
            }
            if (isMatch) { tRowIndex = i + 1; projectId = tData[i][1]; loanAmount = parseFloat(tData[i][15] || 0); break; }
        }
        if (tRowIndex === -1) throw new Error("ไม่พบรายการกู้ยืม (Timestamp ไม่ตรง)");

        var currentPaid = parseFloat(tData[tRowIndex - 1][20] || 0);
        var payAmount = parseFloat(data.paidAmount);
        var newPaid = currentPaid + payAmount;
        var newBalance = loanAmount - newPaid;
        var status = (newBalance <= 0.01) ? "คืนครบ" : "คืนบางส่วน";
        if (newBalance < 0) newBalance = 0;

        tSheet.getRange(tRowIndex, 20).setValue(status);
        tSheet.getRange(tRowIndex, 21).setValue(newPaid);
        tSheet.getRange(tRowIndex, 22).setValue(newBalance);
        tSheet.getRange(tRowIndex, 23).setValue(formatToStorageDate(data.payDate));

        if (projectId) {
            var mSheet = ss.getSheetByName('m_actionplan');
            if (mSheet) {
                var mData = mSheet.getDataRange().getValues();
                var mRowIndex = -1;
                for (var j = 1; j < mData.length; j++) {
                    if (String(mData[j][0]) == String(projectId)) { mRowIndex = j + 1; break; }
                }
                if (mRowIndex !== -1) {
                    var colAlloc = 17; var colSpent = 18; var colBalance = 20;
                    var cellAlloc = mSheet.getRange(mRowIndex, colAlloc);
                    var allocated = parseFloat(cellAlloc.getValue()) || 0;
                    var cellSpent = mSheet.getRange(mRowIndex, colSpent);
                    var currentSpent = parseFloat(cellSpent.getValue()) || 0;
                    var updatedSpent = currentSpent + payAmount;
                    var updatedBalance = allocated - updatedSpent;
                    cellSpent.setValue(updatedSpent);
                    mSheet.getRange(mRowIndex, colBalance).setValue(updatedBalance);
                }
            }
        }
        return { status: 'success', message: 'บันทึกคืนเงินและตัดงบประมาณเรียบร้อย' };
    } catch (e) {
        return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.toString() };
    } finally { lock.releaseLock(); }
}

// ==========================================
// 8. HISTORY GETTERS
// ==========================================
function getTransactionHistory() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName('t_actionplan');
        if (!tSheet) return [];
        var mSheet = ss.getSheetByName('m_actionplan');
        var masterMap = {};
        
        if (mSheet) {
            var map = getColumnMap(mSheet); 
            var mData = mSheet.getDataRange().getValues();
            for (var i = 1; i < mData.length; i++) {
                var pid = String(mData[i][map[COL_NAME.ID]]).trim();
                if (pid) {
                    masterMap[pid] = {
                        allocated: parseFloat(String(mData[i][map[COL_NAME.ALLOC]]).replace(/,/g, '')) || 0,
                        spent: parseFloat(String(mData[i][map[COL_NAME.SPENT]]).replace(/,/g, '')) || 0,
                        loan: parseFloat(String(mData[i][map[COL_NAME.LOAN]]).replace(/,/g, '')) || 0,
                        balance: parseFloat(String(mData[i][map[COL_NAME.BAL]]).replace(/,/g, '')) || 0
                    };
                }
            }
        }
        var data = tSheet.getDataRange().getDisplayValues();
        if (data.length < 2) return [];
        var headers = data[0];
        
        var colProjId = headers.indexOf(COL_NAME.ID) > -1 ? headers.indexOf(COL_NAME.ID) : 1;
        var colOrder = headers.indexOf(COL_NAME.ORDER) > -1 ? headers.indexOf(COL_NAME.ORDER) : 4;
        var colProj = headers.indexOf(COL_NAME.PROJ) > -1 ? headers.indexOf(COL_NAME.PROJ) : 7;
        var colAct = headers.indexOf(COL_NAME.ACT) > -1 ? headers.indexOf(COL_NAME.ACT) : 8;
        var colSub = headers.indexOf(COL_NAME.SUB) > -1 ? headers.indexOf(COL_NAME.SUB) : 9;
        var colSource = headers.indexOf(COL_NAME.SOURCE) > -1 ? headers.indexOf(COL_NAME.SOURCE) : 11;
        var colAmount = headers.indexOf('ยอดเบิกจ่าย') > -1 ? headers.indexOf('ยอดเบิกจ่าย') : 15;
        var colDate = headers.indexOf('วันที่เบิกจ่าย') > -1 ? headers.indexOf('วันที่เบิกจ่าย') : 17;
        var colExpType = headers.indexOf('หมวดรายจ่าย') > -1 ? headers.indexOf('หมวดรายจ่าย') : 18;
        var colDesc = headers.indexOf('รายละเอียด/หมายเหตุ') > -1 ? headers.indexOf('รายละเอียด/หมายเหตุ') : 19;

        var result = [];
        var parseAmount = function (v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };

        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row || (!row[0] && !row[colProjId])) continue;
            var projId = String(row[colProjId]).trim();
            var masterInfo = masterMap[projId] || { allocated: 0, spent: 0, loan: 0, balance: 0 };
            var item = {
                rowId: i + 1, 
                order: row[colOrder],
                project: row[colProj],
                activity: row[colAct],
                subActivity: row[colSub],
                amount: parseAmount(row[colAmount]),
                date: formatToThaiUI(row[colDate]), 
                type: row[colExpType],
                source: row[colSource],
                desc: row[colDesc],
                id: projId,
                masterAllocated: masterInfo.allocated,
                masterSpent: masterInfo.spent,
                masterLoan: masterInfo.loan,
                masterBalance: masterInfo.balance
            };
            if (item.amount > 0 || item.order) result.push(item);
            if (result.length >= 100) break;
        }
        return result;
    } catch (e) { return []; }
}

// --- ดึงประวัติเงินยืม ---
function getLoanHistory() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    if (!tSheet) return [];
    var projectMap = {};
    try {
        var mSheet = ss.getSheetByName('m_actionplan');
        if (mSheet) {
            var map = getColumnMap(mSheet); 
            var mData = mSheet.getDataRange().getDisplayValues();
            for (var i = 1; i < mData.length; i++) {
                var pid = String(mData[i][map[COL_NAME.ID]]).trim();
                if (pid) { projectMap[pid] = { type: mData[i][map[COL_NAME.TYPE]] || "-", source: mData[i][map[COL_NAME.SOURCE]] || "-" }; }
            }
        }
    } catch (e) { }

    var tData = tSheet.getDataRange().getDisplayValues();
    if (tData.length < 2) return [];
    var headers = tData[0];
    
    var colProjId = headers.indexOf(COL_NAME.ID) > -1 ? headers.indexOf(COL_NAME.ID) : 1;
    var colDept = headers.indexOf(COL_NAME.DEPT) > -1 ? headers.indexOf(COL_NAME.DEPT) : 5;
    var colOrder = headers.indexOf(COL_NAME.ORDER) > -1 ? headers.indexOf(COL_NAME.ORDER) : 4;
    var colProj = headers.indexOf(COL_NAME.PROJ) > -1 ? headers.indexOf(COL_NAME.PROJ) : 7;
    var colAct = headers.indexOf(COL_NAME.ACT) > -1 ? headers.indexOf(COL_NAME.ACT) : 8;
    var colSub = headers.indexOf(COL_NAME.SUB) > -1 ? headers.indexOf(COL_NAME.SUB) : 9;
    var colAmount = headers.indexOf('เงินยืม') > -1 ? headers.indexOf('เงินยืม') : 15;
    var colDate = headers.indexOf('วันที่ยืมเงิน') > -1 ? headers.indexOf('วันที่ยืมเงิน') : 16;
    var colType = headers.indexOf('ประเภทเงินยืม') > -1 ? headers.indexOf('ประเภทเงินยืม') : 17;
    var colDesc = headers.indexOf('รายละเอียดการยืมเงิน') > -1 ? headers.indexOf('รายละเอียดการยืมเงิน') : 18;
    var colStatus = headers.indexOf('สถานะการเบิกจ่าย') > -1 ? headers.indexOf('สถานะการเบิกจ่าย') : 19;
    var colPaid = headers.indexOf('จำนวนเบิกจ่าย') > -1 ? headers.indexOf('จำนวนเบิกจ่าย') : 20;
    var colBalance = headers.indexOf('คงเหลือ') > -1 ? headers.indexOf('คงเหลือ') : 21;

    var result = [];
    var parseNum = function (val) { return parseFloat(String(val).replace(/,/g, '')) || 0; };

    for (var i = tData.length - 1; i >= 1; i--) {
        try {
            var row = tData[i];
            if (!row[0] && !row[colProjId]) continue;
            var pid = String(row[colProjId] || "").trim();
            var meta = projectMap[pid] || { type: '-', source: '-' };
            result.push({
                id: row[0],
                timestamp: row[0],
                projectId: pid, 
                rowId: i + 1,      
                project: row[colProj],
                activity: row[colAct],
                subActivity: row[colSub],
                amount: parseNum(row[colAmount]),
                date: formatToThaiUI(row[colDate]),
                status: row[colStatus],
                paid: parseNum(row[colPaid]),
                balance: parseNum(row[colBalance]),
                order: row[colOrder],
                type: row[colType],
                desc: row[colDesc],
                budgetType: meta.type,
                budgetSource: meta.source,
                dept: row[colDept]
            });
        } catch (e) { }
    }
    return result;
}

// --- ฟังก์ชันดึงประวัติรวม (Generic) ---
function getHistory(sheetName) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var tSheet = ss.getSheetByName(sheetName);
        if (!tSheet) return [];
        var data = tSheet.getDataRange().getDisplayValues();
        if (data.length < 2) return [];
        var result = [];
        var parseAmount = function (v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };

        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row || (!row[0] && !row[1])) continue;
            var item = {};
            if (sheetName === 't_actionplan') {
                item = {
                    rowId: i + 1, order: row[4], project: row[7], activity: row[8], subActivity: row[9], amount: parseAmount(row[15]),
                    date: formatToThaiUI(row[17]), // ✅ อัปเดตวันที่
                    type: row[18], source: row[11], desc: row[19], id: row[1]
                };
            } else {
                item = { timestamp: row[0], order: row[4], project: row[7], amount: parseAmount(row[15]), date: formatToThaiUI(row[16]), status: row[19] };
            }
            if (item.amount > 0 || item.order) result.push(item);
            if (result.length >= 50) break;
        }
        return result;
    } catch (e) { return []; }
}

function searchLoanHistory(criteria) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tSheet = ss.getSheetByName('t_loan');
    var mSheet = ss.getSheetByName('m_actionplan');
    var projectMap = {};
    if (mSheet) {
        var map = getColumnMap(mSheet); // 🌟 FIX: ใช้ Dynamic Column Map แทน hardcoded index
        var mData = mSheet.getDataRange().getDisplayValues();
        for (var i = 1; i < mData.length; i++) {
            var pid = String(mData[i][map[COL_NAME.ID]]).trim();
            projectMap[pid] = { type: mData[i][map[COL_NAME.TYPE]], source: mData[i][map[COL_NAME.SOURCE]] };
        }
    }
    var data = tSheet.getDataRange().getDisplayValues();
    var result = [];
    var parseNum = function (v) { return parseFloat(String(v).replace(/,/g, '')) || 0; };
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var match = true;
        if (criteria.order && String(row[4]) != String(criteria.order)) match = false;
        if (match && criteria.project && String(row[7]).indexOf(criteria.project) === -1) match = false;
        if (match) {
            var pid = String(row[1]).trim();
            var meta = projectMap[pid] || { type: '-', source: '-' };
            result.push({
                id: row[0], timestamp: row[0], project: row[7], activity: row[8], subActivity: row[9], amount: parseNum(row[15]),
                date: row[16], status: row[19], paid: parseNum(row[20]), balance: parseNum(row[21]), order: row[4],
                type: row[17], desc: row[18], budgetType: meta.type, budgetSource: meta.source, dept: row[2]
            });
        }
    }
    return result;
}

// ==========================================
// 📌 บันทึกประวัติจัดสรรงบ (Dynamic Column & Anti-Shift)
// ==========================================
function saveAllocation(form) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var mSheet = ss.getSheetByName('m_actionplan');
        var tAllocSheet = ss.getSheetByName('t_allocate');
        if (!mSheet) return { status: 'error', message: 'ไม่พบ Sheet Master' };
        if (!tAllocSheet) return { status: 'error', message: 'ไม่พบ Sheet t_allocate' };

        var map = getColumnMap(mSheet); // 🌟 เรียกใช้เรดาร์ค้นหาคอลัมน์
        var mData = mSheet.getDataRange().getValues();
        var rowIndex = -1;
        var targetId = String(form.id).trim();
        var targetAct = form.fullData ? String(form.fullData.activity || "").trim() : "";
        var targetSub = form.fullData ? String(form.fullData.subActivity || "").trim() : "";

        for (var i = 1; i < mData.length; i++) {
            var rowId = String(mData[i][map[COL_NAME.ID]]).trim();
            var rowAct = String(mData[i][map[COL_NAME.ACT]]).trim();
            var rowSub = String(mData[i][map[COL_NAME.SUB]]).trim();
            if (rowId === targetId && rowAct === targetAct && rowSub === targetSub) { rowIndex = i + 1; break; }
        }

        if (rowIndex === -1) {
            for (var i = 1; i < mData.length; i++) {
                if (String(mData[i][map[COL_NAME.ID]]).trim() === targetId) { rowIndex = i + 1; break; }
            }
            if (rowIndex === -1) return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ใน Master (ID: ' + form.id + ')' };
        }

        // 🌟 ชี้เป้าไปที่คอลัมน์ 'จัดสรร' อย่างแม่นยำ
        var cellAlloc = mSheet.getRange(rowIndex, map[COL_NAME.ALLOC] + 1);
        var rawVal = String(cellAlloc.getValue());
        var currentTotal = parseFloat(rawVal.replace(/,/g, '')) || 0;
        var addAmount = parseFloat(String(form.currentAlloc).replace(/,/g, '')) || 0;
        var newTotal = currentTotal + addAmount;

        // 🎯 บันทึกเฉพาะช่องจัดสรร
        cellAlloc.setValue(newTotal);
        SpreadsheetApp.flush();

        var r = mData[rowIndex - 1];

        // 🎯 จัดเรียงคอลัมน์ประวัติใหม่ (เพิ่ม BUDGET_CODE อุดรอยรั่วข้อมูลเบี้ยว)
        var logRow = [
            new Date(),
            r[map[COL_NAME.ID]],
            r[map[COL_NAME.YEAR]],
            r[map[COL_NAME.CAT]],
            r[map[COL_NAME.ORDER]],
            r[map[COL_NAME.DEPT]],
            r[map[COL_NAME.PLAN]],
            r[map[COL_NAME.PROJ]],
            r[map[COL_NAME.ACT]],
            r[map[COL_NAME.SUB]],
            r[map[COL_NAME.TYPE]],
            r[map[COL_NAME.SOURCE]],
            r[map[COL_NAME.BUDGET_CODE]],   // 🌟 เติมรหัสงบประมาณที่หายไป
            r[map[COL_NAME.ACT_CODE]],
            r[map[COL_NAME.APPROVE]],
            newTotal,                       // ยอดจัดสรรสะสมใหม่
            addAmount,                      // ยอดเงินที่เพิ่งได้รับจัดสรร
            formatToStorageDate(form.date), // วันที่เอกสาร
            form.letterNo,                  // เลขที่หนังสือ
            form.remark                     // หมายเหตุ
        ];
        tAllocSheet.appendRow(logRow);

        return { status: 'success', message: 'บันทึกจัดสรรสำเร็จ (ยอดใหม่: ' + newTotal.toLocaleString() + ')' };
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
    } finally { lock.releaseLock(); }
}

// --- ดึงประวัติจัดสรรงบประมาณ ---
function getAllocationHistory() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('t_allocate');
        if (!sheet) return [];
        var data = sheet.getDataRange().getDisplayValues();
        if (data.length < 2) return [];
        var result = [];
        var parseNum = (v) => parseFloat(String(v).replace(/,/g, '')) || 0;

        for (var i = data.length - 1; i >= 1; i--) {
            var row = data[i];
            if (!row[1]) continue;
            result.push({
                id: row[1],
                order: row[4],
                dept: row[5],
                project: row[7],
                activity: row[8],
                subActivity: row[9],
                type: row[10],
                source: row[11],
                accumulatedAlloc: parseNum(row[14]),
                currentAlloc: parseNum(row[15]),
                date: formatToThaiUI(row[16]), // ✅ อัปเดตวันที่
                letterNo: row[17]
            });
            if (result.length >= 100) break;
        }
        return result;
    } catch (e) { return []; }
}

function formatDateThai(dateStr) {
    if (!dateStr) return "-";
    try {
        var d = new Date(dateStr);
        if (isNaN(d.getTime())) return dateStr;
        var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
        return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
    } catch (e) { return dateStr; }
}

function getDeptDetails(deptName, quarterFilter, monthFilter) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        if (!sheet) return { error: "ไม่พบ Sheet m_actionplan" };

        var map = getColumnMap(sheet); // 🌟 ใช้ Map แทนตัวเลข
        var data = sheet.getDataRange().getValues();

        var I_MONTHS = [
            map[COL_NAME.M_OCT], map[COL_NAME.M_NOV], map[COL_NAME.M_DEC],
            map[COL_NAME.M_JAN], map[COL_NAME.M_FEB], map[COL_NAME.M_MAR],
            map[COL_NAME.M_APR], map[COL_NAME.M_MAY], map[COL_NAME.M_JUN],
            map[COL_NAME.M_JUL], map[COL_NAME.M_AUG], map[COL_NAME.M_SEP]
        ];

        var quarters = { 'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11] };
        var parseNum = function (val) { var v = parseFloat(String(val).replace(/,/g, '')); return isNaN(v) ? 0 : v; };
        var cleanName = function (str) { return String(str).replace(/[\s\/\-_]+/g, "").trim(); };

        var projectsAll = [], projectsMoph = [], projectsNon = [];
        var sumAll = { allocated: 0, spent: 0 }, sumMoph = { allocated: 0, spent: 0 }, sumNon = { allocated: 0, spent: 0 };
        var targetClean = cleanName(deptName);

        for (var i = 1; i < data.length; i++) {
            var row = data[i];
            var rowDeptRaw = String(row[map[COL_NAME.DEPT]] || "");
            if (cleanName(rowDeptRaw).indexOf(targetClean) === -1 && targetClean.indexOf(cleanName(rowDeptRaw)) === -1) continue;

            var timeline = I_MONTHS.map(function (idx) { return (idx !== undefined && String(row[idx] || "").trim() !== '') ? 1 : 0; });
            var passTime = true;
            if (quarterFilter && quarters[quarterFilter]) { if (!quarters[quarterFilter].some(function (mIdx) { return timeline[mIdx] === 1; })) passTime = false; }
            if (monthFilter && String(monthFilter) !== "") { if (timeline[parseInt(monthFilter)] !== 1) passTime = false; }

            if (passTime) {
                var approve = parseNum(row[map[COL_NAME.APPROVE]]);
                var alloc = parseNum(row[map[COL_NAME.ALLOC]]);
                var spent = parseNum(row[map[COL_NAME.SPENT]]);
                var typeVal = String(row[map[COL_NAME.TYPE]] || "").trim();

                var projObj = {
                    project: String(row[map[COL_NAME.PROJ]] || "-"),
                    activity: String(row[map[COL_NAME.ACT]] || "-"),
                    approved: approve, allocated: alloc, spent: spent, balance: alloc - spent,
                    progress: (alloc > 0) ? (spent / alloc * 100) : 0
                };

                projectsAll.push(projObj);
                sumAll.allocated += alloc; sumAll.spent += spent;

                var isNonMoph = (typeVal.indexOf('เงินนอก') > -1 || typeVal.indexOf('เงินบำรุง') > -1 || typeVal.indexOf('บริจาค') > -1 || typeVal.toUpperCase().indexOf('NON') > -1);

                if (!isNonMoph) {
                    projectsMoph.push(projObj);
                    sumMoph.allocated += alloc; sumMoph.spent += spent;
                } else {
                    projectsNon.push(projObj);
                    sumNon.allocated += alloc; sumNon.spent += spent;
                }
            }
        }

        var sortFn = function (a, b) { return b.progress - a.progress; };
        projectsAll.sort(sortFn); projectsMoph.sort(sortFn); projectsNon.sort(sortFn);

        return {
            projectsAll: projectsAll, projectsMoph: projectsMoph, projectsNon: projectsNon,
            sumAll: sumAll, sumMoph: sumMoph, sumNon: sumNon,
            deptName: deptName
        };
    } catch (e) { return { error: "Server Error: " + e.message }; }
}

// ==========================================
// 📌 ฟังก์ชันทดสอบความถูกต้องของชื่อคอลัมน์ (รันเพื่อเช็ค)
// ==========================================
function testDynamicMapping() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('m_actionplan');
    var map = getColumnMap(sheet);
    var keysToTest = Object.keys(COL_NAME);

    var errorCount = 0;
    Logger.log("=== 🔍 เริ่มตรวจสอบ Dynamic Mapping ===");

    for (var i = 0; i < keysToTest.length; i++) {
        var key = keysToTest[i];
        var expectedName = COL_NAME[key];
        var foundIndex = map[expectedName];

        if (foundIndex !== undefined) {
            Logger.log("✅ พบคอลัมน์ [" + expectedName + "] อยู่ที่ Index: " + foundIndex + " (คอลัมน์ที่ " + (foundIndex + 1) + ")");
        } else {
            Logger.log("❌ ไม่พบ! หัวคอลัมน์ [" + expectedName + "] โปรดเช็คตัวสะกด ช่องว่าง หรือการเว้นวรรคในชีตครับ");
            errorCount++;
        }
    }

    if (errorCount === 0) {
        Logger.log("\n🎉 ยอดเยี่ยม! ระบบค้นหาคอลัมน์เจอครบ 100% โครงสร้าง Dynamic พร้อมใช้งานแล้ว!");
    } else {
        Logger.log("\n⚠️ พบ " + errorCount + " จุดที่หาไม่เจอ โปรดแก้ไขให้ตรงกันก่อนไปต่อครับ");
    }
}

// 2. คำนวณปีงบประมาณ และ สร้างรหัสโครงการอัตโนมัติ (เช่น P-2569-1074)
function generateNextProjectID() {
    var d = new Date();
    var year = d.getFullYear() + 543; // ค.ศ. -> พ.ศ.
    var month = d.getMonth() + 1;     // 1-12
    if (month >= 10) year += 1;       // ต.ค. เป็นต้นไป ปัดเป็นปีงบถัดไป

    var fy = year.toString();
    var prefix = "P-" + fy + "-";

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('m_actionplan');
    var map = getColumnMap(sheet);
    var data = sheet.getDataRange().getValues();

    var maxNum = 0;
    for (var i = 1; i < data.length; i++) {
        var currentId = String(data[i][map[COL_NAME.ID]]).trim();
        if (currentId.indexOf(prefix) === 0) {
            var numStr = currentId.replace(prefix, "");
            var num = parseInt(numStr, 10);
            if (!isNaN(num) && num > maxNum) {
                maxNum = num;
            }
        }
    }

    var nextNumStr = (maxNum + 1).toString();
    while (nextNumStr.length < 4) nextNumStr = "0" + nextNumStr; // เติม 0 ให้ครบ 4 หลัก

    return prefix + nextNumStr;
}

// 3. บันทึกข้อมูล (เพิ่มโครงการใหม่ หรือ อัปเดตโครงการเดิม)
function saveMasterPlan(payload) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName('m_actionplan');
        var map = getColumnMap(sheet);

        var selectedMonths = [];
        if (payload.months) {
            selectedMonths = payload.months.split(',').map(function (m) { return m.trim(); });
        }
        var checkMonth = function (monthName) { return selectedMonths.indexOf(monthName) !== -1 ? 1 : ""; };

        if (payload.id) { // ✏️ กรณีแก้ไข (UPDATE)
            var data = sheet.getDataRange().getValues();
            var rowIndex = -1;
            for (var i = 1; i < data.length; i++) {
                if (String(data[i][map[COL_NAME.ID]]) === String(payload.id)) { rowIndex = i + 1; break; }
            }
            if (rowIndex > -1) {
                if (map[COL_NAME.YEAR] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.YEAR] + 1).setValue(payload.year);
                if (map[COL_NAME.CAT] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.CAT] + 1).setValue(payload.cat);

                sheet.getRange(rowIndex, map[COL_NAME.ORDER] + 1).setValue(payload.order);
                sheet.getRange(rowIndex, map[COL_NAME.DEPT] + 1).setValue(payload.dept);
                sheet.getRange(rowIndex, map[COL_NAME.PLAN] + 1).setValue(payload.plan);
                sheet.getRange(rowIndex, map[COL_NAME.PROJ] + 1).setValue(payload.proj);
                sheet.getRange(rowIndex, map[COL_NAME.ACT] + 1).setValue(payload.act);
                sheet.getRange(rowIndex, map[COL_NAME.SUB] + 1).setValue(payload.sub);
                sheet.getRange(rowIndex, map[COL_NAME.TYPE] + 1).setValue(payload.type);
                sheet.getRange(rowIndex, map[COL_NAME.SOURCE] + 1).setValue(payload.source);

                if (map[COL_NAME.BUDGET_CODE] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.BUDGET_CODE] + 1).setValue(payload.budgetCode);
                if (map[COL_NAME.ACT_CODE] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.ACT_CODE] + 1).setValue(payload.actCode);

                sheet.getRange(rowIndex, map[COL_NAME.APPROVE] + 1).setValue(payload.approve);
                sheet.getRange(rowIndex, map[COL_NAME.ALLOC] + 1).setValue(payload.allocated);
                sheet.getRange(rowIndex, map[COL_NAME.STATUS] + 1).setValue(payload.status);
                if (map[COL_NAME.REMARK] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.REMARK] + 1).setValue(payload.remark);

                // 🌟 เพิ่มการอัปเดต 4 คอลัมน์ใหม่
                if (map[COL_NAME.OP_TYPE] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.OP_TYPE] + 1).setValue(payload.opType || "");
                if (map[COL_NAME.EXP_TYPE] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.EXP_TYPE] + 1).setValue(payload.expType || "");
                if (map[COL_NAME.PROJ_REMARK] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.PROJ_REMARK] + 1).setValue(payload.projRemark || "");
                if (map[COL_NAME.EXPENSE_DETAIL] !== undefined) sheet.getRange(rowIndex, map[COL_NAME.EXPENSE_DETAIL] + 1).setValue(payload.expDetail || "");

                sheet.getRange(rowIndex, map[COL_NAME.M_OCT] + 1).setValue(checkMonth('ต.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_NOV] + 1).setValue(checkMonth('พ.ย.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_DEC] + 1).setValue(checkMonth('ธ.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_JAN] + 1).setValue(checkMonth('ม.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_FEB] + 1).setValue(checkMonth('ก.พ.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_MAR] + 1).setValue(checkMonth('มี.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_APR] + 1).setValue(checkMonth('เม.ย.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_MAY] + 1).setValue(checkMonth('พ.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_JUN] + 1).setValue(checkMonth('มิ.ย.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_JUL] + 1).setValue(checkMonth('ก.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_AUG] + 1).setValue(checkMonth('ส.ค.'));
                sheet.getRange(rowIndex, map[COL_NAME.M_SEP] + 1).setValue(checkMonth('ก.ย.'));

                return { status: 'success', message: "อัปเดตข้อมูลสำเร็จ" };
            }
        } else { // ➕ กรณีเพิ่มใหม่ (INSERT)
            var newRow = new Array(sheet.getLastColumn()).fill("");
            var newId = generateNextProjectID();

            newRow[map[COL_NAME.ID]] = newId;

            if (map[COL_NAME.YEAR] !== undefined) newRow[map[COL_NAME.YEAR]] = payload.year;
            if (map[COL_NAME.CAT] !== undefined) newRow[map[COL_NAME.CAT]] = payload.cat;

            newRow[map[COL_NAME.ORDER]] = payload.order;
            newRow[map[COL_NAME.DEPT]] = payload.dept;
            newRow[map[COL_NAME.PLAN]] = payload.plan;
            newRow[map[COL_NAME.PROJ]] = payload.proj;
            newRow[map[COL_NAME.ACT]] = payload.act;
            newRow[map[COL_NAME.SUB]] = payload.sub;
            newRow[map[COL_NAME.TYPE]] = payload.type;
            newRow[map[COL_NAME.SOURCE]] = payload.source;

            if (map[COL_NAME.BUDGET_CODE] !== undefined) newRow[map[COL_NAME.BUDGET_CODE]] = payload.budgetCode;
            if (map[COL_NAME.ACT_CODE] !== undefined) newRow[map[COL_NAME.ACT_CODE]] = payload.actCode;

            newRow[map[COL_NAME.APPROVE]] = payload.approve;
            newRow[map[COL_NAME.ALLOC]] = payload.allocated;
            newRow[map[COL_NAME.STATUS]] = payload.status;
            if (map[COL_NAME.REMARK] !== undefined) newRow[map[COL_NAME.REMARK]] = payload.remark;

            // 🌟 เพิ่มการเขียน 4 คอลัมน์ใหม่
            if (map[COL_NAME.OP_TYPE] !== undefined) newRow[map[COL_NAME.OP_TYPE]] = payload.opType || "";
            if (map[COL_NAME.EXP_TYPE] !== undefined) newRow[map[COL_NAME.EXP_TYPE]] = payload.expType || "";
            if (map[COL_NAME.PROJ_REMARK] !== undefined) newRow[map[COL_NAME.PROJ_REMARK]] = payload.projRemark || "";
            if (map[COL_NAME.EXPENSE_DETAIL] !== undefined) newRow[map[COL_NAME.EXPENSE_DETAIL]] = payload.expDetail || "";

            newRow[map[COL_NAME.M_OCT]] = checkMonth('ต.ค.'); newRow[map[COL_NAME.M_NOV]] = checkMonth('พ.ย.');
            newRow[map[COL_NAME.M_DEC]] = checkMonth('ธ.ค.'); newRow[map[COL_NAME.M_JAN]] = checkMonth('ม.ค.');
            newRow[map[COL_NAME.M_FEB]] = checkMonth('ก.พ.'); newRow[map[COL_NAME.M_MAR]] = checkMonth('มี.ค.');
            newRow[map[COL_NAME.M_APR]] = checkMonth('เม.ย.'); newRow[map[COL_NAME.M_MAY]] = checkMonth('พ.ค.');
            newRow[map[COL_NAME.M_JUN]] = checkMonth('มิ.ย.'); newRow[map[COL_NAME.M_JUL]] = checkMonth('ก.ค.');
            newRow[map[COL_NAME.M_AUG]] = checkMonth('ส.ค.'); newRow[map[COL_NAME.M_SEP]] = checkMonth('ก.ย.');

            sheet.appendRow(newRow);
            return { status: 'success', message: 'สร้างโครงการใหม่สำเร็จ รหัส: ' + newId, newId: newId };
        }
    } catch (e) {
        return { status: 'error', message: 'System Error: ' + e.message };
    } finally {
        lock.releaseLock();
    }
}

// ==========================================
// 📌 เปลี่ยนสถานะทีละหลายรายการ & คืนเงินลงถัง (Dynamic Column)
// ==========================================
function submitBulkUpdate(ids, newStatus, remark) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("m_actionplan");
        const logSheet = ss.getSheetByName("log_refunded_budget");

        if (!sheet) return { success: false, message: "หาชีต m_actionplan ไม่พบ" };

        const data = sheet.getDataRange().getValues();
        const map = getColumnMap(sheet); // 🌟 เรียกใช้เรดาร์ค้นหาคอลัมน์
        const timestamp = new Date();
        const logsToAppend = [];

        for (let i = 1; i < data.length; i++) {
            let row = data[i];
            let rowId = row[map[COL_NAME.ID]];

            if (ids.includes(rowId)) {
                // 1. อัปเดตสถานะและหมายเหตุลงชีต แบบ Dynamic
                sheet.getRange(i + 1, map[COL_NAME.STATUS] + 1).setValue(newStatus);
                if (remark) {
                    sheet.getRange(i + 1, map[COL_NAME.REMARK] + 1).setValue(remark);
                }

                // 2. Logic โยนเงินลงถัง
                let type = row[map[COL_NAME.TYPE]];
                let source = row[map[COL_NAME.SOURCE]];

                if (newStatus === "REFUNDED" && (type === "เงินนอก สป." || type === "NONMOPH" || String(source).includes("เงินบำรุง"))) {
                    let balance = row[map[COL_NAME.BAL]] === '' ? 0 : Number(row[map[COL_NAME.BAL]]);

                    if (balance > 0) {
                        logsToAppend.push([
                            timestamp,
                            rowId,
                            row[map[COL_NAME.YEAR]],
                            row[map[COL_NAME.CAT]],
                            row[map[COL_NAME.ORDER]],
                            row[map[COL_NAME.DEPT]],
                            row[map[COL_NAME.PLAN]],
                            row[map[COL_NAME.PROJ]],
                            row[map[COL_NAME.ACT]],
                            row[map[COL_NAME.SUB]],
                            row[map[COL_NAME.TYPE]],
                            row[map[COL_NAME.SOURCE]],
                            row[map[COL_NAME.BUDGET_CODE]],
                            row[map[COL_NAME.ACT_CODE]],
                            row[map[COL_NAME.APPROVE]],
                            row[map[COL_NAME.ALLOC]] === '' ? 0 : Number(row[map[COL_NAME.ALLOC]]),
                            row[map[COL_NAME.SPENT]] === '' ? 0 : Number(row[map[COL_NAME.SPENT]]),
                            row[map[COL_NAME.LOAN]] === '' ? 0 : Number(row[map[COL_NAME.LOAN]]),
                            balance,
                            remark || "ส่งคืนจากการทำ Bulk Action"
                        ]);
                    }
                }
            }
        }

        if (logsToAppend.length > 0 && logSheet) {
            logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToAppend.length, logsToAppend[0].length).setValues(logsToAppend);
        }
        return { success: true, message: "อัปเดตข้อมูลสำเร็จ" };
    } catch (error) {
        return { success: false, message: error.toString() };
    }
}

// ==========================================
// 📌 ฟังก์ชันดึงข้อมูลตัวเลือกสำหรับ Modal เพิ่มโครงการ
// ==========================================
function getModalDropdownData() {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var fetchColumnData = function (sheetName) {
            var sheet = ss.getSheetByName(sheetName);
            if (!sheet) return [];
            var lastRow = sheet.getLastRow();
            if (lastRow < 2) return [];
            var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            return [...new Set(data.map(function (r) { return String(r[0]).trim(); }).filter(String))];
        };

        return {
            cats: fetchColumnData('c_category'),   // 👈 เพิ่มใหม่: ดึงข้อมูลหมวด
            depts: fetchColumnData('c_deparment'),
            types: fetchColumnData('c_budget_type'),
            plans: fetchColumnData('c_plan')
        };
    } catch (e) {
        return { error: e.message };
    }
}

// 📌 นำเข้าข้อมูล Master Plan แบบกลุ่ม (Batch Import)
function importMasterPlanBatch(dataArray) {
    var lock = LockService.getScriptLock();
    try {
        // ล็อคสคริปต์ 15 วินาที ป้องกันคนอัปโหลดพร้อมกัน
        lock.waitLock(15000);

        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(SHEET_NAME); // ชีต m_actionplan
        var data = sheet.getDataRange().getValues();
        var headers = data[0]; // หัวตารางหลังบ้าน

        // 1. หาปีงบประมาณจากข้อมูลแรกเพื่อนำไปสร้าง รหัสโครงการ
        var importYear = String(dataArray[0]["ปีงบประมาณ"] || new Date().getFullYear() + 543).trim();
        var maxRunning = 0;

        // 2. วนลูปหา ID ล่าสุดของปีงบประมาณนั้น (เช่น P-2569-XXXX)
        var idColIndex = headers.indexOf(COL_NAME.ID);
        for (var i = 1; i < data.length; i++) {
            var idVal = String(data[i][idColIndex] || "");
            if (idVal.indexOf("P-" + importYear + "-") === 0) {
                var numPart = parseInt(idVal.split("-")[2], 10);
                if (!isNaN(numPart) && numPart > maxRunning) {
                    maxRunning = numPart;
                }
            }
        }

        var newRows = [];
        // ตัวแปรสำหรับจับคู่เดือนที่ผู้ใช้พิมพ์ "ต.ค., พ.ย." ให้ตรงกับหัวคอลัมน์
        var monthMap = ['ต.ค.', 'พ.ย.', 'ธ.ค.', 'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.'];

        // 3. จัดเตรียมข้อมูลสำหรับเขียนลงชีต
        for (var j = 0; j < dataArray.length; j++) {
            var rowObj = dataArray[j];
            maxRunning++; // บวก Running Number ขึ้น 1 อัตโนมัติ
            var runningStr = ("000" + maxRunning).slice(-4); // เติม 0 ข้างหน้าให้ครบ 4 หลัก
            var newId = "P-" + importYear + "-" + runningStr;

            var newRow = new Array(headers.length).fill(""); // สร้าง Array ว่างๆ ความยาวเท่าจำนวนคอลัมน์

            // ฟังก์ชันย่อยสำหรับเขียนค่าลงคอลัมน์ให้ตรงช่อง
            function setVal(colName, val) {
                var idx = headers.indexOf(colName);
                if (idx > -1) newRow[idx] = val;
            }

            // --- 📌 ทำการ Mapping ข้อมูล (เอาข้อมูลจากหน้าบ้าน หยอดลงหลังบ้าน) ---
            setVal(COL_NAME.ID, newId);
            setVal(COL_NAME.ORDER, maxRunning);
            setVal("ปีงบประมาณ", rowObj["ปีงบประมาณ"]);
            setVal(COL_NAME.DEPT, rowObj["กลุ่มงาน/งาน"]);
            setVal(COL_NAME.PROJ, rowObj["โครงการ"]);
            setVal(COL_NAME.ACT, rowObj["กิจกรรมหลัก"]);
            setVal(COL_NAME.SUB, rowObj["กิจกรรมย่อย"]);
            setVal(COL_NAME.TYPE, rowObj["ประเภทงบ"]);
            setVal(COL_NAME.SOURCE, rowObj["แหล่งงบประมาณ"]);
            setVal("ผู้รับผิดชอบ", rowObj["ผู้รับผิดชอบ"]);
            setVal("รหัสงบประมาณ", rowObj["รหัสงบประมาณ"]);
            setVal("รหัสกิจกรรม", rowObj["รหัสกิจกรรม"]);

            // คำนวณยอดเงินเบื้องต้น
            var approveAmt = parseFloat(rowObj["อนุมัติตามแผน"]) || 0;
            var allocAmt = parseFloat(rowObj["จัดสรร"]) || 0;
            setVal(COL_NAME.APPROVE, approveAmt);
            setVal(COL_NAME.ALLOC, allocAmt);
            setVal(COL_NAME.SPENT, 0); // เริ่มต้นเบิกจ่ายเป็น 0
            setVal(COL_NAME.LOAN, 0);  // เริ่มต้นเงินยืมเป็น 0
            setVal(COL_NAME.BAL, allocAmt); // คงเหลือตั้งต้นเท่ากับยอดจัดสรร
            setVal(COL_NAME.STATUS, "ACTIVE");

            // --- 📌 การจัดการคอลัมน์เดือน (ถ้ามี เช่น "ต.ค., พ.ย.") ---
            var monthsStr = String(rowObj["เดือนที่ดำเนินการ"] || "");
            var monthsArr = monthsStr.split(",").map(function (item) { return item.trim(); });

            for (var m = 0; m < monthsArr.length; m++) {
                var mName = monthsArr[m];
                // เช็คว่ามีชื่อเดือนนี้อยู่ใน monthMap ไหม (กันคนพิมพ์ผิด)
                if (monthMap.indexOf(mName) > -1) {
                    setVal(mName, 1); // ไปใส่เลข 1 ที่คอลัมน์ชื่อเดือนนั้นๆ
                }
            }

            newRows.push(newRow); // เก็บเข้าห่อใหญ่รอส่ง
        }

        // 4. บันทึกก้อนข้อมูลลงชีตในรอบเดียว (เสี้ยววินาที!)
        if (newRows.length > 0) {
            var startRow = sheet.getLastRow() + 1;
            sheet.getRange(startRow, 1, newRows.length, headers.length).setValues(newRows);
        }

        return { success: true, count: newRows.length };

    } catch (error) {
        return { success: false, message: error.toString() };
    } finally {
        lock.releaseLock(); // ปลดล็อค
    }
}

// 🛠️ ONE-TIME SCRIPT: อัปเดตรูปแบบวันที่ในฐานข้อมูล
function migrateAllDateFormats() {
    var SPREADSHEET_ID = '1BhZDqEU7XKhgYgYnBrbFI7IMbr_SLdhU8rvhAMddodQ';
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // กำหนดชีตและตำแหน่งคอลัมน์ที่มีวันที่ (เริ่มนับ A=1, B=2...)
    var targetSheets = [
        { name: 't_actionplan', cols: [18] }, // R = 18
        { name: 't_loan', cols: [17, 23, 26, 27] }, // Q=17, W=23, Z=26, AA=27
        { name: 't_allocate', cols: [17] } // Q = 17
    ];

    var totalUpdated = 0;

    for (var s = 0; s < targetSheets.length; s++) {
        var sheetInfo = targetSheets[s];
        var sheet = ss.getSheetByName(sheetInfo.name);
        if (!sheet) continue;

        var lastRow = sheet.getLastRow();
        if (lastRow < 2) continue; // ไม่มีข้อมูลให้ข้ามไป

        for (var c = 0; c < sheetInfo.cols.length; c++) {
            var colIndex = sheetInfo.cols[c];
            var range = sheet.getRange(2, colIndex, lastRow - 1, 1);
            var values = range.getValues();

            for (var r = 0; r < values.length; r++) {
                var val = values[r][0];
                if (val !== "" && val !== null) {
                    values[r][0] = convertToStandardDate_(val);
                    totalUpdated++;
                }
            }

            // บังคับ Format เซลล์ให้เป็น Text เพื่อล็อกไม่ให้ Google Sheets แปลงกลับเป็น M/D/Y อัตโนมัติ
            range.setNumberFormat("@");
            // บันทึกค่ากลับลงชีต
            range.setValues(values);
        }
    }

    SpreadsheetApp.flush();
    Logger.log("✅ แปลงรูปแบบวันที่เรียบร้อยแล้ว จำนวน " + totalUpdated + " เซลล์");
}

// 🔧 ฟังก์ชันย่อยสำหรับแปลงวันที่ (ไม่ต้องกด Run ตัวนี้)
function convertToStandardDate_(val) {
    try {
        var d;

        // 1. ถ้าเป็น Date Object อยู่แล้ว
        if (Object.prototype.toString.call(val) === '[object Date]') {
            d = val;
        }
        // 2. ถ้าเป็น String แบบ YYYY-MM-DD
        else if (typeof val === 'string' && val.match(/^\d{4}-\d{2}-\d{2}$/)) {
            var parts = val.split('-');
            d = new Date(parts[0], parseInt(parts[1]) - 1, parts[2]);
        }
        // 3. ถ้ามี / อยู่แล้ว (อาจจะตรงฟอร์แมตแล้ว)
        else if (typeof val === 'string' && val.indexOf('/') > -1) {
            // เช็คว่าถ้าเป็น DD/MM/YYYY หรือ D/M/YYYY อยู่แล้ว ให้ปล่อยผ่านได้เลย
            if (val.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) return val;
            d = new Date(val);
        }
        // 4. กรณีอื่นๆ
        else {
            d = new Date(val);
        }

        if (!isNaN(d.getTime())) {
            var dd = d.getDate(); // ไม่เติม 0 ข้างหน้าเพื่อให้ออกมาเป็น 15/1/2026 ตามที่ท่านต้องการ (ไม่ใช่ 15/01/2026)
            var mm = d.getMonth() + 1;
            var yyyy = d.getFullYear();
            return dd + '/' + mm + '/' + yyyy;
        }
    } catch (e) {
        // ถ้าจับ Format ไม่ได้จริงๆ ให้คืนค่าเดิมกลับไป
    }
    return val;
}

// 📅 1. ส่วนกลาง: จัดการรูปแบบวันที่ (มาตรฐานใหม่)
// แปลงจากหน้าเว็บ (YYYY-MM-DD) -> (DD/MM/YYYY) เพื่อบันทึกลง Sheet
function formatToStorageDate(dateStr) {
    if (!dateStr) return "";
    try {
        var parts = dateStr.split('-');
        if (parts.length === 3) {
            return parseInt(parts[2], 10) + '/' + parseInt(parts[1], 10) + '/' + parts[0];
        }
    } catch (e) { }
    return String(dateStr);
}

// แปลงจาก Sheet -> (15 ม.ค. 2569) เพื่อแสดงผลบนหน้าเว็บ
function formatToThaiUI(dateVal) {
    if (!dateVal) return "-";
    try {
        var d;
        var strVal = String(dateVal).trim();

        if (Object.prototype.toString.call(dateVal) === '[object Date]') {
            d = dateVal;
        } else if (strVal.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
            var parts = strVal.split('/');
            d = new Date(parts[2], parseInt(parts[1], 10) - 1, parts[0]);
        } else if (strVal.match(/^\d{4}-\d{2}-\d{2}$/)) {
            var parts = strVal.split('-');
            d = new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2]);
        } else {
            d = new Date(dateVal);
        }

        if (isNaN(d.getTime())) return strVal;

        var months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
        return d.getDate() + " " + months[d.getMonth()] + " " + (d.getFullYear() + 543);
    } catch (ex) {
        return String(dateVal);
    }
}

// ==========================================
// 📌 1. ฟังก์ชันล้างหนี้ / คืนเงินยืม (Repayment)
// ==========================================
function saveRepayment(payload) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetM = ss.getSheetByName('m_actionplan');
        var sheetL = ss.getSheetByName('t_loan'); // 🌟 เพิ่มการเชื่อมต่อชีต t_loan

        var projectId = String(payload.projectId).trim();
        var actualSpent = Number(payload.actualSpent) || 0;
        var returnedAmount = Number(payload.returnedAmount) || 0;
        var totalClear = actualSpent + returnedAmount;

        // 🎯 1. อัปเดตชีต t_loan (ประวัติยืมเงิน) ค้นหาจากล่างขึ้นบนด้วย projectId
        if (sheetL) {
            var tData = sheetL.getDataRange().getValues();
            var targetRowL = -1;
            
            for (var j = tData.length - 1; j >= 1; j--) {
                if (String(tData[j][1]).trim() === projectId) { // คอลัมน์ B (Index 1) คือ รหัสโครงการ
                    targetRowL = j + 1;
                    
                    var currentLoanAmt = Number(tData[j][15]) || 0; // คอลัมน์ P (Index 15): ยอดเงินยืม
                    var currentPaid = Number(tData[j][20]) || 0;    // คอลัมน์ U (Index 20): จำนวนเบิกจ่าย
                    
                    var newPaid = currentPaid + actualSpent;
                    // ยอดคงเหลือ = เงินยืม - (ยอดเบิกจ่ายสะสม + เงินทอนที่คืนมา)
                    var newBalance = currentLoanAmt - (newPaid + returnedAmount); 
                    
                    if (newBalance < 0) newBalance = 0;
                    var status = (newBalance === 0) ? "คืนครบ" : "คืนบางส่วน";

                    sheetL.getRange(targetRowL, 20).setValue(status); // คอลัมน์ T (Index 19): สถานะ
                    sheetL.getRange(targetRowL, 21).setValue(newPaid); // คอลัมน์ U (Index 20): จำนวนเบิกจ่าย
                    sheetL.getRange(targetRowL, 22).setValue(newBalance); // คอลัมน์ V (Index 21): คงเหลือ
                    
                    // บันทึกวันที่เบิกจ่าย
                    if (payload.repayDate) {
                        sheetL.getRange(targetRowL, 23).setValue(formatToStorageDate(payload.repayDate)); // คอลัมน์ W (Index 22)
                    }
                    
                    break;
                }
            }
        }

        // 🎯 2. อัปเดตชีต m_actionplan (Master Plan)
        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === projectId) {
                targetRowM = i + 1;
                break;
            }
        }

        if (targetRowM > -1) {
            var spentIdx = map[COL_NAME.SPENT] + 1;
            var loanIdx = map[COL_NAME.LOAN] + 1;

            var curSpent = Number(sheetM.getRange(targetRowM, spentIdx).getValue()) || 0;
            var curLoan = Number(sheetM.getRange(targetRowM, loanIdx).getValue()) || 0;

            if (totalClear > curLoan) {
                return { success: false, message: 'ยอดเคลียร์บิล+เงินทอน มากกว่ายอดเงินยืมคงค้างในระบบ' };
            }

            // 🎯 อัปเดตยอด "เบิกจ่าย" (บวกยอดที่ใช้จริง) 
            sheetM.getRange(targetRowM, spentIdx).setValue(curSpent + actualSpent);
            
            // 🎯 อัปเดตยอด "เงินยืม" (หักยอดเคลียร์บิล+เงินทอน)
            sheetM.getRange(targetRowM, loanIdx).setValue(curLoan - totalClear);

            SpreadsheetApp.flush();
            return { success: true, message: 'บันทึกเคลียร์เงินยืมสำเร็จ!' };
        } else {
            return { success: false, message: 'ไม่พบรหัสโครงการนี้ในระบบ (m_actionplan)' };
        }
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message };
    } finally {
        lock.releaseLock();
    }
}

// ==========================================
// 📌 2. ลบรายการเงินยืม (ปลอดภัยกับสูตร 100%)
// ==========================================
function deleteLoanBackend(rowIndex, projectId, amount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetL = ss.getSheetByName('t_loan');
        var sheetM = ss.getSheetByName('m_actionplan');

        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(projectId).trim()) {
                targetRowM = i + 1; break;
            }
        }

        if (targetRowM > -1) {
            var loanIdx = map[COL_NAME.LOAN] + 1;
            var currentLoan = Number(sheetM.getRange(targetRowM, loanIdx).getValue()) || 0;
            var refundAmount = Number(amount) || 0;

            var newLoan = currentLoan - refundAmount;
            if (newLoan < 0) newLoan = 0;

            // 🎯 อัปเดตลดช่องเงินยืมเท่านั้น (สูตรคงเหลือจะเด้งคืนเอง)
            sheetM.getRange(targetRowM, loanIdx).setValue(newLoan);
        }

        if (sheetL) sheetL.deleteRow(rowIndex);
        SpreadsheetApp.flush();
        return { success: true, status: 'success' };
    } catch (e) {
        return { success: false, message: 'System Error: ' + e.message, status: 'error' };
    } finally { lock.releaseLock(); }
}

// ==========================================
// 📌 3. แก้ไขยอดเงินยืม (ปลอดภัยกับสูตร 100%)
// ==========================================
function updateLoanAmount(projectId, newAmount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetL = ss.getSheetByName('t_loan');
        var sheetM = ss.getSheetByName('m_actionplan');

        var parsedAmount = Number(newAmount) || 0;
        if (parsedAmount < 0) return { status: 'error', message: 'ยอดไม่สามารถติดลบได้' };

        // 🎯 1. ค้นหาและอัปเดตยอดในชีต t_loan (ประวัติยืมเงิน) โดยใช้ รหัสโครงการ (projectId)
        var tData = sheetL.getDataRange().getValues();
        var targetRowL = -1;
        var oldAmount = 0;

        // ค้นหาจากล่างขึ้นบน เพื่อเจอรายการยืมล่าสุดของโครงการนี้
        for (var j = tData.length - 1; j >= 1; j--) {
            if (String(tData[j][1]).trim() === String(projectId).trim()) {
                targetRowL = j + 1;
                oldAmount = Number(tData[j][15]) || 0; // ดึงยอดเงินยืมเดิม (คอลัมน์ P / Index 15)
                break;
            }
        }

        if (targetRowL === -1) {
            return { status: 'error', message: 'ไม่พบประวัติการยืมเงินของโครงการนี้ในระบบ' };
        }

        // 🎯 2. ค้นหาและอัปเดตยอดในชีต m_actionplan
        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(projectId).trim()) {
                targetRowM = i + 1; 
                break;
            }
        }

        if (targetRowM > -1) {
            var loanIdx = map[COL_NAME.LOAN] + 1;
            var balIdx = map[COL_NAME.BAL] + 1;

            var currentMasterLoan = Number(sheetM.getRange(targetRowM, loanIdx).getValue()) || 0;
            var currentBal = Number(sheetM.getRange(targetRowM, balIdx).getValue()) || 0;
            
            // คำนวณส่วนต่างระหว่างยอดใหม่และยอดเดิม
            var diffAmount = parsedAmount - oldAmount;

            if (diffAmount > currentBal) {
                return { status: 'error', message: 'ยอดเงินคงเหลือไม่เพียงพอสำหรับการเพิ่มยอด' };
            }

            // เขียนอัปเดตลงชีต t_loan คอลัมน์ P (16)
            sheetL.getRange(targetRowL, 16).setValue(parsedAmount);

            // อัปเดตเฉพาะช่องเงินยืมใน m_actionplan โดยบวกส่วนต่างเข้าไป
            sheetM.getRange(targetRowM, loanIdx).setValue(currentMasterLoan + diffAmount);
            
            SpreadsheetApp.flush();
            return { status: 'success', message: 'แก้ไขยอดเงินยืมสำเร็จ!' };
        }
        
        return { status: 'error', message: 'ไม่พบรหัสโครงการนี้ใน Master Plan' };
    } catch (e) {
        return { status: 'error', message: e.message };
    } finally { 
        lock.releaseLock(); 
    }
}

// ==========================================
// 📌 4. ฟังก์ชันใหม่! คืนเงินยืม (ปลอดภัยกับสูตร 100%)
// ==========================================
function returnLoanBackend(rowIndex, projectId, amount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetL = ss.getSheetByName('t_loan');
        var sheetM = ss.getSheetByName('m_actionplan');

        // 1. เปลี่ยนสถานะใน t_loan (คอลัมน์ที่ 20 คือคอลัมน์สถานะ)
        if (sheetL) {
            sheetL.getRange(rowIndex, 20).setValue("คืนเงินแล้ว");
        }

        // 2. หักยอดออกจาก m_actionplan
        var mData = sheetM.getDataRange().getValues();
        var map = getColumnMap(sheetM);
        var targetRowM = -1;

        for (var i = 1; i < mData.length; i++) {
            if (String(mData[i][map[COL_NAME.ID]]).trim() === String(projectId).trim()) {
                targetRowM = i + 1; break;
            }
        }

        if (targetRowM > -1) {
            var loanIdx = map[COL_NAME.LOAN] + 1;
            var currentLoan = Number(sheetM.getRange(targetRowM, loanIdx).getValue()) || 0;
            var refundAmount = Number(amount) || 0;

            var newLoan = currentLoan - refundAmount;
            if (newLoan < 0) newLoan = 0;

            // 🎯 อัปเดตลดช่องเงินยืม (สูตรจะคืนเงินเข้าคงเหลือเอง)
            sheetM.getRange(targetRowM, loanIdx).setValue(newLoan);
        }

        SpreadsheetApp.flush();
        return { success: true, status: 'success', message: 'บันทึกการคืนเงินสำเร็จ!' };
    } catch (e) {
        return { success: false, status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
    } finally { lock.releaseLock(); }
}

// ==========================================
// 📌 อัปเดตยอดเบิกจ่ายในฐานข้อมูล (อัปเดตชีต t_actionplan คอลัมน์ P)
// ==========================================
function updateTransaction(rowId, projectId, newAmount, oldAmount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheetTx = ss.getSheetByName('t_actionplan');
        var mSheet = ss.getSheetByName('m_actionplan');

        // 🌟 FIX: rowId คือเลขแถวจริงในชีต → ใช้ได้โดยตรง ไม่ต้อง loop หา
        var tRow = parseInt(rowId);

        // 🛡️ ตรวจสอบว่ารหัสโครงการตรงกัน (ป้องกันแก้ผิดแถว กรณีข้อมูลถูกลบ/เพิ่มระหว่างทาง)
        var checkID = String(sheetTx.getRange(tRow, 2).getValue()).trim();
        if (checkID !== String(projectId).trim()) {
            return { status: 'error', message: 'ข้อมูลไม่ตรงกัน อาจมีการเปลี่ยนแปลงข้อมูล โปรดรีเฟรชหน้าเว็บ' };
        }

        // 📝 อัปเดตยอดเบิกจ่ายในประวัติ t_actionplan (คอลัมน์ P = 16)
        sheetTx.getRange(tRow, 16).setValue(newAmount);

        // 📊 อัปเดตยอดเบิกจ่ายรวมใน m_actionplan (ลบยอดเก่า + บวกยอดใหม่)
        if (mSheet) {
            var map = getColumnMap(mSheet);
            var mData = mSheet.getDataRange().getValues();
            var mRowIndex = -1;

            for (var i = 1; i < mData.length; i++) {
                if (String(mData[i][map[COL_NAME.ID]]).trim() === String(projectId).trim()) {
                    mRowIndex = i + 1;
                    break;
                }
            }

            if (mRowIndex !== -1) {
                var cellSpent = mSheet.getRange(mRowIndex, map[COL_NAME.SPENT] + 1);
                var currentSpent = parseFloat(String(cellSpent.getValue()).replace(/,/g, '')) || 0;
                var oldVal = parseFloat(oldAmount) || 0;
                var newVal = parseFloat(newAmount) || 0;
                var newSpentTotal = currentSpent - oldVal + newVal;
                cellSpent.setValue(newSpentTotal);
            }
        }

        SpreadsheetApp.flush();
        return { status: 'success', message: 'แก้ไขยอดเบิกจ่ายเรียบร้อย' };
    } catch (e) {
        return { status: 'error', message: 'เกิดข้อผิดพลาด: ' + e.message };
    } finally {
        lock.releaseLock();
    }
}

// ==========================================
// 📌 หลังบ้าน: อัปเดต/แก้ไข ประวัติจัดสรรงบ
// ==========================================
function updateAllocationRecord(rowId, projectId, newAmount, oldAmount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

        // 🚨 ตรวจสอบ: เปลี่ยนชื่อชีตให้ตรงกับชีตประวัติจัดสรรงบของเจ้านาย
        var sheetAlloc = ss.getSheetByName('t_allocation');
        var data = sheetAlloc.getDataRange().getValues();
        var headers = data[0];

        // 🚨 ตรวจสอบ: เปลี่ยน 'จำนวนเงิน' ให้ตรงกับชื่อหัวตารางของเจ้านาย
        var amountColIdx = headers.indexOf('จำนวนเงิน');

        if (amountColIdx === -1) return { status: 'error', message: 'ไม่พบคอลัมน์จำนวนเงินในตารางครับ' };

        var targetRow = -1;
        for (var i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === String(rowId).trim()) { // สมมติรหัสประวัติอยู่คอลัมน์ A (index 0)
                targetRow = i + 1;
                break;
            }
        }

        if (targetRow > -1) {
            sheetAlloc.getRange(targetRow, amountColIdx + 1).setValue(newAmount);
            return { status: 'success' };
        } else {
            return { status: 'error', message: 'ไม่พบรายการนี้ในระบบ' };
        }
    } catch (e) {
        return { status: 'error', message: e.message };
    } finally {
        lock.releaseLock();
    }
}

// ==========================================
// 📌 หลังบ้าน: ลบ ประวัติจัดสรรงบ (อัปเกรด: Smart Match ขั้นสูงสุด)
// ==========================================
function deleteAllocationRecord(rowId, projectId, amount) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

        var sheetAlloc = ss.getSheetByName('t_allocate');
        var sheetM = ss.getSheetByName('m_actionplan');

        if (!sheetAlloc) return { status: 'error', message: 'ไม่พบชีต t_allocate' };
        if (!sheetM) return { status: 'error', message: 'ไม่พบชีต m_actionplan' };

        var data = sheetAlloc.getDataRange().getValues();
        var targetRow = -1;

        // 🎯 แก้บั๊ก 1: ล้าง "ลูกน้ำ" ออกก่อนแปลงเป็นตัวเลขเสมอ
        var paramProjectId = String(projectId).trim();
        var paramAmount = parseFloat(String(amount).replace(/,/g, '')) || 0;
        var paramRowId = String(rowId).trim();

        // ค้นหาจากล่างขึ้นบน 
        for (var i = data.length - 1; i >= 1; i--) {
            var colA = String(data[i][0]).trim();    // วันที่/Timestamp
            var recProj = String(data[i][1]).trim(); // รหัสโครงการ

            // 🎯 แก้บั๊ก 2: รองรับทั้งประวัติเก่า (Col 15) และประวัติใหม่ (Col 16)
            var recAmtOld = parseFloat(String(data[i][15]).replace(/,/g, '')) || 0;
            var recAmtNew = parseFloat(String(data[i][16]).replace(/,/g, '')) || 0;

            // เงื่อนไขหลัก: รหัสโครงการตรง และ ยอดเงินตรง (เทียบจุดทศนิยมให้ปลอดภัย)
            if (recProj === paramProjectId &&
                (Math.abs(recAmtOld - paramAmount) < 0.01 || Math.abs(recAmtNew - paramAmount) < 0.01)) {
                targetRow = i + 1;
                break;
            }

            // เงื่อนไขรอง: ถ้าหน้าบ้านส่ง Timestamp มาตรงเป๊ะ
            if (paramRowId !== "" && paramRowId !== "undefined" && colA === paramRowId) {
                targetRow = i + 1;
                paramAmount = recAmtNew > 0 ? recAmtNew : recAmtOld; // ดึงยอดจริงมาใช้หักลบ
                paramProjectId = recProj;
                break;
            }
        }

        // 🎯 แผนสำรองไม้ตาย: ถ้าหายอดเงินไม่เจอ ให้จับ "รายการล่าสุด" ของโครงการนี้มาลบเลย
        if (targetRow === -1 && paramProjectId !== "" && paramProjectId !== "undefined") {
            for (var i = data.length - 1; i >= 1; i--) {
                if (String(data[i][1]).trim() === paramProjectId) {
                    targetRow = i + 1;
                    // ดึงยอดที่เจอในชีตกลับมาใช้หักลบ
                    paramAmount = parseFloat(String(data[i][16]).replace(/,/g, '')) || parseFloat(String(data[i][15]).replace(/,/g, '')) || 0;
                    break;
                }
            }
        }

        // 🎯 กระบวนการลบและหักยอดคืน
        if (targetRow > -1) {
            var mData = sheetM.getDataRange().getValues();
            var map = getColumnMap(sheetM);
            var targetRowM = -1;

            for (var j = 1; j < mData.length; j++) {
                if (String(mData[j][map[COL_NAME.ID]]).trim() === paramProjectId) {
                    targetRowM = j + 1;
                    break;
                }
            }

            if (targetRowM > -1) {
                var allocIdx = map[COL_NAME.ALLOC] + 1;
                // ป้องกันกรณีช่องจัดสรรใน Master มีลูกน้ำ
                var currentAlloc = parseFloat(String(sheetM.getRange(targetRowM, allocIdx).getValue()).replace(/,/g, '')) || 0;

                var newAlloc = currentAlloc - paramAmount;
                if (newAlloc < 0) newAlloc = 0;

                sheetM.getRange(targetRowM, allocIdx).setValue(newAlloc);
            }

            sheetAlloc.deleteRow(targetRow);
            SpreadsheetApp.flush();
            return { status: 'success' };
        } else {
            return { status: 'error', message: 'ค้นหาด้วยทุกวิธีแล้ว ไม่พบรายการนี้ครับ (อาจถูกลบไปก่อนหน้าแล้ว)' };
        }
    } catch (e) {
        return { status: 'error', message: e.message };
    } finally {
        lock.releaseLock();
    }
}

// Start function system setting  

// ==========================================
// 📌 LOOKUP SETTINGS MANAGEMENT
// ==========================================

function getLookupList(sheetName) {
    try {
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet: ' + sheetName };
        
        var data = sheet.getDataRange().getValues();
        var list = [];
        
        // เริ่มนับที่ 1 เพื่อข้าม Header (สมมติว่าแถว 1 คือ Header)
        for (var i = 1; i < data.length; i++) {
            var itemText = String(data[i][0]).trim();
            if (itemText !== '') {
                list.push({ text: itemText, row: i + 1 });
            }
        }
        return { status: 'success', data: list };
    } catch (e) { 
        return { status: 'error', message: e.message }; 
    }
}

function addLookupItem(sheetName, itemValue) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet: ' + sheetName };
        
        sheet.appendRow([String(itemValue).trim()]);
        SpreadsheetApp.flush();
        return { status: 'success', message: 'เพิ่มรายการสำเร็จ' };
    } catch (e) { 
        return { status: 'error', message: e.message }; 
    } finally { 
        lock.releaseLock(); 
    }
}

function editLookupItem(sheetName, rowIndex, newValue) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet: ' + sheetName };
        
        sheet.getRange(rowIndex, 1).setValue(String(newValue).trim());
        SpreadsheetApp.flush();
        return { status: 'success', message: 'แก้ไขรายการสำเร็จ' };
    } catch (e) { 
        return { status: 'error', message: e.message }; 
    } finally { 
        lock.releaseLock(); 
    }
}

function deleteLookupItem(sheetName, rowIndex) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
        var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) return { status: 'error', message: 'ไม่พบ Sheet: ' + sheetName };
        
        sheet.deleteRow(rowIndex);
        SpreadsheetApp.flush();
        return { status: 'success', message: 'ลบรายการสำเร็จ' };
    } catch (e) { 
        return { status: 'error', message: e.message }; 
    } finally { 
        lock.releaseLock(); 
    }
}
// End function system setting  

// ==========================================
// 📌 ฟังก์ชันทดสอบระบบ Dynamic Column (กด Run เพื่อเช็คความชัวร์!)
// ==========================================
function testDynamicMapping() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('m_actionplan'); // สแกนตารางหลัก

    if (!sheet) {
        Logger.log("❌ ไม่พบ Sheet ชื่อ 'm_actionplan' โปรดตรวจสอบชื่อ Sheet ครับ");
        return;
    }

    var map = getColumnMap(sheet);
    var keysToTest = Object.keys(COL_NAME);
    var errorCount = 0;

    Logger.log("=== 🔍 เริ่มตรวจสอบ Dynamic Mapping (Plan B) ===");
    Logger.log("จำลองการอ่านหัวตารางจากไฟล์ Google Sheet จริง...\n");

    for (var i = 0; i < keysToTest.length; i++) {
        var key = keysToTest[i];
        var expectedName = COL_NAME[key];
        var foundIndex = map[key];

        if (foundIndex !== undefined) {
            // ถ้าเจอ จะบอกว่าอยู่ Index ที่เท่าไหร่ (และเป็นคอลัมน์ที่เท่าไหร่ในชีต)
            Logger.log("✅ พบคอลัมน์ [" + expectedName + "] อยู่ที่ Index: " + foundIndex + " (คอลัมน์ที่ " + (foundIndex + 1) + ")");
        } else {
            // ถ้าไม่เจอ จะแจ้งเตือนสีแดง
            Logger.log("❌ ไม่พบ! หัวคอลัมน์ [" + expectedName + "] โปรดเช็คตัวสะกด หรือการเว้นวรรคในชีตครับ");
            errorCount++;
        }
    }

    Logger.log("\n=========================================");
    if (errorCount === 0) {
        Logger.log("🎉 ยอดเยี่ยมมากครับเจ้านาย! ระบบค้นหาคอลัมน์เจอครบ 100% (ทั้ง " + keysToTest.length + " คอลัมน์)");
        Logger.log("🛡️ เกราะป้องกันตัวแปรเคลื่อนทำงานสมบูรณ์แบบ พร้อมลุยงานจริงครับ! 🚀");
    } else {
        Logger.log("⚠️ พบปัญหา " + errorCount + " จุดที่หาไม่เจอ");
        Logger.log("💡 คำแนะนำ: โปรดไปที่ชีต m_actionplan แถวที่ 1 แล้วตรวจสอบชื่อหัวตารางให้ตรงกับคำในวงเล็บก้ามปูเป๊ะๆ ครับ (ระวังช่องว่างซ่อนอยู่ด้านหลังคำ)");
    }
    Logger.log("=========================================");
}