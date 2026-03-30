const SPREADSHEET_ID = '1L7jenTSA4Jwmjq5QPsPn4nR-BAam9TfLiYcfrV_B0zU';
const DISTRICT_ORDER = ["เมืองเลย", "นาด้วง", "เชียงคาน", "ปากชม", "ด่านซ้าย", "นาแห้ว", "ภูเรือ", "ท่าลี่", "วังสะพุง", "ภูกระดึง", "ภูหลวง", "ผาขาว", "เอราวัณ", "หนองหิน"];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('LOEI Investment Budget Management System : LIBMS')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function checkConnection() { return "OK"; }

// ==========================================
// 1. DASHBOARD DATA (อัปเดตดึงจากชีต m_budget_...)
// ==========================================
function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hospMap = getHospitalMap(ss);
    
    // เปลี่ยนมาดึงจากชีตหลักที่บันทึกความก้าวหน้า
    const eqData = getSheetData(ss.getSheetByName('m_budget_equipment'));
    const bdData = getSheetData(ss.getSheetByName('m_budget_building'));

    const mapObj = (list, type) => {
      let result = [];
      for(let i=0; i<list.length; i++) {
        let d = list[i];
        let hospId = String(d['รหัสหน่วยบริการ'] || '').trim();
        let info = hospMap[hospId] || { name: hospId, amphoe: '-', unitType: '-' };
        
        result.push({
          year: String(d['ปีงบประมาณ'] || ''), 
          status: String(d['สถานะการดำเนินการ'] || '-').trim(),
          totalBudget: parseMoney(d['วงเงินจัดหาได้'] || d['วงเงินสัญญา'] || d['จำนวนเงินที่เบิกจ่าย']),
          name: String(d['ชื่อรายการ'] || d['รายการ'] || '-'),
          hospId: hospId, 
          hospName: info.name, 
          amphoe: info.amphoe, 
          unitType: info.unitType, 
          dataType: type
        });
      }
      return result;
    };

    const equipment = mapObj(eqData, 'ครุภัณฑ์');
    const building = mapObj(bdData, 'สิ่งก่อสร้าง');
    
    let yearSet = new Set();
    equipment.forEach(x => { if (x.year) yearSet.add(x.year); });
    building.forEach(x => { if (x.year) yearSet.add(x.year); });
    const allYears = Array.from(yearSet).sort().reverse();

    return { success: true, data: { equipment: equipment, building: building, years: allYears, amphoes: DISTRICT_ORDER, version: Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmm") } };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 2. REPORT DATA
// ==========================================
function getReportData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hospMap = getHospitalMap(ss);
    return { success: true, data: {
      equipment: getMBudgetList(ss.getSheetByName('m_budget_equipment'), 'ครุภัณฑ์', hospMap),
      building: getMBudgetList(ss.getSheetByName('m_budget_building'), 'สิ่งก่อสร้าง', hospMap)
    }};
  } catch (e) { return { success: false, error: e.message }; }
}

function getMBudgetList(sheet, type, hospMap) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const rows = data.slice(1);
  return rows.map(row => {
    const getVal = (idx) => {
      let v = row[idx];
      if (v instanceof Date) return Utilities.formatDate(v, "GMT+7", "yyyy-MM-dd");
      return v !== undefined && v !== null ? String(v).trim() : '';
    };

    const hospId = getVal(6);
    const hospInfo = hospMap[hospId] || { name: hospId, amphoe: '-', unitType: '-' };

    if (type === 'ครุภัณฑ์') {
      return {
        id: getVal(0), dataType: 'ครุภัณฑ์', year: getVal(1), name: getVal(5), hospName: hospInfo.name, amphoe: hospInfo.amphoe,
        unitPrice: parseMoney(getVal(8)), totalBudget: parseMoney(getVal(13)), contractAmount: parseMoney(getVal(19)),
        method: getVal(14), procStep: getVal(23), status: getVal(24), spentStatus: getVal(25), risk: getVal(26),
        contractSignDate: getVal(15), contractEndDate: getVal(16), deliveryDate: getVal(17), inspectionDate: getVal(18), paymentDate: getVal(20),
        spentAmount: parseMoney(getVal(21)), balance: parseMoney(getVal(22)), note: getVal(27), period: '-'
      };
    } else {
      return {
        id: getVal(0), dataType: 'สิ่งก่อสร้าง', year: getVal(1), name: getVal(5), hospName: hospInfo.name, amphoe: hospInfo.amphoe,
        unitPrice: parseMoney(getVal(10)), totalBudget: parseMoney(getVal(15)), contractAmount: parseMoney(getVal(21)),
        method: getVal(16), procStep: getVal(31), status: getVal(32), spentStatus: getVal(33), risk: getVal(34),
        contractSignDate: getVal(17), contractEndDate: getVal(18), deliveryDate: getVal(19), inspectionDate: getVal(20), paymentDate: getVal(22),
        spentAmount: parseMoney(getVal(23)), balance: parseMoney(getVal(24)), note: getVal(35),
        period: (getVal(27) || '-') + '/' + (getVal(25) || '-'),
        totalPeriod: parseNum(getVal(25)), yearPeriod: getVal(26), currentPeriod: parseNum(getVal(27)), delayPeriod: parseNum(getVal(28)), delayReason: getVal(29)
      };
    }
  });
}

    // ==========================================
    // 3. PROGRESS SUMMARY
    // ==========================================
    function getProgressSummaryData() {
      try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const hospMap = getHospitalMap(ss);
        const eqRaw = getSheetData(ss.getSheetByName('m_budget_equipment'));
        const bdRaw = getSheetData(ss.getSheetByName('m_budget_building'));

        const mapProgress = (list, type) => {
          return list.map(row => {
            let hospId = String(row['รหัสหน่วยบริการ'] || '').trim();
            let info = hospMap[hospId] || { name: hospId, shortName: hospId, amphoe: '-', hospLevel: '-', sapLevel: '-' };
            
            let subType = '-';
            if (type === 'สิ่งก่อสร้าง') {
              subType = String(row['ประเภทสิ่งก่อสร้าง'] || row['ประเภท'] || row['ลักษณะผูกพัน'] || row['ลักษณะ'] || '-');
            } else {
              subType = String(row['ประเภท'] || row['หมวดหมู่'] || row['ลักษณะ'] || '-');
            }

            let endDate = row['วันสิ้นสุดสัญญา'] || row['สิ้นสุดสัญญา'] || '';
            if (endDate && endDate instanceof Date) {
              endDate = Utilities.formatDate(endDate, "GMT+7", "yyyy-MM-dd");
            } else {
              endDate = String(endDate).trim();
            }

            return {
              id: String(row['id'] || row['รหัสรายการ'] || ''), 
              year: String(row['ปีงบประมาณ'] || ''), 
              hospId: hospId, 
              hospName: info.name, 
              hospShortName: info.shortName, // ส่งชื่อย่อไปให้หน้าเว็บด้วย
              amphoe: info.amphoe,
              hospLevel: String(row['ระดับหน่วยบริการเดิม'] || info.hospLevel || '-'), 
              sapLevel: String(row['ระดับ SAP'] || info.sapLevel || '-'),
              name: String(row['ชื่อรายการ'] || row['รายการ'] || ''), 
              subType: subType,
              unitPrice: parseMoney(row['ราคาต่อหน่วย']), 
              contractAmount: parseMoney(row['วงเงินจัดหาได้'] || row['วงเงินสัญญา']),
              spentAmount: parseMoney(row['จำนวนเงินที่เบิกจ่าย']), 
              status: String(row['สถานะการดำเนินการ'] || ''), 
              spentStatus: String(row['สถานะการเบิกจ่าย'] || ''), 
              risk: String(row['สถานะความเสี่ยง'] || ''),
              totalPeriod: parseNum(row['งวดงานทั้งสิ้น']), 
              yearPeriod: parseNum(row['จำนวนงวดงานในปี']),
              
              // 🚨 แก้ไขตรงนี้: เพิ่ม "ปัจจุบันดำเนินการถึงงวดงาน" ให้ตรงกับใน Sheet
              currentPeriod: parseNum(row['ปัจจุบันดำเนินการถึงงวดงาน'] || row['ปัจจุบันดำเนินการถึงงวด']), 
              
              delayPeriod: parseNum(row['จำนวนงวดงานที่ล่าช้า']), 
              contractEndDate: endDate,
              dataType: type
            };
          });
        };

        const actualBdRaw = getSheetData(ss.getSheetByName('m_budget_building'));

        return { success: true, data: { equipment: mapProgress(eqRaw, 'ครุภัณฑ์'), building: mapProgress(actualBdRaw, 'สิ่งก่อสร้าง') } };
      } catch(e) { return { success: false, error: e.message }; }
    }

// ==========================================
// 4. FORM OPTIONS & SEARCH & SERVICE PLAN
// ==========================================
function getFormOptions() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const options = {};
    const hospSheet = ss.getSheetByName('c_hospital');
    if (hospSheet && hospSheet.getLastRow() > 1) {
      const data = hospSheet.getDataRange().getValues().slice(1);
      options.hospitals = data.map(r => ({ id: String(r[8] || '').trim(), name: String(r[10] || '').trim(), amphoe: String(r[5] || '').trim() })).filter(h => h.id);
    } else { options.hospitals = []; }

    ['c_risk', 'c_procedure', 'c_status', 'c_spent', 'c_proc_ebid_equipment', 'c_proc_specific_equipment', 'c_proc_ebid_building', 'c_proc_specific_building', 'c_proc_selection'].forEach(name => {
      const s = ss.getSheetByName(name);
      options[name] = (s && s.getLastRow() > 1) ? s.getRange(2, 1, s.getLastRow() - 1, 1).getValues().flat().map(String).filter(String) : [];
    });
    return { success: true, data: options };
  } catch (e) { return { success: false, error: e.message }; }
}

function getSearchData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hospMap = getHospitalMap(ss);
    const eqData = getSheetData(ss.getSheetByName('bureau_equipment'));
    const bdData = getSheetData(ss.getSheetByName('bureau_building'));

    const mapObj = (list, type) => {
      let result = [];
      for (let i = 0; i < list.length; i++) {
        let d = list[i];
        let hospId = String(d['รหัสหน่วยบริการ'] || '').trim();
        let info = hospMap[hospId] || { name: hospId, amphoe: '-', unitType: '-' };
        result.push({
          type: type, year: String(d['ปีงบประมาณ'] || ''), hospId: hospId,
          hospName: info.name !== hospId ? info.name : (d['ชื่อหน่วยบริการ'] || hospId),
          amphoe: info.amphoe !== '-' ? info.amphoe : (String(d['อำเภอ'] || '-').trim()),
          unitType: String(d['ประเภทหน่วย'] || info.unitType || '-').trim(),
          hospLevel: String(d['ระดับหน่วยบริการเดิม'] || info.hospLevel || '-').trim(),
          sapLevel: String(d['ระดับ SAP'] || info.sapLevel || '-').trim(),
          name: String(d['รายการ'] || d['ชื่อรายการ'] || '-'),
          price: parseMoney(d['ราคาต่อหน่วย'] || 0),
          budget: parseMoney(d['วงเงินรวม'] || d['วงเงิน'] || 0),
          status: String(d['การพิจารณา'] || '-').trim()
        });
      }
      return result;
    };
    return { success: true, data: { equipment: mapObj(eqData, 'ครุภัณฑ์'), building: mapObj(bdData, 'สิ่งก่อสร้าง') } };
  } catch (e) { return { success: false, error: e.message }; }
}

function getServicePlanData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('m_eq_serviceplan');
    if (!sheet) return { success: false, error: 'ไม่พบชีต m_eq_serviceplan' };

    const data = getSheetData(sheet);
    const result = [];
    for (let i = 0; i < data.length; i++) {
      let d = data[i];
      let name = String(d['ชื่อรายการ'] || d['รายการ'] || '').trim();
      let branch = String(d['สาขา SP'] || d['สาขา Service plan'] || d['สาขา'] || '-').trim();
      let hospName = String(d['หน่วยบริการ'] || d['ชื่อหน่วยบริการ'] || '-').trim();
      
      if (name) {
        result.push({
          branch: branch, hospName: hospName, name: name, unitPrice: parseMoney(d['ราคาต่อหน่วย'] || 0),
          y2571: parseInt(d['แผนคำขอปีงบประมาณ 2571']) || 0, y2572: parseInt(d['แผนคำขอปีงบประมาณ 2572']) || 0,
          y2573: parseInt(d['แผนคำขอปีงบประมาณ 2573']) || 0, qty: parseInt(d['รวมจำนวน']) || 0, budget: parseMoney(d['วงเงินรวม'] || 0)
        });
      }
    }
    return { success: true, data: result };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 5. UPDATE RECORD
// ==========================================
function updateBudgetRecord(form) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const isEq = form.budgetType === 'ครุภัณฑ์';
    const sheetName = isEq ? 'm_budget_equipment' : 'm_budget_building';
    const logSheetName = isEq ? 't_equipment_log' : 't_building_log';
    
    const sheet = ss.getSheetByName(sheetName);
    let logSheet = ss.getSheetByName(logSheetName);
    if (!sheet) return { success: false, error: 'ไม่พบชีต ' + sheetName };
    if (!logSheet) logSheet = ss.insertSheet(logSheetName);

    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(r => String(r[0]) === String(form.id));
    if (rowIndex === -1) return { success: false, error: 'ไม่พบ ID: ' + form.id };
    const rowNum = rowIndex + 1;

    const contract = parseMoney(form.contract);
    const spentAmount = parseMoney(form.spentAmount);
    const balance = contract - spentAmount;

    if (isEq) {
        sheet.getRange(rowNum, 15).setValue(form.method || ''); sheet.getRange(rowNum, 20).setValue(contract);
        sheet.getRange(rowNum, 24).setValue(form.procStep || ''); sheet.getRange(rowNum, 25).setValue(form.status || '');
        sheet.getRange(rowNum, 26).setValue(form.spentStatus || ''); sheet.getRange(rowNum, 27).setValue(form.risk || '');
        if (form.contractSignDate !== undefined) sheet.getRange(rowNum, 16).setValue(form.contractSignDate || '');
        if (form.contractEndDate !== undefined) sheet.getRange(rowNum, 17).setValue(form.contractEndDate || '');
        if (form.deliveryDate !== undefined) sheet.getRange(rowNum, 18).setValue(form.deliveryDate || '');
        if (form.inspectionDate !== undefined) sheet.getRange(rowNum, 19).setValue(form.inspectionDate || '');
        if (form.paymentDate !== undefined) sheet.getRange(rowNum, 21).setValue(form.paymentDate || '');
        sheet.getRange(rowNum, 22).setValue(spentAmount); sheet.getRange(rowNum, 23).setValue(balance);
        if (form.note !== undefined) sheet.getRange(rowNum, 28).setValue(form.note || '');
    } else {
        sheet.getRange(rowNum, 17).setValue(form.method || ''); sheet.getRange(rowNum, 22).setValue(contract);
        sheet.getRange(rowNum, 32).setValue(form.procStep || ''); sheet.getRange(rowNum, 33).setValue(form.status || '');
        sheet.getRange(rowNum, 34).setValue(form.spentStatus || ''); sheet.getRange(rowNum, 35).setValue(form.risk || '');
        if (form.contractSignDate !== undefined) sheet.getRange(rowNum, 18).setValue(form.contractSignDate || '');
        if (form.contractEndDate !== undefined) sheet.getRange(rowNum, 19).setValue(form.contractEndDate || '');
        if (form.deliveryDate !== undefined) sheet.getRange(rowNum, 20).setValue(form.deliveryDate || '');
        if (form.inspectionDate !== undefined) sheet.getRange(rowNum, 21).setValue(form.inspectionDate || '');
        if (form.paymentDate !== undefined) sheet.getRange(rowNum, 23).setValue(form.paymentDate || '');
        sheet.getRange(rowNum, 24).setValue(spentAmount); sheet.getRange(rowNum, 25).setValue(balance);
        if (form.totalPeriod !== undefined) sheet.getRange(rowNum, 26).setValue(form.totalPeriod || '');
        if (form.yearPeriod !== undefined) sheet.getRange(rowNum, 27).setValue(form.yearPeriod || '');
        if (form.currentPeriod !== undefined) sheet.getRange(rowNum, 28).setValue(form.currentPeriod || '');
        if (form.delayPeriod !== undefined) sheet.getRange(rowNum, 29).setValue(form.delayPeriod || '');
        if (form.delayReason !== undefined) sheet.getRange(rowNum, 30).setValue(form.delayReason || '');
        if (form.note !== undefined) sheet.getRange(rowNum, 36).setValue(form.note || '');
    }

    const updatedRow = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
    logSheet.appendRow([new Date(), ...updatedRow.slice(1)]);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; } finally { lock.releaseLock(); }
}

function saveBudgetRecord(form) { return { success: true, id: 'dummy' }; } 

// ==========================================
// 6. HELPERS
// ==========================================
function parseMoney(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  let s = String(val).trim();
  let neg = false;
  if (s.startsWith('(') && s.endsWith(')')) { neg = true; s = s.slice(1, -1); }
  s = s.replace(/[, ]/g, '').replace(/[^\d.-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : (neg ? -n : n);
}
function parseNum(val) { return parseMoney(val); }

function getHospitalMap(ss) {
  const sheet = ss.getSheetByName('c_hospital');
  if(!sheet) return {};
  const data = getSheetData(sheet);
  const map = {};
  data.forEach(r => {
    const id = String(r['รหัสหน่วยบริการ'] || r['รหัส'] || '').trim();
    let fullName = String(r['ชื่อเต็มหน่วยบริการ'] || r['ชื่อหน่วยบริการ'] || id).trim();
    let shortName = String(r['ชื่อย่อหน่วยบริการ'] || r['ชื่อย่อ'] || fullName).trim(); // ค้นหาชื่อย่อ ถ้าไม่มีใช้ชื่อเต็มแทน
    
    if(id) {
      map[id] = { 
        name: fullName, 
        shortName: shortName, 
        amphoe: String(r['อำเภอ'] || '').trim(), 
        unitType: String(r['ประเภทหน่วย'] || '').trim(), 
        hospLevel: String(r['ระดับหน่วยบริการเดิม'] || '-').trim(), 
        sapLevel: String(r['ระดับ SAP'] || '-').trim()
      };
    }
  });
  return map;
}

function getSheetData(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    let rowObj = {};
    for (let c = 0; c < headers.length; c++) {
      let val = data[r][c];
      if (val instanceof Date) val = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
      else if (val === undefined || val === null) val = '';
      rowObj[headers[c]] = val;
    }
    rows.push(rowObj);
  }
  return rows;
}

// ==========================================
// 4.5 UC BUDGET DATA (งบค่าเสื่อม) - Dynamic & Batch Update
// ==========================================

    function getUCBudgetData() {
    try {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sheet = ss.getSheetByName('m_uc_budget');
      if (!sheet) return { success: false, error: 'ไม่พบชีต m_uc_budget' };

      const data = getSheetData(sheet); 
      const hospMap = getHospitalMap(ss);
      
      const result = data.map(d => {
        const hospId = String(d['หน่วยบริการลูกข่าย'] || '').trim();
        const info = hospMap[hospId] || { name: hospId, amphoe: '-', unitType: '-' };
        
        return {
          id: String(d['รหัสรายการ'] || ''),
          year: String(d['ปีงบประมาณ'] || ''),
          motherHosp: String(d['หน่วยบริการแม่ข่าย'] || '-'),
          hospId: hospId,
          hospName: info.name,
          amphoe: info.amphoe,
          affiliation: String(d['สังกัด'] || '-').trim(), 
          fundType: String(d['วงเงิน'] || '-').trim(), 
          category: String(d['ประเภท'] || '-'),
          itemName: String(d['รายการ'] || '-'),
          
          // 💡 สิ่งที่เพิ่มเข้ามาใหม่ (จำนวน & สมทบเงินบำรุง)
          qty: String(d['จำนวน'] || '-').trim(),
          contribution: parseMoney(d['สมทบเงินบำรุง'] || d['สมทบเงินบำรุง (บาท)']),
          
          budgetUC: parseMoney(d['งบค่าเสื่อมUC'] || d['งบค่าเสื่อมUC (บาท)']), 
          totalAmount: parseMoney(d['รวมเงิน'] || d['รวมเงิน (บาท)']),
          
          statusRaw: String(d['สถานะการดำเนินงาน'] || '0').split('-')[0].trim(),
          statusText: String(d['สถานะการดำเนินงาน'] || '0-ยังไม่ดำเนินการ'),
          spentUC: parseMoney(d['งบค่าเสื่อมUCเบิกจ่ายแล้ว'] || d['งบค่าเสื่อมUCเบิกจ่ายแล้ว (บาท)']),
          balanceUC: parseMoney(d['งบค่าเสื่อมUCเหลือจ่าย'] || d['งบค่าเสื่อมUCเหลือจ่าย (บาท)']),
          percentBalance: parseNum(d['%UCเหลือจ่าย']),
          dataType: 'UC'
        };
      });
      return { success: true, data: result };
    } catch (e) { return { success: false, error: e.message }; }
  }

/**
 * ฟังก์ชันสำหรับอัปเดตข้อมูล UC แบบ Multiple (Batch Update) 
 * รองรับ Dynamic Column Mapping และ Audit Log
 */
function batchUpdateUCRecords(payload) {
  // payload.items จะเป็น Array: [{id, status, spentAmount, balanceAmount, percentBalance}]
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('m_uc_budget');
    let logSheet = ss.getSheetByName('t_uc_budget');
    
    if (!sheet) return { success: false, error: 'ไม่พบชีต m_uc_budget' };
    if (!logSheet) return { success: false, error: 'ไม่พบชีต t_uc_budget สำหรับเก็บ Log' };

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    const headers = data[0].map(h => String(h).trim());
    
    // 1. Dynamic Mapping: หา Index ของคอลัมน์เป้าหมาย
    const idIdx = headers.indexOf('รหัสรายการ');
    const statusIdx = headers.indexOf('สถานะการดำเนินงาน');
    // ใช้ findIndex เผื่อหัวตารางมีวงเล็บ (บาท) ติดมาด้วย
    const spentIdx = headers.findIndex(h => h.includes('เบิกจ่ายแล้ว')); 
    const balanceIdx = headers.findIndex(h => h.includes('เหลือจ่าย') && !h.includes('%')); 
    const pctIdx = headers.findIndex(h => h.includes('%UCเหลือจ่าย'));
    
    if (idIdx === -1 || statusIdx === -1) {
      return { success: false, error: 'หัวตาราง m_uc_budget ไม่ถูกต้อง (หารหัสรายการ หรือ สถานะ ไม่พบ)' };
    }

    // เตรียมหาหัวตารางของ t_uc_budget ไว้ทำ Audit Log
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    
    const updates = payload.items; 
    let updatedCount = 0;

    // 2. ลูปหาแถวและแก้ไขข้อมูล
    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][idIdx]);
      const updateData = updates.find(u => u.id === rowId);
      
      if (updateData) {
        const rowNum = i + 1;
        const timestamp = new Date();
        
        // อัปเดตลง m_uc_budget
        sheet.getRange(rowNum, statusIdx + 1).setValue(updateData.status);
        data[i][statusIdx] = updateData.status;

        // ถ้ามีการส่งค่ายอดเบิกจ่ายมา (กรณีสถานะ 5)
        if (updateData.status.startsWith('5') && spentIdx !== -1) {
            sheet.getRange(rowNum, spentIdx + 1).setValue(updateData.spentAmount);
            data[i][spentIdx] = updateData.spentAmount;
            
            if (balanceIdx !== -1) {
                sheet.getRange(rowNum, balanceIdx + 1).setValue(updateData.balanceAmount);
                data[i][balanceIdx] = updateData.balanceAmount;
            }
            if (pctIdx !== -1) {
                sheet.getRange(rowNum, pctIdx + 1).setValue(updateData.percentBalance);
                data[i][pctIdx] = updateData.percentBalance;
            }
        }

        // 3. สร้าง Audit Log Map ลงชีต t_uc_budget อัตโนมัติ
        let logRow = new Array(logHeaders.length).fill('');
        for(let c = 0; c < logHeaders.length; c++) {
           let lHeader = logHeaders[c];
           // ถ้าเป็นคอลัมน์เก็บเวลา (สร้างไว้เผื่อ)
           if(lHeader.toLowerCase() === 'timestamp' || lHeader === 'วันเวลาที่แก้ไข') {
               logRow[c] = timestamp;
           } else {
               // แมปข้อมูลจาก m_uc_budget ที่อัปเดตแล้ว มาใส่ log
               let mIdx = headers.indexOf(lHeader);
               if(mIdx !== -1) logRow[c] = data[i][mIdx];
           }
        }
        
        // ถ้าหาหัวตาราง Log ไม่ตรงเลย ให้ยัด Timestamp แถวหน้าสุด แล้วตามด้วย Data ทั้งหมด
        if (logRow.join('').trim() === String(timestamp).trim()) {
            logRow = [timestamp, ...data[i]];
        }
        
        logSheet.appendRow(logRow);
        updatedCount++;
      }
    }

    return { success: true, count: updatedCount };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// ฟังก์ชันอัปเดตข้อมูลจากไฟล์ Excel แบบเหมาเข่ง
// ==========================================
function batchUpdateFromExcel(updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // ป้องกันคนบันทึกพร้อมกัน
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('m_uc_budget');
    if (!sheet) throw new Error('ไม่พบชีต m_uc_budget');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // สร้าง Index ของหัวคอลัมน์เพื่อความรวดเร็ว
    const hIdx = {};
    headers.forEach((h, i) => hIdx[h.trim()] = i);

    if (hIdx['รหัสรายการ'] === undefined) throw new Error('ไม่พบคอลัมน์ "รหัสรายการ" ในฐานข้อมูล');

    // แปลงข้อมูลที่ส่งมาให้อยู่ในรูป Object Map (อ้างอิงด้วย ID) เพื่อให้ค้นหาได้ทันที
    const updateMap = {};
    updates.forEach(u => updateMap[u.id] = u);

    let matchCount = 0;
    let isModified = false;

    // วนลูปข้อมูลในชีต (เริ่มจากบรรทัดที่ 2 คือ index 1)
    for (let r = 1; r < data.length; r++) {
      const rowId = String(data[r][hIdx['รหัสรายการ']]).trim();
      
      // ถ้ารหัสตรงกับข้อมูลใน Excel ที่ส่งมา
      if (updateMap[rowId]) {
        const u = updateMap[rowId];
        
        // อัปเดตข้อมูลทับลงไป (เช็คก่อนว่ามีคอลัมน์นั้นในระบบไหม)
        if(hIdx['ปีงบประมาณ'] !== undefined && u.year !== '') data[r][hIdx['ปีงบประมาณ']] = u.year;
        if(hIdx['หน่วยบริการแม่ข่าย'] !== undefined && u.mother !== '') data[r][hIdx['หน่วยบริการแม่ข่าย']] = u.mother;
        if(hIdx['หน่วยบริการลูกข่าย'] !== undefined && u.child !== '') data[r][hIdx['หน่วยบริการลูกข่าย']] = u.child;
        if(hIdx['สังกัด'] !== undefined && u.aff !== '') data[r][hIdx['สังกัด']] = u.aff;
        if(hIdx['วงเงิน'] !== undefined && u.fund !== '') data[r][hIdx['วงเงิน']] = u.fund;
        if(hIdx['ประเภท'] !== undefined && u.type !== '') data[r][hIdx['ประเภท']] = u.type;
        if(hIdx['ประเภทครุภัณฑ์'] !== undefined && u.subType !== '') data[r][hIdx['ประเภทครุภัณฑ์']] = u.subType;
        if(hIdx['รายการ'] !== undefined && u.itemName !== '') data[r][hIdx['รายการ']] = u.itemName;
        if(hIdx['จำนวน'] !== undefined && u.qty !== '') data[r][hIdx['จำนวน']] = u.qty;
        
        if(hIdx['งบค่าเสื่อมUC'] !== undefined && u.budgetUC !== '') data[r][hIdx['งบค่าเสื่อมUC']] = u.budgetUC;
        if(hIdx['สมทบเงินบำรุง'] !== undefined && u.contrib !== '') data[r][hIdx['สมทบเงินบำรุง']] = u.contrib;
        if(hIdx['งบอื่นๆ'] !== undefined && u.other !== '') data[r][hIdx['งบอื่นๆ']] = u.other;
        if(hIdx['รวมเงิน'] !== undefined && u.total !== '') data[r][hIdx['รวมเงิน']] = u.total;
        
        if(hIdx['สถานะการดำเนินงาน'] !== undefined && u.status !== '') data[r][hIdx['สถานะการดำเนินงาน']] = u.status;
        if(hIdx['งบค่าเสื่อมUCเบิกจ่ายแล้ว'] !== undefined && u.spent !== '') data[r][hIdx['งบค่าเสื่อมUCเบิกจ่ายแล้ว']] = u.spent;
        if(hIdx['งบค่าเสื่อมUCเหลือจ่าย'] !== undefined && u.balance !== '') data[r][hIdx['งบค่าเสื่อมUCเหลือจ่าย']] = u.balance;
        if(hIdx['%UCเหลือจ่าย'] !== undefined && u.pct !== '') data[r][hIdx['%UCเหลือจ่าย']] = u.pct;

        matchCount++;
        isModified = true;
      }
    }

    // เขียนข้อมูลทั้งหมดทับลงไปในทีเดียว (Batch Write - เร็วมาก)
    if (isModified) {
      sheet.getDataRange().setValues(data);
    }

    return { success: true, count: matchCount };

  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
