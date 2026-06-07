/**
 * Google Apps Script - Student Profile System (ใบประวัตินักศึกษา)
 * Backend: Google Sheets + Google Docs Template -> PDF
 */

const CONFIG = {
  SPREADSHEET_ID: '1EsICliZj8uBZycYCmXAu3GK1CmTPkFwRg3yiXQ2Kyxo',
  TEMPLATE_DOC_ID: '1XF09U2q7vLojLkN5PoFbNpVWwIPF0tEIc9-b0la_Fxc', 
  DESTINATION_FOLDER_ID: '1Ez4qSB5bw4i1EkeM7FIFxeK2WxzmC1p8',
  SHEET_NAME: 'ข้อมูลนักศึกษา',
  ADMIN_PASSWORD: '1234'
};

// Web App entry point
function doGet(e) {
  const mode = e && e.parameter && e.parameter.mode;
  const action = e && e.parameter && e.parameter.action;
  
  if (mode === 'create_template') {
    return HtmlService.createHtmlOutput('<h3>ไม่ต้องสร้างเทมเพลตแล้วครับ!</h3><p>ระบบได้ถูกอัปเกรดให้ใช้ HTML Template ໃນการสร้าง PDF โดยตรงแล้ว ไม่ต้องพึ่งพาการวาดตารางใน Google Docs อีกต่อไป ระบบจะจัดฟอร์มและสร้าง PDF ให้เป๊ะขึ้นและทำงานเร็วขึ้นครับ</p>');
  }
  if (mode === 'test_pdf' || mode === 'inspect') {
    return HtmlService.createHtmlOutput('<h3>ไม่มีความจำเป็นต้องตรวจสอบเทมเพลต Google Docs แล้วครับ</h3><p>ปัจจุบันระบบใช้ HTML ในการแปลง PDF แทนการดึงจาก Google Docs</p>');
  }
  if (mode === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('ระบบจัดการข้อมูลนักศึกษา (Admin)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (mode === 'qa') {
    if (action === 'create_mock') {
      try {
        const results = createThreeMockSubmissions();
        let html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>สร้างข้อมูลจำลองสำเร็จ</title><script src="https://cdn.tailwindcss.com"></script><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet"><style>body { font-family: "Sarabun", sans-serif; }</style></head><body class="bg-slate-100 p-6 min-h-screen flex items-center justify-center"><div class="max-w-md w-full bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-200 p-6 text-center space-y-4">';
        html += '<div class="bg-emerald-50 p-4 rounded-full text-emerald-600 w-16 h-16 flex items-center justify-center mx-auto border border-emerald-200">';
        html += '<svg class="w-8 h-8" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path></svg></div>';
        html += '<h3 class="text-xl font-bold text-slate-800">สร้างข้อมูลจำลองสำเร็จแล้ว!</h3>';
        html += '<p class="text-xs text-slate-500">สร้างข้อมูลนักศึกษาจำลอง 3 รายการพร้อมทั้งสร้างไฟล์ PDF เรียบร้อยแล้ว</p>';
        html += '<div class="bg-slate-50 p-3 rounded-lg border text-left space-y-1 text-xs">';
        results.forEach((r, idx) => {
          if (r.success) {
            html += `<p class="text-emerald-700 font-semibold">&#10004; STD-20260${idx+1}: สำเร็จ (<a href="${r.url}" target="_blank" class="underline text-blue-600 font-bold">ดู PDF</a>)</p>`;
          } else {
            html += `<p class="text-red-700 font-semibold">&#10008; STD-20260${idx+1}: ล้มเหลว - ${r.message}</p>`;
          }
        });
        html += '</div>';
        html += '<div class="pt-2"><a href="?mode=admin" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded-xl text-sm transition shadow-md block">เข้าสู่ระบบแอดมิน (Admin Panel)</a></div>';
        html += '<a href="?" class="text-xs text-slate-400 hover:text-slate-600 block transition">&larr; กลับหน้าฟอร์มกรอกประวัติ</a>';
        html += '</div></body></html>';
        return HtmlService.createHtmlOutput(html)
          .setTitle('สร้างข้อมูลจำลองสำเร็จ')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
          .addMetaTag('viewport', 'width=device-width, initial-scale=1');
      } catch (err) {
        return HtmlService.createHtmlOutput('เกิดข้อผิดพลาดในการสร้างข้อมูลจำลอง: ' + err.toString())
          .setTitle('เกิดข้อผิดพลาด');
      }
    }
    return runQATests();
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบบันทึกใบประวัตินักศึกษา')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * รับข้อมูลจากฟอร์มและดำเนินการ
 */
/**
 * ดึงลำดับคอลัมน์ที่กำหนดไว้สำหรับ Google Sheets แบบสม่ำเสมอ
 */
function getExpectedColumns() {
  const columns = ['Timestamp'];

  // Tab 1: ข้อมูลส่วนตัว
  const personalFields = [
    'applyDate', 'applySource', 'applyBranch', 'applyEduLevel', 'applyMajor',
    'titleTh', 'fullNameTh', 'nicknameTh', 'fullNameEn', 'nicknameEn',
    'birthDate', 'ageYears', 'ageMonths',
    'maritalStatus', 'childCount',
    'addressRegister', 'addressCurrent',
    'phonePersonal', 'facebook', 'lineId', 'email',
    'phoneParent', 'livingWith', 'livingWithOtherDetails', 'parentRelationship',
    'height', 'weight', 'bmi', 'bodyShape', 'handedness', 'bodyNotes',
    'eyesightLeftVal', 'eyesightLeftStatus', 'eyesightRightVal', 'eyesightRightStatus',
    'glassesLeft', 'contactsLeft', 'glassesRight', 'contactsRight',
    'colorBlindness', 'tattoo', 'tattooPoints', 'tattooSize', 'tattooArea',
    'militaryStatus', 'militaryYear',
    'nameChanged', 'nameChangedOldName', 'nameChangedOldLastName', 'nameChangedCount', 'nameChangedReason'
  ];
  columns.push(...personalFields);

  // Tab 2: สุขภาพและเอกสารส่วนตัว
  const healthFields = [
    'illness_surgery', 'illness_accident', 'illness_fracture', 'illness_metal', 'illness_joint', 'illness_other',
    'illnessTime', 'illnessCause', 'illnessSymptoms',
    'disease_none', 'disease_asthma', 'disease_depression', 'disease_anemia', 'disease_epilepsy', 'disease_anxiety', 'disease_thalassemia', 'disease_hepb', 'disease_thyroid', 'disease_allergy', 'disease_other',
    'diseaseDetails',
    'bloodGroup', 'plasticSurgery', 'plasticSurgeryDetails', 'dentalBraces', 'dentalBracesEnd',
    'bodyMark', 'bodyMarkSymptoms', 'bodyMarkCause', 'periodPain',
    'smoke', 'smokeQty', 'vape', 'vapeQty', 'cannabis', 'cannabisQty', 'kratom', 'kratomQty', 'alcohol', 'alcoholQty',
    'driverLicense', 'licenseCar', 'licenseMotorcycle',
    'passport', 'passportReason', 'passportCountry', 'passportTime', 'passportExpireYear',
    'japanVisit', 'japanVisitReason', 'japanVisitVisaType', 'japanVisitTime',
    'japanApply', 'japanApplyCompany', 'japanApplyTime', 'japanApplyStep', 'japanApplyWithdrawReason',
    'criminalRecord', 'criminalRecordDetails',
    'studentDebt', 'studentDebtSource', 'studentDebtReason', 'studentDebtAmount',
    'parentDebt', 'parentDebtSource', 'parentDebtReason', 'parentDebtAmount'
  ];
  columns.push(...healthFields);

  // Tab 3: ประวัติครอบครัว
  for (let i = 1; i <= 8; i++) {
    columns.push(`fam_name_${i}`, `fam_rel_${i}`, `fam_age_${i}`, `fam_job_${i}`, `fam_work_${i}`, `fam_inc_${i}`, `fam_prov_${i}`);
  }
  columns.push('familyStatus', 'guardian');
  columns.push('japanContact', 'japanContactName', 'japanContactGender', 'japanContactAge', 'japanContactRelationship', 'japanContactAddress', 'japanContactPhone', 'japanContactYears', 'japanContactFacebook', 'japanContactVisaType');

  // Tab 4: ประวัติการศึกษา
  const eduFields = [
    'edu_m3_school', 'edu_m3_major', 'edu_m3_start_year', 'edu_m3_start_date', 'edu_m3_end_year', 'edu_m3_end_date',
    'edu_voc_school', 'edu_voc_major', 'edu_voc_start_year', 'edu_voc_start_date', 'edu_voc_end_year', 'edu_voc_end_date',
    'edu_hvoc_school', 'edu_hvoc_major', 'edu_hvoc_start_year', 'edu_hvoc_start_date', 'edu_hvoc_end_year', 'edu_hvoc_end_date',
    'edu_uni_school', 'edu_uni_major', 'edu_uni_start_year', 'edu_uni_start_date', 'edu_uni_end_year', 'edu_uni_end_date',
    'eduInterrupted', 'eduInterruptedDetails'
  ];
  columns.push(...eduFields);

  // Tab 5: ประวัติการทำงาน
  for (let i = 1; i <= 9; i++) {
    columns.push(`work_${i}_company`, `work_${i}_province`, `work_${i}_bizType`, `work_${i}_jobDesc`, `work_${i}_salary`, `work_${i}_period`, `work_${i}_type`);
  }

  // Tab 6: เป้าหมายและบันทึกเพิ่มเติม
  const otherFields = ['lifeGoal', 'hobbies', 'strengths', 'weaknesses', 'specialSkills', 'pdfUrl', 'illness_none', 'studentId'];
  columns.push(...otherFields);

  return columns;
}

/**
 * ตรวจสอบและอัปเดตหัวตารางในชีตให้ตรงกับเวอร์ชันล่าสุดโดยอัตโนมัติ
 */
function checkAndSyncHeaders(sheet) {
  const expected = getExpectedColumns();
  const lastCol = sheet.getLastColumn();
  
  let currentHeaders = [];
  if (lastCol > 0) {
    currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  
  let needsSync = false;
  if (currentHeaders.length !== expected.length) {
    needsSync = true;
  } else {
    for (let i = 0; i < expected.length; i++) {
      if (currentHeaders[i] !== expected[i]) {
        needsSync = true;
        break;
      }
    }
  }
  
  if (needsSync) {
    // ปรับปรุงหัวข้อคอลัมน์ในแถวแรก
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
  }
}

function processForm(formData) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }
    
    // ตรวจสอบและอัปเดตหัวตารางทันที
    checkAndSyncHeaders(sheet);
    const columns = getExpectedColumns();
    
    // 1. สร้างไฟล์ PDF จากเทมเพลต Google Docs ก่อน เพื่อนำ URL ไปเซฟในตาราง
    const pdfUrl = generatePDF(formData);
    formData.pdfUrl = pdfUrl;

    // 2. จัดเก็บแถวข้อมูลลงใน Sheet
    const rowData = [new Date()];
    // ลูปตามตำแหน่งคอลัมน์ที่กำหนดไว้เท่านั้นเพื่อหลีกเลี่ยงการสลับตำแหน่ง
    for (let colIdx = 1; colIdx < columns.length; colIdx++) {
      const fieldName = columns[colIdx];
      const val = formData[fieldName];
      
      // แปลงค่า boolean เป็นข้อความเข้าใจง่ายสำหรับแสดงบน Google Sheet
      if (val === true) {
        rowData.push("ใช่");
      } else if (val === false) {
        rowData.push("ไม่");
      } else {
        rowData.push(val !== undefined && val !== null ? val : "");
      }
    }
    
    sheet.appendRow(rowData);

    return {
      success: true,
      url: pdfUrl
    };
  } catch (error) {
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * สร้าง PDF โดยใช้ HtmlService แทนการแก้ไขไฟล์ Google Docs
 */
function generatePDF(data) {
  const folder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
  
  // จัดเตรียมข้อมูลชื่อประกอบไฟล์
  const studentName = data.fullNameTh || 'นักศึกษาใหม่';
  const studentId = data.studentId || '';
  const fileName = `ใบประวัติ_${studentId}_${studentName}`;

  // 1. ปรับปรุงข้อมูลวันเกิด
  if (data.birthDate) {
    try {
      const parts = data.birthDate.split('-');
      if (parts.length === 3) {
        let y = parseInt(parts[0]);
        let m = parseInt(parts[1]) - 1;
        let d = parseInt(parts[2]);
        
        data.birthDate_day = d.toString();
        const thMonths = [
          "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
          "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
        ];
        data.birthDate_month = thMonths[m] || '';
        
        if (y > 2400) {
          data.birthDate_yearTh = y.toString();
          data.birthDate_yearEn = (y - 543).toString();
        } else {
          data.birthDate_yearEn = y.toString();
          data.birthDate_yearTh = (y + 543).toString();
        }
      }
    } catch(e) {
      console.error("Error parsing birthDate: " + e.toString());
    }
  }

  // 2. ใช้ Proxy เพื่อจัดการ Checkbox อัตโนมัติใน HTML
  const templateData = new Proxy(data, {
    get: function(target, prop) {
      if (typeof prop === 'symbol') return target[prop];
      
      // ถ้ามีค่านั้นตรงๆ ในข้อมูล
      if (prop in target && target[prop] !== undefined && target[prop] !== null) {
        if (target[prop] === true || target[prop] === "true" || target[prop] === "on" || target[prop] === "Yes") {
           return "☑";
        } else if (target[prop] === false || target[prop] === "false" || target[prop] === "No") {
           return "☐";
        }
        return target[prop];
      }
      
      // ถ้าเป็นคีย์ที่ไม่มี แต่มีลักษณะเป็น Checkbox Radio (มี _)
      if (typeof prop === 'string' && prop.includes('_')) {
        let parts = prop.split('_');
        let key = parts[0];
        let val = parts.slice(1).join('_');
        
        if (prop.startsWith('work_') && prop.includes('_type_')) {
          const typeMatch = prop.match(/^(work_\d+_type)_(.+)$/);
          if (typeMatch) {
            key = typeMatch[1];
            val = typeMatch[2];
          }
        }
        
        if (target[key] === val) return "☑";
        return "☐";
      }
      
      // ถ้าเป็น boolean ทั่วไปที่ไม่มีในระบบ
      if (typeof prop === 'string' && (prop.startsWith('illness') || prop.startsWith('disease') || prop.startsWith('glasses') || prop.startsWith('contacts') || prop.startsWith('tattoo') || prop.startsWith('license') || prop.startsWith('japanApply') || prop.startsWith('japanContact'))) {
        return "☐";
      }
      
      return ""; // Default
    }
  });

  // 3. เตรียม Html Template
  const template = HtmlService.createTemplateFromFile('PdfTemplate');
  template.data = templateData;
  const htmlOutput = template.evaluate();
  
  // 4. สร้าง PDF
  const blob = htmlOutput.getAs(MimeType.PDF);
  blob.setName(`${fileName}.pdf`);
  const pdfFile = folder.createFile(blob);
  
  // ตั้งค่าสิทธิ์ให้ทุกคนที่มีลิงก์สามารถอ่านได้
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return pdfFile.getUrl();
}

/**
 * ฟังก์ชันสำหรับดีบั๊ก ตรวจสอบความยาวตัวอักษรของเทมเพลต Google Docs
 * สามารถเลือกและรันฟังก์ชันนี้เพื่อตรวจสอบว่าสคริปต์สามารถเปิดและอ่านข้อมูลจากไฟล์เทมเพลตได้หรือไม่
 */
function inspectTemplateDoc() {
  try {
    const doc = DocumentApp.openById(CONFIG.TEMPLATE_DOC_ID);
    const text = doc.getBody().getText();
    console.log("Template Character Length: " + text.length);
    console.log("Template Text Sample: " + text.substring(0, 500));
    return {
      success: true,
      length: text.length,
      sample: text.substring(0, 200)
    };
  } catch (e) {
    console.error("Error inspecting template: " + e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับออกแบบและสร้างเทมเพลต Google Docs จากศูนย์ให้มีฟอร์มเหมือน PDF
 */
function createTemplateDoc_old() {
  const doc = DocumentApp.openById(CONFIG.TEMPLATE_DOC_ID);
  const body = doc.getBody();
  body.clear();
  
  // ตั้งระยะขอบ (Margins)
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);
  
  // ตัวช่วยตั้งฟอนต์ย่อหน้า
  function styleParagraph(p, size, bold, align, color) {
    p.setFontFamily('Sarabun');
    p.setFontSize(size);
    p.setBold(bold);
    if (align) p.setAlignment(align);
    if (color) p.setForegroundColor(color);
    return p;
  }

  // ตัวช่วยสร้างหัวข้อ
  function addHeading(text) {
    const p = body.appendParagraph(text);
    styleParagraph(p, 16, true, DocumentApp.HorizontalAlignment.CENTER, '#1e293b');
    p.setSpacingBefore(12);
    p.setSpacingAfter(12);
    return p;
  }
  
  // ตัวช่วยสร้างหัวข้อย่อย
  function addSubHeading(text) {
    const p = body.appendParagraph(text);
    styleParagraph(p, 11, true, DocumentApp.HorizontalAlignment.LEFT, '#2563eb');
    p.setSpacingBefore(10);
    p.setSpacingAfter(6);
    return p;
  }

  // ตัวช่วยสร้างข้อความปกติ
  function addPara(text) {
    const p = body.appendParagraph(text);
    styleParagraph(p, 9.5, false, DocumentApp.HorizontalAlignment.LEFT, '#334155');
    p.setSpacingAfter(4);
    return p;
  }
  
  // จัดสไตล์ตาราง
  function styleTable(table) {
    const numRows = table.getNumRows();
    for (let r = 0; r < numRows; r++) {
      const row = table.getRow(r);
      const numCells = row.getNumCells();
      for (let c = 0; c < numCells; c++) {
        const cell = row.getCell(c);
        cell.setFontFamily('Sarabun');
        cell.setFontSize(8.5);
        cell.setPaddingTop(4);
        cell.setPaddingBottom(4);
        cell.setPaddingLeft(4);
        cell.setPaddingRight(4);
        if (r === 0) {
          cell.setBold(true);
          cell.setBackgroundColor('#f1f5f9');
          cell.setForegroundColor('#1e293b');
        } else {
          cell.setForegroundColor('#475569');
        }
      }
    }
  }

  // === PAGE 1 ===
  addHeading("ใบประวัตินักศึกษา");
  addPara("วันที่สมัคร: <<applyDate>>   ทราบข้อมูลการสมัครจาก: <<applySource>>   สำนักงานสาขาที่รับสมัคร: <<applyBranch>>");
  
  addSubHeading("ข้อมูลการสมัครและวุฒิการศึกษา");
  addPara("รหัสนักศึกษา: <<studentId>>   สาขาวิชาที่สมัคร: <<applyMajor>>");
  addPara("วุฒิที่ใช้สมัคร: [<<applyEduLevel_m6>>] ม.6   [<<applyEduLevel_nfe>>] กศน.   [<<applyEduLevel_voc>>] ปวช.   [<<applyEduLevel_hvoc>>] ปวส.   [<<applyEduLevel_uni>>] ป.ตรี");
  
  addSubHeading("ประวัติส่วนตัว");
  addPara("ชื่อ-นามสกุล (ไทย): <<titleTh>><<fullNameTh>>   ชื่อเล่น: <<nicknameTh>>");
  addPara("Full Name (English): <<fullNameEn>>   Nickname (English): <<nicknameEn>>");
  addPara("วัน/เดือน/ปีเกิด: <<birthDate>>   อายุ: <<ageYears>> ปี <<ageMonths>> เดือน");
  addPara("สถานภาพ: [<<maritalStatus_single>>] โสด   [<<maritalStatus_married_unreg>>] แต่งงานไม่จดทะเบียน   [<<maritalStatus_married_reg>>] แต่งงานจดทะเบียน   [<<maritalStatus_divorced>>] หย่าร้าง   [<<maritalStatus_has_child>>] มีบุตร (จำนวน: <<childCount>> คน)");
  
  addSubHeading("ข้อมูลติดต่อและที่อยู่");
  addPara("ที่อยู่ตามทะเบียนบ้าน: <<addressRegister>>");
  addPara("ที่อยู่ปัจจุบัน: <<addressCurrent>>");
  addPara("เบอร์โทรส่วนตัว: <<phonePersonal>>   Facebook: <<facebook>>   ID LINE: <<lineId>>   Email: <<email>>");
  addPara("เบอร์โทรผู้ปกครอง: <<phoneParent>>   พักอาศัยอยู่กับ: [<<livingWith_parents>>] พ่อแม่   [<<livingWith_other>>] อื่นๆ (ระบุ: <<livingWithOtherDetails>>)   เกี่ยวเนื่องเป็น: <<parentRelationship>>");
  
  addSubHeading("ข้อมูลสัดส่วนและการตรวจร่างกาย");
  addPara("ส่วนสูง: <<height>> ซม.   น้ำหนัก: <<weight>> กก.   BMI: <<bmi>>   รูปร่าง: [<<bodyShape_thin>>] ผอม   [<<bodyShape_normal>>] ปกติ   [<<bodyShape_fat>>] อ้วน   มือที่ถนัด: [<<handedness_right>>] ขวา   [<<handedness_left>>] ซ้าย");
  addPara("หมายเหตุร่างกาย/สัดส่วนจริง: <<bodyNotes>>");
  addPara("สายตาซ้าย: ค่าสายตา <<eyesightLeftVal>> [<<eyesightLeftStatus_normal>>] ปกติ   [<<eyesightLeftStatus_short>>] สั้น   [<<eyesightLeftStatus_astigmatism>>] เอียง");
  addPara("สายตาขวา: ค่าสายตา <<eyesightRightVal>> [<<eyesightRightStatus_normal>>] ปกติ   [<<eyesightRightStatus_short>>] สั้น   [<<eyesightRightStatus_astigmatism>>] เอียง");
  addPara("การใส่เครื่องช่วยสายตา: [<<glassesLeft>>] แว่นตาซ้าย   [<<contactsLeft>>] คอนแทคเลนส์ซ้าย   [<<glassesRight>>] แว่นตาขวา   [<<contactsRight>>] คอนแทคเลนส์ขวา");
  addPara("ตาบอดสี: [<<colorBlindness_none>>] ไม่บอดสี   [<<colorBlindness_red_green>>] ตาบอดสีแดง-เขียว   [<<colorBlindness_blue_yellow>>] ตาบอดสีน้ำเงิน-เหลือง   [<<colorBlindness_all>>] ตาบอดทุกสี");
  addPara("รอยสัก: [<<tattoo_none>>] ไม่มี   [<<tattoo_has>>] มี (จำนวน: <<tattooPoints>> จุด  ขนาด: <<tattooSize>> ซม.  บริเวณ: <<tattooArea>>)");
  addPara("ภาระทางทหาร: [<<militaryStatus_done>>] เกณฑ์แล้ว   [<<militaryStatus_exempted>>] ยกเว้น   [<<militaryStatus_deferred>>] ผ่อนผัน   [<<militaryStatus_not_yet>>] ยังไม่ถึงเกณฑ์   [<<militaryStatus_active>>] เป็นทหาร (ปีที่เกณฑ์/กำหนดเกณฑ์: <<militaryYear>>)");
  addPara("การเปลี่ยนชื่อ-นามสกุล: [<<nameChanged_none>>] ไม่เคย   [<<nameChanged_has>>] เคย (ชื่อเดิม: <<nameChangedOldName>> นามสกุลเดิม: <<nameChangedOldLastName>> จำนวน: <<nameChangedCount>> ครั้ง สาเหตุ: <<nameChangedReason>>)");

  body.appendPageBreak();

  // === PAGE 2 ===
  addHeading("ข้อมูลสุขภาพและประวัติต่างประเทศ");
  
  addSubHeading("ประวัติการเจ็บป่วย (Medical History)");
  addPara("ประวัติสุขภาพ: [<<illness_none>>] ไม่เคยมีประวัติเจ็บป่วย   [<<illness_surgery>>] เคยผ่าตัด   [<<illness_accident>>] เคยเกิดอุบัติเหตุหนัก   [<<illness_fracture>>] เคยกระดูกหัก   [<<illness_metal>>] มีเหล็กดาม   [<<illness_joint>>] เจ็บหลัง/ข้อต่างๆ   [<<illness_other>>] อื่นๆ");
  addPara("รายละเอียดการเจ็บป่วย: เมื่อใด: <<illnessTime>>   สาเหตุ: <<illnessCause>>   อาการ: <<illnessSymptoms>>");
  
  addSubHeading("โรคประจำตัว");
  addPara("[<<disease_none>>] ไม่มีโรคประจำตัว   [<<disease_asthma>>] หอบหืด   [<<disease_depression>>] ซึมเศร้า   [<<disease_anemia>>] เลือดจาง   [<<disease_epilepsy>>] ลมชัก   [<<disease_anxiety>>] วิตกกังวล   [<<disease_thalassemia>>] ธาลัสซีเมีย   [<<disease_hepb>>] ไวรัสตับอักเสบบี   [<<disease_thyroid>>] ไทรอยด์   [<<disease_allergy>>] ภูมิแพ้   [<<disease_other>>] โรคประจำตัวอื่นๆ (ระบุ: <<diseaseDetails>>)");
  
  addSubHeading("สภาพทางร่างกายทั่วไป");
  addPara("กรุ๊ปเลือด: [<<bloodGroup_A>>] A   [<<bloodGroup_B>>] B   [<<bloodGroup_AB>>] AB   [<<bloodGroup_O>>] O");
  addPara("ศัลยกรรม: [<<plasticSurgery_none>>] ไม่เคยทำศัลยกรรม   [<<plasticSurgery_has>>] เคยทำศัลยกรรม (ระบุรายละเอียด: <<plasticSurgeryDetails>>)");
  addPara("การจัดฟัน: [<<dentalBraces_none>>] ไม่เคยจัดฟัน   [<<dentalBraces_has>>] เคยจัดฟัน (เสร็จสิ้นเมื่อ/สิ้นสุดเมื่อ: <<dentalBracesEnd>>)");
  addPara("ตำหนิทางร่างกาย: [<<bodyMark_none>>] ไม่มีตำหนิชัดเจน   [<<bodyMark_has>>] มีตำหนิ (อาการ/ลักษณะ: <<bodyMarkSymptoms>> สาเหตุ: <<bodyMarkCause>>)");
  addPara("อาการปวดประจำเดือน (เฉพาะผู้หญิง): [<<periodPain_none>>] ไม่ปวด   [<<periodPain_mild>>] ปวดเล็กน้อย   [<<periodPain_severe>>] ปวดมากจนมีผลต่อการดำเนินชีวิตประจำวัน");
  
  addSubHeading("พฤติกรรมการบริโภคและสารเสพติด");
  addPara("การสูบบุหรี่: [<<smoke_none>>] ไม่สูบ   [<<smoke_regular>>] สูบประจำ   [<<smoke_sometimes>>] สูบบางครั้ง (เฉลี่ยวันละ: <<smokeQty>> มวน)");
  addPara("บุหรี่ไฟฟ้า: [<<vape_none>>] ไม่สูบ   [<<vape_regular>>] สูบประจำ   [<<vape_sometimes>>] สูบบางครั้ง (เฉลี่ยวันละ: <<vapeQty>> ครั้ง)");
  addPara("กัญชา/สารเสพติด: [<<cannabis_none>>] ไม่สูบ   [<<cannabis_regular>>] สูบประจำ   [<<cannabis_sometimes>>] สูบบางครั้ง (เฉลี่ยวันละ: <<cannabisQty>> ครั้ง)");
  addPara("การดื่มน้ำกระท่อม: [<<kratom_none>>] ไม่ดื่ม   [<<kratom_regular>>] ดื่มประจำ   [<<kratom_sometimes>>] ดื่มบางครั้ง (เฉลี่ยสัปดาห์ละ: <<kratomQty>> วัน)");
  addPara("เครื่องดื่มแอลกอฮอล์: [<<alcohol_none>>] ไม่ดื่ม   [<<alcohol_regular>>] ดื่มประจำ   [<<alcohol_sometimes>>] ดื่มบางครั้ง (เฉลี่ยสัปดาห์ละ: <<alcoholQty>> วัน)");
  
  addSubHeading("ประวัติทางกฎหมาย เอกสารสิทธิ์ และการเดินทางต่างประเทศ");
  addPara("ใบอนุญาตขับขี่: [<<driverLicense_none>>] ไม่มี   [<<driverLicense_has>>] มี (ประเภท: [<<licenseCar>>] รถยนต์   [<<licenseMotorcycle>>] รถจักรยานยนต์)");
  addPara("พาสปอร์ต: [<<passport_none>>] ไม่มี   [<<passport_has>>] มี (เหตุผล: <<passportReason>>  ประเทศ: <<passportCountry>>  เมื่อใด: <<passportTime>>  หมดอายุปี: <<passportExpireYear>>)");
  addPara("ประวัติการเดินทางไปญี่ปุ่น: [<<japanVisit_none>>] ไม่เคย   [<<japanVisit_has>>] เคยเดินทาง (เหตุผล: <<japanVisitReason>>  ประเภทวีซ่า: <<japanVisitVisaType>>  เมื่อใด: <<japanVisitTime>>)");
  addPara("เคยสมัครไปญี่ปุ่นกับที่อื่น: [<<japanApply_none>>] ไม่เคย   [<<japanApply_has>>] เคยสมัคร (บริษัท: <<japanApplyCompany>>  เมื่อใด: <<japanApplyTime>>  ถึงขั้นตอน: <<japanApplyStep>>  เหตุผลที่ถอนตัว: <<japanApplyWithdrawReason>>)");
  addPara("ประวัติคดีความอาญา: [<<criminalRecord_none>>] ไม่มีคดีความ   [<<criminalRecord_has>>] เคยมีคดีความ (รายละเอียดคดี: <<criminalRecordDetails>>)");
  addPara("ภาระหนี้สินของนักศึกษา: [<<studentDebt_none>>] ไม่มีภาระหนี้สิน   [<<studentDebt_has>>] มีหนี้สิน (กู้จาก: <<studentDebtSource>>  เหตุผล: <<studentDebtReason>>  จำนวน: <<studentDebtAmount>> บาท)");
  addPara("ภาระหนี้สินของครอบครัว: [<<parentDebt_none>>] ไม่มีภาระหนี้สิน   [<<parentDebt_has>>] มีหนี้สิน (กู้จาก: <<parentDebtSource>>  เหตุผล: <<parentDebtReason>>  จำนวน: <<parentDebtAmount>> บาท)");

  body.appendPageBreak();

  // === PAGE 3 ===
  addHeading("ประวัติความสัมพันธ์ครอบครัว");
  addPara("สถานภาพครอบครัว (บิดา-มารดา): [<<familyStatus_together>>] อยู่ร่วมกัน   [<<familyStatus_divorced>>] พ่อแม่หย่าร้าง   [<<familyStatus_separated>>] พ่อแม่แยกทางกัน   [<<familyStatus_father_deceased>>] บิดาเสียชีวิต   [<<familyStatus_mother_deceased>>] มารดาเสียชีวิต");
  addPara("ปัจจุบันอยู่ในความดูแลปกครองของ: <<guardian>>");
  
  addSubHeading("ข้อมูลประวัติครอบครัว (บิดา มารดา และพี่น้องร่วมบิดามารดา สูงสุด 8 ท่าน)");
  const famTable = body.appendTable();
  const famHeader = famTable.appendTableRow();
  famHeader.appendTableCell("ลำดับ").setWidth(40);
  famHeader.appendTableCell("ชื่อ-นามสกุล").setWidth(150);
  famHeader.appendTableCell("ความสัมพันธ์").setWidth(80);
  famHeader.appendTableCell("อายุ (ปี)").setWidth(50);
  famHeader.appendTableCell("อาชีพ").setWidth(90);
  famHeader.appendTableCell("สถานที่ทำงาน/บริษัท").setWidth(120);
  famHeader.appendTableCell("รายได้ (ต่อเดือน)").setWidth(80);
  famHeader.appendTableCell("จังหวัด").setWidth(70);
  
  for (let i = 1; i <= 8; i++) {
    const row = famTable.appendTableRow();
    row.appendTableCell(i.toString());
    row.appendTableCell(`<<fam_name_${i}>>`);
    row.appendTableCell(`<<fam_rel_${i}>>`);
    row.appendTableCell(`<<fam_age_${i}>>`);
    row.appendTableCell(`<<fam_job_${i}>>`);
    row.appendTableCell(`<<fam_work_${i}>>`);
    row.appendTableCell(`<<fam_inc_${i}>>`);
    row.appendTableCell(`<<fam_prov_${i}>>`);
  }
  styleTable(famTable);
  
  addSubHeading("ประวัติบุคคลที่รู้จักหรือมีญาติพำนักอยู่ในประเทศญี่ปุ่น");
  addPara("คนรู้จักในญี่ปุ่น: [<<japanContact_none>>] ไม่มีคนรู้จักพำนักอยู่ในญี่ปุ่น   [<<japanContact_has>>] มีญาติหรือคนรู้จักพำนักอยู่ในญี่ปุ่น");
  addPara("รายละเอียดคนรู้จัก: ชื่อ-นามสกุล: <<japanContactName>>   เพศ: <<japanContactGender>>   อายุ: <<japanContactAge>> ปี   ความสัมพันธ์: <<japanContactRelationship>>");
  addPara("ที่อยู่ติดต่อในญี่ปุ่น: <<japanContactAddress>>   เบอร์โทรศัพท์: <<japanContactPhone>>   ระยะเวลาพำนัก: <<japanContactYears>> ปี");
  addPara("ช่องทางโซเชียล: Facebook: <<japanContactFacebook>>   ประเภทวีซ่าที่ถือครอง: <<japanContactVisaType>>");

  body.appendPageBreak();

  // === PAGE 4 ===
  addHeading("ประวัติการศึกษาและการทำงาน");
  
  addSubHeading("ข้อมูลประวัติการศึกษา");
  const eduTable = body.appendTable();
  const eduHeader = eduTable.appendTableRow();
  eduHeader.appendTableCell("ระดับการศึกษา").setWidth(120);
  eduHeader.appendTableCell("ชื่อสถานศึกษา").setWidth(180);
  eduHeader.appendTableCell("สาขาวิชา / แผนก").setWidth(150);
  eduHeader.appendTableCell("เข้าศึกษา (ปีที่เข้า/วันเริ่ม)").setWidth(120);
  eduHeader.appendTableCell("สำเร็จการศึกษา (ปีที่จบ/วันสำเร็จ)").setWidth(120);
  
  const eduRows = [
    { name: "มัธยมศึกษาตอนต้น (ม.3)", prefix: "edu_m3" },
    { name: "มัธยมปลาย / ปวช.", prefix: "edu_voc" },
    { name: "ปวส. / ทวิภาคี", prefix: "edu_hvoc" },
    { name: "ปริญญาตรี (ถ้ามี)", prefix: "edu_uni" }
  ];
  
  eduRows.forEach(item => {
    const row = eduTable.appendTableRow();
    row.appendTableCell(item.name);
    row.appendTableCell(`<<${item.prefix}_school>>`);
    row.appendTableCell(`<<${item.prefix}_major>>`);
    row.appendTableCell(`<<${item.prefix}_start_year>> (<<${item.prefix}_start_date>>)`);
    row.appendTableCell(`<<${item.prefix}_end_year>> (<<${item.prefix}_end_date>>)`);
  });
  styleTable(eduTable);
  
  addPara("");
  addPara("ประวัติการเรียนสะดุด/ดรอปเรียน/ลาออก: [<<eduInterrupted_none>>] ไม่มีประวัติ   [<<eduInterrupted_drop>>] ดรอปเรียน   [<<eduInterrupted_quit>>] ลาออกระหว่างเรียน (สาเหตุ/รายละเอียด: <<eduInterruptedDetails>>)");
  
  addSubHeading("ประวัติการทำงานและการฝึกงานย้อนหลัง (ลำดับที่ 1 - 3)");
  const workTable1 = body.appendTable();
  const workHeader1 = workTable1.appendTableRow();
  workHeader1.appendTableCell("#").setWidth(30);
  workHeader1.appendTableCell("ชื่อบริษัท/สถานประกอบการ").setWidth(150);
  workHeader1.appendTableCell("จังหวัด").setWidth(70);
  workHeader1.appendTableCell("ประเภทกิจการ").setWidth(90);
  workHeader1.appendTableCell("ลักษณะงานที่ทำ / ตำแหน่ง").setWidth(150);
  workHeader1.appendTableCell("เงินเดือน").setWidth(70);
  workHeader1.appendTableCell("เริ่มงาน-ลาออก").setWidth(100);
  workHeader1.appendTableCell("ประเภทงาน").setWidth(85);
  
  for (let i = 1; i <= 3; i++) {
    const row = workTable1.appendTableRow();
    row.appendTableCell(i.toString());
    row.appendTableCell(`<<work_${i}_company>>`);
    row.appendTableCell(`<<work_${i}_province>>`);
    row.appendTableCell(`<<work_${i}_bizType>>`);
    row.appendTableCell(`<<work_${i}_jobDesc>>`);
    row.appendTableCell(`<<work_${i}_salary>>`);
    row.appendTableCell(`<<work_${i}_period>>`);
    row.appendTableCell(`<<work_${i}_type>>`);
  }
  styleTable(workTable1);

  body.appendPageBreak();

  // === PAGE 5 ===
  addHeading("ประวัติการทำงานและการฝึกงานย้อนหลัง (ต่อ)");
  addSubHeading("ประวัติการทำงานและการฝึกงานย้อนหลัง (ลำดับที่ 4 - 9)");
  
  const workTable2 = body.appendTable();
  const workHeader2 = workTable2.appendTableRow();
  workHeader2.appendTableCell("#").setWidth(30);
  workHeader2.appendTableCell("ชื่อบริษัท/สถานประกอบการ").setWidth(150);
  workHeader2.appendTableCell("จังหวัด").setWidth(70);
  workHeader2.appendTableCell("ประเภทกิจการ").setWidth(90);
  workHeader2.appendTableCell("ลักษณะงานที่ทำ / ตำแหน่ง").setWidth(150);
  workHeader2.appendTableCell("เงินเดือน").setWidth(70);
  workHeader2.appendTableCell("เริ่มงาน-ลาออก").setWidth(100);
  workHeader2.appendTableCell("ประเภทงาน").setWidth(85);
  
  for (let i = 4; i <= 9; i++) {
    const row = workTable2.appendTableRow();
    row.appendTableCell(i.toString());
    row.appendTableCell(`<<work_${i}_company>>`);
    row.appendTableCell(`<<work_${i}_province>>`);
    row.appendTableCell(`<<work_${i}_bizType>>`);
    row.appendTableCell(`<<work_${i}_jobDesc>>`);
    row.appendTableCell(`<<work_${i}_salary>>`);
    row.appendTableCell(`<<work_${i}_period>>`);
    row.appendTableCell(`<<work_${i}_type>>`);
  }
  styleTable(workTable2);

  body.appendPageBreak();

  // === PAGE 6 ===
  addHeading("เป้าหมายในอนาคต & แบบประเมินระดับการรับสมัคร");
  
  addSubHeading("เป้าหมาย เจตคติ และจุดเด่นของนักศึกษา");
  addPara("เป้าหมายในชีวิต: <<lifeGoal>>");
  addPara("กิจกรรมยามว่าง/งานอดิเรก: <<hobbies>>");
  addPara("ข้อดี/จุดแข็งของตนเอง: <<strengths>>");
  addPara("ข้อเสีย/ข้อปรับปรุงของตนเอง: <<weaknesses>>");
  addPara("ความสามารถพิเศษอื่นๆ: <<specialSkills>>");
  
  addSubHeading("แบบประเมินผู้สมัคร (ส่วนการทดสอบและประเมินผลสำหรับเจ้าหน้าที่)");
  
  const evalTable = body.appendTable();
  const evalHeader = evalTable.appendTableRow();
  evalHeader.appendTableCell("หัวข้อการประเมิน").setWidth(250);
  evalHeader.appendTableCell("ผ่าน").setWidth(80);
  evalHeader.appendTableCell("ไม่ผ่าน").setWidth(80);
  evalHeader.appendTableCell("หมายเหตุ / คะแนนสะสม").setWidth(250);
  
  const evalRows = [
    { title: "ทดสอบคณิตศาสตร์ (30 ข้อ เวลา 10 นาที)", score: "......................... / 30 คะแนน (ผ่านเกณฑ์ 15 ข้อ)" },
    { title: "ทดสอบเรียงไพ่ (ดอกจิก หัวใจ โพดำ ข้าวหลามตัด A-K) ครั้งที่ 1", score: "เวลา ................... นาที   ผิด ................... จุด" },
    { title: "ทดสอบเรียงไพ่ (ดอกจิก หัวใจ โพดำ ข้าวหลามตัด A-K) ครั้งที่ 2", score: "เวลา ................... นาที   ผิด ................... จุด" },
    { title: "ทดสอบสมรรถภาพร่างกาย: ลุกนั่ง (1 นาที)", score: "ลุกนั่งได้ ......................... ครั้ง/นาที (ผ่านเกณฑ์ 30 ครั้ง)" },
    { title: "ทดสอบสมรรถภาพร่างกาย: ซิทอัพ (1 นาที)", score: "ซิทอัพได้ ......................... ครั้ง/นาที (ผ่านเกณฑ์ 30 ครั้ง)" },
    { title: "ทดสอบสมรรถภาพร่างกาย: วิดพื้น (1 นาที)", score: "วิดพื้นได้ ......................... ครั้ง/นาที (ผ่านเกณฑ์ 30 ครั้ง)" }
  ];
  
  evalRows.forEach(item => {
    const row = evalTable.appendTableRow();
    row.appendTableCell(item.title);
    row.appendTableCell("      [   ]");
    row.appendTableCell("      [   ]");
    row.appendTableCell(item.score);
  });
  styleTable(evalTable);
  
  addPara("");
  addPara("รายละเอียดและข้อสังเกตเพิ่มเติมของเจ้าหน้าที่ประเมิน:");
  addPara("......................................................................................................................................................................");
  addPara("......................................................................................................................................................................");
  addPara("......................................................................................................................................................................");
  
  addPara("");
  addPara("");
  
  // Signature layout
  const sigTable = body.appendTable();
  const sigRow1 = sigTable.appendTableRow();
  sigRow1.appendTableCell("ลงชื่อ ................................................. ผู้สมัคร\n( <<applicantSignatureName>> )\nวันที่สมัคร: <<applyDate>>").setWidth(300);
  sigRow1.appendTableCell("ลงชื่อ ................................................. ผู้ประเมิน\n( ................................................. )\nวันที่ประเมิน: .................................................").setWidth(300);
  
  // Style sigTable with no borders
  sigTable.setBorderColor('#ffffff'); // ซ่อนเส้นตาราง
  for (let r = 0; r < sigTable.getNumRows(); r++) {
    const row = sigTable.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setFontFamily('Sarabun');
      cell.setFontSize(9.5);
      cell.setPaddingTop(12);
      cell.setPaddingBottom(12);
    }
  }
  
  doc.saveAndClose();
  return { success: true, message: "สร้างเอกสารเทมเพลตใบประวัติและจัดสไตล์โครงสร้างตาราง 6 หน้า เรียบร้อยแล้วครับ" };
}

/**
 * ฟังก์ชันสำหรับจำลองส่งฟอร์มเพื่อทดสอบระบบ (UnitTest / Mock Run)
 * สามารถเลือกฟังก์ชันนี้แล้วกด "เรียกทำงาน" (Run) ใน Apps Script Editor เพื่อทดสอบระบบได้ทันที
 */
function testProcessForm() {
  const mockData = {
    applyDate: "2026-06-02",
    applySource: "การแนะแนว",
    applyBranch: "เชียงใหม่",
    applyEduLevel: "voc",
    applyMajor: "อิเล็กทรอนิกส์",
    studentId: "640001",
    titleTh: "นาย",
    fullNameTh: "ณดูด กะปิไหมจ๊ะ",
    nicknameTh: "โด่ง",
    fullNameEn: "Mr. Nadud Kapimaicha",
    nicknameEn: "Doung",
    birthDate: "2548-01-24",
    ageYears: 20,
    ageMonths: 6,
    maritalStatus: "single",
    childCount: 0,
    addressRegister: "35/4 หมู่ 6 ต.สันกลาง อ.สันกำแพง จ.เชียงใหม่ 50130",
    addressCurrent: "82/5 ต.ต้นเปา อ.สันกำแพง จ.เชียงใหม่ 50130",
    phonePersonal: "0811122233",
    facebook: "Nadud dud dud",
    lineId: "Jabjab888",
    email: "Kapimaicha_888@hotmail.com",
    phoneParent: "0822244288",
    livingWith: "parents",
    livingWithOtherDetails: "",
    parentRelationship: "มารดา",
    height: 178,
    weight: 65,
    bmi: 20.52,
    bodyShape: "normal",
    handedness: "right",
    bodyNotes: "",
    eyesightLeftVal: "1.2",
    eyesightLeftStatus: "normal",
    eyesightRightVal: "1.2",
    eyesightRightStatus: "normal",
    glassesLeft: true,
    contactsLeft: false,
    glassesRight: true,
    contactsRight: false,
    colorBlindness: "none",
    tattoo: "none",
    tattooPoints: 0,
    tattooSize: "",
    tattooArea: "",
    militaryStatus: "nst",
    militaryYear: "",
    nameChanged: "none",
    nameChangedOldName: "",
    nameChangedOldLastName: "",
    nameChangedCount: 0,
    nameChangedReason: "",
    
    // Tab 2
    illness_surgery: false,
    illness_accident: false,
    illness_fracture: false,
    illness_metal: false,
    illness_joint: false,
    illness_other: false,
    illnessTime: "",
    illnessCause: "",
    illnessSymptoms: "",
    disease_none: true,
    disease_asthma: false,
    disease_depression: false,
    disease_anemia: false,
    disease_epilepsy: false,
    disease_anxiety: false,
    disease_thalassemia: false,
    disease_hepb: false,
    disease_thyroid: false,
    disease_allergy: false,
    disease_other: false,
    diseaseDetails: "",
    bloodGroup: "O",
    plasticSurgery: "none",
    plasticSurgeryDetails: "",
    dentalBraces: "none",
    dentalBracesEnd: "",
    bodyMark: "none",
    bodyMarkSymptoms: "",
    bodyMarkCause: "",
    periodPain: "none",
    smoke: "never",
    smokeQty: "",
    vape: "never",
    vapeQty: "",
    cannabis: "never",
    cannabisQty: "",
    kratom: "never",
    kratomQty: "",
    alcohol: "never",
    alcoholQty: "",
    driverLicense: "none",
    licenseCar: false,
    licenseMotorcycle: false,
    passport: "none",
    passportReason: "",
    passportCountry: "",
    passportTime: "",
    passportExpireYear: "",
    japanVisit: "none",
    japanVisitReason: "",
    japanVisitVisaType: "",
    japanVisitTime: "",
    japanApply: "none",
    japanApplyCompany: "",
    japanApplyTime: "",
    japanApplyStep: "",
    japanApplyWithdrawReason: "",
    criminalRecord: "none",
    criminalRecordDetails: "",
    studentDebt: "has",
    studentDebtSource: "กยศ",
    studentDebtReason: "ใช้ในการเรียน",
    studentDebtAmount: "200,000",
    parentDebt: "none",
    parentDebtSource: "",
    parentDebtReason: "",
    parentDebtAmount: "",
    
    // Tab 3
    fam_name_1: "นาย อยู่ดี กะปิไหมจ๊ะ", fam_rel_1: "บิดา", fam_age_1: "50", fam_job_1: "ค้าขาย", fam_work_1: "ตลาดต้นพะยอม", fam_inc_1: "12,000", fam_prov_1: "เชียงใหม่",
    fam_name_2: "นาง มีสุข กะปิไหมจ๊ะ", fam_rel_2: "มารดา", fam_age_2: "45", fam_job_2: "ค้าขาย", fam_work_2: "ตลาดต้นพะยอม", fam_inc_2: "12,000", fam_prov_2: "เชียงใหม่",
    familyStatus: "together",
    guardian: "",
    japanContact: "none",
    
    // Tab 4
    edu_m3_school: "โรงเรียนบ้านหนองกก", edu_m3_major: "ศึกษาทั่วไป", edu_m3_start_year: "2560", edu_m3_start_date: "2560-03-09", edu_m3_end_year: "2563", edu_m3_end_date: "2563-06-24",
    edu_voc_school: "วิทยาลัยเทคโนโลยีแม่กก", edu_voc_major: "อิเล็กทรอนิกส์", edu_voc_start_year: "2563", edu_voc_start_date: "2563-07-05", edu_voc_end_year: "2566", edu_voc_end_date: "2566-08-12",
    edu_hvoc_school: "วิทยาลัยเทคโนโลยีแม่กก", edu_hvoc_major: "อิเล็กทรอนิกส์", edu_hvoc_start_year: "2566", edu_hvoc_start_date: "2566-09-30", edu_hvoc_end_year: "2568", edu_hvoc_end_date: "2568-02-08",
    eduInterrupted: "none",
    
    // Tab 5
    work_1_company: "ตั้งตัวพานิช จำกัด", work_1_province: "เชียงใหม่", work_1_bizType: "ขายเครื่องใช้ไฟฟ้า", work_1_jobDesc: "ซ่อมเครื่องใช้ไฟฟ้าและช่วยงานช่าง", work_1_period: "สิงหาคม 2564 - มกราคม 2565", work_1_type: "intern_voc",
    work_2_company: "บริษัท ร่ำรวย888 จำกัด", work_2_province: "เชียงใหม่", work_2_bizType: "ผลิตและจำหน่ายเครื่องใช้ไฟฟ้า", work_2_jobDesc: "ซ่อมเครื่องใช้ไฟฟ้าต่างๆ เช่นแอร์ ทีวี เครื่องน้ำอุ่น", work_2_period: "มีนาคม 2567 - กันยายน 2567", work_2_type: "intern_hvoc",
    
    // Tab 6
    lifeGoal: "มีเงินเลี้ยงดูพ่อแม่",
    hobbies: "เล่นกีฬา , ดูหนัง , ฟังเพลง",
    strengths: "อารมณ์ดี หัวเราะง่าย",
    weaknesses: "ขี้อาย เข้าสังคมไม่เก่ง",
    specialSkills: "เล่นกีฬาได้หลากหลาย"
  };

  console.log("Starting testProcessForm...");
  const result = processForm(mockData);
  console.log("Result of testProcessForm: " + JSON.stringify(result));
}

/**
 * ดึงข้อมูลนักศึกษาทั้งหมดจาก Google Sheet เพื่อแสดงในหน้า Admin
 * เพิ่มการตรวจสอบรหัสผ่านและแปลงโครงสร้างข้อมูลประเภท วันที่ (Date) ป้องกันปัญหา JSON Serialization
 */
function getStudentSubmissions(pwd) {
  const inputPwd = (pwd || "").toString().trim();
  if (inputPwd !== CONFIG.ADMIN_PASSWORD) {
    throw new Error("สิทธิ์การเข้าถึงปฏิเสธ: รหัสผ่านแอดมินไม่ถูกต้อง");
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return [];
    
    // ตรวจสอบและซิงค์คอลัมน์ก่อนที่จะอ่านข้อมูล
    checkAndSyncHeaders(sheet);
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1) return []; // ไม่มีข้อมูลนอกจากหัวตาราง
    
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // แปลงแต่ละแถวข้อมูลเป็น Object ที่ระบุคีย์ตามหัวตาราง
    const submissions = dataRange.map((row, index) => {
      const obj = {};
      headers.forEach((header, idx) => {
        let val = row[idx];
        
        // แก้ไข Bug: หากข้อมูลเป็นประเภท Date วัตถุของ Google Apps Script จะส่งผ่าน API กลับไปตรงๆ ไม่ได้ (จะเกิดการล้มเหลว)
        // จำเป็นต้องใช้ Utilities.formatDate แปลงเป็นข้อความ String รูปแบบวันเวลาสากล
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
        obj[header] = val;
      });
      
      // Fallback สำหรับกรณีประวัติเก่ายังไม่มี studentId ให้กำหนดตามลำดับแถวชั่วคราว
      if (!obj.studentId) {
        obj.studentId = "OLD-" + (index + 1);
      }
      
      return obj;
    });
    
    return submissions;
  } catch (e) {
    console.error("Error fetching submissions: " + e.toString());
    return [];
  }
}

/**
 * บังคับสร้างไฟล์ PDF ใหม่ของนักศึกษาจากข้อมูลที่มีอยู่ในชีต
 */
function regenerateStudentPDF(pwd, studentId) {
  const inputPwd = (pwd || "").toString().trim();
  if (inputPwd !== CONFIG.ADMIN_PASSWORD) {
    throw new Error("สิทธิ์การเข้าถึงปฏิเสธ: รหัสผ่านแอดมินไม่ถูกต้อง");
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error("ไม่พบชีตข้อมูลนักศึกษา");
    
    // ตรวจสอบและซิงค์คอลัมน์ก่อนที่จะอ่านข้อมูล
    checkAndSyncHeaders(sheet);
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    // ค้นหานักศึกษาที่มี Student ID ตรงกัน
    let studentRowIndex = -1;
    let studentData = null;
    
    for (let i = 0; i < dataRange.length; i++) {
      const row = dataRange[i];
      const rowObj = {};
      headers.forEach((header, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
        rowObj[header] = val;
      });
      
      // กำหนด Fallback เช่นเดียวกันเพื่อความสอดคล้องในการดึงข้อมูลเก่า
      if (!rowObj.studentId) {
        rowObj.studentId = "OLD-" + (i + 1);
      }
      
      if (rowObj.studentId && studentId && rowObj.studentId.toString() === studentId.toString()) {
        studentRowIndex = i + 2; // +2 เพราะเป็น 1-indexed และเว้นบรรทัดหัวข้อ
        studentData = rowObj;
        break;
      }
    }
    
    if (!studentData) throw new Error("ไม่พบข้อมูลนักศึกษารหัสดังกล่าว");
    
    // จัดรูปแบบข้อมูลเพื่อให้เข้ากันได้กับ generatePDF
    // โดยเฉพาะพวกค่า 'ใช่' หรือ 'ไม่' ใน Sheet แปลงกลับเป็น Boolean
    for (let key in studentData) {
      if (studentData[key] === "ใช่") {
        studentData[key] = true;
      } else if (studentData[key] === "ไม่") {
        studentData[key] = false;
      }
    }
    
    const pdfUrl = generatePDF(studentData);
    
    // อัปเดตคอลัมน์ pdfUrl ในชีต
    const pdfColIdx = headers.indexOf('pdfUrl') + 1;
    if (pdfColIdx > 0) {
      sheet.getRange(studentRowIndex, pdfColIdx).setValue(pdfUrl);
    }
    
    return {
      success: true,
      url: pdfUrl
    };
  } catch (e) {
    return {
      success: false,
      message: e.toString()
    };
  }
}

/**
 * ฟังก์ชันสำหรับตรวจสอบความถูกต้องของรหัสผ่านผู้ดูแลระบบ (Admin)
 */
function verifyAdminPassword(pwd) {
  const ADMIN_PASSWORD = '1234'; // รหัสผ่านสำหรับระบบจัดการแอดมิน
  return pwd === ADMIN_PASSWORD;
}

/**
 * ดึง URL ของ Web App ปัจจุบัน เพื่อใช้ในการนำทางแบบสองทิศทางใน iframe
 */
function getScriptURL() {
  return ScriptApp.getService().getUrl();
}

/**
 * QA Test Runner - ตรวจสอบความถูกต้องและทดสอบระบบทั้งหมด
 */
function runQATests() {
  const results = [];
  
  // Test Case 1: Sheet Connection
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (sheet) {
      results.push({ name: "1. การเชื่อมต่อ Google Sheets", status: "PASS", message: "เชื่อมต่อได้ปกติ พบบัญชีหัวตาราง " + sheet.getLastColumn() + " คอลัมน์" });
    } else {
      results.push({ name: "1. การเชื่อมต่อ Google Sheets", status: "FAIL", message: "เชื่อมต่อกับ Spreadsheet ได้ แต่ไม่พบแท็บชีตชื่อ '" + CONFIG.SHEET_NAME + "'" });
    }
  } catch (e) {
    results.push({ name: "1. การเชื่อมต่อ Google Sheets", status: "FAIL", message: "ไม่สามารถเชื่อมต่อได้: " + e.toString() });
  }

  // Test Case 2: Template Doc Access
  try {
    const doc = DocumentApp.openById(CONFIG.TEMPLATE_DOC_ID);
    const length = doc.getBody().getText().length;
    results.push({ name: "2. การเชื่อมต่อเทมเพลต Google Docs", status: "PASS", message: "เชื่อมต่อและเปิดเทมเพลตได้สำเร็จ ความยาวเนื้อหา " + length + " ตัวอักษร" });
  } catch (e) {
    results.push({ name: "2. การเชื่อมต่อเทมเพลต Google Docs", status: "FAIL", message: "ไม่สามารถเปิดเทมเพลตได้: " + e.toString() });
  }

  // Test Case 3: Output Folder Access
  try {
    const folder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
    results.push({ name: "3. สิทธิ์การเข้าถึงโฟลเดอร์ปลายทาง (Google Drive)", status: "PASS", message: "เปิดโฟลเดอร์ได้ปกติ ชื่อโฟลเดอร์: " + folder.getName() });
  } catch (e) {
    results.push({ name: "3. สิทธิ์การเข้าถึงโฟลเดอร์ปลายทาง (Google Drive)", status: "FAIL", message: "ไม่สามารถเปิดโฟลเดอร์ปลายทางได้: " + e.toString() });
  }

  // Test Case 4: Authentication Check
  try {
    const checkTrue = verifyAdminPassword(CONFIG.ADMIN_PASSWORD);
    const checkFalse = verifyAdminPassword("wrong_pwd");
    if (checkTrue === true && checkFalse === false) {
      results.push({ name: "4. ระบบการตรวจสอบรหัสผ่านแอดมิน", status: "PASS", message: "ระบบรหัสผ่าน '" + CONFIG.ADMIN_PASSWORD + "' ทำงานถูกต้อง (ผ่านเมื่อถูก และปฏิเสธเมื่อผิด)" });
    } else {
      results.push({ name: "4. ระบบการตรวจสอบรหัสผ่านแอดมิน", status: "FAIL", message: "ผลตรวจสอบผิดพลาด: รหัสผ่านถูกต้องผ่าน=" + checkTrue + ", รหัสผิดผ่าน=" + checkFalse });
    }
  } catch (e) {
    results.push({ name: "4. ระบบการตรวจสอบรหัสผ่านแอดมิน", status: "FAIL", message: "ฟังก์ชันตรวจสอบรหัสทำงานผิดพลาด: " + e.toString() });
  }

  // Test Case 5: End-to-End Form Processing & PDF Generation
  let tempRowIndex = -1;
  let tempPdfFileId = "";
  try {
    // ใช้ข้อมูล Mock
    const mockData = {
      applyDate: "2026-06-02",
      applySource: "QA Testing System",
      applyBranch: "QA Branch",
      applyEduLevel: "voc",
      applyMajor: "QA Engineering",
      studentId: "QA-99999",
      titleTh: "นาย",
      fullNameTh: "ทดสอบ ระบบคิวเอ",
      nicknameTh: "คิวเอ",
      fullNameEn: "Mr. QA Testing",
      nicknameEn: "QA",
      birthDate: "2548-01-01",
      ageYears: 21,
      ageMonths: 5,
      maritalStatus: "single",
      childCount: 0,
      addressRegister: "ที่อยู่สำหรับทดสอบระบบ QA",
      addressCurrent: "ที่อยู่สำหรับทดสอบระบบ QA",
      phonePersonal: "0999999999",
      facebook: "QA.Facebook",
      lineId: "QA.Line",
      email: "qa@test.com",
      phoneParent: "0999999999",
      livingWith: "parents",
      parentRelationship: "บิดา",
      height: 180,
      weight: 75,
      bmi: 23.15,
      bodyShape: "normal",
      handedness: "right",
      illness_none: true,
      disease_none: true,
      lifeGoal: "QA Goal Test",
      hobbies: "QA Hobbies",
      strengths: "QA Strengths",
      weaknesses: "QA Weaknesses",
      specialSkills: "QA Skills"
    };

    const processResult = processForm(mockData);
    if (processResult.success && processResult.url) {
      // ดึง ID ของ PDF มาลบทีหลังเพื่อไม่ให้ขยะรก Drive
      const urlParts = processResult.url.split('/d/');
      if (urlParts.length > 1) {
        tempPdfFileId = urlParts[1].split('/')[0];
      }

      // ค้นหาแถวที่เพิ่มเข้าไปในชีตเพื่อทำการลบ
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      const lastRowVals = sheet.getRange(lastRow, 1, 1, lastCol).getValues()[0];
      
      // ตรวจสอบข้อมูลแถวสุดท้ายว่าเป็นของ QA จริงหรือไม่
      const idIdx = sheet.getRange(1, 1, 1, lastCol).getValues()[0].indexOf("studentId");
      if (idIdx >= 0 && lastRowVals[idIdx] === "QA-99999") {
        tempRowIndex = lastRow;
        // ลบแถวทดสอบออกทันที
        sheet.deleteRow(lastRow);
      }

      // ลบไฟล์ PDF ทดสอบ
      if (tempPdfFileId) {
        try {
          DriveApp.getFileById(tempPdfFileId).setTrashed(true);
        } catch (errPdf) {}
      }

      results.push({ name: "5. การบันทึกและสร้าง PDF (End-to-End)", status: "PASS", message: "สร้าง PDF สำเร็จ และเขียนข้อมูลลงแถวสุดท้ายใน Sheets สำเร็จ (ทำการเคลียร์ข้อมูล QA อัตโนมัติหลังทดสอบเรียบร้อย)" });
    } else {
      results.push({ name: "5. การบันทึกและสร้าง PDF (End-to-End)", status: "FAIL", message: "ประมวลผลไม่สำเร็จ: " + (processResult.message || "ไม่ได้รับ URL ไฟล์ PDF") });
    }
  } catch (e) {
    results.push({ name: "5. การบันทึกและสร้าง PDF (End-to-End)", status: "FAIL", message: "ล้มเหลวระหว่างประมวลผล: " + e.toString() });
  }

  // สร้างหน้าแสดงรายงานผล
  let html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>QA Test Runner</title><script src="https://cdn.tailwindcss.com"></script><link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;600;700&display=swap" rel="stylesheet"><style>body { font-family: "Sarabun", sans-serif; }</style></head><body class="bg-slate-100 p-6 min-h-screen"><div class="max-w-3xl mx-auto bg-white rounded-2xl shadow-xl overflow-hidden border border-slate-200"><div class="bg-gradient-to-r from-blue-700 to-indigo-800 text-white p-6"><h1 class="text-2xl font-bold">QA System Test Runner & Diagnostic</h1><p class="text-xs text-blue-100 mt-1">รันผลการตรวจสอบระบบเชื่อมต่อและการประมวลผลของ Apps Script ทั้งหมด</p></div><div class="p-6 space-y-4">';
  
  let allPass = true;
  results.forEach(r => {
    const isPass = r.status === "PASS";
    if (!isPass) allPass = false;
    
    html += '<div class="p-4 rounded-xl border ' + (isPass ? 'bg-emerald-50/50 border-emerald-200' : 'bg-red-50/50 border-red-200') + ' flex items-start gap-3">';
    html += '<span class="px-2.5 py-1 rounded text-xs font-bold ' + (isPass ? 'bg-emerald-500 text-white' : 'bg-red-500 text-white') + '">' + r.status + '</span>';
    html += '<div><h4 class="font-bold text-slate-800 text-sm">' + r.name + '</h4><p class="text-xs text-slate-650 mt-0.5">' + r.message + '</p></div></div>';
  });

  html += '<div class="mt-6 p-4 rounded-xl border text-center ' + (allPass ? 'bg-blue-50 border-blue-200 text-blue-800' : 'bg-amber-50 border-amber-200 text-amber-800') + '">';
  html += '<h3 class="font-bold text-base">' + (allPass ? '🎉 ทุกระบบทำงานได้อย่างถูกต้องสมบูรณ์แบบ!' : '⚠ มีบางระบบที่ยังทำงานไม่ถูกต้อง กรุณาตรวจสอบสิทธิ์หรือแก้ไข') + '</h3>';
  html += '<p class="text-xs mt-1">วันที่ตรวจสอบข้อมูล: ' + Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss") + '</p></div>';
  
  html += '<div class="mt-4 text-center flex flex-col sm:flex-row items-center justify-center gap-3">';
  html += '<a href="?mode=qa&action=create_mock" class="bg-amber-600 hover:bg-amber-700 text-white font-bold py-2 px-4 rounded-xl text-xs transition shadow-md flex items-center gap-1">⚡ สร้างข้อมูลจำลอง 3 รายการ</a>';
  html += '<div class="text-xs text-slate-400">|</div>';
  html += '<a href="?mode=admin" class="text-blue-650 hover:underline text-xs">เข้าสู่หน้าแอดมิน &rarr;</a>';
  html += '<div class="text-xs text-slate-400">|</div>';
  html += '<a href="?" class="text-blue-650 hover:underline text-xs">&larr; กลับหน้าฟอร์มหลัก</a></div>';
  html += '</div></div></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('QA Test Runner')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * สร้างข้อมูลจำลอง 3 รายการเพื่อใช้ในการทดสอบระบบ Admin
 */
function createThreeMockSubmissions() {
  const mockStudents = [
    {
      applyDate: "2026-06-02",
      applySource: "ระบบทดสอบ QA",
      applyBranch: "เชียงใหม่",
      applyEduLevel: "hvoc",
      applyMajor: "วิศวกรรมไฟฟ้า",
      studentId: "STD-202601",
      titleTh: "นาย",
      fullNameTh: "มิ่งขวัญ ปัญญาดี",
      nicknameTh: "ขวัญ",
      fullNameEn: "Mr. Mingkhwan Panyadee",
      nicknameEn: "Khwan",
      birthDate: "2004-05-15",
      ageYears: 22,
      ageMonths: 1,
      maritalStatus: "single",
      childCount: 0,
      addressRegister: "123 ม.4 ต.ช้างเผือก อ.เมือง จ.เชียงใหม่ 50300",
      addressCurrent: "123 ม.4 ต.ช้างเผือก อ.เมือง จ.เชียงใหม่ 50300",
      phonePersonal: "0811112222",
      facebook: "Mingkhwan.P",
      lineId: "khwan_p",
      email: "mingkhwan@example.com",
      phoneParent: "0811113333",
      livingWith: "parents",
      parentRelationship: "บิดา",
      height: 172,
      weight: 65,
      bmi: "21.97",
      bodyShape: "normal",
      handedness: "right",
      eyesightLeftVal: "1.0",
      eyesightLeftStatus: "normal",
      eyesightRightVal: "1.0",
      eyesightRightStatus: "normal",
      glassesLeft: false,
      contactsLeft: false,
      glassesRight: false,
      contactsRight: false,
      colorBlindness: "ไม่มี",
      tattoo: "ไม่มี",
      militaryStatus: "exempted",
      nameChanged: "ไม่เคย",
      illness_none: true,
      disease_none: true,
      bloodGroup: "O",
      plasticSurgery: "ไม่เคย",
      dentalBraces: "ไม่",
      smoke: "ไม่",
      vape: "ไม่",
      cannabis: "ไม่",
      kratom: "ไม่",
      alcohol: "ไม่",
      driverLicense: "มี",
      licenseCar: "ใช่",
      licenseMotorcycle: "ใช่",
      passport: "ไม่มี",
      japanVisit: "ไม่เคย",
      japanApply: "none",
      criminalRecord: "ไม่มี",
      studentDebt: "ไม่มี",
      parentDebt: "ไม่มี",
      familyStatus: "together",
      guardian: "บิดา",
      japanContact: "none",
      edu_m3_school: "โรงเรียนบ้านช้างเผือก", edu_m3_major: "ทั่วไป", edu_m3_start_year: "2560", edu_m3_start_date: "2560-05-16", edu_m3_end_year: "2563", edu_m3_end_date: "2563-03-15",
      edu_voc_school: "เทคนิคเชียงใหม่", edu_voc_major: "ไฟฟ้ากำลัง", edu_voc_start_year: "2563", edu_voc_start_date: "2563-05-16", edu_voc_end_year: "2566", edu_voc_end_date: "2566-03-15",
      edu_hvoc_school: "เทคนิคเชียงใหม่", edu_hvoc_major: "ไฟฟ้าควบคุม", edu_hvoc_start_year: "2566", edu_hvoc_start_date: "2566-06-01", edu_hvoc_end_year: "2568", edu_hvoc_end_date: "2568-03-15",
      eduInterrupted: "none",
      fam_name_1: "สมปอง ปัญญาดี", fam_rel_1: "บิดา", fam_age_1: "55", fam_job_1: "รับจ้างทั่วไป", fam_work_1: "โรงสีข้าวเชียงใหม่", fam_inc_1: "15000", fam_prov_1: "เชียงใหม่",
      fam_name_2: "สมศรี ปัญญาดี", fam_rel_2: "มารดา", fam_age_2: "52", fam_job_2: "แม่บ้าน", fam_work_2: "บริษัททำความสะอาด", fam_inc_2: "12000", fam_prov_2: "เชียงใหม่",
      fam_name_3: "สมหมาย ปัญญาดี", fam_rel_3: "พี่ชาย", fam_age_3: "28", fam_job_3: "พนักงานบริษัท", fam_work_3: "CP All", fam_inc_3: "25000", fam_prov_3: "ลำพูน",
      fam_name_4: "สมหญิง ปัญญาดี", fam_rel_4: "พี่สาว", fam_age_4: "26", fam_job_4: "พยาบาล", fam_work_4: "รพ.มหาราช", fam_inc_4: "30000", fam_prov_4: "เชียงใหม่",
      fam_name_5: "สมชาย ปัญญาดี", fam_rel_5: "น้องชาย", fam_age_5: "18", fam_job_5: "นักศึกษา", fam_work_5: "ม.เชียงใหม่", fam_inc_5: "-", fam_prov_5: "เชียงใหม่",
      fam_name_6: "สมรักษ์ ปัญญาดี", fam_rel_6: "น้องชาย", fam_age_6: "15", fam_job_6: "นักเรียน", fam_work_6: "ร.ร.ยุพราช", fam_inc_6: "-", fam_prov_6: "เชียงใหม่",
      fam_name_7: "สมจิตร ปัญญาดี", fam_rel_7: "ลุง", fam_age_7: "60", fam_job_7: "เกษตรกร", fam_work_7: "สวนลำไย", fam_inc_7: "20000", fam_prov_7: "เชียงใหม่",
      fam_name_8: "สมใจ ปัญญาดี", fam_rel_8: "ป้า", fam_age_8: "58", fam_job_8: "ค้าขาย", fam_work_8: "ตลาดสด", fam_inc_8: "18000", fam_prov_8: "เชียงใหม่",
      work_1_company: "เชียงใหม่การช่าง", work_1_province: "เชียงใหม่", work_1_bizType: "รับเหมาไฟฟ้า", work_1_jobDesc: "ผู้ช่วยช่างไฟ", work_1_salary: "12,000", work_1_period: "มิ.ย.66 - ธ.ค.66", work_1_type: "parttime",
      work_2_company: "ลำพูนอิเล็กทรอนิกส์", work_2_province: "ลำพูน", work_2_bizType: "โรงงานอุตสาหกรรม", work_2_jobDesc: "ฝ่ายผลิต", work_2_salary: "15,000", work_2_period: "ม.ค.67 - เม.ย.67", work_2_type: "work",
      work_3_company: "การไฟฟ้าส่วนภูมิภาค", work_3_province: "เชียงใหม่", work_3_bizType: "รัฐวิสาหกิจ", work_3_jobDesc: "นักศึกษาฝึกงาน", work_3_salary: "3,000", work_3_period: "พ.ค.67 - ส.ค.67", work_3_type: "intern_hvoc",
      work_4_company: "บริษัท สมาร์ทโฮม จำกัด", work_4_province: "เชียงใหม่", work_4_bizType: "รับติดตั้งระบบบ้านอัจฉริยะ", work_4_jobDesc: "ช่างติดตั้ง", work_4_salary: "14,000", work_4_period: "ก.ย.67 - ต.ค.67", work_4_type: "work",
      work_5_company: "ร้านเกมส์ E-Sport", work_5_province: "เชียงใหม่", work_5_bizType: "บริการ", work_5_jobDesc: "พนักงานดูแลร้าน", work_5_salary: "9,000", work_5_period: "ช่วงปิดเทอม ปี 65", work_5_type: "break",
      work_6_company: "เซเว่น อีเลฟเว่น", work_6_province: "เชียงใหม่", work_6_bizType: "ค้าปลีก", work_6_jobDesc: "พนักงานพาร์ทไทม์", work_6_salary: "45 บาท/ชม.", work_6_period: "ม.ค.65 - มี.ค.65", work_6_type: "parttime",
      work_7_company: "บิ๊กซี ซูเปอร์เซ็นเตอร์", work_7_province: "เชียงใหม่", work_7_bizType: "ค้าปลีก", work_7_jobDesc: "พนักงานจัดเรียงสินค้า", work_7_salary: "10,000", work_7_period: "เม.ย.65 - ส.ค.65", work_7_type: "work",
      work_8_company: "ร้านกาแฟ อเมซอน", work_8_province: "ลำพูน", work_8_bizType: "ร้านอาหารและเครื่องดื่ม", work_8_jobDesc: "บาริสต้าชั่วคราว", work_8_salary: "11,000", work_8_period: "ต.ค.65 - ธ.ค.65", work_8_type: "parttime",
      work_9_company: "ศูนย์ซ่อมคอมพิวเตอร์ IT Center", work_9_province: "เชียงใหม่", work_9_bizType: "ซ่อมบำรุง IT", work_9_jobDesc: "ช่างซ่อมคอมพิวเตอร์", work_9_salary: "13,500", work_9_period: "ม.ค.66 - พ.ค.66", work_9_type: "work",
      lifeGoal: "อยากทำงานที่มั่นคงและมีเงินส่งเสียครอบครัว",
      hobbies: "เล่นฟุตบอล, อ่านหนังสือ",
      strengths: "มีความอดทนและตั้งใจทำงานสูง",
      weaknesses: "ภาษาอังกฤษยังไม่แข็งแรง",
      specialSkills: "ซ่อมเครื่องใช้ไฟฟ้าพื้นฐาน"
    },
    {
      applyDate: "2026-06-02",
      applySource: "ระบบทดสอบ QA",
      applyBranch: "กรุงเทพฯ",
      applyEduLevel: "uni",
      applyMajor: "บัญชีและการเงิน",
      studentId: "STD-202602",
      titleTh: "นางสาว",
      fullNameTh: "กิ่งแก้ว พวงเงิน",
      nicknameTh: "แก้ว",
      fullNameEn: "Miss Kingkaew Phuangngoen",
      nicknameEn: "Kaew",
      birthDate: "2003-11-20",
      ageYears: 22,
      ageMonths: 6,
      maritalStatus: "single",
      childCount: 0,
      addressRegister: "456 ต.ศรีภูมิ อ.เมือง จ.เชียงใหม่ 50200",
      addressCurrent: "456 ต.ศรีภูมิ อ.เมือง จ.เชียงใหม่ 50200",
      phonePersonal: "0823334444",
      facebook: "Kingkaew.K",
      lineId: "kaew_k",
      email: "kingkaew@example.com",
      phoneParent: "0823335555",
      livingWith: "parents",
      parentRelationship: "มารดา",
      height: 160,
      weight: 48,
      bmi: "18.75",
      bodyShape: "normal",
      handedness: "right",
      eyesightLeftVal: "0.8",
      eyesightLeftStatus: "normal",
      eyesightRightVal: "0.8",
      eyesightRightStatus: "normal",
      glassesLeft: true,
      contactsLeft: false,
      glassesRight: true,
      contactsRight: false,
      colorBlindness: "ไม่มี",
      tattoo: "ไม่มี",
      militaryStatus: "exempted",
      nameChanged: "ไม่เคย",
      illness_none: true,
      disease_none: true,
      bloodGroup: "B",
      plasticSurgery: "ไม่เคย",
      dentalBraces: "ไม่",
      smoke: "ไม่",
      vape: "ไม่",
      cannabis: "ไม่",
      kratom: "ไม่",
      alcohol: "ไม่",
      driverLicense: "ไม่มี",
      passport: "มี",
      passportReason: "ท่องเที่ยว",
      passportCountry: "เวียดนาม",
      passportTime: "1 ครั้ง",
      passportExpireYear: "2570",
      japanVisit: "ไม่เคย",
      japanApply: "none",
      criminalRecord: "ไม่มี",
      studentDebt: "ไม่มี",
      parentDebt: "ไม่มี",
      familyStatus: "together",
      guardian: "มารดา",
      japanContact: "none",
      edu_m3_school: "โรงเรียนสตรีเชียงใหม่", edu_m3_major: "ศิลป์-คำนวณ", edu_m3_start_year: "2559", edu_m3_start_date: "2559-05-16", edu_m3_end_year: "2562", edu_m3_end_date: "2562-03-15",
      edu_voc_school: "พณิชยการเชียงใหม่", edu_voc_major: "การบัญชี", edu_voc_start_year: "2562", edu_voc_start_date: "2562-05-16", edu_voc_end_year: "2565", edu_voc_end_date: "2565-03-15",
      edu_uni_school: "มหาวิทยาลัยเชียงใหม่", edu_uni_major: "บัญชี", edu_uni_start_year: "2565", edu_uni_start_date: "2565-06-01", edu_uni_end_year: "2569", edu_uni_end_date: "2569-03-15",
      eduInterrupted: "none",
      work_1_company: "สำนักงานบัญชีสุขใจ", work_1_province: "เชียงใหม่", work_1_bizType: "สำนักงานบัญชี", work_1_jobDesc: "พนักงานฝึกหัดทำบัญชี", work_1_salary: "15,000", work_1_period: "3 เดือน", work_1_type: "intern_uni",
      lifeGoal: "ต้องการเป็นผู้สอบบัญชีรับอนุญาต (CPA)",
      hobbies: "ฟังเพลง, ดูซีรีส์",
      strengths: "ทำงานละเอียด รอบคอบ และตรงต่อเวลา",
      weaknesses: "ค่อนข้างพูดไม่เก่งเมื่ออยู่กับคนแปลกหน้า",
      specialSkills: "ใช้โปรแกรมบัญชีสำเร็จรูป Express ได้ดี"
    },
    {
      applyDate: "2026-06-02",
      applySource: "ระบบทดสอบ QA",
      applyBranch: "เชียงใหม่",
      applyEduLevel: "voc",
      applyMajor: "เทคโนโลยีสารสนเทศ",
      studentId: "STD-202603",
      titleTh: "นาย",
      fullNameTh: "ชัยชนะ ภักดีชน",
      nicknameTh: "ชัย",
      fullNameEn: "Mr. Chaichana Phakdichon",
      nicknameEn: "Chai",
      birthDate: "2006-08-05",
      ageYears: 19,
      ageMonths: 10,
      maritalStatus: "single",
      childCount: 0,
      addressRegister: "789 ต.สุเทพ อ.เมือง จ.เชียงใหม่ 50200",
      addressCurrent: "789 ต.สุเทพ อ.เมือง จ.เชียงใหม่ 50200",
      phonePersonal: "0835556666",
      facebook: "Chaichana.P",
      lineId: "chai_p",
      email: "chaichana@example.com",
      phoneParent: "0835557777",
      livingWith: "parents",
      parentRelationship: "บิดา",
      height: 178,
      weight: 70,
      bmi: "22.09",
      bodyShape: "normal",
      handedness: "right",
      eyesightLeftVal: "1.2",
      eyesightLeftStatus: "normal",
      eyesightRightVal: "1.2",
      eyesightRightStatus: "normal",
      glassesLeft: false,
      contactsLeft: false,
      glassesRight: false,
      contactsRight: false,
      colorBlindness: "ไม่มี",
      tattoo: "ไม่มี",
      militaryStatus: "active",
      militaryYear: "2570",
      nameChanged: "ไม่เคย",
      illness_none: true,
      disease_none: true,
      bloodGroup: "A",
      plasticSurgery: "ไม่เคย",
      dentalBraces: "ไม่",
      smoke: "ไม่",
      vape: "ไม่",
      cannabis: "ไม่",
      kratom: "ไม่",
      alcohol: "ไม่",
      driverLicense: "มี",
      licenseMotorcycle: "ใช่",
      passport: "ไม่มี",
      japanVisit: "ไม่เคย",
      japanApply: "none",
      criminalRecord: "ไม่มี",
      studentDebt: "ไม่มี",
      parentDebt: "ไม่มี",
      familyStatus: "together",
      guardian: "บิดา",
      japanContact: "none",
      edu_m3_school: "โรงเรียนบ้านสุเทพ", edu_m3_major: "ทั่วไป", edu_m3_start_year: "2562", edu_m3_start_date: "2562-05-16", edu_m3_end_year: "2565", edu_m3_end_date: "2565-03-15",
      edu_voc_school: "โปลิเทคนิคลานนา", edu_voc_major: "ไอที/คอมพิวเตอร์", edu_voc_start_year: "2565", edu_voc_start_date: "2565-05-16", edu_voc_end_year: "2568", edu_voc_end_date: "2568-03-15",
      eduInterrupted: "none",
      lifeGoal: "อยากเป็นโปรแกรมเมอร์และทำสตาร์ทอัพของตัวเอง",
      hobbies: "เขียนโค้ดเล่น, เล่นเกมคอมพิวเตอร์",
      strengths: "ชอบเรียนรู้เทคโนโลยีใหม่ๆ อยู่เสมอ",
      weaknesses: "บางครั้งทำงานจนลืมพักผ่อน",
      specialSkills: "เขียนภาษา HTML, CSS, JavaScript เบื้องต้นได้"
    }
  ];

  const results = [];
  mockStudents.forEach(student => {
    const res = processForm(student);
    results.push(res);
  });
  return results;
}

/**
 * ฟังก์ชันตรวจสอบและวินิจฉัยปัญหาการสร้างไฟล์ PDF
 */
function runPdfDiagnostic() {
  const result = {
    step1_templateDoc: {},
    step2_destinationFolder: {},
    step3_copyDoc: {},
    step4_pdfExport: {},
    errors: []
  };
  
  // 1. Check template
  try {
    const docFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
    result.step1_templateDoc.name = docFile.getName();
    result.step1_templateDoc.size = docFile.getSize();
    
    const doc = DocumentApp.openById(CONFIG.TEMPLATE_DOC_ID);
    const text = doc.getBody().getText();
    result.step1_templateDoc.textLength = text.length;
    result.step1_templateDoc.hasContent = text.length > 50;
    result.step1_templateDoc.sample = text.substring(0, 100);
  } catch (e) {
    result.errors.push("Step 1 Template failed: " + e.toString());
  }
  
  // 2. Check Folder
  try {
    const folder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
    result.step2_destinationFolder.name = folder.getName();
    result.step2_destinationFolder.sharingAccess = folder.getSharingAccess().toString();
  } catch (e) {
    result.errors.push("Step 2 Folder failed: " + e.toString());
  }
  
  // 3. Try copy
  let tempDocFile;
  try {
    const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
    const folder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);
    tempDocFile = templateFile.makeCopy("Temp_Diagnostic_Test", folder);
    
    result.step3_copyDoc.copiedFileId = tempDocFile.getId();
    result.step3_copyDoc.copiedFileName = tempDocFile.getName();
    
    const tempDoc = DocumentApp.openById(tempDocFile.getId());
    const body = tempDoc.getBody();
    
    result.step3_copyDoc.textLengthBefore = body.getText().length;
    
    // Replace something dummy
    body.replaceText("<<fullNameTh>>", "ทดสอบ ระบบดีบั๊ก");
    tempDoc.saveAndClose();
    
    Utilities.sleep(1000);
    
    const tempDocVerify = DocumentApp.openById(tempDocFile.getId());
    result.step3_copyDoc.textLengthAfter = tempDocVerify.getBody().getText().length;
  } catch (e) {
    result.errors.push("Step 3 Copy failed: " + e.toString());
  }
  
  // 4. Try export
  if (tempDocFile) {
    const docId = tempDocFile.getId();
    
    // Test UrlFetch
    try {
      const url = `https://docs.google.com/document/d/${docId}/export?format=pdf`;
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      
      const contentType = response.getHeaders()['Content-Type'] || '';
      result.step4_pdfExport.urlFetchResponseCode = response.getResponseCode();
      result.step4_pdfExport.urlFetchContentType = contentType;
      
      if (response.getResponseCode() === 200) {
        const blob = response.getBlob();
        result.step4_pdfExport.urlFetchBlobSize = blob.getBytes().length;
      }
    } catch (e) {
      result.step4_pdfExport.urlFetchError = e.toString();
    }
    
    // Test DriveApp
    try {
      const blob = DriveApp.getFileById(docId).getBlob().getAs(MimeType.PDF);
      result.step4_pdfExport.driveAppBlobSize = blob.getBytes().length;
      
      const testPdfFile = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID).createFile(blob).setName("Diagnostic_Test.pdf");
      testPdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      result.step4_pdfExport.pdfFileUrl = testPdfFile.getUrl();
    } catch (e) {
      result.step4_pdfExport.driveAppError = e.toString();
    }
    
    // Cleanup temp doc
    try {
      tempDocFile.setTrashed(true);
      result.step3_copyDoc.cleanedUp = true;
    } catch (e) {
      result.step3_copyDoc.cleanupError = e.toString();
    }
  }
  
  return result;
}

/**
 * ฟังก์ชันสำหรับออกแบบและสร้างเทมเพลต Google Docs จากศูนย์ให้มีฟอร์มเหมือน PDF ต้นฉบับ
 */
function createTemplateDoc() {
  const doc = DocumentApp.openById(CONFIG.TEMPLATE_DOC_ID);
  const body = doc.getBody();
  body.clear();
  
  // ตั้งระยะขอบ (Margins)
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(36);
  body.setMarginRight(36);
  
  // ตัวช่วยตั้งฟอนต์ย่อหน้า
  function styleParagraph(p, size, bold, align, color) {
    p.setFontFamily('Sarabun');
    p.setFontSize(size);
    p.setBold(bold);
    if (align) p.setAlignment(align);
    if (color) p.setForegroundColor(color);
    return p;
  }

  // ตัวช่วยจัดสไตล์เซลล์และไฮไลต์แท็ก <<...>> เป็นสีแดง
  function formatCellText(cell, text, isLabel, align, size, bg) {
    cell.setFontFamily('Sarabun');
    cell.setFontSize(size || 8);
    cell.setBold(isLabel ? true : false);
    
    if (text !== undefined && text !== null) {
      cell.clear();
      const p = cell.getChild(0).asParagraph();
      if (text !== "") {
        p.setText(text);
      } else {
        p.setText(" "); // Prevent empty text error
      }
      if (align) p.setAlignment(align);
      p.setSpacingAfter(0);
      p.setSpacingBefore(0);
      p.setLineSpacing(1.05);
      
      const textObj = p.editAsText();
      if (textObj && textObj.getText().length > 0) {
        // ตั้งค่าสีพื้นฐานเป็นสีดำ
        textObj.setForegroundColor('#000000');
        
        // ค้นหาแท็ก <<...>> และเปลี่ยนสีเป็นสีแดง
        const textVal = textObj.getText();
        const regex = /<<.*?>>/g;
        let match;
        while ((match = regex.exec(textVal)) !== null) {
          const startIdx = match.index;
          const endIdx = regex.lastIndex - 1;
          textObj.setForegroundColor(startIdx, endIdx, '#dc2626');
        }
      }
    }
    
    cell.setPaddingTop(3);
    cell.setPaddingBottom(3);
    cell.setPaddingLeft(4);
    cell.setPaddingRight(4);
    
    // ตั้งค่าสีพื้นหลัง:
    // 1. ใช้ bg หากระบุ
    // 2. หากไม่ระบุ ให้เช็ค isLabel ถ้าเป็นหัวข้อ/ป้ายกำกับให้ใช้สีเทาอ่อน #f8fafc ถ้าไม่ใช่ใช้สีขาว #ffffff
    const bgColor = bg || (isLabel ? '#f8fafc' : '#ffffff');
    cell.setBackgroundColor(bgColor);
  }

  // จัดสไตล์ตารางทั่วไปแบบมีขอบดำ
  function styleTableBorder(table, borderHex) {
    table.setBorderColor(borderHex || '#000000');
    table.setBorderWidth(1.5);
  }

  // ตัวช่วยสร้างระยะห่างย่อหน้าขนาดเล็ก
  function addSpacer(body, size) {
    const p = body.appendParagraph("");
    p.setFontSize(size || 3);
    p.setSpacingAfter(0);
    p.setSpacingBefore(0);
  }

  // ตัวช่วยสร้างบัตรประวัติการทำงาน (Job Card) แบบกริด
  function createWorkCard(body, i) {
    const outerTable = body.appendTable();
    outerTable.setBorderColor('#000000');
    outerTable.setBorderWidth(1);
    
    const row = outerTable.appendTableRow();
    const cellLard = row.appendTableCell(i.toString());
    cellLard.setWidth(30);
    formatCellText(cellLard, i.toString(), true, DocumentApp.HorizontalAlignment.CENTER, 9.5);
    
    const cellContent = row.appendTableCell();
    cellContent.setWidth(470);
    cellContent.setPaddingTop(0); cellContent.setPaddingBottom(0); cellContent.setPaddingLeft(0); cellContent.setPaddingRight(0);
    
    // Nested inner table
    const inner = cellContent.appendTable();
    inner.setBorderColor('#000000');
    inner.setBorderWidth(0.5);
    
    // Row 1: ชื่อบริษัท & จังหวัด
    const r1 = inner.appendTableRow();
    const r1_l1 = r1.appendTableCell("ชื่อบริษัท"); r1_l1.setWidth(80);
    const r1_v1 = r1.appendTableCell(`<<work_${i}_company>>`); r1_v1.setWidth(180);
    const r1_l2 = r1.appendTableCell("จังหวัด"); r1_l2.setWidth(60);
    const r1_v2 = r1.appendTableCell(`<<work_${i}_province>>`); r1_v2.setWidth(150);
    
    // Row 2: ประเภทกิจการ
    const r2 = inner.appendTableRow();
    const r2_l1 = r2.appendTableCell("ประเภทกิจการ"); r2_l1.setWidth(80);
    const r2_v1 = r2.appendTableCell(`<<work_${i}_bizType>>`); r2_v1.setWidth(390);
    r2.appendTableCell("");
    r2.appendTableCell("");
    r2.getCell(3).merge();
    r2.getCell(2).merge();
    
    // Row 3: ลักษณะงานที่ทำ & เงินเดือน
    const r3 = inner.appendTableRow();
    const r3_l1 = r3.appendTableCell("ลักษณะงานที่ทำ / ตำแหน่ง"); r3_l1.setWidth(80);
    const r3_v1 = r3.appendTableCell(`<<work_${i}_jobDesc>>`); r3_v1.setWidth(180);
    const r3_l2 = r3.appendTableCell("เงินเดือน"); r3_l2.setWidth(60);
    const r3_v2 = r3.appendTableCell(`<<work_${i}_salary>>`); r3_v2.setWidth(150);
    
    // Row 4: เริ่มงาน-ลาออก
    const r4 = inner.appendTableRow();
    const r4_l1 = r4.appendTableCell("เริ่มงาน-ลาออก"); r4_l1.setWidth(80);
    const r4_v1 = r4.appendTableCell(`<<work_${i}_period>>`); r4_v1.setWidth(390);
    r4.appendTableCell("");
    r4.appendTableCell("");
    r4.getCell(3).merge();
    r4.getCell(2).merge();
    
    // Row 5: ประเภทงาน
    const r5 = inner.appendTableRow();
    const r5_l1 = r5.appendTableCell("ประเภทงาน"); r5_l1.setWidth(80);
    const r5_v1 = r5.appendTableCell(
      `<<work_${i}_type_work>> ทำงานประจำ   <<work_${i}_type_parttime>> พาร์ทไทม์   <<work_${i}_type_intern_voc>> ฝึกงาน ปวช.\n<<work_${i}_type_intern_hvoc>> ฝึกงาน ปวส.   <<work_${i}_type_intern_uni>> ฝึกงาน ป.ตรี   <<work_${i}_type_break>> ช่วงปิดเทอม   <<work_${i}_type_other>> อื่นๆ`
    ); r5_v1.setWidth(390);
    r5.appendTableCell("");
    r5.appendTableCell("");
    r5.getCell(3).merge();
    r5.getCell(2).merge();
    
    // Set widths & style for inner cells
    for (let r = 0; r < inner.getNumRows(); r++) {
      const rowObj = inner.getRow(r);
      for (let c = 0; c < rowObj.getNumCells(); c++) {
        const cell = rowObj.getCell(c);
        const isLabel = (c === 0) || (r === 0 && c === 2) || (r === 2 && c === 2);
        const textVal = cell.getText();
        formatCellText(cell, textVal, isLabel, DocumentApp.HorizontalAlignment.LEFT, 7.5);
      }
    }
    
    if (cellContent.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
      cellContent.removeChild(cellContent.getChild(0));
    }
    
    // Add spacer paragraph after job card
    addSpacer(body, 3);
  }

  // === PAGE 1 ===
  // Header Table (Invisible borders)
  const headerTable = body.appendTable();
  headerTable.setBorderWidth(0);
  headerTable.setBorderColor('#ffffff');
  const headerRow = headerTable.appendTableRow();
  
  const leftCell = headerRow.appendTableCell();
  leftCell.setWidth(210);
  const midCell = headerRow.appendTableCell();
  midCell.setWidth(100);
  const rightCell = headerRow.appendTableCell();
  rightCell.setWidth(190);
  
  // Left Cell Content: Logo and Title
  const pLogo = leftCell.appendParagraph("AIU | 大 shogun Project\nโครงการไดโชกุน");
  styleParagraph(pLogo, 9.5, true, DocumentApp.HorizontalAlignment.CENTER, '#000000');
  pLogo.setSpacingAfter(4);
  
  const pTitle = leftCell.appendParagraph("ใบประวัตินักศึกษา");
  styleParagraph(pTitle, 16, true, DocumentApp.HorizontalAlignment.CENTER, '#000000');
  pTitle.setSpacingAfter(8);
  
  const pMeta1 = leftCell.appendParagraph("วันที่สมัคร..................... <<applyDate>>");
  const pMeta2 = leftCell.appendParagraph("ทราบข้อมูลการสมัครจาก..................... <<applySource>>");
  const pMeta3 = leftCell.appendParagraph("สำนักงานสาขาที่รับสมัคร..................... <<applyBranch>>");
  [pMeta1, pMeta2, pMeta3].forEach(p => {
    styleParagraph(p, 9, false, DocumentApp.HorizontalAlignment.LEFT, '#000000');
    p.setSpacingAfter(3);
    // เปลี่ยนสีเฉพาะแท็กในย่อหน้า
    const textObj = p.editAsText();
    if (textObj) {
      const textVal = textObj.getText();
      const regex = /<<.*?>>/g;
      let match;
      while ((match = regex.exec(textVal)) !== null) {
        textObj.setForegroundColor(match.index, regex.lastIndex - 1, '#dc2626');
      }
    }
  });
  
  // Mid Cell Content: Photo Box
  const photoTable = midCell.appendTable();
  photoTable.setBorderColor('#000000');
  photoTable.setBorderWidth(1);
  const photoInnerRow = photoTable.appendTableRow();
  const photoInnerCell = photoInnerRow.appendTableCell();
  photoInnerCell.setWidth(90);
  photoInnerRow.setMinimumHeight(115);
  const pPhoto = photoInnerCell.appendParagraph("\n\nรูปถ่าย\nขนาด 1 นิ้ว\n(ติดหนีบคลิป)");
  styleParagraph(pPhoto, 8, false, DocumentApp.HorizontalAlignment.CENTER, '#000000');
  pPhoto.setSpacingAfter(0);
  
  // Right Cell Content: สำหรับเจ้าหน้าที่ table
  const staffTable = rightCell.appendTable();
  staffTable.setBorderColor('#000000');
  staffTable.setBorderWidth(1);

  // Row 0: "สำหรับเจ้าหน้าที่" (span 3)
  const sr0 = staffTable.appendTableRow();
  const sc0_0 = sr0.appendTableCell("สำหรับเจ้าหน้าที่");
  sc0_0.setWidth(180);
  const sc0_1 = sr0.appendTableCell("");
  const sc0_2 = sr0.appendTableCell("");
  formatCellText(sc0_0, "สำหรับเจ้าหน้าที่", true, DocumentApp.HorizontalAlignment.CENTER, 8, '#e2e8f0');

  // Row 1: "รหัสนักศึกษา : <<studentId>>" (span 3)
  const sr1 = staffTable.appendTableRow();
  const sc1_0 = sr1.appendTableCell("รหัสนักศึกษา : <<studentId>>");
  sc1_0.setWidth(180);
  const sc1_1 = sr1.appendTableCell("");
  const sc1_2 = sr1.appendTableCell("");
  formatCellText(sc1_0, "รหัสนักศึกษา : <<studentId>>", false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 2: Two cells: Cell 0 (span 2) and Cell 1 (span 1)
  const sr2 = staffTable.appendTableRow();
  const sc2_0 = sr2.appendTableCell("<<applySource_industry>> อุตสาหกรรม   <<applySource_service>> ธุรกิจการบริการ   <<applySource_elderly>> การจัดการสุขภาพผู้สูงอายุ");
  sc2_0.setWidth(120);
  const sc2_1 = sr2.appendTableCell("");
  const sc2_2 = sr2.appendTableCell("<<applyBranch_regular>> ปกติ   <<applyBranch_dvc>> ทวิภาคี   <<applyBranch_intern>> ฝึกงาน");
  sc2_2.setWidth(60);
  formatCellText(sc2_0, sc2_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 6.5);
  formatCellText(sc2_2, sc2_2.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 6.5);

  // Row 3: "กำหนดสอบทุน..............................................." (span 3)
  const sr3 = staffTable.appendTableRow();
  const sc3_0 = sr3.appendTableCell("กำหนดสอบทุน...............................................");
  sc3_0.setWidth(180);
  const sc3_1 = sr3.appendTableCell("");
  const sc3_2 = sr3.appendTableCell("");
  formatCellText(sc3_0, sc3_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 4: "ประชุมผู้ปกครองวันที่ : ....................................................." (span 3)
  const sr4 = staffTable.appendTableRow();
  const sc4_0 = sr4.appendTableCell("ประชุมผู้ปกครองวันที่ : .....................................................");
  sc4_0.setWidth(180);
  const sc4_1 = sr4.appendTableCell("");
  const sc4_2 = sr4.appendTableCell("");
  formatCellText(sc4_0, sc4_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 5: "ผู้เข้าร่วมประชุม : .............................................................." (span 3)
  const sr5 = staffTable.appendTableRow();
  const sc5_0 = sr5.appendTableCell("ผู้เข้าร่วมประชุม : ..............................................................");
  sc5_0.setWidth(180);
  const sc5_1 = sr5.appendTableCell("");
  const sc5_2 = sr5.appendTableCell("");
  formatCellText(sc5_0, sc5_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 6: "ตรวจสุขภาพวันที่ : .............................................................." (span 3)
  const sr6 = staffTable.appendTableRow();
  const sc6_0 = sr6.appendTableCell("ตรวจสุขภาพวันที่ : ..............................................................");
  sc6_0.setWidth(180);
  const sc6_1 = sr6.appendTableCell("");
  const sc6_2 = sr6.appendTableCell("");
  formatCellText(sc6_0, sc6_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 7: "ชื่อโรงพยาบาล : ..............................................................." (span 3)
  const sr7 = staffTable.appendTableRow();
  const sc7_0 = sr7.appendTableCell("ชื่อโรงพยาบาล : ...............................................................");
  sc7_0.setWidth(180);
  const sc7_1 = sr7.appendTableCell("");
  const sc7_2 = sr7.appendTableCell("");
  formatCellText(sc7_0, sc7_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 8: "ผลตรวจสุขภาพ   ☐ ผ่าน   ☐ ไม่ผ่าน" (span 3)
  const sr8 = staffTable.appendTableRow();
  const sc8_0 = sr8.appendTableCell("ผลตรวจสุขภาพ   ☐ ผ่าน   ☐ ไม่ผ่าน");
  sc8_0.setWidth(180);
  const sc8_1 = sr8.appendTableCell("");
  const sc8_2 = sr8.appendTableCell("");
  formatCellText(sc8_0, sc8_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 9: "☐ ชำระเงินสด" (span 3)
  const sr9 = staffTable.appendTableRow();
  const sc9_0 = sr9.appendTableCell("☐ ชำระเงินสด");
  sc9_0.setWidth(180);
  const sc9_1 = sr9.appendTableCell("");
  const sc9_2 = sr9.appendTableCell("");
  formatCellText(sc9_0, sc9_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  // Row 10: Cell 0 (span 1) "☐ กู้กองทุน" and Cell 1 (span 2) "☐ หลักทรัพย์   ☐ บุคคลค้ำ"
  const sr10 = staffTable.appendTableRow();
  const sc10_0 = sr10.appendTableCell("☐ กู้กองทุน");
  sc10_0.setWidth(60);
  const sc10_1 = sr10.appendTableCell("☐ หลักทรัพย์   ☐ บุคคลค้ำ");
  sc10_1.setWidth(120);
  const sc10_2 = sr10.appendTableCell("");
  formatCellText(sc10_0, sc10_0.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7);
  formatCellText(sc10_1, sc10_1.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7);

  // Perform merges at the end after the table layout is fully populated
  sr0.getCell(2).merge(); sr0.getCell(1).merge();
  sr1.getCell(2).merge(); sr1.getCell(1).merge();
  sr2.getCell(1).merge();
  sr3.getCell(2).merge(); sr3.getCell(1).merge();
  sr4.getCell(2).merge(); sr4.getCell(1).merge();
  sr5.getCell(2).merge(); sr5.getCell(1).merge();
  sr6.getCell(2).merge(); sr6.getCell(1).merge();
  sr7.getCell(2).merge(); sr7.getCell(1).merge();
  sr8.getCell(2).merge(); sr8.getCell(1).merge();
  sr9.getCell(2).merge(); sr9.getCell(1).merge();
  sr10.getCell(2).merge();

  // Spacer
  addSpacer(body, 4);

  // Main Details Table 1: Personal Info (Unified)
  const personalTable = body.appendTable();
  styleTableBorder(personalTable);
  
  const addDetailsRow = (table, label1, val1, label2, val2, size) => {
    const row = table.appendTableRow();
    
    // Cell 1: Label 1
    const c1 = row.appendTableCell();
    c1.setWidth(90);
    formatCellText(c1, label1, true, DocumentApp.HorizontalAlignment.LEFT, size || 8);
    
    if (label2 !== undefined && label2 !== null) {
      // Cell 2: Val 1
      const c2 = row.appendTableCell();
      c2.setWidth(210);
      formatCellText(c2, val1, false, DocumentApp.HorizontalAlignment.LEFT, size || 8);
      
      // Cell 3: Label 2
      const c3 = row.appendTableCell();
      c3.setWidth(60);
      formatCellText(c3, label2, true, DocumentApp.HorizontalAlignment.LEFT, size || 8);
      
      // Cell 4: Val 2
      const c4 = row.appendTableCell();
      c4.setWidth(140);
      formatCellText(c4, val2, false, DocumentApp.HorizontalAlignment.LEFT, size || 8);
    } else {
      // Cell 2: Val 1 spanning columns 2, 3, and 4
      const c2 = row.appendTableCell();
      c2.setWidth(410);
      formatCellText(c2, val1, false, DocumentApp.HorizontalAlignment.LEFT, size || 8);
      
      row.appendTableCell("");
      row.appendTableCell("");
      row.getCell(3).merge();
      row.getCell(2).merge();
    }
  };

  addDetailsRow(personalTable, "วุฒิที่ใช้สมัคร", "<<applyEduLevel_m6>> ม.6   <<applyEduLevel_nfe>> กศน.   <<applyEduLevel_voc>> ปวช.   <<applyEduLevel_hvoc>> ปวส.   <<applyEduLevel_uni>> ป.ตรี", "สาขาวิชา", "<<applyMajor>>");
  addDetailsRow(personalTable, "ชื่อ-นามสกุล (ไทย)", "<<titleTh>><<fullNameTh>>", "ชื่อเล่น", "<<nicknameTh>>");
  addDetailsRow(personalTable, "ชื่อ-นามสกุล (อังกฤษ)", "<<fullNameEn>>", "ชื่อเล่น (อังกฤษ)", "<<nicknameEn>>");
  addDetailsRow(personalTable, "วัน/เดือน/ปี เกิด", "วันที่ <<birthDate_day>> เดือน <<birthDate_month>> พ.ศ. <<birthDate_yearTh>> ค.ศ. <<birthDate_yearEn>>   อายุ <<ageYears>> ปี <<ageMonths>> เดือน");
  addDetailsRow(personalTable, "สถานภาพ", "<<maritalStatus_single>> โสด   <<maritalStatus_married_unreg>> แต่งงานไม่จดทะเบียน   <<maritalStatus_married_reg>> แต่งงานจดทะเบียน   <<maritalStatus_divorced>> หย่าร้าง   <<maritalStatus_has_child>> มีบุตร <<childCount>> คน");
  addDetailsRow(personalTable, "ที่อยู่ตามทะเบียนบ้าน", "<<addressRegister>>");
  addDetailsRow(personalTable, "ที่อยู่ปัจจุบัน", "<<addressCurrent>>");
  addDetailsRow(personalTable, "เบอร์โทรส่วนตัว", "<<phonePersonal>>", "Facebook", "<<facebook>>");
  addDetailsRow(personalTable, "ID LINE", "<<lineId>>", "Email", "<<email>>");
  addDetailsRow(personalTable, "เบอร์โทรผู้ปกครอง", "<<phoneParent>>", "พักอาศัยอยู่กับ", "<<livingWith_parents>> พ่อแม่   <<livingWith_other>> อื่นๆ");
  addDetailsRow(personalTable, "เกี่ยวเนื่องเป็น", "<<parentRelationship>>", "รายละเอียดอื่น", "<<livingWithOtherDetails>>");

  // Row for สัดส่วน
  const rowSize = personalTable.appendTableRow();
  const sc1 = rowSize.appendTableCell();
  sc1.setWidth(90);
  formatCellText(sc1, "สัดส่วน", true, DocumentApp.HorizontalAlignment.LEFT, 8);
  
  const sc2 = rowSize.appendTableCell();
  sc2.setWidth(410);
  sc2.setPaddingTop(0); sc2.setPaddingBottom(0); sc2.setPaddingLeft(0); sc2.setPaddingRight(0);
  
  rowSize.appendTableCell("");
  rowSize.appendTableCell("");
  rowSize.getCell(3).merge();
  rowSize.getCell(2).merge();
  
  const nestedSize = sc2.appendTable();
  nestedSize.setBorderColor('#000000');
  nestedSize.setBorderWidth(0.5);
  
  // Row 1 of nested
  const nr1 = nestedSize.appendTableRow();
  const nr1_c1 = nr1.appendTableCell("ส่วนสูง\n(ซม.)"); nr1_c1.setWidth(60);
  const nr1_c2 = nr1.appendTableCell("น้ำหนัก\n(กก.)"); nr1_c2.setWidth(60);
  const nr1_c3 = nr1.appendTableCell("ค่า BMI"); nr1_c3.setWidth(50);
  const nr1_c4 = nr1.appendTableCell("สัดส่วนจริง"); nr1_c4.setWidth(90);
  const nr1_c5 = nr1.appendTableCell("ถนัดมือ"); nr1_c5.setWidth(70);
  const nr1_c6 = nr1.appendTableCell("หมายเหตุ"); nr1_c6.setWidth(80);
  
  // Row 2 of nested
  const nr2 = nestedSize.appendTableRow();
  const nr2_c1 = nr2.appendTableCell("<<height>>"); nr2_c1.setWidth(60);
  const nr2_c2 = nr2.appendTableCell("<<weight>>"); nr2_c2.setWidth(60);
  const nr2_c3 = nr2.appendTableCell("<<bmi>>"); nr2_c3.setWidth(50);
  const nr2_c4 = nr2.appendTableCell("<<bodyShape_normal>> ปกติ   <<bodyShape_thin>> ผอม   <<bodyShape_fat>> อ้วน"); nr2_c4.setWidth(90);
  const nr2_c5 = nr2.appendTableCell("<<handedness_left>> ซ้าย   <<handedness_right>> ขวา"); nr2_c5.setWidth(70);
  const nr2_c6 = nr2.appendTableCell("<<bodyNotes>>"); nr2_c6.setWidth(80);
  
  [nr1, nr2].forEach((nrow, rIdx) => {
    for (let c = 0; c < nrow.getNumCells(); c++) {
      const ncell = nrow.getCell(c);
      formatCellText(ncell, ncell.getText(), rIdx === 0, DocumentApp.HorizontalAlignment.CENTER, 7.5);
    }
  });
  
  if (sc2.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
    sc2.removeChild(sc2.getChild(0));
  }

  // Row for สายตา
  const rowEye = personalTable.appendTableRow();
  const ec1 = rowEye.appendTableCell();
  ec1.setWidth(90);
  formatCellText(ec1, "สายตา", true, DocumentApp.HorizontalAlignment.LEFT, 8);
  
  const ec2 = rowEye.appendTableCell();
  ec2.setWidth(410);
  ec2.setPaddingTop(0); ec2.setPaddingBottom(0); ec2.setPaddingLeft(0); ec2.setPaddingRight(0);
  
  rowEye.appendTableCell("");
  rowEye.appendTableCell("");
  rowEye.getCell(3).merge();
  rowEye.getCell(2).merge();
  
  const nestedEye = ec2.appendTable();
  nestedEye.setBorderColor('#000000');
  nestedEye.setBorderWidth(0.5);
  
  // Row 1: Headers
  const er1 = nestedEye.appendTableRow();
  const er1_c1 = er1.appendTableCell("ค่าสายตา"); er1_c1.setWidth(205);
  const er1_c2 = er1.appendTableCell("ใส่แว่นตา / คอนแทคเลนส์"); er1_c2.setWidth(205);
  
  // Row 2: Left
  const er2 = nestedEye.appendTableRow();
  const er2_c1 = er2.appendTableCell("ซ้าย .....<<eyesightLeftVal>>.....   <<eyesightLeftStatus_normal>> ปกติ   <<eyesightLeftStatus_short>> สั้น   <<eyesightLeftStatus_astigmatism>> เอียง"); er2_c1.setWidth(205);
  const er2_c2 = er2.appendTableCell("ซ้าย...................   <<glassesLeft>> แว่นตา"); er2_c2.setWidth(205);
  
  // Row 3: Right
  const er3 = nestedEye.appendTableRow();
  const er3_c1 = er3.appendTableCell("ขวา .....<<eyesightRightVal>>.....   <<eyesightRightStatus_normal>> ปกติ   <<eyesightRightStatus_short>> สั้น   <<eyesightRightStatus_astigmatism>> เอียง"); er3_c1.setWidth(205);
  const er3_c2 = er3.appendTableCell("ขวา...................   <<contactsLeft>> คอนแทคเลนส์"); er3_c2.setWidth(205);
  
  [er1, er2, er3].forEach((nrow, rIdx) => {
    for (let c = 0; c < nrow.getNumCells(); c++) {
      const ncell = nrow.getCell(c);
      formatCellText(ncell, ncell.getText(), rIdx === 0, DocumentApp.HorizontalAlignment.LEFT, 7.5);
    }
  });
  
  if (ec2.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
    ec2.removeChild(ec2.getChild(0));
  }

  // Other Details on Page 1
  addDetailsRow(personalTable, "ตาบอดสี", "<<colorBlindness_none>> ไม่บอดสี   <<colorBlindness_red_green>> ตาบอดสีแดง-เขียว   <<colorBlindness_blue_yellow>> ตาบอดสีน้ำเงิน-เหลือง   <<colorBlindness_all>> ตาบอดทุกสี");
  addDetailsRow(personalTable, "รอยสัก", "<<tattoo_none>> ไม่มี   <<tattoo_has>> มี (จำนวน: <<tattooPoints>> จุด   ขนาด: <<tattooSize>> ซม.   บริเวณ: <<tattooArea>>)");
  addDetailsRow(personalTable, "ภาระทางทหาร", "<<militaryStatus_served>> เกณฑ์แล้ว (ปี พ.ศ. <<militaryYear>>)   <<militaryStatus_nst>> รด.   <<militaryStatus_exempted>> ผ่อนผัน/ยกเว้น   <<militaryStatus_not_reached>> ยังไม่ถึงเกณฑ์ (กำหนดปี พ.ศ. <<militaryYear>>)");
  addDetailsRow(personalTable, "เปลี่ยนชื่อ-นามสกุล", "<<nameChanged_none>> ไม่เคย   <<nameChanged_has>> เคย (ชื่อเดิม: <<nameChangedOldName>>  นามสกุล: <<nameChangedOldLastName>>  จำนวน: <<nameChangedCount>> ครั้ง  สาเหตุ: <<nameChangedReason>>)");

  body.appendPageBreak();

  // === PAGE 2 ===
  const page2Table = body.appendTable();
  styleTableBorder(page2Table);

  const addPage2Row = (label, content) => {
    const row = page2Table.appendTableRow();
    const c1 = row.appendTableCell();
    c1.setWidth(90);
    formatCellText(c1, label, true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
    
    const c2 = row.appendTableCell();
    c2.setWidth(410);
    formatCellText(c2, content, false, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  };

  // Row 1: Header
  const p2TitleRow = page2Table.appendTableRow();
  const p2TitleCell = p2TitleRow.appendTableCell("สุขภาพ/ร่างกาย");
  p2TitleCell.setWidth(90);
  formatCellText(p2TitleCell, "สุขภาพ/ร่างกาย", true, DocumentApp.HorizontalAlignment.CENTER, 9, '#e2e8f0');
  const p2TitleCell2 = p2TitleRow.appendTableCell("");
  p2TitleCell2.setWidth(410);
  formatCellText(p2TitleCell2, "", false, DocumentApp.HorizontalAlignment.LEFT, 9, '#e2e8f0');

  addPage2Row("ประวัติการเจ็บป่วย", "<<illness_none>> ไม่เคย   <<illness_surgery>> เคยผ่าตัด   <<illness_accident>> เคยเกิดอุบัติเหตุหนัก   <<illness_fracture>> เคยกระดูกหัก   <<illness_metal>> มีเหล็กดาม   <<illness_joint>> มีอาการเจ็บหลัง/ข้อต่างๆ   <<illness_other>> อื่นๆ\nเมื่อใด: <<illnessTime>>   สาเหตุ: <<illnessCause>>   อาการ: <<illnessSymptoms>>");
  // โรคประจำตัว Row (3-column nested table structure matching DOCX)
  const diseaseRow = page2Table.appendTableRow();
  const dc1 = diseaseRow.appendTableCell();
  dc1.setWidth(90);
  formatCellText(dc1, "โรคประจำตัว", true, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  const dc2 = diseaseRow.appendTableCell();
  dc2.setWidth(410);
  dc2.setPaddingTop(0); dc2.setPaddingBottom(0); dc2.setPaddingLeft(0); dc2.setPaddingRight(0);

  const dnested = dc2.appendTable();
  dnested.setBorderColor('#000000');
  dnested.setBorderWidth(0.5);

  const dr1 = dnested.appendTableRow();
  const dr1_c1 = dr1.appendTableCell("<<disease_none>> ไม่มีโรคประจำตัว   <<disease_anemia>> โรคเลือดจาง/ภาวะเลือดจาง   <<disease_thalassemia>> ธาลัสซีเมีย/พาหะ   <<disease_thyroid>> ไทรอยด์");
  dr1_c1.setWidth(136);
  const dr1_c2 = dr1.appendTableCell("<<disease_asthma>> หอบหืด   <<disease_epilepsy>> ลมชัก/ลมบ้าหมู   <<disease_hepb>> ไวรัสตับอักเสบบี   <<disease_allergy>> ภูมิแพ้");
  dr1_c2.setWidth(136);
  const dr1_c3 = dr1.appendTableCell("<<disease_depression>> ภาวะซึมเศร้า   <<disease_anxiety>> วิตกกังวล/แพนิค   <<disease_other>> มีโรคประจำตัวอื่นๆ");
  dr1_c3.setWidth(138);

  const dr2 = dnested.appendTableRow();
  const dr2_c1 = dr2.appendTableCell("โรค/อาการ: <<diseaseDetails>>");
  dr2_c1.setWidth(410);
  dr2.appendTableCell("");
  dr2.appendTableCell("");
  dr2.getCell(2).merge();
  dr2.getCell(1).merge();

  [dr1, dr2].forEach(row => {
    for (let c = 0; c < row.getNumCells(); c++) {
      formatCellText(row.getCell(c), row.getCell(c).getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7);
    }
  });

  if (dc2.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
    dc2.removeChild(dc2.getChild(0));
  }
  
  // ร่างกาย Row
  const bodyRow = page2Table.appendTableRow();
  const bc1 = bodyRow.appendTableCell();
  bc1.setWidth(90);
  formatCellText(bc1, "ร่างกาย", true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  
  const bc2 = bodyRow.appendTableCell();
  bc2.setWidth(410);
  bc2.setPaddingTop(0); bc2.setPaddingBottom(0); bc2.setPaddingLeft(0); bc2.setPaddingRight(0);
  
  const bnested = bc2.appendTable();
  bnested.setBorderColor('#000000');
  bnested.setBorderWidth(0.5);
  
  const br1 = bnested.appendTableRow();
  const br1_c1 = br1.appendTableCell("กรุ๊ปเลือด"); br1_c1.setWidth(136);
  const br1_c2 = br1.appendTableCell("ศัลยกรรม"); br1_c2.setWidth(136);
  const br1_c3 = br1.appendTableCell("จัดฟัน"); br1_c3.setWidth(138);
  
  const br2 = bnested.appendTableRow();
  const br2_c1 = br2.appendTableCell("<<bloodGroup_A>> A   <<bloodGroup_B>> B   <<bloodGroup_AB>> AB   <<bloodGroup_O>> O"); br2_c1.setWidth(136);
  const br2_c2 = br2.appendTableCell("<<plasticSurgery_none>> ไม่เคยทำ   <<plasticSurgery_has>> เคยทำ: <<plasticSurgeryDetails>>"); br2_c2.setWidth(136);
  const br2_c3 = br2.appendTableCell("<<dentalBraces_none>> ไม่จัดฟัน   <<dentalBraces_bracing>> จัดฟัน (สิ้นสุดเมื่อ: <<dentalBracesEnd>>)"); br2_c3.setWidth(138);
  
  const br3 = bnested.appendTableRow();
  const br3_c1 = br3.appendTableCell("มีตำหนิของร่างกาย: <<bodyMark_has>> มี   <<bodyMark_none>> ไม่มี   อาการ/ลักษณะ: <<bodyMarkSymptoms>>   สาเหตุ: <<bodyMarkCause>>"); br3_c1.setWidth(410);
  br3.appendTableCell("");
  br3.appendTableCell("");
  br3.getCell(2).merge();
  br3.getCell(1).merge();
  
  [br1].forEach(row => {
    for (let c = 0; c < row.getNumCells(); c++) {
      formatCellText(row.getCell(c), row.getCell(c).getText(), true, DocumentApp.HorizontalAlignment.CENTER, 7.5);
    }
  });
  [br2, br3].forEach(row => {
    for (let c = 0; c < row.getNumCells(); c++) {
      formatCellText(row.getCell(c), row.getCell(c).getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);
    }
  });
  
  if (bc2.getChild(0).getType() === DocumentApp.ElementType.PARAGRAPH) {
    bc2.removeChild(bc2.getChild(0));
  }

  addPage2Row("ปวดประจำเดือน", "<<periodPain_none>> ไม่ปวด   <<periodPain_mild>> ปวดเล็กน้อย   <<periodPain_severe>> ปวดมากจนมีผลต่อการดำเนินชีวิตประจำวัน");
  addPage2Row("สูบบุหรี่", "<<smoke_never>> ไม่สูบ   <<smoke_regular>> สูบประจำ   <<smoke_sometimes>> สูบบางครั้ง (เฉลี่ยวันละ: <<smokeQty>> มวน)");
  addPage2Row("บุหรี่ไฟฟ้า", "<<vape_never>> ไม่สูบ   <<vape_regular>> สูบประจำ   <<vape_sometimes>> สูบบางครั้ง (เฉลี่ยวันละ: <<vapeQty>> ครั้ง)");
  addPage2Row("สูบกัญชา", "<<cannabis_never>> ไม่สูบ   <<cannabis_regular>> สูบประจำ   <<cannabis_sometimes>> สูบบางครั้ง (เฉลี่ยวันละ: <<cannabisQty>> ครั้ง)");
  addPage2Row("ดื่มน้ำกระท่อม", "<<kratom_never>> ไม่ดื่ม   <<kratom_regular>> ดื่มประจำ   <<kratom_sometimes>> ดื่มบางครั้ง (เฉลี่ยสัปดาห์ละ: <<kratomQty>> วัน)");
  addPage2Row("ดื่มแอลกอฮอล์", "<<alcohol_never>> ไม่ดื่ม   <<alcohol_regular>> ดื่มประจำ   <<alcohol_sometimes>> ดื่มบางครั้ง (เฉลี่ยสัปดาห์ละ: <<alcoholQty>> วัน)");

  // Section Header for documents
  const page2DocHeader = page2Table.appendTableRow();
  const hCell = page2DocHeader.appendTableCell("ด้านเอกสารส่วนตัว");
  hCell.setWidth(90);
  formatCellText(hCell, "ด้านเอกสารส่วนตัว", true, DocumentApp.HorizontalAlignment.CENTER, 9, '#e2e8f0');
  const hCell2 = page2DocHeader.appendTableCell("");
  hCell2.setWidth(410);
  formatCellText(hCell2, "", false, DocumentApp.HorizontalAlignment.LEFT, 9, '#e2e8f0');

  addPage2Row("ใบอนุญาตขับขี่", "<<driverLicense_none>> ไม่มี   <<driverLicense_has>> มี (ประเภท: <<licenseCar>> รถยนต์   <<licenseMotorcycle>> รถจักรยานยนต์)");
  addPage2Row("เคยมีพาสปอร์ตหรือไม่", "<<passport_none>> ไม่มี   <<passport_has>> มี (เหตุผล: <<passportReason>>  ประเทศ: <<passportCountry>>  เมื่อ: <<passportTime>>  พาสปอร์ตหมดอายุปี: <<passportExpireYear>>)");
  addPage2Row("เคยไปญี่ปุ่นหรือไม่", "<<japanVisit_none>> ไม่เคย   <<japanVisit_has>> เคย (เหตุผล: <<japanVisitReason>>  ประเภทวีซ่า: <<japanVisitVisaType>>  เมื่อ: <<japanVisitTime>>)");
  addPage2Row("เคยสมัครงานไปญี่ปุ่น\nกับบริษัทอื่นหรือไม่", "<<japanApply_none>> ไม่เคย   <<japanApply_has>> เคย (บริษัท: <<japanApplyCompany>>  เมื่อ: <<japanApplyTime>>  ถึงขั้นตอน: <<japanApplyStep>>  เหตุผลที่ถอนตัว: <<japanApplyWithdrawReason>>)");
  addPage2Row("เคยมีคดีความอาญา\nหรือไม่", "<<criminalRecord_none>> ไม่มีคดีความ   <<criminalRecord_has>> เคยมี (รายละเอียดคดี: <<criminalRecordDetails>>)");
  addPage2Row("มีภาระหนี้สินหรือไม่", "นักศึกษา: <<studentDebt_none>> ไม่มี   <<studentDebt_has>> มี กู้มาจาก: <<studentDebtSource>>   เหตุผล: <<studentDebtReason>>   จำนวน: <<studentDebtAmount>> บาท   ผู้ปกครอง: <<parentDebt_none>> ไม่มี   <<parentDebt_has>> มี กู้มาจาก: <<parentDebtSource>>   เหตุผล: <<parentDebtReason>>   จำนวน: <<parentDebtAmount>> บาท");

  // Merge headers after table structure is fully populated
  p2TitleRow.getCell(1).merge();
  page2DocHeader.getCell(1).merge();

  body.appendPageBreak();

  // === PAGE 3 ===
  const famTable = body.appendTable();
  styleTableBorder(famTable);
  
  // Row 1: Title
  const famTitleRow = famTable.appendTableRow();
  const fCell = famTitleRow.appendTableCell("ประวัติครอบครัว");
  fCell.setWidth(30);
  formatCellText(fCell, "ประวัติครอบครัว", true, DocumentApp.HorizontalAlignment.CENTER, 9, '#e2e8f0');
  for (let c = 1; c < 8; c++) {
    const fc = famTitleRow.appendTableCell("");
    fc.setWidth(c === 1 ? 130 : (c === 2 ? 60 : (c === 3 ? 30 : (c === 4 ? 70 : (c === 5 ? 100 : (c === 6 ? 50 : 30))))));
    formatCellText(fc, "", false, DocumentApp.HorizontalAlignment.LEFT, 9, '#e2e8f0');
  }

  // Row 2: Headers
  const famHeader = famTable.appendTableRow();
  const fh1 = famHeader.appendTableCell("ลำดับ"); fh1.setWidth(30);
  const fh2 = famHeader.appendTableCell("ชื่อ-นามสกุล"); fh2.setWidth(130);
  const fh3 = famHeader.appendTableCell("ความสัมพันธ์"); fh3.setWidth(60);
  const fh4 = famHeader.appendTableCell("อายุ"); fh4.setWidth(30);
  const fh5 = famHeader.appendTableCell("อาชีพ"); fh5.setWidth(70);
  const fh6 = famHeader.appendTableCell("สถานที่ทำงาน/บริษัท"); fh6.setWidth(100);
  const fh7 = famHeader.appendTableCell("รายได้ (ต่อเดือน)"); fh7.setWidth(50);
  const fh8 = famHeader.appendTableCell("จังหวัด"); fh8.setWidth(30);
  
  [fh1, fh2, fh3, fh4, fh5, fh6, fh7, fh8].forEach(cell => {
    formatCellText(cell, cell.getText(), true, DocumentApp.HorizontalAlignment.CENTER, 7.5);
  });

  for (let i = 1; i <= 8; i++) {
    const row = famTable.appendTableRow();
    
    const c1 = row.appendTableCell(i.toString()); c1.setWidth(30);
    formatCellText(c1, i.toString(), false, DocumentApp.HorizontalAlignment.CENTER, 7.5, '#f8fafc');
    
    const c2 = row.appendTableCell(`<<fam_name_${i}>>`); c2.setWidth(130);
    const c3 = row.appendTableCell(`<<fam_rel_${i}>>`); c3.setWidth(60);
    const c4 = row.appendTableCell(`<<fam_age_${i}>>`); c4.setWidth(30);
    const c5 = row.appendTableCell(`<<fam_job_${i}>>`); c5.setWidth(70);
    const c6 = row.appendTableCell(`<<fam_work_${i}>>`); c6.setWidth(100);
    const c7 = row.appendTableCell(`<<fam_inc_${i}>>`); c7.setWidth(50);
    const c8 = row.appendTableCell(`<<fam_prov_${i}>>`); c8.setWidth(30);
    
    [c2, c3, c4, c5, c6, c7, c8].forEach(cell => {
      formatCellText(cell, cell.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);
    });
  }

  // Family Status Row in family table (two merged sections matching PDF layout)
  const fsRow = famTable.appendTableRow();
  
  // Left side: "สถานะครอบครัว" (spans columns 1-2, width 160)
  const fsCellL = fsRow.appendTableCell("สถานะครอบครัว");
  formatCellText(fsCellL, "สถานะครอบครัว", true, DocumentApp.HorizontalAlignment.CENTER, 8);
  
  const fsCell1 = fsRow.appendTableCell("");
  formatCellText(fsCell1, "", true, DocumentApp.HorizontalAlignment.CENTER, 8);

  // Right side: Checkboxes and Guardian (spans columns 3-8, width 340)
  const fsCellR = fsRow.appendTableCell(
    "<<familyStatus_together>> อยู่ร่วมกัน   <<familyStatus_divorced>> พ่อแม่หย่าร้าง (จดทะเบียนหย่า)   <<familyStatus_separated>> พ่อแม่แยกทาง (ไม่ได้จดทะเบียน)   " +
    "<<familyStatus_father_deceased>> บิดาเสียชีวิต   <<familyStatus_mother_deceased>> มารดาเสียชีวิต   " +
    "ปัจจุบันอยู่ในความดูแล/ปกครองของใคร: <<guardian>>"
  );
  formatCellText(fsCellR, fsCellR.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 8);

  for (let c = 3; c < 8; c++) {
    const fsc = fsRow.appendTableCell("");
    formatCellText(fsc, "", false, DocumentApp.HorizontalAlignment.LEFT, 8);
  }

  // Merge family table rows after table is populated
  for (let c = 7; c >= 1; c--) famTitleRow.getCell(c).merge();
  for (let c = 7; c >= 3; c--) fsRow.getCell(c).merge();
  fsRow.getCell(1).merge();

  addSpacer(body, 4);

  // Japan Contact Details Box
  const japanContactTable = body.appendTable();
  styleTableBorder(japanContactTable);
  
  const jcRow1 = japanContactTable.appendTableRow();
  const jc1 = jcRow1.appendTableCell(); jc1.setWidth(90);
  formatCellText(jc1, "มีคนรู้จักที่อยู่ญี่ปุ่นหรือไม่", true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  const jc2 = jcRow1.appendTableCell(); jc2.setWidth(410);
  formatCellText(jc2, "<<japanContact_none>> ไม่มี   <<japanContact_has>> มี", false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  const jcRow2 = japanContactTable.appendTableRow();
  const jc3 = jcRow2.appendTableCell(); jc3.setWidth(90);
  formatCellText(jc3, "ชื่อ – นามสกุล คนรู้จัก", true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  const jc4 = jcRow2.appendTableCell(); jc4.setWidth(410);
  formatCellText(jc4, "<<japanContactName>>   เพศ: <<japanContactGender_male>> ชาย  <<japanContactGender_female>> หญิง   อายุ: <<japanContactAge>> ปี", false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  const jcRow3 = japanContactTable.appendTableRow();
  const jc5 = jcRow3.appendTableCell(); jc5.setWidth(90);
  formatCellText(jc5, "ความสัมพันธ์", true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  const jc6 = jcRow3.appendTableCell(); jc6.setWidth(410);
  formatCellText(jc6, "<<japanContactRelationship>>   ที่อยู่ญี่ปุ่น: <<japanContactAddress>>", false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  const jcRow4 = japanContactTable.appendTableRow();
  const jc7 = jcRow4.appendTableCell(); jc7.setWidth(90);
  formatCellText(jc7, "เบอร์โทร", true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  const jc8 = jcRow4.appendTableCell(); jc8.setWidth(410);
  formatCellText(jc8, "<<japanContactPhone>>   อยู่ญี่ปุ่นกี่ปี: <<japanContactYears>> ปี", false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  const jcRow5 = japanContactTable.appendTableRow();
  const jc9 = jcRow5.appendTableCell(); jc9.setWidth(90);
  formatCellText(jc9, "Facebook / Social", true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  const jc10 = jcRow5.appendTableCell(); jc10.setWidth(410);
  formatCellText(jc10, "<<japanContactFacebook>>   ประเภทวีซ่า: <<japanContactVisaType>>", false, DocumentApp.HorizontalAlignment.LEFT, 7.5);

  body.appendPageBreak();

  // === PAGE 4 ===
  const eduTable = body.appendTable();
  styleTableBorder(eduTable);
  
  // Row 1: Merged Title
  const eduTitleRow = eduTable.appendTableRow();
  const edCell = eduTitleRow.appendTableCell("ประวัติการศึกษา");
  edCell.setWidth(90);
  formatCellText(edCell, "ประวัติการศึกษา", true, DocumentApp.HorizontalAlignment.CENTER, 9, '#e2e8f0');
  for (let c = 1; c < 4; c++) {
    const edc = eduTitleRow.appendTableCell("");
    edc.setWidth(c === 1 ? 180 : (c === 2 ? 110 : 120));
    formatCellText(edc, "", false, DocumentApp.HorizontalAlignment.LEFT, 9, '#e2e8f0');
  }

  // Row 2: Headers
  const eduHeader = eduTable.appendTableRow();
  const ehCol1 = eduHeader.appendTableCell("ระดับการศึกษา"); ehCol1.setWidth(90);
  const ehCol2 = eduHeader.appendTableCell("ชื่อสถานศึกษา"); ehCol2.setWidth(180);
  const ehCol3 = eduHeader.appendTableCell("สาขา/แผนก"); ehCol3.setWidth(110);
  const ehCol4 = eduHeader.appendTableCell("เข้าศึกษา-สำเร็จการศึกษา"); ehCol4.setWidth(120);
  [ehCol1, ehCol2, ehCol3, ehCol4].forEach(cell => {
    formatCellText(cell, cell.getText(), true, DocumentApp.HorizontalAlignment.CENTER, 7.5);
  });

  const eduRows = [
    { label: "มัธยมต้น", prefix: "edu_m3" },
    { label: "มัธยมปลาย/ปวช.", prefix: "edu_voc" },
    { label: "ปวส./ทวิภาคี", prefix: "edu_hvoc" },
    { label: "ปริญญาตรี", prefix: "edu_uni" }
  ];

  eduRows.forEach(item => {
    const row = eduTable.appendTableRow();
    const c1 = row.appendTableCell(item.label); c1.setWidth(90);
    formatCellText(c1, item.label, true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
    
    const c2 = row.appendTableCell(`<<${item.prefix}_school>>`); c2.setWidth(180);
    const c3 = row.appendTableCell(`<<${item.prefix}_major>>`); c3.setWidth(110);
    const c4 = row.appendTableCell(`ปีที่เข้า: <<${item.prefix}_start_year>> (วันที่: <<${item.prefix}_start_date>>)\nปีที่จบ: <<${item.prefix}_end_year>> (วันที่: <<${item.prefix}_end_date>>)`); c4.setWidth(120);
    
    [c2, c3, c4].forEach(cell => {
      formatCellText(cell, cell.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 7.5);
    });
  });

  // Interrupted Row (split matching PDF: grey label left, choices right)
  const intRow = eduTable.appendTableRow();
  const intCellL = intRow.appendTableCell("กรณีเรียนไม่ต่อเนื่อง/\nดรอป/\nลาออกระหว่าง\nเรียน");
  formatCellText(intCellL, "กรณีเรียนไม่ต่อเนื่อง/\nดรอป/\nลาออกระหว่าง\nเรียน", true, DocumentApp.HorizontalAlignment.CENTER, 8);
  
  const intCellR = intRow.appendTableCell(
    "<<eduInterrupted_none>> ไม่มีประวัติ   <<eduInterrupted_drop>> ดรอปเรียน   <<eduInterrupted_quit>> ลาออกระหว่างเรียน   รายละเอียด/สาเหตุ: <<eduInterruptedDetails>>"
  );
  formatCellText(intCellR, intCellR.getText(), false, DocumentApp.HorizontalAlignment.LEFT, 8);

  for (let c = 2; c < 4; c++) {
    const intc = intRow.appendTableCell("");
    formatCellText(intc, "", false, DocumentApp.HorizontalAlignment.LEFT, 8);
  }

  // Merge rows after table is populated
  for (let c = 3; c >= 1; c--) eduTitleRow.getCell(c).merge();
  for (let c = 3; c >= 2; c--) intRow.getCell(c).merge();

  addSpacer(body, 4);

  const pWorkTitle = body.appendParagraph("ประวัติการทำงานและการฝึกงานย้อนหลัง (ลำดับที่ 1 - 3)");
  styleParagraph(pWorkTitle, 10, true, DocumentApp.HorizontalAlignment.LEFT, '#000000');
  pWorkTitle.setSpacingAfter(4);

  // Job Cards 1-3
  for (let i = 1; i <= 3; i++) {
    createWorkCard(body, i);
  }

  body.appendPageBreak();

  // === PAGE 5 ===
  const pWorkTitle2 = body.appendParagraph("ประวัติการทำงานและการฝึกงานย้อนหลัง (ลำดับที่ 4 - 9)");
  styleParagraph(pWorkTitle2, 10, true, DocumentApp.HorizontalAlignment.LEFT, '#000000');
  pWorkTitle2.setSpacingAfter(4);

  // Job Cards 4-9
  for (let i = 4; i <= 9; i++) {
    createWorkCard(body, i);
  }

  body.appendPageBreak();

  // === PAGE 6 ===
  // Goal Box Table
  const goalTable = body.appendTable();
  styleTableBorder(goalTable);
  
  const addGoalRow = (label, valTag) => {
    const row = goalTable.appendTableRow();
    const c1 = row.appendTableCell();
    c1.setWidth(100);
    formatCellText(c1, label, true, DocumentApp.HorizontalAlignment.LEFT, 8);
    
    const c2 = row.appendTableCell();
    c2.setWidth(400);
    formatCellText(c2, valTag, false, DocumentApp.HorizontalAlignment.LEFT, 8);
  };
  
  addGoalRow("เป้าหมายในชีวิต:", "<<lifeGoal>>");
  addGoalRow("สิ่งที่ทำเวลาว่าง:", "<<hobbies>>");
  addGoalRow("ข้อดีของตนเอง:", "<<strengths>>");
  addGoalRow("ข้อเสียของตนเอง:", "<<weaknesses>>");
  addGoalRow("ความสามารถพิเศษ:", "<<specialSkills>>");

  addSpacer(body, 4);

  // Evaluation Table
  const evalTable = body.appendTable();
  styleTableBorder(evalTable);
  
  // Row 1: Merged Title
  const evTitleRow = evalTable.appendTableRow();
  const evTitleCell = evTitleRow.appendTableCell("แบบประเมิน");
  evTitleCell.setWidth(200);
  formatCellText(evTitleCell, "แบบประเมิน", true, DocumentApp.HorizontalAlignment.CENTER, 9, '#e2e8f0');
  for (let c = 1; c < 4; c++) {
    const evc = evTitleRow.appendTableCell("");
    evc.setWidth(c === 1 ? 50 : (c === 2 ? 50 : 200));
    formatCellText(evc, "", false, DocumentApp.HorizontalAlignment.LEFT, 9, '#e2e8f0');
  }

  // Row 2: Headers
  const evHeader = evalTable.appendTableRow();
  const evH1 = evHeader.appendTableCell("หัวข้อการประเมิน"); evH1.setWidth(200);
  const evH2 = evHeader.appendTableCell("ผ่าน"); evH2.setWidth(50);
  const evH3 = evHeader.appendTableCell("ไม่ผ่าน"); evH3.setWidth(50);
  const evH4 = evHeader.appendTableCell("หมายเหตุ"); evH4.setWidth(200);
  
  [evH1, evH2, evH3, evH4].forEach(cell => {
    formatCellText(cell, cell.getText(), true, DocumentApp.HorizontalAlignment.CENTER, 8);
  });

  const addEvalRow = (title, scoreText) => {
    const row = evalTable.appendTableRow();
    const c1 = row.appendTableCell(title); c1.setWidth(200);
    const c2 = row.appendTableCell("☐"); c2.setWidth(50);
    const c3 = row.appendTableCell("☐"); c3.setWidth(50);
    const c4 = row.appendTableCell(scoreText); c4.setWidth(200);
    
    formatCellText(c1, title, true, DocumentApp.HorizontalAlignment.LEFT, 7.5);
    formatCellText(c2, "☐", false, DocumentApp.HorizontalAlignment.CENTER, 7.5);
    formatCellText(c3, "☐", false, DocumentApp.HorizontalAlignment.CENTER, 7.5);
    formatCellText(c4, scoreText, false, DocumentApp.HorizontalAlignment.LEFT, 7.5);
  };

  addEvalRow("ทดสอบคณิตศาสตร์", "......................... / 30 คะแนน");
  addEvalRow("ทดสอบเรียงไพ่", "ครั้งที่ 1................นาที ผิด................จุด\nครั้งที่ 2 ...............นาที ผิด................จุด");
  addEvalRow("ทดสอบสมรรถภาพร่างกาย", "ลุกนั่ง............................ครั้ง/1นาที\nซิทอัพ...........................ครั้ง/1นาที\nวิดพื้น...........................ครั้ง/1นาที");

  // Instruction Row (merged)
  const instRow = evalTable.appendTableRow();
  const instCell = instRow.appendTableCell();
  instCell.setWidth(200);
  formatCellText(instCell, "วิธีการประเมินและทดสอบ\n1. คุณสมบัติเบื้องต้น : พิจารณาตามเกณฑ์\n2. การทดสอบคณิตศาสตร์ : 30 ข้อ เวลา 10 นาที ผ่านเกณฑ์การประเมิน 15 ข้อขึ้นไป\n3. ทดสอบเรียงไพ่ :\n   - แยกสีดำ สีแดง\n   - แยกประเภทของไพ่ ดอกจิก หัวใจ โพดำ ข้าวหลามตัด\n   - เรียงจากจำนวนน้อยไปหาจำนวนมาก A-10 และตามด้วย J,Q,K\n   - เวลา 5 นาที และสามารถทำซ้ำได้ 2 รอบ\nทดสอบสมรรถภาพ : วิดพื้น ลุกนั่ง และซิทอัพ เวลา 1 นาที ผ่านเกณฑ์การประเมิน 30 ครั้งขึ้นไป", false, DocumentApp.HorizontalAlignment.LEFT, 7.5, '#f8fafc');
  for (let c = 1; c < 4; c++) {
    const instc = instRow.appendTableCell("");
    instc.setWidth(c === 1 ? 50 : (c === 2 ? 50 : 200));
    formatCellText(instc, "", false, DocumentApp.HorizontalAlignment.LEFT, 7.5, '#f8fafc');
  }

  // Merge rows after table is populated
  for (let c = 3; c >= 1; c--) evTitleRow.getCell(c).merge();
  for (let c = 3; c >= 1; c--) instRow.getCell(c).merge();

  addSpacer(body, 4);

  // Summary Result Table (3-column structure matching DOCX)
  const summaryTable = body.appendTable();
  styleTableBorder(summaryTable);
  const sumRow = summaryTable.appendTableRow();
  const sumCell1 = sumRow.appendTableCell("สรุปผลการรับสมัคร"); sumCell1.setWidth(200);
  const sumCell2 = sumRow.appendTableCell("☐ ผ่าน"); sumCell2.setWidth(150);
  const sumCell3 = sumRow.appendTableCell("☐ ไม่ผ่าน"); sumCell3.setWidth(150);
  
  formatCellText(sumCell1, "สรุปผลการรับสมัคร", true, DocumentApp.HorizontalAlignment.LEFT, 8.5, '#e2e8f0');
  formatCellText(sumCell2, "☐ ผ่าน", false, DocumentApp.HorizontalAlignment.CENTER, 8.5);
  formatCellText(sumCell3, "☐ ไม่ผ่าน", false, DocumentApp.HorizontalAlignment.CENTER, 8.5);

  addSpacer(body, 4);

  // Additional Details Table
  const addDetailsTable = body.appendTable();
  styleTableBorder(addDetailsTable);
  const dtRow = addDetailsTable.appendTableRow();
  const dtCell = dtRow.appendTableCell();
  dtCell.setWidth(500);
  formatCellText(dtCell, "รายละเอียดเพิ่มเติม (สำหรับเจ้าหน้าที่)\n......................................................................................................................................................................\n......................................................................................................................................................................\n......................................................................................................................................................................\n......................................................................................................................................................................", false, DocumentApp.HorizontalAlignment.LEFT, 8);

  addSpacer(body, 6);

  // Signatures Table (Invisible borders, side-by-side)
  const sigTable = body.appendTable();
  sigTable.setBorderWidth(0);
  sigTable.setBorderColor('#ffffff');
  const sigRow = sigTable.appendTableRow();
  const scL = sigRow.appendTableCell(); scL.setWidth(250);
  const scR = sigRow.appendTableCell(); scR.setWidth(250);
  
  formatCellText(scL, "ลงชื่อ ................................................. ผู้ประเมิน\n\nวันที่ ...................................................", false, DocumentApp.HorizontalAlignment.LEFT, 8.5);
  formatCellText(scR, "ลงชื่อ ................................................. ผู้สมัคร\n( <<applicantSignatureName>> )\n\nวันที่ ...................................................", false, DocumentApp.HorizontalAlignment.LEFT, 8.5);

  doc.saveAndClose();
  return { success: true, message: "สร้างเอกสารเทมเพลตใบประวัติและจัดสไตล์โครงสร้างตาราง 6 หน้า เรียบร้อยแล้วครับ." };
}

