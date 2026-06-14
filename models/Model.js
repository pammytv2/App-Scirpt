const Model = {
  /**
   * ดึงลำดับคอลัมน์ที่กำหนดไว้สำหรับ Google Sheets แบบสม่ำเสมอ
   */
  getExpectedColumns() {
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
  },

  /**
   * ตรวจสอบและอัปเดตหัวตารางในชีตให้ตรงกับเวอร์ชันล่าสุดโดยอัตโนมัติ
   */
  checkAndSyncHeaders(sheet) {
    const expected = this.getExpectedColumns();
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
  },

  /**
   * จัดเก็บข้อมูลฟอร์มลง Google Sheets
   */
  saveStudent(formData) {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }
    
    this.checkAndSyncHeaders(sheet);
    const columns = this.getExpectedColumns();
    
    const rowData = [new Date()];
    for (let colIdx = 1; colIdx < columns.length; colIdx++) {
      const fieldName = columns[colIdx];
      const val = formData[fieldName];
      
      if (val === true) {
        rowData.push("ใช่");
      } else if (val === false) {
        rowData.push("ไม่");
      } else {
        rowData.push(val !== undefined && val !== null ? val : "");
      }
    }
    
    sheet.appendRow(rowData);
    return sheet;
  },

  /**
   * ดึงข้อมูลนักศึกษาทั้งหมดจาก Google Sheet
   */
  getStudentSubmissions(pwd) {
    const inputPwd = (pwd || "").toString().trim();
    if (inputPwd !== CONFIG.ADMIN_PASSWORD) {
      throw new Error("สิทธิ์การเข้าถึงปฏิเสธ: รหัสผ่านแอดมินไม่ถูกต้อง");
    }

    try {
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      if (!sheet) return [];
      
      this.checkAndSyncHeaders(sheet);
      
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow <= 1) return []; // ไม่มีข้อมูลนอกจากหัวตาราง
      
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      
      const submissions = dataRange.map((row, index) => {
        const obj = {};
        headers.forEach((header, idx) => {
          let val = row[idx];
          if (val instanceof Date) {
            val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
          obj[header] = val;
        });
        
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
  },

  /**
   * ค้นหาแถวนักศึกษาในชีตตาม Student ID
   */
  findStudentRowIndexAndData(studentId) {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error("ไม่พบชีตข้อมูลนักศึกษา");
    
    this.checkAndSyncHeaders(sheet);
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
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
      
      if (!rowObj.studentId) {
        rowObj.studentId = "OLD-" + (i + 1);
      }
      
      if (rowObj.studentId && studentId && rowObj.studentId.toString() === studentId.toString()) {
        studentRowIndex = i + 2; // +2 เพราะเป็น 1-indexed และเว้นบรรทัดหัวข้อ
        studentData = rowObj;
        break;
      }
    }
    
    return {
      sheet,
      headers,
      rowIndex: studentRowIndex,
      data: studentData
    };
  },

  /**
   * อัปเดตคอลัมน์ pdfUrl ในแถวที่ระบุ
   */
  updatePdfUrl(sheet, rowIndex, pdfColIdx, pdfUrl) {
    sheet.getRange(rowIndex, pdfColIdx).setValue(pdfUrl);
  }
};
