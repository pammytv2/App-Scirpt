const Controller = {
  /**
   * จัดการกับ HTTP GET Requests (doGet Router)
   */
  handleGetRequest(e) {
    const mode = e && e.parameter && e.parameter.mode;
    const action = e && e.parameter && e.parameter.action;
    
    if (mode === 'create_template') {
      return View.renderHtml(
        '<h3>ไม่ต้องสร้างเทมเพลตแล้วครับ!</h3><p>ระบบได้ถูกอัปเกรดให้ใช้ HTML Template ในการสร้าง PDF โดยตรงแล้ว ไม่ต้องพึ่งพาการวาดตารางใน Google Docs อีกต่อไป ระบบจะจัดฟอร์มและสร้าง PDF ให้เป๊ะขึ้นและทำงานเร็วขึ้นครับ</p>',
        'ไม่ต้องสร้างเทมเพลตแล้วครับ'
      );
    }
    
    if (mode === 'test_pdf' || mode === 'inspect') {
      return View.renderHtml(
        '<h3>ไม่มีความจำเป็นต้องตรวจสอบเทมเพลต Google Docs แล้วครับ</h3><p>ปัจจุบันระบบใช้ HTML ในการแปลง PDF แทนการดึงจาก Google Docs</p>',
        'ไม่มีความจำเป็นต้องตรวจสอบเทมเพลต'
      );
    }
    
    if (mode === 'admin') {
      return View.render('views/Admin', 'ระบบจัดการข้อมูลนักศึกษา (Admin)');
    }
    
    if (mode === 'qa') {
      if (action === 'create_mock') {
        try {
          const results = QAService.createThreeMockSubmissions();
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
          return View.renderHtml(html, 'สร้างข้อมูลจำลองสำเร็จ');
        } catch (err) {
          return View.renderHtml('เกิดข้อผิดพลาดในการสร้างข้อมูลจำลอง: ' + err.toString(), 'เกิดข้อผิดพลาด');
        }
      }
      return QAService.runQATests();
    }
    
    return View.render('views/Index', 'ระบบบันทึกใบประวัตินักศึกษา');
  },

  /**
   * รับข้อมูลจากฟอร์ม, ดำเนินการสร้าง PDF และบันทึกลงชีต
   */
  processFormSubmit(formData) {
    try {
      // 1. สร้างไฟล์ PDF จากข้อมูลที่ส่งมา
      const pdfUrl = View.generatePDF(formData);
      formData.pdfUrl = pdfUrl;

      // 2. จัดเก็บแถวข้อมูลลงใน Sheet
      Model.saveStudent(formData);

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
  },

  /**
   * บังคับสร้างไฟล์ PDF ใหม่ของนักศึกษาจากข้อมูลที่มีอยู่ในชีต
   */
  regeneratePDF(pwd, studentId) {
    const inputPwd = (pwd || "").toString().trim();
    if (inputPwd !== CONFIG.ADMIN_PASSWORD) {
      throw new Error("สิทธิ์การเข้าถึงปฏิเสธ: รหัสผ่านแอดมินไม่ถูกต้อง");
    }

    try {
      // ค้นหาแถวและข้อมูลนักศึกษา
      const searchResult = Model.findStudentRowIndexAndData(studentId);
      const sheet = searchResult.sheet;
      const headers = searchResult.headers;
      const studentRowIndex = searchResult.rowIndex;
      const studentData = searchResult.data;
      
      if (!studentData) {
        throw new Error("ไม่พบข้อมูลนักศึกษารหัสดังกล่าว");
      }
      
      // แปลงค่า 'ใช่' หรือ 'ไม่' ใน Sheet กลับเป็น Boolean
      for (let key in studentData) {
        if (studentData[key] === "ใช่") {
          studentData[key] = true;
        } else if (studentData[key] === "ไม่") {
          studentData[key] = false;
        }
      }
      
      // สร้าง PDF ใหม่
      const pdfUrl = View.generatePDF(studentData);
      
      // อัปเดตคอลัมน์ pdfUrl ในชีต
      const pdfColIdx = headers.indexOf('pdfUrl') + 1;
      if (pdfColIdx > 0) {
        Model.updatePdfUrl(sheet, studentRowIndex, pdfColIdx, pdfUrl);
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
  },

  /**
   * ตรวจสอบรหัสผ่านแอดมิน
   */
  verifyAdminPassword(pwd) {
    return pwd === CONFIG.ADMIN_PASSWORD;
  },

  /**
   * ดึง URL ของ Apps Script Web App ปัจจุบัน
   */
  getScriptURL() {
    return ScriptApp.getService().getUrl();
  }
};
