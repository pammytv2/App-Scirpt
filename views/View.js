const View = {
  /**
   * สร้างผลลัพธ์ HTML จากไฟล์เทมเพลตเพื่อแสดงผลบนเว็บเบราว์เซอร์
   */
  render(fileName, title) {
    return HtmlService.createHtmlOutputFromFile(fileName)
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  },

  /**
   * สร้างผลลัพธ์ HTML จากสตริงข้อความเพื่อแสดงผลบนเว็บเบราว์เซอร์
   */
  renderHtml(htmlContent, title) {
    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle(title)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  },

  /**
   * สร้าง PDF โดยใช้ HtmlService จากไฟล์เทมเพลต PdfTemplate.html
   */
  generatePDF(data) {
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
    const template = HtmlService.createTemplateFromFile('views/PdfTemplate');
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
};
