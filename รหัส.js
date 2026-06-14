/**
 * Google Apps Script - Student Profile System (ใบประวัตินักศึกษา)
 * Main Entry Point - Router File (MVC Architecture)
 * Exposes global functions required by Apps Script engine & client-side google.script.run
 */

/**
 * Web App entry point (HTTP GET request handler)
 */
function doGet(e) {
  return Controller.handleGetRequest(e);
}

/**
 * รับข้อมูลจากฟอร์มและดำเนินการบันทึกข้อมูลและสร้าง PDF
 */
function processForm(formData) {
  return Controller.processFormSubmit(formData);
}

/**
 * ดึงข้อมูลนักศึกษาทั้งหมดจาก Google Sheet เพื่อแสดงในหน้า Admin
 */
function getStudentSubmissions(pwd) {
  return Model.getStudentSubmissions(pwd);
}

/**
 * บังคับสร้างไฟล์ PDF ใหม่ของนักศึกษาจากข้อมูลที่มีอยู่ในชีต
 */
function regenerateStudentPDF(pwd, studentId) {
  return Controller.regeneratePDF(pwd, studentId);
}

/**
 * ฟังก์ชันสำหรับตรวจสอบความถูกต้องของรหัสผ่านผู้ดูแลระบบ (Admin)
 */
function verifyAdminPassword(pwd) {
  return Controller.verifyAdminPassword(pwd);
}

/**
 * ดึง URL ของ Web App ปัจจุบัน เพื่อใช้ในการนำทางแบบสองทิศทางใน iframe
 */
function getScriptURL() {
  return Controller.getScriptURL();
}

/**
 * QA Test Runner - ตรวจสอบความถูกต้องและทดสอบระบบทั้งหมด
 */
function runQATests() {
  return QAService.runQATests();
}

/**
 * สร้างข้อมูลจำลอง 3 รายการเพื่อใช้ในการทดสอบระบบ Admin
 */
function createThreeMockSubmissions() {
  return QAService.createThreeMockSubmissions();
}

/**
 * ฟังก์ชันตรวจสอบและวินิจฉัยปัญหาการสร้างไฟล์ PDF
 */
function runPdfDiagnostic() {
  return QAService.runPdfDiagnostic();
}

/**
 * ฟังก์ชันสำหรับดีบั๊ก ตรวจสอบความยาวตัวอักษรของเทมเพลต Google Docs
 */
function inspectTemplateDoc() {
  return QAService.inspectTemplateDoc();
}

/**
 * ฟังก์ชันสำหรับออกแบบและสร้างเทมเพลต Google Docs จากศูนย์
 */
function createTemplateDoc() {
  return QAService.createTemplateDoc();
}
