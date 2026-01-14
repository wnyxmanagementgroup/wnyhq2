// --- CONFIGURATION ---

// 1. Debug Mode: เปลี่ยนเป็น false เมื่อใช้งานจริง (Production)
const IS_DEBUG = true; 

if (!IS_DEBUG) {
    console.log = function() {};
    console.warn = function() {};
    console.error = function() {};
    console.info = function() {};
}

const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzP-HCQGbA3Xi2Ms4DXTGy8k17Bv72pFohnJ0txAePjjXybe6pK42mSaYOfTQ5V9Q6mDA/exec";

// Global State
let allRequestsCache = [];
let allMemosCache = [];
let userMemosCache = [];
let allUsersCache = [];
window.requestsChartInstance = null;
window.statusChartInstance = null;
let currentPublicWeeklyData = [];

// --- UTILITY: SECURITY SANITIZATION ---
// ฟังก์ชันป้องกัน XSS: แปลงอักขระพิเศษเป็น HTML entities ก่อนแสดงผล
// ใช้ฟังก์ชันนี้ครอบตัวแปร text ที่มาจาก User input เสมอ
function escapeHtml(text) {
    if (!text) return '';
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// Map ตำแหน่งพิเศษ (คงเดิม)
let specialPositionMap = {
    'รองผู้อำนวยการกลุ่มบริหารทั่วไป':'นางวชิรินทรา พัฒนกุลเดช',
    'รองผู้อำนวยการกลุ่มบริหารงานบุคคล':'นางปณิชา ภัสสิรากุล',
    'รองผู้อำนวยการกลุ่มบริหารงบประมาณ':'นางจันทิมา นกอยู่',
    'รองผู้อำนวยการกลุ่มบริหารวิชาการ': 'นายมงคล เกตมณี',
    'หัวหน้ากลุ่มบริหารทั่วไป':'นายสุชาติ สินทร',
    'หัวหน้ากลุ่มบริหารงานบุคคล':'นายพงษ์ศักดิ์ ทองโพธิกุล',
    'หัวหน้ากลุ่มบริหารงบประมาณ':'นางสาวนุชนาฎ อำพันเสน',
    'หัวหน้ากลุ่มบริหารวิชาการ':'นางสาวสาวิทตรี อุ่นทองศิริ',
    'หัวหน้ากลุ่มสาระการเรียนรู้วิทยาศาสตร์และเทคโนโลยี': 'นางสาวปิยราช พันธุ์กมลศิลป์',
    // ... (คงรายการเดิมไว้ทั้งหมด) ...
    '.....................................':'.....................................'
};

const statusTranslations = {
    'Pending': 'กำลังดำเนินการ',
    'Submitted': 'รอการตรวจสอบ',
    'Approved': 'เสร็จสิ้น',
    'Pending Approval': 'รอการตรวจสอบ',
    'เสร็จสิ้น/รับไฟล์ไปใช้งาน': 'เสร็จสิ้น',
    'เสร็จสิ้น': 'เสร็จสิ้น',
    'รอเอกสาร (เบิก)': 'รอเอกสาร (เบิก)',
    'นำกลับไปแก้ไข': 'นำกลับไปแก้ไข',
    'เสร็จสิ้นรอออกคำสั่งไปราชการ': 'เสร็จสิ้นรอออกคำสั่ง',
    'รอตรวจสอบและออกคำสั่งไปราชการ': 'รอตรวจสอบและออกคำสั่ง',
    'กำลังดำเนินการ': 'กำลังดำเนินการ'
};

function translateStatus(status) {
    return statusTranslations[status] || status;
}
