// --- PAGE NAVIGATION & EVENT LISTENERS ---

async function switchPage(targetPageId) {
    console.log("🔄 Switching to page:", targetPageId);
    
    document.querySelectorAll('.page-view').forEach(page => { page.classList.add('hidden'); });
    const targetPage = document.getElementById(targetPageId);
    if (targetPage) { targetPage.classList.remove('hidden'); }

    document.querySelectorAll('.nav-button').forEach(btn => {
        btn.classList.remove('active');
        if(btn.dataset.target === targetPageId) { btn.classList.add('active'); }
    });

    if (targetPageId === 'edit-page') { setTimeout(() => { setupEditPageEventListeners(); }, 100); }
    if (targetPageId === 'dashboard-page') await fetchUserRequests();
    if (targetPageId === 'form-page') { await resetRequestForm(); setTimeout(() => { tryAutoFillRequester(); }, 100); }
    if (targetPageId === 'profile-page') loadProfileData();
    if (targetPageId === 'stats-page') await loadStatsData(); // Default: Use Cache
    if (targetPageId === 'admin-users-page') await fetchAllUsers();
    if (targetPageId === 'command-generation-page') { document.getElementById('admin-view-requests-tab').click(); }
}

function setupVehicleOptions() {
    document.querySelectorAll('input[name="vehicle_option"].vehicle-checkbox').forEach(checkbox => { checkbox.addEventListener('change', toggleVehicleDetails); });
    document.querySelectorAll('input[name="edit-vehicle_option"].vehicle-checkbox').forEach(checkbox => { checkbox.addEventListener('change', toggleEditVehicleDetails); });
}

function setupEventListeners() {
    // Auth & User Management
    document.getElementById('login-form').addEventListener('submit', handleLogin);
    document.getElementById('logout-button').addEventListener('click', handleLogout);
    document.getElementById('show-register-modal-button').addEventListener('click', () => document.getElementById('register-modal').style.display = 'flex');
    document.getElementById('register-form').addEventListener('submit', handleRegister);
    document.getElementById('show-forgot-password-modal').addEventListener('click', () => { document.getElementById('forgot-password-modal').style.display = 'flex'; });
    document.getElementById('forgot-password-modal-close-button').addEventListener('click', () => { document.getElementById('forgot-password-modal').style.display = 'none'; });
    document.getElementById('forgot-password-cancel-button').addEventListener('click', () => { document.getElementById('forgot-password-modal').style.display = 'none'; });
    document.getElementById('forgot-password-form').addEventListener('submit', handleForgotPassword);
    
    // Modals
    document.getElementById('public-attendee-modal-close-button').addEventListener('click', () => { document.getElementById('public-attendee-modal').style.display = 'none'; });
    document.getElementById('public-attendee-modal-close-btn2').addEventListener('click', () => { document.getElementById('public-attendee-modal').style.display = 'none'; });
    document.querySelectorAll('.modal').forEach(modal => { modal.addEventListener('click', (e) => { if (e.target === modal) modal.style.display = 'none'; }); });
    document.getElementById('register-modal-close-button').addEventListener('click', () => document.getElementById('register-modal').style.display = 'none');
    document.getElementById('register-modal-close-button2').addEventListener('click', () => document.getElementById('register-modal').style.display = 'none');
    document.getElementById('alert-modal-close-button').addEventListener('click', () => document.getElementById('alert-modal').style.display = 'none');
    document.getElementById('alert-modal-ok-button').addEventListener('click', () => document.getElementById('alert-modal').style.display = 'none');
    document.getElementById('confirm-modal-close-button').addEventListener('click', () => document.getElementById('confirm-modal').style.display = 'none');
    
    // Admin Commands & Memos
    document.getElementById('back-to-admin-command').addEventListener('click', async () => { await switchPage('command-generation-page'); });
    document.getElementById('admin-generate-command-button').addEventListener('click', handleAdminGenerateCommand);
    document.getElementById('command-approval-form').addEventListener('submit', handleCommandApproval);
    document.getElementById('command-approval-modal-close-button').addEventListener('click', () => document.getElementById('command-approval-modal').style.display = 'none');
    document.getElementById('command-approval-cancel-button').addEventListener('click', () => document.getElementById('command-approval-modal').style.display = 'none');
    
    document.getElementById('dispatch-form').addEventListener('submit', handleDispatchFormSubmit);
    document.getElementById('dispatch-modal-close-button').addEventListener('click', () => document.getElementById('dispatch-modal').style.display = 'none');
    document.getElementById('dispatch-cancel-button').addEventListener('click', () => document.getElementById('dispatch-modal').style.display = 'none');
    
    document.getElementById('admin-memo-action-form').addEventListener('submit', handleAdminMemoActionSubmit);
    document.getElementById('admin-memo-action-modal-close-button').addEventListener('click', () => document.getElementById('admin-memo-action-modal').style.display = 'none');
    document.getElementById('admin-memo-cancel-button').addEventListener('click', () => document.getElementById('admin-memo-action-modal').style.display = 'none');
    
    document.getElementById('send-memo-modal-close-button').addEventListener('click', () => document.getElementById('send-memo-modal').style.display = 'none');
    document.getElementById('send-memo-cancel-button').addEventListener('click', () => document.getElementById('send-memo-modal').style.display = 'none');
    document.getElementById('send-memo-form').addEventListener('submit', handleMemoSubmitFromModal);

    // Stats
    document.getElementById('refresh-stats').addEventListener('click', async () => { 
        await loadStatsData(true); // Force Refresh
        showAlert('สำเร็จ', 'อัปเดตข้อมูลสถิติเรียบร้อยแล้ว'); 
    });
    document.getElementById('export-stats').addEventListener('click', exportStatsReport);

    // Navigation
    document.getElementById('navigation').addEventListener('click', async (e) => {
        const navButton = e.target.closest('.nav-button');
        if (navButton && navButton.dataset.target) { await switchPage(navButton.dataset.target); }
    });

    // Forms
    setupVehicleOptions();
    document.getElementById('admin-memo-status').addEventListener('change', function(e) {
        const fileUploads = document.getElementById('admin-file-uploads');
        if (e.target.value === 'เสร็จสิ้น/รับไฟล์ไปใช้งาน') { fileUploads.classList.remove('hidden'); } else { fileUploads.classList.add('hidden'); }
    });

    document.getElementById('request-form').addEventListener('submit', handleRequestFormSubmit);
    document.getElementById('form-add-attendee').addEventListener('click', () => addAttendeeField());
    document.getElementById('form-import-excel').addEventListener('click', () => document.getElementById('excel-file-input').click());
    document.getElementById('excel-file-input').addEventListener('change', handleExcelImport); 
    document.getElementById('form-download-template').addEventListener('click', downloadAttendeeTemplate); 
    document.querySelectorAll('input[name="expense_option"]').forEach(radio => radio.addEventListener('change', toggleExpenseOptions));
    
    document.querySelectorAll('input[name="modal_memo_type"]').forEach(radio => radio.addEventListener('change', (e) => {
        const fileContainer = document.getElementById('modal-memo-file-container');
        const fileInput = document.getElementById('modal-memo-file');
        const isReimburse = e.target.value === 'reimburse';
        fileContainer.classList.toggle('hidden', isReimburse);
        fileInput.required = !isReimburse;
    }));
    document.querySelectorAll('input[name="vehicle_option"]').forEach(checkbox => {checkbox.addEventListener('change', toggleVehicleDetails);});
    document.getElementById('profile-form').addEventListener('submit', handleProfileUpdate);
    document.getElementById('password-form').addEventListener('submit', handlePasswordUpdate);
    document.getElementById('show-password-toggle').addEventListener('change', togglePasswordVisibility);
    
    document.getElementById('form-department').addEventListener('change', (e) => {
        const selectedPosition = e.target.value;
        const headNameInput = document.getElementById('form-head-name');
        headNameInput.value = specialPositionMap[selectedPosition] || '';
    });
    
    document.getElementById('search-requests').addEventListener('input', (e) => renderRequestsList(allRequestsCache, userMemosCache, e.target.value));

    // Admin User Mgmt
    document.getElementById('add-user-button').addEventListener('click', openAddUserModal);
    document.getElementById('download-user-template-button').addEventListener('click', downloadUserTemplate);
    document.getElementById('import-users-button').addEventListener('click', () => document.getElementById('user-excel-input').click());
    document.getElementById('user-excel-input').addEventListener('change', handleUserImport);
    
    // Admin Tabs
    document.getElementById('admin-view-requests-tab').addEventListener('click', async (e) => {
        document.getElementById('admin-view-memos-tab').classList.remove('active');
        e.target.classList.add('active');
        document.getElementById('admin-requests-view').classList.remove('hidden');
        document.getElementById('admin-memos-view').classList.add('hidden');
        await fetchAllRequestsForCommand();
    });
    
    document.getElementById('admin-view-memos-tab').addEventListener('click', async (e) => {
        document.getElementById('admin-view-requests-tab').classList.remove('active');
        e.target.classList.add('active');
        document.getElementById('admin-memos-view').classList.remove('hidden');
        document.getElementById('admin-requests-view').classList.add('hidden');
        await fetchAllMemos();
    });

    // Error Handling
    window.addEventListener('error', (event) => {
        console.error('Global error:', event.error);
        if (event.error.message && event.error.message.includes('openEditPageDirect')) return;
        // Don't show alert for every error to avoid annoyance, rely on console
    });
    window.addEventListener('unhandledrejection', (event) => {
        console.error('Unhandled promise rejection:', event.reason);
        if (event.reason && event.reason.message && event.reason.message.includes('openEditPageDirect')) return;
    });
}

function handleExcelImport(e) {
    const file = e.target.files[0]; if (!file) return;
    try {
        const reader = new FileReader();
        reader.onload = async function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            const attendeesList = document.getElementById('form-attendees-list');
            attendeesList.innerHTML = '';
            
            jsonData.forEach(row => {
                if (row['ชื่อ-นามสกุล'] && row['ตำแหน่ง']) {
                    const list = document.getElementById('form-attendees-list');
                    const attendeeDiv = document.createElement('div');
                    attendeeDiv.className = 'grid grid-cols-1 md:grid-cols-3 gap-2 items-center mb-2';
                    attendeeDiv.innerHTML = `
                    <input type="text" class="form-input attendee-name md:col-span-1" value="${escapeHtml(row['ชื่อ-นามสกุล'])}" required>
                    <div class="attendee-position-wrapper md:col-span-1">
                        <select class="form-input attendee-position-select"><option value="other">อื่นๆ</option></select>
                        <input type="text" class="form-input attendee-position-other mt-1" value="${escapeHtml(row['ตำแหน่ง'])}">
                    </div>
                    <button type="button" class="btn btn-danger btn-sm" onclick="this.parentElement.remove()">ลบ</button>`;
                    list.appendChild(attendeeDiv);
                }
            });
            showAlert('สำเร็จ', 'นำเข้าข้อมูลผู้ร่วมเดินทางสำเร็จ');
        };
        reader.readAsArrayBuffer(file);
    } catch (error) { showAlert('ผิดพลาด', error.message); } finally { e.target.value = ''; }
}

function downloadAttendeeTemplate() {
    const ws = XLSX.utils.aoa_to_sheet([['ชื่อ-นามสกุล', 'ตำแหน่ง'],['ตัวอย่าง ผู้ใช้', 'ครู']]);
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'attendee_template.xlsx');
}

function enhanceEditFunctionSafety() {
    const requiredFunctions = ['openEditPage', 'generateDocumentFromDraft', 'getEditFormData'];
    requiredFunctions.forEach(funcName => {
        if (typeof window[funcName] !== 'function') {
            console.error(`Required function ${funcName} is missing`);
            window[funcName] = function() { showAlert("ระบบผิดพลาด", "ฟังก์ชันไม่พร้อมใช้งาน กรุณารีเฟรชหน้า"); };
        }
    });
}

document.addEventListener('DOMContentLoaded', () => {
    console.log('App Initializing...');
    
    // Check Config
    if (typeof escapeHtml !== 'function') {
        console.error("Config.js not loaded or missing escapeHtml!");
        alert("System Error: Configuration missing. Please refresh.");
        return;
    }

    loadPublicWeeklyData();
    setupEventListeners();
    enhanceEditFunctionSafety();
    
    Chart.defaults.font.family = "'Sarabun', sans-serif";
    Chart.defaults.font.size = 14;
    Chart.defaults.color = '#374151';
    
    const navEdit = document.getElementById('nav-edit');
    if (navEdit) navEdit.classList.add('hidden');
    
    if (typeof resetEditPage === 'function') resetEditPage();
    
    const user = getCurrentUser();
    if (user) { initializeUserSession(user); } else { showLoginScreen(); }
});
