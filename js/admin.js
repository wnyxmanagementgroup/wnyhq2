// --- ADMIN FUNCTIONS ---

function checkAdminAccess() {
    const user = getCurrentUser();
    if (!user || user.role !== 'admin') {
        showAlert('ผิดพลาด', 'คุณไม่มีสิทธิ์เข้าถึงหน้านี้');
        return false;
    }
    return true;
}

async function fetchAllRequestsForCommand() {
    try {
        if (!checkAdminAccess()) return;
        const result = await apiCall('GET', 'getAllRequests');
        if (result.status === 'success') renderAdminRequestsList(result.data);
    } catch (error) { showAlert('ผิดพลาด', 'ไม่สามารถโหลดข้อมูลคำขอได้'); }
}

async function fetchAllMemos() {
    try {
        if (!checkAdminAccess()) return;
        const result = await apiCall('GET', 'getAllMemos');
        if (result.status === 'success') renderAdminMemosList(result.data);
    } catch (error) { showAlert('ผิดพลาด', 'ไม่สามารถโหลดข้อมูลบันทึกข้อความได้'); }
}

async function fetchAllUsers() {
    try {
        if (!checkAdminAccess()) return;
        const result = await apiCall('GET', 'getAllUsers');
        if (result.status === 'success') { 
            allUsersCache = result.data; 
            renderUsersList(allUsersCache); 
        }
    } catch (error) { showAlert('ผิดพลาด', 'ไม่สามารถโหลดข้อมูลผู้ใช้ได้'); }
}

// ... (openAdminGenerateCommand, addAdminAttendeeField, handleAdminGenerateCommand - Logic เดิม) ...

function renderUsersList(users) {
    const container = document.getElementById('users-content');
    if (!users || users.length === 0) { 
        container.innerHTML = '<p class="text-center text-gray-500">ไม่พบข้อมูลผู้ใช้</p>'; 
        return; 
    }
    
    // Responsive Table Structure with data-labels and Sanitization
    container.innerHTML = `
    <div class="overflow-x-auto">
        <table class="min-w-full bg-white responsive-table">
            <thead>
                <tr class="bg-gray-100">
                    <th class="px-4 py-2 text-left">ชื่อผู้ใช้</th>
                    <th class="px-4 py-2 text-left">ชื่อ-นามสกุล</th>
                    <th class="px-4 py-2 text-left">ตำแหน่ง</th>
                    <th class="px-4 py-2 text-left">กลุ่มสาระ/งาน</th>
                    <th class="px-4 py-2 text-left">บทบาท</th>
                    <th class="px-4 py-2 text-left">การจัดการ</th>
                </tr>
            </thead>
            <tbody>
                ${users.map(user => `
                <tr class="border-b">
                    <td class="px-4 py-2" data-label="ชื่อผู้ใช้">${escapeHtml(user.username)}</td>
                    <td class="px-4 py-2" data-label="ชื่อ-นามสกุล">${escapeHtml(user.fullName)}</td>
                    <td class="px-4 py-2" data-label="ตำแหน่ง">${escapeHtml(user.position)}</td>
                    <td class="px-4 py-2" data-label="กลุ่มสาระ">${escapeHtml(user.department)}</td>
                    <td class="px-4 py-2" data-label="บทบาท">${escapeHtml(user.role)}</td>
                    <td class="px-4 py-2" data-label="การจัดการ">
                        <button onclick="deleteUser('${escapeHtml(user.username)}')" class="btn btn-danger btn-sm">ลบ</button>
                    </td>
                </tr>`).join('')}
            </tbody>
        </table>
    </div>`;
}

// ... (deleteUser, openAddUserModal, downloadUserTemplate, handleUserImport - Logic เดิม) ...

function renderAdminRequestsList(requests) {
    const container = document.getElementById('admin-requests-list');
    if (!requests || requests.length === 0) { container.innerHTML = '<p class="text-center text-gray-500">ไม่พบคำขอไปราชการ</p>'; return; }
    
    container.innerHTML = requests.map(request => {
        const attendeeCount = request.attendeeCount || 0;
        const totalPeople = attendeeCount + 1;
        let peopleCategory = totalPeople === 1 ? "คำสั่งเดี่ยว (1 คน)" : (totalPeople <= 5 ? "คำสั่งกลุ่มเล็ก (2-5 คน)" : "คำสั่งกลุ่มใหญ่ (6 คนขึ้นไป)");
        
        // Sanitization
        const safeId = escapeHtml(request.id);
        const safeName = escapeHtml(request.requesterName);
        const safePurpose = escapeHtml(request.purpose);
        const safeLocation = escapeHtml(request.location);
        const safeDate = `${formatDisplayDate(request.startDate)} - ${formatDisplayDate(request.endDate)}`;

        return `
        <div class="border rounded-lg p-4 bg-white">
            <div class="flex justify-between items-start flex-wrap gap-4">
                <div class="flex-1 min-w-[200px]">
                    <h4 class="font-bold text-indigo-700">${safeId}</h4>
                    <p class="text-sm text-gray-600">โดย: ${safeName} | ${safePurpose}</p>
                    <p class="text-sm text-gray-500">${safeLocation} | ${safeDate}</p>
                    <p class="text-sm text-gray-700">ผู้ร่วมเดินทาง: ${attendeeCount} คน</p>
                    <p class="text-sm font-medium text-blue-700">👥 รวมทั้งหมด: ${totalPeople} คน (${peopleCategory})</p>
                    <p class="text-sm">สถานะคำขอ: <span class="font-medium">${translateStatus(request.status)}</span></p>
                    <p class="text-sm">สถานะคำสั่ง: <span class="font-medium">${request.commandStatus || 'รอดำเนินการ'}</span></p>
                </div>
                <div class="flex flex-col gap-2 w-full sm:w-auto">
                    ${request.pdfUrl ? `<a href="${request.pdfUrl}" target="_blank" class="btn btn-success btn-sm">ดูคำขอ</a>` : ''}
                    <div class="flex gap-1 flex-wrap">
                        ${request.commandPdfUrl ? 
                            `<a href="${request.commandPdfUrl}" target="_blank" class="btn bg-blue-500 text-white btn-sm">ดูคำสั่ง</a>` : 
                            `<button onclick="openAdminGenerateCommand('${safeId}')" class="btn bg-green-500 text-white btn-sm">ออกคำสั่ง</button>`
                        }
                        ${!request.dispatchBookPdfUrl ? `<button onclick="openDispatchModal('${safeId}')" class="btn bg-orange-500 text-white btn-sm">ออกหนังสือส่ง</button>` : ''}
                    </div>
                </div>
            </div>
        </div>`;
    }).join('');
}

function renderAdminMemosList(memos) {
    const container = document.getElementById('admin-memos-list');
    if (!memos || memos.length === 0) { container.innerHTML = '<p class="text-center text-gray-500">ไม่พบบันทึกข้อความ</p>'; return; }
    
    container.innerHTML = memos.map(memo => {
        const hasCompletedFiles = memo.completedMemoUrl || memo.completedCommandUrl || memo.dispatchBookUrl;
        const safeId = escapeHtml(memo.id);
        const safeRef = escapeHtml(memo.refNumber);
        const safeUser = escapeHtml(memo.submittedBy);

        return `
        <div class="border rounded-lg p-4 bg-white">
            <div class="flex justify-between items-start flex-wrap gap-4">
                <div class="flex-1">
                    <h4 class="font-bold">${safeId}</h4>
                    <p class="text-sm text-gray-600">โดย: ${safeUser} | อ้างอิง: ${safeRef}</p>
                    <p class="text-sm">สถานะ: <span class="font-medium">${translateStatus(memo.status)}</span></p>
                    <div class="mt-2 text-xs text-gray-500">
                        ${memo.completedMemoUrl ? `<div>✓ บันทึกข้อความสมบูรณ์</div>` : ''}
                        ${memo.completedCommandUrl ? `<div>✓ คำสั่งสมบูรณ์</div>` : ''}
                        ${memo.dispatchBookUrl ? `<div>✓ หนังสือส่งสมบูรณ์</div>` : ''}
                    </div>
                </div>
                <div class="flex flex-col gap-2 w-full sm:w-auto">
                    ${memo.fileURL ? `<a href="${memo.fileURL}" target="_blank" class="btn btn-success btn-sm">ดูไฟล์ต้นทาง</a>` : ''}
                    ${memo.completedMemoUrl ? `<a href="${memo.completedMemoUrl}" target="_blank" class="btn bg-blue-500 text-white btn-sm">ดูบันทึกสมบูรณ์</a>` : ''}
                    ${memo.completedCommandUrl ? `<a href="${memo.completedCommandUrl}" target="_blank" class="btn bg-blue-500 text-white btn-sm">ดูคำสั่งสมบูรณ์</a>` : ''}
                    ${memo.dispatchBookUrl ? `<a href="${memo.dispatchBookUrl}" target="_blank" class="btn bg-purple-500 text-white btn-sm">ดูหนังสือส่ง</a>` : ''}
                    <button onclick="openAdminMemoAction('${safeId}')" class="btn bg-green-500 text-white btn-sm">${hasCompletedFiles ? 'จัดการไฟล์' : 'อัพโหลดไฟล์'}</button>
                </div>
            </div>
        </div>`;
    }).join('');
}

// ... (openCommandApproval, openDispatchModal, openAdminMemoAction, handleCommandApproval, handleDispatchFormSubmit, handleAdminMemoActionSubmit, sendCompletionEmail - Logic เดิม) ...
