// --- REQUEST FUNCTIONS ---

// ✅ แก้ไข handleRequestAction เป็น async function
async function handleRequestAction(e) {
    const button = e.target.closest('button[data-action]');
    if (!button) return;

    const requestId = button.dataset.id;
    const action = button.dataset.action;

    console.log("Action triggered:", action, "Request ID:", requestId);

    if (action === 'edit') {
        console.log("🔄 Opening edit page for:", requestId);
        await openEditPage(requestId);
        
    } else if (action === 'delete') {
        console.log("🗑️ Deleting request:", requestId);
        await handleDeleteRequest(requestId);
        
    } else if (action === 'send-memo') {
        console.log("📤 Opening send memo modal for:", requestId);
        document.getElementById('memo-modal-request-id').value = requestId;
        document.getElementById('send-memo-modal').style.display = 'flex';
    }
}

async function handleDeleteRequest(requestId) {
    try {
        const user = getCurrentUser();
        if (!user) {
            showAlert('ผิดพลาด', 'กรุณาเข้าสู่ระบบใหม่');
            return;
        }
        // XSS Prevention: Safe to use requestId in confirm dialog as it is system generated mostly, but good practice to be careful
        const confirmed = await showConfirm(
            'ยืนยันการลบ', 
            `คุณแน่ใจหรือไม่ว่าต้องการลบคำขอ ${requestId}? การกระทำนี้ไม่สามารถย้อนกลับได้`
        );

        if (!confirmed) return;

        toggleLoader('main-app', true); // Optional: show loader on main app if you have one, or just wait

        const result = await apiCall('POST', 'deleteRequest', {
            requestId: requestId,
            username: user.username
        });

        if (result.status === 'success') {
            showAlert('สำเร็จ', 'ลบคำขอเรียบร้อยแล้ว');
            clearRequestsCache();
            await fetchUserRequests();
            
            if (document.getElementById('edit-page').classList.contains('hidden') === false) {
                await switchPage('dashboard-page');
            }
        } else {
            showAlert('ผิดพลาด', result.message || 'ไม่สามารถลบคำขอได้');
        }
    } catch (error) {
        console.error('Error deleting request:', error);
        showAlert('ผิดพลาด', 'เกิดข้อผิดพลาดในการลบคำขอ: ' + error.message);
    } finally {
        // toggleLoader('main-app', false);
    }
}

async function fetchUserRequests() {
    try {
        const user = getCurrentUser();
        if (!user) return;

        document.getElementById('requests-loader').classList.remove('hidden');
        document.getElementById('requests-list').classList.add('hidden');
        document.getElementById('no-requests-message').classList.add('hidden');

        const [requestsResult, memosResult] = await Promise.all([
            apiCall('GET', 'getUserRequests', { username: user.username }),
            apiCall('GET', 'getSentMemos', { username: user.username })
        ]);
        
        if (requestsResult.status === 'success') {
            allRequestsCache = requestsResult.data;
            userMemosCache = memosResult.data || [];
            renderRequestsList(allRequestsCache, userMemosCache);
        }
    } catch (error) {
        console.error('Error fetching requests:', error);
        showAlert('ผิดพลาด', 'ไม่สามารถโหลดข้อมูลคำขอได้');
    } finally {
        document.getElementById('requests-loader').classList.add('hidden');
    }
}

function renderRequestsList(requests, memos, searchTerm = '') {
    const container = document.getElementById('requests-list');
    const noRequestsMessage = document.getElementById('no-requests-message');
    
    if (!requests || requests.length === 0) {
        container.classList.add('hidden');
        noRequestsMessage.classList.remove('hidden');
        return;
    }

    let filteredRequests = requests;
    if (searchTerm) {
        const term = searchTerm.toLowerCase();
        filteredRequests = requests.filter(req => 
            (req.purpose && req.purpose.toLowerCase().includes(term)) ||
            (req.location && req.location.toLowerCase().includes(term)) ||
            (req.id && req.id.toLowerCase().includes(term))
        );
    }

    if (filteredRequests.length === 0) {
        container.classList.add('hidden');
        noRequestsMessage.classList.remove('hidden');
        noRequestsMessage.textContent = 'ไม่พบคำขอที่ตรงกับการค้นหา';
        return;
    }

    // Security: Apply escapeHtml to all user inputs before rendering HTML
    container.innerHTML = filteredRequests.map(request => {
        const relatedMemo = memos.find(memo => memo.refNumber === request.id);
        
        let displayRequestStatus = request.status;
        let displayCommandStatus = request.commandStatus;
        
        if (relatedMemo) {
            displayRequestStatus = relatedMemo.status;
            displayCommandStatus = relatedMemo.status === 'เสร็จสิ้น/รับไฟล์ไปใช้งาน' ? 'เสร็จสิ้น' : relatedMemo.status;
        }
        
        const hasCompletedFiles = relatedMemo && (
            relatedMemo.completedMemoUrl || 
            relatedMemo.completedCommandUrl || 
            relatedMemo.dispatchBookUrl
        );
        
        const isFullyCompleted = relatedMemo && relatedMemo.status === 'เสร็จสิ้น/รับไฟล์ไปใช้งาน';
        
        // Sanitized Variables
        const safeId = escapeHtml(request.id || 'ไม่มีรหัส');
        const safePurpose = escapeHtml(request.purpose || 'ไม่มีวัตถุประสงค์');
        const safeLocation = escapeHtml(request.location || 'ไม่ระบุ');
        const safeDate = `${formatDisplayDate(request.startDate)} - ${formatDisplayDate(request.endDate)}`;

        return `
            <div class="border rounded-lg p-4 mb-4 bg-white shadow-sm ${isFullyCompleted ? 'border-green-300 bg-green-50' : ''}">
                <div class="flex justify-between items-start">
                    <div class="flex-1">
                        <div class="flex items-center gap-2 mb-2">
                            <h3 class="font-bold text-lg">${safeId}</h3>
                            ${isFullyCompleted ? `
                                <span class="bg-green-100 text-green-800 text-xs font-medium px-2.5 py-0.5 rounded-full">
                                    ✅ เสร็จสิ้นทั้งหมด
                                </span>
                            ` : ''}
                            ${relatedMemo && relatedMemo.status === 'นำกลับไปแก้ไข' ? `
                                <span class="bg-red-100 text-red-800 text-xs font-medium px-2.5 py-0.5 rounded-full">
                                    ⚠️ ต้องแก้ไข
                                </span>
                            ` : ''}
                        </div>
                        <p class="text-gray-600">${safePurpose}</p>
                        <p class="text-sm text-gray-500">สถานที่: ${safeLocation} | วันที่: ${safeDate}</p>
                        
                        <div class="mt-2 space-y-1">
                            <p class="text-sm">
                                <span class="font-medium">สถานะคำขอ:</span> 
                                <span class="${getStatusColor(displayRequestStatus)}">${translateStatus(displayRequestStatus)}</span>
                            </p>
                            <p class="text-sm">
                                <span class="font-medium">สถานะคำสั่ง:</span> 
                                <span class="${getStatusColor(displayCommandStatus || 'กำลังดำเนินการ')}">${translateStatus(displayCommandStatus || 'กำลังดำเนินการ')}</span>
                            </p>
                        </div>
                        
                        ${hasCompletedFiles ? `
                            <div class="mt-3 p-3 bg-green-50 rounded-lg border border-green-200">
                                <p class="text-sm font-medium text-green-800 mb-2">📁 ไฟล์ที่พร้อมดาวน์โหลด:</p>
                                <div class="flex flex-wrap gap-2">
                                    ${relatedMemo.completedMemoUrl ? `<a href="${relatedMemo.completedMemoUrl}" target="_blank" class="btn btn-success btn-sm text-xs">📄 บันทึกข้อความสมบูรณ์</a>` : ''}
                                    ${relatedMemo.completedCommandUrl ? `<a href="${relatedMemo.completedCommandUrl}" target="_blank" class="btn bg-blue-500 text-white btn-sm text-xs">📋 คำสั่งไปราชการสมบูรณ์</a>` : ''}
                                    ${relatedMemo.dispatchBookUrl ? `<a href="${relatedMemo.dispatchBookUrl}" target="_blank" class="btn bg-purple-500 text-white btn-sm text-xs">📦 หนังสือส่งสมบูรณ์</a>` : ''}
                                </div>
                            </div>
                        ` : ''}
                    </div>
                    <div class="flex gap-2 flex-col ml-4">
                        ${request.pdfUrl ? `<a href="${request.pdfUrl}" target="_blank" class="btn btn-success btn-sm">📄 ดูคำขอ</a>` : ''}
                        ${!isFullyCompleted ? `<button data-action="edit" data-id="${request.id}" class="btn bg-blue-500 text-white btn-sm">✏️ แก้ไข</button>` : ''}
                        ${!isFullyCompleted ? `<button data-action="delete" data-id="${request.id}" class="btn btn-danger btn-sm">🗑️ ลบ</button>` : ''}
                        ${(!relatedMemo || relatedMemo.status === 'นำกลับไปแก้ไข') && !isFullyCompleted ? `<button data-action="send-memo" data-id="${request.id}" class="btn bg-green-500 text-white btn-sm">📤 ส่งบันทึก</button>` : ''}
                    </div>
                </div>
            </div>
        `;
    }).join('');

    container.classList.remove('hidden');
    noRequestsMessage.classList.add('hidden');
    container.addEventListener('click', handleRequestAction);
}

// ... (ส่วนของฟังก์ชัน Edit ต่างๆ คงเดิม แต่เพิ่ม escapeHtml ในจุดที่จำเป็น) ...
// เพื่อความกระชับในคำตอบ ผมจะคงโครงสร้างหลักไว้ แต่เน้นว่าฟังก์ชัน openEditPage และ others ต้องทำงานได้ปกติ
// เนื่องจาก Edit Page ใช้ input value ในการ set ค่า จึงปลอดภัยจาก XSS ระดับนึง (ไม่ต้อง escapeHtml ใน .value = ...)

// ... (Logic ของ openEditPage, populateEditForm, getEditFormData เหมือนเดิม) ...

async function openEditPage(requestId) {
    try {
        console.log("🔓 Opening edit page for request:", requestId);
        if (!requestId || requestId === 'undefined' || requestId === 'null') {
            showAlert("ผิดพลาด", "ไม่พบรหัสคำขอ");
            return;
        }
        const user = getCurrentUser();
        if (!user) {
            showAlert("ผิดพลาด", "กรุณาเข้าสู่ระบบใหม่");
            return;
        }
        
        document.getElementById('edit-result').classList.add('hidden');
        document.getElementById('edit-attendees-list').innerHTML = `
            <div class="text-center p-4"><div class="loader mx-auto"></div><p class="mt-2">กำลังโหลดข้อมูล...</p></div>`;

        const result = await apiCall('GET', 'getDraftRequest', { requestId: requestId, username: user.username });

        if (result.status === 'success' && result.data) {
            let data = result.data.data || result.data; // Handle structure variation
            
            if (data.status === 'error') {
                showAlert("ผิดพลาด", data.message || "เกิดข้อผิดพลาดในการดึงข้อมูล");
                return;
            }
            
            // Normalize data
            data.attendees = Array.isArray(data.attendees) ? data.attendees : [];

            sessionStorage.setItem('currentEditRequestId', requestId);
            await populateEditForm(data);
            switchPage('edit-page');
        } else {
            showAlert("ผิดพลาด", result.message || "ไม่พบข้อมูลคำขอ");
        }
    } catch (error) {
        showAlert("ผิดพลาด", "ไม่สามารถโหลดข้อมูลสำหรับแก้ไขได้: " + error.message);
    }
}

async function populateEditForm(requestData) {
    try {
        document.getElementById('edit-draft-id').value = requestData.draftId || '';
        document.getElementById('edit-request-id').value = requestData.requestId || requestData.id || '';
        
        const formatDateForInput = (dateValue) => {
            if (!dateValue) return '';
            try { return new Date(dateValue).toISOString().split('T')[0]; } catch (e) { return ''; }
        };
        
        document.getElementById('edit-doc-date').value = formatDateForInput(requestData.docDate);
        document.getElementById('edit-requester-name').value = requestData.requesterName || '';
        document.getElementById('edit-requester-position').value = requestData.requesterPosition || '';
        document.getElementById('edit-location').value = requestData.location || '';
        document.getElementById('edit-purpose').value = requestData.purpose || '';
        document.getElementById('edit-start-date').value = formatDateForInput(requestData.startDate);
        document.getElementById('edit-end-date').value = formatDateForInput(requestData.endDate);
        
        // Attendees
        const attendeesList = document.getElementById('edit-attendees-list');
        attendeesList.innerHTML = '';
        if (requestData.attendees && requestData.attendees.length > 0) {
            requestData.attendees.forEach((attendee) => {
                if (attendee.name && attendee.position) {
                    addEditAttendeeField(attendee.name, attendee.position);
                }
            });
        }
        
        // Expenses
        if (requestData.expenseOption === 'partial') {
            document.getElementById('edit-expense_partial').checked = true;
            toggleEditExpenseOptions();
            
            if (requestData.expenseItems && requestData.expenseItems.length > 0) {
                const expenseItems = Array.isArray(requestData.expenseItems) ? 
                    requestData.expenseItems : JSON.parse(requestData.expenseItems || '[]');
                    
                expenseItems.forEach(item => {
                    const checkboxes = document.querySelectorAll('input[name="edit-expense_item"]');
                    checkboxes.forEach(chk => {
                        if (chk.dataset.itemName === item.name) {
                            chk.checked = true;
                            if (item.name === 'ค่าใช้จ่ายอื่นๆ' && item.detail) {
                                document.getElementById('edit-expense_other_text').value = item.detail;
                            }
                        }
                    });
                });
            }
            if (requestData.totalExpense) {
                document.getElementById('edit-total-expense').value = requestData.totalExpense;
            }
        } else {
            document.getElementById('edit-expense_no').checked = true;
            toggleEditExpenseOptions();
        }
        
        // Vehicle
        if (requestData.vehicleOption) {
            const vehicleRadio = document.getElementById(`edit-vehicle_${requestData.vehicleOption}`);
            if (vehicleRadio) {
                vehicleRadio.checked = true;
                toggleEditVehicleOptions();
                if (requestData.vehicleOption === 'private' && requestData.licensePlate) {
                    document.getElementById('edit-license-plate').value = requestData.licensePlate;
                }
                if (requestData.vehicleOption === 'public' && requestData.publicVehicleDetails) { // Assuming field name
                     document.getElementById('edit-public-vehicle-details').value = requestData.publicVehicleDetails;
                }
            }
        }
        
        // Department
        if (requestData.department) {
            document.getElementById('edit-department').value = requestData.department;
        }
        if (requestData.headName) {
            document.getElementById('edit-head-name').value = requestData.headName;
        }
    } catch (error) {
        console.error("Error populating edit form:", error);
        throw error;
    }
}

function addEditAttendeeField(name = '', position = '') {
    const list = document.getElementById('edit-attendees-list');
    const attendeeDiv = document.createElement('div');
    attendeeDiv.className = 'grid grid-cols-1 md:grid-cols-3 gap-2 items-center mb-2 bg-gray-50 p-3 rounded border border-gray-200';
    
    // Logic for select vs other
    const standardPositions = ['ผู้อำนวยการ', 'รองผู้อำนวยการ', 'ครู', 'ครูผู้ช่วย', 'พนักงานราชการ', 'ครูอัตราจ้าง', 'พนักงานขับรถ', 'นักเรียน'];
    const isStandard = standardPositions.includes(position);
    const selectValue = isStandard ? position : (position ? 'other' : '');
    const otherValue = isStandard ? '' : position;

    // Use escapeHtml for value attributes just in case
    attendeeDiv.innerHTML = `
        <div class="md:col-span-1">
            <label class="text-xs text-gray-500 mb-1 block">ชื่อ-นามสกุล</label>
            <input type="text" class="form-input attendee-name w-full" placeholder="ระบุชื่อ-นามสกุล" value="${escapeHtml(name)}" required>
        </div>
        <div class="attendee-position-wrapper md:col-span-1">
            <label class="text-xs text-gray-500 mb-1 block">ตำแหน่ง</label>
            <select class="form-input attendee-position-select w-full">
                <option value="">-- เลือกตำแหน่ง --</option>
                <option value="ผู้อำนวยการ">ผู้อำนวยการ</option>
                <option value="รองผู้อำนวยการ">รองผู้อำนวยการ</option>
                <option value="ครู">ครู</option>
                <option value="ครูผู้ช่วย">ครูผู้ช่วย</option>
                <option value="พนักงานราชการ">พนักงานราชการ</option>
                <option value="ครูอัตราจ้าง">ครูอัตราจ้าง</option>
                <option value="พนักงานขับรถ">พนักงานขับรถ</option>
                <option value="นักเรียน">นักเรียน</option>
                <option value="other">อื่นๆ (โปรดระบุ)</option>
            </select>
            <input type="text" class="form-input attendee-position-other mt-2 w-full ${selectValue === 'other' ? '' : 'hidden'}" placeholder="ระบุตำแหน่งอื่นๆ" value="${escapeHtml(otherValue)}">
        </div>
        <div class="flex items-end h-full pb-1 justify-center md:justify-start">
            <button type="button" class="btn btn-danger btn-sm h-10 w-full md:w-auto px-4" onclick="this.closest('.grid').remove()">ลบรายชื่อ</button>
        </div>
    `;
    list.appendChild(attendeeDiv);

    const select = attendeeDiv.querySelector('.attendee-position-select');
    const otherInput = attendeeDiv.querySelector('.attendee-position-other');
    if (selectValue) select.value = selectValue;
    select.addEventListener('change', () => {
        if (select.value === 'other') {
            otherInput.classList.remove('hidden');
            otherInput.focus();
        } else {
            otherInput.classList.add('hidden');
            otherInput.value = '';
        }
    });
}

// ... (Functions: toggleEditExpenseOptions, toggleEditVehicleOptions, toggleEditVehicleDetails, generateDocumentFromDraft, getEditFormData, validateEditForm - Same logic) ...
// เพื่อความสะดวก ผมจะละส่วนที่ซ้ำซ้อน แต่คุณต้องมี function เหล่านี้ในไฟล์จริง
function toggleEditExpenseOptions() {
    const partialOptions = document.getElementById('edit-partial-expense-options');
    const totalContainer = document.getElementById('edit-total-expense-container');
    if (document.getElementById('edit-expense_partial')?.checked) {
        partialOptions.classList.remove('hidden');
        totalContainer.classList.remove('hidden');
    } else {
        partialOptions.classList.add('hidden');
        totalContainer.classList.add('hidden');
    }
}
function toggleEditVehicleOptions() {
    const privateDetails = document.getElementById('edit-private-vehicle-details');
    if (document.getElementById('edit-vehicle_private')?.checked) {
        privateDetails.classList.remove('hidden');
    } else {
        privateDetails.classList.add('hidden');
    }
}
function toggleEditVehicleDetails() {
    const privateDetails = document.getElementById('edit-private-vehicle-details'); 
    const publicDetails = document.getElementById('edit-public-vehicle-details'); 
    const privateCheckbox = document.querySelector('input[name="edit-vehicle_option"][value="private"]');
    const publicCheckbox = document.querySelector('input[name="edit-vehicle_option"][value="public"]');

    if (privateDetails) privateDetails.classList.toggle('hidden', !privateCheckbox?.checked);
    if (publicDetails) publicDetails.classList.toggle('hidden', !publicCheckbox?.checked);
}
// ... (Logic การ submit form และ generate document เหมือนเดิม)

// Public Data - ปรับปรุง HTML ให้รองรับ Mobile Table
async function loadPublicWeeklyData() {
    try {
        const [requestsResult, memosResult] = await Promise.all([apiCall('GET', 'getAllRequests'), apiCall('GET', 'getAllMemos')]);
        if (requestsResult.status === 'success') {
            const requests = requestsResult.data;
            const memos = memosResult.status === 'success' ? memosResult.data : [];
            const enrichedRequests = requests.map(req => {
                const relatedMemo = memos.find(m => m.refNumber === req.id);
                return { ...req, completedCommandUrl: relatedMemo ? relatedMemo.completedCommandUrl : null, realStatus: relatedMemo ? relatedMemo.status : req.status };
            });
            currentPublicWeeklyData = enrichedRequests;
            renderPublicTable(enrichedRequests);
        } else {
            document.getElementById('public-weekly-list').innerHTML = `<tr><td colspan="4" class="text-center py-4 text-red-500">ไม่สามารถโหลดข้อมูลได้</td></tr>`;
            document.getElementById('current-week-display').textContent = "Connection Error";
        }
    } catch (error) { document.getElementById('public-weekly-list').innerHTML = `<tr><td colspan="4" class="text-center py-4 text-gray-500">ไม่พบข้อมูล</td></tr>`; }
}

function renderPublicTable(requests) {
    const tbody = document.getElementById('public-weekly-list');
    // เพิ่ม class เพื่อบอกว่า table นี้เป็น responsive
    tbody.parentElement.classList.add('responsive-table'); // หาก parent เป็น table tag

    const now = new Date();
    const dayOfWeek = now.getDay();
    const daysToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
    const monday = new Date(now);
    monday.setDate(now.getDate() - daysToMonday); monday.setHours(0, 0, 0, 0);
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6); sunday.setHours(23, 59, 59, 999);
    
    document.getElementById('current-week-display').textContent = `${formatDisplayDate(monday)} - ${formatDisplayDate(sunday)}`;
    
    const weeklyRequests = requests.filter(req => {
        if (!req.startDate || !req.endDate) return false;
        const reqStart = new Date(req.startDate); const reqEnd = new Date(req.endDate);
        reqStart.setHours(0,0,0,0); reqEnd.setHours(0,0,0,0);
        return (reqStart <= sunday && reqEnd >= monday);
    }).sort((a, b) => new Date(a.startDate) - new Date(b.startDate));
    
    currentPublicWeeklyData = weeklyRequests;
    
    if (weeklyRequests.length === 0) { 
        tbody.innerHTML = `<tr><td colspan="4" class="text-center py-10 text-gray-500">ไม่มีรายการไปราชการในสัปดาห์นี้</td></tr>`; 
        return; 
    }
    
    tbody.innerHTML = weeklyRequests.map((req, index) => {
        // Sanitize
        const safeName = escapeHtml(req.requesterName);
        const safePosition = escapeHtml(req.requesterPosition || '');
        const safePurpose = escapeHtml(req.purpose);
        const safeLocation = escapeHtml(req.location);
        
        let attendeesText = "";
        const count = req.attendees ? (typeof req.attendees === 'string' ? JSON.parse(req.attendees).length : req.attendees.length) : (req.attendeeCount || 0);
        if (count > 0) { attendeesText = `<div class="text-xs text-indigo-500 mt-1 cursor-pointer hover:underline" onclick="openPublicAttendeeModal(${index})">👥 และคณะรวม ${count + 1} คน</div>`; }
        
        const dateText = `${formatDisplayDate(req.startDate)} - ${formatDisplayDate(req.endDate)}`;
        
        // Status Badge Logic
        let actionHtml = '';
        if (req.completedCommandUrl && req.completedCommandUrl.trim() !== "") {
             actionHtml = `<a href="${req.completedCommandUrl}" target="_blank" class="btn bg-green-600 hover:bg-green-700 text-white btn-sm">ดูคำสั่ง</a>`;
        } else {
             const displayStatus = req.realStatus || req.status;
             actionHtml = `<span class="badge badge-gray text-xs">${translateStatus(displayStatus)}</span>`;
        }

        // เพิ่ม data-label สำหรับ Mobile View
        return `
        <tr class="border-b hover:bg-gray-50 transition">
            <td class="px-6 py-4 whitespace-nowrap font-medium text-indigo-600" data-label="วัน-เวลา">${dateText}</td>
            <td class="px-6 py-4" data-label="ชื่อผู้ขอ">
                <div class="font-bold text-gray-800">${safeName}</div>
                <div class="text-xs text-gray-500">${safePosition}</div>
            </td>
            <td class="px-6 py-4" data-label="เรื่อง / สถานที่">
                <div class="font-medium text-gray-900 truncate max-w-xs" title="${safePurpose}">${safePurpose}</div>
                <div class="text-xs text-gray-500">ณ ${safeLocation}</div>
                ${attendeesText}
            </td>
            <td class="px-6 py-4 text-center align-middle" data-label="ไฟล์คำสั่ง">${actionHtml}</td>
        </tr>`;
    }).join('');
}
