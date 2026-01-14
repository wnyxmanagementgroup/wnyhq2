// --- STATS FUNCTIONS WITH CACHING ---

let lastStatsLoadTime = 0;
const STATS_CACHE_DURATION = 5 * 60 * 1000; // 5 นาที

async function loadStatsData(forceRefresh = false) {
    try {
        console.log("🔄 Loading stats data...");
        const user = getCurrentUser();
        if (!user) return;

        // Check Cache
        const now = Date.now();
        if (!forceRefresh && (now - lastStatsLoadTime < STATS_CACHE_DURATION) && allRequestsCache.length > 0) {
             console.log("⚡ Using cached stats data");
             const userRequests = user.role === 'admin' ? allRequestsCache : allRequestsCache.filter(req => req.username === user.username);
             const userMemos = user.role === 'admin' ? allMemosCache : userMemosCache; 
             renderStatsOverview(userRequests, userMemos, allUsersCache, user);
             return;
        }

        // Show Loading UI
        document.getElementById('stats-overview').innerHTML = `
            <div class="text-center p-8">
                <div class="loader mx-auto"></div>
                <p class="mt-4">กำลังโหลดสถิติ...</p>
            </div>`;
        document.getElementById('stats-charts').classList.add('hidden');

        // Fetch Data
        const [requestsResult, memosResult, usersResult] = await Promise.all([
            apiCall('GET', 'getAllRequests').catch(() => ({ status: 'success', data: [] })),
            apiCall('GET', 'getAllMemos').catch(() => ({ status: 'success', data: [] })),
            apiCall('GET', 'getAllUsers').catch(() => ({ status: 'success', data: [] }))
        ]);

        // Update Cache
        if(requestsResult.status === 'success') allRequestsCache = requestsResult.data;
        if(memosResult.status === 'success') allMemosCache = memosResult.data;
        if(usersResult.status === 'success') allUsersCache = usersResult.data;
        
        lastStatsLoadTime = Date.now();

        // Process Data for View
        const requests = requestsResult?.data || [];
        const memos = memosResult?.data || [];
        const users = usersResult?.data || [];

        const userRequests = user.role === 'admin' ? requests : requests.filter(req => req.username === user.username);
        const userMemos = user.role === 'admin' ? memos : memos.filter(memo => memo.submittedBy === user.username);

        renderStatsOverview(userRequests, userMemos, users, user);

    } catch (error) {
        console.error('❌ Error loading stats:', error);
        document.getElementById('stats-overview').innerHTML = `
            <div class="text-center p-8 text-red-500">
                <p>เกิดข้อผิดพลาดในการโหลดข้อมูล</p>
                <button onclick="loadStatsData(true)" class="btn btn-primary mt-4">ลองอีกครั้ง</button>
            </div>`;
    }
}

function renderStatsOverview(requests, memos, users, currentUser) {
    const stats = calculateStats(requests, memos, users, currentUser);
    const container = document.getElementById('stats-overview');
    
    container.innerHTML = `
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-6">
            <div class="stat-card bg-white rounded-lg shadow p-4 border-l-4 border-blue-500">
                <div class="flex items-center">
                    <div class="bg-blue-100 p-3 rounded-lg">📋</div>
                    <div class="ml-4">
                        <p class="text-sm font-medium text-gray-600">คำขอทั้งหมด</p>
                        <p class="text-2xl font-bold text-gray-900">${stats.totalRequests}</p>
                    </div>
                </div>
            </div>
            <div class="stat-card bg-white rounded-lg shadow p-4 border-l-4 border-green-500">
                <div class="flex items-center">
                    <div class="bg-green-100 p-3 rounded-lg">✅</div>
                    <div class="ml-4">
                        <p class="text-sm font-medium text-gray-600">คำขอที่เสร็จสิ้น</p>
                        <p class="text-2xl font-bold text-gray-900">${stats.completedRequests}</p>
                    </div>
                </div>
            </div>
            <div class="stat-card bg-white rounded-lg shadow p-4 border-l-4 border-purple-500">
                <div class="flex items-center">
                    <div class="bg-purple-100 p-3 rounded-lg">📤</div>
                    <div class="ml-4">
                        <p class="text-sm font-medium text-gray-600">บันทึกข้อความ</p>
                        <p class="text-2xl font-bold text-gray-900">${stats.totalMemos}</p>
                    </div>
                </div>
            </div>
            <div class="stat-card bg-white rounded-lg shadow p-4 border-l-4 border-yellow-500">
                <div class="flex items-center">
                    <div class="bg-yellow-100 p-3 rounded-lg">👥</div>
                    <div class="ml-4">
                        <p class="text-sm font-medium text-gray-600">ผู้ใช้ทั้งหมด</p>
                        <p class="text-2xl font-bold text-gray-900">${stats.totalUsers}</p>
                    </div>
                </div>
            </div>
        </div>
        <div id="stats-charts" class="mt-8">
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div class="chart-container">
                    <h3 class="text-lg font-bold mb-4 text-gray-800">คำขอรายเดือน (6 เดือนล่าสุด)</h3>
                    <canvas id="requests-chart"></canvas>
                </div>
                <div class="chart-container">
                    <h3 class="text-lg font-bold mb-4 text-gray-800">สรุปสถานะคำขอ</h3>
                    <canvas id="status-chart"></canvas>
                </div>
            </div>
        </div>`;
        
    // Destroy old charts if exist
    if (window.requestsChartInstance) { window.requestsChartInstance.destroy(); window.requestsChartInstance = null; }
    if (window.statusChartInstance) { window.statusChartInstance.destroy(); window.statusChartInstance = null; }
    
    // Create new charts with delay to ensure DOM is ready
    setTimeout(() => { createCharts(stats); }, 100);
}

function createCharts(stats) {
    const monthlyCtx = document.getElementById('requests-chart');
    if (monthlyCtx) {
        window.requestsChartInstance = new Chart(monthlyCtx, {
            type: 'bar', 
            data: { 
                labels: stats.monthlyStats.map(m => m.month), 
                datasets: [{ 
                    label: 'จำนวนคำขอ', 
                    data: stats.monthlyStats.map(m => m.count), 
                    backgroundColor: 'rgba(79, 70, 229, 0.6)', 
                    borderColor: 'rgba(79, 70, 229, 1)', 
                    borderWidth: 1, 
                    borderRadius: 6 
                }] 
            },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { display: false } }, 
                scales: { y: { beginAtZero: true, ticks: { precision: 0 } }, x: { grid: { display: false } } } 
            }
        });
    }
    const statusCtx = document.getElementById('status-chart');
    if (statusCtx) {
        const statusEntries = Object.entries(stats.requestStatus);
        window.statusChartInstance = new Chart(statusCtx, {
            type: 'doughnut', 
            data: { 
                labels: statusEntries.map(([status, count]) => `${translateStatus(status)} (${count})`), 
                datasets: [{ 
                    data: statusEntries.map(([status, count]) => count), 
                    backgroundColor: ['#16a34a', '#3b82f6', '#f59e0b', '#ef4444', '#a855f7', '#f97316'], 
                    borderColor: 'rgba(255, 255, 255, 0.8)', 
                    borderWidth: 2 
                }] 
            },
            options: { 
                responsive: true, 
                maintainAspectRatio: false, 
                plugins: { legend: { position: 'bottom' } }, 
                cutout: '50%' 
            }
        });
    }
    document.getElementById('stats-charts')?.classList.remove('hidden');
}

function calculateStats(requests, memos, users, currentUser) {
    // ... (Logic เดิมของการคำนวณ stats) ...
    const requestStatus = {};
    requests.forEach(req => { const status = req.status || 'กำลังดำเนินการ'; requestStatus[status] = (requestStatus[status] || 0) + 1; });
    const completedRequests = requests.filter(req => req.status === 'เสร็จสิ้น/รับไฟล์ไปใช้งาน' || req.status === 'Approved' || req.commandStatus === 'เสร็จสิ้นรอออกคำสั่งไปราชการ').length;
    const departmentStats = {};
    requests.forEach(req => { const dept = req.department || 'ไม่ระบุแผนก'; departmentStats[dept] = (departmentStats[dept] || 0) + 1; });
    const userStats = { total: users.length, admins: users.filter(u => u.role === 'admin').length, regularUsers: users.filter(u => u.role === 'user').length };
    
    // Monthly stats calculation (6 months back)
    const monthlyStats = []; const now = new Date();
    for (let i = 5; i >= 0; i--) {
        const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
        const monthKey = date.toLocaleDateString('th-TH', { year: 'numeric', month: 'long' });
        const monthStart = new Date(date.getFullYear(), date.getMonth(), 1);
        const monthEnd = new Date(date.getFullYear(), date.getMonth() + 1, 0);
        const monthRequests = requests.filter(req => {
            const dateString = req.timestamp || req.startDate || req.docDate || req.createdAt;
            if (!dateString) return false;
            try { const reqDate = new Date(dateString); return reqDate >= monthStart && reqDate <= monthEnd; } catch (e) { return false; }
        });
        monthlyStats.push({ month: monthKey, count: monthRequests.length, completed: monthRequests.filter(req => req.status === 'เสร็จสิ้น').length });
    }
    return { totalRequests: requests.length, completedRequests, totalMemos: memos.length, totalUsers: users.length, requestStatus, departmentStats, userStats, monthlyStats };
}

async function exportStatsReport() {
    // ... (Logic เดิม) ...
    try {
        const user = getCurrentUser(); if (!user) return;
        toggleLoader('export-stats', true);
        const [requestsResult, memosResult, usersResult] = await Promise.all([apiCall('GET', 'getAllRequests'), apiCall('GET', 'getAllMemos'), apiCall('GET', 'getAllUsers')]);
        const requests = requestsResult.data || []; const memos = memosResult.data || []; const users = usersResult.data || [];
        const userRequests = user.role === 'admin' ? requests : requests.filter(req => req.username === user.username);
        const stats = calculateStats(userRequests, memos, users, user);
        
        const reportData = [['รายงานสถิติ', '', '', ''], ['วันที่', new Date().toLocaleDateString('th-TH'), '', ''], ['', '', '', ''], ['สถิติภาพรวม', '', '', ''], ['คำขอทั้งหมด', stats.totalRequests, '', ''], ['เสร็จสิ้น', stats.completedRequests, '', '']];
        const ws = XLSX.utils.aoa_to_sheet(reportData); const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Report');
        XLSX.writeFile(wb, `Report_${new Date().toISOString().split('T')[0]}.xlsx`);
        showAlert('สำเร็จ', 'ส่งออกรายงานเรียบร้อยแล้ว');
    } catch (error) { showAlert('ผิดพลาด', 'ไม่สามารถส่งออกรายงานได้'); } finally { toggleLoader('export-stats', false); }
}
