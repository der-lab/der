 // 登录相关变量
        const validUsers = [
            { username: 'lxh', password: 'lxh123' }
        ];
        
        // 全局变量
        let originalData = [];
        let processedData = [];
        let filteredData = [];
        let workbook = null;
        let fileName = "";
        let colorMap = new Map();
        let lineColorMap = new Map();
        let colorIndex = 0;
        let isDataVisible = false;
        let lineInfo = {};
        let selectedRows = new Set();
        let isSearchActive = false;
        let vehiclePickupStats = {};
        let lineOrder = [];
        let lineUnassignedCount = {};
        let vehicleCapacityMap = new Map();
        let lastLineInfo = {};
        let plateStats = {};
        let manuallyAssignedOrders = new Set();
        let lineManuallyAssignedCount = {};
        let plateLineMap = new Map();
        let manuallyModifiedOrders = new Map(); // 订单ID -> 修改的字段和值
        let manuallyAddedOrders = new Set(); // 手动添加的订单
        let manuallyDeletedOrders = new Set(); // 手动删除的订单
        
        // 登录验证函数
        function validateLogin(username, password) {
            return validUsers.find(user => user.username === username && user.password === password);
        }
        
        // 登录处理函数
        function handleLogin() {
            const username = document.getElementById('username').value.trim();
            const password = document.getElementById('password').value;
            const loginError = document.getElementById('loginError');
            
            if (!username || !password) {
                loginError.textContent = '请输入用户名和密码';
                return;
            }
            
            const user = validateLogin(username, password);
            if (user) {
                // 登录成功，存储用户信息使用sessionStorage，关闭浏览器后会自动清除
                const userToStore = { username: user.username };
                sessionStorage.setItem('loggedInUser', JSON.stringify(userToStore));
                loginError.textContent = '';
                
                // 隐藏登录界面，显示主系统
                document.getElementById('loginContainer').style.display = 'none';
                document.getElementById('mainContainer').style.display = 'block';
            } else {
                loginError.textContent = '用户名或密码错误';
            }
        }
        
        // 检查用户是否已登录
        function checkLoginStatus() {
            const loggedInUser = sessionStorage.getItem('loggedInUser');
            if (loggedInUser) {
                // 用户已登录，显示主系统
                document.getElementById('loginContainer').style.display = 'none';
                document.getElementById('mainContainer').style.display = 'block';
            }
        }
        
        // 新增：线路统计面板显示状态
        let lineStatsVisible = true;
        
        // 新增：调整大小功能变量
        let isResizing = false;
        let startY, startHeight;
        
        // 新增：车辆信息检查结果
        let vehicleInfoErrors = [];
        
        // 新增：车辆总人数共享跟踪
        let vehicleTotalAssigned = new Map(); // 车牌 -> 已分配总人数
        
        // 新增：线路时间地点统计数据
        let lineTimeLocationData = [];
        
        // 新增：导出排序方式
        let exportSortType = 'line'; // 默认按线路排序
        
        // 新增：数据库相关
        let db = null;
        const DB_NAME = 'TableAnalysisSystemDB';
        const DB_VERSION = 2;
        const STORES = {
            ORIGINAL_DATA: 'originalData',
            PROCESSED_DATA: 'processedData',
            LINE_INFO: 'lineInfo',
            LAST_LINE_INFO: 'lastLineInfo',
            MANUALLY_MODIFIED: 'manuallyModified',
            MANUALLY_ADDED: 'manuallyAdded',
            MANUALLY_DELETED: 'manuallyDeleted'
        };
        
        // 预定义的颜色集合
        const colorPalette = [
            "#E6F3FF", "#E6FFE6", "#FFF0E6", "#FFF6E6", "#F0E6FF",
            "#FFE6F2", "#E6FFFF", "#FFF0F5", "#F5F5DC", "#F0FFF0",
            "#FFE6E6", "#E6E6FF", "#FFFFE6", "#E6F5FF", "#F0F8FF",
            "#FAFAD2", "#F5F5F5", "#FFFACD", "#E0FFFF", "#F0FFF0",
            "#F5F5DC", "#FFF8DC", "#FFFAF0", "#F8F8FF", "#F5F5F5",
            "#FFF5EE", "#F5FFFA", "#FFFAFA", "#F0F8FF", "#F8F8FF"
        ];
        
        // 数据库初始化函数
        function initDatabase() {
            return new Promise((resolve, reject) => {
                const request = indexedDB.open(DB_NAME, DB_VERSION);
                
                request.onerror = function(event) {
                    console.error('数据库打开失败:', event.target.error);
                    reject(event.target.error);
                };
                
                request.onsuccess = function(event) {
                    db = event.target.result;
                    console.log('数据库连接成功');
                    resolve(db);
                };
                
                request.onupgradeneeded = function(event) {
                    db = event.target.result;
                    console.log('数据库版本升级，当前版本:', event.oldVersion, '目标版本:', event.newVersion);
                    
                    // 创建存储对象
                    if (!db.objectStoreNames.contains(STORES.ORIGINAL_DATA)) {
                        db.createObjectStore(STORES.ORIGINAL_DATA);
                        console.log('创建存储对象:', STORES.ORIGINAL_DATA);
                    }
                    
                    if (!db.objectStoreNames.contains(STORES.PROCESSED_DATA)) {
                        db.createObjectStore(STORES.PROCESSED_DATA);
                        console.log('创建存储对象:', STORES.PROCESSED_DATA);
                    }
                    
                    if (!db.objectStoreNames.contains(STORES.LINE_INFO)) {
                        db.createObjectStore(STORES.LINE_INFO);
                        console.log('创建存储对象:', STORES.LINE_INFO);
                    }
                    
                    if (!db.objectStoreNames.contains(STORES.LAST_LINE_INFO)) {
                        db.createObjectStore(STORES.LAST_LINE_INFO);
                        console.log('创建存储对象:', STORES.LAST_LINE_INFO);
                    }
                    
                    if (!db.objectStoreNames.contains(STORES.MANUALLY_MODIFIED)) {
                        db.createObjectStore(STORES.MANUALLY_MODIFIED);
                        console.log('创建存储对象:', STORES.MANUALLY_MODIFIED);
                    }
                    
                    if (!db.objectStoreNames.contains(STORES.MANUALLY_ADDED)) {
                        db.createObjectStore(STORES.MANUALLY_ADDED);
                        console.log('创建存储对象:', STORES.MANUALLY_ADDED);
                    }
                    
                    if (!db.objectStoreNames.contains(STORES.MANUALLY_DELETED)) {
                        db.createObjectStore(STORES.MANUALLY_DELETED);
                        console.log('创建存储对象:', STORES.MANUALLY_DELETED);
                    }
                    
                    // 添加vehicleCapacityMap存储
                    if (!db.objectStoreNames.contains('vehicleCapacityMap')) {
                        db.createObjectStore('vehicleCapacityMap');
                        console.log('创建存储对象: vehicleCapacityMap');
                    }
                };
            });
        }
        
        // 保存数据到数据库
        function saveDataToDB(storeName, key, value) {
            return new Promise((resolve, reject) => {
                if (!db) {
                    reject(new Error('数据库未初始化'));
                    return;
                }
                
                const transaction = db.transaction([storeName], 'readwrite');
                const store = transaction.objectStore(storeName);
                const request = store.put(value, key);
                
                request.onsuccess = function() {
                    console.log(`数据已保存到 ${storeName} 存储`);
                    resolve();
                };
                
                request.onerror = function(event) {
                    console.error(`保存数据到 ${storeName} 失败:`, event.target.error);
                    reject(event.target.error);
                };
            });
        }
        
        // 从数据库加载数据
        function loadDataFromDB(storeName, key) {
            return new Promise((resolve, reject) => {
                if (!db) {
                    reject(new Error('数据库未初始化'));
                    return;
                }
                
                const transaction = db.transaction([storeName], 'readonly');
                const store = transaction.objectStore(storeName);
                const request = store.get(key);
                
                request.onsuccess = function() {
                    resolve(request.result);
                };
                
                request.onerror = function(event) {
                    console.error(`从 ${storeName} 加载数据失败:`, event.target.error);
                    reject(event.target.error);
                };
            });
        }
        
        // 清空数据库
        function clearDB() {
            return new Promise((resolve, reject) => {
                if (!db) {
                    reject(new Error('数据库未初始化'));
                    return;
                }
                
                const storeNames = Object.values(STORES);
                const transaction = db.transaction(storeNames, 'readwrite');
                
                let completed = 0;
                let error = null;
                
                storeNames.forEach(storeName => {
                    const store = transaction.objectStore(storeName);
                    const request = store.clear();
                    
                    request.onsuccess = function() {
                        completed++;
                        if (completed === storeNames.length) {
                            if (error) {
                                reject(error);
                            } else {
                                resolve();
                            }
                        }
                    };
                    
                    request.onerror = function(event) {
                        error = event.target.error;
                        completed++;
                        if (completed === storeNames.length) {
                            reject(error);
                        }
                    };
                });
            });
        }
        
        // 保存所有数据到数据库
        async function saveAllData() {
            try {
                await saveDataToDB(STORES.ORIGINAL_DATA, 'data', originalData);
                await saveDataToDB(STORES.PROCESSED_DATA, 'data', processedData);
                await saveDataToDB(STORES.LINE_INFO, 'data', lineInfo);
                await saveDataToDB(STORES.LAST_LINE_INFO, 'data', lastLineInfo);
                await saveDataToDB(STORES.MANUALLY_MODIFIED, 'data', Object.fromEntries(manuallyModifiedOrders));
                await saveDataToDB(STORES.MANUALLY_ADDED, 'data', Array.from(manuallyAddedOrders));
                await saveDataToDB(STORES.MANUALLY_DELETED, 'data', Array.from(manuallyDeletedOrders));
                try {
                    await saveDataToDB('vehicleCapacityMap', 'data', Object.fromEntries(vehicleCapacityMap));
                } catch (error) {
                    console.error('保存vehicleCapacityMap失败:', error);
                    // 继续执行其他保存操作
                }
                
                showNotification('数据已保存到数据库', 'success');
            } catch (error) {
                console.error('保存数据失败:', error);
                showNotification('保存数据失败', 'error');
            }
        }
        
        // 从数据库加载所有数据
        async function loadAllData() {
            try {
                // 逐个加载数据，确保即使某些数据加载失败，其他数据仍然能够被加载
                let original = null;
                let processed = null;
                let line = null;
                let lastLine = null;
                let modified = null;
                let added = null;
                let deleted = null;
                let vehicleCapacity = null;
                
                try {
                    original = await loadDataFromDB(STORES.ORIGINAL_DATA, 'data');
                } catch (error) {
                    console.error('加载originalData失败:', error);
                }
                
                try {
                    processed = await loadDataFromDB(STORES.PROCESSED_DATA, 'data');
                } catch (error) {
                    console.error('加载processedData失败:', error);
                }
                
                try {
                    line = await loadDataFromDB(STORES.LINE_INFO, 'data');
                } catch (error) {
                    console.error('加载lineInfo失败:', error);
                }
                
                try {
                    lastLine = await loadDataFromDB(STORES.LAST_LINE_INFO, 'data');
                } catch (error) {
                    console.error('加载lastLineInfo失败:', error);
                }
                
                try {
                    modified = await loadDataFromDB(STORES.MANUALLY_MODIFIED, 'data');
                } catch (error) {
                    console.error('加载manuallyModifiedOrders失败:', error);
                }
                
                try {
                    added = await loadDataFromDB(STORES.MANUALLY_ADDED, 'data');
                } catch (error) {
                    console.error('加载manuallyAddedOrders失败:', error);
                }
                
                try {
                    deleted = await loadDataFromDB(STORES.MANUALLY_DELETED, 'data');
                } catch (error) {
                    console.error('加载manuallyDeletedOrders失败:', error);
                }
                
                try {
                    vehicleCapacity = await loadDataFromDB('vehicleCapacityMap', 'data');
                } catch (error) {
                    console.error('加载vehicleCapacityMap失败:', error);
                }
                
                // 加载所有数据，包括originalData
                // 这样当用户再次点击确定并分析数据时，系统会从数据库中保存的最后更新的数据分析
                if (original) originalData = original;
                if (processed) processedData = processed;
                if (line) lineInfo = line;
                if (lastLine) lastLineInfo = lastLine;
                if (modified) {
                    try {
                        manuallyModifiedOrders = new Map(Object.entries(modified));
                    } catch (error) {
                        console.error('解析manuallyModifiedOrders失败:', error);
                        manuallyModifiedOrders = new Map();
                    }
                }
                if (added) {
                    try {
                        manuallyAddedOrders = new Set(added);
                    } catch (error) {
                        console.error('解析manuallyAddedOrders失败:', error);
                        manuallyAddedOrders = new Set();
                    }
                }
                if (deleted) {
                    try {
                        manuallyDeletedOrders = new Set(deleted);
                    } catch (error) {
                        console.error('解析manuallyDeletedOrders失败:', error);
                        manuallyDeletedOrders = new Set();
                    }
                }
                if (vehicleCapacity) {
                    try {
                        vehicleCapacityMap = new Map(Object.entries(vehicleCapacity));
                    } catch (error) {
                        console.error('解析vehicleCapacityMap失败:', error);
                        vehicleCapacityMap = new Map();
                    }
                }
                
                // 重新计算相关数据结构
                updatePlateLineMap();
                updateVehicleTotalAssigned();
                
                // 更新界面状态
                if (originalData.length > 0) {
                    analyzeBtn.disabled = false;
                }
                
                if (processedData.length > 0) {
                    toggleBtn.disabled = false;
                    exportBtn.disabled = false;
                    ticketStatsSection.style.display = 'block';
                    displayLineTicketStats();
                    displayLineTimeLocationStats();
                    
                    // 直接显示整合后的数据和车次统计
                    dataSection.style.display = 'flex';
                    vehiclePickupStatsPanel.style.display = 'block';
                    displayProcessedData();
                    displayVehiclePickupStatsPanel();
                    ticketStatsSection.classList.add('ticket-stats-down');
                    
                    // 隐藏切换按钮，但保留其功能
                    toggleBtn.style.display = 'none';
                }
                
                showNotification('数据已成功从数据库加载', 'success');
            } catch (error) {
                console.error('加载数据失败:', error);
                showNotification('加载数据失败', 'error');
            }
        }
        
        // DOM元素
        const fileInput = document.getElementById('fileInput');
        const uploadBtn = document.getElementById('uploadBtn');
        const uploadArea = document.getElementById('uploadArea');
        const fileNameDisplay = document.getElementById('fileName');
        const analyzeBtn = document.getElementById('analyzeBtn');
        const toggleBtn = document.getElementById('toggleBtn');
        const exportBtn = document.getElementById('exportBtn');
        const processedDataBody = document.getElementById('processedData');
        const ticketStatsSection = document.getElementById('ticketStatsSection');
        const lineStatsPanel = document.getElementById('lineStatsPanel');
        const loading = document.getElementById('loading');
        const dataSection = document.getElementById('dataSection');
        const refreshBtn = document.getElementById('refreshBtn');
        const vehiclePickupStatsPanel = document.getElementById('vehiclePickupStatsPanel');
        const selectAllPlateBtn = document.getElementById('selectAllPlateBtn');
        const selectAllUnassignedBtn = document.getElementById('selectAllUnassignedBtn');
        const loginBtn = document.getElementById('loginBtn');
        const usernameInput = document.getElementById('username');
        const passwordInput = document.getElementById('password');
        
        // 页面加载时初始化数据库
        document.addEventListener('DOMContentLoaded', async function() {
            // 检查登录状态
            checkLoginStatus();
            
            // 登录按钮事件监听器
            if (loginBtn) {
                loginBtn.addEventListener('click', handleLogin);
            }
            
            // 回车键登录
            if (usernameInput && passwordInput) {
                [usernameInput, passwordInput].forEach(input => {
                    input.addEventListener('keypress', function(e) {
                        if (e.key === 'Enter') {
                            handleLogin();
                        }
                    });
                });
            }
            
            try {
                await initDatabase();
                await loadAllData();
            } catch (error) {
                console.error('初始化数据库失败:', error);
            }
            
            // 导航栏功能
            const navbarLinks = document.querySelectorAll('.navbar a');
            const sections = document.querySelectorAll('.section');
            
            // 默认显示第一个板块
            sections[0].classList.add('active');
            navbarLinks[0].classList.add('active');
            
            navbarLinks.forEach(link => {
                link.addEventListener('click', function(e) {
                    e.preventDefault();
                    
                    // 移除所有板块的active类
                    sections.forEach(section => {
                        section.classList.remove('active');
                    });
                    
                    // 移除所有导航链接的active类
                    navbarLinks.forEach(navLink => {
                        navLink.classList.remove('active');
                    });
                    
                    // 添加当前板块的active类
                    const sectionId = this.getAttribute('data-section');
                    const targetSection = document.getElementById(sectionId);
                    if (targetSection) {
                        targetSection.classList.add('active');
                    }
                    
                    // 添加当前导航链接的active类
                    this.classList.add('active');
                    
                    // 当切换到订单数据板块时，设置isDataVisible为true
                    if (sectionId === 'order-data-section') {
                        isDataVisible = true;
                        // 重新显示数据，确保筛选功能正常工作
                        displayProcessedData();
                        displayVehiclePickupStatsPanel();
                    }
                });
            });
        });
        
        // 新增：线路时间地点统计相关元素
        const lineTimeLocationPanel = document.getElementById('lineTimeLocationPanel');
        const lineTimeLocationBody = document.getElementById('lineTimeLocationBody');
        const lineTimeLocationSearch = document.getElementById('lineTimeLocationSearch');
        const lineTimeLocationTableContainer = document.getElementById('lineTimeLocationTableContainer');
        const resizeHandle = document.getElementById('resizeHandle');
        
        // 搜索相关元素
        const searchBtn = document.getElementById('searchBtn');
        const clearSearchBtn = document.getElementById('clearSearchBtn');
        const modifyBtn = document.getElementById('modifyBtn');
        const addPassengerBtn = document.getElementById('addPassengerBtn');
        const searchResultsInfo = document.getElementById('searchResultsInfo');
        const selectAllCheckbox = document.getElementById('selectAllCheckbox');
        
        // 搜索输入框
        const searchContact = document.getElementById('searchContact');
        const searchPhone = document.getElementById('searchPhone');
        const searchPlate = document.getElementById('searchPlate');
        const searchLine = document.getElementById('searchLine');
        const searchPickup = document.getElementById('searchPickup');
        const searchDropoff = document.getElementById('searchDropoff');
        const modifyTimeSort = document.getElementById('modifyTimeSort');
        
        // 新增：搜索框刷新按钮元素
        const refreshSearchBtn = document.getElementById('refreshSearchBtn');
        // 新增：数据库操作按钮
        const saveDataBtn = document.getElementById('saveDataBtn');
        const loadDataBtn = document.getElementById('loadDataBtn');

        // 在事件监听器部分添加
        refreshSearchBtn.addEventListener('click', refreshSearchData);
        saveDataBtn.addEventListener('click', saveAllData);
        loadDataBtn.addEventListener('click', loadAllData);
        
        // 新增：修改时间排序事件监听器
        if (modifyTimeSort) {
            modifyTimeSort.addEventListener('change', function() {
                if (isDataVisible) {
                    displayProcessedData();
                    displayVehiclePickupStatsPanel();
                }
            });
        }

        // 新增：刷新搜索数据函数
        function refreshSearchData() {
            // 清空所有搜索输入框
            searchContact.value = '';
            searchPhone.value = '';
            searchPlate.value = '';
            searchLine.value = '';
            searchPickup.value = '';
            searchDropoff.value = '';
            
            // 重置搜索状态
            isSearchActive = false;
            filteredData = [];
            
            // 更新搜索结果信息
            searchResultsInfo.textContent = `显示全部 ${processedData.length} 条数据`;
            searchResultsInfo.style.color = '#4a6491';
            
            // 清空选中行
            selectedRows.clear();
            selectAllCheckbox.checked = false;
            updateModifyButtonState();
            selectAllPlateBtn.style.display = 'none';
            selectAllUnassignedBtn.style.display = 'none';
            
            // 重新显示数据
            if (isDataVisible) {
                displayProcessedData();
                displayVehiclePickupStatsPanel();
            }
            
            showNotification('搜索数据已刷新，显示全部数据', 'success');
        }


        // 线路信息填写弹窗相关元素
        const modalOverlay = document.getElementById('modalOverlay');
        const closeModalBtn = document.getElementById('closeModalBtn');
        const cancelBtn = document.getElementById('cancelBtn');
        const confirmBtn = document.getElementById('confirmBtn');
        const lineInputsBody = document.getElementById('lineInputsBody');
        
        // 弹窗搜索输入框
        const modalSearchLine = document.getElementById('modalSearchLine');
        const modalSearchPlate = document.getElementById('modalSearchPlate');
        
        // 修改数据弹窗相关元素
        const modifyModalOverlay = document.getElementById('modifyModalOverlay');
        const closeModifyModalBtn = document.getElementById('closeModifyModalBtn');
        const cancelModifyBtn = document.getElementById('cancelModifyBtn');
        const confirmModifyBtn = document.getElementById('confirmModifyBtn');
        const modifyForm = document.getElementById('modifyForm');
        const deletePassengerBtn = document.getElementById('deletePassengerBtn');
        
        // 新增：添加乘客弹窗相关元素
        const addPassengerModalOverlay = document.getElementById('addPassengerModalOverlay');
        const closeAddPassengerModalBtn = document.getElementById('closeAddPassengerModalBtn');
        const cancelAddPassengerBtn = document.getElementById('cancelAddPassengerBtn');
        const confirmAddPassengerBtn = document.getElementById('confirmAddPassengerBtn');
        const addPassengerForm = document.getElementById('addPassengerForm');
        const identifyLineBtn = document.getElementById('identifyLineBtn');
        
        // 新增：导出排序对话框元素
        const exportSortDialogOverlay = document.getElementById('exportSortDialogOverlay');
        const exportSortDialogCancelBtn = document.getElementById('exportSortDialogCancelBtn');
        const exportSortDialogOkBtn = document.getElementById('exportSortDialogOkBtn');
        const sortByLineOption = document.getElementById('sortByLineOption');
        const sortByCampusOption = document.getElementById('sortByCampusOption');
        
        // 新增：删除确认对话框元素
        const deleteConfirmDialogOverlay = document.getElementById('deleteConfirmDialogOverlay');
        const deleteConfirmDialogCancelBtn = document.getElementById('deleteConfirmDialogCancelBtn');
        const deleteConfirmDialogOkBtn = document.getElementById('deleteConfirmDialogOkBtn');
        const deleteConfirmDialogBody = document.getElementById('deleteConfirmDialogBody');
        
        // 新增：确认对话框元素
        const confirmDialogOverlay = document.getElementById('confirmDialogOverlay');
        const confirmDialogCancelBtn = document.getElementById('confirmDialogCancelBtn');
        const confirmDialogOkBtn = document.getElementById('confirmDialogOkBtn');
        
        // 新增：车牌错误对话框元素
        const vehicleErrorDialogOverlay = document.getElementById('vehicleErrorDialogOverlay');
        const vehicleErrorList = document.getElementById('vehicleErrorList');
        const vehicleErrorDialogOkBtn = document.getElementById('vehicleErrorDialogOkBtn');
        
        // 事件监听器
        uploadBtn.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('click', () => fileInput.click());
        refreshBtn.addEventListener('click', showRefreshConfirm);
        
        // 拖放事件监听器
        uploadArea.addEventListener('dragover', handleDragOver);
        uploadArea.addEventListener('dragleave', handleDragLeave);
        uploadArea.addEventListener('drop', handleDrop);
        
        fileInput.addEventListener('change', handleFileUpload);
        analyzeBtn.addEventListener('click', showLineInfoModal);
        toggleBtn.addEventListener('click', toggleDataView);
        exportBtn.addEventListener('click', showExportSortDialog);
        
        // 搜索相关事件
        searchBtn.addEventListener('click', performSearch);
        clearSearchBtn.addEventListener('click', clearSearch);
        
        // 智能识别按钮事件
        const smartIdentifyBtn = document.getElementById('smartIdentifyBtn');
        smartIdentifyBtn.addEventListener('click', smartIdentify);
        modifyBtn.addEventListener('click', showModifyModal);
        addPassengerBtn.addEventListener('click', showAddPassengerModal);
        selectAllCheckbox.addEventListener('change', toggleSelectAll);
        selectAllPlateBtn.addEventListener('click', selectAllCurrentPlate);
        selectAllUnassignedBtn.addEventListener('click', selectAllCurrentLineUnassigned);
        
        // 为搜索输入框添加回车键搜索
        searchContact.addEventListener('keyup', (e) => { if (e.key === 'Enter') performSearch(); });
        searchPhone.addEventListener('keyup', (e) => { if (e.key === 'Enter') performSearch(); });
        searchPlate.addEventListener('keyup', (e) => { if (e.key === 'Enter') performSearch(); });
        searchLine.addEventListener('keyup', (e) => { if (e.key === 'Enter') performSearch(); });
        searchPickup.addEventListener('keyup', (e) => { if (e.key === 'Enter') performSearch(); });
        searchDropoff.addEventListener('keyup', (e) => { if (e.key === 'Enter') performSearch(); });
        
        // 新增：线路时间地点统计搜索事件
        lineTimeLocationSearch.addEventListener('input', filterLineTimeLocationData);
        
        // 新增：调整大小事件
        if (resizeHandle) {
            resizeHandle.addEventListener('mousedown', startResize);
        }
        
        // 弹窗搜索输入框事件
        modalSearchLine.addEventListener('input', filterModalRows);
        modalSearchPlate.addEventListener('input', filterModalRows);
        
        // 线路信息填写弹窗事件监听器
        closeModalBtn.addEventListener('click', closeModal);
        cancelBtn.addEventListener('click', closeModal);
        confirmBtn.addEventListener('click', processDataWithLineInfo);
        
        // 修改弹窗事件监听器
        closeModifyModalBtn.addEventListener('click', closeModifyModal);
        cancelModifyBtn.addEventListener('click', closeModifyModal);
        confirmModifyBtn.addEventListener('click', confirmModify);
        deletePassengerBtn.addEventListener('click', showDeleteConfirmDialog);
        
        // 新增：添加乘客弹窗事件监听器
        closeAddPassengerModalBtn.addEventListener('click', closeAddPassengerModal);
        cancelAddPassengerBtn.addEventListener('click', closeAddPassengerModal);
        confirmAddPassengerBtn.addEventListener('click', confirmAddPassenger);
        
        // 新增：导出排序对话框事件监听器
        exportSortDialogCancelBtn.addEventListener('click', closeExportSortDialog);
        exportSortDialogOkBtn.addEventListener('click', confirmExportSort);
        sortByLineOption.addEventListener('click', () => {
            document.getElementById('sortByLine').checked = true;
            exportSortType = 'line';
        });
        sortByCampusOption.addEventListener('click', () => {
            document.getElementById('sortByCampus').checked = true;
            exportSortType = 'campus';
        });
        
        // 新增：删除确认对话框事件监听器
        deleteConfirmDialogCancelBtn.addEventListener('click', closeDeleteConfirmDialog);
        deleteConfirmDialogOkBtn.addEventListener('click', confirmDeletePassenger);
        
        // 新增：确认对话框事件监听器
        confirmDialogCancelBtn.addEventListener('click', closeConfirmDialog);
        confirmDialogOkBtn.addEventListener('click', confirmRefresh);
        vehicleErrorDialogOkBtn.addEventListener('click', closeVehicleErrorDialog);
        
        // 新增：警告对话框事件监听器
        const warningDialogOverlay = document.getElementById('warningDialogOverlay');
        const confirmWarningBtn = document.getElementById('confirmWarningBtn');
        if (confirmWarningBtn) {
            confirmWarningBtn.addEventListener('click', function() {
                warningDialogOverlay.style.display = 'none';
            });
        }
        
        let uploadedFiles = [];
        
        // 处理文件上传
        function handleFileUpload(event) {
            const files = event.target.files;
            if (!files || files.length === 0) return;
            
            processFiles(Array.from(files));
            
            // 清空文件输入，以便可以重复选择同一文件
            event.target.value = '';
        }
        
        // 处理拖放文件
        function handleDragOver(event) {
            event.preventDefault();
            uploadArea.classList.add('drag-over');
        }
        
        function handleDragLeave(event) {
            event.preventDefault();
            uploadArea.classList.remove('drag-over');
        }
        
        function handleDrop(event) {
            event.preventDefault();
            uploadArea.classList.remove('drag-over');
            
            const files = event.dataTransfer.files;
            if (!files || files.length === 0) return;
            
            processFiles(Array.from(files));
        }
        
        // 处理文件列表
        function processFiles(files) {
            const excelFiles = files.filter(file => file.name.endsWith('.xlsx') || file.name.endsWith('.xls'));
            
            if (excelFiles.length === 0) {
                showNotification('请选择Excel文件', 'error');
                return;
            }
            
            // 清空之前的文件列表
            uploadedFiles = [];
            const fileList = document.getElementById('fileList');
            fileList.innerHTML = '';
            
            // 添加文件到列表
            excelFiles.forEach(file => {
                uploadedFiles.push(file);
                addFileToUI(file);
            });
            
            // 处理第一个文件
            if (excelFiles.length > 0) {
                processSingleFile(excelFiles[0]);
            }
        }
        
        // 添加文件到UI
        function addFileToUI(file) {
            const fileList = document.getElementById('fileList');
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <span>${file.name}</span>
                <button onclick="removeFile('${file.name}')">删除</button>
            `;
            fileList.appendChild(fileItem);
        }
        
        // 移除文件
        window.removeFile = function(fileName) {
            uploadedFiles = uploadedFiles.filter(file => file.name !== fileName);
            const fileItems = document.querySelectorAll('.file-item');
            fileItems.forEach(item => {
                if (item.querySelector('span').textContent === fileName) {
                    item.remove();
                }
            });
        }
        
        // 处理单个文件
        function processSingleFile(file) {
            fileName = file.name;
            fileNameDisplay.textContent = `已选择文件: ${fileName}`;
            
            resetDisplay();
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    workbook = XLSX.read(data, { type: 'array' });
                    
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                    
                    originalData = convertSheetData(jsonData);
                    
                    // 清除之前的手动修改记录
                    manuallyModifiedOrders.clear();
                    manuallyAddedOrders.clear();
                    manuallyDeletedOrders.clear();
                    
                    analyzeBtn.disabled = false;
                    
                    showNotification(`成功加载文件: ${fileName}`, 'success');
                    // 保存原始数据到数据库
                    saveDataToDB(STORES.ORIGINAL_DATA, 'data', originalData);
                    // 清除数据库中的手动修改记录
                    saveDataToDB(STORES.MANUALLY_MODIFIED, 'data', {});
                    saveDataToDB(STORES.MANUALLY_ADDED, 'data', []);
                    saveDataToDB(STORES.MANUALLY_DELETED, 'data', []);
                } catch (error) {
                    console.error('Error reading file:', error);
                    showNotification('读取失败，请联网后重新进入系统', 'error');
                }
            };
            
            reader.readAsArrayBuffer(file);
        }
        
        // 转换工作表数据为对象数组
        function convertSheetData(sheetData) {
            if (sheetData.length < 2) return [];
            
            const headers = sheetData[0].map(h => h ? h.toString().trim() : '');
            
            const findColumnIndex = (keywords) => {
                for (let keyword of keywords) {
                    for (let i = 0; i < headers.length; i++) {
                        if (headers[i] && headers[i].toLowerCase().includes(keyword.toLowerCase())) {
                            return i;
                        }
                    }
                }
                return -1;
            };
            
            const orderIdIndex = findColumnIndex(['订单id', '订单', 'id']);
            const lineIndex = findColumnIndex(['线路', '路线', 'line']);
            const plateIndex = findColumnIndex(['车牌', '车牌号码', '车牌号']);
            const contactIndex = findColumnIndex(['联系人', '姓名', '名字']);
            const phoneIndex = findColumnIndex(['电话', '联系电话', '手机']);
            const ticketIndex = findColumnIndex(['票数', '票', '人数']);
            const noticeIndex = findColumnIndex(['通知', '通知内容', '内容']);
            const dateIndex = findColumnIndex(['日期', '乘车日期', '出发日期']);
            const timeIndex = findColumnIndex(['时间', '乘车时间', '出发时间']);
            const pickupIndex = findColumnIndex(['上车', '上车地点', '出发地']);
            const dropoffIndex = findColumnIndex(['下车', '下车地点', '目的地']);
            const managerIndex = findColumnIndex(['负责人', '随车负责人', '车长']);
            const managerPhoneIndex = findColumnIndex(['负责人电话', '车长电话']);
            const notesIndex = findColumnIndex(['注意事项', '注意', '备注']);
            // 新增：分销员手机列
            const distributorPhoneIndex = findColumnIndex(['分销员手机', '分销员电话', '分销手机', '推广员手机', '分销商手机']);
            
            const data = [];
            for (let i = 1; i < sheetData.length; i++) {
                const row = sheetData[i];
                if (!row || row.length === 0) continue;
                
                let ticketCount = 1;
                if (ticketIndex >= 0 && row[ticketIndex] !== undefined && row[ticketIndex] !== null && row[ticketIndex] !== '') {
                    const ticketValue = Number(row[ticketIndex]);
                    ticketCount = isNaN(ticketValue) ? 1 : ticketValue;
                }
                
                const rowObj = {
                    订单ID: orderIdIndex >= 0 && row[orderIdIndex] ? row[orderIdIndex].toString() : `ORD${1000 + i}`,
                    线路: lineIndex >= 0 && row[lineIndex] ? row[lineIndex].toString() : `线路${(i % 5) + 1}`,
                    车牌号码: "",
                    联系人: contactIndex >= 0 && row[contactIndex] ? row[contactIndex].toString() : `联系人${i}`,
                    联系电话: phoneIndex >= 0 && row[phoneIndex] ? row[phoneIndex].toString() : `138${String(10000000 + i).substring(1)}`,
                    票数: ticketCount,
                    通知内容: "",
                    乘车日期: dateIndex >= 0 && row[dateIndex] ? row[dateIndex].toString() : getRandomDate(),
                    乘车时间: timeIndex >= 0 && row[timeIndex] ? row[timeIndex].toString() : `${8 + (i % 10)}:${i % 2 === 0 ? '00' : '30'}`,
                    上车地点: pickupIndex >= 0 && row[pickupIndex] ? row[pickupIndex].toString() : `地点${String.fromCharCode(65 + (i % 5))}`,
                    下车地点: dropoffIndex >= 0 && row[dropoffIndex] ? row[dropoffIndex].toString() : `地点${String.fromCharCode(70 + (i % 5))}`,
                    随车负责人: "",
                    随车负责人电话: "",
                    修改时间: new Date().toLocaleString('zh-CN'),
                    分销员手机: distributorPhoneIndex >= 0 && row[distributorPhoneIndex] ? row[distributorPhoneIndex].toString() : ''
                };
                
                data.push(rowObj);
            }
            
            return data;
        }
        
        // 生成随机日期
        function getRandomDate() {
            const start = new Date(2023, 0, 1);
            const end = new Date(2023, 11, 31);
            const date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
            return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
        }
        
        // 生成示例数据（用于演示）
        function generateSampleData() {
            const sampleData = [];
            const lines = ['线路 1', '线路 2', '线路 3'];
            const pickupLocations = ['出发地 1', '出发地 2', '出发地 3'];
            const dropoffLocations = [' 目的地 1', '目的地 2', '目的地 3'];
            
            for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
                const line = lines[lineIndex];
                
                let orderCount;
                if (line === '线路 1') orderCount = 1;
                else if (line === ' 线路 2') orderCount = 1;
                else orderCount = 1;
                
                for (let i = 1; i <= orderCount; i++) {
                    const pickupIndex = Math.floor(Math.random() * pickupLocations.length);
                    const pickup = pickupLocations[pickupIndex];
                    
                    const dropoffIndex = lineIndex % dropoffLocations.length;
                    const dropoff = dropoffLocations[dropoffIndex];
                    
                    const ticketCount = 1;
                    
                    sampleData.push({
                        订单ID: `${1000 + (lineIndex * 100 + i)}`,
                        线路: line,
                        车牌号码: "",
                        联系人: `乘客${lineIndex * 100 + i}`,
                        联系电话: `123${String(10000000 + (lineIndex * 100 + i)).substring(1)}`,
                        票数: ticketCount,
                        通知内容: "",
                        乘车日期: '2025-06-15',
                        乘车时间: `${8 + (i % 5)}:${i % 2 === 0 ? '00' : '30'}`,
                        上车地点: pickup,
                        下车地点: dropoff,
                        随车负责人: "",
                        随车负责人电话: "",
                        修改时间: new Date().toLocaleString('zh-CN'),
                        分销员手机: ''
                    });
                }
            }
            
            return sampleData;
        }
        
        // 显示线路信息填写弹窗（已增加线路限制输入框）
        function showLineInfoModal() {
            if (originalData.length === 0) {
                showNotification('请先上传Excel文件', 'error', true);
                return;
            }
            
            const lines = [...new Set(originalData.map(item => item.线路))].sort();
            
            // 按线路分组，收集每个线路的时间、上车地点、下车地点
            const lineTimeMap = {};
            const linePickupMap = {};
            const lineDropoffMap = {};
            
            lines.forEach(line => {
                // 获取当前线路的所有数据
                const lineData = originalData.filter(item => item.线路 === line);
                
                // 收集当前线路的所有时间（去重、过滤空值、排序）
                const times = [...new Set(lineData.map(item => item.乘车时间).filter(time => time && time.trim()))].sort();
                lineTimeMap[line] = times;
                
                // 收集当前线路的所有上车地点（去重、过滤空值、排序）
                const pickups = [...new Set(lineData.map(item => item.上车地点).filter(pickup => pickup && pickup.trim()))].sort();
                linePickupMap[line] = pickups;
                
                // 收集当前线路的所有下车地点（去重、过滤空值、排序）
                const dropoffs = [...new Set(lineData.map(item => item.下车地点).filter(dropoff => dropoff && dropoff.trim()))].sort();
                lineDropoffMap[line] = dropoffs;
            });
            
            lineInputsBody.innerHTML = '';
            
            lines.forEach((line, index) => {
                // 如果有多个车辆，按照顺序字段排序
                let sortedVehicles = [];
                if (lastLineInfo[line] && lastLineInfo[line].length > 0) {
                    sortedVehicles = [...lastLineInfo[line]].sort((a, b) => {
                        // 未设置顺序的车辆排在最后
                        if (a.顺序 === null && b.顺序 !== null) return 1;
                        if (b.顺序 === null && a.顺序 !== null) return -1;
                        if (a.顺序 === null && b.顺序 === null) return 0;
                        
                        // 按顺序字段排序
                        return a.顺序 - b.顺序;
                    });
                }
                
                // 生成第一行
                if (sortedVehicles.length > 0) {
                    for (let i = 0; i < sortedVehicles.length; i++) {
                        const lastInfo = sortedVehicles[i];
                        const row = document.createElement('tr');
                        row.dataset.line = line;
                        row.dataset.plate = "";
                        
                        // 获取当前线路的时间选项
                        const timeOptions = lineTimeMap[line] || [];
                        const timeOptionsHtml = timeOptions.map(time => 
                            `<option value="${time}" ${lastInfo && lastInfo.乘车时间 === time ? 'selected' : ''}>${time}</option>`
                        ).join('');
                        
                        // 获取当前线路的上车地点选项
                        const pickupOptions = linePickupMap[line] || [];
                        const pickupOptionsHtml = pickupOptions.map(pickup => 
                            `<option value="${pickup}" ${lastInfo && lastInfo.上车地点 === pickup ? 'selected' : ''}>${pickup}</option>`
                        ).join('');
                        
                        // 获取当前线路的下车地点选项
                        const dropoffOptions = lineDropoffMap[line] || [];
                        const dropoffOptionsHtml = dropoffOptions.map(dropoff => 
                            `<option value="${dropoff}" ${lastInfo && lastInfo.下车地点 === dropoff ? 'selected' : ''}>${dropoff}</option>`
                        ).join('');
                        
                        // 生成行HTML，注意新增的“线路限制”列放在乘车时间之前
                        row.innerHTML = `
                            <td class="line-action-buttons">
                                <button class="add-line-btn" type="button" data-line="${line}">
                                    <i class="fas fa-plus"></i>
                                </button>
                                <button class="remove-line-btn" type="button" data-line="${line}" ${i === 0 ? 'disabled' : ''}>
                                    <i class="fas fa-minus"></i>
                                </button>
                            </td>
                            <td>${line}</td>
                            <td>
                                <input type="number" class="input-field order-input" id="order-${index}-${i}" data-line="${line}" placeholder="顺序" min="1" value="${lastInfo && lastInfo.顺序 !== null ? lastInfo.顺序 : ''}">
                            </td>
                            <td>
                                <input type="number" class="input-field line-limit-input" id="lineLimit-${index}-${i}" data-line="${line}" placeholder="不限" min="0" value="${lastInfo && lastInfo.线路限制 !== null ? lastInfo.线路限制 : ''}">
                            </td>
                            <td>
                                <select class="input-field time-select" id="time-${index}-${i}" data-line="${line}">
                                    <option value="">--不限时间--</option>
                                    ${timeOptionsHtml}
                                </select>
                            </td>
                            <td>
                                <select class="input-field pickup-select" id="pickup-${index}-${i}" data-line="${line}">
                                    <option value="">--不限上车地点--</option>
                                    ${pickupOptionsHtml}
                                </select>
                            </td>
                            <td>
                                <select class="input-field dropoff-select" id="dropoff-${index}-${i}" data-line="${line}">
                                    <option value="">--不限下车地点--</option>
                                    ${dropoffOptionsHtml}
                                </select>
                            </td>
                            <td>
                                <input type="number" class="input-field capacity-input" id="count-${index}-${i}" data-line="${line}" placeholder="请输入座位数" min="0" value="${lastInfo && lastInfo.人数 !== null ? lastInfo.人数 : ''}">
                            </td>
                            <td><input type="text" class="input-field plate-input" id="plate-${index}-${i}" data-line="${line}" placeholder="请输入车牌号码" value="${lastInfo ? lastInfo.车牌号码 || '' : ''}"></td>
                            <td><input type="text" class="input-field manager-input" id="manager-${index}-${i}" data-line="${line}" placeholder="请输入随车负责人" value="${lastInfo ? lastInfo.随车负责人 || '' : ''}"></td>
                            <td><input type="text" class="input-field manager-phone-input" id="managerPhone-${index}-${i}" data-line="${line}" placeholder="请输入负责人电话" value="${lastInfo ? lastInfo.随车负责人电话 || '' : ''}"></td>
                            <td><input type="text" class="input-field notice-input" id="notice-${index}-${i}" data-line="${line}" placeholder="请输入通知内容" value="${lastInfo ? lastInfo.通知内容 || '' : ''}"></td>
                        `;
                        lineInputsBody.appendChild(row);
                    }
                } else {
                    // 没有车辆信息，生成空行
                    const row = document.createElement('tr');
                    row.dataset.line = line;
                    row.dataset.plate = "";
                    
                    // 获取当前线路的时间选项
                    const timeOptions = lineTimeMap[line] || [];
                    const timeOptionsHtml = timeOptions.map(time => 
                        `<option value="${time}">${time}</option>`
                    ).join('');
                    
                    // 获取当前线路的上车地点选项
                    const pickupOptions = linePickupMap[line] || [];
                    const pickupOptionsHtml = pickupOptions.map(pickup => 
                        `<option value="${pickup}">${pickup}</option>`
                    ).join('');
                    
                    // 获取当前线路的下车地点选项
                    const dropoffOptions = lineDropoffMap[line] || [];
                    const dropoffOptionsHtml = dropoffOptions.map(dropoff => 
                        `<option value="${dropoff}">${dropoff}</option>`
                    ).join('');
                    
                    // 生成行HTML，注意新增的“线路限制”列放在乘车时间之前
                    row.innerHTML = `
                        <td class="line-action-buttons">
                            <button class="add-line-btn" type="button" data-line="${line}">
                                <i class="fas fa-plus"></i>
                            </button>
                            <button class="remove-line-btn" type="button" data-line="${line}" disabled>
                                <i class="fas fa-minus"></i>
                            </button>
                        </td>
                        <td>${line}</td>
                        <td>
                            <input type="number" class="input-field order-input" id="order-${index}" data-line="${line}" placeholder="顺序" min="1">
                        </td>
                        <td>
                            <input type="number" class="input-field line-limit-input" id="lineLimit-${index}" data-line="${line}" placeholder="不限" min="0">
                        </td>
                        <td>
                            <select class="input-field time-select" id="time-${index}" data-line="${line}">
                                <option value="">--不限时间--</option>
                                ${timeOptionsHtml}
                            </select>
                        </td>
                        <td>
                            <select class="input-field pickup-select" id="pickup-${index}" data-line="${line}">
                                <option value="">--不限上车地点--</option>
                                ${pickupOptionsHtml}
                            </select>
                        </td>
                        <td>
                            <select class="input-field dropoff-select" id="dropoff-${index}" data-line="${line}">
                                <option value="">--不限下车地点--</option>
                                ${dropoffOptionsHtml}
                            </select>
                        </td>
                        <td>
                            <input type="number" class="input-field capacity-input" id="count-${index}" data-line="${line}" placeholder="请输入座位数" min="0">
                        </td>
                        <td><input type="text" class="input-field plate-input" id="plate-${index}" data-line="${line}" placeholder="请输入车牌号码"></td>
                        <td><input type="text" class="input-field manager-input" id="manager-${index}" data-line="${line}" placeholder="请输入随车负责人"></td>
                        <td><input type="text" class="input-field manager-phone-input" id="managerPhone-${index}" data-line="${line}" placeholder="请输入负责人电话"></td>
                        <td><input type="text" class="input-field notice-input" id="notice-${index}" data-line="${line}" placeholder="请输入通知内容"></td>
                    `;
                    lineInputsBody.appendChild(row);
                }
            });
            
            updateRemoveButtonsState();
            
            document.querySelectorAll('.add-line-btn').forEach(btn => {
                btn.addEventListener('click', addLineHandler);
            });
            
            document.querySelectorAll('.remove-line-btn').forEach(btn => {
                btn.addEventListener('click', removeLineHandler);
            });
            
            modalOverlay.style.display = 'flex';
        }
        
        // 更新减号按钮状态
        function updateRemoveButtonsState() {
            const rows = document.querySelectorAll('#lineInputsBody tr');
            const lineCounts = {};
            
            rows.forEach(row => {
                const lineCell = row.querySelector('td:nth-child(2)');
                if (lineCell) {
                    const line = lineCell.textContent;
                    lineCounts[line] = (lineCounts[line] || 0) + 1;
                }
            });
            
            rows.forEach(row => {
                const lineCell = row.querySelector('td:nth-child(2)');
                const removeBtn = row.querySelector('.remove-line-btn');
                
                if (lineCell && removeBtn) {
                    const line = lineCell.textContent;
                    removeBtn.disabled = lineCounts[line] <= 1;
                }
            });
        }
        
        // 加号按钮事件处理器（增加一行，同样包含线路限制输入框）
        function addLineHandler() {
            const line = this.getAttribute('data-line');
            const row = this.closest('tr');
            
            // 获取当前线路的时间、上车地点、下车地点选项
            const lineData = originalData.filter(item => item.线路 === line);
            
            // 收集当前线路的所有时间（去重、过滤空值、排序）
            const times = [...new Set(lineData.map(item => item.乘车时间).filter(time => time && time.trim()))].sort();
            const timeOptionsHtml = times.map(time => `<option value="${time}">${time}</option>`).join('');
            
            // 收集当前线路的所有上车地点（去重、过滤空值、排序）
            const pickups = [...new Set(lineData.map(item => item.上车地点).filter(pickup => pickup && pickup.trim()))].sort();
            const pickupOptionsHtml = pickups.map(pickup => `<option value="${pickup}">${pickup}</option>`).join('');
            
            // 收集当前线路的所有下车地点（去重、过滤空值、排序）
            const dropoffs = [...new Set(lineData.map(item => item.下车地点).filter(dropoff => dropoff && dropoff.trim()))].sort();
            const dropoffOptionsHtml = dropoffs.map(dropoff => `<option value="${dropoff}">${dropoff}</option>`).join('');
            
            const newRow = row.cloneNode(true);
            
            const newIndex = document.querySelectorAll('#lineInputsBody tr').length;
            // 为所有输入框更新ID并清空值
            newRow.querySelectorAll('input, select').forEach(input => {
                const oldId = input.id;
                if (oldId) {
                    const field = oldId.split('-')[0];
                    input.id = `${field}-${newIndex}`;
                }
                if (input.tagName === 'INPUT') {
                    if (input.type === 'text' || input.type === 'number') {
                        input.value = "";
                    }
                } else if (input.tagName === 'SELECT') {
                    input.selectedIndex = 0;
                }
            });
            
            // 替换时间、上车地点、下车地点的选项
            const timeSelect = newRow.querySelector('.time-select');
            const pickupSelect = newRow.querySelector('.pickup-select');
            const dropoffSelect = newRow.querySelector('.dropoff-select');
            
            if (timeSelect) {
                timeSelect.innerHTML = `<option value="">--不限时间--</option>${timeOptionsHtml}`;
            }
            
            if (pickupSelect) {
                pickupSelect.innerHTML = `<option value="">--不限上车地点--</option>${pickupOptionsHtml}`;
            }
            
            if (dropoffSelect) {
                dropoffSelect.innerHTML = `<option value="">--不限下车地点--</option>${dropoffOptionsHtml}`;
            }
            
            const addBtn = newRow.querySelector('.add-line-btn');
            const removeBtn = newRow.querySelector('.remove-line-btn');
            addBtn.addEventListener('click', addLineHandler);
            removeBtn.addEventListener('click', removeLineHandler);
            
            row.parentNode.insertBefore(newRow, row.nextSibling);
            
            updateRemoveButtonsState();
        }
        
        // 减号按钮事件处理器
        function removeLineHandler() {
            const row = this.closest('tr');
            const line = this.getAttribute('data-line');
            
            const lineRows = document.querySelectorAll(`tr td:nth-child(2)`);
            let lineCount = 0;
            
            lineRows.forEach(cell => {
                if (cell.textContent === line) {
                    lineCount++;
                }
            });
            
            if (lineCount > 1) {
                row.remove();
                updateRemoveButtonsState();
            }
        }
        
        // 关闭弹窗
        function closeModal() {
            modalOverlay.style.display = 'none';
        }
        
        // 新增：检查车辆信息是否有重复车牌且人数不同
        function checkVehicleInfoConsistency() {
            const plateInfoMap = new Map(); // 车牌 -> {人数, 线路列表}
            vehicleInfoErrors = [];
            
            const lineRows = document.querySelectorAll('#lineInputsBody tr');
            
            lineRows.forEach((row, rowIndex) => {
                const lineCell = row.querySelector('td:nth-child(2)');
                const orderInput = row.querySelector('.order-input');
                const countInput = row.querySelector('.capacity-input');
                const plateInput = row.querySelector('.plate-input');
                
                if (lineCell && plateInput) {
                    const line = lineCell.textContent;
                    const plate = plateInput.value.trim();
                    const order = orderInput ? (orderInput.value.trim() === "" ? null : parseInt(orderInput.value)) : null;
                    const count = countInput.value.trim() === "" ? null : parseInt(countInput.value);
                    
                    if (plate) { // 只检查有车牌的行
                        if (plateInfoMap.has(plate)) {
                            const existingInfo = plateInfoMap.get(plate);
                            
                            // 检查人数是否一致
                            if (existingInfo.count !== count) {
                                // 记录错误
                                if (!existingInfo.errorAdded) {
                                    vehicleInfoErrors.push({
                                        plate: plate,
                                        lines: [...existingInfo.lines, line],
                                        counts: [existingInfo.count, count]
                                    });
                                    existingInfo.errorAdded = true;
                                } else {
                                    // 如果已经记录过错误，更新线路列表
                                    const error = vehicleInfoErrors.find(e => e.plate === plate);
                                    if (error && !error.lines.includes(line)) {
                                        error.lines.push(line);
                                        error.counts.push(count);
                                    }
                                }
                            } else if (existingInfo.order !== null && order !== null && existingInfo.order !== order) {
                                // 检查顺序是否一致
                                if (!existingInfo.errorAdded) {
                                    vehicleInfoErrors.push({
                                        plate: plate,
                                        lines: [...existingInfo.lines, line],
                                        orders: [existingInfo.order, order]
                                    });
                                    existingInfo.errorAdded = true;
                                } else {
                                    // 如果已经记录过错误，更新线路列表
                                    const error = vehicleInfoErrors.find(e => e.plate === plate);
                                    if (error && !error.lines.includes(line)) {
                                        error.lines.push(line);
                                        if (!error.orders) {
                                            error.orders = [existingInfo.order, order];
                                        } else if (!error.orders.includes(order)) {
                                            error.orders.push(order);
                                        }
                                    }
                                }
                            } else {
                                // 人数和顺序都一致，添加线路到列表中
                                if (!existingInfo.lines.includes(line)) {
                                    existingInfo.lines.push(line);
                                }
                            }
                        } else {
                            // 第一次遇到这个车牌
                            plateInfoMap.set(plate, {
                                count: count,
                                order: order,
                                lines: [line],
                                errorAdded: false
                            });
                        }
                    }
                }
            });
            
            return vehicleInfoErrors.length === 0;
        }
        
        // 新增：显示车辆信息错误对话框
        function showVehicleErrorDialog() {
            if (vehicleInfoErrors.length === 0) {
                return false;
            }
            
            // 清空错误列表
            vehicleErrorList.innerHTML = '';
            
            // 填充错误列表
            vehicleInfoErrors.forEach((error, index) => {
                const errorItem = document.createElement('div');
                errorItem.className = 'vehicle-error-item';
                
                let errorDetails = `线路: ${error.lines.join(', ')}<br>`;
                
                if (error.counts) {
                    const countsText = error.counts.map(c => c === null ? '未设置' : `${c}人`).join(', ');
                    errorDetails += `人数设置: ${countsText}<br>`;
                }
                
                if (error.orders) {
                    const ordersText = error.orders.map(o => o === null ? '未设置' : o).join(', ');
                    errorDetails += `顺序设置: ${ordersText}<br>`;
                }
                
                errorItem.innerHTML = `
                    <div>
                        <span class="vehicle-error-plate">${error.plate}</span>
                        <div class="vehicle-error-details">
                            ${errorDetails}
                        </div>
                    </div>
                `;
                
                vehicleErrorList.appendChild(errorItem);
            });
            
            // 显示错误对话框
            vehicleErrorDialogOverlay.style.display = 'flex';
            
            return true;
        }
        
        // 新增：关闭车辆错误对话框
        function closeVehicleErrorDialog() {
            vehicleErrorDialogOverlay.style.display = 'none';
        }
        
        // 使用线路信息处理数据（读取新增的线路限制字段）
        function processDataWithLineInfo() {
            // 首先检查车辆信息是否一致
            if (!checkVehicleInfoConsistency()) {
                // 显示错误对话框
                showVehicleErrorDialog();
                return;
            }
            
            const inputs = document.querySelectorAll('.input-field');
            lineInfo = {};
            // 保留vehicleCapacityMap中的现有车辆信息
            plateLineMap.clear();
            vehicleTotalAssigned.clear(); // 清空车辆总人数跟踪
            
            const lineRows = document.querySelectorAll('#lineInputsBody tr');
            
            // 第一遍：收集所有车辆信息
            const vehicleInfoMap = new Map(); // 车牌 -> 车辆信息
            
            lineRows.forEach((row, rowIndex) => {
                const lineCell = row.querySelector('td:nth-child(2)');
                const orderInput = row.querySelector('.order-input');
                const timeSelect = row.querySelector('.time-select');
                const pickupSelect = row.querySelector('.pickup-select');
                const dropoffSelect = row.querySelector('.dropoff-select');
                const lineLimitInput = row.querySelector('.line-limit-input'); // 新增
                const countInput = row.querySelector('.capacity-input');
                const plateInput = row.querySelector('.plate-input');
                const managerInput = row.querySelector('.manager-input');
                const managerPhoneInput = row.querySelector('.manager-phone-input');
                const noticeInput = row.querySelector('.notice-input');
                
                if (lineCell) {
                    const line = lineCell.textContent;
                    const order = orderInput ? (orderInput.value.trim() === "" ? null : parseInt(orderInput.value)) : null;
                    const time = timeSelect ? timeSelect.value.trim() : "";
                    const pickup = pickupSelect ? pickupSelect.value : "";
                    const dropoff = dropoffSelect ? dropoffSelect.value : "";
                    const lineLimit = lineLimitInput ? (lineLimitInput.value.trim() === "" ? null : parseInt(lineLimitInput.value)) : null; // 读取线路限制
                    const count = countInput ? (countInput.value.trim() === "" ? null : parseInt(countInput.value)) : null;
                    const plate = plateInput ? plateInput.value.trim() : "";
                    const manager = managerInput ? managerInput.value.trim() : "";
                    const managerPhone = managerPhoneInput ? managerPhoneInput.value.trim() : "";
                    const notice = noticeInput ? noticeInput.value.trim() : "";
                    
                    if (plate) {
                        // 如果这个车牌已经存在，使用已存在的信息（确保一致性）
                        if (vehicleInfoMap.has(plate)) {
                            const existingInfo = vehicleInfoMap.get(plate);
                            
                            // 使用已存在的车辆信息
                            if (!lineInfo[line]) {
                                lineInfo[line] = [];
                            }
                            
                            lineInfo[line].push({
                                乘车时间: time,
                                上车地点: pickup,
                                下车地点: dropoff,
                                线路限制: lineLimit, // 保存线路限制
                                顺序: existingInfo.order, // 使用已存在的顺序
                                人数: existingInfo.count, // 使用已存在的人数
                                车牌号码: plate,
                                随车负责人: manager || existingInfo.manager,
                                随车负责人电话: managerPhone || existingInfo.managerPhone,
                                通知内容: notice || existingInfo.notice,
                                已分配人数: 0
                            });
                        } else {
                            // 第一次遇到这个车牌
                            if (!lineInfo[line]) {
                                lineInfo[line] = [];
                            }
                            
                            lineInfo[line].push({
                                乘车时间: time,
                                上车地点: pickup,
                                下车地点: dropoff,
                                线路限制: lineLimit, // 保存线路限制
                                顺序: order,
                                人数: count,
                                车牌号码: plate,
                                随车负责人: manager,
                                随车负责人电话: managerPhone,
                                通知内容: notice,
                                已分配人数: 0
                            });
                            
                            // 保存车辆信息
                            vehicleInfoMap.set(plate, {
                                order: order,
                                count: count,
                                manager: manager,
                                managerPhone: managerPhone,
                                notice: notice
                            });
                            
                            // 初始化车辆总人数跟踪
                            vehicleTotalAssigned.set(plate, 0);
                        }
                        
                        // 添加或更新车辆容量映射
                        if (plate) {
                            vehicleCapacityMap.set(plate, {
                                乘车时间: time,
                                上车地点: pickup,
                                下车地点: dropoff,
                                顺序: order,
                                容量: count,
                                随车负责人: manager,
                                随车负责人电话: managerPhone,
                                通知内容: notice,
                                已分配人数: 0
                            });
                        }
                        
                        if (plate) {
                            if (!plateLineMap.has(plate)) {
                                plateLineMap.set(plate, new Set());
                            }
                            plateLineMap.get(plate).add(line);
                        }
                    } else {
                        // 没有车牌的情况
                        if (!lineInfo[line]) {
                            lineInfo[line] = [];
                        }
                        
                        lineInfo[line].push({
                            乘车时间: time,
                            上车地点: pickup,
                            下车地点: dropoff,
                            线路限制: lineLimit, // 保存线路限制
                            人数: count,
                            车牌号码: plate,
                            随车负责人: manager,
                            随车负责人电话: managerPhone,
                            通知内容: notice,
                            已分配人数: 0
                        });
                    }
                }
            });
            
            // 第二遍：处理没有处理的行（比如没有车牌的）
            lineRows.forEach((row, rowIndex) => {
                const lineCell = row.querySelector('td:nth-child(2)');
                const timeSelect = row.querySelector('.time-select');
                const pickupSelect = row.querySelector('.pickup-select');
                const dropoffSelect = row.querySelector('.dropoff-select');
                const lineLimitInput = row.querySelector('.line-limit-input');
                const countInput = row.querySelector('.capacity-input');
                const plateInput = row.querySelector('.plate-input');
                const managerInput = row.querySelector('.manager-input');
                const managerPhoneInput = row.querySelector('.manager-phone-input');
                const noticeInput = row.querySelector('.notice-input');
                
                if (lineCell) {
                    const line = lineCell.textContent;
                    const time = timeSelect ? timeSelect.value.trim() : "";
                    const pickup = pickupSelect ? pickupSelect.value : "";
                    const dropoff = dropoffSelect ? dropoffSelect.value : "";
                    const lineLimit = lineLimitInput ? (lineLimitInput.value.trim() === "" ? null : parseInt(lineLimitInput.value)) : null;
                    const count = countInput ? (countInput.value.trim() === "" ? null : parseInt(countInput.value)) : null;
                    const plate = plateInput ? plateInput.value.trim() : "";
                    const manager = managerInput ? managerInput.value.trim() : "";
                    const managerPhone = managerPhoneInput ? managerPhoneInput.value.trim() : "";
                    const notice = noticeInput ? noticeInput.value.trim() : "";
                    
                    // 如果这个行已经处理过（有车牌且已经在第一遍处理），跳过
                    if (plate && vehicleInfoMap.has(plate)) {
                        return;
                    }
                    
                    // 处理没有车牌或者车牌未在第一遍处理的情况
                    if (!lineInfo[line]) {
                        lineInfo[line] = [];
                    }
                    
                    // 检查是否已经添加过（可能在第一遍已经添加了没有车牌的行）
                    const alreadyAdded = lineInfo[line].some(item => 
                        item.乘车时间 === time && 
                        item.上车地点 === pickup && 
                        item.下车地点 === dropoff && 
                        item.车牌号码 === plate
                    );
                    
                    if (!alreadyAdded) {
                        lineInfo[line].push({
                            乘车时间: time,
                            上车地点: pickup,
                            下车地点: dropoff,
                            线路限制: lineLimit,
                            人数: count,
                            车牌号码: plate,
                            随车负责人: manager,
                            随车负责人电话: managerPhone,
                            通知内容: notice,
                            已分配人数: 0
                        });
                        
                        if (plate) {
                            if (!plateLineMap.has(plate)) {
                                plateLineMap.set(plate, new Set());
                            }
                            plateLineMap.get(plate).add(line);
                        }
                    }
                }
            });
            
            // 对每条线路的车辆按顺序字段排序
            Object.keys(lineInfo).forEach(line => {
                lineInfo[line].sort((a, b) => {
                    // 未设置顺序的车辆排在最后
                    if (a.顺序 === null && b.顺序 !== null) return 1;
                    if (b.顺序 === null && a.顺序 !== null) return -1;
                    if (a.顺序 === null && b.顺序 === null) return 0;
                    
                    // 按顺序字段排序
                    return a.顺序 - b.顺序;
                });
            });
            
            lastLineInfo = JSON.parse(JSON.stringify(lineInfo));
            
            closeModal();
            
            analyzeData();
            
            // 保存线路信息和车辆容量信息到数据库
            saveDataToDB(STORES.LINE_INFO, 'data', lineInfo);
            saveDataToDB(STORES.LAST_LINE_INFO, 'data', lastLineInfo);
            try {
                saveDataToDB('vehicleCapacityMap', 'data', Object.fromEntries(vehicleCapacityMap));
            } catch (error) {
                console.error('保存vehicleCapacityMap失败:', error);
            }
        }
        
        // 分析数据并按照新的规则整合（增加线路限制检查和手动修改保留）
        async function analyzeData() {
            loading.style.display = 'block';
            
            setTimeout(async () => {
                try {
                    colorMap.clear();
                    lineColorMap.clear();
                    colorIndex = 0;
                    vehiclePickupStats = {};
                    lineUnassignedCount = {};
                    plateStats = {};
                    manuallyAssignedOrders.clear();
                    lineManuallyAssignedCount = {};
                    vehicleTotalAssigned.clear(); // 清空车辆总人数跟踪
                    lineTimeLocationData = []; // 清空线路时间地点统计数据
                    
                    const lines = [...new Set(originalData.map(item => item.线路))].sort();
                    lineOrder = lines;
                    lines.forEach(line => {
                        if (!lineColorMap.has(line)) {
                            lineColorMap.set(line, colorPalette[colorIndex % colorPalette.length]);
                            colorIndex++;
                        }
                        lineUnassignedCount[line] = 0;
                        lineManuallyAssignedCount[line] = 0;
                    });
                    
                    // 收集手动修改的订单和手动添加的订单
                    const manuallyModifiedOrderIds = new Set(manuallyModifiedOrders.keys());
                    const manuallyAddedOrderIds = new Set(manuallyAddedOrders);
                    
                    // 先处理手动修改和手动添加的订单，保留其数据
                    const preservedOrders = [];
                    const remainingOriginalData = [];
                    
                    // 处理手动添加的订单
                    manuallyAddedOrders.forEach(orderId => {
                        // 从processedData中找到手动添加的订单
                        let existingOrder = processedData.find(item => item.订单ID === orderId);
                        if (existingOrder) {
                            preservedOrders.push(existingOrder);
                        } else {
                            // 如果在processedData中找不到，尝试从数据库加载
                            console.log(`手动添加的订单 ${orderId} 在processedData中找不到，尝试从数据库加载`);
                            // 这里可以添加从数据库加载的逻辑
                        }
                    });
                    
                    // 确保所有手动添加的订单都被保留
                    if (preservedOrders.length < manuallyAddedOrders.size) {
                        console.log('警告：部分手动添加的订单未被找到');
                        // 这里可以添加额外的处理逻辑
                    }
                    
                    // 处理原始数据中的订单，区分手动修改和未修改的
                    originalData.forEach(item => {
                        if (manuallyModifiedOrderIds.has(item.订单ID)) {
                            // 查找手动修改后的数据
                            const modifiedData = manuallyModifiedOrders.get(item.订单ID);
                            if (modifiedData) {
                                const preservedOrder = {
                                    ...item,
                                    ...modifiedData
                                };
                                preservedOrders.push(preservedOrder);
                            }
                        } else {
                            remainingOriginalData.push(item);
                        }
                    });
                    
                    // 初始化处理后的数据
                    processedData = [...preservedOrders];
                    
                    const lineTotalTickets = {};
                    
                    // 先计算原始数据中的票数
                    remainingOriginalData.forEach(item => {
                        const line = item.线路;
                        if (!lineTotalTickets[line]) {
                            lineTotalTickets[line] = 0;
                        }
                        lineTotalTickets[line] += (item.票数 || 0);
                    });
                    
                    // 再加上手动添加的订单的票数
                    preservedOrders.forEach(order => {
                        // 只处理手动添加的订单（不在originalData中的订单）
                        const isInOriginalData = originalData.some(item => item.订单ID === order.订单ID);
                        if (!isInOriginalData) {
                            const line = order.线路;
                            if (!lineTotalTickets[line]) {
                                lineTotalTickets[line] = 0;
                            }
                            lineTotalTickets[line] += (order.票数 || 0);
                        }
                    });
                    
                    // 新增：全局车辆已分配人数跟踪（跨线路共享）
                    const vehicleAssignedCount = new Map(); // 车牌 -> 已分配人数（跨线路）
                    // 新增：每个车牌每个线路的已分配人数（用于线路限制检查）
                    const vehicleLineAssignedCount = new Map(); // key: 车牌|线路 -> 已分配人数
                    // 新增：车辆负责人电话映射
                    const vehicleManagerPhoneMap = new Map(); // 车牌 -> 随车负责人电话
                    
                    // 收集车辆负责人信息
                    Object.keys(lineInfo).forEach(line => {
                        const vehicles = lineInfo[line];
                        vehicles.forEach(vehicle => {
                            if (vehicle.车牌号码 && vehicle.随车负责人电话) {
                                vehicleManagerPhoneMap.set(vehicle.车牌号码, vehicle.随车负责人电话);
                            }
                        });
                    });
                    
                    // 先计算已保存订单的分配情况
                    preservedOrders.forEach(order => {
                        const plate = order.车牌号码;
                        const line = order.线路;
                        const tickets = order.票数 || 0;
                        
                        if (plate && plate.trim()) {
                            // 更新全局车辆已分配人数
                            const currentAssigned = vehicleAssignedCount.get(plate) || 0;
                            vehicleAssignedCount.set(plate, currentAssigned + tickets);
                            
                            // 更新线路在该车上的已分配人数
                            const key = `${plate}|${line}`;
                            const currentLineAssigned = vehicleLineAssignedCount.get(key) || 0;
                            vehicleLineAssignedCount.set(key, currentLineAssigned + tickets);
                        }
                    });
                    
                    Object.keys(lineTotalTickets).forEach(line => {
                        const totalTickets = lineTotalTickets[line];
                        const vehicles = lineInfo[line] || [];
                        
                        if (vehicles.length === 0) {
                            vehicles.push({
                                乘车时间: "",
                                上车地点: "",
                                下车地点: "",
                                人数: null,
                                车牌号码: "",
                                随车负责人: "",
                                随车负责人电话: "",
                                通知内容: "",
                                已分配人数: 0,
                                线路限制: null // 默认无限制
                            });
                        }
                        
                        const lineOrders = remainingOriginalData.filter(item => item.线路 === line);
                        
                        // 按照四要素分组：线路、上车地点、下车地点、时间
                        let ordersByFourElements = {};
                        lineOrders.forEach(order => {
                            const key = `${order.线路}|${order.上车地点 || ''}|${order.下车地点 || ''}|${order.乘车时间 || ''}`;
                            if (!ordersByFourElements[key]) {
                                ordersByFourElements[key] = [];
                            }
                            ordersByFourElements[key].push(order);
                        });
                        
                        let totalSpecifiedCapacity = 0;
                        const specifiedVehicles = vehicles.filter(v => v.人数 !== null);
                        specifiedVehicles.forEach(v => {
                            totalSpecifiedCapacity += v.人数;
                        });
                        
                        if (totalSpecifiedCapacity > totalTickets) {
                            showNotification(`警告：线路"${line}"的人数限制(${totalSpecifiedCapacity})大于实际票数(${totalTickets})，请调整人数设置`, 'warning');
                        }
                        
                        vehicles.forEach(v => {
                            v.已分配人数 = 0;
                        });
                        
                        // 收集所有车辆信息（跨线路）
                        const allVehicles = [];
                        Object.keys(lineInfo).forEach(lineKey => {
                            const vehicles = lineInfo[lineKey] || [];
                            allVehicles.push(...vehicles);
                        });
                        
                        // 第一步：处理所有随车负责人的订单，不考虑线路和其他因素
                        const nonManagerOrders = [];
                        lineOrders.forEach(order => {
                            let assigned = false;
                            
                            // 查找匹配的车辆（根据随车负责人电话）
                            for (let vehicle of allVehicles) {
                                const plate = vehicle.车牌号码;
                                const managerPhone = vehicle.随车负责人电话;
                                
                                if (plate && plate.trim() && managerPhone && order.联系电话 && 
                                    String(order.联系电话).trim() === String(managerPhone).trim()) {
                                    // 检查车辆容量是否足够
                                    const orderTickets = parseInt(order.票数) || 0;
                                    const currentAssigned = vehicleAssignedCount.get(plate) || 0;
                                    const totalAfterAssignment = currentAssigned + orderTickets;
                                    
                                    // 即使容量不足，也分配随车负责人的订单
                                    // 查找同车同线路的其他订单，保持通知内容一致
                                    const sameVehicleOrders = processedData.filter(item => 
                                        item.车牌号码 === vehicle.车牌号码 && 
                                        item.线路 === order.线路
                                    );
                                    
                                    // 确定通知内容
                                    let noticeContent = vehicle.通知内容 || order.通知内容;
                                    if (sameVehicleOrders.length > 0) {
                                        const existingNotice = sameVehicleOrders.find(item => item.通知内容);
                                        if (existingNotice && existingNotice.通知内容) {
                                            noticeContent = existingNotice.通知内容;
                                        }
                                    }
                                    
                                    const processedItem = {
                                        ...order,
                                        车牌号码: vehicle.车牌号码,
                                        随车负责人: vehicle.随车负责人,
                                        随车负责人电话: vehicle.随车负责人电话,
                                        通知内容: noticeContent
                                    };
                                    processedData.push(processedItem);
                                    
                                    // 更新全局车辆已分配人数
                                    if (vehicle.车牌号码 && vehicle.车牌号码.trim()) {
                                        vehicleAssignedCount.set(vehicle.车牌号码, totalAfterAssignment);
                                        
                                        if (vehicleCapacityMap.has(vehicle.车牌号码)) {
                                            vehicleCapacityMap.get(vehicle.车牌号码).已分配人数 = totalAfterAssignment;
                                        }
                                        
                                        // 更新线路在该车上的已分配人数
                                        const key = `${vehicle.车牌号码}|${order.线路}`;
                                        const currentLineAssigned = vehicleLineAssignedCount.get(key) || 0;
                                        vehicleLineAssignedCount.set(key, currentLineAssigned + orderTickets);
                                    }
                                    
                                    // 更新车辆对象的已分配人数
                                    vehicle.已分配人数 = totalAfterAssignment;
                                    
                                    assigned = true;
                                    break; // 找到匹配的车辆后停止查找
                                }
                            }
                            
                            if (!assigned) {
                                nonManagerOrders.push(order);
                            }
                        });
                        
                        // 第二步：处理非随车负责人的订单
                        // 清空并重新填充四要素分组：线路、上车地点、下车地点、时间
                        ordersByFourElements = {};
                        nonManagerOrders.forEach(order => {
                            const key = `${order.线路}|${order.上车地点 || ''}|${order.下车地点 || ''}|${order.乘车时间 || ''}`;
                            if (!ordersByFourElements[key]) {
                                ordersByFourElements[key] = [];
                            }
                            ordersByFourElements[key].push(order);
                        });
                        
                        // 第三步：处理有明确人数限制的车辆
                        specifiedVehicles.forEach(vehicle => {
                            let capacity = vehicle.人数;
                            let assigned = 0;
                            
                            // 检查车辆已分配的总人数（跨线路）
                            const plate = vehicle.车牌号码;
                            if (plate && plate.trim()) {
                                const alreadyAssigned = vehicleAssignedCount.get(plate) || 0;
                                const remainingCapacity = capacity - alreadyAssigned;
                                
                                if (remainingCapacity <= 0) {
                                    // 车辆已满，跳过
                                    console.log(`车牌 ${plate} 已满，跳过分配`);
                                    return;
                                }
                                
                                capacity = remainingCapacity; // 使用剩余容量
                            }
                            
                            const vehicleTime = vehicle.乘车时间 || "";
                            const vehiclePickup = vehicle.上车地点 || "";
                            const vehicleDropoff = vehicle.下车地点 || "";
                            const lineLimit = vehicle.线路限制; // 线路限制（可选）
                            
                            // 获取匹配车辆条件的四要素组
                            const matchingGroups = Object.keys(ordersByFourElements).filter(groupKey => {
                                const [groupLine, pickup, dropoff, time] = groupKey.split('|');
                                
                                if (vehicleTime && time !== vehicleTime) {
                                    return false;
                                }
                                
                                if (vehiclePickup && pickup !== vehiclePickup) {
                                    return false;
                                }
                                
                                if (vehicleDropoff && dropoff !== vehicleDropoff) {
                                    return false;
                                }
                                
                                return true;
                            });
                            
                            // 按组内人数从大到小排序
                            matchingGroups.sort((a, b) => {
                                const aTotal = ordersByFourElements[a].reduce((sum, order) => sum + (order.票数 || 0), 0);
                                const bTotal = ordersByFourElements[b].reduce((sum, order) => sum + (order.票数 || 0), 0);
                                return bTotal - aTotal;
                            });
                            
                            // 分配逻辑：优先分配能凑满的组，并考虑线路限制
                            for (let groupKey of matchingGroups) {
                                if (assigned >= capacity) break;
                                
                                const groupOrders = ordersByFourElements[groupKey];
                                if (!groupOrders || groupOrders.length === 0) continue;
                                
                                const groupTotalTickets = groupOrders.reduce((sum, order) => sum + (order.票数 || 0), 0);
                                
                                // 计算该线路在该车辆上已分配的人数
                                const lineKey = `${plate}|${line}`;
                                const alreadyLineAssigned = vehicleLineAssignedCount.get(lineKey) || 0;
                                
                                // 检查线路限制
                                let lineRemaining = Infinity;
                                if (lineLimit !== null && lineLimit > 0) {
                                    lineRemaining = lineLimit - alreadyLineAssigned;
                                    if (lineRemaining <= 0) {
                                        // 该线路在此车上已达上限，跳过此车辆
                                        console.log(`线路 ${line} 在车牌 ${plate} 上已达上限 ${lineLimit} 人，跳过`);
                                        break; // 跳出循环，不再分配该车辆
                                    }
                                }
                                
                                if (groupTotalTickets <= Math.min(capacity - assigned, lineRemaining)) {
                                    // 整组分配
                                    groupOrders.forEach(order => {
                                        const processedItem = {
                                            ...order,
                                            车牌号码: vehicle.车牌号码,
                                            随车负责人: vehicle.随车负责人,
                                            随车负责人电话: vehicle.随车负责人电话,
                                            通知内容: vehicle.通知内容 || order.通知内容
                                        };
                                        processedData.push(processedItem);
                                        
                                        const orderTickets = order.票数 || 0;
                                        assigned += orderTickets;
                                        vehicle.已分配人数 = assigned;
                                        
                                        // 更新全局车辆已分配人数
                                        if (vehicle.车牌号码 && vehicle.车牌号码.trim()) {
                                            const currentAssigned = vehicleAssignedCount.get(vehicle.车牌号码) || 0;
                                            vehicleAssignedCount.set(vehicle.车牌号码, currentAssigned + orderTickets);
                                            
                                            if (vehicleCapacityMap.has(vehicle.车牌号码)) {
                                                vehicleCapacityMap.get(vehicle.车牌号码).已分配人数 = currentAssigned + orderTickets;
                                            }
                                            
                                            // 更新线路在该车上的已分配人数
                                            const key = `${vehicle.车牌号码}|${line}`;
                                            const currentLineAssigned = vehicleLineAssignedCount.get(key) || 0;
                                            vehicleLineAssignedCount.set(key, currentLineAssigned + orderTickets);
                                        }
                                    });
                                    
                                    // 从原始数据中移除已分配的订单
                                    delete ordersByFourElements[groupKey];
                                } else {
                                    // 组人数超过剩余容量或线路限制，尝试部分分配
                                    let remainingInGroup = groupTotalTickets;
                                    const ordersToAssign = [];
                                    const ordersToKeep = [];
                                    
                                    // 按订单票数从大到小排序
                                    const sortedOrders = [...groupOrders].sort((a, b) => {
                                        return (b.票数 || 0) - (a.票数 || 0);
                                    });
                                    
                                    for (let order of sortedOrders) {
                                        if (assigned >= capacity) {
                                            ordersToKeep.push(order);
                                            continue;
                                        }
                                        
                                        const orderTickets = order.票数 || 0;
                                        // 检查线路限制
                                        const lineKey = `${plate}|${line}`;
                                        const currentLineAssigned = vehicleLineAssignedCount.get(lineKey) || 0;
                                        let lineRemainingNow = Infinity;
                                        if (lineLimit !== null) {
                                            lineRemainingNow = lineLimit - currentLineAssigned;
                                            if (lineRemainingNow <= 0) {
                                                // 线路已达上限，剩余订单不能分配到此车
                                                ordersToKeep.push(order);
                                                continue;
                                            }
                                        }
                                        
                                        if (assigned + orderTickets <= capacity && orderTickets <= lineRemainingNow) {
                                            ordersToAssign.push(order);
                                            assigned += orderTickets;
                                            vehicle.已分配人数 = assigned;
                                            
                                            // 更新全局车辆已分配人数
                                            if (vehicle.车牌号码 && vehicle.车牌号码.trim()) {
                                                const currentAssigned = vehicleAssignedCount.get(vehicle.车牌号码) || 0;
                                                vehicleAssignedCount.set(vehicle.车牌号码, currentAssigned + orderTickets);
                                                
                                                if (vehicleCapacityMap.has(vehicle.车牌号码)) {
                                                    vehicleCapacityMap.get(vehicle.车牌号码).已分配人数 = currentAssigned + orderTickets;
                                                }
                                                
                                                // 更新线路在该车上的已分配人数
                                                const key = `${vehicle.车牌号码}|${line}`;
                                                const newLineAssigned = (vehicleLineAssignedCount.get(key) || 0) + orderTickets;
                                                vehicleLineAssignedCount.set(key, newLineAssigned);
                                            }
                                        } else {
                                            ordersToKeep.push(order);
                                        }
                                    }
                                    
                                    // 分配符合条件的订单
                                    ordersToAssign.forEach(order => {
                                        const processedItem = {
                                            ...order,
                                            车牌号码: vehicle.车牌号码,
                                            随车负责人: vehicle.随车负责人,
                                            随车负责人电话: vehicle.随车负责人电话,
                                            通知内容: vehicle.通知内容 || order.通知内容
                                        };
                                        processedData.push(processedItem);
                                    });
                                    
                                    // 更新组内剩余订单
                                    if (ordersToKeep.length > 0) {
                                        ordersByFourElements[groupKey] = ordersToKeep;
                                    } else {
                                        delete ordersByFourElements[groupKey];
                                    }
                                }
                            }
                        });
                        
                        // 第二步：处理无明确人数限制的车辆
                        const unspecifiedVehicles = vehicles.filter(v => v.人数 === null);
                        const remainingOrders = [];
                        
                        // 收集所有剩余的订单
                        Object.keys(ordersByFourElements).forEach(groupKey => {
                            const groupOrders = ordersByFourElements[groupKey];
                            if (groupOrders && groupOrders.length > 0) {
                                remainingOrders.push(...groupOrders);
                            }
                        });
                        
                        if (remainingOrders.length > 0) {
                            if (unspecifiedVehicles.length > 0) {
                                // 将剩余订单按照时间分组
                                const ordersByTime = {};
                                remainingOrders.forEach(order => {
                                    const time = order.乘车时间 || '';
                                    if (!ordersByTime[time]) {
                                        ordersByTime[time] = [];
                                    }
                                    ordersByTime[time].push(order);
                                });
                                
                                const times = Object.keys(ordersByTime);
                                
                                times.forEach(time => {
                                    const timeOrders = ordersByTime[time];
                                    let currentVehicleIndex = 0;
                                    
                                    // 先找匹配时间的车辆
                                    let availableVehicles = unspecifiedVehicles.filter(v => 
                                        v.乘车时间 === '' || v.乘车时间 === time
                                    );
                                    
                                    if (availableVehicles.length === 0) {
                                        availableVehicles = unspecifiedVehicles;
                                    }
                                    
                                    timeOrders.forEach(order => {
                                        let vehicle = availableVehicles[currentVehicleIndex % availableVehicles.length];
                                        
                                        // 检查线路限制
                                        const plate = vehicle.车牌号码 || '';
                                        const lineLimit = vehicle.线路限制;
                                        if (plate && lineLimit !== null) {
                                            const lineKey = `${plate}|${line}`;
                                            const currentLineAssigned = vehicleLineAssignedCount.get(lineKey) || 0;
                                            if (currentLineAssigned >= lineLimit) {
                                                // 该线路在此车上已达上限，尝试下一个车辆
                                                let found = false;
                                                for (let i = 1; i < availableVehicles.length; i++) {
                                                    const nextVehicle = availableVehicles[(currentVehicleIndex + i) % availableVehicles.length];
                                                    const nextPlate = nextVehicle.车牌号码 || '';
                                                    const nextLimit = nextVehicle.线路限制;
                                                    if (!nextPlate || nextLimit === null) {
                                                        // 无限制或未分配，可以使用
                                                        vehicle = nextVehicle;
                                                        currentVehicleIndex = (currentVehicleIndex + i) % availableVehicles.length;
                                                        found = true;
                                                        break;
                                                    }
                                                    const nextKey = `${nextPlate}|${line}`;
                                                    const nextAssigned = vehicleLineAssignedCount.get(nextKey) || 0;
                                                    if (nextAssigned < nextLimit) {
                                                        vehicle = nextVehicle;
                                                        currentVehicleIndex = (currentVehicleIndex + i) % availableVehicles.length;
                                                        found = true;
                                                        break;
                                                    }
                                                }
                                                if (!found) {
                                                    console.warn(`线路 ${line} 在所有无限制车辆上已达上限，订单无法分配`);
                                                    return;
                                                }
                                            }
                                        }
                                        
                                        const processedItem = {
                                            ...order,
                                            车牌号码: vehicle.车牌号码,
                                            随车负责人: vehicle.随车负责人,
                                            随车负责人电话: vehicle.随车负责人电话,
                                            通知内容: vehicle.通知内容 || order.通知内容
                                        };
                                        processedData.push(processedItem);
                                        
                                        const orderTickets = order.票数 || 0;
                                        vehicle.已分配人数 += orderTickets;
                                        
                                        // 更新全局车辆已分配人数
                                        if (vehicle.车牌号码 && vehicle.车牌号码.trim()) {
                                            const currentAssigned = vehicleAssignedCount.get(vehicle.车牌号码) || 0;
                                            vehicleAssignedCount.set(vehicle.车牌号码, currentAssigned + orderTickets);
                                            
                                            if (vehicleCapacityMap.has(vehicle.车牌号码)) {
                                                vehicleCapacityMap.get(vehicle.车牌号码).已分配人数 = currentAssigned + orderTickets;
                                            }
                                            
                                            // 更新线路在该车上的已分配人数
                                            const key = `${vehicle.车牌号码}|${line}`;
                                            const newLineAssigned = (vehicleLineAssignedCount.get(key) || 0) + orderTickets;
                                            vehicleLineAssignedCount.set(key, newLineAssigned);
                                        }
                                        
                                        currentVehicleIndex++;
                                    });
                                });
                            } else {
                                // 没有无限制车辆，创建未分配车辆
                                const unassignedVehicle = {
                                    乘车时间: "",
                                    上车地点: "",
                                    下车地点: "",
                                    人数: null,
                                    车牌号码: "",
                                    随车负责人: "",
                                    随车负责人电话: "",
                                    通知内容: "",
                                    已分配人数: 0,
                                    线路限制: null
                                };
                                
                                remainingOrders.forEach(order => {
                                    const processedItem = {
                                        ...order,
                                        车牌号码: unassignedVehicle.车牌号码,
                                        随车负责人: unassignedVehicle.随车负责人,
                                        随车负责人电话: unassignedVehicle.随车负责人电话,
                                        通知内容: unassignedVehicle.通知内容 || order.通知内容
                                    };
                                    processedData.push(processedItem);
                                    
                                    unassignedVehicle.已分配人数 += (order.票数 || 0);
                                });
                                
                                lineUnassignedCount[line] = unassignedVehicle.已分配人数;
                                
                                if (!lineInfo[line]) {
                                    lineInfo[line] = [];
                                }
                                lineInfo[line].push(unassignedVehicle);
                            }
                        }
                        
                        // 第三步：尝试将未分配的订单分配到有剩余容量的车辆中（考虑线路限制）
                        const vehiclesWithCapacity = vehicles.filter(v => v.人数 !== null && v.已分配人数 < v.人数);
                        const unassignedOrdersInLine = processedData.filter(item => 
                            item.线路 === line && (!item.车牌号码 || item.车牌号码.trim() === "")
                        );
                        
                        if (vehiclesWithCapacity.length > 0 && unassignedOrdersInLine.length > 0) {
                            vehiclesWithCapacity.forEach(vehicle => {
                                // 检查车辆剩余容量（考虑跨线路已分配人数）
                                const plate = vehicle.车牌号码;
                                let remainingCapacity = vehicle.人数 - vehicle.已分配人数;
                                
                                if (plate && plate.trim()) {
                                    const alreadyAssigned = vehicleAssignedCount.get(plate) || 0;
                                    const vehicleCapacity = vehicleCapacityMap.has(plate) ? vehicleCapacityMap.get(plate).容量 : vehicle.人数;
                                    remainingCapacity = vehicleCapacity - alreadyAssigned;
                                    
                                    if (remainingCapacity <= 0) {
                                        return; // 车辆已满
                                    }
                                }
                                
                                const vehicleTime = vehicle.乘车时间 || "";
                                const vehiclePickup = vehicle.上车地点 || "";
                                const vehicleDropoff = vehicle.下车地点 || "";
                                const lineLimit = vehicle.线路限制;
                                
                                // 计算线路在该车上已分配人数
                                let currentLineAssigned = 0;
                                if (plate) {
                                    const key = `${plate}|${line}`;
                                    currentLineAssigned = vehicleLineAssignedCount.get(key) || 0;
                                }
                                
                                let lineRemaining = Infinity;
                                if (lineLimit !== null) {
                                    lineRemaining = lineLimit - currentLineAssigned;
                                    if (lineRemaining <= 0) {
                                        return; // 线路已达上限
                                    }
                                }
                                
                                const matchingUnassigned = unassignedOrdersInLine.filter(order => {
                                    if (vehicleTime && order.乘车时间 !== vehicleTime) {
                                        return false;
                                    }
                                    
                                    if (vehiclePickup && order.上车地点 !== vehiclePickup) {
                                        return false;
                                    }
                                    
                                    if (vehicleDropoff && order.下车地点 !== vehicleDropoff) {
                                        return false;
                                    }
                                    
                                    return (order.票数 || 0) <= Math.min(remainingCapacity, lineRemaining);
                                });
                                
                                matchingUnassigned.forEach(order => {
                                    const orderIndex = processedData.findIndex(item => item.订单ID === order.订单ID);
                                    if (orderIndex !== -1) {
                                        const orderTickets = order.票数 || 0;
                                        
                                        processedData[orderIndex].车牌号码 = vehicle.车牌号码;
                                        processedData[orderIndex].随车负责人 = vehicle.随车负责人;
                                        processedData[orderIndex].随车负责人电话 = vehicle.随车负责人电话;
                                        processedData[orderIndex].通知内容 = vehicle.通知内容 || order.通知内容;
                                        
                                        vehicle.已分配人数 += orderTickets;
                                        
                                        // 更新全局车辆已分配人数
                                        if (vehicle.车牌号码 && vehicle.车牌号码.trim()) {
                                            const currentAssigned = vehicleAssignedCount.get(vehicle.车牌号码) || 0;
                                            vehicleAssignedCount.set(vehicle.车牌号码, currentAssigned + orderTickets);
                                            
                                            if (vehicleCapacityMap.has(vehicle.车牌号码)) {
                                                vehicleCapacityMap.get(vehicle.车牌号码).已分配人数 = currentAssigned + orderTickets;
                                            }
                                            
                                            // 更新线路在该车上的已分配人数
                                            const key = `${vehicle.车牌号码}|${line}`;
                                            const newLineAssigned = (vehicleLineAssignedCount.get(key) || 0) + orderTickets;
                                            vehicleLineAssignedCount.set(key, newLineAssigned);
                                        }
                                        
                                        lineUnassignedCount[line] -= orderTickets;
                                        if (lineUnassignedCount[line] < 0) lineUnassignedCount[line] = 0;
                                    }
                                });
                            });
                        }
                    });
                    
                    // 保存车辆总分配人数到全局变量
                    vehicleAssignedCount.forEach((value, key) => {
                        vehicleTotalAssigned.set(key, value);
                    });
                    
                    processedData.sort((a, b) => {
                        const plateA = a.车牌号码 || "";
                        const plateB = b.车牌号码 || "";
                        const lineA = a.线路;
                        const lineB = b.线路;
                        
                        if (plateA && plateB) {
                            if (plateA !== plateB) {
                                return plateA.localeCompare(plateB);
                            } else {
                                return lineA.localeCompare(lineB);
                            }
                        } else if (plateA && !plateB) {
                            return -1;
                        } else if (!plateA && plateB) {
                            return 1;
                        } else {
                            return lineA.localeCompare(lineB);
                        }
                    });
                    
                    lineUnassignedCount = {};
                    lines.forEach(line => {
                        lineUnassignedCount[line] = 0;
                    });
                    
                    processedData.forEach(item => {
                        const line = item.线路;
                        const plate = item.车牌号码 || "";
                        const ticketCount = item.票数 || 0;
                        
                        if (!plate || plate.trim() === "" || 
                            plate.includes("未分配") || 
                            plate.includes("未匹配") ||
                            plate.includes("未安排") ||
                            plate === "未分配车辆") {
                            lineUnassignedCount[line] += ticketCount;
                        }
                    });
                    
                    isSearchActive = false;
                    filteredData = [];
                    searchResultsInfo.textContent = '';
                    
                    // 清空搜索输入框
                    searchContact.value = '';
                    searchPhone.value = '';
                    searchPlate.value = '';
                    searchLine.value = '';
                    searchPickup.value = '';
                    searchDropoff.value = '';
                    
                    selectedRows.clear();
                    selectAllCheckbox.checked = false;
                    updateModifyButtonState();
                    selectAllPlateBtn.style.display = 'none';
                    selectAllUnassignedBtn.style.display = 'none';
                    
                    displayLineTicketStats();
                    
                    // 新增：显示线路时间地点统计
                    displayLineTimeLocationStats();
                    
                    ticketStatsSection.style.display = 'block';
                    
                    toggleBtn.disabled = false;
                    exportBtn.disabled = false;
                    
                    // 直接显示整合后的数据和车次统计
                    dataSection.style.display = 'flex';
                    vehiclePickupStatsPanel.style.display = 'block';
                    displayProcessedData();
                    displayVehiclePickupStatsPanel();
                    ticketStatsSection.classList.add('ticket-stats-down');
                    
                    // 隐藏切换按钮，但保留其功能
                    toggleBtn.style.display = 'none';
                    
                    showNotification('已成功分配各车辆信息（考虑线路人数限制）', 'success');
                    
                    // 对processedData进行排序，确保随车联系人的订单排在前面
                    processedData.sort((a, b) => {
                        // 首先按车牌号分组
                        const plateA = a.车牌号码 || '未分配车辆';
                        const plateB = b.车牌号码 || '未分配车辆';
                        if (plateA !== plateB) {
                            // 未分配车辆始终排在最后
                            if (plateA === '未分配车辆' || plateA === '') return 1;
                            if (plateB === '未分配车辆' || plateB === '') return -1;
                            
                            // 按照顺序排序
                            const orderA = vehicleCapacityMap.has(plateA) ? (vehicleCapacityMap.get(plateA).顺序 || 9999) : 9999;
                            const orderB = vehicleCapacityMap.has(plateB) ? (vehicleCapacityMap.get(plateB).顺序 || 9999) : 9999;
                            if (orderA !== orderB) {
                                return orderA - orderB;
                            }
                            
                            // 顺序相同，按车牌号排序
                            return plateA.localeCompare(plateB);
                        }
                        
                        // 同一车辆内，优先将随车联系人的订单排在前面
                        const isManagerA = a.联系电话 === a.随车负责人电话;
                        const isManagerB = b.联系电话 === b.随车负责人电话;
                        if (isManagerA && !isManagerB) return -1;
                        if (!isManagerA && isManagerB) return 1;
                        
                        // 然后按线路排序
                        const lineA = a.线路 || '';
                        const lineB = b.线路 || '';
                        if (lineA !== lineB) {
                            return lineA.localeCompare(lineB);
                        }
                        
                        // 线路相同，按上车地点排序
                        const pickupA = a.上车地点 || '';
                        const pickupB = b.上车地点 || '';
                        if (pickupA !== pickupB) {
                            return pickupA.localeCompare(pickupB);
                        }
                        
                        return 0;
                    });
                    
                    // 自动保存数据到数据库
                    await saveAllData();
                } catch (error) {
                    console.error('Error analyzing data:', error);
                    showNotification('数据处理失败', 'error');
                } finally {
                    loading.style.display = 'none';
                }
            }, 800);
        }
        
        // 切换数据显示
        function toggleDataView() {
            isDataVisible = !isDataVisible;
            
            if (isDataVisible) {
                dataSection.style.display = 'flex';
                toggleBtn.innerHTML = '<i class="fas fa-eye-slash"></i> 隐藏整合后的数据';
                
                displayProcessedData();
                displayVehiclePickupStatsPanel();
                
                ticketStatsSection.classList.add('ticket-stats-down');
            } else {
                dataSection.style.display = 'none';
                toggleBtn.innerHTML = '<i class="fas fa-eye"></i> 显示整合后的数据';
                
                vehiclePickupStatsPanel.style.display = 'none';
                
                ticketStatsSection.classList.remove('ticket-stats-down');
            }
        }
        
        // 显示处理后的数据
        function displayProcessedData() {
            processedDataBody.innerHTML = '';
            
            const dataToDisplay = isSearchActive ? filteredData : processedData;
            
            if (dataToDisplay.length === 0) {
                processedDataBody.innerHTML = '<tr><td colspan="15" style="text-align:center;">没有数据</td></tr>';
                return;
            }
            
            vehiclePickupStats = {};
            plateStats = {};
            
            dataToDisplay.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                const line = item.线路 || '未知线路';
                const pickup = item.上车地点 || '未知上车地点';
                const dropoff = item.下车地点 || '未知下车地点';
                const tickets = item.票数 || 0;
                
                const key = `${plate}|${line}`;
                
                if (!vehiclePickupStats[key]) {
                    vehiclePickupStats[key] = {
                        plate: plate,
                        line: line,
                        stats: {}
                    };
                }
                
                const locationKey = `${pickup}→${dropoff}`;
                if (!vehiclePickupStats[key].stats[locationKey]) {
                    vehiclePickupStats[key].stats[locationKey] = 0;
                }
                
                vehiclePickupStats[key].stats[locationKey] += tickets;
                
                if (!plateStats[plate]) {
                    plateStats[plate] = {
                        total: 0,
                        lines: new Set()
                    };
                }
                plateStats[plate].total += tickets;
                plateStats[plate].lines.add(line);
            });
            
            // 按照与导出文件相同的逻辑对数据进行排序
            const plateGroups = {};
            dataToDisplay.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                if (!plateGroups[plate]) {
                    plateGroups[plate] = [];
                }
                plateGroups[plate].push(item);
            });
            
            // 获取修改时间排序方式
            const modifyTimeSort = document.getElementById('modifyTimeSort').value;
            
            // 按照车辆顺序和车牌号排序
            const plates = Object.keys(plateGroups).sort((a, b) => {
                // 未分配车辆始终排在最后
                if (a === '未分配车辆' || a === '') return 1;
                if (b === '未分配车辆' || b === '') return -1;
                
                // 获取车辆顺序
                const orderA = vehicleCapacityMap.has(a) ? (vehicleCapacityMap.get(a).顺序 || 9999) : 9999;
                const orderB = vehicleCapacityMap.has(b) ? (vehicleCapacityMap.get(b).顺序 || 9999) : 9999;
                
                // 按照顺序排序
                if (orderA !== orderB) {
                    return orderA - orderB;
                }
                
                // 顺序相同，按车牌号排序
                return a.localeCompare(b);
            });
            
            // 对每个车辆组内的订单进行排序
            const sortedData = [];
            plates.forEach(plate => {
                let groupData = plateGroups[plate];
                
                // 优先将随车联系人的订单排在前面
                groupData.sort((a, b) => {
                    // 检查是否是随车联系人
                    const isManagerA = a.联系电话 === a.随车负责人电话;
                    const isManagerB = b.联系电话 === b.随车负责人电话;
                    
                    if (isManagerA && !isManagerB) return -1;
                    if (!isManagerA && isManagerB) return 1;
                    
                    // 然后根据修改时间排序方式排序
                    if (modifyTimeSort !== 'none') {
                        const timeA = new Date(a.修改时间 || 0);
                        const timeB = new Date(b.修改时间 || 0);
                        if (modifyTimeSort === 'asc') {
                            return timeA - timeB;
                        } else if (modifyTimeSort === 'desc') {
                            return timeB - timeA;
                        }
                    } else {
                        // 然后根据选择的排序方式排序
                        if (exportSortType === 'line') {
                            // 按线路排序
                            const lineA = a.线路 || '';
                            const lineB = b.线路 || '';
                            return lineA.localeCompare(lineB);
                        } else if (exportSortType === 'campus') {
                            // 按校区（上车地点）排序
                            const pickupA = a.上车地点 || '';
                            const pickupB = b.上车地点 || '';
                            return pickupA.localeCompare(pickupB);
                        }
                    }
                    
                    return 0;
                });
                
                sortedData.push(...groupData);
            });
            
            let currentPlate = '';
            let currentLine = '';
            
            sortedData.forEach((item, index) => {
                const row = document.createElement('tr');
                row.dataset.orderId = item.订单ID;
                row.dataset.line = item.线路;
                row.dataset.plate = item.车牌号码 || '';
                
                if (selectedRows.has(item.订单ID)) {
                    row.classList.add('selected-row');
                }
                
                if (item.车牌号码 !== currentPlate) {
                    currentPlate = item.车牌号码;
                    currentLine = '';
                    
                    const plateSeparatorRow = document.createElement('tr');
                    
                    const isUnassigned = !currentPlate || currentPlate === '未分配车辆';
                    if (isUnassigned) {
                        plateSeparatorRow.className = 'unassigned-plate-group';
                    } else {
                        plateSeparatorRow.className = 'plate-group';
                    }
                    
                    let seatStatus = '';
                    let plateTotal = plateStats[currentPlate] ? plateStats[currentPlate].total : 0;
                    
                    if (currentPlate && vehicleCapacityMap.has(currentPlate)) {
                        const vehicleInfo = vehicleCapacityMap.get(currentPlate);
                        const capacity = vehicleInfo.容量;
                        const vehicleTime = vehicleInfo.乘车时间 || '';
                        const timeInfo = vehicleTime ? `时间:${vehicleTime} ` : '';
                        
                        // 检查车辆总分配人数（跨线路）
                        const totalAssigned = vehicleTotalAssigned.get(currentPlate) || plateTotal;
                        
                        if (capacity !== null) {
                            if (totalAssigned >= capacity) {
                                seatStatus = `<span class="seat-status seat-status-full">满员(${capacity}座, 已分配${totalAssigned}人)</span>`;
                            } else {
                                const emptySeats = capacity - totalAssigned;
                                seatStatus = `<span class="seat-status seat-status-available">${timeInfo}${totalAssigned}/${capacity}人(空${emptySeats}座)</span>`;
                            }
                        } else {
                            seatStatus = `<span class="seat-status seat-status-unlimited">${timeInfo}${totalAssigned}人(未限制)</span>`;
                        }
                    } else {
                        seatStatus = `<span class="seat-status seat-status-unassigned">未分配车辆</span>`;
                    }
                    
                    let lineText = '';
                    if (currentPlate && plateLineMap.has(currentPlate)) {
                        const lines = Array.from(plateLineMap.get(currentPlate));
                        if (lines.length === 1) {
                            lineText = `<span class="line-in-plate">${lines[0]}</span> - `;
                        } else if (lines.length > 1) {
                            lineText = `<span class="line-in-plate">${lines.join(',')}</span> - `;
                        }
                    } else if (plateStats[currentPlate] && plateStats[currentPlate].lines.size > 0) {
                        const lines = Array.from(plateStats[currentPlate].lines);
                        if (lines.length === 1) {
                            lineText = `<span class="line-in-plate">${lines[0]}</span> - `;
                        } else if (lines.length > 1) {
                            lineText = `<span class="line-in-plate">${lines.join(',')}</span> - `;
                        }
                    }
                    
                    const plateText = currentPlate ? `${lineText}车牌号码: ${currentPlate} ${seatStatus}` : `未分配车辆 ${seatStatus}`;
                    
                    const plateCheckboxId = `plate-checkbox-${currentPlate || 'unassigned'}`;
                    plateSeparatorRow.innerHTML = `<td colspan="15" style="background-color: ${isUnassigned ? '#fff5f5' : '#f0f0f0'}; font-weight: bold; padding: 8px;">
                        <div class="plate-header">
                            <input type="checkbox" class="plate-checkbox" id="${plateCheckboxId}" data-plate="${currentPlate || 'unassigned'}">
                            <label for="${plateCheckboxId}" style="cursor: pointer;">${plateText}</label>
                        </div>
                    </td>`;
                    processedDataBody.appendChild(plateSeparatorRow);
                    
                    const plateCheckbox = plateSeparatorRow.querySelector('.plate-checkbox');
                    plateCheckbox.addEventListener('change', function() {
                        const plate = this.getAttribute('data-plate');
                        const plateRows = processedDataBody.querySelectorAll(`tr[data-plate="${plate}"]`);
                        
                        plateRows.forEach(row => {
                            const checkbox = row.querySelector('.row-checkbox');
                            const orderId = row.dataset.orderId;
                            
                            if (this.checked) {
                                checkbox.checked = true;
                                selectedRows.add(orderId);
                                row.classList.add('selected-row');
                            } else {
                                checkbox.checked = false;
                                selectedRows.delete(orderId);
                                row.classList.remove('selected-row');
                                selectAllCheckbox.checked = false;
                            }
                        });
                        
                        updateModifyButtonState();
                    });
                    
                    const key = `${currentPlate}|${item.线路}`;
                    if (vehiclePickupStats[key] && vehiclePickupStats[key].stats) {
                        const statsRow = document.createElement('tr');
                        statsRow.className = 'pickup-stats-row';
                        let statsHTML = '<td colspan="15" class="pickup-stats-cell">';
                        statsHTML += `<strong>上车→下车地点人员统计:</strong> `;
                        
                        const stats = vehiclePickupStats[key].stats;
                        for (let location in stats) {
                            statsHTML += `<span class="location-count">${location}: ${stats[location]}人</span>`;
                        }
                        
                        statsHTML += '</td>';
                        statsRow.innerHTML = statsHTML;
                        processedDataBody.appendChild(statsRow);
                    }
                }
                
                if (item.线路 !== currentLine) {
                    currentLine = item.线路;
                    
                    const lineSeparatorRow = document.createElement('tr');
                    lineSeparatorRow.className = 'line-group';
                    
                    const isUnassigned = !currentPlate || currentPlate === '未分配车辆';
                    
                    if (isUnassigned) {
                        const lineCheckboxId = `line-checkbox-${item.线路}-${currentPlate || 'unassigned'}`;
                        lineSeparatorRow.innerHTML = `<td colspan="15" style="background-color: #f8f8f8; padding: 6px;">
                            <div class="unassigned-line-label">
                                <input type="checkbox" class="unassigned-line-checkbox" id="${lineCheckboxId}" data-line="${item.线路}" data-plate="${currentPlate || ''}">
                                <label for="${lineCheckboxId}">线路: ${item.线路}</label>
                            </div>
                        </td>`;
                        
                        const lineCheckbox = lineSeparatorRow.querySelector('.unassigned-line-checkbox');
                        lineCheckbox.addEventListener('change', function() {
                            const line = this.getAttribute('data-line');
                            const plate = this.getAttribute('data-plate');
                            
                            const lineRows = processedDataBody.querySelectorAll(`tr[data-line="${line}"][data-plate="${plate}"]`);
                            
                            lineRows.forEach(row => {
                                const checkbox = row.querySelector('.row-checkbox');
                                const orderId = row.dataset.orderId;
                                
                                if (this.checked) {
                                    checkbox.checked = true;
                                    selectedRows.add(orderId);
                                    row.classList.add('selected-row');
                                } else {
                                    checkbox.checked = false;
                                    selectedRows.delete(orderId);
                                    row.classList.remove('selected-row');
                                    selectAllCheckbox.checked = false;
                                }
                            });
                            
                            updateModifyButtonState();
                        });
                    } else {
                        lineSeparatorRow.innerHTML = `<td colspan="15" style="background-color: #f8f8f8; padding: 6px; font-weight: bold;">
                            线路: ${item.线路}
                        </td>`;
                    }
                    processedDataBody.appendChild(lineSeparatorRow);
                }
                
                row.style.backgroundColor = lineColorMap.get(item.线路) || '#FFFFFF';
                
                row.innerHTML = `
                    <td style="text-align:center;">
                        <input type="checkbox" class="select-checkbox row-checkbox" data-order-id="${item.订单ID}" ${selectedRows.has(item.订单ID) ? 'checked' : ''}>
                    </td>
                    <td>${item.订单ID}</td>
                    <td><strong>${item.线路}</strong></td>
                    <td>${item.车牌号码 || ""}</td>
                    <td>${item.联系人}</td>
                    <td>${item.联系电话}</td>
                    <td class="ticket-count">${item.票数}</td>
                    <td>${item.乘车日期}</td>
                    <td>${item.乘车时间}</td>
                    <td>${item.上车地点}</td>
                    <td>${item.下车地点}</td>
                    <td>${item.随车负责人 || ""}</td>
                    <td>${item.随车负责人电话 || ""}</td>
                    <td>${item.通知内容}</td>
                    <td>${item.修改时间}</td>
                `;
                processedDataBody.appendChild(row);
            });
            
            document.querySelectorAll('.row-checkbox').forEach(checkbox => {
                checkbox.addEventListener('change', function() {
                    const orderId = this.getAttribute('data-order-id');
                    const row = this.closest('tr');
                    const line = row.dataset.line;
                    const plate = row.dataset.plate;
                    
                    if (this.checked) {
                        selectedRows.add(orderId);
                        row.classList.add('selected-row');
                    } else {
                        selectedRows.delete(orderId);
                        row.classList.remove('selected-row');
                        selectAllCheckbox.checked = false;
                        
                        const lineCheckbox = document.querySelector(`.unassigned-line-checkbox[data-line="${line}"][data-plate="${plate}"]`);
                        if (lineCheckbox) {
                            lineCheckbox.checked = false;
                        }
                        
                        const plateCheckbox = document.querySelector(`.plate-checkbox[data-plate="${plate}"]`);
                        if (plateCheckbox) {
                            plateCheckbox.checked = false;
                        }
                    }
                    
                    updateModifyButtonState();
                });
            });
            
            updateSelectAllButtons();
        }
        
        // 更新全选按钮的显示状态
        function updateSelectAllButtons() {
            const selectedPlates = new Set();
            const selectedLines = new Set();
            let allUnassigned = true;
            
            selectedRows.forEach(orderId => {
                const row = document.querySelector(`tr[data-order-id="${orderId}"]`);
                if (row) {
                    selectedPlates.add(row.dataset.plate);
                    selectedLines.add(row.dataset.line);
                    
                    if (row.dataset.plate) {
                        allUnassigned = false;
                    }
                }
            });
            
            if (selectedPlates.size === 1 && selectedRows.size > 0 && !allUnassigned) {
                const plate = Array.from(selectedPlates)[0];
                selectAllPlateBtn.style.display = 'inline-flex';
                selectAllPlateBtn.setAttribute('data-plate', plate);
                selectAllPlateBtn.innerHTML = `<i class="fas fa-check-square"></i> 全选${plate}`;
            } else {
                selectAllPlateBtn.style.display = 'none';
            }
            
            if (selectedLines.size === 1 && selectedRows.size > 0 && allUnassigned) {
                const line = Array.from(selectedLines)[0];
                selectAllUnassignedBtn.style.display = 'inline-flex';
                selectAllUnassignedBtn.setAttribute('data-line', line);
                selectAllUnassignedBtn.innerHTML = `<i class="fas fa-check-square"></i> 全选${line}未分配`;
            } else {
                selectAllUnassignedBtn.style.display = 'none';
            }
        }
        
        // 全选当前车牌
        function selectAllCurrentPlate() {
            const plate = selectAllPlateBtn.getAttribute('data-plate');
            if (!plate) return;
            
            const plateRows = processedDataBody.querySelectorAll(`tr[data-plate="${plate}"]`);
            const plateCheckbox = document.querySelector(`.plate-checkbox[data-plate="${plate}"]`);
            
            plateRows.forEach(row => {
                const checkbox = row.querySelector('.row-checkbox');
                const orderId = row.dataset.orderId;
                
                if (checkbox) {
                    checkbox.checked = true;
                    selectedRows.add(orderId);
                    row.classList.add('selected-row');
                }
            });
            
            if (plateCheckbox) {
                plateCheckbox.checked = true;
            }
            
            updateModifyButtonState();
        }
        
        // 全选当前线路的未分配订单
        function selectAllCurrentLineUnassigned() {
            const line = selectAllUnassignedBtn.getAttribute('data-line');
            if (!line) return;
            
            const unassignedRows = processedDataBody.querySelectorAll(`tr[data-line="${line}"][data-plate=""]`);
            
            unassignedRows.forEach(row => {
                const checkbox = row.querySelector('.row-checkbox');
                const orderId = row.dataset.orderId;
                
                if (checkbox) {
                    checkbox.checked = true;
                    selectedRows.add(orderId);
                    row.classList.add('selected-row');
                }
            });
            
            const lineCheckbox = document.querySelector(`.unassigned-line-checkbox[data-line="${line}"]`);
            if (lineCheckbox) {
                lineCheckbox.checked = true;
            }
            
            updateModifyButtonState();
        }
        
        // 显示各车次分组地点统计面板
        function displayVehiclePickupStatsPanel() {
            if (!processedData.length) {
                vehiclePickupStatsPanel.style.display = 'none';
                return;
            }
            
            vehiclePickupStats = {};
            processedData.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                const line = item.线路 || '未知线路';
                const pickup = item.上车地点 || '未知上车地点';
                const dropoff = item.下车地点 || '未知下车地点';
                const tickets = item.票数 || 0;
                
                const key = `${plate}|${line}`;
                
                if (!vehiclePickupStats[key]) {
                    vehiclePickupStats[key] = {
                        plate: plate,
                        line: line,
                        stats: {}
                    };
                }
                
                const locationKey = `${pickup}→${dropoff}`;
                if (!vehiclePickupStats[key].stats[locationKey]) {
                    vehiclePickupStats[key].stats[locationKey] = 0;
                }
                
                vehiclePickupStats[key].stats[locationKey] += tickets;
            });
            
            let statsHTML = `
                <div class="stats-header">
                    <div class="vehicle-pickup-stats-title">
                        <i class="fas fa-bus"></i> 各车次上车→下车地点人数统计
                    </div>
                    <button class="stats-refresh-btn" id="refreshStatsBtn">
                        <i class="fas fa-redo"></i> 刷新统计
                    </button>
                </div>
                
                <div class="stats-search-container">
                    <div class="stats-search-dropdown">
                        <button class="stats-search-dropdown-btn" id="statsSearchDropdownBtn">
                            <span id="statsSearchTypeText">车牌号</span>
                            <i class="fas fa-chevron-down"></i>
                        </button>
                        <div class="stats-search-dropdown-content" id="statsSearchDropdownContent">
                            <div class="stats-search-dropdown-item" data-type="车牌号">车牌号</div>
                            <div class="stats-search-dropdown-item" data-type="线路">线路</div>
                        </div>
                    </div>
                    <input type="text" class="stats-search-input" id="statsSearchInput" placeholder="请输入搜索内容...">
                </div>
            `;
            
            const plateGroups = {};
            
            for (let key in vehiclePickupStats) {
                const plate = vehiclePickupStats[key].plate;
                if (!plateGroups[plate]) {
                    plateGroups[plate] = [];
                }
                plateGroups[plate].push(vehiclePickupStats[key]);
            }
            
            // 按照车辆顺序字段排序
            Object.keys(plateGroups).sort((a, b) => {
                // 未分配车辆始终排在最后
                if (a === '未分配车辆' || a === '') return 1;
                if (b === '未分配车辆' || b === '') return -1;
                
                // 获取车辆顺序
                const orderA = vehicleCapacityMap.has(a) ? (vehicleCapacityMap.get(a).顺序 || 9999) : 9999;
                const orderB = vehicleCapacityMap.has(b) ? (vehicleCapacityMap.get(b).顺序 || 9999) : 9999;
                
                // 按照顺序排序
                if (orderA !== orderB) {
                    return orderA - orderB;
                }
                
                // 顺序相同，按车牌号排序
                return a.localeCompare(b);
            }).forEach(plate => {
                const vehicles = plateGroups[plate];
                const isUnassigned = plate === '未分配车辆' || plate === '';
                
                let plateTotal = 0;
                vehicles.forEach(vehicle => {
                    for (let location in vehicle.stats) {
                        plateTotal += vehicle.stats[location];
                    }
                });
                
                let capacityInfo = '';
                if (plate && vehicleCapacityMap.has(plate)) {
                    const vehicleInfo = vehicleCapacityMap.get(plate);
                    const capacity = vehicleInfo.容量;
                    const vehicleTime = vehicleInfo.乘车时间 || '';
                    const timeInfo = vehicleTime ? `时间:${vehicleTime} ` : '';
                    
                    // 使用车辆总分配人数（跨线路）
                    const totalAssigned = vehicleTotalAssigned.get(plate) || plateTotal;
                    
                    if (capacity !== null) {
                        if (totalAssigned > capacity) {
                            const overloaded = totalAssigned - capacity;
                            capacityInfo = `<div style="margin-top:5px; font-size:0.85rem;"><span style="color: #e74c3c; font-weight: 700;">超载(${capacity}座, 已分配${totalAssigned}人, 超载${overloaded}人)</span></div>`;
                        } else if (totalAssigned === capacity) {
                            capacityInfo = `<div style="margin-top:5px; font-size:0.85rem;"><span class="full-status">满员(${capacity}座, 已分配${totalAssigned}人)</span></div>`;
                        } else {
                            const emptySeats = capacity - totalAssigned;
                            capacityInfo = `<div style="margin-top:5px; font-size:0.85rem;"><span class="available-status">${timeInfo}${totalAssigned}/${capacity}人(空${emptySeats}座)</span></div>`;
                        }
                    } else {
                        capacityInfo = `<div style="margin-top:5px; font-size:0.85rem;"><span class="warning-status">${timeInfo}${totalAssigned}人(未限制)</span></div>`;
                    }
                } else if (isUnassigned) {
                    capacityInfo = `<div style="margin-top:5px; font-size:0.85rem;"><span class="unassigned-label">未分配车辆</span></div>`;
                }
                
                let lineText = '';
                if (plate && plateLineMap.has(plate)) {
                    const lines = Array.from(plateLineMap.get(plate));
                    if (lines.length === 1) {
                        lineText = `<span class="line-in-plate">${lines[0]}</span> - `;
                    } else if (lines.length > 1) {
                        lineText = `<span class="line-in-plate">${lines.join(',')}</span> - `;
                    }
                } else if (plate && plateStats[plate] && plateStats[plate].lines.size > 0) {
                    const lines = Array.from(plateStats[plate].lines);
                    if (lines.length === 1) {
                        lineText = `<span class="line-in-plate">${lines[0]}</span> - `;
                    } else if (lines.length > 1) {
                        lineText = `<span class="line-in-plate">${lines.join(',')}</span> - `;
                    }
                }
                
                const plateTitle = plate ? `${lineText}车牌号码: ${plate}` : '未分配车辆';
                statsHTML += `
                    <div class="line-group-title" data-plate="${plate}" data-line="${vehicles.map(v => v.line).join(',')}">
                        <i class="fas fa-car"></i> ${plateTitle} ${capacityInfo}
                        ${plate && plate !== '未分配车辆' ? `<button class="edit-vehicle-btn" data-plate="${plate}" style="margin-left: 10px; padding: 4px 12px; background-color: #3498db; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 0.85rem;"><i class="fas fa-edit"></i> 编辑</button>` : ''}
                    </div>
                    <div class="vehicle-pickup-stats-grid">
                `;
                
                vehicles.forEach(vehicle => {
                    const plateName = vehicle.plate || '未分配车辆';
                    const lineName = vehicle.line;
                    
                    let totalCount = 0;
                    for (let location in vehicle.stats) {
                        totalCount += vehicle.stats[location];
                    }
                    
                    const isUnassigned = plateName === '未分配车辆' || plateName === '';
                    
                    const isManuallyAssigned = plateName && manuallyAssignedOrders.size > 0 && 
                        Array.from(manuallyAssignedOrders).some(orderId => {
                            const order = processedData.find(item => item.订单ID === orderId);
                            return order && order.车牌号码 === plateName && order.线路 === lineName;
                        });
                    
                    const manuallyAssignedClass = isManuallyAssigned ? 'manually-assigned-box' : '';
                    const manuallyAssignedText = isManuallyAssigned ? `<div class="manual-assignment-text">${totalCount}人乘坐${plateName}</div>` : '';
                    
                    const locationScrollContent = Object.keys(vehicle.stats).map(location => {
                        return `<div class="pickup-stat-item">
                            <span class="pickup-location">${location}</span>
                            <span class="pickup-count ${isUnassigned ? 'unassigned-count' : ''}">${vehicle.stats[location]}人</span>
                        </div>`;
                    }).join('');
                    
                    statsHTML += `
                        <div class="vehicle-stat-item ${isUnassigned ? 'unassigned-stat-item' : ''} ${manuallyAssignedClass}" style="${isManuallyAssigned ? 'border: 2px solid #e74c3c !important; box-shadow: 0 0 5px rgba(231, 76, 60, 0.5);' : ''}" data-plate="${plateName}" data-line="${lineName}">
                            <div class="vehicle-title ${isUnassigned ? 'unassigned-title' : ''} ${isManuallyAssigned ? 'manually-assigned-label' : ''}">
                                <span class="vehicle-plate">${lineName}</span>
                            </div>
                            <div class="vehicle-pickup-stats">
                                <div class="location-scroll">
                                    ${locationScrollContent}
                                </div>
                            </div>
                            <div class="vehicle-total">
                                <span class="vehicle-total-label ${isUnassigned ? 'unassigned-label' : ''}">总人数</span>
                                <span class="vehicle-total-count ${isUnassigned ? 'unassigned-count' : ''}">${totalCount}人</span>
                            </div>
                            ${manuallyAssignedText}
                        </div>
                    `;
                });
                
                statsHTML += '</div>';
            });
            
            vehiclePickupStatsPanel.innerHTML = statsHTML;
            vehiclePickupStatsPanel.style.display = 'block';
            
            // 设置车辆统计搜索功能
            setupStatsSearch();
            
            // 设置刷新按钮事件
            const refreshStatsBtn = document.getElementById('refreshStatsBtn');
            if (refreshStatsBtn) {
                refreshStatsBtn.addEventListener('click', function() {
                    // 重新计算车辆统计
                    displayVehiclePickupStatsPanel();
                    showNotification('车辆统计已刷新', 'success');
                });
            }
            
            // 设置编辑车辆按钮事件
            const editVehicleBtns = document.querySelectorAll('.edit-vehicle-btn');
            editVehicleBtns.forEach(btn => {
                btn.addEventListener('click', function() {
                    const plate = this.getAttribute('data-plate');
                    showEditVehicleModal(plate);
                });
            });
        }
        
        // 设置车辆统计搜索功能
        function setupStatsSearch() {
            const statsSearchDropdownBtn = document.getElementById('statsSearchDropdownBtn');
            const statsSearchDropdownContent = document.getElementById('statsSearchDropdownContent');
            const statsSearchTypeText = document.getElementById('statsSearchTypeText');
            const statsSearchInput = document.getElementById('statsSearchInput');
            
            if (!statsSearchDropdownBtn) return;
            
            let statsSearchType = '车牌号';
            
            // 设置下拉菜单
            document.querySelectorAll('.stats-search-dropdown-item').forEach(item => {
                item.addEventListener('click', function() {
                    statsSearchType = this.getAttribute('data-type');
                    statsSearchTypeText.textContent = statsSearchType;
                    statsSearchDropdownContent.classList.remove('show');
                    
                    // 触发搜索
                    filterStatsItems();
                });
            });
            
            statsSearchDropdownBtn.addEventListener('click', function() {
                statsSearchDropdownContent.classList.toggle('show');
            });
            
            // 点击外部关闭下拉菜单
            document.addEventListener('click', function(e) {
                if (!statsSearchDropdownBtn.contains(e.target) && !statsSearchDropdownContent.contains(e.target)) {
                    statsSearchDropdownContent.classList.remove('show');
                }
            });
            
            // 输入框搜索
            statsSearchInput.addEventListener('input', filterStatsItems);
            
            // 过滤车辆统计项
            function filterStatsItems() {
                const searchTerm = document.getElementById('statsSearchInput').value.trim().toLowerCase();
                const searchType = statsSearchType;
                
                if (!searchTerm) {
                    // 显示所有
                    document.querySelectorAll('.vehicle-stat-item, .line-group-title').forEach(item => {
                        item.classList.remove('filtered-out');
                    });
                    
                    // 移除之前的无结果提示
                    const existingNoResults = vehiclePickupStatsPanel.querySelector('.no-results');
                    if (existingNoResults) {
                        existingNoResults.remove();
                    }
                    return;
                }
                
                let hasResults = false;
                
                // 隐藏所有车辆统计项
                document.querySelectorAll('.vehicle-stat-item, .line-group-title').forEach(item => {
                    item.classList.add('filtered-out');
                });
                
                // 根据搜索类型显示匹配的项
                if (searchType === '车牌号') {
                    document.querySelectorAll('.line-group-title').forEach(title => {
                        const plate = title.getAttribute('data-plate');
                        if (plate && plate.toLowerCase().includes(searchTerm)) {
                            title.classList.remove('filtered-out');
                            
                            // 显示该车牌下的所有车辆项
                            const plateName = title.getAttribute('data-plate');
                            document.querySelectorAll(`.vehicle-stat-item[data-plate="${plateName}"]`).forEach(item => {
                                item.classList.remove('filtered-out');
                            });
                            
                            hasResults = true;
                        }
                    });
                } else if (searchType === '线路') {
                    document.querySelectorAll('.vehicle-stat-item').forEach(item => {
                        const line = item.getAttribute('data-line');
                        if (line && line.toLowerCase().includes(searchTerm)) {
                            item.classList.remove('filtered-out');
                            
                            // 显示对应的车牌标题
                            const plate = item.getAttribute('data-plate');
                            document.querySelectorAll(`.line-group-title[data-plate="${plate}"]`).forEach(title => {
                                title.classList.remove('filtered-out');
                            });
                            
                            hasResults = true;
                        }
                    });
                }
                
                // 如果没有结果，显示提示
                const existingNoResults = vehiclePickupStatsPanel.querySelector('.no-results');
                if (!hasResults) {
                    if (!existingNoResults) {
                        const noResults = document.createElement('div');
                        noResults.className = 'no-results';
                        noResults.textContent = `未找到匹配"${searchTerm}"的${searchType}`;
                        vehiclePickupStatsPanel.appendChild(noResults);
                    }
                } else if (existingNoResults) {
                    existingNoResults.remove();
                }
            }
        }
        
        // 显示线路票数统计
        function displayLineTicketStats() {
            const lineTicketStats = {};
            
            processedData.forEach(item => {
                const line = item.线路;
                if (!lineTicketStats[line]) {
                    lineTicketStats[line] = 0;
                }
                lineTicketStats[line] += (item.票数 || 0);
            });
            
            const totalTickets = Object.values(lineTicketStats).reduce((sum, tickets) => sum + tickets, 0);
            
            let lineStatsHTML = `
                <div class="line-stats-title">
                    <i class="fas fa-users"></i> 各线路票数统计
                    <button id="toggleLineStatsBtn" class="toggle-line-stats-btn">
                        <i class="fas fa-eye-slash"></i> 隐藏
                    </button>
                </div>
                <div class="line-stats-content" id="lineStatsContent">
                    <div class="line-stats-grid">
            `;
            
            const lines = lineOrder.length > 0 ? lineOrder : [...new Set(processedData.map(item => item.线路))].sort();
            
            lines.forEach(line => {
                const tickets = lineTicketStats[line] || 0;
                
                const vehicles = lineInfo[line] || [];
                let vehicleInfo = '';
                if (vehicles.length > 0) {
                    vehicleInfo = `<br><span style="font-size:0.95rem;">${vehicles.length}辆车`;
                    vehicles.forEach((vehicle, index) => {
                        const capacity = vehicle.人数 === null ? '未限制' : `${vehicle.人数}人限制`;
                        const timeInfo = vehicle.乘车时间 ? `时间:${vehicle.乘车时间}, ` : '';
                        const pickupInfo = vehicle.上车地点 ? `上:${vehicle.上车地点}, ` : '';
                        const dropoffInfo = vehicle.下车地点 ? `下:${vehicle.下车地点}, ` : '';
                        const assigned = vehicle.已分配人数 || 0;
                        
                        let statusText = "";
                        if (vehicle.人数 !== null) {
                            if (assigned >= vehicle.人数) {
                                statusText = `车${index+1}:满员(${vehicle.人数}座)`;
                            } else {
                                const emptySeats = vehicle.人数 - assigned;
                                statusText = `车${index+1}:${timeInfo}${pickupInfo}${dropoffInfo}${assigned}/${vehicle.人数}人(空${emptySeats}座)`;
                            }
                        } else {
                            statusText = `车${index+1}:${timeInfo}${pickupInfo}${dropoffInfo}${assigned}人(未限制)`;
                        }
                        
                        vehicleInfo += `，${statusText}`;
                    });
                    vehicleInfo += '</span>';
                }
                
                const unassignedCount = lineUnassignedCount[line] || 0;
                
                lineStatsHTML += `
                    <div class="line-stat-item" data-line="${line}">
                        <div class="line-name">${line}</div>
                        <div class="line-count">${tickets} 人</div>
                        <div class="line-label">${vehicleInfo}</div>
                    </div>
                `;
            });
            
            lineStatsHTML += `
                    </div>
                    <div class="total-tickets">
                        <div class="total-tickets-value">${totalTickets} 人</div>
                        <div class="total-tickets-label">总票数</div>
                    </div>
                </div>
            `;
            
            lineStatsPanel.innerHTML = lineStatsHTML;
            
            // 添加显示/隐藏按钮事件监听器
            const toggleLineStatsBtn = document.getElementById('toggleLineStatsBtn');
            const lineStatsContent = document.getElementById('lineStatsContent');
            
            if (toggleLineStatsBtn && lineStatsContent) {
                toggleLineStatsBtn.addEventListener('click', function() {
                    lineStatsVisible = !lineStatsVisible;
                    
                    if (lineStatsVisible) {
                        lineStatsContent.style.display = 'block';
                        toggleLineStatsBtn.innerHTML = '<i class="fas fa-eye-slash"></i> 隐藏';
                        toggleLineStatsBtn.style.backgroundColor = '#3498db';
                    } else {
                        lineStatsContent.style.display = 'none';
                        toggleLineStatsBtn.innerHTML = '<i class="fas fa-eye"></i> 显示';
                        toggleLineStatsBtn.style.backgroundColor = '#2ecc71';
                    }
                });
            }
        }
        
        // 新增：显示线路时间地点统计
        function displayLineTimeLocationStats() {
            // 清空现有数据
            lineTimeLocationData = [];
            lineTimeLocationBody.innerHTML = '';
            
            if (processedData.length === 0) {
                lineTimeLocationPanel.style.display = 'none';
                return;
            }
            
            // 按线路、时间、上车地点、下车地点分组统计
            const statsMap = {};
            
            processedData.forEach(item => {
                const line = item.线路 || '未知线路';
                const time = item.乘车时间 || '未指定时间';
                const pickup = item.上车地点 || '未指定上车地点';
                const dropoff = item.下车地点 || '未指定下车地点';
                const tickets = item.票数 || 0;
                
                const key = `${line}|${time}|${pickup}|${dropoff}`;
                
                if (!statsMap[key]) {
                    statsMap[key] = {
                        线路: line,
                        时间: time,
                        上车地点: pickup,
                        下车地点: dropoff,
                        票数: 0
                    };
                }
                
                statsMap[key].票数 += tickets;
            });
            
            // 转换为数组并按线路、时间排序
            lineTimeLocationData = Object.values(statsMap).sort((a, b) => {
                if (a.线路 !== b.线路) return a.线路.localeCompare(b.线路);
                if (a.时间 !== b.时间) return a.时间.localeCompare(b.时间);
                if (a.上车地点 !== b.上车地点) return a.上车地点.localeCompare(b.上车地点);
                return a.下车地点.localeCompare(b.下车地点);
            });
            
            // 填充表格
            if (lineTimeLocationData.length === 0) {
                lineTimeLocationBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">暂无统计数据</td></tr>';
            } else {
                lineTimeLocationData.forEach(item => {
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${item.线路}</td>
                        <td>${item.时间}</td>
                        <td>${item.上车地点}</td>
                        <td>${item.下车地点}</td>
                        <td style="font-weight:bold; color:#27ae60;">${item.票数} 人</td>
                    `;
                    lineTimeLocationBody.appendChild(row);
                });
            }
            
            // 显示面板
            lineTimeLocationPanel.style.display = 'flex';
            
            // 确保表格容器的初始高度设置正确
            const headerHeight = lineTimeLocationPanel.querySelector('.line-time-location-header').offsetHeight;
            const searchHeight = lineTimeLocationPanel.querySelector('.line-time-location-search').offsetHeight;
            const panelHeight = parseInt(getComputedStyle(lineTimeLocationPanel).height, 10);
            const tableContainerHeight = panelHeight - headerHeight - searchHeight - 40;
            lineTimeLocationTableContainer.style.maxHeight = Math.max(100, tableContainerHeight) + 'px';
            
            // 清空搜索框
            lineTimeLocationSearch.value = '';
            
            // 显示按上车地点和下车地点统计的人员数量
            displayLocationStats();
        }
        
        // 新增：显示按车牌号分组的上车和下车地点人员统计
        function displayLocationStats() {
            // 按车牌号分组统计
            const plateStatsMap = {};
            
            processedData.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                const pickup = item.上车地点 || '未指定上车地点';
                const dropoff = item.下车地点 || '未指定下车地点';
                const tickets = item.票数 || 0;
                
                if (plate && plate !== '未分配车辆') {
                    if (!plateStatsMap[plate]) {
                        plateStatsMap[plate] = {
                            上车地点: {},
                            下车地点: {}
                        };
                    }
                    
                    // 统计上车地点
                    if (!plateStatsMap[plate].上车地点[pickup]) {
                        plateStatsMap[plate].上车地点[pickup] = 0;
                    }
                    plateStatsMap[plate].上车地点[pickup] += tickets;
                    
                    // 统计下车地点
                    if (!plateStatsMap[plate].下车地点[dropoff]) {
                        plateStatsMap[plate].下车地点[dropoff] = 0;
                    }
                    plateStatsMap[plate].下车地点[dropoff] += tickets;
                }
            });
            
            // 转换为数组并排序
            const plateStatsData = Object.entries(plateStatsMap).map(([plate, stats]) => {
                // 上车地点数据
                const pickupData = Object.entries(stats.上车地点)
                    .map(([location, count]) => ({ 类型: '上车地点', 地点: location, 人数: count }))
                    .sort((a, b) => b.人数 - a.人数);
                
                // 下车地点数据
                const dropoffData = Object.entries(stats.下车地点)
                    .map(([location, count]) => ({ 类型: '下车地点', 地点: location, 人数: count }))
                    .sort((a, b) => b.人数 - a.人数);
                
                return {
                    车牌号: plate,
                    上车地点数据: pickupData,
                    下车地点数据: dropoffData
                };
            }).sort((a, b) => a.车牌号.localeCompare(b.车牌号));
            
            // 检查是否已存在地点统计面板
            let locationStatsPanel = document.getElementById('locationStatsPanel');
            if (!locationStatsPanel) {
                // 创建新的面板
                locationStatsPanel = document.createElement('div');
                locationStatsPanel.id = 'locationStatsPanel';
                locationStatsPanel.className = 'line-time-location-panel';
                
                locationStatsPanel.innerHTML = `
                    <div class="line-time-location-header">
                        <div class="line-time-location-title">
                            <i class="fas fa-car"></i> 各车牌号上车下车地点人员统计
                        </div>
                        <div class="line-time-location-controls">
                            <button class="resize-handle" title="拖拽调整高度">
                                <i class="fas fa-arrows-alt-v"></i>
                            </button>
                        </div>
                    </div>
                    <div class="line-time-location-search">
                        <input type="text" class="line-time-location-search-input" id="locationStatsSearch" 
                               placeholder="输入车牌号或地点名称进行筛选...">
                    </div>
                    <div class="line-time-location-table-container" id="locationStatsTableContainer">
                        <table class="line-time-location-table" id="locationStatsTable">
                            <thead>
                                <tr>
                                    <th>车牌号</th>
                                    <th>统计类型</th>
                                    <th>地点</th>
                                    <th>人数</th>
                                </tr>
                            </thead>
                            <tbody id="locationStatsBody"></tbody>
                        </table>
                    </div>
                `;
                
                // 将面板添加到人员统计板块
                const personStatsSection = document.getElementById('person-stats-section');
                if (personStatsSection) {
                    personStatsSection.appendChild(locationStatsPanel);
                }
            }
            
            // 填充表格
            const locationStatsBody = document.getElementById('locationStatsBody');
            if (locationStatsBody) {
                locationStatsBody.innerHTML = '';
                
                let hasData = false;
                
                // 填充数据
                plateStatsData.forEach(plateStats => {
                    const plate = plateStats.车牌号;
                    const totalRows = plateStats.上车地点数据.length + plateStats.下车地点数据.length;
                    
                    if (totalRows > 0) {
                        hasData = true;
                        
                        // 添加上车地点统计
                        plateStats.上车地点数据.forEach((item, index) => {
                            const row = document.createElement('tr');
                            if (index === 0) {
                                // 第一个上车地点，显示车牌号并设置rowspan
                                row.innerHTML = `
                                    <td rowspan="${totalRows}" style="vertical-align: top; font-weight:bold;">${plate}</td>
                                    <td style="font-weight:bold; color:#3498db;">上车地点</td>
                                    <td>${item.地点}</td>
                                    <td style="font-weight:bold; color:#27ae60;">${item.人数} 人</td>
                                `;
                            } else {
                                // 后续上车地点，不显示车牌号
                                row.innerHTML = `
                                    <td style="font-weight:bold; color:#3498db;">上车地点</td>
                                    <td>${item.地点}</td>
                                    <td style="font-weight:bold; color:#27ae60;">${item.人数} 人</td>
                                `;
                            }
                            locationStatsBody.appendChild(row);
                        });
                        
                        // 添加下车地点统计
                        plateStats.下车地点数据.forEach((item, index) => {
                            const row = document.createElement('tr');
                            row.innerHTML = `
                                <td style="font-weight:bold; color:#e74c3c;">下车地点</td>
                                <td>${item.地点}</td>
                                <td style="font-weight:bold; color:#27ae60;">${item.人数} 人</td>
                            `;
                            locationStatsBody.appendChild(row);
                        });
                    }
                });
                
                if (!hasData) {
                    locationStatsBody.innerHTML = '<tr><td colspan="4" style="text-align:center;">暂无统计数据</td></tr>';
                }
            }
            
            // 显示面板
            if (locationStatsPanel) {
                locationStatsPanel.style.display = 'flex';
                
                // 确保表格容器的初始高度设置正确
                const headerHeight = locationStatsPanel.querySelector('.line-time-location-header').offsetHeight;
                const searchHeight = locationStatsPanel.querySelector('.line-time-location-search').offsetHeight;
                const panelHeight = parseInt(getComputedStyle(locationStatsPanel).height, 10);
                const tableContainerHeight = panelHeight - headerHeight - searchHeight - 40;
                const locationStatsTableContainer = document.getElementById('locationStatsTableContainer');
                if (locationStatsTableContainer) {
                    locationStatsTableContainer.style.maxHeight = Math.max(100, tableContainerHeight) + 'px';
                }
                
                // 清空搜索框
                const locationStatsSearch = document.getElementById('locationStatsSearch');
                if (locationStatsSearch) {
                    locationStatsSearch.value = '';
                    
                    // 添加搜索事件
                    locationStatsSearch.addEventListener('input', filterLocationStatsData);
                }
                
                // 添加调整大小事件
                const resizeHandle = locationStatsPanel.querySelector('.resize-handle');
                if (resizeHandle) {
                    resizeHandle.addEventListener('mousedown', startResize);
                }
            }
        }
        
        // 新增：过滤地点统计数据
        function filterLocationStatsData() {
            const searchTerm = document.getElementById('locationStatsSearch').value.trim().toLowerCase();
            const locationStatsBody = document.getElementById('locationStatsBody');
            
            if (!locationStatsBody) return;
            
            // 按车牌号分组统计
            const plateStatsMap = {};
            
            processedData.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                const pickup = item.上车地点 || '未指定上车地点';
                const dropoff = item.下车地点 || '未指定下车地点';
                const tickets = item.票数 || 0;
                
                if (plate && plate !== '未分配车辆') {
                    if (!plateStatsMap[plate]) {
                        plateStatsMap[plate] = {
                            上车地点: {},
                            下车地点: {}
                        };
                    }
                    
                    // 统计上车地点
                    if (!plateStatsMap[plate].上车地点[pickup]) {
                        plateStatsMap[plate].上车地点[pickup] = 0;
                    }
                    plateStatsMap[plate].上车地点[pickup] += tickets;
                    
                    // 统计下车地点
                    if (!plateStatsMap[plate].下车地点[dropoff]) {
                        plateStatsMap[plate].下车地点[dropoff] = 0;
                    }
                    plateStatsMap[plate].下车地点[dropoff] += tickets;
                }
            });
            
            // 转换为数组并排序
            let plateStatsData = Object.entries(plateStatsMap).map(([plate, stats]) => {
                // 上车地点数据
                const pickupData = Object.entries(stats.上车地点)
                    .map(([location, count]) => ({ 类型: '上车地点', 地点: location, 人数: count }))
                    .sort((a, b) => b.人数 - a.人数);
                
                // 下车地点数据
                const dropoffData = Object.entries(stats.下车地点)
                    .map(([location, count]) => ({ 类型: '下车地点', 地点: location, 人数: count }))
                    .sort((a, b) => b.人数 - a.人数);
                
                return {
                    车牌号: plate,
                    上车地点数据: pickupData,
                    下车地点数据: dropoffData
                };
            }).sort((a, b) => a.车牌号.localeCompare(b.车牌号));
            
            // 过滤数据
            if (searchTerm) {
                plateStatsData = plateStatsData.filter(plateStats => {
                    // 检查车牌号是否匹配
                    if (plateStats.车牌号.toLowerCase().includes(searchTerm)) {
                        return true;
                    }
                    
                    // 检查上车地点是否匹配
                    const hasMatchingPickup = plateStats.上车地点数据.some(item => 
                        item.地点.toLowerCase().includes(searchTerm)
                    );
                    if (hasMatchingPickup) {
                        return true;
                    }
                    
                    // 检查下车地点是否匹配
                    const hasMatchingDropoff = plateStats.下车地点数据.some(item => 
                        item.地点.toLowerCase().includes(searchTerm)
                    );
                    return hasMatchingDropoff;
                });
                
                // 过滤每个车牌号的数据
                plateStatsData = plateStatsData.map(plateStats => {
                    return {
                        车牌号: plateStats.车牌号,
                        上车地点数据: plateStats.上车地点数据.filter(item => 
                            item.地点.toLowerCase().includes(searchTerm)
                        ),
                        下车地点数据: plateStats.下车地点数据.filter(item => 
                            item.地点.toLowerCase().includes(searchTerm)
                        )
                    };
                }).filter(plateStats => 
                    plateStats.上车地点数据.length > 0 || plateStats.下车地点数据.length > 0
                );
            }
            
            // 显示过滤后的数据
            locationStatsBody.innerHTML = '';
            
            let hasData = false;
            
            // 填充数据
            plateStatsData.forEach(plateStats => {
                const plate = plateStats.车牌号;
                const totalRows = plateStats.上车地点数据.length + plateStats.下车地点数据.length;
                
                if (totalRows > 0) {
                    hasData = true;
                    
                    // 添加上车地点统计
                    plateStats.上车地点数据.forEach((item, index) => {
                        const row = document.createElement('tr');
                        if (index === 0) {
                            // 第一个上车地点，显示车牌号并设置rowspan
                            row.innerHTML = `
                                <td rowspan="${totalRows}" style="vertical-align: top; font-weight:bold;">${plate}</td>
                                <td style="font-weight:bold; color:#3498db;">上车地点</td>
                                <td>${item.地点}</td>
                                <td style="font-weight:bold; color:#27ae60;">${item.人数} 人</td>
                            `;
                        } else {
                            // 后续上车地点，不显示车牌号
                            row.innerHTML = `
                                <td style="font-weight:bold; color:#3498db;">上车地点</td>
                                <td>${item.地点}</td>
                                <td style="font-weight:bold; color:#27ae60;">${item.人数} 人</td>
                            `;
                        }
                        locationStatsBody.appendChild(row);
                    });
                    
                    // 添加下车地点统计
                    plateStats.下车地点数据.forEach((item, index) => {
                        const row = document.createElement('tr');
                        row.innerHTML = `
                            <td style="font-weight:bold; color:#e74c3c;">下车地点</td>
                            <td>${item.地点}</td>
                            <td style="font-weight:bold; color:#27ae60;">${item.人数} 人</td>
                        `;
                        locationStatsBody.appendChild(row);
                    });
                }
            });
            
            if (!hasData) {
                locationStatsBody.innerHTML = '<tr><td colspan="4" style="text-align:center;">未找到匹配数据</td></tr>';
            }
        }
        
        // 新增：过滤线路时间地点数据
        function filterLineTimeLocationData() {
            const searchTerm = lineTimeLocationSearch.value.trim().toLowerCase();
            
            if (!searchTerm) {
                // 显示所有数据
                lineTimeLocationBody.innerHTML = '';
                lineTimeLocationData.forEach(item => {
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${item.线路}</td>
                        <td>${item.时间}</td>
                        <td>${item.上车地点}</td>
                        <td>${item.下车地点}</td>
                        <td style="font-weight:bold; color:#27ae60;">${item.票数} 人</td>
                    `;
                    lineTimeLocationBody.appendChild(row);
                });
                return;
            }
            
            // 过滤数据
            const filteredData = lineTimeLocationData.filter(item => 
                item.线路.toLowerCase().includes(searchTerm)
            );
            
            // 显示过滤后的数据
            lineTimeLocationBody.innerHTML = '';
            
            if (filteredData.length === 0) {
                lineTimeLocationBody.innerHTML = '<tr><td colspan="5" style="text-align:center;">未找到匹配线路的数据</td></tr>';
            } else {
                filteredData.forEach(item => {
                    const row = document.createElement('tr');
                    row.innerHTML = `
                        <td>${item.线路}</td>
                        <td>${item.时间}</td>
                        <td>${item.上车地点}</td>
                        <td>${item.下车地点}</td>
                        <td style="font-weight:bold; color:#27ae60;">${item.票数} 人</td>
                    `;
                    lineTimeLocationBody.appendChild(row);
                });
            }
        }
        
        // 新增：底部拖拽区域元素
        const bottomResizeArea = document.getElementById('bottomResizeArea');

        // 新增：为底部拖拽区域添加事件监听
        if (bottomResizeArea) {
            bottomResizeArea.addEventListener('mousedown', startResize);
        }

        // 新增：开始调整大小
        function startResize(e) {
            isResizing = true;
            startY = e.clientY;
            startHeight = parseInt(getComputedStyle(lineTimeLocationPanel).height, 10);
            
            document.addEventListener('mousemove', resize);
            document.addEventListener('mouseup', stopResize);
            
            // 阻止文本选择
            e.preventDefault();
        }
        
        // 新增：调整大小
        function resize(e) {
            if (!isResizing) return;
            
            const deltaY = e.clientY - startY;
            let newHeight = startHeight + deltaY;
            
            // 限制最小和最大高度
            const minHeight = 200;
            const maxHeight = window.innerHeight * 0.8; // 最大为视口高度的80%
            
            newHeight = Math.max(minHeight, Math.min(newHeight, maxHeight));
            
            lineTimeLocationPanel.style.height = newHeight + 'px';
            
            // 更新表格容器的高度
            const headerHeight = lineTimeLocationPanel.querySelector('.line-time-location-header').offsetHeight;
            const searchHeight = lineTimeLocationPanel.querySelector('.line-time-location-search').offsetHeight;
            const tableContainerHeight = newHeight - headerHeight - searchHeight - 40; // 减去内边距
            lineTimeLocationTableContainer.style.maxHeight = Math.max(100, tableContainerHeight) + 'px';
        }
        
        // 新增：停止调整大小
        function stopResize() {
            isResizing = false;
            document.removeEventListener('mousemove', resize);
            document.removeEventListener('mouseup', stopResize);
        }
        
        // 新增：显示导出排序选择对话框
        function showExportSortDialog() {
            if (processedData.length === 0) {
                showNotification('没有可导出的数据', 'error');
                return;
            }
            
            // 设置默认选项
            document.getElementById('sortByLine').checked = exportSortType === 'line';
            document.getElementById('sortByCampus').checked = exportSortType === 'campus';
            
            exportSortDialogOverlay.style.display = 'flex';
        }
        
        // 新增：关闭导出排序对话框
        function closeExportSortDialog() {
            exportSortDialogOverlay.style.display = 'none';
        }
        
        // 新增：确认导出排序方式
        function confirmExportSort() {
            const selectedSort = document.querySelector('input[name="exportSort"]:checked');
            if (selectedSort) {
                exportSortType = selectedSort.value;
                closeExportSortDialog();
                exportData();
            } else {
                showNotification('请选择排序方式', 'error');
            }
        }
        
        // 导出数据为Excel文件
        function checkInconsistencies() {
            const inconsistencies = [];
            
            // Group data by vehicle and line
            const vehicleLineGroups = {};
            processedData.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                const line = item.线路 || '';
                const key = `${plate}-${line}`;
                
                if (!vehicleLineGroups[key]) {
                    vehicleLineGroups[key] = [];
                }
                vehicleLineGroups[key].push(item);
            });
            
            // Check each group for inconsistencies
            Object.keys(vehicleLineGroups).forEach(key => {
                const [plate, line] = key.split('-');
                const group = vehicleLineGroups[key];
                
                if (group.length > 1) {
                    // Get the first item as reference
                    const reference = group[0];
                    const referenceEscort = reference.随车负责人 || '';
                    const referenceEscortPhone = reference.随车负责人电话 || '';
                    const referenceNotification = reference.通知内容 || '';
                    
                    // Check other items in the group
                    const inconsistentItems = [];
                    group.forEach((item, index) => {
                        if (index === 0) return; // Skip the reference item
                        
                        const currentEscort = item.随车负责人 || '';
                        const currentEscortPhone = item.随车负责人电话 || '';
                        const currentNotification = item.通知内容 || '';
                        
                        if (currentEscort !== referenceEscort || 
                            currentEscortPhone !== referenceEscortPhone || 
                            currentNotification !== referenceNotification) {
                            inconsistentItems.push({
                                联系人: item.联系人,
                                联系电话: item.联系电话,
                                随车负责人: currentEscort,
                                随车负责人电话: currentEscortPhone,
                                通知内容: currentNotification
                            });
                        }
                    });
                    
                    if (inconsistentItems.length > 0) {
                        inconsistencies.push({
                            车牌号码: plate,
                            线路: line,
                            参考信息: {
                                随车负责人: referenceEscort,
                                随车负责人电话: referenceEscortPhone,
                                通知内容: referenceNotification
                            },
                            不一致项: inconsistentItems
                        });
                    }
                }
            });
            
            return inconsistencies;
        }

        function showInconsistencyDialog(inconsistencies) {
            const dialogOverlay = document.getElementById('inconsistencyDialogOverlay');
            const inconsistencyList = document.getElementById('inconsistencyList');
            
            // Clear previous content
            inconsistencyList.innerHTML = '';
            
            // Generate HTML for each inconsistency
            inconsistencies.forEach(item => {
                const inconsistencyItem = document.createElement('div');
                inconsistencyItem.className = 'inconsistency-item';
                
                // Vehicle and line information
                inconsistencyItem.innerHTML = `
                    <h4>${item.车牌号码} - ${item.线路}</h4>
                    <div class="reference-info">
                        <p><strong>参考信息：</strong></p>
                        <p>随车负责人: ${item.参考信息.随车负责人 || '空'}</p>
                        <p>随车负责人电话: ${item.参考信息.随车负责人电话 || '空'}</p>
                        <p>通知内容: ${item.参考信息.通知内容 || '空'}</p>
                    </div>
                    <div class="inconsistent-orders">
                        <p><strong>不一致的订单：</strong></p>
                `;
                
                // Add each inconsistent order
                item.不一致项.forEach(order => {
                    const orderDiv = document.createElement('div');
                    orderDiv.className = 'inconsistent-order';
                    
                    // Check which fields are inconsistent
                    const reference = item.参考信息;
                    const inconsistentFields = [];
                    if (order.随车负责人 !== reference.随车负责人) inconsistentFields.push('随车负责人');
                    if (order.随车负责人电话 !== reference.随车负责人电话) inconsistentFields.push('随车负责人电话');
                    if (order.通知内容 !== reference.通知内容) inconsistentFields.push('通知内容');
                    
                    orderDiv.innerHTML = `
                        <p><strong>联系人：</strong>${order.联系人}</p>
                        <p><strong>联系电话：</strong>${order.联系电话}</p>
                        ${inconsistentFields.map(field => `
                            <p><span class="inconsistent-field">${field}：</span>${order[field] || '空'}</p>
                        `).join('')}
                    `;
                    
                    inconsistencyItem.appendChild(orderDiv);
                });
                
                inconsistencyList.appendChild(inconsistencyItem);
            });
            
            // Show the dialog
            dialogOverlay.style.display = 'flex';
            
            // Handle cancel button
            document.getElementById('inconsistencyDialogCancelBtn').onclick = function() {
                dialogOverlay.style.display = 'none';
            };
            
            // Handle continue button
            document.getElementById('inconsistencyDialogOkBtn').onclick = function() {
                dialogOverlay.style.display = 'none';
                // Continue with export
                exportDataWithoutCheck();
            };
        }

        // Check vehicles without managers
        function checkVehiclesWithoutManagers() {
            const vehiclesWithoutManagers = [];
            
            // Group data by vehicle
            const vehicleGroups = {};
            processedData.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                if (!vehicleGroups[plate]) {
                    vehicleGroups[plate] = [];
                }
                vehicleGroups[plate].push(item);
            });
            
            // Check each vehicle for manager
            Object.keys(vehicleGroups).forEach(plate => {
                if (plate === '未分配车辆' || plate === '') return;
                
                const group = vehicleGroups[plate];
                const hasManager = group.some(item => item.随车负责人 && item.随车负责人.trim() !== '');
                
                if (!hasManager) {
                    vehiclesWithoutManagers.push(plate);
                }
            });
            
            return vehiclesWithoutManagers;
        }

        // Show vehicles without managers dialog
        function showVehiclesWithoutManagersDialog(vehicles) {
            const dialogOverlay = document.createElement('div');
            dialogOverlay.className = 'confirm-dialog-overlay';
            dialogOverlay.style.display = 'flex';
            dialogOverlay.innerHTML = `
                <div class="confirm-dialog">
                    <div class="confirm-dialog-header">
                        <h3><i class="fas fa-exclamation-triangle"></i> 车辆缺少随车负责人</h3>
                    </div>
                    <div class="confirm-dialog-body">
                        <p>以下车辆没有设置随车负责人：</p>
                        <div class="inconsistency-list" style="margin: 15px 0;">
                            ${vehicles.map(plate => `<div class="inconsistency-item" style="margin-bottom: 10px;"><h4>${plate}</h4></div>`).join('')}
                        </div>
                        <p style="margin-top: 15px; color: #e74c3c; font-weight: bold;">
                            注意：每辆车都应该有随车负责人，以确保行程安全。
                        </p>
                    </div>
                    <div class="confirm-dialog-footer">
                        <button class="confirm-dialog-btn confirm-dialog-cancel-btn">取消</button>
                        <button class="confirm-dialog-btn confirm-dialog-ok-btn">继续导出</button>
                    </div>
                </div>
            `;
            
            document.body.appendChild(dialogOverlay);
            
            // Handle cancel button
            dialogOverlay.querySelector('.confirm-dialog-cancel-btn').onclick = function() {
                dialogOverlay.remove();
            };
            
            // Handle continue button
            dialogOverlay.querySelector('.confirm-dialog-ok-btn').onclick = function() {
                dialogOverlay.remove();
                // Continue with export
                exportDataWithoutCheck();
            };
        }

        // Check overloaded vehicles
        function checkOverloadedVehicles() {
            const overloadedVehicles = [];
            
            // Group data by vehicle
            const vehicleGroups = {};
            processedData.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                if (!vehicleGroups[plate]) {
                    vehicleGroups[plate] = {
                        totalTickets: 0,
                        orders: []
                    };
                }
                vehicleGroups[plate].totalTickets += parseInt(item.票数) || 0;
                vehicleGroups[plate].orders.push(item);
            });
            
            // Check each vehicle for overload
            Object.keys(vehicleGroups).forEach(plate => {
                if (plate === '未分配车辆' || plate === '') return;
                
                const group = vehicleGroups[plate];
                const capacity = vehicleCapacityMap.has(plate) ? (vehicleCapacityMap.get(plate).容量 || 0) : 0;
                
                if (capacity > 0 && group.totalTickets > capacity) {
                    overloadedVehicles.push({
                        车牌号码: plate,
                        总票数: group.totalTickets,
                        座位数: capacity,
                        超载数: group.totalTickets - capacity
                    });
                }
            });
            
            return overloadedVehicles;
        }

        // Show overloaded vehicles dialog
        function showOverloadedVehiclesDialog(vehicles) {
            const dialogOverlay = document.getElementById('overloadDialogOverlay');
            const overloadList = document.getElementById('overloadList');
            
            // Clear previous content
            overloadList.innerHTML = '';
            
            // Generate HTML for each overloaded vehicle
            vehicles.forEach(vehicle => {
                const overloadItem = document.createElement('div');
                overloadItem.className = 'inconsistency-item';
                
                overloadItem.innerHTML = `
                    <h4>${vehicle.车牌号码}</h4>
                    <p>总票数: ${vehicle.总票数}</p>
                    <p>座位数: ${vehicle.座位数}</p>
                    <p style="color: #e74c3c; font-weight: bold;">超载数: ${vehicle.超载数}</p>
                `;
                
                overloadList.appendChild(overloadItem);
            });
            
            // Show the dialog
            dialogOverlay.style.display = 'flex';
            
            // Handle cancel button
            document.getElementById('overloadDialogCancelBtn').onclick = function() {
                dialogOverlay.style.display = 'none';
            };
            
            // Handle continue button
            document.getElementById('overloadDialogOkBtn').onclick = function() {
                dialogOverlay.style.display = 'none';
                // Continue with export
                exportDataWithoutCheck();
            };
        }

        function exportData() {
            if (processedData.length === 0) {
                showNotification('没有可导出的数据', 'error');
                return;
            }
            
            // Check for inconsistencies before exporting
            const inconsistencies = checkInconsistencies();
            if (inconsistencies.length > 0) {
                showInconsistencyDialog(inconsistencies);
                return;
            }
            
            // Check for vehicles without managers
            const vehiclesWithoutManagers = checkVehiclesWithoutManagers();
            if (vehiclesWithoutManagers.length > 0) {
                showVehiclesWithoutManagersDialog(vehiclesWithoutManagers);
                return;
            }
            
            // Check for overloaded vehicles
            const overloadedVehicles = checkOverloadedVehicles();
            if (overloadedVehicles.length > 0) {
                showOverloadedVehiclesDialog(overloadedVehicles);
                return;
            }
            
            // If no issues, proceed with export
            exportDataWithoutCheck();
        }

        function exportDataWithoutCheck() {
            
            try {
                const sheetData = [];
                
                // 表头行：增加分销员手机列
                sheetData.push({
                    '订单ID': '订单ID',                    
                    '线路': '线路',
                    '车牌号码': '车牌号码',
                    '联系人': '联系人',
                    '联系电话': '联系电话',
                    '票数': '票数',
                    '乘车日期': '乘车日期',
                    '乘车时间': '乘车时间',
                    '上车地点': '上车地点',
                    '下车地点': '下车地点',
                    '随车负责人': '随车负责人',
                    '随车负责人电话': '随车负责人电话',
                    '通知内容': '通知内容',
                    '分销员手机': '分销员手机', 
                    '修改时间': '修改时间'
                });
                
                // 使用与displayProcessedData相同的数据源
                const dataToExport = isSearchActive ? filteredData : processedData;
                
                const plateGroups = {};
                dataToExport.forEach(item => {
                    const plate = item.车牌号码 || '未分配车辆';
                    if (!plateGroups[plate]) {
                        plateGroups[plate] = [];
                    }
                    plateGroups[plate].push(item);
                });
                
                const plates = Object.keys(plateGroups).sort((a, b) => {
                    // 未分配车辆始终排在最后
                    if (a === '未分配车辆' || a === '') return 1;
                    if (b === '未分配车辆' || b === '') return -1;
                    
                    // 获取车辆顺序
                    const orderA = vehicleCapacityMap.has(a) ? (vehicleCapacityMap.get(a).顺序 || 9999) : 9999;
                    const orderB = vehicleCapacityMap.has(b) ? (vehicleCapacityMap.get(b).顺序 || 9999) : 9999;
                    
                    // 按照顺序排序
                    if (orderA !== orderB) {
                        return orderA - orderB;
                    }
                    
                    // 顺序相同，按车牌号排序
                    return a.localeCompare(b);
                });
                
                let firstGroup = true;
                plates.forEach(plate => {
                    if (!firstGroup) {
                        sheetData.push({});
                    }
                    firstGroup = false;
                    
                    // 根据选择的排序方式对组内数据进行排序
                    let groupData = plateGroups[plate];
                    
                    // 优先将随车联系人的订单排在前面
                    groupData.sort((a, b) => {
                        // 检查是否是随车联系人
                        const isManagerA = a.联系电话 === a.随车负责人电话;
                        const isManagerB = b.联系电话 === b.随车负责人电话;
                        
                        if (isManagerA && !isManagerB) return -1;
                        if (!isManagerA && isManagerB) return 1;
                        
                        // 然后根据选择的排序方式排序
                        if (exportSortType === 'line') {
                            // 按线路排序
                            const lineA = a.线路 || '';
                            const lineB = b.线路 || '';
                            return lineA.localeCompare(lineB);
                        } else if (exportSortType === 'campus') {
                            // 按校区（上车地点）排序
                            const pickupA = a.上车地点 || '';
                            const pickupB = b.上车地点 || '';
                            return pickupA.localeCompare(pickupB);
                        }
                        
                        return 0;
                    });
                    
                    groupData.forEach(item => {
                        sheetData.push({
                            '订单ID': item.订单ID,
                            '线路': item.线路,
                            '车牌号码': item.车牌号码 || "",
                            '联系人': item.联系人,
                            '联系电话': item.联系电话,
                            '票数': item.票数,
                            '乘车日期': item.乘车日期,
                            '乘车时间': item.乘车时间,
                            '上车地点': item.上车地点,
                            '下车地点': item.下车地点,
                            '随车负责人': item.随车负责人 || "",
                            '随车负责人电话': item.随车负责人电话 || "",
                            '通知内容': item.通知内容,
                            '分销员手机': item.分销员手机 || "",
                            '修改时间': item.修改时间
                        });
                    });
                });
                
                const worksheet = XLSX.utils.json_to_sheet(sheetData, {skipHeader: true});
                
                // 调整列宽，与预览表格字段顺序一致
                const colWidths = [
                    { wch: 10 }, // 订单ID
                    { wch: 8 },  // 线路
                    { wch: 12 }, // 车牌号码
                    { wch: 8 },  // 联系人
                    { wch: 12 }, // 联系电话
                    { wch: 8 },  // 票数
                    { wch: 12 }, // 乘车日期
                    { wch: 10 }, // 乘车时间
                    { wch: 10 }, // 上车地点
                    { wch: 15 }, // 下车地点
                    { wch: 10 }, // 随车负责人
                    { wch: 12 }, // 随车负责人电话
                    { wch: 20 }, // 通知内容
                    { wch: 12 }, // 分销员手机
                    { wch: 15 }  // 修改时间
                ];
                worksheet['!cols'] = colWidths;
                
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                
                // 表头样式
                for (let col = range.s.c; col <= range.e.c; col++) {
                    const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
                    if (!worksheet[cellAddress]) continue;
                    
                    if (!worksheet[cellAddress].s) worksheet[cellAddress].s = {};
                    worksheet[cellAddress].s.fill = {
                        patternType: "solid",
                        fgColor: { rgb: "4A6491" }
                    };
                    worksheet[cellAddress].s.font = {
                        name: '微软雅黑',
                        sz: 11,
                        bold: true,
                        color: { rgb: "FFFFFF" }
                    };
                    worksheet[cellAddress].s.alignment = {
                        vertical: "center",
                        horizontal: "center"
                    };
                }
                
                // 数据行按线路颜色标记
                let currentRow = 1;
                plates.forEach(plate => {
                    const groupData = plateGroups[plate];
                    groupData.forEach(item => {
                        const color = lineColorMap.get(item.线路);
                        if (color) {
                            const excelColor = color.replace('#', '');
                            for (let c = range.s.c; c <= range.e.c; c++) {
                                const cellAddress = XLSX.utils.encode_cell({ r: currentRow, c: c });
                                if (!worksheet[cellAddress]) continue;
                                if (!worksheet[cellAddress].s) worksheet[cellAddress].s = {};
                                worksheet[cellAddress].s.fill = {
                                    patternType: "solid",
                                    fgColor: { rgb: excelColor }
                                };
                                worksheet[cellAddress].s.font = {
                                    name: '微软雅黑',
                                    sz: 11
                                };
                            }
                        }
                        currentRow++;
                    });
                    currentRow++; // 组间空行
                });
                
                const newWorkbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(newWorkbook, worksheet, '线路订单整合数据');
                
                const sortText = exportSortType === 'line' ? '按线路' : '按校区';
                const exportFileName = fileName ? 
                    `整合后的_${fileName.replace(/\.[^/.]+$/, "")}_按车牌分组_${sortText}排序.xlsx` : 
                    `线路订单整合数据_按车牌分组_${sortText}排序.xlsx`;
                
                XLSX.writeFile(newWorkbook, exportFileName);
                
                showNotification(`文件已导出: ${exportFileName}`, 'success');
                showNotification(`已按${sortText}排序导出文件`, 'info');
            } catch (error) {
                console.error('Error exporting data:', error);
                showNotification('导出文件失败: ' + error.message, 'error');
            }
        }
        
        // 重置显示
        function resetDisplay() {
            originalData = [];
            processedData = [];
            filteredData = [];
            processedDataBody.innerHTML = '<tr><td colspan="15" style="text-align:center;">请先上传并分析数据</td></tr>';
            ticketStatsSection.style.display = 'none';
            ticketStatsSection.classList.remove('ticket-stats-down');
            dataSection.style.display = 'none';
            vehiclePickupStatsPanel.style.display = 'none';
            lineTimeLocationPanel.style.display = 'none';
            isDataVisible = false;
            isSearchActive = false;
            analyzeBtn.disabled = true;
            toggleBtn.disabled = true;
            toggleBtn.innerHTML = '<i class="fas fa-eye"></i> 显示整合后的数据';
            exportBtn.disabled = true;
            lineInfo = {};
            vehiclePickupStats = {};
            lineOrder = [];
            selectedRows.clear();
            selectAllCheckbox.checked = false;
            searchResultsInfo.textContent = '';
            
            // 清空搜索输入框
            searchContact.value = '';
            searchPhone.value = '';
            searchPlate.value = '';
            searchLine.value = '';
            searchPickup.value = '';
            searchDropoff.value = '';
            
            updateModifyButtonState();
            lineUnassignedCount = {};
            vehicleCapacityMap.clear();
            plateStats = {};
            selectAllPlateBtn.style.display = 'none';
            selectAllUnassignedBtn.style.display = 'none';
            manuallyAssignedOrders.clear();
            lineManuallyAssignedCount = {};
            plateLineMap.clear();
            vehicleTotalAssigned.clear(); // 清空车辆总人数跟踪
            
            // 清空车辆信息错误
            vehicleInfoErrors = [];
            
            // 清空线路时间地点统计数据
            lineTimeLocationData = [];
            lineTimeLocationBody.innerHTML = '';
            
            // 重置线路统计面板显示状态
            lineStatsVisible = true;
        }
        
        // 新增：显示刷新确认对话框
        function showRefreshConfirm() {
            confirmDialogOverlay.style.display = 'flex';
        }
        
        // 新增：关闭确认对话框
        function closeConfirmDialog() {
            confirmDialogOverlay.style.display = 'none';
        }
        
        // 新增：确认刷新
        function confirmRefresh() {
            closeConfirmDialog();
            
            // 执行刷新操作
            fileInput.value = '';
            fileNameDisplay.textContent = '';
            
            resetDisplay();
            
            colorMap.clear();
            lineColorMap.clear();
            colorIndex = 0;
            
            lastLineInfo = {};
            
            originalData = generateSampleData();
            
            analyzeBtn.disabled = false;
            
            showNotification('页面已刷新，可以重新上传文件或使用示例数据进行分析', 'info');
        }
        
        // 显示通知
        function showNotification(message, type, center = false) {
            const oldNotifications = document.querySelectorAll('.notification');
            oldNotifications.forEach(notification => {
                notification.style.animation = 'slideOut 0.3s ease-out';
                setTimeout(() => {
                    if (notification.parentNode) {
                        notification.parentNode.removeChild(notification);
                    }
                }, 300);
            });
            
            const notification = document.createElement('div');
            notification.className = `notification ${type}`;
            if (center) {
                notification.classList.add('center');
            }
            notification.textContent = message;
            
            document.body.appendChild(notification);
            
            setTimeout(() => {
                notification.style.animation = 'slideOut 0.3s ease-out';
                setTimeout(() => {
                    if (notification.parentNode) {
                        document.body.removeChild(notification);
                    }
                }, 300);
            }, 3000);
        }
        
        // 执行搜索
        function performSearch() {
            // 获取所有搜索框的值
            const contactSearch = searchContact.value.trim().toLowerCase();
            const phoneSearch = searchPhone.value.trim().toLowerCase();
            const plateSearch = searchPlate.value.trim().toLowerCase();
            const lineSearch = searchLine.value.trim().toLowerCase();
            const pickupSearch = searchPickup.value.trim().toLowerCase();
            const dropoffSearch = searchDropoff.value.trim().toLowerCase();
            
            // 检查是否有任何搜索条件
            const hasSearchConditions = contactSearch || phoneSearch || plateSearch || lineSearch || pickupSearch || dropoffSearch;
            
            if (!hasSearchConditions) {
                isSearchActive = false;
                filteredData = [];
                searchResultsInfo.textContent = `显示全部 ${processedData.length} 条数据`;
                searchResultsInfo.style.color = '#4a6491';
            } else {
                isSearchActive = true;
                
                filteredData = processedData.filter(item => {
                    // 检查每个搜索条件，如果条件为空，则视为匹配
                    const contactMatch = !contactSearch || (item.联系人 && item.联系人.toLowerCase().includes(contactSearch));
                    const phoneMatch = !phoneSearch || (item.联系电话 && item.联系电话.toLowerCase().includes(phoneSearch));
                    const plateMatch = !plateSearch || (item.车牌号码 && item.车牌号码.toLowerCase().includes(plateSearch));
                    const lineMatch = !lineSearch || (item.线路 && item.线路.toLowerCase().includes(lineSearch));
                    const pickupMatch = !pickupSearch || (item.上车地点 && item.上车地点.toLowerCase().includes(pickupSearch));
                    const dropoffMatch = !dropoffSearch || (item.下车地点 && item.下车地点.toLowerCase().includes(dropoffSearch));
                    
                    // 所有条件都必须匹配
                    return contactMatch && phoneMatch && plateMatch && lineMatch && pickupMatch && dropoffMatch;
                });
                
                // 构建搜索条件文本
                const searchConditions = [];
                if (contactSearch) searchConditions.push(`联系人: ${contactSearch}`);
                if (phoneSearch) searchConditions.push(`电话: ${phoneSearch}`);
                if (plateSearch) searchConditions.push(`车牌: ${plateSearch}`);
                if (lineSearch) searchConditions.push(`线路: ${lineSearch}`);
                if (pickupSearch) searchConditions.push(`上车: ${pickupSearch}`);
                if (dropoffSearch) searchConditions.push(`下车: ${dropoffSearch}`);
                
                const conditionsText = searchConditions.join('; ');
                searchResultsInfo.textContent = `找到 ${filteredData.length} 条数据 (${conditionsText})`;
                searchResultsInfo.style.color = '#2ecc71';
            }
            
            if (isDataVisible) {
                displayProcessedData();
                displayVehiclePickupStatsPanel();
            }
        }
        
        // 清除搜索
        function clearSearch() {
            isSearchActive = false;
            filteredData = [];
            searchResultsInfo.textContent = `显示全部 ${processedData.length} 条数据`;
            searchResultsInfo.style.color = '#4a6491';
            
            // 清空所有搜索输入框
            searchContact.value = '';
            searchPhone.value = '';
            searchPlate.value = '';
            searchLine.value = '';
            searchPickup.value = '';
            searchDropoff.value = '';
            
            if (isDataVisible) {
                displayProcessedData();
                displayVehiclePickupStatsPanel();
            }
        }
        
        // 过滤弹窗中的行
        function filterModalRows() {
            const lineSearch = modalSearchLine.value.trim().toLowerCase();
            const plateSearch = modalSearchPlate.value.trim().toLowerCase();
            
            const rows = document.querySelectorAll('#lineInputsBody tr');
            
            // 检查是否有任何搜索条件
            const hasSearchConditions = lineSearch || plateSearch;
            
            if (!hasSearchConditions) {
                rows.forEach(row => {
                    row.classList.remove('filtered-out');
                });
                
                // 移除之前的无结果提示
                const existingNoResults = lineInputsBody.parentNode.querySelector('.no-results');
                if (existingNoResults) {
                    existingNoResults.remove();
                }
                return;
            }
            
            let hasResults = false;
            
            rows.forEach(row => {
                // 获取行数据
                const lineCell = row.querySelector('td:nth-child(2)');
                const plateInput = row.querySelector('.plate-input'); // 通过class获取
                
                const lineText = lineCell ? lineCell.textContent.toLowerCase() : '';
                const plateText = plateInput ? plateInput.value.toLowerCase() : '';
                
                // 检查是否匹配
                const lineMatch = !lineSearch || lineText.includes(lineSearch);
                const plateMatch = !plateSearch || plateText.includes(plateSearch);
                
                const shouldShow = lineMatch && plateMatch;
                
                if (shouldShow) {
                    row.classList.remove('filtered-out');
                    hasResults = true;
                } else {
                    row.classList.add('filtered-out');
                }
            });
            
            // 如果没有结果，显示提示
            const existingNoResults = lineInputsBody.parentNode.querySelector('.no-results');
            if (!hasResults) {
                if (!existingNoResults) {
                    const noResults = document.createElement('div');
                    noResults.className = 'no-results';
                    noResults.textContent = `未找到匹配的数据`;
                    lineInputsBody.parentNode.appendChild(noResults);
                }
            } else if (existingNoResults) {
                existingNoResults.remove();
            }
        }
        
        // 更新修改按钮状态
        function updateModifyButtonState() {
            modifyBtn.disabled = selectedRows.size === 0;
            
            if (isDataVisible) {
                updateSelectAllButtons();
            }
        }
        
        // 全选/取消全选
        function toggleSelectAll() {
            const checkboxes = document.querySelectorAll('.row-checkbox');
            const dataToSelect = isSearchActive ? filteredData : processedData;
            
            if (selectAllCheckbox.checked) {
                dataToSelect.forEach(item => {
                    selectedRows.add(item.订单ID);
                });
                
                checkboxes.forEach(checkbox => {
                    checkbox.checked = true;
                    const row = checkbox.closest('tr');
                    if (row) row.classList.add('selected-row');
                });
                
                document.querySelectorAll('.plate-checkbox').forEach(checkbox => {
                    checkbox.checked = true;
                });
                
                document.querySelectorAll('.unassigned-line-checkbox').forEach(checkbox => {
                    checkbox.checked = true;
                });
            } else {
                dataToSelect.forEach(item => {
                    selectedRows.delete(item.订单ID);
                });
                
                checkboxes.forEach(checkbox => {
                    checkbox.checked = false;
                    const row = checkbox.closest('tr');
                    if (row) row.classList.remove('selected-row');
                });
                
                document.querySelectorAll('.plate-checkbox').forEach(checkbox => {
                    checkbox.checked = false;
                });
                
                document.querySelectorAll('.unassigned-line-checkbox').forEach(checkbox => {
                    checkbox.checked = false;
                });
            }
            
            updateModifyButtonState();
        }
        
        // 显示修改弹窗
        function showModifyModal() {
            if (selectedRows.size === 0) {
                showNotification('请先选择要修改的订单', 'error');
                return;
            }
            
            const firstOrderId = Array.from(selectedRows)[0];
            const orderData = processedData.find(item => item.订单ID === firstOrderId);
            
            if (!orderData) {
                showNotification('未找到选中的订单数据', 'error');
                return;
            }
            
            const selectedLines = new Set();
            let totalTickets = 0;
            selectedRows.forEach(orderId => {
                const order = processedData.find(item => item.订单ID === orderId);
                if (order) {
                    selectedLines.add(order.线路);
                    totalTickets += (order.票数 || 0);
                }
            });
            
            let formHTML = `
                <div class="form-group">
                    <label>选中订单数量</label>
                    <div class="readonly-field">${selectedRows.size} 个订单</div>
                </div>
                <div class="form-group">
                    <label>涉及线路</label>
                    <div class="readonly-field">${Array.from(selectedLines).join('、')}</div>
                </div>
                <div class="form-group">
                    <label>总票数</label>
                    <div class="readonly-field">${totalTickets} 票</div>
                </div>
            `;

            if (selectedRows.size === 1) {
                formHTML += `
                    <div class="form-group">
                        <label>订单ID</label>
                        <div class="readonly-field">${orderData.订单ID}</div>
                    </div>
                    <div class="form-group">
                        <label>联系人</label>
                        <div class="readonly-field">${orderData.联系人}</div>
                    </div>
                    <div class="form-group">
                        <label>联系电话</label>
                        <div class="readonly-field">${orderData.联系电话}</div>
                    </div>
                    <div class="form-group">
                        <label>乘车时间</label>
                        <input type="text" id="modifyTime" class="input-field" value="${orderData.乘车时间 || ''}" placeholder="请输入乘车时间（如：08:30）">
                    </div>
                    <div class="form-group">
                        <label>上车地点</label>
                        <input type="text" id="modifyPickup" class="input-field" value="${orderData.上车地点 || ''}" placeholder="请输入上车地点">
                    </div>
                    <div class="form-group">
                        <label>下车地点</label>
                        <input type="text" id="modifyDropoff" class="input-field" value="${orderData.下车地点 || ''}" placeholder="请输入下车地点">
                    </div>
                `;
            } else {
                formHTML += `
                    <div class="form-group">
                        <label>乘车时间</label>
                        <input type="text" id="modifyTime" class="input-field" placeholder="请输入乘车时间">
                        <small style="color:#666;">为空则不修改此字段</small>
                    </div>
                    <div class="form-group">
                        <label>上车地点</label>
                        <input type="text" id="modifyPickup" class="input-field" placeholder="请输入上车地点">
                        <small style="color:#666;">为空则不修改此字段</small>
                    </div>
                    <div class="form-group">
                        <label>下车地点</label>
                        <input type="text" id="modifyDropoff" class="input-field" placeholder="请输入下车地点">
                        <small style="color:#666;">为空则不修改此字段</small>
                    </div>
                `;
            }

            formHTML += `
                <div class="form-group">
                    <label>车牌号码</label>
                    <div style="display: flex; gap: 10px; align-items: center;">
                        <input type="text" id="modifyPlate" class="input-field" style="flex: 1;" value="${orderData.车牌号码 || ''}" placeholder="请输入车牌号码">
                        <button type="button" id="identifyInfoBtn" class="process-btn" style="background-color: #3498db; color: white; padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer;">
                            识别信息
                        </button>
                    </div>
                    <small style="color:#666;">如果输入已有车牌号，订单将归并到该车，并更新座位统计</small>
                </div>
                <div class="form-group">
                    <label>随车负责人</label>
                    <input type="text" id="modifyManager" class="input-field" value="${orderData.随车负责人 || ''}" placeholder="请输入随车负责人">
                    <small style="color:#666;">为空则自动匹配输入车牌号中的信息</small>
                </div>
                <div class="form-group">
                    <label>随车负责人电话</label>
                    <input type="text" id="modifyManagerPhone" class="input-field" value="${orderData.随车负责人电话 || ''}" placeholder="请输入随车负责人电话">
                    <small style="color:#666;">为空则自动匹配输入车牌号中的信息</small>
                </div>
                <div class="form-group">
                    <label>通知内容</label>
                    ${selectedLines.size > 1 ? `
                    <div style="border: 1px solid #ddd; border-radius: 4px; padding: 10px;">
                        ${Array.from(selectedLines).map(line => {
                            // 查找该线路的通知内容
                            const lineOrders = Array.from(selectedRows).map(orderId => processedData.find(item => item.订单ID === orderId)).filter(order => order && order.线路 === line);
                            const notice = lineOrders.length > 0 ? lineOrders[0].通知内容 || '' : '';
                            return `
                            <div style="margin-bottom: 10px;">
                                <strong>${line}:</strong>
                                <div style="margin-top: 5px; padding: 8px; background-color: #f9f9f9; border-radius: 4px;">
                                    ${notice || '（无通知内容）'}
                                </div>
                            </div>
                            `;
                        }).join('')}
                    </div>
                    <small style="color:#666;">批量选择了多个路线，显示各路线的通知内容</small>
                    ` : `
                    <textarea id="modifyNotice" class="input-field" rows="3" placeholder="请输入通知内容">${orderData.通知内容 || ''}</textarea>
                    <small style="color:#666;">为空则自动匹配输入车牌号中的信息</small>
                    `}
                </div>
            `;

            modifyForm.innerHTML = formHTML;
            
            // 添加识别信息按钮的事件监听器
            const identifyInfoBtn = document.getElementById('identifyInfoBtn');
            if (identifyInfoBtn) {
                identifyInfoBtn.addEventListener('click', function() {
                    const plate = document.getElementById('modifyPlate').value.trim();
                    if (!plate) {
                        showNotification('请先输入车牌号码', 'error');
                        return;
                    }
                    
                    // 按线路分组处理
                    const ordersByLine = {};
                    selectedRows.forEach(orderId => {
                        const order = processedData.find(item => item.订单ID === orderId);
                        if (order) {
                            const line = order.线路;
                            if (!ordersByLine[line]) {
                                ordersByLine[line] = [];
                            }
                            ordersByLine[line].push(order);
                        }
                    });
                    
                    // 对每个线路查找相同车牌和线路的其他订单
                    let foundInfo = false;
                    Object.keys(ordersByLine).forEach(line => {
                        // 查找相同车牌和线路的其他订单
                        const samePlateLineOrders = processedData.filter(item => 
                            item.车牌号码 === plate && 
                            item.线路 === line
                        );
                        
                        if (samePlateLineOrders.length > 0) {
                            // 使用第一个找到的订单的信息
                            const referenceOrder = samePlateLineOrders[0];
                            
                            // 填充表单
                            if (!document.getElementById('modifyManager').value) {
                                document.getElementById('modifyManager').value = referenceOrder.随车负责人 || '';
                            }
                            if (!document.getElementById('modifyManagerPhone').value) {
                                document.getElementById('modifyManagerPhone').value = referenceOrder.随车负责人电话 || '';
                            }
                            if (!document.getElementById('modifyNotice').value) {
                                document.getElementById('modifyNotice').value = referenceOrder.通知内容 || '';
                            }
                            
                            foundInfo = true;
                        }
                    });
                    
                    if (foundInfo) {
                        showNotification('已自动填充相同路线的车辆信息', 'success');
                    } else {
                        showNotification('未找到该车牌相同路线的信息', 'info');
                    }
                });
            }
            
            // 显示或隐藏删除乘客按钮
            if (selectedRows.size === 1) {
                deletePassengerBtn.style.display = 'flex';
            } else {
                deletePassengerBtn.style.display = 'none';
            }

            modifyModalOverlay.style.display = 'flex';
        }
        
        // 关闭修改弹窗
        function closeModifyModal() {
            modifyModalOverlay.style.display = 'none';
        }
        
        // 新增：显示添加乘客弹窗
        function showAddPassengerModal() {
            // 清空所有输入框
            document.getElementById('addPassengerSmartInput').value = '';
            document.getElementById('addPassengerOrderId').value = '';
            document.getElementById('addPassengerLine').value = '';
            document.getElementById('addPassengerContact').value = '';
            document.getElementById('addPassengerPhone').value = '';
            document.getElementById('addPassengerTicketCount').value = '1';
            document.getElementById('addPassengerPickup').value = '';
            document.getElementById('addPassengerDropoff').value = '';
            document.getElementById('addPassengerDate').value = '';
            document.getElementById('addPassengerTime').value = '';
            document.getElementById('addPassengerPlate').innerHTML = '<option value="">请选择车牌号码</option>';
            document.getElementById('addPassengerManager').value = '';
            document.getElementById('addPassengerManagerPhone').value = '';
            document.getElementById('addPassengerNotice').value = '';
            
            // 填充车牌号码选项
            const plateSelect = document.getElementById('addPassengerPlate');
            const existingPlates = new Set();
            processedData.forEach(item => {
                if (item.车牌号码 && item.车牌号码.trim()) {
                    existingPlates.add(item.车牌号码.trim());
                }
            });
            existingPlates.forEach(plate => {
                const option = document.createElement('option');
                option.value = plate;
                option.textContent = plate;
                plateSelect.appendChild(option);
            });
            
            // 添加识别信息按钮的事件监听器
            const addPassengerIdentifyInfoBtn = document.getElementById('addPassengerIdentifyInfoBtn');
            if (addPassengerIdentifyInfoBtn) {
                // 先移除可能存在的旧监听器
                addPassengerIdentifyInfoBtn.onclick = null;
                
                addPassengerIdentifyInfoBtn.addEventListener('click', function() {
                    const plate = document.getElementById('addPassengerPlate').value;
                    const line = document.getElementById('addPassengerLine').value.trim();
                    
                    if (!plate) {
                        showNotification('请先选择车牌号码', 'error');
                        return;
                    }
                    
                    if (!line) {
                        showNotification('请先输入线路名称', 'error');
                        return;
                    }
                    
                    // 查找相同车牌和线路的其他订单
                    const samePlateLineOrders = processedData.filter(item => 
                        item.车牌号码 === plate && 
                        item.线路 === line
                    );
                    
                    if (samePlateLineOrders.length > 0) {
                        // 使用第一个找到的订单的信息
                        const referenceOrder = samePlateLineOrders[0];
                        
                        // 填充表单
                        if (!document.getElementById('addPassengerManager').value) {
                            document.getElementById('addPassengerManager').value = referenceOrder.随车负责人 || '';
                        }
                        if (!document.getElementById('addPassengerManagerPhone').value) {
                            document.getElementById('addPassengerManagerPhone').value = referenceOrder.随车负责人电话 || '';
                        }
                        if (!document.getElementById('addPassengerNotice').value) {
                            document.getElementById('addPassengerNotice').value = referenceOrder.通知内容 || '';
                        }
                        
                        showNotification('已自动填充相同路线的车辆信息', 'success');
                    } else {
                        showNotification('未找到该车牌相同路线的信息', 'info');
                    }
                });
            }
            
            addPassengerModalOverlay.style.display = 'flex';
        }
        
        // 新增：识别线路函数
        // 智能识别函数
        function smartIdentify() {
            const smartInput = document.getElementById('addPassengerSmartInput').value.trim();
            
            if (!smartInput) {
                showNotification('请输入智能识别信息', 'error');
                return;
            }
            
            // 提取订单号
            const orderIdMatch = smartInput.match(/订单号：\s*(\d+)/);
            const orderId = orderIdMatch ? orderIdMatch[1] : '';
            
            // 提取线路
            const lineMatch = smartInput.match(/线路：\s*(.+?)(?=\n|$)/);
            let line = lineMatch ? lineMatch[1] : '';
            // 转换线路格式：阜南→安庆 → 阜南至安庆
            line = line.replace(/→/g, '至');
            
            // 提取出发时间
            const departureMatch = smartInput.match(/出发时间：\s*(\d{4}-\d{2}-\d{2})\s*(\d{2}:\d{2})/);
            const date = departureMatch ? departureMatch[1] : '';
            const time = departureMatch ? departureMatch[2] : '';
            
            // 提取手机
            const phoneMatch = smartInput.match(/手机：\s*(\d{11})/);
            const phone = phoneMatch ? phoneMatch[1] : '';
            
            // 提取联系人
            const contactMatch = smartInput.match(/联系人：\s*(.+?)(?=\n|$)/);
            const contact = contactMatch ? contactMatch[1] : '';
            
            // 提取全票和童票
            const fullTicketMatch = smartInput.match(/全票：\s*(\d+)/);
            const childTicketMatch = smartInput.match(/童票：\s*(\d+)/);
            const fullTicket = fullTicketMatch ? parseInt(fullTicketMatch[1]) : 0;
            const childTicket = childTicketMatch ? parseInt(childTicketMatch[1]) : 0;
            const ticketCount = fullTicket + childTicket;
            
            // 提取上车地点
            const pickupMatch = smartInput.match(/上车：\s*(.+?)(?=\n|$)/);
            const pickup = pickupMatch ? pickupMatch[1] : '';
            
            // 提取下车地点
            const dropoffMatch = smartInput.match(/下车：\s*(.+?)(?=\n|$)/);
            const dropoff = dropoffMatch ? dropoffMatch[1] : '';
            
            // 填充表单字段
            document.getElementById('addPassengerOrderId').value = orderId;
            document.getElementById('addPassengerLine').value = line;
            document.getElementById('addPassengerDate').value = date;
            document.getElementById('addPassengerTime').value = time;
            document.getElementById('addPassengerPhone').value = phone;
            document.getElementById('addPassengerContact').value = contact;
            document.getElementById('addPassengerTicketCount').value = ticketCount || 1;
            document.getElementById('addPassengerPickup').value = pickup;
            document.getElementById('addPassengerDropoff').value = dropoff;
            
            // 如果线路不为空，调用identifyLine函数
            if (line) {
                identifyLine();
            }
            
            showNotification('智能识别成功，已填充相关数据', 'success');
        }
        
        function identifyLine() {
            const line = document.getElementById('addPassengerLine').value.trim();
            
            if (!line) {
                showNotification('请输入线路名称', 'error');
                return;
            }
            
            // 从processedData中筛选出该线路的数据
            const lineData = processedData.filter(item => item.线路 === line);
            
            if (lineData.length === 0) {
                showNotification(`未找到线路 "${line}" 的数据`, 'error');
                return;
            }
            
            // 提取该线路已有的车牌号及其人数和座位数信息
            const plates = [...new Set(lineData.map(item => item.车牌号码).filter(Boolean))];
            const plateInfo = [];
            
            plates.forEach(plate => {
                // 计算该车牌的总人数（所有线路的总和）
                const allPlateData = processedData.filter(item => item.车牌号码 === plate);
                const totalPassengers = allPlateData.reduce((sum, item) => sum + (item.票数 || 0), 0);
                
                // 获取座位数（从vehicleCapacityMap中获取）
                let capacity = '未知';
                if (vehicleCapacityMap.has(plate)) {
                    capacity = vehicleCapacityMap.get(plate).容量 || '未知';
                }
                
                // 计算状态
                let status = '';
                if (capacity !== '未知') {
                    const capacityNum = parseInt(capacity);
                    if (totalPassengers >= capacityNum) {
                        status = totalPassengers > capacityNum ? `超载${totalPassengers - capacityNum}人` : '满员';
                    } else {
                        status = `空${capacityNum - totalPassengers}座`;
                    }
                }
                
                plateInfo.push({ plate, totalPassengers, capacity, status });
            });
            
            // 填充车牌号码下拉列表（包含人数、座位数和状态信息）
            const plateSelect = document.getElementById('addPassengerPlate');
            plateSelect.innerHTML = '<option value="">请选择车牌号码（为空则表示未分配）</option>';
            plateInfo.forEach(info => {
                const option = document.createElement('option');
                option.value = info.plate;
                option.textContent = `${info.plate} (${info.totalPassengers}/${info.capacity}, ${info.status})`;
                plateSelect.appendChild(option);
            });
            
            showNotification(`成功识别线路 "${line}"，已填充相关数据`, 'success');
        }
        
        // 新增：关闭添加乘客弹窗
        function closeAddPassengerModal() {
            addPassengerModalOverlay.style.display = 'none';
        }
        
        // 新增：确认添加乘客
        function confirmAddPassenger() {
            const orderId = document.getElementById('addPassengerOrderId').value.trim();
            const line = document.getElementById('addPassengerLine').value.trim();
            const contact = document.getElementById('addPassengerContact').value.trim();
            const phone = document.getElementById('addPassengerPhone').value.trim();
            const ticketCount = parseInt(document.getElementById('addPassengerTicketCount').value) || 1;
            const pickup = document.getElementById('addPassengerPickup').value.trim();
            const dropoff = document.getElementById('addPassengerDropoff').value.trim();
            const date = document.getElementById('addPassengerDate').value.trim();
            const time = document.getElementById('addPassengerTime').value.trim();
            const plate = document.getElementById('addPassengerPlate').value.trim();
            const manager = document.getElementById('addPassengerManager').value.trim();
            const managerPhone = document.getElementById('addPassengerManagerPhone').value.trim();
            const notice = document.getElementById('addPassengerNotice').value.trim();
            
            // 验证必填字段
            if (!orderId || !line || !contact || !phone || !pickup || !dropoff || !date) {
                showNotification('请填写所有必填字段（订单ID、线路、联系人、联系电话、上车地点、下车地点、乘车日期）', 'error');
                return;
            }
            
            // 检查订单ID是否已存在
            const orderExists = processedData.some(item => item.订单ID === orderId);
            if (orderExists) {
                showNotification(`订单ID "${orderId}" 已存在，请使用不同的订单ID`, 'error');
                return;
            }

            // 使用用户输入的随车负责人、随车负责人电话和通知内容
            // 如果用户未输入，则从分配的车辆中获取相同线路的信息
            let finalNotice = notice;
            let finalManager = manager;
            let finalManagerPhone = managerPhone;
            
            if (!finalNotice || !finalManager || !finalManagerPhone) {
                if (plate) {
                    const sameLineVehicleData = processedData.find(item => item.车牌号码 === plate && item.线路 === line);
                    if (sameLineVehicleData) {
                        if (!finalNotice) finalNotice = sameLineVehicleData.通知内容 || '';
                        if (!finalManager) finalManager = sameLineVehicleData.随车负责人 || '';
                        if (!finalManagerPhone) finalManagerPhone = sameLineVehicleData.随车负责人电话 || '';
                    }
                }
            }

            // 创建新的订单对象，增加分销员手机字段（默认为空）
            const newOrder = {
                订单ID: orderId,
                线路: line,
                车牌号码: plate,
                联系人: contact,
                联系电话: phone,
                票数: ticketCount,
                通知内容: finalNotice,
                乘车日期: date,
                乘车时间: time,
                上车地点: pickup,
                下车地点: dropoff,
                随车负责人: finalManager,
                随车负责人电话: finalManagerPhone,
                修改时间: new Date().toLocaleString('zh-CN'),
                分销员手机: '' // 手动添加的乘客默认无分销员手机
            };
            
            // 所有信息均由用户手动输入，不再自动填充

            // 添加到processedData
            processedData.push(newOrder);
            
            // 将新订单添加到手动添加的订单集合中
            manuallyAddedOrders.add(orderId);
            
            // 如果指定了车牌，更新车辆统计
            if (plate && plate.trim()) {
                // 更新车辆总分配人数
                if (vehicleCapacityMap.has(plate)) {
                    const vehicleInfo = vehicleCapacityMap.get(plate);
                    const currentAssigned = vehicleTotalAssigned.get(plate) || 0;
                    vehicleTotalAssigned.set(plate, currentAssigned + ticketCount);
                    vehicleInfo.已分配人数 = currentAssigned + ticketCount;
                } else {
                    // 如果车辆不存在，创建新的车辆信息
                    vehicleCapacityMap.set(plate, {
                        乘车时间: time,
                        容量: null,
                        随车负责人: finalManager,
                        随车负责人电话: finalManagerPhone,
                        通知内容: finalNotice,
                        已分配人数: ticketCount
                    });
                    vehicleTotalAssigned.set(plate, ticketCount);
                }
                
                // 更新车牌线路映射
                if (!plateLineMap.has(plate)) {
                    plateLineMap.set(plate, new Set());
                }
                plateLineMap.get(plate).add(line);
            } else {
                // 未分配车辆，更新线路未分配人数
                if (!lineUnassignedCount[line]) {
                    lineUnassignedCount[line] = 0;
                }
                lineUnassignedCount[line] += ticketCount;
            }
            
            // 重新排序processedData
            processedData.sort((a, b) => {
                const plateA = a.车牌号码 || "";
                const plateB = b.车牌号码 || "";
                const lineA = a.线路;
                const lineB = b.线路;
                
                if (plateA && plateB) {
                    if (plateA !== plateB) {
                        return plateA.localeCompare(plateB);
                    } else {
                        return lineA.localeCompare(lineB);
                    }
                } else if (plateA && !plateB) {
                    return -1;
                } else if (!plateA && plateB) {
                    return 1;
                } else {
                    return lineA.localeCompare(lineB);
                }
            });
            
            closeAddPassengerModal();
            
            // 更新显示
            if (isDataVisible) {
                displayProcessedData();
                displayVehiclePickupStatsPanel();
                displayLineTicketStats();
                displayLineTimeLocationStats();
            }
            
            showNotification(`成功添加乘客: ${contact} (${orderId})`, 'success');
            
            // 自动保存数据到数据库
            saveAllData();
        }
        
        // 新增：显示删除确认对话框
        function showDeleteConfirmDialog() {
            if (selectedRows.size !== 1) {
                showNotification('请只选择一个订单进行删除', 'error');
                return;
            }
            
            const orderId = Array.from(selectedRows)[0];
            const orderData = processedData.find(item => item.订单ID === orderId);
            
            if (!orderData) {
                showNotification('未找到选中的订单数据', 'error');
                return;
            }
            
            deleteConfirmDialogBody.textContent = `确定要删除乘客 "${orderData.联系人}" (订单ID: ${orderId}) 吗？此操作无法撤销。`;
            deleteConfirmDialogOverlay.style.display = 'flex';
        }
        
        // 新增：关闭删除确认对话框
        function closeDeleteConfirmDialog() {
            deleteConfirmDialogOverlay.style.display = 'none';
        }
        
        // 新增：确认删除乘客
        function confirmDeletePassenger() {
            if (selectedRows.size !== 1) {
                showNotification('请只选择一个订单进行删除', 'error');
                return;
            }
            
            const orderId = Array.from(selectedRows)[0];
            const orderIndex = processedData.findIndex(item => item.订单ID === orderId);
            
            if (orderIndex === -1) {
                showNotification('未找到选中的订单数据', 'error');
                closeDeleteConfirmDialog();
                return;
            }
            
            const orderData = processedData[orderIndex];
            const plate = orderData.车牌号码 || '';
            const line = orderData.线路;
            const ticketCount = orderData.票数 || 0;
            
            // 删除订单
            processedData.splice(orderIndex, 1);
            
            // 如果订单有分配车牌，更新车辆统计
            if (plate && plate.trim()) {
                if (vehicleCapacityMap.has(plate)) {
                    const vehicleInfo = vehicleCapacityMap.get(plate);
                    const currentAssigned = vehicleTotalAssigned.get(plate) || 0;
                    const newAssigned = Math.max(0, currentAssigned - ticketCount);
                    vehicleTotalAssigned.set(plate, newAssigned);
                    vehicleInfo.已分配人数 = newAssigned;
                    
                    // 如果车辆已分配人数为0，可以考虑移除车辆信息，但这里保留
                }
                
                // 更新车牌线路映射
                if (plateLineMap.has(plate)) {
                    // 检查是否还有该线路的其他订单
                    const hasOtherOrders = processedData.some(item => 
                        item.车牌号码 === plate && item.线路 === line
                    );
                    
                    if (!hasOtherOrders) {
                        plateLineMap.get(plate).delete(line);
                        if (plateLineMap.get(plate).size === 0) {
                            plateLineMap.delete(plate);
                        }
                    }
                }
            } else {
                // 未分配车辆，更新线路未分配人数
                if (lineUnassignedCount[line]) {
                    lineUnassignedCount[line] = Math.max(0, lineUnassignedCount[line] - ticketCount);
                }
            }
            
            // 从选中行中移除
            selectedRows.delete(orderId);
            
            closeDeleteConfirmDialog();
            closeModifyModal();
            
            // 更新显示
            if (isDataVisible) {
                displayProcessedData();
                displayVehiclePickupStatsPanel();
                displayLineTicketStats();
                displayLineTimeLocationStats();
            }
            
            showNotification(`成功删除乘客: ${orderData.联系人} (${orderId})`, 'success');
            updateModifyButtonState();
            
            // 自动保存数据到数据库
            saveAllData();
        }
        
        // 显示警告对话框
        function showWarningDialog(message) {
            const warningDialogOverlay = document.getElementById('warningDialogOverlay');
            const warningMessage = document.getElementById('warningMessage');
            
            warningMessage.textContent = message;
            warningDialogOverlay.style.display = 'flex';
        }
        
        // 确认修改
        function confirmModify() {
            const modifyTime = document.getElementById('modifyTime') ? document.getElementById('modifyTime').value : null;
            const modifyPickup = document.getElementById('modifyPickup') ? document.getElementById('modifyPickup').value : null;
            const modifyDropoff = document.getElementById('modifyDropoff') ? document.getElementById('modifyDropoff').value : null;
            const modifyPlate = document.getElementById('modifyPlate').value.trim();
            const modifyManager = document.getElementById('modifyManager').value.trim();
            const modifyManagerPhone = document.getElementById('modifyManagerPhone').value.trim();
            const modifyNotice = document.getElementById('modifyNotice') ? document.getElementById('modifyNotice').value.trim() : '';
            
            // 检查是否修改了第一个订单的通知内容、随车负责人或随车负责人电话
            let hasFirstOrderWithVehicle = false;
            const vehicleLineGroups = {};
            
            // 构建车辆线路分组
            processedData.forEach(item => {
                const plate = item.车牌号码 || '未分配车辆';
                const line = item.线路 || '';
                const key = `${plate}-${line}`;
                if (!vehicleLineGroups[key]) vehicleLineGroups[key] = [];
                vehicleLineGroups[key].push(item);
            });
            
            // 检查是否包含第一个订单且修改了通知内容、随车负责人或随车负责人电话，并且该订单已分配车辆且未修改车牌号
            Array.from(selectedRows).forEach(orderId => {
                const orderIndex = processedData.findIndex(item => item.订单ID === orderId);
                if (orderIndex !== -1) {
                    const order = processedData[orderIndex];
                    const plate = order.车牌号码 || '未分配车辆';
                    const line = order.线路 || '';
                    const key = `${plate}-${line}`;
                    
                    // 检查是否修改了车牌号
                    const plateChanged = modifyPlate && modifyPlate !== order.车牌号码;
                    
                    if (vehicleLineGroups[key]) {
                        const group = vehicleLineGroups[key];
                        // 检查是否是该组的第一个订单
                        if (group[0].订单ID === orderId) {
                            // 检查是否修改了通知内容、随车负责人或随车负责人电话，并且该订单已分配车辆且未修改车牌号
                            if ((modifyNotice && modifyNotice !== order.通知内容 || 
                                 modifyManager && modifyManager !== order.随车负责人 || 
                                 modifyManagerPhone && modifyManagerPhone !== order.随车负责人电话) && 
                                order.车牌号码 && order.车牌号码.trim() !== '' && 
                                !plateChanged) {
                                hasFirstOrderWithVehicle = true;
                            }
                        }
                    }
                }
            });
            
            if (hasFirstOrderWithVehicle) {
                showWarningDialog('第一个订单不可以修改通知内容、随车负责人或随车负责人电话，会影响整车信息，请在导出表格后自行修改');
                return;
            }
            
            const existingPlates = new Set();
            processedData.forEach(item => {
                if (item.车牌号码 && item.车牌号码.trim()) {
                    existingPlates.add(item.车牌号码.trim());
                }
            });
            
            let ordersToModify = Array.from(selectedRows);
            
            let updatedCount = 0;
            
            const newlyAssignedOrders = [];
            
            ordersToModify.forEach(orderId => {
                const orderIndex = processedData.findIndex(item => item.订单ID === orderId);
                if (orderIndex !== -1) {
                    const originalOrder = processedData[orderIndex];
                    
                    const oldPlate = originalOrder.车牌号码 || '';
                    const newPlate = modifyPlate;
                    const plateChanged = newPlate && newPlate !== oldPlate;
                    
                    const ticketCount = originalOrder.票数 || 0;
                    
                    const wasUnassigned = !oldPlate || oldPlate.trim() === '';
                    const nowAssigned = newPlate && newPlate.trim() !== '';
                    const becameAssigned = wasUnassigned && nowAssigned && plateChanged;
                    
                    if (newPlate && existingPlates.has(newPlate) && plateChanged) {
                        let foundVehicleInfo = null;
                        
                        if (vehicleCapacityMap.has(newPlate)) {
                            foundVehicleInfo = vehicleCapacityMap.get(newPlate);
                        } else {
                            const existingOrder = processedData.find(item => 
                                item.车牌号码 && item.车牌号码.trim() === newPlate
                            );
                            if (existingOrder) {
                                // 优先查找同车同线路的订单
                                const sameLineOrder = processedData.find(item => 
                                    item.车牌号码 && item.车牌号码.trim() === newPlate &&
                                    item.线路 === originalOrder.线路
                                );
                                
                                if (sameLineOrder) {
                                    foundVehicleInfo = {
                                        乘车时间: sameLineOrder.乘车时间 || "",
                                        随车负责人: sameLineOrder.随车负责人 || "",
                                        随车负责人电话: sameLineOrder.随车负责人电话 || "",
                                        通知内容: sameLineOrder.通知内容 || ""
                                    };
                                } else {
                                    foundVehicleInfo = {
                                        乘车时间: existingOrder.乘车时间 || "",
                                        随车负责人: existingOrder.随车负责人 || "",
                                        随车负责人电话: existingOrder.随车负责人电话 || "",
                                        通知内容: existingOrder.通知内容 || ""
                                    };
                                }
                            }
                        }
                        
                        if (selectedRows.size === 1) {
                            if (modifyTime !== null) processedData[orderIndex].乘车时间 = modifyTime || originalOrder.乘车时间;
                            if (modifyPickup !== null) processedData[orderIndex].上车地点 = modifyPickup || originalOrder.上车地点;
                            if (modifyDropoff !== null) processedData[orderIndex].下车地点 = modifyDropoff || originalOrder.下车地点;
                        } else {
                            if (modifyTime !== null && modifyTime.trim() !== '') processedData[orderIndex].乘车时间 = modifyTime;
                            if (modifyPickup !== null && modifyPickup.trim() !== '') processedData[orderIndex].上车地点 = modifyPickup;
                            if (modifyDropoff !== null && modifyDropoff.trim() !== '') processedData[orderIndex].下车地点 = modifyDropoff;
                        }
                        
                        processedData[orderIndex].车牌号码 = newPlate;
                        
                        // 优先使用用户输入的信息，否则使用同车同线路的信息
                        processedData[orderIndex].随车负责人 = modifyManager || (foundVehicleInfo ? foundVehicleInfo.随车负责人 : originalOrder.随车负责人);
                        processedData[orderIndex].随车负责人电话 = modifyManagerPhone || (foundVehicleInfo ? foundVehicleInfo.随车负责人电话 : originalOrder.随车负责人电话);
                        
                        // 更新修改时间
                        processedData[orderIndex].修改时间 = new Date().toLocaleString('zh-CN');
                        
                        // 确定通知内容：优先使用用户输入，否则使用同车同线路的信息
                        let noticeContent = modifyNotice || (foundVehicleInfo ? foundVehicleInfo.通知内容 : originalOrder.通知内容);
                        
                        // 检查是否选择了本车中全部的订单或本车中相同路线的所有订单
                        const allVehicleOrders = processedData.filter(item => 
                            item.车牌号码 === newPlate
                        );
                        const allLineOrders = processedData.filter(item => 
                            item.车牌号码 === newPlate && 
                            item.线路 === originalOrder.线路
                        );
                        
                        const selectedOrderIds = new Set(selectedRows);
                        const allVehicleSelected = allVehicleOrders.every(item => selectedOrderIds.has(item.订单ID));
                        const allLineSelected = allLineOrders.every(item => selectedOrderIds.has(item.订单ID));
                        
                        // 查找同车同线路的其他订单，保持通知内容一致（仅当选择了全部订单时）
                        if (allVehicleSelected || allLineSelected) {
                            const sameLineOrders = processedData.filter(item => 
                                item.车牌号码 === newPlate && 
                                item.线路 === originalOrder.线路
                            );
                            
                            if (sameLineOrders.length > 0) {
                                const existingNotice = sameLineOrders.find(item => item.通知内容);
                                if (existingNotice && existingNotice.通知内容 && !modifyNotice) {
                                    noticeContent = existingNotice.通知内容;
                                }
                            }
                        }
                        
                        processedData[orderIndex].通知内容 = noticeContent;
                        
                        // 更新车辆总分配人数
                        if (foundVehicleInfo && vehicleCapacityMap.has(newPlate)) {
                            const vehicleInfo = vehicleCapacityMap.get(newPlate);
                            const currentAssigned = vehicleTotalAssigned.get(newPlate) || 0;
                            vehicleTotalAssigned.set(newPlate, currentAssigned + ticketCount);
                            vehicleInfo.已分配人数 = currentAssigned + ticketCount;
                        }
                        
                        if (newPlate) {
                            if (!plateLineMap.has(newPlate)) {
                                plateLineMap.set(newPlate, new Set());
                            }
                            plateLineMap.get(newPlate).add(originalOrder.线路);
                        }
                        
                        if (becameAssigned && originalOrder.线路 && lineUnassignedCount[originalOrder.线路]) {
                            lineUnassignedCount[originalOrder.线路] -= ticketCount;
                            if (lineUnassignedCount[originalOrder.线路] < 0) lineUnassignedCount[originalOrder.线路] = 0;
                            
                            manuallyAssignedOrders.add(orderId);
                            if (!lineManuallyAssignedCount[originalOrder.线路]) {
                                lineManuallyAssignedCount[originalOrder.线路] = 0;
                            }
                            lineManuallyAssignedCount[originalOrder.线路] += ticketCount;
                            newlyAssignedOrders.push({orderId, plate: newPlate, tickets: ticketCount, line: originalOrder.线路});
                        }
                        
                        if (oldPlate && vehicleCapacityMap.has(oldPlate)) {
                            const oldVehicleInfo = vehicleCapacityMap.get(oldPlate);
                            const currentAssigned = vehicleTotalAssigned.get(oldPlate) || 0;
                            const newAssigned = Math.max(0, currentAssigned - ticketCount);
                            vehicleTotalAssigned.set(oldPlate, newAssigned);
                            oldVehicleInfo.已分配人数 = newAssigned;
                            
                            // 检查原车是否还有其他订单
                            const oldVehicleOrders = processedData.filter(item => 
                                item.车牌号码 === oldPlate
                            );
                            
                            // 如果原车还有其他订单，保留原车信息
                            if (oldVehicleOrders.length === 0) {
                                // 如果原车没有其他订单，移除原车信息
                                vehicleCapacityMap.delete(oldPlate);
                            }
                        }
                        
                        if (oldPlate && plateLineMap.has(oldPlate)) {
                            plateLineMap.get(oldPlate).delete(originalOrder.线路);
                            if (plateLineMap.get(oldPlate).size === 0) {
                                plateLineMap.delete(oldPlate);
                            }
                        }
                        
                        // 确保只更新当前订单的信息，不影响原车的其他订单
                        // 这里不需要更新原车的其他订单信息，因为我们只是将当前订单转移到新车
                        
                        showNotification(`订单 ${originalOrder.订单ID} 已归并到车牌号 ${newPlate} 下`, 'info');
                    } else {
                        if (selectedRows.size === 1) {
                            if (modifyTime !== null) processedData[orderIndex].乘车时间 = modifyTime || originalOrder.乘车时间;
                            if (modifyPickup !== null) processedData[orderIndex].上车地点 = modifyPickup || originalOrder.上车地点;
                            if (modifyDropoff !== null) processedData[orderIndex].下车地点 = modifyDropoff || originalOrder.下车地点;
                        } else {
                            if (modifyTime !== null && modifyTime.trim() !== '') processedData[orderIndex].乘车时间 = modifyTime;
                            if (modifyPickup !== null && modifyPickup.trim() !== '') processedData[orderIndex].上车地点 = modifyPickup;
                            if (modifyDropoff !== null && modifyDropoff.trim() !== '') processedData[orderIndex].下车地点 = modifyDropoff;
                        }
                        
                        if (newPlate && plateChanged) {
                            processedData[orderIndex].车牌号码 = newPlate;
                            
                            // 检查是否选择了本车中全部的订单或本车中相同路线的所有订单
                            const allVehicleOrders = processedData.filter(item => 
                                item.车牌号码 === newPlate
                            );
                            const allLineOrders = processedData.filter(item => 
                                item.车牌号码 === newPlate && 
                                item.线路 === originalOrder.线路
                            );
                            
                            const selectedOrderIds = new Set(selectedRows);
                            const allVehicleSelected = allVehicleOrders.every(item => selectedOrderIds.has(item.订单ID));
                            const allLineSelected = allLineOrders.every(item => selectedOrderIds.has(item.订单ID));
                            
                            // 查找同车同线路、同时间、同上下车地点的其他订单，保持通知内容一致（仅当选择了全部订单时）
                            let noticeContent = modifyNotice || originalOrder.通知内容;
                            
                            if (allVehicleSelected || allLineSelected) {
                                const sameVehicleOrders = processedData.filter(item => 
                                    item.车牌号码 === newPlate && 
                                    item.线路 === originalOrder.线路 && 
                                    item.乘车时间 === (modifyTime || originalOrder.乘车时间) && 
                                    item.上车地点 === (modifyPickup || originalOrder.上车地点) && 
                                    item.下车地点 === (modifyDropoff || originalOrder.下车地点)
                                );
                                
                                if (sameVehicleOrders.length > 0) {
                                    const existingNotice = sameVehicleOrders.find(item => item.通知内容);
                                    if (existingNotice && existingNotice.通知内容) {
                                        noticeContent = existingNotice.通知内容;
                                    }
                                }
                            }
                            
                            // 只有当选择了本车中全部的订单或本车中相同路线的所有订单时，才更新车辆信息
                            if (allVehicleSelected || allLineSelected) {
                                if (vehicleCapacityMap.has(newPlate)) {
                                    const vehicleInfo = vehicleCapacityMap.get(newPlate);
                                    const currentAssigned = vehicleTotalAssigned.get(newPlate) || 0;
                                    vehicleTotalAssigned.set(newPlate, currentAssigned + ticketCount);
                                    vehicleInfo.已分配人数 = currentAssigned + ticketCount;
                                    
                                    // 更新车辆信息
                                    if (modifyTime) vehicleInfo.乘车时间 = modifyTime;
                                    if (modifyManager) vehicleInfo.随车负责人 = modifyManager;
                                    if (modifyManagerPhone) vehicleInfo.随车负责人电话 = modifyManagerPhone;
                                    if (noticeContent) vehicleInfo.通知内容 = noticeContent;
                                } else {
                                    vehicleCapacityMap.set(newPlate, {
                                        乘车时间: modifyTime || "",
                                        容量: null,
                                        随车负责人: modifyManager || "",
                                        随车负责人电话: modifyManagerPhone || "",
                                        通知内容: noticeContent || "",
                                        已分配人数: ticketCount
                                    });
                                    vehicleTotalAssigned.set(newPlate, ticketCount);
                                }
                            } else {
                                // 只更新已分配人数
                                if (vehicleCapacityMap.has(newPlate)) {
                                    const vehicleInfo = vehicleCapacityMap.get(newPlate);
                                    const currentAssigned = vehicleTotalAssigned.get(newPlate) || 0;
                                    vehicleTotalAssigned.set(newPlate, currentAssigned + ticketCount);
                                    vehicleInfo.已分配人数 = currentAssigned + ticketCount;
                                } else {
                                    // 创建基本车辆信息，但不包含随车负责人等详细信息
                                    vehicleCapacityMap.set(newPlate, {
                                        乘车时间: "",
                                        容量: null,
                                        随车负责人: "",
                                        随车负责人电话: "",
                                        通知内容: "",
                                        已分配人数: ticketCount
                                    });
                                    vehicleTotalAssigned.set(newPlate, ticketCount);
                                }
                            }
                            
                            if (newPlate) {
                                if (!plateLineMap.has(newPlate)) {
                                    plateLineMap.set(newPlate, new Set());
                                }
                                plateLineMap.get(newPlate).add(originalOrder.线路);
                            }
                            
                            existingPlates.add(newPlate);
                            
                            if (becameAssigned && originalOrder.线路 && lineUnassignedCount[originalOrder.线路]) {
                                lineUnassignedCount[originalOrder.线路] -= ticketCount;
                                if (lineUnassignedCount[originalOrder.线路] < 0) lineUnassignedCount[originalOrder.线路] = 0;
                                
                                manuallyAssignedOrders.add(orderId);
                                if (!lineManuallyAssignedCount[originalOrder.线路]) {
                                    lineManuallyAssignedCount[originalOrder.线路] = 0;
                                }
                                lineManuallyAssignedCount[originalOrder.线路] += ticketCount;
                                newlyAssignedOrders.push({orderId, plate: newPlate, tickets: ticketCount, line: originalOrder.线路});
                            }
                            
                            if (oldPlate && vehicleCapacityMap.has(oldPlate)) {
                                const oldVehicleInfo = vehicleCapacityMap.get(oldPlate);
                                const currentAssigned = vehicleTotalAssigned.get(oldPlate) || 0;
                                const newAssigned = Math.max(0, currentAssigned - ticketCount);
                                vehicleTotalAssigned.set(oldPlate, newAssigned);
                                oldVehicleInfo.已分配人数 = newAssigned;
                                
                                // 检查原车是否还有其他订单
                                const oldVehicleOrders = processedData.filter(item => 
                                    item.车牌号码 === oldPlate
                                );
                                
                                // 如果原车还有其他订单，保留原车信息
                                if (oldVehicleOrders.length === 0) {
                                    // 如果原车没有其他订单，移除原车信息
                                    vehicleCapacityMap.delete(oldPlate);
                                }
                            }
                            
                            if (oldPlate && plateLineMap.has(oldPlate)) {
                                plateLineMap.get(oldPlate).delete(originalOrder.线路);
                                if (plateLineMap.get(oldPlate).size === 0) {
                                    plateLineMap.delete(oldPlate);
                                }
                            }
                            
                            // 更新通知内容
                            processedData[orderIndex].通知内容 = noticeContent;
                        } else {
                            processedData[orderIndex].车牌号码 = newPlate || originalOrder.车牌号码;
                        }
                        
                        processedData[orderIndex].随车负责人 = modifyManager || originalOrder.随车负责人;
                        processedData[orderIndex].随车负责人电话 = modifyManagerPhone || originalOrder.随车负责人电话;
                        processedData[orderIndex].通知内容 = modifyNotice || originalOrder.通知内容;
                    }
                    
                    // 保存修改后的信息到manuallyModifiedOrders集合中
                    const modifiedOrder = processedData[orderIndex];
                    const originalOrderFromData = originalData.find(item => item.订单ID === orderId);
                    
                    if (originalOrderFromData) {
                        // 计算修改的字段
                        const modifiedFields = {};
                        
                        if (modifiedOrder.乘车时间 !== originalOrderFromData.乘车时间) {
                            modifiedFields.乘车时间 = modifiedOrder.乘车时间;
                        }
                        if (modifiedOrder.上车地点 !== originalOrderFromData.上车地点) {
                            modifiedFields.上车地点 = modifiedOrder.上车地点;
                        }
                        if (modifiedOrder.下车地点 !== originalOrderFromData.下车地点) {
                            modifiedFields.下车地点 = modifiedOrder.下车地点;
                        }
                        if (modifiedOrder.车牌号码 !== originalOrderFromData.车牌号码) {
                            modifiedFields.车牌号码 = modifiedOrder.车牌号码;
                        }
                        if (modifiedOrder.随车负责人 !== originalOrderFromData.随车负责人) {
                            modifiedFields.随车负责人 = modifiedOrder.随车负责人;
                        }
                        if (modifiedOrder.随车负责人电话 !== originalOrderFromData.随车负责人电话) {
                            modifiedFields.随车负责人电话 = modifiedOrder.随车负责人电话;
                        }
                        if (modifiedOrder.通知内容 !== originalOrderFromData.通知内容) {
                            modifiedFields.通知内容 = modifiedOrder.通知内容;
                        }
                        
                        // 如果有修改的字段，保存到manuallyModifiedOrders中
                        if (Object.keys(modifiedFields).length > 0) {
                            manuallyModifiedOrders.set(orderId, modifiedFields);
                        } else {
                            // 如果没有修改的字段，从manuallyModifiedOrders中移除
                            manuallyModifiedOrders.delete(orderId);
                        }
                    }
                    
                    updatedCount++;
                }
            });
            
            if (isSearchActive) {
                ordersToModify.forEach(orderId => {
                    const orderIndex = filteredData.findIndex(item => item.订单ID === orderId);
                    if (orderIndex !== -1) {
                        if (selectedRows.size === 1) {
                            if (modifyTime !== null) filteredData[orderIndex].乘车时间 = modifyTime || filteredData[orderIndex].乘车时间;
                            if (modifyPickup !== null) filteredData[orderIndex].上车地点 = modifyPickup || filteredData[orderIndex].上车地点;
                            if (modifyDropoff !== null) filteredData[orderIndex].下车地点 = modifyDropoff || filteredData[orderIndex].下车地点;
                        } else {
                            if (modifyTime !== null && modifyTime.trim() !== '') filteredData[orderIndex].乘车时间 = modifyTime;
                            if (modifyPickup !== null && modifyPickup.trim() !== '') filteredData[orderIndex].上车地点 = modifyPickup;
                            if (modifyDropoff !== null && modifyDropoff.trim() !== '') filteredData[orderIndex].下车地点 = modifyDropoff;
                        }
                        
                        filteredData[orderIndex].车牌号码 = modifyPlate || filteredData[orderIndex].车牌号码;
                        filteredData[orderIndex].随车负责人 = modifyManager || filteredData[orderIndex].随车负责人;
                        filteredData[orderIndex].随车负责人电话 = modifyManagerPhone || filteredData[orderIndex].随车负责人电话;
                        filteredData[orderIndex].通知内容 = modifyNotice || filteredData[orderIndex].通知内容;
                    }
                });
            }
            
            closeModifyModal();
            
            if (newlyAssignedOrders.length > 0) {
                let totalTickets = 0;
                const plates = new Set();
                newlyAssignedOrders.forEach(item => {
                    totalTickets += item.tickets;
                    plates.add(item.plate);
                });
                
                const platesText = Array.from(plates).join('、');
                showNotification(`已将 ${totalTickets} 人手动分配到车辆 ${platesText}，请查看车辆统计面板`, 'info');
            }
            
            if (isDataVisible) {
                displayProcessedData();
                displayVehiclePickupStatsPanel();
                displayLineTicketStats();
                displayLineTimeLocationStats();
                
                ordersToModify.forEach(orderId => {
                    const row = document.querySelector(`tr[data-order-id="${orderId}"]`);
                    if (row) {
                        row.classList.add('highlight-row');
                        setTimeout(() => {
                            row.classList.remove('highlight-row');
                        }, 1500);
                    }
                });
            }
            
            showNotification(`成功修改 ${updatedCount} 个订单的信息`, 'success');
            
            selectedRows.clear();
            updateModifyButtonState();
            
            // 自动保存数据到数据库
            saveAllData();
        }
        
        // 显示编辑车辆模态框
        function showEditVehicleModal(plate) {
            if (!plate || plate === '未分配车辆') {
                showNotification('无法编辑未分配车辆', 'error');
                return;
            }
            
            const vehicleInfo = vehicleCapacityMap.get(plate) || {
                顺序: null,
                容量: null,
                乘车时间: '',
                随车负责人: '',
                随车负责人电话: '',
                通知内容: ''
            };
            
            // 获取该车辆包含的所有线路
            const lines = new Set();
            processedData.forEach(item => {
                if (item.车牌号码 === plate && item.线路) {
                    lines.add(item.线路);
                }
            });
            const lineArray = Array.from(lines);
            
            let formHTML = `
                <div class="form-group">
                    <label>原车牌号</label>
                    <div class="readonly-field">${plate}</div>
                </div>
                <div class="form-group">
                    <label>新车牌号</label>
                    <input type="text" id="newPlate" class="input-field" value="${plate}" placeholder="请输入新车牌号">
                </div>
                <div class="form-group">
                    <label>顺序</label>
                    <input type="number" id="editOrder" class="input-field" value="${vehicleInfo.顺序 || ''}" placeholder="请输入顺序，留空表示不设置">
                </div>
                <div class="form-group">
                    <label>座位数</label>
                    <input type="number" id="editCapacity" class="input-field" value="${vehicleInfo.容量 || ''}" placeholder="请输入座位数，留空表示不限制">
                </div>
                <div class="form-group">
                    <label>随车负责人</label>
                    <input type="text" id="editManager" class="input-field" value="${vehicleInfo.随车负责人 || ''}" placeholder="请输入随车负责人">
                </div>
                <div class="form-group">
                    <label>随车负责人电话</label>
                    <input type="text" id="editManagerPhone" class="input-field" value="${vehicleInfo.随车负责人电话 || ''}" placeholder="请输入随车负责人电话">
                </div>
            `;
            
            // 如果车辆包含多条路线，为每个线路添加通知内容修改框
            if (lineArray.length > 1) {
                lineArray.forEach(line => {
                    // 获取该线路的通知内容
                    let lineNotice = '';
                    const lineOrder = processedData.find(item => item.车牌号码 === plate && item.线路 === line);
                    if (lineOrder && lineOrder.通知内容) {
                        lineNotice = lineOrder.通知内容;
                    }
                    
                    formHTML += `
                        <div class="form-group">
                            <label>通知内容（${line}）</label>
                            <textarea id="editNotice_${line}" class="input-field" rows="3" placeholder="请输入${line}线路的通知内容">${lineNotice || ''}</textarea>
                        </div>
                    `;
                });
            } else if (lineArray.length === 1) {
                // 如果只有一条路线，显示单个通知内容修改框
                let notice = vehicleInfo.通知内容;
                const lineOrder = processedData.find(item => item.车牌号码 === plate && item.线路 === lineArray[0]);
                if (lineOrder && lineOrder.通知内容) {
                    notice = lineOrder.通知内容;
                }
                
                formHTML += `
                    <div class="form-group">
                        <label>通知内容</label>
                        <textarea id="editNotice" class="input-field" rows="3" placeholder="请输入通知内容">${notice || ''}</textarea>
                    </div>
                `;
            }
            
            document.getElementById('editVehicleForm').innerHTML = formHTML;
            document.getElementById('editVehicleModalOverlay').style.display = 'flex';
            
            // 保存原车牌号和线路信息，用于后续更新
            document.getElementById('saveEditVehicleBtn').setAttribute('data-original-plate', plate);
            document.getElementById('saveEditVehicleBtn').setAttribute('data-lines', lineArray.join(','));
        }
        
        // 关闭编辑车辆模态框
        function closeEditVehicleModal() {
            document.getElementById('editVehicleModalOverlay').style.display = 'none';
        }
        
        // 保存车辆信息
        function saveEditVehicleInfo() {
            const originalPlate = document.getElementById('saveEditVehicleBtn').getAttribute('data-original-plate');
            const newPlate = document.getElementById('newPlate').value.trim();
            const order = document.getElementById('editOrder').value.trim() ? parseInt(document.getElementById('editOrder').value) : null;
            const capacity = document.getElementById('editCapacity').value.trim() ? parseInt(document.getElementById('editCapacity').value) : null;
            const manager = document.getElementById('editManager').value.trim();
            const managerPhone = document.getElementById('editManagerPhone').value.trim();
            const lines = document.getElementById('saveEditVehicleBtn').getAttribute('data-lines').split(',');
            
            if (!newPlate) {
                showNotification('请输入车牌号', 'error');
                return;
            }
            
            // 更新vehicleCapacityMap
            const vehicleInfo = {
                顺序: order,
                容量: capacity,
                乘车时间: '', // 不再更新乘车时间
                随车负责人: manager,
                随车负责人电话: managerPhone,
                通知内容: '' // 通知内容按线路单独更新
            };
            
            // 如果车牌号有变更，只更新当前订单的车牌号
            if (originalPlate !== newPlate) {
                // 只更新当前订单的车牌号
                processedData[orderIndex].车牌号码 = newPlate;
                
                // 检查原车是否还有其他订单
                const oldVehicleOrders = processedData.filter(item => 
                    item.车牌号码 === originalPlate
                );
                
                // 如果原车没有其他订单，从vehicleCapacityMap中删除原车牌号
                if (oldVehicleOrders.length === 0) {
                    vehicleCapacityMap.delete(originalPlate);
                    // 从plateLineMap中删除原车牌号
                    plateLineMap.delete(originalPlate);
                    // 从vehicleTotalAssigned中删除原车牌号
                    vehicleTotalAssigned.delete(originalPlate);
                }
            }
            
            // 添加或更新新车牌号的信息
            vehicleCapacityMap.set(newPlate, vehicleInfo);
            
            // 更新所有使用该车辆的订单信息
            processedData.forEach(item => {
                if (item.车牌号码 === newPlate) {
                    if (manager) item.随车负责人 = manager;
                    if (managerPhone) item.随车负责人电话 = managerPhone;
                }
            });
            
            // 按线路更新通知内容
            lines.forEach(line => {
                if (lines.length > 1) {
                    // 多条线路，按线路更新通知内容
                    const noticeElement = document.getElementById(`editNotice_${line}`);
                    if (noticeElement) {
                        const notice = noticeElement.value.trim();
                        if (notice) {
                            processedData.forEach(item => {
                                if (item.车牌号码 === newPlate && item.线路 === line) {
                                    item.通知内容 = notice;
                                }
                            });
                        }
                    }
                } else if (lines.length === 1) {
                    // 单条线路，更新通知内容
                    const noticeElement = document.getElementById('editNotice');
                    if (noticeElement) {
                        const notice = noticeElement.value.trim();
                        if (notice) {
                            processedData.forEach(item => {
                                if (item.车牌号码 === newPlate) {
                                    item.通知内容 = notice;
                                }
                            });
                        }
                    }
                }
            });
            
            // 如果更新了负责人电话，需要重新分配相关订单
            if (managerPhone) {
                // 收集所有联系电话匹配新负责人电话的订单
                const managerOrders = processedData.filter(item => item.联系电话 === managerPhone);
                
                // 将这些订单重新分配到当前车辆
                managerOrders.forEach(order => {
                    order.车牌号码 = newPlate;
                    order.随车负责人 = manager;
                    order.随车负责人电话 = managerPhone;
                    
                    // 查找同车同线路、同时间、同上下车地点的其他订单，保持通知内容一致
                    const sameVehicleOrders = processedData.filter(item => 
                        item.车牌号码 === newPlate && 
                        item.线路 === order.线路 && 
                        item.乘车时间 === order.乘车时间 && 
                        item.出发地点 === order.出发地点 && 
                        item.目的地 === order.目的地
                    );
                    
                    // 如果有其他订单，使用它们的通知内容
                    if (sameVehicleOrders.length > 0) {
                        const existingNotice = sameVehicleOrders.find(item => item.通知内容);
                        if (existingNotice && existingNotice.通知内容) {
                            order.通知内容 = existingNotice.通知内容;
                        }
                    }
                });
            }
            
            // 更新lineInfo对象，确保在为各线路填写车辆信息框中也能看到更新后的车辆信息
            lines.forEach(line => {
                if (lineInfo[line]) {
                    lineInfo[line].forEach(vehicle => {
                        if (vehicle.车牌号码 === originalPlate) {
                            // 如果车牌号有变更，创建新的车辆信息，而不是修改原车辆信息
                            if (originalPlate !== newPlate) {
                                // 移除原车辆信息
                                lineInfo[line] = lineInfo[line].filter(v => v.车牌号码 !== originalPlate);
                                // 添加新车辆信息
                                lineInfo[line].push({
                                    ...vehicle,
                                    车牌号码: newPlate,
                                    顺序: order,
                                    容量: capacity,
                                    随车负责人: manager,
                                    随车负责人电话: managerPhone
                                });
                            } else {
                                // 如果车牌号没有变更，更新原车辆信息
                                vehicle.顺序 = order;
                                vehicle.容量 = capacity;
                                vehicle.随车负责人 = manager;
                                vehicle.随车负责人电话 = managerPhone;
                            }
                        }
                    });
                }
                
                // 同时更新lastLineInfo，确保下次打开为各线路填写车辆信息框时能看到更新后的车辆信息
                if (lastLineInfo[line]) {
                    lastLineInfo[line].forEach(vehicle => {
                        if (vehicle.车牌号码 === originalPlate) {
                            // 如果车牌号有变更，创建新的车辆信息，而不是修改原车辆信息
                            if (originalPlate !== newPlate) {
                                // 移除原车辆信息
                                lastLineInfo[line] = lastLineInfo[line].filter(v => v.车牌号码 !== originalPlate);
                                // 添加新车辆信息
                                lastLineInfo[line].push({
                                    ...vehicle,
                                    车牌号码: newPlate,
                                    顺序: order,
                                    容量: capacity,
                                    随车负责人: manager,
                                    随车负责人电话: managerPhone
                                });
                            } else {
                                // 如果车牌号没有变更，更新原车辆信息
                                vehicle.顺序 = order;
                                vehicle.容量 = capacity;
                                vehicle.随车负责人 = manager;
                                vehicle.随车负责人电话 = managerPhone;
                            }
                        }
                    });
                }
            });
            
            // 同步更新为各线路填写车辆信息中的顺序字段
            const lineRows = document.querySelectorAll('#lineInputsBody tr');
            lineRows.forEach((row, rowIndex) => {
                const plateInput = row.querySelector('.plate-input');
                const orderInput = row.querySelector('.order-input');
                
                if (plateInput && orderInput) {
                    const plate = plateInput.value.trim();
                    if (plate === newPlate) {
                        orderInput.value = order !== null ? order : '';
                    }
                }
            });
            
            // 重新计算plateLineMap和vehicleTotalAssigned
            updatePlateLineMap();
            updateVehicleTotalAssigned();
            
            closeEditVehicleModal();
            
            // 更新显示
            displayProcessedData();
            displayVehiclePickupStatsPanel();
            displayLineTicketStats();
            
            showNotification(`车辆信息已更新：${originalPlate} → ${newPlate}`, 'success');
            
            // 自动保存数据到数据库
            saveAllData();
        }
        
        // 更新plateLineMap
        function updatePlateLineMap() {
            plateLineMap.clear();
            processedData.forEach(item => {
                if (item.车牌号码 && item.线路) {
                    if (!plateLineMap.has(item.车牌号码)) {
                        plateLineMap.set(item.车牌号码, new Set());
                    }
                    plateLineMap.get(item.车牌号码).add(item.线路);
                }
            });
        }
        
        // 更新vehicleTotalAssigned
        function updateVehicleTotalAssigned() {
            vehicleTotalAssigned.clear();
            processedData.forEach(item => {
                if (item.车牌号码) {
                    const currentCount = vehicleTotalAssigned.get(item.车牌号码) || 0;
                    vehicleTotalAssigned.set(item.车牌号码, currentCount + (item.票数 || 0));
                }
            });
        }
        


        // 页面加载完成后初始化
        document.addEventListener('DOMContentLoaded', function() {
            // 初始化数据库
            initDatabase().then(() => {
                showNotification('系统已就绪，请上传Excel文件或点击"刷新"按钮生成示例数据', 'info');
                
                // 尝试加载已保存的数据
                loadAllData().then(() => {
                    if (processedData.length > 0) {
                        // 数据加载成功，更新UI
                        analyzeBtn.disabled = false;
                        toggleBtn.disabled = false;
                        exportBtn.disabled = false;
                        ticketStatsSection.style.display = 'block';
                        displayLineTicketStats();
                        displayLineTimeLocationStats();
                    }
                });
            }).catch(error => {
                console.error('数据库初始化失败:', error);
                showNotification('数据库初始化失败', 'error');
                showNotification('系统已就绪，请上传Excel文件或点击"刷新"按钮生成示例数据', 'info');
            });
            
            analyzeBtn.disabled = true;
            
            // 编辑车辆模态框事件
            document.getElementById('closeEditVehicleModal').addEventListener('click', closeEditVehicleModal);
            document.getElementById('cancelEditVehicleBtn').addEventListener('click', closeEditVehicleModal);
            document.getElementById('saveEditVehicleBtn').addEventListener('click', saveEditVehicleInfo);
            
            // 数据库操作按钮事件
            document.getElementById('saveDataBtn').addEventListener('click', saveAllData);
            document.getElementById('loadDataBtn').addEventListener('click', loadAllData);
        });