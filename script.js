import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-app.js";
import { getFirestore, collection, getDocs, doc, setDoc, addDoc, deleteDoc, getDoc } from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

// ====== ⭐️ Firebaseの設定（ここを書き換えてください） ======
const firebaseConfig = {
    apiKey: "AIzaSyBjXSoRLe3FQd1hmSRkuUoC3NZtsOa3i8w",
    authDomain: "timecardapp-e9505.firebaseapp.com",
    projectId: "timecardapp-e9505",
    storageBucket: "timecardapp-e9505.firebasestorage.app",
    messagingSenderId: "84945572712",
    appId: "1:84945572712:web:a9851b7db6cd7d653dbcdd",
    measurementId: "G-2GYDTL0ZMZ"
};
// =========================================================

let app, db;
try {
    app = initializeApp(firebaseConfig);
    db = getFirestore(app);
} catch (e) {
    console.error("Firebase Initialization Error:", e);
}

document.addEventListener('DOMContentLoaded', () => {
    /* --- State --- */
    let users = [];
    let adminPassword = '0000';
    let isAdminUnlocked = false;
    let history = [];
    let currentSessions = {};
    let currentUserForPunch = null;
    let deletedUserIds = new Set();
    const loadingOverlay = document.getElementById('loadingOverlay');

    function showLoading() { if (loadingOverlay) loadingOverlay.style.display = 'flex'; }
    function hideLoading() { if (loadingOverlay) loadingOverlay.style.display = 'none'; }

    // load db data
    async function loadServerData() {
        showLoading();
        if (!db) {
            console.warn("DB is not initialized. Skipping Firebase loading.");
            setTimeout(hideLoading, 500); // 失敗時でも少しだけローディングを見せる
            return;
        }
        try {
            const confSnap = await getDoc(doc(db, "settings", "config"));
            if (confSnap.exists()) { adminPassword = confSnap.data().adminPassword || '0000'; }

            const usersSnap = await getDocs(collection(db, "users"));
            users = [];
            usersSnap.forEach(s => users.push({ id: s.id, ...s.data() }));

            // 初期ユーザーの自動作成を削除しました

            const histSnap = await getDocs(collection(db, "history"));
            history = [];
            histSnap.forEach(s => history.push({ docId: s.id, ...s.data() }));

            const sessSnap = await getDocs(collection(db, "sessions"));
            currentSessions = {};
            sessSnap.forEach(s => currentSessions[s.id] = s.data());
        } catch (e) {
            console.error("Firebase load error:", e);
            alert("データ読込エラー。Firebaseの設定(firebaseConfig)とルールを確認して下さい。");
        }
        hideLoading();
    }

    /* --- Elements --- */
    const currentDateEl = document.getElementById('currentDate');
    const currentTimeEl = document.getElementById('currentTime');

    // Login & Punch
    const userLoginSection = document.getElementById('userLoginSection');
    const userPunchPin = document.getElementById('userPunchPin');
    const userLoginBtn = document.getElementById('userLoginBtn');

    const loggedInSection = document.getElementById('loggedInSection');
    const loggedInUserName = document.getElementById('loggedInUserName');
    const userLogoutBtn = document.getElementById('userLogoutBtn');

    // Change Password
    const changePwdBtn = document.getElementById('changePwdBtn');
    const changePasswordModal = document.getElementById('changePasswordModal');
    const closeChangePasswordModalBtn = document.getElementById('closeChangePasswordModalBtn');
    const newPasswordInput = document.getElementById('newPassword');
    const newPasswordConfirmInput = document.getElementById('newPasswordConfirm');
    const changePasswordSubmitBtn = document.getElementById('changePasswordSubmitBtn');

    const punchInBtn = document.getElementById('punchInBtn');
    const punchOutBtn = document.getElementById('punchOutBtn');
    const statusMessageEl = document.getElementById('statusMessage');

    // History
    const historyBody = document.getElementById('historyBody');
    const clearHistoryBtn = document.getElementById('clearHistoryBtn');

    // Admin Auth
    const adminLockBtn = document.getElementById('adminLockBtn');
    const authModal = document.getElementById('authModal');
    const closeAuthModalBtn = document.getElementById('closeAuthModalBtn');
    const authPasswordInput = document.getElementById('authPassword');
    const authSubmitBtn = document.getElementById('authSubmitBtn');
    const adminOnlyElements = document.querySelectorAll('.admin-only');

    // Help Modals
    const helpModal = document.getElementById('helpModal');
    const closeHelpModalBtn = document.getElementById('closeHelpModalBtn');
    const userHelpBtn = document.getElementById('userHelpBtn');
    const helpTabBtns = document.querySelectorAll('.help-tab-btn');
    const helpContents = document.querySelectorAll('.help-content');

    const settingsHelpModal = document.getElementById('settingsHelpModal');
    const closeSettingsHelpModalBtn = document.getElementById('closeSettingsHelpModalBtn');
    const adminHelpBtn = document.getElementById('adminHelpBtn');

    // Settings Modal
    const settingsModal = document.getElementById('settingsModal');
    const settingsBtn = document.getElementById('settingsBtn');
    const closeModalBtn = document.getElementById('closeModalBtn');
    const adminPasswordSetInput = document.getElementById('adminPasswordSet');
    const settingsUsersBody = document.getElementById('settingsUsersBody');
    const addUserBtn = document.getElementById('addUserBtn');
    const saveSettingsBtn = document.getElementById('saveSettingsBtn');
    const userCountBadge = document.getElementById('userCountBadge');

    // Aggregation
    const startDateInput = document.getElementById('startDate');
    const endDateInput = document.getElementById('endDate');
    const calculateBtn = document.getElementById('calculateBtn');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const aggregateResults = document.getElementById('aggregateResults');
    const resDays = document.getElementById('resDays');
    const resHolidays = document.getElementById('resHolidays');
    const resTotalHours = document.getElementById('resTotalHours');
    const resOvertime = document.getElementById('resOvertime');
    const resAllowance = document.getElementById('resAllowance');
    const userSelectAggregate = document.getElementById('userSelectAggregate');
    const aggregateResultsContainer = document.getElementById('aggregateResultsContainer');

    /* --- Initialization --- */
    async function init() {
        updateClock();
        setInterval(updateClock, 1000);
        await loadServerData();
        applyAdminLockState();
        renderAggregateUserDropdown();
        renderHistory();
        checkLoginState();

        // Default date range
        const today = new Date();
        const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
        startDateInput.value = formatDateForInput(startOfMonth);
        endDateInput.value = formatDateForInput(today);
    }

    function getUserById(id) {
        return users.find(u => u.id === id);
    }

    /* --- Clock --- */
    function updateClock() {
        const now = new Date();
        const days = ['日', '月', '火', '水', '木', '金', '土'];
        currentDateEl.textContent = `${now.getFullYear()}年${String(now.getMonth() + 1).padStart(2, '0')}月${String(now.getDate()).padStart(2, '0')}日 (${days[now.getDay()]})`;
        currentTimeEl.textContent = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`;
    }

    function formatDateForInput(date) {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    }

    /* --- User Login Logic --- */
    userLoginBtn.addEventListener('click', () => {
        const pin = userPunchPin.value.trim();
        if (!pin) {
            alert('パスワードを入力してください。');
            return;
        }

        const user = users.find(u => {
            if (u.password && u.password !== "") {
                return u.password === pin;
            } else {
                return u.phone === pin;
            }
        });

        if (user) {
            currentUserForPunch = user.id;
            userPunchPin.value = '';
            checkLoginState();
        } else {
            alert('入力されたパスワード（または初期パスワードとしての電話番号）が間違っています。');
        }
    });

    /* --- Password Change Logic --- */
    if (changePwdBtn) {
        changePwdBtn.addEventListener('click', () => {
            newPasswordInput.value = '';
            newPasswordConfirmInput.value = '';
            changePasswordModal.classList.add('show');
        });
    }

    if (closeChangePasswordModalBtn) {
        closeChangePasswordModalBtn.addEventListener('click', () => {
            changePasswordModal.classList.remove('show');
        });
    }

    if (changePasswordSubmitBtn) {
        changePasswordSubmitBtn.addEventListener('click', async () => {
            const np = newPasswordInput.value.trim();
            const npc = newPasswordConfirmInput.value.trim();

            if (!np) {
                alert('新しいパスワードを入力してください。');
                return;
            }
            if (np !== npc) {
                alert('確認用パスワードが一致しません。');
                return;
            }

            // 重複チェック
            const isDuplicate = users.some(u => {
                if (u.id === currentUserForPunch) return false;
                if (u.password === np) return true;
                if (!u.password && u.phone === np) return true;
                return false;
            });

            if (isDuplicate) {
                alert('このパスワードは他のユーザーが使用しているため設定できません。');
                return;
            }

            const user = getUserById(currentUserForPunch);
            if (!user) return;

            user.password = np;

            showLoading();
            try {
                if (db) await setDoc(doc(db, "users", user.id), user);
                alert('パスワードを変更しました。');
                changePasswordModal.classList.remove('show');
            } catch (e) {
                console.error(e);
                alert("パスワードの保存に失敗しました");
            }
            hideLoading();
        });
    }

    userPunchPin.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            userLoginBtn.click();
        }
    });

    userLogoutBtn.addEventListener('click', () => {
        logoutUser();
    });

    function logoutUser() {
        currentUserForPunch = null;
        checkLoginState();
    }

    function checkLoginState() {
        if (currentUserForPunch) {
            const user = getUserById(currentUserForPunch);
            if (user) {
                userLoginSection.style.display = 'none';
                loggedInSection.style.display = 'block';
                loggedInUserName.textContent = user.name;
                updatePunchButtons();
            } else {
                logoutUser();
            }
        } else {
            userLoginSection.style.display = 'flex';
            loggedInSection.style.display = 'none';
            loggedInUserName.textContent = '';
        }
    }

    function updatePunchButtons() {
        if (!currentUserForPunch) return;

        if (currentSessions[currentUserForPunch]) {
            punchInBtn.disabled = true;
            punchOutBtn.disabled = false;
        } else {
            punchInBtn.disabled = false;
            punchOutBtn.disabled = true;
        }
    }

    /* --- Admin Auth Logic --- */
    function applyAdminLockState() {
        if (isAdminUnlocked) {
            adminLockBtn.textContent = '🔓';
            adminLockBtn.setAttribute('aria-label', '管理者ロックをかける');
            adminOnlyElements.forEach(el => el.style.display = '');
        } else {
            adminLockBtn.textContent = '🔒';
            adminLockBtn.setAttribute('aria-label', '管理者ロックを解除');
            adminOnlyElements.forEach(el => el.style.display = 'none');
        }
    }

    adminLockBtn.addEventListener('click', () => {
        if (isAdminUnlocked) {
            isAdminUnlocked = false;
            applyAdminLockState();
        } else {
            authPasswordInput.value = '';
            authModal.classList.add('show');
        }
    });

    closeAuthModalBtn.addEventListener('click', () => {
        authModal.classList.remove('show');
    });

    authSubmitBtn.addEventListener('click', () => {
        if (authPasswordInput.value === adminPassword) {
            isAdminUnlocked = true;
            applyAdminLockState();
            authModal.classList.remove('show');
        } else {
            alert('パスワードが間違っています。');
        }
    });

    authPasswordInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') authSubmitBtn.click();
    });

    /* --- Settings Modal (Multi-user) --- */
    function renderSettingsUsers() {
        settingsUsersBody.innerHTML = '';
        userCountBadge.textContent = `(${users.length}/20)`;

        users.forEach(u => {
            const hasPassword = u.password && u.password !== "";
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type="text" class="edit-name" value="${u.name}" required></td>
                <td><input type="tel" class="edit-phone" value="${u.phone || ''}" placeholder="09012345678"></td>
                <td style="font-size: 0.8rem; color: var(--text-secondary); min-width: 90px;">
                    ${hasPassword ? `
                    <div style="display:flex; align-items:center; gap:0.2rem;">
                        <span>********</span>
                        <button class="btn danger small reset-pw-btn" data-id="${u.id}" title="PWリセット">消</button>
                    </div>
                    ` : '未設定'}
                </td>
                <td><input type="time" class="edit-in" value="${u.basicIn}"></td>
                <td><input type="time" class="edit-out" value="${u.basicOut}"></td>
                <td><input type="number" class="edit-allowance" value="${u.allowance}"></td>
                <td>
                    <button class="btn danger small delete-user-btn" data-id="${u.id}">削除</button>
                </td>
            `;
            settingsUsersBody.appendChild(tr);
        });

        // Disable Add Button if 20 reached
        addUserBtn.disabled = users.length >= 20;

        document.querySelectorAll('.delete-user-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.target.dataset.id;
                deletedUserIds.add(id);
                users = users.filter(u => u.id !== id);
                renderSettingsUsers();
            });
        });

        document.querySelectorAll('.reset-pw-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const id = e.target.dataset.id;
                const u = users.find(user => user.id === id);
                if (u && confirm(`${u.name} のパスワードを初期化（電話番号でのログインに戻す）しますか？`)) {
                    u.password = "";
                    renderSettingsUsers();
                }
            });
        });
    }

    settingsBtn.addEventListener('click', () => {
        adminPasswordSetInput.value = adminPassword;
        renderSettingsUsers();
        settingsModal.classList.add('show');
    });

    closeModalBtn.addEventListener('click', async () => {
        deletedUserIds.clear();
        await loadServerData();
        renderHistory();
        renderAggregateUserDropdown();
        checkLoginState();
        settingsModal.classList.remove('show');
    });

    addUserBtn.addEventListener('click', () => {
        if (users.length >= 20) return;
        users.push({
            id: 'u_' + Date.now(),
            name: '新規ユーザー',
            phone: '',
            basicIn: '09:00',
            basicOut: '18:00',
            allowance: 0
        });
        renderSettingsUsers();
    });

    saveSettingsBtn.addEventListener('click', async () => {
        const rows = settingsUsersBody.querySelectorAll('tr');
        const newUsers = [];
        let hasError = false;
        let phones = new Set();

        rows.forEach((row, index) => {
            const u = users[index];
            const nameStr = row.querySelector('.edit-name').value.trim();
            const phoneStr = row.querySelector('.edit-phone').value.trim();

            if (!nameStr) {
                alert('名前が空のユーザーがいます。');
                hasError = true;
                return;
            }
            if (!phoneStr) {
                alert('電話番号が空のユーザーがいます。打刻パスワードに必要です。');
                hasError = true;
                return;
            }
            if (phones.has(phoneStr)) {
                alert('電話番号が重複しているユーザーがいます。電話番号は一意である必要があります。');
                hasError = true;
                return;
            }
            phones.add(phoneStr);

            newUsers.push({
                id: u.id,
                name: nameStr,
                phone: phoneStr,
                password: u.password || "",
                basicIn: row.querySelector('.edit-in').value || '09:00',
                basicOut: row.querySelector('.edit-out').value || '18:00',
                allowance: Number(row.querySelector('.edit-allowance').value) || 0
            });
        });

        if (hasError) return;

        showLoading();
        try {
            if (db) {
                for (let id of deletedUserIds) await deleteDoc(doc(db, "users", id));
                for (let nu of newUsers) await setDoc(doc(db, "users", nu.id), nu);
            }
            users = newUsers;
            deletedUserIds.clear();

            const newPwd = adminPasswordSetInput.value.trim();
            if (newPwd && newPwd !== adminPassword) {
                adminPassword = newPwd;
                if (db) await setDoc(doc(db, "settings", "config"), { adminPassword });
            }

            checkLoginState(); // Validate current login just in case
            renderHistory(); // Re-render to update names
            settingsModal.classList.remove('show');
        } catch (e) {
            console.error(e);
            alert("設定の保存に失敗しました");
        }
        hideLoading();
    });

    /* --- Help Modal Logic --- */
    function openHelpModal(tabId) {
        // 全てのタブとコンテンツをリセット
        helpTabBtns.forEach(btn => btn.classList.remove('active'));
        helpContents.forEach(content => content.classList.remove('active'));

        // 指定されたタブとコンテンツをアクティブにする
        const targetBtn = Array.from(helpTabBtns).find(btn => btn.dataset.target === tabId);
        const targetContent = document.getElementById(tabId);

        if (targetBtn && targetContent) {
            targetBtn.classList.add('active');
            targetContent.classList.add('active');
        }

        helpModal.classList.add('show');
    }

    if (userHelpBtn) {
        userHelpBtn.addEventListener('click', () => {
            openHelpModal('userHelpContent');
        });
    }

    if (closeHelpModalBtn) {
        closeHelpModalBtn.addEventListener('click', () => {
            helpModal.classList.remove('show');
        });
    }

    helpTabBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            const targetId = e.currentTarget.dataset.target;
            openHelpModal(targetId);
        });
    });

    // Settings 専用ヘルプのロジック
    if (adminHelpBtn) {
        adminHelpBtn.addEventListener('click', () => {
            settingsHelpModal.classList.add('show');
        });
    }

    if (closeSettingsHelpModalBtn) {
        closeSettingsHelpModalBtn.addEventListener('click', () => {
            settingsHelpModal.classList.remove('show');
        });
    }

    /* --- Aggregate User Dropdown --- */
    function renderAggregateUserDropdown() {
        const selectedAgg = userSelectAggregate.value;
        userSelectAggregate.innerHTML = '<option value="ALL">全員</option>';
        users.forEach(u => {
            const opt = document.createElement('option');
            opt.value = u.id;
            opt.textContent = u.name;
            userSelectAggregate.appendChild(opt);
        });
        if (selectedAgg && (selectedAgg === 'ALL' || getUserById(selectedAgg))) {
            userSelectAggregate.value = selectedAgg;
        }
    }

    /* --- Punch Logic --- */
    function showMessage(msg) {
        statusMessageEl.textContent = msg;
        setTimeout(() => statusMessageEl.textContent = '', 3000);
    }

    punchInBtn.addEventListener('click', async () => {
        if (!currentUserForPunch) return;
        const user = getUserById(currentUserForPunch);
        if (!user) return;

        const now = new Date();
        const session = {
            userId: user.id,
            userNameSnapshot: user.name,
            dateStr: formatDateForInput(now),
            inTime: now.toISOString(),
            status: []
        };

        const inTimeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
        if (inTimeStr > user.basicIn) {
            session.status.push('遅刻');
        }

        currentSessions[user.id] = session;

        showLoading();
        try {
            if (db) await setDoc(doc(db, "sessions", user.id), session);
            showMessage(`${user.name}：${inTimeStr} に出勤しました`);
            logoutUser(); // 安全のため打刻後はログアウト
        } catch (e) {
            delete currentSessions[user.id];
            alert("出勤の記録に失敗しました");
        }
        hideLoading();
    });

    punchOutBtn.addEventListener('click', async () => {
        if (!currentUserForPunch) return;
        const user = getUserById(currentUserForPunch);
        if (!user || !currentSessions[user.id]) return;

        const session = currentSessions[user.id];
        const now = new Date();
        session.outTime = now.toISOString();

        const outTimeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
        if (outTimeStr < user.basicOut) {
            session.status.push('早退');
        }

        if (session.status.length === 0) {
            session.status.push('通常');
        }

        calculateSessionTimes(session, user);

        showLoading();
        try {
            if (db) {
                const histRef = await addDoc(collection(db, "history"), session);
                session.docId = histRef.id;
                await deleteDoc(doc(db, "sessions", user.id));
            }
            history.push(session);
            delete currentSessions[user.id];

            renderHistory();
            showMessage(`${user.name}：${outTimeStr} に退勤しました`);
            logoutUser(); // 安全のため打刻後はログアウト
        } catch (e) {
            console.error(e);
            alert("退勤の記録に失敗しました");
        }
        hideLoading();
    });

    function calculateSessionTimes(session, user) {
        const inTime = new Date(session.inTime);
        const outTime = new Date(session.outTime);
        const diffMs = outTime - inTime;
        let diffHours = diffMs / (1000 * 60 * 60);

        let breakHours = 0;
        if (diffHours > 8) {
            breakHours = 1; // 60 mins
        } else if (diffHours > 6) {
            breakHours = 45 / 60; // 45 mins
        }

        let actualWorkHours = Math.max(0, diffHours - breakHours);

        const basicInParts = user.basicIn.split(':').map(Number);
        const basicOutParts = user.basicOut.split(':').map(Number);
        let basicTotalHours = (basicOutParts[0] + basicOutParts[1] / 60) - (basicInParts[0] + basicInParts[1] / 60);

        let basicBreak = 0;
        if (basicTotalHours > 8) basicBreak = 1;
        else if (basicTotalHours > 6) basicBreak = 45 / 60;

        let basicActual = Math.max(0, basicTotalHours - basicBreak);
        let overtime = Math.max(0, actualWorkHours - basicActual);

        session.workHours = actualWorkHours;
        session.breakHours = breakHours;
        session.overtime = overtime;
        session.allowance = user.allowance;
    }

    /* --- History --- */
    function renderHistory() {
        historyBody.innerHTML = '';
        const sorted = [...history].sort((a, b) => new Date(b.dateStr) - new Date(a.dateStr));

        sorted.forEach((record, index) => {
            const tr = document.createElement('tr');

            const currentUser = getUserById(record.userId);
            const displayName = currentUser ? currentUser.name : (record.userNameSnapshot || '不明');

            const inT = new Date(record.inTime);
            const inStr = `${String(inT.getHours()).padStart(2, '0')}:${String(inT.getMinutes()).padStart(2, '0')}`;
            let outStr = '-';
            if (record.outTime) {
                const outT = new Date(record.outTime);
                outStr = `${String(outT.getHours()).padStart(2, '0')}:${String(outT.getMinutes()).padStart(2, '0')}`;
            }

            let statusHtml = '';
            if (record.status) {
                record.status.forEach(st => {
                    let className = 'status-normal';
                    if (st === '遅刻') className = 'status-late';
                    if (st === '早退') className = 'status-early';
                    statusHtml += `<span class="status-badge ${className}">${st}</span> `;
                });
            }

            tr.innerHTML = `
                <td>${displayName}</td>
                <td>${record.dateStr}</td>
                <td>${inStr}</td>
                <td>${outStr}</td>
                <td>${statusHtml}</td>
                <td>${record.breakHours ? record.breakHours.toFixed(2) + 'h' : '0.00h'}</td>
                <td>${record.workHours ? record.workHours.toFixed(2) + 'h' : '0.00h'}</td>
                <td>${record.overtime ? record.overtime.toFixed(2) + 'h' : '0.00h'}</td>
                <td>${record.allowance !== undefined ? record.allowance + '円' : '0円'}</td>
                <td><button class="btn danger small delete-btn" data-index="${index}">削除</button></td>
            `;
            historyBody.appendChild(tr);
        });

        document.querySelectorAll('.delete-btn').forEach(btn => {
            btn.addEventListener('click', async (e) => {
                const globalIndex = history.findIndex(r => r === sorted[e.target.dataset.index]);
                if (globalIndex > -1) {
                    const record = history[globalIndex];
                    if (confirm('この記録を削除しますか？')) {
                        showLoading();
                        try {
                            if (db && record.docId) {
                                await deleteDoc(doc(db, "history", record.docId));
                            }
                            history.splice(globalIndex, 1);
                            renderHistory();
                        } catch (err) {
                            alert("削除に失敗しました");
                        }
                        hideLoading();
                    }
                }
            });
        });
    }

    clearHistoryBtn.addEventListener('click', async () => {
        if (confirm('すべての履歴を削除しますか？')) {
            showLoading();
            try {
                if (db) {
                    for (let record of history) {
                        if (record.docId) await deleteDoc(doc(db, "history", record.docId));
                    }
                }
                history = [];
                renderHistory();
            } catch (e) {
                alert("一括削除中にエラーが発生しました");
            }
            hideLoading();
        }
    });

    /* --- Aggregation & Export --- */
    function getRecordsInRange(startStr, endStr) {
        return history.filter(record => {
            return record.dateStr >= startStr && record.dateStr <= endStr;
        });
    }

    calculateBtn.addEventListener('click', () => {
        const start = startDateInput.value;
        const end = endDateInput.value;

        if (!start || !end) {
            alert('期間を指定してください');
            return;
        }
        if (start > end) {
            alert('開始日は締め日より前である必要があります');
            return;
        }

        const records = getRecordsInRange(start, end);
        const container = document.getElementById('aggregateResultsContainer');
        container.innerHTML = '';

        if (records.length === 0) {
            container.innerHTML = '<p style="text-align: center; color: #cbd5e1;"> 対象期間のデータがありません </p>';
            container.style.display = 'block';
            return;
        }

        // ユーザーごとに集計
        const summaryByUser = {};

        users.forEach(u => {
            summaryByUser[u.id] = {
                name: u.name,
                days: new Set(),
                totalHours: 0,
                overtime: 0,
                allowance: 0
            };
        });

        records.forEach(r => {
            if (!summaryByUser[r.userId]) return; // 削除されたユーザー等の対策
            summaryByUser[r.userId].days.add(r.dateStr);
            summaryByUser[r.userId].totalHours += (r.workHours || 0);
            summaryByUser[r.userId].overtime += (r.overtime || 0);
            summaryByUser[r.userId].allowance += (r.allowance || 0);
        });

        const startDate = new Date(start);
        const endDate = new Date(end);
        const diffTime = Math.abs(endDate - startDate);
        const totalDaysInt = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) + 1;

        let resultHtml = '';

        // 各ユーザーのカードを生成
        Object.values(summaryByUser).forEach(summary => {
            const workDays = summary.days.size;
            // 未稼働ユーザーは表示をスキップするかどうか（ここでは全員表示する方針）
            const holidays = Math.max(0, totalDaysInt - workDays);

            resultHtml += `
            <div class="user-result-card">
                <h4>${summary.name}</h4>
                <div class="aggregate-results" style="background: transparent; padding: 0;">
                    <div class="result-item"><span>出勤:</span> <span>${workDays}日</span></div>
                    <div class="result-item"><span>休日:</span> <span>${holidays}日</span></div>
                    <div class="result-item"><span>総勤務:</span> <span>${summary.totalHours.toFixed(2)}h</span></div>
                    <div class="result-item"><span>残業:</span> <span>${summary.overtime.toFixed(2)}h</span></div>
                    <div class="result-item"><span>手当:</span> <span>${summary.allowance.toLocaleString()}円</span></div>
                </div>
            </div>`;
        });

        container.innerHTML = resultHtml;
        container.style.display = 'block';
    });

    exportExcelBtn.addEventListener('click', () => {
        const start = startDateInput.value;
        const end = endDateInput.value;

        if (!start || !end) {
            alert('期間を指定してください');
            return;
        }

        const records = getRecordsInRange(start, end);
        if (records.length === 0) {
            alert('対象期間のデータがありません');
            return;
        }

        const wb = XLSX.utils.book_new();

        // ユーザー毎にシートを分けて出力
        users.forEach(user => {
            const userRecords = records.filter(r => r.userId === user.id);
            // データがなくてもシートは作る（要望次第でスキップ可能）

            const exportData = [];
            exportData.push(["【集計結果】"]);
            exportData.push([`期間: ${start} 〜 ${end}`]);
            exportData.push([`対象: ${user.name}`]);
            exportData.push(["出勤日数", "休日日数", "総勤務時間", "残業時間", "出動手当"]);

            const days = new Set();
            let th = 0, to = 0, ta = 0;
            userRecords.forEach(r => {
                days.add(r.dateStr);
                th += (r.workHours || 0);
                to += (r.overtime || 0);
                ta += (r.allowance || 0);
            });

            const workDays = days.size;
            const startDate = new Date(start);
            const endDate = new Date(end);
            const totalDaysInt = Math.ceil(Math.abs(endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
            const holidays = Math.max(0, totalDaysInt - workDays);

            exportData.push([
                `${workDays}日`,
                `${holidays}日`,
                `${th.toFixed(2)}h`,
                `${to.toFixed(2)}h`,
                `${ta}円`
            ]);
            exportData.push([]);

            exportData.push(["【打刻履歴詳細】"]);
            exportData.push(["日付", "出勤時間", "退勤時間", "状態", "休憩時間", "実働時間", "残業時間", "出動手当"]);

            userRecords.sort((a, b) => new Date(a.dateStr) - new Date(b.dateStr)).forEach(r => {
                const inT = new Date(r.inTime);
                const inStr = `${String(inT.getHours()).padStart(2, '0')}:${String(inT.getMinutes()).padStart(2, '0')}`;
                let outStr = '';
                if (r.outTime) {
                    const outT = new Date(r.outTime);
                    outStr = `${String(outT.getHours()).padStart(2, '0')}:${String(outT.getMinutes()).padStart(2, '0')}`;
                }

                exportData.push([
                    r.dateStr,
                    inStr,
                    outStr,
                    (r.status || []).join(' / '),
                    r.breakHours ? r.breakHours.toFixed(2) : '0',
                    r.workHours ? r.workHours.toFixed(2) : '0',
                    r.overtime ? r.overtime.toFixed(2) : '0',
                    r.allowance || 0
                ]);
            });

            const ws = XLSX.utils.aoa_to_sheet(exportData);
            const wscols = [
                { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }
            ];
            ws['!cols'] = wscols;

            // Sheet名が31文字を超えたり重複しないようにする
            let safeName = user.name.substring(0, 30);
            XLSX.utils.book_append_sheet(wb, ws, safeName);
        });

        const fileName = `Timecard_AllUsers_${start}_${end}.xlsx`;
        XLSX.writeFile(wb, fileName);
    });

    // 初期化実行
    init();
});
