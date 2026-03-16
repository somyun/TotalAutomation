// app.js

// --- Global Data Store ---
let appConfig = {};
let selectedUserId = null;
let loginUsers = []; // Store user data for login validation
let isWorkLogInitialized = false; // Flag for persistence
let pendingRestoreState = null; // State waiting for config load
let currentSafetyEduData = {}; // [New] Store for Safety Edu Presets

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    // Initial Nav Setup
    switchMainTab('view-work-log');

    // Manual Window Drag Handler
    const titleBar = document.getElementById('title-bar');
    if (titleBar) {
        titleBar.addEventListener('mousedown', (e) => {
            if (e.button !== 0) return;
            if (e.target.closest('.control-btn') || e.target.tagName === 'INPUT') return;
            sendMessageToAHK({ command: 'dragWindow' });
        });
    }

    // Event Listener for ERP Refresh Button
    const btnRefresh = document.getElementById('btn-refresh-erp');
    if (btnRefresh) {
        btnRefresh.addEventListener('click', () => {
            // Clear previous status text temporarily or show loading state if desired
            document.getElementById('erp-last-update').innerText = "조회 중...";
            sendMessageToAHK({ command: 'refreshERPOrder' });
        });
    }

    // Request Initial Data
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({ command: 'ready' });
    }

    // Password Validation Listeners
    setupPasswordValidation('new-webpw', 'new-webpw2');
    setupPasswordValidation('new-pw2', 'new-pw2-confirm');
    setupPasswordValidation('new-sappw', 'new-sappw2');

    // --- UX Improvements ---
    // 1. Numeric Input Restrictions
    const numericInputs = ['login-pw2', 'new-id', 'new-pw2', 'new-pw2-confirm', 'user-pw2'];
    numericInputs.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', (e) => {
                e.target.value = e.target.value.replace(/[^0-9]/g, '');

                // 2. Auto-Login Trigger (Only for login-pw2)
                if (id === 'login-pw2') {
                    checkAutoLogin(e.target.value);
                }
            });
        }
    });

    // [New] Auto-formatting for Time and Phone inputs
    const phoneInputs = ['ta-driver-phone', 'ta-worker-phone', 'ta-safety-phone'];
    const timeInputs = ['ta-work-start', 'ta-work-end', 'ta-op-start', 'ta-op-end', 'vl-start-time', 'vl-end-time'];

    phoneInputs.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', (e) => formatPhone(e.target));
    });

    timeInputs.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', (e) => formatTime(e.target));
    });

    // ERP Check Order Number Restriction (Dynamic elements handled elsewhere or if static)
    // Assuming adding a global delegation or checking specific view load in future if needed.
    // For now, let's look for specific ID if it exists? 
    // The user mentioned "ERP Check tab order number", but that might be dynamically generated.
    // I will add a helper for it.

    // [UX] Enter Key Handler for User Selection View
    document.addEventListener('keydown', (e) => {
        const loginView = document.getElementById('login-view');
        // Check if Login View is visible (offsetParent is null if display:none or parent is hidden)
        if (e.key === 'Enter' && loginView && loginView.offsetParent !== null) {
            // Only proceed if a user is actually selected
            if (selectedUserId) {
                e.preventDefault(); // Prevent double triggers if button is focused
                moveToPasswordView();
            }
        }
    });
});

function checkAutoLogin(inputPw) {
    if (!selectedUserId) return;

    // Find selected user data
    // loginUsers contains profile objects directly
    const user = loginUsers.find(u => u.id === selectedUserId);
    if (user && user.pw2 === inputPw) {
        // Match found! Login immediately.
        tryLogin();
    }
}

// --- AHK Bridge ---
if (window.chrome && window.chrome.webview) {
    window.chrome.webview.addEventListener('message', event => {
        try {
            // PostWebMessageAsJson sends a parsed object, no need to parse again
            const msg = event.data;
            handleAhkMessage(msg);
        } catch (e) {
            console.error('Error handling AHK message:', e);
        }
    });
} else {
    // Browser fallback
    console.warn('WebView2 environment not detected.');
}

function sendMessageToAHK(payload) {
    // [Refactor] Inject UI State for fast reload on 'runTask'
    if (payload && payload.command === 'runTask') {
        try {
            const state = collectUiState();
            payload.uiState = state;
        } catch (e) {
            console.error("Failed to collect UI state for runTask:", e);
        }
    }

    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage(payload);
    } else {
        console.log('To AHK:', payload);
    }
}

function handleAhkMessage(msg) {
    switch (msg.type) {
        case 'initLogin':
            loginUsers = msg.users;
            renderUserList(msg.users);
            break;
        case 'loginSuccess':
            handleLoginSuccess(msg.profile);
            break;
        case 'loginFail':
            showNativeMsgBox(msg.message, "로그인 실패");
            break;
        case 'updateShiftStatus':
            updateShiftUI(msg.data);
            break;
        case 'loadConfig':
            appConfig = msg.data;
            initPresets(); // Initialize Presets for Track/Vehicle Views
            if (document.getElementById('settings-view').style.display !== 'none') {
                loadSettingsToUI();
            }
            // Initialize Work Log IF logged in and not yet done
            if (selectedUserId) {
                if (!isWorkLogInitialized) {
                    renderWorkLogUI(); // 최초 렌더링
                    isWorkLogInitialized = true;
                }

                // [Hot Reload] 통합 UI 갱신 (설정 변경 사항 즉시 반영)
                refreshUI();
            }

            // Apply pending restore if exists (Race Condition Fix)
            if (pendingRestoreState) {
                restoreUiStateData(pendingRestoreState);
                pendingRestoreState = null;
            }
            break;
        case 'releaseNotes':
            if (msg.error) {
                renderReleaseNotes(null, msg.error);
            } else {
                renderReleaseNotes(msg.data);
            }
            break;
        case 'getUiState':
            const state = collectUiState();
            sendMessageToAHK({ command: 'saveUiState', data: state });
            break;
        case 'restoreUiState':
            if (isWorkLogInitialized) {
                restoreUiStateData(msg.data);
            } else {
                pendingRestoreState = msg.data;
            }
            break;
        case 'updateERPStatus':
            handleERPStatusUpdate(msg.status);
            break;
        case 'approvalInfo':
            if (msg.data) {
                setVal('vl-approve-no', msg.data.승인번호);
                setVal('vl-dept', msg.data.승인부서);
                setVal('vl-approver', msg.data.승인자);
            } else {
                showNativeMsgBox("승인정보를 불러오지 못했습니다.");
            }
            break;
        case 'updateERPOrderList': // New Case
            handleERPOrderListUpdate(msg.orders);
            break;
        case 'updateTitle':
            // document.title = msg.title; // 윈도우 캡션용 (선택사항)
            const tEl = document.getElementById('app-title');
            if (tEl) tEl.innerText = msg.title; // 사용자 눈에 보이는 커스텀 타이틀바 갱신
            break;
        case 'headlessReady':
            // [New] Headless 준비 완료 시 버튼 활성화
            const btnImport = document.getElementById('btn-import-workers');
            if (btnImport) btnImport.disabled = false;

            // ERP 새로고침 버튼도 활성화
            const btnRefresh = document.getElementById('btn-refresh-erp');
            if (btnRefresh) {
                btnRefresh.disabled = false;
                btnRefresh.style.opacity = "1";
            }
            break;
        case 'updateWorkerList':
            // [New] 작업자 명단 수신 처리
            handleWorkerListUpdate(msg.data);
            break;
        case 'showInitOverlay':
            // [New] 초기화 오버레이 표시
            showInitOverlay(msg.message);
            break;
        case 'hideInitOverlay':
            // [New] 초기화 오버레이 숨김
            hideInitOverlay();
            break;
    }
}

// --- Initialization Overlay ---
function showInitOverlay(message) {
    let overlay = document.getElementById('init-overlay');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.id = 'init-overlay';
        overlay.style.cssText = `
            position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(255, 255, 255, 0.95); z-index: 9999;
            display: flex; flex-direction: column; justify-content: center; align-items: center;
        `;
        document.body.appendChild(overlay);
    }

    // 로딩 이미지 경로 (webview 컨텍스트 기준)
    const imgPath = 'img/loading.gif';

    overlay.innerHTML = `
        <img src="${imgPath}" alt="Loading..." style="width: 64px; height: 64px; margin-bottom: 20px;">
        <div style="font-size: 18px; font-weight: bold; color: #333;">${message || '초기 구성 중...'}</div>
        <div style="font-size: 14px; color: #666; margin-top: 10px;">잠시만 기다려주세요</div>
    `;

    overlay.style.display = 'flex';
}

function hideInitOverlay() {
    const overlay = document.getElementById('init-overlay');
    if (overlay) {
        overlay.style.transition = 'opacity 0.5s ease';
        overlay.style.opacity = '0';
        setTimeout(() => {
            overlay.style.display = 'none';
            overlay.style.opacity = '1';
        }, 500);
    }
}

// --- State Persistence ---
function collectUiState() {
    const state = {
        activeView: null,
        mainTab: null,
        formData: {},
        customState: {} // Generic storage for .savable-ui elements
    };

    // 1. Active View detection
    if (document.getElementById('app-container').style.display !== 'none') state.activeView = 'app';
    else if (document.getElementById('settings-view').style.display !== 'none') state.activeView = 'settings';
    else if (document.getElementById('add-user-view').style.display !== 'none') state.activeView = 'add';
    else state.activeView = 'login'; // default

    // 2. Active Main Tab
    const activeTab = document.querySelector('.view-section.active');
    if (activeTab) state.mainTab = activeTab.id;

    // 3. Form Data
    // 모든 input, select, textarea 수집 (비밀번호 제외)
    const elements = document.querySelectorAll('input, select, textarea');
    elements.forEach(el => {
        if (!el.id) return; // ID가 없으면 복구 불가
        if (el.type === 'password') return; // 비밀번호는 제외

        // 제외할 기타 요소들이 있다면 여기서 필터링

        let value;
        if (el.type === 'checkbox') {
            value = el.checked;
        } else if (el.type === 'radio') {
            if (el.checked) value = el.value;
            else return; // 라디오는 선택된 것만 저장
        } else {
            value = el.value;
        }

        state.formData[el.id] = {
            tag: el.tagName,
            type: el.type,
            value: value
        };
    });

    // 4. Custom Savable UI
    document.querySelectorAll('.savable-ui').forEach(el => {
        const key = el.getAttribute('data-save-key');
        if (key) {
            state.customState[key] = getSavableValue(el, key);
        }
    });

    return state;
}

function getSavableValue(el, key) {
    if (key === 'erpLocation') {
        return selectedERPLocation;
    }
    // Future extensions:
    // if (key === 'someOtherWidget') return ...;
    return null;
}

function restoreUiStateData(state) {
    if (!state) return;

    // 1. Restore View
    if (state.activeView) {
        switchView(state.activeView);
    }

    // 2. Restore Main Tab (if in app view)
    if (state.activeView === 'app' && state.mainTab) {
        switchMainTab(state.mainTab);
    }

    // 3. Restore Form Data
    if (state.formData) {
        Object.keys(state.formData).forEach(id => {
            const data = state.formData[id];
            const el = document.getElementById(id);
            if (!el) return;

            // 타입 검사 등 안전장치
            if (el.type === 'checkbox') {
                el.checked = data.value;
            } else if (el.type === 'radio') {
                // 라디오는 같은 Name 그룹 내에서 해당 ID를 체크
                el.checked = true;
            } else {
                el.value = data.value;
            }

            // Trigger input event for auto-save logic or UI updates
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        });
    }



    // 4. Restore Custom UI
    if (state.customState) {
        Object.keys(state.customState).forEach(key => {
            const val = state.customState[key];
            // Find element by key
            const el = document.querySelector(`.savable-ui[data-save-key="${key}"]`);
            if (el) {
                applySavableValue(el, key, val);
            }
        });
    }

}

function applySavableValue(el, key, value) {
    if (key === 'erpLocation') {
        selectedERPLocation = value;
        // Search and Select Button
        // Wait a bit for dynamic content if needed, though usually loaded by now
        setTimeout(() => {
            const buttons = el.querySelectorAll('.erp-btn');
            buttons.forEach(btn => {
                if (btn.innerText === value) {
                    btn.classList.add('selected');
                } else {
                    btn.classList.remove('selected');
                }
            });
        }, 50);
    }
}

// --- Navigation & Views ---

function switchMainTab(viewId) {
    document.querySelectorAll('#content .view-section').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.nav-top .menu-item').forEach(el => el.classList.remove('active'));

    const target = document.getElementById(viewId);
    if (target) target.classList.add('active');

    const navItem = document.querySelector(`.nav-top .menu-item[data-target="${viewId}"]`);
    if (navItem) navItem.classList.add('active');

    if (viewId === 'view-erp-check') {
        // Apply User Preference for Inspector Format
        // Logic: If user has a specific preference, override the current state?
        // Or only set initial state? "Initial value" was requested.
        // Let's set it based on config every time we switch to this tab.
        const uid = selectedUserId;
        if (uid && appConfig.users && appConfig.users[uid]) {
            const pref = appConfig.users[uid].profile.erpFormat || 'summary';
            const toggle = document.getElementById('toggle-worker-format');
            // summary = unchecked, list = checked
            const shouldBeChecked = (pref === 'list');
            if (toggle.checked !== shouldBeChecked) {
                toggle.checked = shouldBeChecked;
                updateToggleStyle();
                // Note: We do NOT dispatch 'change' event here to avoid triggering auto-save indirectly
                // (though we added the guard clause just in case)
            }
        }
        // 통합 UI 갱신 함수 호출 (ERP 점검 포함)
        refreshUI();

        // 4차 요구사항: 탭 진입 시 자동 새로고침(AHK 갱신 요청) 수행
        const btnRefresh = document.getElementById('btn-refresh-erp');
        if (btnRefresh && !btnRefresh.disabled) {
            btnRefresh.click();
        }
    } else if (viewId === 'view-work-log') {
        // Only render if NOT initialized yet (Persistence Fix)
        if (!isWorkLogInitialized) {
            renderWorkLogUI();
        } else {
            // 이미 초기화된 경우 통합 UI 갱신 함수 호출
            refreshUI();
        }
    }
}

function switchView(viewName) {
    // Hide all main containers first
    const loginCon = document.getElementById('login-container');
    const appCon = document.getElementById('app-container');
    const setView = document.getElementById('settings-view');

    loginCon.style.display = 'none';
    appCon.style.display = 'none';
    setView.style.display = 'none';

    if (viewName === 'login') {
        loginCon.style.display = 'flex';
        document.getElementById('login-view').style.display = 'flex';
        document.getElementById('password-view').style.display = 'none';
        document.getElementById('add-user-view').style.display = 'none';
    } else if (viewName === 'add') {
        loginCon.style.display = 'flex';
        document.getElementById('login-view').style.display = 'none';
        document.getElementById('add-user-view').style.display = 'block';
        resetAddUserForm();
    } else if (viewName === 'app') {
        appCon.style.display = 'flex';
    } else if (viewName === 'settings') {
        setView.style.display = 'flex';
        // Ensure data is loaded
        loadSettingsToUI();
    }
}

function openSettings() {
    sendMessageToAHK({ command: 'requestConfig' });
    switchView('settings');
    switchSettingsTab('tab-user');
}

function closeSettings() {
    sendMessageToAHK({ command: 'exitSettings' });
    switchView('app');
}

function switchSettingsTab(tabId) {
    document.querySelectorAll('.settings-sidebar .tab-item').forEach(el => el.classList.remove('active'));
    const clickedTab = document.querySelector(`.settings-sidebar .tab-item[onclick="switchSettingsTab('${tabId}')"]`);
    if (clickedTab) clickedTab.classList.add('active');

    document.querySelectorAll('.settings-tab-pane').forEach(el => el.classList.remove('active'));
    const targetPane = document.getElementById(tabId);
    if (targetPane) targetPane.classList.add('active');

    if (tabId === 'tab-updates') {
        requestReleaseNotes();
    }
}

function requestReleaseNotes() {
    const container = document.getElementById('release-list');
    if (container) container.innerHTML = '<div class="loading-spinner">불러오는 중...</div>';
    sendMessageToAHK({ command: 'getReleaseNotes' });
}

function renderReleaseNotes(data, error = null) {
    const container = document.getElementById('release-list');
    if (!container) return;

    container.innerHTML = '';

    if (error) {
        container.innerHTML = `<div style="color:red; text-align:center;">오류: ${error}</div>`;
        return;
    }

    if (!data || data.length === 0) {
        container.innerHTML = '<div style="text-align:center; padding:20px;">업데이트 내역이 없습니다.</div>';
        return;
    }

    data.forEach(release => {
        const date = new Date(release.published_at).toLocaleDateString('ko-KR');
        const bodyText = release.body || "내용 없음";

        const item = document.createElement('div');
        item.className = 'release-item';
        item.innerHTML = `
            <div class="release-header">
                <span class="release-ver">${release.tag_name}</span>
                <span class="release-date">${date}</span>
            </div>
            <div class="release-body">${escapeHtml(bodyText)}</div>
        `;
        container.appendChild(item);
    });
}

function escapeHtml(text) {
    return text.replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// ... (Rest of Login Logic remains same) ...

// --- Validation Utils ---
function setupPasswordValidation(id1, id2) {
    const el1 = document.getElementById(id1);
    const el2 = document.getElementById(id2);
    if (!el1 || !el2) return;

    const validate = () => {
        const val1 = el1.value;
        const val2 = el2.value;

        // Find label to add checkmark
        const label = el2.parentNode.querySelector('label');

        if (val2.length > 0) {
            if (val1 === val2) {
                // Match
                el2.classList.remove('inputs-error');
                el2.classList.add('inputs-match');
                if (label) label.classList.add('validation-success');
            } else {
                // Mismatch
                el2.classList.remove('inputs-match');
                el2.classList.add('inputs-error');
                if (label) label.classList.remove('validation-success');
            }
        } else {
            // Empty
            el2.classList.remove('inputs-match');
            el2.classList.remove('inputs-error');
            if (label) label.classList.remove('validation-success');
        }
    };

    el1.addEventListener('input', validate);
    el2.addEventListener('input', validate);
}

// --- Login Logic ---
function renderUserList(users) {
    const container = document.getElementById('user-list-container');
    container.innerHTML = '';

    if (!users || users.length === 0) {
        container.innerHTML = '<div style="padding:20px; text-align:center; color:#888;">등록된 사용자가 없습니다.</div>';
        return;
    }

    users.forEach(user => {
        const item = document.createElement('div');
        item.className = 'user-item';
        item.ondblclick = () => {
            selectUser(user.id);
            moveToPasswordView();
        };
        item.onclick = () => selectUser(user.id);

        item.innerHTML = `
            <div class="user-info">
                <span class="user-name">${user.name}</span>
                <span class="user-id">(${user.id})</span>
            </div>
        `;
        item.dataset.id = user.id;
        container.appendChild(item);
    });
}

function selectUser(id) {
    selectedUserId = id;
    // Highlight UI
    document.querySelectorAll('.user-item').forEach(el => el.classList.remove('selected'));
    const item = document.querySelector(`.user-item[data-id="${id}"]`);
    if (item) item.classList.add('selected');

    // Enable Buttons
    document.getElementById('btn-next').disabled = false;
    document.getElementById('btn-delete').disabled = false;
}

function moveToPasswordView() {
    if (!selectedUserId) return;
    const user = loginUsers.find(u => u.id === selectedUserId);
    if (!user) return;

    // Switch to Password View
    document.getElementById('login-view').style.display = 'none';
    const pwdView = document.getElementById('password-view');
    pwdView.style.display = 'flex';

    document.getElementById('pwd-user-name').innerText = user.name;
    document.getElementById('pwd-user-id').innerText = `(${user.id})`;

    const pwInput = document.getElementById('login-pw2');
    pwInput.value = '';
    pwInput.focus();
}

function backToUserList() {
    selectedUserId = null;
    // Clear UI Selection
    document.querySelectorAll('.user-item').forEach(el => el.classList.remove('selected'));

    // Disable Buttons
    document.getElementById('btn-next').disabled = true;
    document.getElementById('btn-delete').disabled = true;

    document.getElementById('password-view').style.display = 'none';
    document.getElementById('login-view').style.display = 'flex';
}

function handleLoginKey(e) {
    if (e.key === 'Enter') {
        tryLogin();
    }
}

function tryLogin() {
    if (!selectedUserId) {
        showNativeMsgBox('사용자를 선택해주세요.', '알림');
        return;
    }

    const pwInput = document.getElementById('login-pw2');
    const enteredPw = pwInput.value;

    const user = loginUsers.find(u => u.id === selectedUserId);
    if (!user) return;

    // Local 2nd Password Check
    if (user.pw2 !== enteredPw) {
        showNativeMsgBox('2차 비밀번호가 일치하지 않습니다.', '로그인 실패');
        pwInput.value = '';
        pwInput.focus();
        return;
    }

    sendMessageToAHK({ command: 'tryLogin', id: selectedUserId });
}

function handleLoginSuccess(profile) {
    switchView('app');
    const titleEl = document.getElementById('app-title');
    if (titleEl) titleEl.innerText = `통합자동화 v3.0 - ${profile.name}`;

    // Set selectedUserId to current logged in user to ensure settings load correct user profile
    selectedUserId = profile.id;

    // Request full config
    sendMessageToAHK({ command: 'requestConfig' });

    // Auto-init Work Log View (Work Type Auto-selection)
    // Wait for Config Load (Race Condition Fix)
    // renderWorkLogUI(); // Removed here, moved to loadConfig
    isWorkLogInitialized = false; // Add this line to ensure reset

    switchMainTab('view-work-log');
}

function deleteUser() {
    if (!selectedUserId) {
        showNativeMsgBox("삭제할 유저를 선택해주세요.");
        return;
    }
    sendMessageToAHK({ command: 'deleteUser', id: selectedUserId });
}

function submitNewUser() {
    const name = getVal('new-name');
    const id = getVal('new-id');
    const team = getVal('new-team');

    const webpw = getVal('new-webpw');
    const webpw2 = getVal('new-webpw2');
    const pw2 = getVal('new-pw2');
    const pw2_confirm = getVal('new-pw2-confirm');
    const sappw = getVal('new-sappw');
    const sappw2 = getVal('new-sappw2');

    if (!name || !id || !webpw || !pw2) {
        showNativeMsgBox('필수 정보를 입력해주세요.');
        return;
    }
    if (webpw !== webpw2) {
        showNativeMsgBox('통합 비밀번호가 일치하지 않습니다.');
        return;
    }
    if (pw2 !== pw2_confirm) {
        showNativeMsgBox('2차 비밀번호가 일치하지 않습니다.');
        return;
    }
    if (sappw && sappw !== sappw2) {
        showNativeMsgBox('SAP 비밀번호가 일치하지 않습니다.');
        return;
    }

    const newUser = { id, name, team, webPW: webpw, pw2: pw2, sapPW: sappw };
    sendMessageToAHK({ command: 'addUser', data: newUser });

    switchView('login');
}

function resetAddUserForm() {
    ['new-name', 'new-id', 'new-webpw', 'new-webpw2', 'new-pw2', 'new-pw2-confirm', 'new-sappw', 'new-sappw2'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
}

// --- Settings Logic ---
let saveTimeout = null;
function autoSaveSettings() {
    if (saveTimeout) clearTimeout(saveTimeout);
    saveTimeout = setTimeout(saveSettings, 500); // 500ms debounce
}



function loadSettingsToUI() {
    const uid = selectedUserId;
    if (!uid || !appConfig.users) return;

    // Revert: users is an Object (Map) keyed by ID
    const user = appConfig.users[uid];
    if (!user) return;

    const profile = user.profile || {};

    setVal('user-name', profile.name);
    setVal('user-id', profile.id);
    setVal('user-dept', profile.arbpl);
    setVal('user-team', profile.team);
    setVal('user-webpw', profile.webPW);
    setVal('user-pw2', profile.pw2);
    setVal('user-pw2', profile.pw2);
    setVal('user-sappw', profile.sapPW);
    // New Setting: Inspector Display Format Preference (Default: summary)
    setVal('erp-format-pref', profile.erpFormat || 'summary');

    // DEBUG: Trace Data
    // showNativeMsgBox("UI Load for " + uid + " | Workers: " + appConfig.appSettings?.colleagues?.length, "Debug");

    // Workers (Global from appSettings)
    const wBody = document.querySelector('#worker-table tbody');
    wBody.innerHTML = '';
    const globalWorkers = appConfig.appSettings?.colleagues || [];
    globalWorkers.forEach(w => addWorkerRowToTable(wBody, w));

    // Locations (Global)
    const lBody = document.querySelector('#location-table tbody');
    lBody.innerHTML = '';
    const globalLocs = (appConfig.appSettings && appConfig.appSettings.locations) ? appConfig.appSettings.locations : [];
    globalLocs.forEach(l => addLocationRowToTable(lBody, l));

    // Hotkeys
    const hBody = document.querySelector('#hotkey-table tbody');
    hBody.innerHTML = '';
    const hotkeys = user.hotkeys || [];
    renderHotkeyTable(hotkeys);

    // Presets (Track Access)
    try {
        renderPresetList(user.presets || {});
    } catch (e) { console.error("Error rendering presets:", e); }

    // --- NEW: Load Daily Log Defaults ---
    const defaults = user.dailyLogDefaults || {};

    // 1. Safety Management (Wrapped)
    try {
        if (defaults.safety) {
            defaults.safety.forEach((safe, idx) => {
                setVal(`fps-safe${idx + 1}-content`, safe.content);
                setVal(`fps-safe${idx + 1}-start`, safe.start);
                setVal(`fps-safe${idx + 1}-end`, safe.end);
            });
        }
    } catch (e) { console.error("Error loading Safety defaults:", e); }

    // 2. Driver Fitness Check & Others (Wrapped)
    try {
        const drvToggle = document.getElementById('fps-driver-check-toggle');
        if (drvToggle) drvToggle.checked = !!defaults.driverCheck;

        // 3. General Work Table
        const gwBody = document.querySelector('#general-work-table tbody');
        if (gwBody) {
            gwBody.innerHTML = '';
            const gwData = defaults.generalWork || [];
            gwData.forEach(row => addGeneralWorkRow(row));
        }

        // 4. Auto Input Reservation
        setVal('fps-auto-input-time', defaults.autoInputTime);
    } catch (e) { console.error("Error loading General/Driver defaults:", e); }

    // 5. Safety Education (Wrapped in Try-Catch)
    try {
        const safeEdu = defaults.safetyEdu || {};
        // Ensure keys
        const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
        days.forEach(d => {
            if (!safeEdu[d]) safeEdu[d] = "";
        });
        currentSafetyEduData = JSON.parse(JSON.stringify(safeEdu));

        const ddlSafeEdu = document.getElementById('fps-safe-edu-day');
        if (ddlSafeEdu) {
            ddlSafeEdu.value = 'Mon'; // Reset to Mon
            // Initial UI Update
            handleSafetyEduDayChange();
        }
    } catch (e) {
        console.error("Error loading Safety Edu defaults:", e);
    }

    // 6. Setup Global AutoSave Delegate
    setupGlobalAutoSaveDelegate();
}
function setupGlobalAutoSaveDelegate() {
    const settingsView = document.getElementById('settings-view');
    if (!settingsView || settingsView.dataset.delegateAttached) return;

    const handleEvent = (e) => {
        const target = e.target;
        // Filter for inputs, selects, textareas
        if (!target.matches('input, select, textarea')) return;

        // --- SPECIFIC HANDLERS (Delegated) ---

        // 1. Safety Education
        if (target.id === 'fps-safe-edu-day') {
            handleSafetyEduDayChange();
            return; // Do NOT autosave on day switch alone, wait for content
        }
        if (target.id === 'fps-safe-edu-content') {
            handleSafetyEduContentChange();
            return; // handleSafetyEduContentChange already calls autoSaveSettings
        }

        // --- GENERIC AUTOSAVE ---
        // For all other fields, trigger auto-save
        autoSaveSettings();
    };

    settingsView.addEventListener('input', handleEvent);
    settingsView.addEventListener('change', handleEvent);

    settingsView.dataset.delegateAttached = "true";
    console.log("Global Settings Auto-Save Delegate Attached");
}

function saveSettings() {
    try {
        // [CRITICAL FIX] Prevent Auto-Save when Settings View is hidden
        // This prevents wiping data if called unintentionally from other views
        if (document.getElementById('settings-view').style.display === 'none') return;

        const uid = selectedUserId;
        if (!uid || !appConfig.users) return;

        // Revert: users is an Object (Map) keyed by ID
        const user = appConfig.users[uid];
        if (!user) return; // Should not happen if logged in

        if (!user.profile) user.profile = {};

        const deptSel = document.getElementById('user-dept');
        if (deptSel) {
            user.profile.arbpl = deptSel.value; // Store ID (e.g. 5129)
            user.profile.department = deptSel.options[deptSel.selectedIndex] ? deptSel.options[deptSel.selectedIndex].text : ""; // Store Name
        }
        user.profile.team = getVal('user-team');
        user.profile.webPW = getVal('user-webpw');
        user.profile.pw2 = getVal('user-pw2');
        user.profile.sapPW = getVal('user-sappw');
        // New Setting: Inspector Display Format Preference
        user.profile.erpFormat = getVal('erp-format-pref');


        // Gather Workers with Logic
        const newWorkers = [];
        const rows = document.querySelectorAll('#worker-table tbody tr');

        // Constraint Checking Maps
        const managers = [];
        const drivers = { 'A조': { '정': 0, '부': 0 }, 'B조': { '정': 0, '부': 0 }, 'C조': { '정': 0, '부': 0 }, 'D조': { '정': 0, '부': 0 }, '일근': { '정': 0, '부': 0 } };

        rows.forEach(row => {
            const inputs = row.querySelectorAll('input, select');
            const team = inputs[2].value;
            const isManager = inputs[4].checked; // Checkbox
            const driverRole = inputs[5].value; // Select

            if (isManager) managers.push(row);
            if (driverRole !== '-' && drivers[team]) {
                drivers[team][driverRole]++;
            }

            newWorkers.push({
                name: inputs[0].value,
                id: inputs[1].value,
                team: team,
                phone: inputs[3].value,
                isManager: isManager ? 1 : 0,
                driverRole: driverRole
            });
        });

        // Enforce Manager Constraint (Last one checked wins, handled by UI event usually, but here we validate state)
        // Actually, UI event is better for UX. But let's assume the state is what it is.
        // If multiple managers are checked, we might warn or just save.
        // User requested: "Only 1 manager possible". Let's enforce in UI changes mostly.

        if (!appConfig.appSettings) appConfig.appSettings = {};
        appConfig.appSettings.colleagues = newWorkers;

        // Gather Locations (Global)
        const newLocs = [];
        document.querySelectorAll('#location-table tbody tr').forEach(row => {
            const inputs = row.querySelectorAll('input, select');
            newLocs.push({
                name: inputs[0].value,
                order: inputs[1].value,
                type: inputs[2].value
            });
        });

        if (!appConfig.appSettings) appConfig.appSettings = {};
        appConfig.appSettings.locations = newLocs;

        // Gather Hotkeys
        const newHotkeys = [];
        document.querySelectorAll('#hotkey-table tbody tr').forEach(row => {
            const action = row.dataset.action;
            const key = row.dataset.key;
            const desc = row.dataset.desc;
            const enabled = row.querySelector('input[type="checkbox"]').checked;
            newHotkeys.push({ action, key, desc, enabled });
        });
        user.hotkeys = newHotkeys;

        // --- NEW: Gather Daily Log Defaults ---
        const dailyLogDefaults = {
            safety: [],
            driverCheck: document.getElementById('fps-driver-check-toggle') ? document.getElementById('fps-driver-check-toggle').checked : false,
            generalWork: [],
            safetyEdu: currentSafetyEduData, // [New]
            autoInputTime: getVal('fps-auto-input-time')
        };

        // 1. Safety (4 Sets)
        for (let i = 1; i <= 4; i++) {
            dailyLogDefaults.safety.push({
                content: getVal(`fps-safe${i}-content`),
                start: getVal(`fps-safe${i}-start`),
                end: getVal(`fps-safe${i}-end`)
            });
        }

        // 3. General Work
        document.querySelectorAll('#general-work-table tbody tr').forEach(row => {
            const selects = row.querySelectorAll('select');
            const inputs = row.querySelectorAll('input');
            dailyLogDefaults.generalWork.push({
                workType: selects[0].value,
                category: selects[1].value,
                content: inputs[0].value,
                manager: inputs[1].value,
                start: inputs[2].value,
                end: inputs[3].value
            });
        });

        user.dailyLogDefaults = dailyLogDefaults;

        sendMessageToAHK({ command: 'saveConfig', data: appConfig });
    } catch (e) {
        showNativeMsgBox("설정 저장 중 오류: " + e.message);
    }
}

// --- Worker Logic ---
function handleManagerCheck(checkbox) {
    if (checkbox.checked) {
        // Uncheck all others
        const allChecks = document.querySelectorAll('#worker-table input[type="checkbox"]');
        allChecks.forEach(cb => {
            if (cb !== checkbox) cb.checked = false;
        });
    }
    autoSaveSettings();
}

function handleDriverChange(select) {
    // Unique Driver Per Team Logic?
    // "Each Team can have 1 Jung, 1 Bu". 
    // Complexity: We need to know the team of this row.
    const row = select.closest('tr');
    const teamSelect = row.querySelector('select:first-of-type'); // First select is Team
    const team = teamSelect.value;
    const role = select.value;

    if (role !== '-') {
        // Check if another row with same team has same role
        const rows = document.querySelectorAll('#worker-table tbody tr');
        for (const r of rows) {
            if (r === row) continue;
            const t = r.querySelector('select:first-of-type').value;
            const d = r.querySelector('select:last-of-type').value; // Last select is Driver
            if (t === team && d === role) {
                // Conflict
                showNativeMsgBox(`${team}에는 이미 ${role}운전원이 있습니다.`);
                select.value = '-'; // Revert
                return;
            }
        }
    }
    autoSaveSettings();
}

function addWorkerRow(data = {}) {
    const tbody = document.querySelector('#worker-table tbody');
    // Fix: Check if data has keys (loaded config) or is empty (new row button)
    const rowData = Object.keys(data).length > 0 ? data : { name: '', id: '', team: '일근', phone: '', isManager: 0, driverRole: '-' };
    addWorkerRowToTable(tbody, rowData);
    // Auto scroll to bottom of the settings content area
    const contentArea = document.querySelector('.settings-content-area');
    if (contentArea) contentArea.scrollTop = contentArea.scrollHeight;
}

function addWorkerRowToTable(tbody, data) {
    const tr = document.createElement('tr');

    // Team Options
    const teams = ['A조', 'B조', 'C조', 'D조', '일근'];
    let teamOpts = teams.map(t => `<option value="${t}" ${data.team === t ? 'selected' : ''}>${t}</option>`).join('');

    // Driver Options
    const drivers = ['-', '정', '부'];
    let driverOpts = drivers.map(d => `<option value="${d}" ${data.driverRole === d ? 'selected' : ''}>${d}</option>`).join('');

    tr.innerHTML = `
        <td><input type="text" value="${data.name || ''}" placeholder="이름" oninput="autoSaveSettings()"></td>
        <td><input type="text" value="${data.id || ''}" placeholder="사번" maxlength="6" oninput="this.value=this.value.replace(/[^0-9]/g,''); autoSaveSettings()"></td>
        <td><select onchange="autoSaveSettings()">${teamOpts}</select></td>
        <td><input type="text" value="${data.phone || ''}" placeholder="   -    -    " maxlength="13" oninput="formatPhone(this); autoSaveSettings()"></td>
        <td class="center"><input type="checkbox" ${data.isManager ? 'checked' : ''} onchange="handleManagerCheck(this)"></td>
        <td class="center"><select onchange="handleDriverChange(this)">${driverOpts}</select></td>
        <td class="center"><button class="small-btn danger" onclick="this.closest('tr').remove(); autoSaveSettings()">X</button></td>
    `;
    tbody.appendChild(tr);
}

// --- Location Logic ---
function addLocationRow(data = {}) {
    const tbody = document.querySelector('#location-table tbody');
    // Fix: Check if data has keys (loaded config) or is empty (new row button)
    const rowData = Object.keys(data).length > 0 ? data : { name: '', order: '', type: '기타업무' };
    addLocationRowToTable(tbody, rowData);
    // Auto scroll to bottom of the settings content area
    const contentArea = document.querySelector('.settings-content-area');
    if (contentArea) contentArea.scrollTop = contentArea.scrollHeight;
}

function addLocationRowToTable(tbody, data) {
    const tr = document.createElement('tr');

    // Type Options
    const types = ['변전소', '전기실(그룹1)', '전기실(그룹2)', '전기실(그룹3)', '기타업무'];
    let typeOpts = types.map(t => `<option value="${t}" ${data.type === t ? 'selected' : ''}>${t}</option>`).join('');

    tr.innerHTML = `
        <td><input type="text" value="${data.name || ''}" placeholder="점검명" oninput="autoSaveSettings()"></td>
        <td><input type="text" value="${data.order || ''}" placeholder="오더번호" maxlength="8" oninput="this.value=this.value.replace(/[^0-9]/g,''); autoSaveSettings()"></td>
        <td><select onchange="autoSaveSettings()">${typeOpts}</select></td>
        <td class="center"><button class="small-btn danger" onclick="this.closest('tr').remove(); autoSaveSettings()">X</button></td>
    `;
    tbody.appendChild(tr);
}

// --- Hotkey Logic ---
const defaultHotkeys = [
    { action: "AutoLogin", key: "#z", desc: "자동 로그인" },
    { action: "OpenLog", key: "#!z", desc: "업무일지 실행" },
    { action: "ConvertExcel", key: "#!a", desc: "일반업무 -> 엑셀 변환" },
    { action: "CopyExcel", key: "#!c", desc: "엑셀 데이터 복사" },
    { action: "PasteExcel", key: "#!v", desc: "일반업무에 붙여넣기" },
    { action: "ForceExit", key: "^Esc", desc: "강제 종료" }
];

const hotkeyTranslations = {
    "AutoLogin": "자동 로그인",
    "OpenLog": "업무일지 실행",
    "ConvertExcel": "엑셀 변환",
    "CopyExcel": "엑셀 복사",
    "PasteExcel": "붙여넣기",
    "ForceExit": "강제 종료"
};

function renderHotkeyTable(savedHotkeys) {
    const tbody = document.querySelector('#hotkey-table tbody');
    tbody.innerHTML = '';

    defaultHotkeys.forEach(def => {
        // Find saved state
        const saved = savedHotkeys.find(h => h.action === def.action);
        const isEnabled = saved ? saved.enabled : true;

        const tr = document.createElement('tr');
        tr.dataset.action = def.action;
        tr.dataset.key = def.key;
        tr.dataset.desc = def.desc;

        // Render Keycap
        const keyHtml = renderKeycap(def.key);
        // Translate Action Name
        const actionName = hotkeyTranslations[def.action] || def.action;

        tr.innerHTML = `
            <td>${actionName}</td>
            <td>${keyHtml}</td>
            <td class="desc-cell" title="${def.desc}">${def.desc}</td>
            <td class="center"><input type="checkbox" ${isEnabled ? 'checked' : ''} onchange="autoSaveSettings()"></td>
        `;
        tbody.appendChild(tr);
    });
}

function renderKeycap(keyStr) {
    // Replace modifiers with <kbd>
    let html = keyStr;
    // Better rendering with less spacing
    let parts = [];
    if (keyStr.includes('#')) parts.push('<kbd>Win</kbd>');
    if (keyStr.includes('^')) parts.push('<kbd>Ctrl</kbd>');
    if (keyStr.includes('!')) parts.push('<kbd>Alt</kbd>');
    if (keyStr.includes('+')) parts.push('<kbd>Shift</kbd>');

    // Extract the main key (strip modifiers)
    let mainKey = keyStr.replace(/[#^!+]/g, '');
    parts.push(`<kbd>${mainKey.toUpperCase()}</kbd>`);

    return parts.join('+'); // Removed spaces around + for tighter look
}

// --- Preset Logic ---
function renderPresetList(presetsMap) {
    const sel = document.getElementById('track-preset-sel'); // Fixed ID
    if (!sel) return; // Guard against missing element
    sel.innerHTML = '<option value="">(새 프리셋)</option>';
    if (!presetsMap) return;

    // If presetsMap is array (from v3 structure update?) or object
    // Assuming object for now based on legacy code or map
    // Check if array
    let list = Array.isArray(presetsMap) ? presetsMap : Object.keys(presetsMap).map(k => presetsMap[k]);

    list.forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.name;
        opt.text = p.name;
        sel.appendChild(opt);
    });
}

// --- Utils ---
function getVal(id) {
    const el = document.getElementById(id);
    return el ? el.value : '';
}
function setVal(id, val) {
    const el = document.getElementById(id);
    if (el) el.value = val !== undefined ? val : '';
}

function showNativeMsgBox(text, title = "알림") {
    sendMessageToAHK({ command: 'msgbox', text: text, title: title });
}

function formatPhone(input) {
    let value = input.value.replace(/[^0-9]/g, '');
    let formatted = '';

    if (value.length < 4) {
        formatted = value;
    } else if (value.length < 7) {
        formatted = value.substr(0, 3) + '-' + value.substr(3);
    } else if (value.length < 11) {
        // 010-123-4567 (3-3-4)
        formatted = value.substr(0, 3) + '-' + value.substr(3, 3) + '-' + value.substr(6);
    } else {
        // 010-1234-5678 (3-4-4)
        formatted = value.substr(0, 3) + '-' + value.substr(3, 4) + '-' + value.substr(7);
    }

    // Safety cut
    if (formatted.length > 13) formatted = formatted.substr(0, 13);

    input.value = formatted;

    // Auto Save is handled by global delegate on 'input' event
}

function formatTime(input) {
    let value = input.value.replace(/[^0-9]/g, '');
    // Safety cut for HHmm (4 digits)
    if (value.length > 4) value = value.substr(0, 4);

    let formatted = '';
    if (value.length < 3) {
        formatted = value;
    } else {
        // HH:mm
        formatted = value.substr(0, 2) + ':' + value.substr(2);
    }
    input.value = formatted;

    // Auto Save is handled by global delegate on 'input' event
}

// --- ERP Check Logic ---
let selectedERPLocation = null;

function renderERPCheck() {
    // 1. Get global locations
    const locations = (appConfig.appSettings && appConfig.appSettings.locations) ? appConfig.appSettings.locations : [];

    // [State Preservation] 렌더링 전 현재 상태 저장 (선택된 장소, 체크표시)
    const savedSelectedLocation = selectedERPLocation;
    const savedCheckmarks = new Set();
    const existingButtons = document.querySelectorAll('.erp-btn[data-order-num]');
    existingButtons.forEach(btn => {
        const checkMark = btn.querySelector('.order-check-mark');
        if (checkMark) {
            const orderNum = btn.dataset.orderNum;
            if (orderNum) {
                savedCheckmarks.add(orderNum);
            }
        }
    });

    // 2. Clear containers
    const gridSub = document.getElementById('grid-substation');
    const gridEtc = document.getElementById('grid-etc');
    const elG1 = document.getElementById('elec-g1');
    const elG2 = document.getElementById('elec-g2');
    const elG3 = document.getElementById('elec-g3');

    if (gridSub) gridSub.innerHTML = '';
    if (gridEtc) gridEtc.innerHTML = '';
    if (elG1) elG1.innerHTML = '';
    if (elG2) elG2.innerHTML = '';
    if (elG3) elG3.innerHTML = '';

    // 선택된 장소는 나중에 복원하므로 여기서는 null로 설정하지 않음
    // selectedERPLocation = null; // Reset selection - 주석 처리

    // 3. Process Items
    locations.forEach(loc => {
        const btn = document.createElement('div');
        btn.className = 'erp-btn';
        // Use TextNode to avoid overwriting span later if we appended
        btn.appendChild(document.createTextNode(loc.name));

        btn.dataset.locName = loc.name;
        btn.dataset.orderNum = loc.order; // Store Order Number for lookup
        btn.onclick = () => selectERPLoc(btn, loc.name);

        // [State Preservation] 저장된 체크표시 복원
        if (savedCheckmarks.has(loc.order)) {
            const checkSpan = document.createElement('span');
            checkSpan.className = 'order-check-mark';
            checkSpan.textContent = '\u2714\uFE0E'; // ✔ (Heavy Check Mark) + Text Presentation Selector
            checkSpan.style.color = '#00C853'; // Vibrant Green
            checkSpan.style.marginRight = '5px';
            checkSpan.style.fontWeight = 'bold';
            btn.prepend(checkSpan);
        }

        // [State Preservation] 저장된 선택 상태 복원
        if (savedSelectedLocation === loc.name) {
            btn.classList.add('selected');
            selectedERPLocation = loc.name;
        }

        if (loc.type === '변전소') {
            // Default Status Dot (Yellow)
            const dot = document.createElement('span');
            dot.className = 'status-dot';
            dot.style.fontSize = '1.2em';
            dot.style.marginLeft = '5px';
            dot.style.fontWeight = 'bold';
            dot.style.color = '#FFC107'; // Amber/Yellow
            dot.innerText = '●';
            dot.title = "데이터 확인 중..."; // Initial tooltip
            btn.appendChild(dot);

            gridSub.appendChild(btn);
        } else if (loc.type === '기타업무') {
            gridEtc.appendChild(btn);
        } else if (loc.type.startsWith('전기실')) {
            // Group Matching
            if (loc.type.includes('그룹1')) {
                elG1.appendChild(btn);
            } else if (loc.type.includes('그룹2')) {
                elG2.appendChild(btn);
            } else if (loc.type.includes('그룹3')) {
                elG3.appendChild(btn);
            } else {
                gridEtc.appendChild(btn);
            }
        }
    });

    // [State Preservation] 저장된 선택 장소가 더 이상 목록에 없는 경우 초기화
    if (savedSelectedLocation && !locations.find(loc => loc.name === savedSelectedLocation)) {
        selectedERPLocation = null;
    }

    // Apply cached status if available
    handleERPStatusUpdate(null);
}

// ... (SelectERPLoc, Toggle, Run logic skipped/unchanged) ...

// --- ERP Status Update (Polling) ---
let latestERPStatus = null;

function handleERPStatusUpdate(statusMap) {
    if (statusMap) {
        latestERPStatus = statusMap;
    } else if (latestERPStatus) {
        statusMap = latestERPStatus;
    } else {
        return; // Keep Yellow
    }

    const btns = document.querySelectorAll('#grid-substation .erp-btn[data-loc-name]');

    btns.forEach(btn => {
        const locName = btn.dataset.locName;
        const dot = btn.querySelector('.status-dot');

        if (dot) {
            // AHK JSON serialization might send 1 instead of true
            const isUpdated = statusMap[locName] ? true : false;
            dot.style.color = isUpdated ? '#4CAF50' : '#FF0000'; // Green vs Red
            dot.title = isUpdated ? "점검 완료 (오늘)" : "점검 미완료";

            // Re-assert visibility (in case)
            dot.style.display = 'inline';
        }
    });

    // 4차 요구사항(일괄모드 연동)
    if (erpBatchAppInstance && typeof erpBatchAppInstance.latestStatusMap !== 'undefined') {
        // Vue3 반응성 시스템 트리거를 위해 객체를 완전히 새로 할당
        erpBatchAppInstance.latestStatusMap = { ...(statusMap || {}) };
    }
}

window.handleERPStatusUpdate = handleERPStatusUpdate;

function handleERPOrderListUpdate(orders) {
    if (!orders || !Array.isArray(orders)) return;

    // 1. Update Timestamp
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const min = String(now.getMinutes()).padStart(2, '0');
    const timeStr = `${yyyy}-${mm}-${dd} ${hh}:${min}`;

    document.getElementById('erp-last-update').innerText = timeStr;

    // 2. Mark Buttons
    orders.forEach(orderNum => {
        // Find buttons with this order number
        // Attribute selector queries need quotes if value contains special chars, but orderNum is safe digits usually.
        const btns = document.querySelectorAll(`.erp-btn[data-order-num="${orderNum}"]`);

        btns.forEach(btn => {
            // Check if already has checkmark to avoid double adding
            // We use a specific class for the checkmark span
            if (!btn.querySelector('.order-check-mark')) {
                const checkSpan = document.createElement('span');
                checkSpan.className = 'order-check-mark';
                checkSpan.textContent = '\u2714\uFE0E'; // ✔ (Heavy Check Mark) + Text Presentation Selector
                checkSpan.style.color = '#00C853'; // Vibrant Green
                checkSpan.style.marginRight = '5px';
                checkSpan.style.fontWeight = 'bold';

                // Prepend to the button content
                btn.prepend(checkSpan);
            }
        });
    });

    // 4차 요구사항(일괄모드 연동)
    if (erpBatchAppInstance && typeof erpBatchAppInstance.completedOrders !== 'undefined') {
        // Vue3 배열 반응성을 위해 새로운 배열 인스턴스 할당
        erpBatchAppInstance.completedOrders = orders ? [...orders] : [];
    }
}

function selectERPLoc(btn, locName) {
    // Deselect all
    document.querySelectorAll('.erp-btn').forEach(el => el.classList.remove('selected'));

    // Select this
    btn.classList.add('selected');
    selectedERPLocation = locName;
}

let erpBatchAppInstance = null;
let isBatchMode = false;

function toggleERPMode() {
    const btn = document.getElementById('btn-erp-mode');
    const indContainer = document.getElementById('erp-individual-container');
    const batchContainer = document.getElementById('erp-batch-container');

    if (btn.innerText.includes('일괄모드')) {
        // 개별 -> 일괄
        btn.innerText = '< 개별모드';
        indContainer.style.display = 'none';
        batchContainer.style.display = 'flex';
        isBatchMode = true;

        // Vue App 초기화 (지연 마운트)
        if (!erpBatchAppInstance) {
            initERPBatchApp();
        } else {
            // 이미 마운트된 경우, 데이터를 최신 워커 목록으로 갱신
            updateERPBatchData();
        }
    } else {
        // 일괄 -> 개별
        btn.innerText = '일괄모드 >';
        batchContainer.style.display = 'none';
        indContainer.style.display = 'flex';
        isBatchMode = false;
    }
}

function initERPBatchApp() {
    const { createApp } = Vue;
    erpBatchAppInstance = createApp({
        data() {
            return {
                locations: [],
                workers: [],
                activeRows: [], // Array of location names
                selections: {},  // { locName: [workerId1, workerId2, ...] }
                latestStatusMap: {},
                completedOrders: []
            };
        },
        mounted() {
            this.syncData();
        },
        methods: {
            syncData() {
                // appConfig에서 변전소, 전기실, 기타업무 등 장소 목록 합치기
                const locs = [];
                if (appConfig && appConfig.appSettings && appConfig.appSettings.locations) {
                    appConfig.appSettings.locations.forEach(loc => {
                        let shortType = loc.type || '';
                        if (shortType.startsWith('전기실')) shortType = '전기실';
                        locs.push({ name: loc.name, type: shortType, order: loc.order });
                    });
                }
                this.locations = locs;

                // 로그인 유저 목록 또는 설정된 분소원 목록에서 워커 데이터 추출
                this.workers = [];
                const uid = selectedUserId;
                if (uid && appConfig && appConfig.users && appConfig.users[uid]) {
                    const userProfile = appConfig.users[uid].profile || {};
                    const myTeam = userProfile.team || '';
                    const allColleagues = appConfig.appSettings?.colleagues || [];

                    // 필터링: 모달팝업(개별모드)과 동일하게 본인 부서만 표출
                    const filteredWorkers = allColleagues.filter(w => {
                        if (myTeam && w.team === myTeam) return true;
                        return false;
                    });

                    // 정렬 로직 적용 (분소장 1순위, 사번순)
                    const sortedWorkers = [...filteredWorkers].sort((a, b) => {
                        if (a.isManager !== b.isManager) return b.isManager - a.isManager;
                        return a.id.localeCompare(b.id);
                    });

                    this.workers = sortedWorkers.map(w => ({ id: w.id, name: w.name }));
                }

                // [버그 수정] 인스턴스 초기화 시, 이미 캐싱된 ERP 신호등 데이터(latestERPStatus)를 불러와 Vue 인스턴스에 주입
                if (typeof latestERPStatus !== 'undefined' && latestERPStatus) {
                    this.latestStatusMap = { ...latestERPStatus };
                }
            },
            isRowActive(locName) {
                return this.activeRows.includes(locName);
            },
            toggleRow(locName) {
                const idx = this.activeRows.indexOf(locName);
                if (idx > -1) {
                    // 비활성화
                    this.activeRows.splice(idx, 1);
                    // 연쇄적으로 해당 행의 선택 내역 지우기
                    delete this.selections[locName];
                } else {
                    // 활성화
                    this.activeRows.push(locName);
                    // 요구사항: 점검장소가 눌러져서 행 활성화 시 점검자 버튼은 모두 눌러진 상태가 기본값
                    this.selections[locName] = this.workers.map(w => w.id);
                }
            },
            isWorkerSelected(locName, workerId) {
                if (!this.selections[locName]) return false;
                return this.selections[locName].includes(workerId);
            },
            toggleWorker(locName, workerId) {
                if (!this.isRowActive(locName)) return; // 방어 코드

                if (!this.selections[locName]) {
                    this.selections[locName] = [];
                }

                const selArr = this.selections[locName];
                const idx = selArr.indexOf(workerId);

                if (idx > -1) {
                    selArr.splice(idx, 1);
                } else {
                    selArr.push(workerId);
                }
            },
            getBatchData() {
                // AHK에 보낼 형태로 데이터 정제
                const result = [];
                for (const locName of this.activeRows) {
                    const selectedWorkerIds = this.selections[locName] || [];
                    result.push({
                        location: locName,
                        workerIds: selectedWorkerIds
                    });
                }
                return result;
            }
        }
    }).mount('#erp-batch-container');
}

function updateERPBatchData() {
    if (erpBatchAppInstance && erpBatchAppInstance.syncData) {
        erpBatchAppInstance.syncData();
    }
}

function openERPHandler() {
    if (isBatchMode) {
        if (!erpBatchAppInstance) return;
        const batchData = erpBatchAppInstance.getBatchData();
        if (batchData.length === 0) {
            showNativeMsgBox("선택된 점검장소가 없습니다.");
            return;
        }

        if (!erpBatchModalInstance) {
            initERPBatchModalApp();
        }
        
        // 동기화
        erpBatchModalInstance.openModal(batchData);
    } else {
        openERPWorkerModal(); // 개별모드 팝업창
    }
}

let erpBatchModalInstance = null;
function initERPBatchModalApp() {
    const { createApp } = Vue;
    erpBatchModalInstance = createApp({
        data() {
            return {
                batchDataRaw: [],
                batchPreviewList: [],
                isListFormat: false,
                previewMaxCount: 0
            };
        },
        methods: {
            openModal(rawData) {
                this.batchDataRaw = rawData;
                const toggle = document.getElementById('toggle-worker-format');
                if (toggle) this.isListFormat = toggle.checked;
                this.updatePreviewList();
                document.getElementById('erp-batch-modal').style.display = 'flex';
            },
            closeModal() {
                document.getElementById('erp-batch-modal').style.display = 'none';
            },
            toggleFormat() {
                this.isListFormat = !this.isListFormat;
                // Update toggle state to Main Tab if needed, but keeping it simple for modal interaction
                this.updatePreviewList();
            },
            updatePreviewList() {
                // 부서 이름 최대 글자수 측정
                let maxLocLen = 0;
                this.batchDataRaw.forEach(item => {
                    if (item.location.length > maxLocLen) maxLocLen = item.location.length;
                });

                let maxLen = 0;
                this.batchPreviewList = this.batchDataRaw.map(item => {
                    const workerNames = [];
                    item.workerIds.forEach(id => {
                        const w = erpBatchAppInstance.workers.find(wk => wk.id === id);
                        if (w) workerNames.push(w.name);
                    });
                    if (workerNames.length > maxLen) maxLen = workerNames.length;

                    let workersText = '';
                    if (workerNames.length === 0) {
                        workersText = '(선택 없음)';
                    } else if (this.isListFormat) {
                        workersText = workerNames.join(', ');
                    } else {
                        if (workerNames.length === 1) {
                            workersText = workerNames[0];
                        } else {
                            workersText = `${workerNames[0]} 외 ${workerNames.length - 1}명`;
                        }
                    }

                    // 하이픈 정렬을 위한 띄어쓰기 패딩
                    const paddingSpaces = ' '.repeat(maxLocLen - item.location.length);
                    const paddedLocName = ` ${paddingSpaces}${item.location} `;

                    return {
                        locName: paddedLocName,
                        workersText: workersText
                    };
                });
                this.previewMaxCount = (maxLen > 1) ? (maxLen - 1) : 0;
            },
            submitBatchTask() {
                // Return structured Array of tasks mapping to original runTask
                const format = this.isListFormat ? 'list' : 'summary';

                const formattedBatchData = this.batchDataRaw.map((item, index) => {
                    // targetType, targetOrder 속성 찾기 (appConfig.appSettings.locations 활용)
                    let locType = "";
                    let locOrder = "";
                    if (appConfig && appConfig.appSettings && appConfig.appSettings.locations) {
                        const locObj = appConfig.appSettings.locations.find(l => l.name === item.location);
                        if (locObj) {
                            locType = locObj.type || "";
                            locOrder = locObj.order || "";
                        }
                    }

                    // members 배열 완성 (id를 기반으로 실제 이름 찾기)
                    const workerNames = [];
                    item.workerIds.forEach(id => {
                        const w = erpBatchAppInstance.workers.find(wk => wk.id === id);
                        if (w) workerNames.push(w.name);
                    });

                    return {
                        location: item.location,
                        workersStr: this.batchPreviewList[index].workersText, // 알림 메시지 팝업에서의 목록 표시용 보존
                        targetType: locType,
                        targetOrder: locOrder,
                        members: workerNames
                    };
                });

                sendMessageToAHK({
                    command: 'runTask',
                    task: 'ERPCheck',
                    batchmode: true,
                    format: format,
                    location: formattedBatchData
                });
                this.closeModal();
            }
        }
    }).mount('#erp-batch-modal');
}



// --- Work Log Logic ---
function renderWorkLogUI() {
    const uid = selectedUserId;
    if (!uid || !appConfig.users || !appConfig.users[uid]) return;

    // 1. Determine Work Type (Day/Night) based on Time
    const now = new Date();
    const dayOfWeek = now.getDay(); // 0=Sun, 6=Sat
    const hours = now.getHours();
    const minutes = now.getMinutes();
    const timeVal = hours * 100 + minutes;

    // Logic: Day if Mon-Fri (1-5) AND 08:30 <= Time < 17:30 [Legacy - Removed]
    // Now handled entirely by updateShiftUI logic via AHK broadcast.

    // 2. Render Worker List (Must be before handleWorkTypeChange for uncheck logic to work)
    renderWorkLogWorkerList();

    // 3. Apply Automation Options based on Work Type
    handleWorkTypeChange(false); // Validates and sets checkboxes

    // 3. Initialize Safety Log Dates (Legacy Logic: -510 mins for Day shift boundary)
    const adjDate = new Date(Date.now() - 510 * 60000);
    const yyyy = adjDate.getFullYear();
    const mm = String(adjDate.getMonth() + 1).padStart(2, '0');
    const dd = String(adjDate.getDate()).padStart(2, '0');
    const dateStr = `${yyyy}${mm}${dd}`;

    isWorkLogInitialized = true;
}

function handleWorkTypeChange(skipRenderWorkers = true) {
    const isDay = document.querySelector('input[name="work-type"][value="day"]').checked;

    // Checkboxes
    const chkMakeLog = document.getElementById('chk-make-log');
    const chkGeneral = document.getElementById('chk-general');
    const chkSafe = document.getElementById('chk-safe-manage');
    const chkDriving = document.getElementById('chk-driving');
    const chkDrink = document.getElementById('chk-drink');
    const chkCal = document.getElementById('chk-drink-cal');

    if (isDay) {
        // Day Mode
        chkMakeLog.checked = true;
        chkMakeLog.disabled = false;

        chkDriving.checked = false;
        chkDriving.disabled = true; // Disabled for Day

        chkDrink.disabled = false; // Enabled for Day

        // Calibration check state depends on Drink check
        if (chkDrink.checked) {
            chkCal.disabled = false;
        } else {
            chkCal.disabled = true;
        }

    } else {
        // Night Mode
        chkMakeLog.checked = false;
        chkMakeLog.disabled = true; // Disabled for Night

        chkDriving.checked = true;
        chkDriving.disabled = false;

        chkDrink.disabled = true;
        chkDrink.checked = false;
        chkCal.disabled = true;
        chkCal.checked = false;

        // [New] 야간 근무 시 일근자 체크 해제
        uncheckDayShiftWorkers();
    }

    // [New] 주간 모드로 변경 시 -> 체크 복구
    if (isDay) {
        checkDayShiftWorkers();
    }

    enableDriverSelects(!isDay);

    // Toggle Driver Column Visibility
    const drvHeader = document.getElementById('col-header-drive');
    if (drvHeader) {
        if (isDay) {
            drvHeader.classList.add('hidden-col');
        } else {
            drvHeader.classList.remove('hidden-col');
        }
    }

    // Toggle Cells
    document.querySelectorAll('.w-drive-cell').forEach(cell => {
        if (isDay) {
            cell.classList.add('hidden-col');
        } else {
            cell.classList.remove('hidden-col');
        }
    });

    // --- [New] Apply Safety Management Presets & Dates ---
    if (typeof appConfig !== 'undefined' && appConfig.users && typeof selectedUserId !== 'undefined' && appConfig.users[selectedUserId]) {
        const user = appConfig.users[selectedUserId];
        const defaults = user.dailyLogDefaults || {};
        const safetyPresets = defaults.safety || []; // Array of 4 items

        // Date Calculation
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        const todayStr = `${yyyy}${mm}${dd}`;

        const tmr = new Date(now);
        tmr.setDate(tmr.getDate() + 1);
        const t_yyyy = tmr.getFullYear();
        const t_mm = String(tmr.getMonth() + 1).padStart(2, '0');
        const t_dd = String(tmr.getDate()).padStart(2, '0');
        const tomorrowStr = `${t_yyyy}${t_mm}${t_dd}`;

        // Select Presets based on Mode
        // Day: idx 0 (Morning), idx 1 (Afternoon)
        // Night: idx 2 (Evening), idx 3 (Dawn)
        let p1, p2;
        let d1, d2;

        if (isDay) {
            p1 = safetyPresets[0] || {};
            p2 = safetyPresets[1] || {};
            d1 = todayStr;
            d2 = todayStr;
        } else {
            p1 = safetyPresets[2] || {};
            p2 = safetyPresets[3] || {};
            d1 = todayStr;
            d2 = tomorrowStr; // Dawn is next day
        }

        // Apply to Row 1
        setVal('safe-content-1', p1.content || '');
        setVal('safe-start-1', p1.start || '');
        setVal('safe-end-1', p1.end || '');
        setVal('safe-date-1', d1);

        // Apply to Row 2
        setVal('safe-content-2', p2.content || '');
        setVal('safe-start-2', p2.start || '');
        setVal('safe-end-2', p2.end || '');
        setVal('safe-date-2', d2);
    }
}

function toggleDrinkCalibration() {
    const chkDrink = document.getElementById('chk-drink');
    const chkCal = document.getElementById('chk-drink-cal');
    if (chkDrink.checked) {
        chkCal.disabled = false;
    } else {
        chkCal.disabled = true;
        chkCal.checked = false;
    }
}

function renderWorkLogWorkerList() {
    const container = document.getElementById('work-log-worker-list');

    // [State Preservation] 렌더링 전 현재 상태 저장 (체크박스, 휴가사유, 운전원 선택)
    const savedState = {};
    const existingRows = container.querySelectorAll('.worker-row');
    existingRows.forEach(row => {
        const workerId = row.dataset.id;
        if (workerId) {
            const chk = row.querySelector('.w-chk');
            const note = row.querySelector('.w-note');
            const drv = row.querySelector('.w-drive');
            savedState[workerId] = {
                checked: chk ? chk.checked : false,
                note: note ? note.value : '',
                driver: drv ? drv.value : ''
            };
        }
    });

    container.innerHTML = '';
    // const countSpan = document.getElementById('worker-count'); // Removed

    // Get Current User ID & Profile
    const uid = selectedUserId;
    if (!uid || !appConfig.users || !appConfig.users[uid]) return;
    const user = appConfig.users[uid];

    // Use Global Colleagues
    const allColleagues = appConfig.appSettings?.colleagues || [];
    const myTeam = (user.profile && user.profile.team) ? user.profile.team : '';

    // Filter: Same Team OR '일근'
    const filteredWorkers = allColleagues.filter(w => {
        if (w.team === '일근') return true;
        if (myTeam && w.team === myTeam) return true;
        return false;
    });

    // Sort: Manager First (isManager=1), then ID (asc)
    const sortedWorkers = [...filteredWorkers].sort((a, b) => {
        if (a.isManager !== b.isManager) return b.isManager - a.isManager; // 1 before 0
        return a.id.localeCompare(b.id);
    });

    sortedWorkers.forEach((worker, index) => {
        const row = document.createElement('div');
        row.className = 'worker-row';
        row.dataset.id = worker.id;
        row.dataset.team = worker.team; // [중요] 근무조 데이터셋 추가
        row.dataset.name = worker.name;
        row.dataset.isManager = worker.isManager;

        // [State Preservation] 저장된 상태가 있으면 사용, 없으면 기본값 (모두 체크)
        const saved = savedState[worker.id];
        const isChecked = saved ? saved.checked : true;
        const savedNote = saved ? saved.note : '';
        const savedDriver = saved ? saved.driver : '';

        // Dynamic Driver Column
        const isDay = document.querySelector('input[name="work-type"][value="day"]').checked;
        const driveClass = isDay ? 'wl-drive hidden-col w-drive-cell' : 'wl-drive w-drive-cell';

        row.innerHTML = `
            <label class="wl-name checkbox-label" style="margin: 0;" for="chk-worker-${worker.id}">
                <input type="checkbox" ${isChecked ? 'checked' : ''} class="w-chk" id="chk-worker-${worker.id}">
                <span>${worker.name}</span>
            </label>
            <div class="wl-note"><input type="text" placeholder="사유" class="w-note" id="note-worker-${worker.id}"></div>
            <div class="${driveClass}">
                <select class="w-drive" disabled id="drv-worker-${worker.id}">
                    <option value="">-</option>
                    <option value="정" ${(savedDriver || worker.driverRole) === '정' ? 'selected' : ''}>정</option>
                    <option value="부" ${(savedDriver || worker.driverRole) === '부' ? 'selected' : ''}>부</option>
                    <option value="검사자" ${(savedDriver || worker.driverRole) === '검사자' ? 'selected' : ''}>검사자</option>
                </select>
            </div>
        `;
        container.appendChild(row);

        // [State Preservation] 저장된 휴가사유 복원 (innerHTML 후에 설정하여 안전하게 처리)
        const noteInput = row.querySelector('.w-note');
        if (noteInput && savedNote) {
            noteInput.value = savedNote;
        }

        // Add Listeners
        const chk = row.querySelector('.w-chk');
        const note = row.querySelector('.w-note');
        const drv = row.querySelector('.w-drive');

        // Logic: If Note has text (vacation), Uncheck.
        note.addEventListener('input', () => {
            if (note.value.trim() !== '' && note.value !== '일근') {
                chk.checked = false;
            } else {
                // Optional: Re-check if cleared?
                // chk.checked = true; 
            }
            updateWorkerStats();
        });

        chk.addEventListener('change', () => {
            updateWorkerStats();
            if (!chk.checked) drv.value = '';
        });
    });

    // Initial Driver Enable Check
    const isDay = document.querySelector('input[name="work-type"][value="day"]').checked;
    enableDriverSelects(!isDay);
}

function enableDriverSelects(enable) {
    document.querySelectorAll('.w-drive').forEach(el => {
        el.disabled = !enable;
    });
}

function toggleAllWorkers(mainChk) {
    const chks = document.querySelectorAll('#work-log-worker-list .w-chk');
    chks.forEach(c => c.checked = mainChk.checked);
    updateWorkerStats();
}

function updateWorkerStats() {
    // Placeholder for stats logic
}

function startWorkLog() {
    const uid = selectedUserId;
    if (!uid) return;

    const workType = document.querySelector('input[name="work-type"]:checked').value; // 'day' or 'night'

    // Gather Options
    const options = {
        makeLog: document.getElementById('chk-make-log').checked,
        general: document.getElementById('chk-general').checked,
        safe: document.getElementById('chk-safe-manage').checked,
        driving: document.getElementById('chk-driving').checked,
        driving: document.getElementById('chk-driving').checked,
        drink: document.getElementById('chk-drink').checked
    };

    // Gather Safety Data
    const safetyData = {
        content1: getVal('safe-content-1'),
        date1: getVal('safe-date-1'),
        start1: getVal('safe-start-1'),
        end1: getVal('safe-end-1'),
        content2: getVal('safe-content-2'),
        date2: getVal('safe-date-2'),
        start2: getVal('safe-start-2'),
        end2: getVal('safe-end-2')
    };

    // Gather Workers
    const workers = [];
    document.querySelectorAll('.worker-row').forEach(row => {
        const chk = row.querySelector('.w-chk');
        const note = row.querySelector('.w-note');
        const drv = row.querySelector('.w-drive');

        workers.push({
            name: row.dataset.name,
            id: row.dataset.id,
            attend: chk.checked,
            reason: note.value,
            driverRole: drv.value,
            team: row.dataset.team,
            boss: row.dataset.isManager
        });
    });

    const payload = {
        command: 'runTask',
        task: 'createWorkLog',
        data: {
            workType,
            options,
            safetyData,
            workers
        }
    };
    sendMessageToAHK(payload);
}


// Global Exports
window.switchMainTab = switchMainTab;
window.runTask = function (task) { sendMessageToAHK({ command: 'runTask', task: task }); };

// --- Safety Education Logic ---
function handleSafetyEduDayChange() {
    const ddl = document.getElementById('fps-safe-edu-day');
    const inp = document.getElementById('fps-safe-edu-content');
    if (!ddl || !inp) return;

    const day = ddl.value;
    if (currentSafetyEduData && currentSafetyEduData[day] !== undefined) {
        inp.value = currentSafetyEduData[day];
    } else {
        inp.value = '';
    }
}
// Removed window exports for handleSafety, attached internally

function handleSafetyEduContentChange() {
    const ddl = document.getElementById('fps-safe-edu-day');
    const inp = document.getElementById('fps-safe-edu-content');
    if (!ddl || !inp) return;

    const day = ddl.value;
    const val = inp.value;

    if (!currentSafetyEduData) currentSafetyEduData = {};
    currentSafetyEduData[day] = val;
    // Debounce is handled inside autoSaveSettings, but here we need to ensure the data is ready
    autoSaveSettings();
}
// Removed window exports for handleSafety, attached internally
window.handleSafetyEduContentChange = handleSafetyEduContentChange; // Expose for inline handlers if any
window.handleSafetyEduDayChange = handleSafetyEduDayChange; // Expose for inline handlers if any
window.openSettings = openSettings;
window.closeSettings = closeSettings;
window.switchSettingsTab = switchSettingsTab;
window.tryLogin = tryLogin;
window.deleteUser = deleteUser;
window.switchView = switchView;
window.submitNewUser = submitNewUser;
window.saveSettings = saveSettings;
window.addWorkerRow = addWorkerRow;
window.addLocationRow = addLocationRow;
window.minimizeWindow = function () { sendMessageToAHK({ command: 'minimize' }); };
window.closeWindow = function () { sendMessageToAHK({ command: 'close' }); };
window.autoSaveSettings = autoSaveSettings;
window.handleManagerCheck = handleManagerCheck;
window.handleDriverChange = handleDriverChange;

// --- Shift Status UI Update ---
function updateShiftUI(shiftData) {
    const panel = document.getElementById('shift-status-panel');
    if (!panel) return;

    panel.style.display = 'block';

    document.getElementById('shift-prev').textContent = shiftData.prev;
    document.getElementById('shift-current').textContent = shiftData.current;
    document.getElementById('shift-next').textContent = shiftData.next;

    // 주간/야간 텍스트 및 스타일
    const label = document.getElementById('shift-label-text');
    label.textContent = shiftData.shiftName; // "주간" or "야간"

    // 색상 등 시각적 구분
    if (shiftData.isNight) {
        label.style.color = '#5e6c84'; // Blue-ish gray

        // 야간: Prev, Next 위에 날짜 표시
        document.getElementById('shift-prev-date').textContent = shiftData.prevDate;
        document.getElementById('shift-current-date').textContent = "";
        document.getElementById('shift-next-date').textContent = shiftData.nextDate;

    } else {
        label.style.color = '#d97008'; // Orange-ish for Day

        // 주간: Current 위에 날짜 표시
        document.getElementById('shift-prev-date').textContent = "";
        document.getElementById('shift-current-date').textContent = shiftData.currDate;
        document.getElementById('shift-next-date').textContent = "";
    }

    // [New] Sync Work Log Radio Buttons
    const radioVal = shiftData.isNight ? 'night' : 'day';
    const radio = document.querySelector(`input[name="work-type"][value="${radioVal}"]`);
    if (radio && !radio.checked) {
        radio.checked = true;

        // 통합 UI 갱신 함수 호출 (작업자 명단, 프리셋, 옵션 등 모두 갱신)
        refreshUI();
    }
}


function uncheckDayShiftWorkers() {
    const listContainer = document.getElementById('work-log-worker-list');
    if (!listContainer) return;

    // dataset.team = "일근" 인 항목 체크 해제 및 상태 저장
    const rows = listContainer.querySelectorAll('.worker-row');
    rows.forEach(row => {
        if (row.dataset.team === '일근') {
            const chk = row.querySelector('.w-chk');
            if (chk && chk.checked) {
                row.dataset.wasChecked = "true"; // 상태 저장
                chk.checked = false;

                const drv = row.querySelector('.w-drive');
                if (drv) drv.value = '';
                updateWorkerStats();
            }
        }
    });
}

function checkDayShiftWorkers() {
    const listContainer = document.getElementById('work-log-worker-list');
    if (!listContainer) return;

    const rows = listContainer.querySelectorAll('.worker-row');
    rows.forEach(row => {
        if (row.dataset.team === '일근') {
            const chk = row.querySelector('.w-chk');
            // 이전에 체크되어 있었던 경우만 복구 (또는 기본적으로 모두 체크)
            if (chk && !chk.checked) {
                // 복구: 이전에 자동 해제되었거나(wasChecked), 명시적 해제 기록이 없는 경우
                if (row.dataset.wasChecked === "true" || !row.hasAttribute('data-was-checked')) {
                    chk.checked = true;
                }
            }
        }
    });

    updateWorkerStats();
}


// Work Log Exports
// --- ERP Worker Modal Logic ---

let erpModalWorkerList = []; // Cache list

function openERPWorkerModal() {
    if (!selectedERPLocation) {
        showNativeMsgBox("점검 장소를 선택해주세요.");
        return;
    }


    const modal = document.getElementById('erp-worker-modal');
    modal.style.display = 'flex';

    // Sync Toggle State from Main Tab
    const mainToggle = document.getElementById('toggle-worker-format');
    updateModalToggleState(mainToggle.checked);

    // Render List
    renderERPWorkerList();
}

function closeERPWorkerModal() {
    document.getElementById('erp-worker-modal').style.display = 'none';
}

function renderERPWorkerList() {
    const uid = selectedUserId;
    if (!uid || !appConfig.users || !appConfig.users[uid]) return;

    // Use global appSettings.colleagues
    const allColleagues = appConfig.appSettings?.colleagues || [];

    // Get My Team
    const user = appConfig.users[uid];
    const myTeam = (user.profile && user.profile.team) ? user.profile.team : '';

    const container = document.getElementById('erp-worker-list');
    container.innerHTML = '';

    // Filter: ID Match (if needed) or Team Match
    // Logic: Config colleagues contains EVERYONE. We need to filter by myTeam.
    const filteredWorkers = allColleagues.filter(w => {
        if (myTeam && w.team === myTeam) return true;
        return false;
    });

    if (filteredWorkers.length === 0) {
        container.innerHTML = '<div style="grid-column:1/-1; text-align:center; padding:20px; color:#999;">점검자 목록이 없습니다.<br>설정에서 작업자를 추가해주세요.</div>';
        return;
    }

    // Sort: Manager First, then ID
    const sortedWorkers = [...filteredWorkers].sort((a, b) => {
        if (a.isManager !== b.isManager) return b.isManager - a.isManager;
        return a.id.localeCompare(b.id);
    });

    erpModalWorkerList = sortedWorkers; // Cache for submission

    sortedWorkers.forEach(worker => {
        const div = document.createElement('div');
        div.className = 'worker-checkbox-item';
        // Remove custom onclick handler (causes double-toggle with label)

        const isChecked = true; // Default Select All

        div.innerHTML = `
            <label style="display:flex; align-items:center; gap:8px; cursor:pointer; width:100%; color:black;">
                <input type="checkbox" id="chk-erp-${worker.id}" ${isChecked ? 'checked' : ''} value="${worker.name}" onchange="updateModalButtonText()" style="width:16px; height:16px; accent-color:#4CAF50;">
                <span style="font-size:14px; margin-top:1px;">${worker.name}</span>
            </label>
        `;
        container.appendChild(div);
    });

    // Initial Button Text Update
    updateModalButtonText();
}

function toggleModalWorkerFormat() {
    // Toggle the Main Tab Switch (Source of Truth)
    const mainToggle = document.getElementById('toggle-worker-format');
    mainToggle.checked = !mainToggle.checked;

    // Trigger change event to save state if needed (savable-ui) and update styles
    mainToggle.dispatchEvent(new Event('change', { bubbles: true }));

    updateModalToggleState(mainToggle.checked);
    updateModalButtonText(); // Update text format
}

function updateModalToggleState(isListMode) {
    // No longer changing button text here based on mode alone, 
    // now we update based on selection + mode in updateModalButtonText
}

// Renamed/Refactored: Updates the button text dynamically
function updateModalButtonText() {
    const isListMode = document.getElementById('toggle-worker-format').checked;
    const btn = document.getElementById('btn-erp-modal-toggle');

    // Gather selected names
    const selectedNames = [];
    const checkboxes = document.querySelectorAll('#erp-worker-list input[type="checkbox"]');
    checkboxes.forEach(c => {
        if (c.checked) selectedNames.push(c.value);
    });

    if (selectedNames.length === 0) {
        btn.innerText = "(선택 없음)";
        return;
    }

    if (isListMode) {
        // List Mode: "Name1, Name2, Name3"
        btn.innerText = selectedNames.join(', ');
    } else {
        // Summary Mode: "Name1 외 N명"
        if (selectedNames.length === 1) {
            btn.innerText = selectedNames[0];
        } else {
            btn.innerText = `${selectedNames[0]} 외 ${selectedNames.length - 1}명`;
        }
    }
}

function submitERPTask() {
    // 1. Gather Selected Workers
    const selectedNames = [];
    const checkboxes = document.querySelectorAll('#erp-worker-list input[type="checkbox"]');
    checkboxes.forEach(c => {
        if (c.checked) selectedNames.push(c.value);
    });

    if (selectedNames.length === 0) {
        showNativeMsgBox("작업자를 한 명 이상 선택해주세요.");
        return;
    }

    // 2. Check Format
    const isListMode = document.getElementById('toggle-worker-format').checked;
    const format = isListMode ? 'list' : 'summary';

    // 3. Find Location Data (Type, Order)
    const locName = selectedERPLocation;
    let locType = "";
    let locOrder = "";

    // appConfig.appSettings.locations should be available
    if (appConfig.appSettings && appConfig.appSettings.locations) {
        const locObj = appConfig.appSettings.locations.find(l => l.name === locName);
        if (locObj) {
            locType = locObj.type || "";
            locOrder = locObj.order || "";
        }
    }

    // 4. Send to AHK
    sendMessageToAHK({
        command: 'runTask',
        task: 'ERPCheck',
        location: locName,
        targetType: locType,
        targetOrder: locOrder,
        members: selectedNames,
        format: format
    });

    closeERPWorkerModal();
}

function updateToggleStyle() {
    const toggle = document.getElementById('toggle-worker-format');
    const lblSummary = document.getElementById('lbl-format-summary');
    const lblList = document.getElementById('lbl-format-list');

    if (toggle.checked) {
        // List Mode Checked
        if (lblSummary) lblSummary.classList.remove('bold-active');
        if (lblList) lblList.classList.add('bold-active');
    } else {
        // Summary Mode Unchecked
        if (lblSummary) lblSummary.classList.add('bold-active');
        if (lblList) lblList.classList.remove('bold-active');
    }
    autoSaveSettings();
}

// Global Exports
window.openERPWorkerModal = openERPWorkerModal;
window.closeERPWorkerModal = closeERPWorkerModal;
window.toggleModalWorkerFormat = toggleModalWorkerFormat;
window.submitERPTask = submitERPTask;
window.handleWorkTypeChange = handleWorkTypeChange;
window.toggleAllWorkers = toggleAllWorkers;
window.updateToggleStyle = updateToggleStyle;
window.startWorkLog = startWorkLog;
window.toggleDrinkCalibration = toggleDrinkCalibration;

// --- Helper: Preset Management ---
function getUserPresets(type) {
    const uid = selectedUserId;
    if (!uid || !appConfig.users || !appConfig.users[uid]) return {};
    const user = appConfig.users[uid];
    if (type === 'track') return user.trackPresets || {};
    if (type === 'vehicle') return user.vehiclePresets || {};
    return {};
}

function saveUserPresets(type, presets) {
    const uid = selectedUserId;
    if (!uid) return;
    if (!appConfig.users[uid]) return;

    if (type === 'track') appConfig.users[uid].trackPresets = presets;
    if (type === 'vehicle') appConfig.users[uid].vehiclePresets = presets;

    sendMessageToAHK({ command: 'saveConfig', data: appConfig });
}

function renderPresetOptions(selectId, presets) {
    const sel = document.getElementById(selectId);
    if (!sel) return;

    // Keep selection if possible
    const currentVal = sel.value;

    // Clear existing options
    sel.innerHTML = '';

    const keys = Object.keys(presets);
    keys.forEach(key => {
        const opt = document.createElement('option');
        opt.value = key;
        opt.text = key;
        sel.appendChild(opt);
    });

    // "New Preset" Option
    const newOpt = document.createElement('option');
    newOpt.value = "__NEW__";
    newOpt.text = "+ 새 프리셋 추가";
    newOpt.style.fontWeight = "bold";
    newOpt.style.color = "#0063B5";
    sel.appendChild(newOpt);

    // Determines Selection
    if (keys.length > 0) {
        // 프리셋이 있으면, 현재 선택된 값이 유효한 프리셋인 경우만 유지하고
        // 그 외(New Preset 포함)에는 첫 번째 프리셋을 강제로 선택합니다.
        if (currentVal && presets[currentVal]) {
            sel.value = currentVal;
        } else {
            sel.value = keys[0];
            sel.dispatchEvent(new Event('change'));
        }
    } else {
        // 프리셋이 없으면 새 프리셋 추가 선택
        sel.value = "__NEW__";
        sel.dispatchEvent(new Event('change'));
    }
}

// --- Track Access Logic ---
function loadTrackPreset() {
    const sel = document.getElementById('track-preset-sel');
    const key = sel.value;

    if (key === '__NEW__') {
        // Clear Form for New Entry
        setVal('ta-work-type', '1');
        setVal('ta-work-content', '');
        setVal('ta-work-from', '');
        setVal('ta-work-to', '');
        setVal('ta-driver-name', '');
        setVal('ta-driver-phone', '');
        setVal('ta-worker-name', '');
        setVal('ta-worker-phone', '');
        setVal('ta-safety-name', '');
        setVal('ta-safety-phone', '');
        setVal('ta-supervisor-name', '');
        setVal('ta-supervisor-id', '');

        setVal('ta-work-start', '');
        setVal('ta-work-end', '');
        setVal('ta-op-start', '');
        setVal('ta-op-end', '');
        setVal('ta-line', '1');
        setVal('ta-track-type', '1');
        if (document.getElementById('ta-track-cutoff')) document.getElementById('ta-track-cutoff').checked = false;
        setVal('ta-agreement-no', '');
        setVal('ta-total-count', '');
        if (document.getElementById('ta-station-input')) document.getElementById('ta-station-input').checked = false;
        return;
    }

    if (!key) return;

    const presets = getUserPresets('track');
    const data = presets[key];

    if (data) {
        setVal('ta-work-type', data.workType);
        setVal('ta-work-content', data.content);
        setVal('ta-work-from', data.workFrom);
        setVal('ta-work-to', data.workTo);
        setVal('ta-driver-name', data.driverName);
        setVal('ta-driver-phone', data.driverPhone);
        setVal('ta-worker-name', data.workerName);
        setVal('ta-worker-phone', data.workerPhone);
        setVal('ta-safety-name', data.safetyName);
        setVal('ta-safety-phone', data.safetyPhone);
        setVal('ta-supervisor-name', data.supervisorName);
        setVal('ta-supervisor-id', data.supervisorId);

        setVal('ta-work-start', data.workStart);
        setVal('ta-work-end', data.workEnd);
        setVal('ta-op-start', data.opStart);
        setVal('ta-op-end', data.opEnd);
        setVal('ta-line', data.line);
        setVal('ta-track-type', data.trackType);

        if (document.getElementById('ta-track-cutoff')) document.getElementById('ta-track-cutoff').checked = !!data.trackCutoff;
        setVal('ta-agreement-no', data.agreementNo);
        setVal('ta-total-count', data.totalCount);
        if (document.getElementById('ta-station-input')) document.getElementById('ta-station-input').checked = !!data.stationInput;
    }
}

function saveTrackPreset() {
    const sel = document.getElementById('track-preset-sel');
    let key = sel.value;

    if (!key || key === '__NEW__') {
        const newName = prompt("새 프리셋 이름을 입력하세요:");
        if (!newName) return;
        key = newName;
    }

    const data = {
        workType: getVal('ta-work-type'),
        content: getVal('ta-work-content'),
        workFrom: getVal('ta-work-from'),
        workTo: getVal('ta-work-to'),
        driverName: getVal('ta-driver-name'),
        driverPhone: getVal('ta-driver-phone'),
        workerName: getVal('ta-worker-name'),
        workerPhone: getVal('ta-worker-phone'),
        safetyName: getVal('ta-safety-name'),
        safetyPhone: getVal('ta-safety-phone'),
        supervisorName: getVal('ta-supervisor-name'),
        supervisorId: getVal('ta-supervisor-id'),

        workStart: getVal('ta-work-start'),
        workEnd: getVal('ta-work-end'),
        opStart: getVal('ta-op-start'),
        opEnd: getVal('ta-op-end'),
        line: getVal('ta-line'),
        trackType: getVal('ta-track-type'),
        trackCutoff: document.getElementById('ta-track-cutoff') ? document.getElementById('ta-track-cutoff').checked : false,
        agreementNo: getVal('ta-agreement-no'),
        totalCount: getVal('ta-total-count'),
        stationInput: document.getElementById('ta-station-input') ? document.getElementById('ta-station-input').checked : false
    };

    const presets = getUserPresets('track');
    presets[key] = data;
    saveUserPresets('track', presets);

    renderPresetOptions('track-preset-sel', presets);
    sel.value = key;

    showNativeMsgBox(`'${key}' 프리셋이 저장되었습니다.`);
}

function renameTrackPreset() {
    const sel = document.getElementById('track-preset-sel');
    const oldKey = sel.value;
    if (!oldKey || oldKey === '__NEW__') {
        showNativeMsgBox("이름을 변경할 프리셋을 선택해주세요.");
        return;
    }

    const newKey = prompt("새 이름을 입력하세요:", oldKey);
    if (!newKey || newKey === oldKey) return;

    const presets = getUserPresets('track');
    if (presets[newKey]) {
        showNativeMsgBox("이미 존재하는 이름입니다.");
        return;
    }

    presets[newKey] = presets[oldKey];
    delete presets[oldKey];
    saveUserPresets('track', presets);

    renderPresetOptions('track-preset-sel', presets);
    sel.value = newKey;
}

function deleteTrackPreset() {
    const sel = document.getElementById('track-preset-sel');
    const key = sel.value;
    if (!key || key === '__NEW__') {
        showNativeMsgBox("삭제할 프리셋을 선택해주세요.");
        return;
    }

    if (!confirm(`'${key}' 프리셋을 삭제하시겠습니까?`)) return;

    const presets = getUserPresets('track');
    delete presets[key];
    saveUserPresets('track', presets);

    renderPresetOptions('track-preset-sel', presets);

    // Auto-select logic is handled inside renderPresetOptions if we pass nothing, 
    // BUT renderPresetOptions expects to respect currentVal if passed.
    // Since we deleted the key, we should let it default.
    // However, our renderPresetOptions helper tries to keep selection.
    // Let's manually trigger the logic again.

    // Quick Fix: renderPresetOptions handles init logic if we don't set value explicitly?
    // Actually, renderPresetOptions uses 'sel.value' to determine previous value.
    // We should clear it before calling? No, it reads it.

    // Correct approach using our new robust renderPresetOptions:
    // 1. Value is still the deleted key technically before we re-render? No, we re-render options.
    // Actually, let's just trigger the 'change' event on the first item if exists.

    const newKeys = Object.keys(presets);
    if (newKeys.length > 0) {
        sel.value = newKeys[0];
    } else {
        sel.value = "__NEW__";
    }
    sel.dispatchEvent(new Event('change'));
}

function runTrackAccessTask() {
    const data = {
        workType: getVal('ta-work-type'),
        content: getVal('ta-work-content'),
        workFrom: getVal('ta-work-from'),
        workTo: getVal('ta-work-to'),
        driverName: getVal('ta-driver-name'),
        driverPhone: getVal('ta-driver-phone'),
        workerName: getVal('ta-worker-name'),
        workerPhone: getVal('ta-worker-phone'),
        safetyName: getVal('ta-safety-name'),
        safetyPhone: getVal('ta-safety-phone'),
        supervisorName: getVal('ta-supervisor-name'),
        supervisorId: getVal('ta-supervisor-id'),

        workStart: getVal('ta-work-start'),
        workEnd: getVal('ta-work-end'),
        opStart: getVal('ta-op-start'),
        opEnd: getVal('ta-op-end'),
        line: getVal('ta-line'),
        trackType: getVal('ta-track-type'),
        trackCutoff: document.getElementById('ta-track-cutoff') ? document.getElementById('ta-track-cutoff').checked : false,
        agreementNo: getVal('ta-agreement-no'),
        totalCount: getVal('ta-total-count'),
        stationInput: document.getElementById('ta-station-input') ? document.getElementById('ta-station-input').checked : false
    };

    sendMessageToAHK({ command: 'runTask', task: 'TrackAccess', data: data });
}


// --- Vehicle Log Logic ---
function loadVehiclePreset() {
    const sel = document.getElementById('vehicle-preset-sel');
    const key = sel.value;

    if (key === '__NEW__') {
        setVal('vl-driver', '');
        setVal('vl-point-1', '');
        setVal('vl-point-2', '');
        setVal('vl-track-type', '상행선');
        setVal('vl-start-time', '');
        setVal('vl-end-time', '');
        setVal('vl-content', '');
        setVal('vl-remarks', '');
        setVal('vl-approve-no', '');
        setVal('vl-dept', '');
        setVal('vl-approver', '');
        setVal('vl-run-time', '');
        setVal('vl-distance', '');
        return;
    }

    if (!key) return;

    const presets = getUserPresets('vehicle');
    const data = presets[key];

    if (data) {
        setVal('vl-driver', data.driver);
        setVal('vl-point-1', data.point1);
        setVal('vl-point-2', data.point2);
        setVal('vl-track-type', data.trackType);
        setVal('vl-start-time', data.startTime);
        setVal('vl-end-time', data.endTime);
        setVal('vl-content', data.content);
        setVal('vl-remarks', data.remarks);
        setVal('vl-approve-no', data.approveNo);
        setVal('vl-dept', data.dept);
        setVal('vl-approver', data.approver);
        setVal('vl-run-time', data.runTime);
        setVal('vl-distance', data.distance);
    }
}

function saveVehiclePreset() {
    const sel = document.getElementById('vehicle-preset-sel');
    let key = sel.value;

    if (!key || key === '__NEW__') {
        const newName = prompt("새 프리셋 이름을 입력하세요:");
        if (!newName) return;
        key = newName;
    }

    const data = {
        driver: getVal('vl-driver'),
        point1: getVal('vl-point-1'),
        point2: getVal('vl-point-2'),
        trackType: getVal('vl-track-type'),
        startTime: getVal('vl-start-time'),
        endTime: getVal('vl-end-time'),
        content: getVal('vl-content'),
        remarks: getVal('vl-remarks'),
        approveNo: getVal('vl-approve-no'),
        dept: getVal('vl-dept'),
        approver: getVal('vl-approver'),
        runTime: getVal('vl-run-time'),
        distance: getVal('vl-distance')
    };

    const presets = getUserPresets('vehicle');
    presets[key] = data;
    saveUserPresets('vehicle', presets);

    renderPresetOptions('vehicle-preset-sel', presets);
    sel.value = key;

    showNativeMsgBox(`'${key}' 프리셋이 저장되었습니다.`);
}

function renameVehiclePreset() {
    const sel = document.getElementById('vehicle-preset-sel');
    const oldKey = sel.value;
    if (!oldKey || oldKey === '__NEW__') {
        showNativeMsgBox("이름을 변경할 프리셋을 선택해주세요.");
        return;
    }

    const newKey = prompt("새 이름을 입력하세요:", oldKey);
    if (!newKey || newKey === oldKey) return;

    const presets = getUserPresets('vehicle');
    if (presets[newKey]) {
        showNativeMsgBox("이미 존재하는 이름입니다.");
        return;
    }

    presets[newKey] = presets[oldKey];
    delete presets[oldKey];
    saveUserPresets('vehicle', presets);

    renderPresetOptions('vehicle-preset-sel', presets);
    sel.value = newKey;
}

function deleteVehiclePreset() {
    const sel = document.getElementById('vehicle-preset-sel');
    const key = sel.value;
    if (!key || key === '__NEW__') {
        showNativeMsgBox("삭제할 프리셋을 선택해주세요.");
        return;
    }

    if (!confirm(`'${key}' 프리셋을 삭제하시겠습니까?`)) return;

    const presets = getUserPresets('vehicle');
    delete presets[key];
    saveUserPresets('vehicle', presets);

    renderPresetOptions('vehicle-preset-sel', presets);

    // Auto-select logic
    const newKeys = Object.keys(presets);
    if (newKeys.length > 0) {
        sel.value = newKeys[0];
    } else {
        sel.value = "__NEW__";
    }
    sel.dispatchEvent(new Event('change'));
}

function runVehicleLogTask() {
    const data = {
        driver: getVal('vl-driver'),
        department: appConfig.users[selectedUserId] ? appConfig.users[selectedUserId].profile.department : '',
        point1: getVal('vl-point-1'),
        point2: getVal('vl-point-2'),
        trackType: getVal('vl-track-type'),
        startTime: getVal('vl-start-time'),
        endTime: getVal('vl-end-time'),
        content: getVal('vl-content'),
        remarks: getVal('vl-remarks'),
        approveNo: getVal('vl-approve-no'),
        dept: getVal('vl-dept'),
        approver: getVal('vl-approver'),
        runTime: getVal('vl-run-time'),
        distance: getVal('vl-distance')
    };

    sendMessageToAHK({ command: 'runTask', task: 'VehicleLog', data: data });
}

// Execute 'runBringApproved' command
function runBringApproved() {
    const driver = getVal('vl-driver');
    if (!driver) {
        showNativeMsgBox("운전자를 입력해주세요.");
        return;
    }
    sendMessageToAHK({
        command: 'runTask',
        task: 'bringApproval',
        data: { driverName: driver }
    });
}

// Helpers
function getVal(id) {
    const el = document.getElementById(id);
    return el ? el.value : '';
}

function setVal(id, val) {
    const el = document.getElementById(id);
    if (el) el.value = val || '';
}

// Helper: Init Presets after login
function initPresets() {
    const trackPresets = getUserPresets('track');
    const vehiclePresets = getUserPresets('vehicle');
    renderPresetOptions('track-preset-sel', trackPresets);
    renderPresetOptions('vehicle-preset-sel', vehiclePresets);
}


// Export new functions
window.loadTrackPreset = loadTrackPreset;
window.saveTrackPreset = saveTrackPreset;
window.renameTrackPreset = renameTrackPreset;
window.deleteTrackPreset = deleteTrackPreset;
window.runTrackAccessTask = runTrackAccessTask;

window.loadVehiclePreset = loadVehiclePreset;
window.saveVehiclePreset = saveVehiclePreset;
window.renameVehiclePreset = renameVehiclePreset;
window.deleteVehiclePreset = deleteVehiclePreset;
window.runVehicleLogTask = runVehicleLogTask;
window.initPresets = initPresets; // To be called after login/load

// --- Daily Log Preset & Track Access Helpers ---

function loadPresetDetail() {
    const sel = document.getElementById('track-preset-sel'); // Fixed ID
    const key = sel.value;

    // This function is bound to the Track Access Preset selector in the Preset Tab.
    // It should load the Track Access preset details into the "ps-name" and "ps-work" fields
    // which seem to be intended for Quick Viewing/Editing of Track Access presets?
    // Or maybe the user INTENDED this selector to load the Daily Log defaults?
    // Based on "일지프리셋(Preset) 탭", and the fact that we have a separate "Daily Log Defaults" section...
    // I will assume this selector controls the "Track Access Presets" section at the top.

    if (!key) return;

    // If it's the "New Preset" option
    // (We don't have logic for it in the settings tab selector usually, but let's handle it)

    const presets = getUserPresets('track');
    const data = presets[key]; // Might be undefined if new

    if (data) {
        setVal('ps-name', key); // Name is the key
        setVal('ps-work', data.content); // "Work Content" from Track Access
    } else {
        setVal('ps-name', '');
        setVal('ps-work', '');
    }
}

function addGeneralWorkRow(data = {}) {
    const tbody = document.querySelector('#general-work-table tbody');
    const rowData = Object.keys(data).length > 0 ? data : {
        workType: '주간',
        category: '전체',
        content: '',
        manager: '',
        start: '',
        end: ''
    };

    const tr = document.createElement('tr');

    // Options
    const workTypes = ['주간', '야간'];
    const categories = ['전체', '내부업무', '점검업무', '유지보수', '협조사항'];

    const workOpts = workTypes.map(t => `<option value='${t}' ${rowData.workType === t ? 'selected' : ''}>${t}</option>`).join('');
    const catOpts = categories.map(c => `<option value='${c}' ${rowData.category === c ? 'selected' : ''}>${c}</option>`).join('');

    tr.innerHTML = `
        <td><select>${workOpts}</select></td>
        <td><select>${catOpts}</select></td>
        <td><input type='text' value='${rowData.content || ''}' placeholder='내용'></td>
        <td><input type='text' value='${rowData.manager || ''}' placeholder='책임자'></td>
        <td><input type='text' value='${rowData.start || ''}' placeholder='   :   ' maxlength='5' oninput='formatTime(this)' class='time-input'></td>
        <td><input type='text' value='${rowData.end || ''}' placeholder='   :   ' maxlength='5' oninput='formatTime(this)' class='time-input'></td>
        <td class='center'><button class='small-btn danger' onclick='this.closest("tr").remove()'>X</button></td>
    `;
    tbody.appendChild(tr);
}

function saveAutoInputSettings() {
    const time = document.getElementById('fps-auto-input-time').value;
    // Just save settings overall
    saveSettings();
    showNativeMsgBox(`자동입력 예약시간(${time})이 저장되었습니다.`);
}

// Global Exports
window.loadPresetDetail = loadPresetDetail;
window.addGeneralWorkRow = addGeneralWorkRow;
window.saveAutoInputSettings = saveAutoInputSettings;
window.formatTime = formatTime;

// [Refactor] 통합 UI 갱신 함수 (Hot Reload 지원)
// 설정 변경 후 호출되어 각 화면의 요소를 강제로 최신화합니다.
// 모든 UI 갱신 로직을 이 함수 하나로 통합하여 스파게티 코드 문제를 해결합니다.
function refreshUI() {
    if (!selectedUserId) return;

    // 1. 업무일지: 작업자 명단 갱신 (팀 변경, 이름 변경 등 반영)
    // renderWorkLogWorkerList는 appConfig.users[selectedUserId]와 appSettings.colleagues를 새로 읽어옵니다.
    renderWorkLogWorkerList();

    // 2. 업무일지: 프리셋 및 옵션 재적용
    // handleWorkTypeChange(true)는 현재 선택된 주/야간 모드에 맞춰 체크박스, 안전관리 내용을 다시 세팅합니다.
    handleWorkTypeChange(true);

    // 3. ERP 점검: 장소 목록 및 오더번호 갱신
    // 탭이 그 때 활성화되어 있지 않더라도 DOM을 갱신해두면 나중에 탭 진입 시 최신 상태가 보입니다.
    renderERPCheck();
    updateToggleStyle();

    // 4. 프리셋 목록 갱신 (선로출입/차량일지)
    // 설정에서 프리셋이 변경되었을 때 반영하기 위해 갱신합니다.
    initPresets();
}

// [New] Headless 작업자 명단 불러오기 요청
function importWorkersFromHeadless() {
    const uid = selectedUserId;
    if (!uid) return;

    // 현재 유저의 '분소' 코드 (arbpl) 가져오기
    // appConfig.users[uid].profile.arbpl 에 저장되어 있다고 가정 (Settings에서 저장됨)
    // 없을 경우 기본값 혹은 오류 처리
    const user = appConfig.users[uid];
    const arbpl = (user.profile && user.profile.arbpl) ? user.profile.arbpl : "";

    if (!arbpl) {
        showNativeMsgBox("분소 정보가 설정되지 않았습니다. [내 정보] 탭에서 분소를 선택해주세요.");
        return;
    }

    const btn = document.getElementById('btn-import-workers');
    if (btn) btn.disabled = true; // 중복 클릭 방지

    // 요청 전송
    sendMessageToAHK({ command: 'importWorkers', arbpl: arbpl });

    // 타임아웃 처리 (혹은 응답 대기 로직)
    setTimeout(() => {
        if (btn && btn.disabled) {
            btn.disabled = false; // 10초 후 복구 (실패 시 등)
        }
    }, 10000);
}

// 수신된 작업자 명단 처리
function handleWorkerListUpdate(newWorkers) {
    if (!newWorkers || !Array.isArray(newWorkers)) return;

    // 기존 테이블 데이터 스캔 (중복 방지)
    const existingIds = new Set();
    document.querySelectorAll('#worker-table tbody tr').forEach(row => {
        const idInput = row.querySelector('input[placeholder="사번"]');
        if (idInput && idInput.value) {
            existingIds.add(idInput.value);
        }
    });

    let addedCount = 0;
    const tbody = document.querySelector('#worker-table tbody');

    newWorkers.forEach(w => {
        // AHK Map keys: "사번", "이름", "휴가종류", "근무조"
        const id = w["사번"];
        const name = w["이름"];
        const team = w["근무조"];
        // const vacation = w["휴가종류"]; // 비고란엔 넣지 않음 (요청사항 없음, 필요 시 추가)

        if (id && !existingIds.has(id)) {
            // 새 작업자 추가
            addWorkerRowToTable(tbody, {
                name: name,
                id: id,
                team: team,
                phone: '',
                isManager: 0,
                driverRole: '-'
            });
            addedCount++;
        }
    });

    if (addedCount > 0) {
        autoSaveSettings(); // 데이터 변경 저장

        // Auto scroll
        const contentArea = document.querySelector('.settings-content-area');
        if (contentArea) contentArea.scrollTop = contentArea.scrollHeight;
    } else {
        showNativeMsgBox("추가할 새로운 분소원이 없습니다.");
    }

    const btn = document.getElementById('btn-import-workers');
    if (btn) btn.disabled = false;
}

window.importWorkersFromHeadless = importWorkersFromHeadless;