// app.js

// --- Global Data Store ---
let appConfig = {};
let selectedUserId = null;
let loginUsers = []; // Store user data for login validation
let isWorkLogInitialized = false; // Flag for persistence
let pendingRestoreState = null; // State waiting for config load
let currentSafetyEduData = {}; // [New] Store for Safety Edu Presets
let guntaeData = null; // 근태 조회 결과 캐시
let guntaeStatus = 'idle'; // idle, loading, success, failure
let configReadyUserId = null;
let pendingAutoWorkerList = null;

const shellState = Vue.reactive({
    title: '통합자동화',
    activeMainTab: 'view-work-log',
    shift: {
        visible: false, prev: '', current: '', next: '',
        prevDate: '', currentDate: '', nextDate: '', name: '', isNight: false
    },
    session: { visible: false, state: 'ready', text: '준비완료' },
    initOverlay: { visible: false, fading: false, message: '초기 구성 중...' },
    missingPassword: { visible: false, types: [], hideNextTime: false }
});

const authState = Vue.reactive({
    view: 'users', users: [], selectedId: null, loginPw: '',
    form: {
        name: '', id: '', department: '', team: 'A조',
        webPW: '', webPWConfirm: '', pw2: '', pw2Confirm: '', sapPW: '', sapPWConfirm: ''
    }
});

const erpControlsState = Vue.reactive({
    isBatchMode: false,
    isListFormat: false,
    lastUpdate: '',
    sessionReady: false
});

// --- Firestore Realtime FormList ---
const ERP_FIREBASE_CONFIG = Object.freeze({
    apiKey: 'AIzaSyD4eSO-idxDepO8knAqLLzxX5ZfNCy9NAM',
    authDomain: 'btcwebapp-551bd.firebaseapp.com',
    projectId: 'btcwebapp-551bd'
});
let erpFormListUnsubscribe = null;
const USER_SETTINGS_COLLECTION = 'userSettings';

function getFirebaseApp() {
    if (!window.firebase || typeof window.firebase.initializeApp !== 'function') {
        throw new Error('Firebase SDK를 로드하지 못했습니다.');
    }

    const appName = 'erp-form-list-realtime';
    return window.firebase.apps.find(app => app.name === appName)
        || window.firebase.initializeApp(ERP_FIREBASE_CONFIG, appName);
}

function firestoreRevisionToIso(value) {
    if (!value) return '';
    if (typeof value === 'string') return value;
    if (typeof value.toDate === 'function') return value.toDate().toISOString();
    if (Number.isFinite(value.seconds)) return new Date(value.seconds * 1000).toISOString();
    return String(value);
}

function seoulDateKey(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return String(value || '').slice(0, 10);

    const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: 'Asia/Seoul',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
    }).formatToParts(date);
    const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
    return `${values.year}-${values.month}-${values.day}`;
}

function normalizeERPFormList(items) {
    if (!Array.isArray(items)) return [];
    return items
        .map(item => {
            const lastModifiedDate = firestoreRevisionToIso(item?.lastModifiedDate);
            return {
                formKey: String(item?.formKey || ''),
                sheetName: String(item?.sheetName || '').trim(),
                lastModifiedDate,
                seoulDate: seoulDateKey(lastModifiedDate)
            };
        })
        .filter(item => item.sheetName && item.lastModifiedDate);
}

function startERPFormListRealtime() {
    if (erpFormListUnsubscribe) return;

    if (!window.firebase || typeof window.firebase.initializeApp !== 'function') {
        sendMessageToAHK({
            command: 'erpFormListError',
            message: 'Firebase SDK를 로드하지 못했습니다.'
        });
        return;
    }

    try {
        const firebaseApp = getFirebaseApp();
        const formListRef = firebaseApp.firestore().collection('publicCache').doc('formList');

        erpFormListUnsubscribe = formListRef.onSnapshot(snapshot => {
            const items = snapshot.exists
                ? normalizeERPFormList(snapshot.data()?.items)
                : [];
            sendMessageToAHK({
                command: 'updateERPFormList',
                source: 'firestore',
                items
            });
        }, error => {
            console.error('ERP FormList Firestore listener failed:', error);
            sendMessageToAHK({
                command: 'erpFormListError',
                message: error?.message || String(error)
            });
        });

        window.addEventListener('beforeunload', () => {
            if (erpFormListUnsubscribe) erpFormListUnsubscribe();
            erpFormListUnsubscribe = null;
        }, { once: true });
    } catch (error) {
        console.error('ERP FormList Firestore initialization failed:', error);
        sendMessageToAHK({
            command: 'erpFormListError',
            message: error?.message || String(error)
        });
    }
}

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    initTitleBarApp();
    initAuthApp();
    initSidebarApp();
    initERPControlsApps();
    initCommonOverlayApps();
    initSettingsApp();
    initWorkLogApp();
    initERPIndividualApp();
    initERPWorkerModalApp();
    initTrackAccessApp();
    initVehicleLogApp();
    initMaterialLocationApp();

    // Initial Nav Setup
    switchMainTab('view-work-log');

    // Request Initial Data
    if (window.chrome && window.chrome.webview) {
        window.chrome.webview.postMessage({ command: 'ready' });
    }
    startERPFormListRealtime();

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

function initTitleBarApp() {
    Vue.createApp({
        data() { return { shell: shellState }; },
        methods: {
            drag(event) {
                if (event.button !== 0 || event.target.closest('.control-btn')) return;
                sendMessageToAHK({ command: 'dragWindow' });
            },
            minimize() { sendMessageToAHK({ command: 'minimize' }); },
            close() { sendMessageToAHK({ command: 'close' }); }
        }
    }).mount('#title-bar');
}

function initAuthApp() {
    Vue.createApp({
        data() {
            return { auth: authState, departments: SETTINGS_DEPARTMENTS, teams: SETTINGS_TEAMS };
        },
        computed: {
            selectedUser() {
                return this.auth.users.find(user => String(user.id) === String(this.auth.selectedId)) || null;
            },
            webPasswordsMatch() {
                return !!this.auth.form.webPWConfirm && this.auth.form.webPW === this.auth.form.webPWConfirm;
            },
            secondPasswordsMatch() {
                return !!this.auth.form.pw2Confirm && this.auth.form.pw2 === this.auth.form.pw2Confirm;
            },
            sapPasswordsMatch() {
                return !!this.auth.form.sapPWConfirm && this.auth.form.sapPW === this.auth.form.sapPWConfirm;
            }
        },
        mounted() {
            window.addEventListener('keydown', this.navigateUsers);
        },
        methods: {
            select: selectUser,
            selectAndContinue(id) { selectUser(id); moveToPasswordView(); },
            next: moveToPasswordView,
            back: backToUserList,
            login: tryLogin,
            showAdd() { switchView('add'); },
            cancelAdd() { switchView('login'); },
            remove: deleteUser,
            submit: submitNewUser,
            loginPwInput(event) {
                this.auth.loginPw = String(event.target.value || '').replace(/\D/g, '').slice(0, 6);
                checkAutoLogin(this.auth.loginPw);
            },
            numericFormInput(field, event) {
                this.auth.form[field] = String(event.target.value || '').replace(/\D/g, '').slice(0, 6);
            },
            navigateUsers(event) {
                if (this.auth.view !== 'users' || document.getElementById('login-container').style.display === 'none') return;
                if (event.key === 'Enter') {
                    if (this.auth.selectedId) {
                        event.preventDefault();
                        moveToPasswordView();
                    }
                    return;
                }
                if (!['ArrowDown', 'ArrowUp'].includes(event.key) || this.auth.users.length === 0) return;
                event.preventDefault();
                const current = this.auth.users.findIndex(user => String(user.id) === String(this.auth.selectedId));
                const base = current < 0 ? (event.key === 'ArrowDown' ? -1 : this.auth.users.length) : current;
                const next = event.key === 'ArrowDown'
                    ? Math.min(base + 1, this.auth.users.length - 1)
                    : Math.max(base - 1, 0);
                selectUser(this.auth.users[next].id);
                Vue.nextTick(() => document.querySelector('.user-item.selected')?.scrollIntoView({ block: 'nearest' }));
            }
        }
    }).mount('#login-container');
}

function initSidebarApp() {
    Vue.createApp({
        data() { return { shell: shellState }; },
        methods: {
            openTab: switchMainTab,
            settings: openSettings
        }
    }).mount('#sidebar');
}

function initERPControlsApps() {
    Vue.createApp({
        data() { return { controls: erpControlsState }; },
        methods: {
            refresh: requestERPRefresh
        }
    }).mount('#erp-header-controls');

    Vue.createApp({
        data() { return { controls: erpControlsState }; },
        methods: {
            toggleMode: toggleERPMode,
            formatChanged: saveERPFormatPreference,
            start: openERPHandler
        }
    }).mount('#erp-footer-controls');
}

function initCommonOverlayApps() {
    Vue.createApp({
        data() { return { warning: shellState.missingPassword }; },
        methods: { close: closeMissingPwModal, openSettings: goSetMissingPw }
    }).mount('#missing-pw-modal');

    Vue.createApp({ data() { return { session: shellState.session }; } })
        .mount('#session-status-overlay');
    Vue.createApp({ data() { return { overlay: shellState.initOverlay }; } })
        .mount('#init-overlay');
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
        case 'materialLocationResult':
            handleMaterialLocationResult(msg);
            break;
        case 'initLogin':
            loginUsers = msg.users;
            renderUserList(msg.users);
            if (msg.preSelectUid) {
                selectUser(msg.preSelectUid);
            }
            break;
        case 'loginSuccess':
            handleLoginSuccess(msg.profile);
            break;
        case 'loginFail':
            showNativeMsgBox(msg.message, "로그인 실패");
            break;
        case 'confirmMsgboxResult':
            if (pendingNativeConfirmations.has(msg.requestId)) {
                const resolve = pendingNativeConfirmations.get(msg.requestId);
                pendingNativeConfirmations.delete(msg.requestId);
                resolve(!!msg.confirmed);
            }
            break;
        case 'updateShiftStatus':
            updateShiftUI(msg.data);
            break;
        case 'loadConfig':
            appConfig = msg.data;
            // trackAccessApp은 설정 수신 전에 마운트되므로 작업자 목록 computed가
            // 최초 빈 값을 계속 캐시하지 않도록 명시적으로 갱신 신호를 보낸다.
            trackAccessState.workerListRevision += 1;
            configReadyUserId = selectedUserId;
            ensureUserHotkeyDefaults();
            syncMaterialLocationDefaults();
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
            processPendingAutoWorkerList();
            processPendingLocationOrderCheck();
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
                vehicleLogState.form.approveNo = String(msg.data.승인번호 || '');
                vehicleLogState.form.dept = String(msg.data.승인부서 || '');
                vehicleLogState.form.approver = String(msg.data.승인자 || '');
            } else {
                showNativeMsgBox(msg.error || "승인정보를 불러오지 못했습니다.");
            }
            break;
        case 'updateERPOrderList': // New Case
            handleERPOrderListUpdate(msg.orders);
            break;
        case 'updateTitle':
            shellState.title = String(msg.title || '통합자동화');
            break;
        case 'updateGuntae':
            // 근태 조회 결과 수신
            guntaeData = msg.data;
            applyGuntaeData(msg.data);
            break;
        case 'updateGuntaeStatus':
            updateGuntaeStatusLabel(msg.status);
            break;
        case 'updateWorkerList':
            if (msg.automatic) {
                pendingAutoWorkerList = msg;
                processPendingAutoWorkerList();
            } else {
                handleWorkerListUpdate(msg.data);
            }
            break;
        case 'updateLocationList':
            handleLocationListUpdate(msg.data, msg.employeeId, msg.requestId);
            break;
        case 'locationImportFailed':
            if (pendingLocationImport?.requestId === msg.requestId) {
                pendingLocationImport = null;
                settingsState.importingLocations = false;
            }
            break;
        case 'locationOrderMaintenanceReady':
            pendingLocationOrderReady = msg;
            processPendingLocationOrderCheck();
            break;
        case 'locationOrderMaintenanceResult':
            handleLocationOrderMaintenance(msg);
            break;
        case 'updateAgreementNo':
            applyTrackAgreementNo(msg.value);
            break;
        case 'trackAgreementResult':
            trackAccessState.agreementLoading = false;
            if (msg.value) applyTrackAgreementNo(msg.value);
            if (msg.error) showNativeMsgBox(msg.error, '철도운행협의서 불러오기');
            break;
        case 'trackAgreementRenewalConfirmResult':
            trackAccessState.renewalConfirmPending = false;
            if (msg.confirmed) requestTrackAgreement();
            break;
        case 'showInitOverlay':
            // [New] 초기화 오버레이 표시
            showInitOverlay(msg.message);
            break;
        case 'hideInitOverlay':
            // [New] 초기화 오버레이 숨김
            hideInitOverlay();
            break;
        case 'sessionStatus':
            handleSessionStatus(msg.state, msg.message);
            break;
        case 'requireAuth':
            if (selectedUserId) {
                switchView('login');
                moveToPasswordView();
            }
            break;
    }
}

// --- Session Status Overlay ---
function handleSessionStatus(state, message) {
    shellState.session.state = state;

    if (state === 'acquiring') {
        setSessionControlsEnabled(false);
        shellState.session.text = '통합 세션 로그인 중...';
        shellState.session.visible = true;
    } else if (state === 'ready') {
        setSessionControlsEnabled(true);
        shellState.session.text = '준비완료';
        shellState.session.visible = true;
        setTimeout(() => {
            if (shellState.session.state === 'ready') shellState.session.visible = false;
        }, 1500);
    } else if (state === 'error') {
        setSessionControlsEnabled(false);
        shellState.session.text = message || '세션확보 실패';
        shellState.session.visible = true;
        setTimeout(() => {
            if (shellState.session.state === 'error') shellState.session.visible = false;
        }, 3000);
    }
}

function setSessionControlsEnabled(enabled) {
    settingsState.sessionControlsEnabled = !!enabled;
    erpControlsState.sessionReady = !!enabled;
    vehicleLogState.sessionReady = !!enabled;
}

// --- Initialization Overlay ---
function showInitOverlay(message) {
    shellState.initOverlay.message = message || '초기 구성 중...';
    shellState.initOverlay.fading = false;
    shellState.initOverlay.visible = true;
}

function hideInitOverlay() {
    shellState.initOverlay.fading = true;
    setTimeout(() => {
        shellState.initOverlay.visible = false;
        shellState.initOverlay.fading = false;
    }, 500);
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
    else if (authState.view === 'add') state.activeView = 'add';
    else state.activeView = 'login'; // default

    // 2. Active Main Tab
    state.mainTab = shellState.activeMainTab;

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

    // Vue가 관리하는 화면은 DOM이 아니라 반응형 원본 상태를 직접 저장합니다.
    state.customState.erpCheck = collectERPRecoveryState();
    state.customState.trackAccess = collectTrackAccessRecoveryState();
    state.customState.workLog = collectWorkLogRecoveryState();
    state.customState.vehicleLog = collectVehicleLogRecoveryState();

    return state;
}

function getSavableValue(el, key) {
    if (key === 'erpLocation') {
        return erpState.selectedLocation;
    }
    // Future extensions:
    // if (key === 'someOtherWidget') return ...;
    return null;
}

function restoreUiStateData(state) {
    if (!state) return;

    const hasTrackAccessSnapshot = !!state.customState?.trackAccess?.form;
    const hasWorkLogSnapshot = !!state.customState?.workLog;
    const hasVehicleLogSnapshot = !!state.customState?.vehicleLog?.form;

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
            // 새 복구 형식에서는 선로출입 Vue 상태를 아래에서 한 번에 직접 적용합니다.
            if (hasTrackAccessSnapshot && el.closest('#view-track-access')) return;
            if (hasWorkLogSnapshot && el.closest('#view-work-log')) return;
            if (hasVehicleLogSnapshot && el.closest('#view-vehicle-log')) return;

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

        restoreERPRecoveryState(state.customState.erpCheck);
        restoreTrackAccessRecoveryState(state.customState.trackAccess);
        restoreWorkLogRecoveryState(state.customState.workLog);
        restoreVehicleLogRecoveryState(state.customState.vehicleLog);
    }

}

function applySavableValue(el, key, value) {
    if (key === 'erpLocation') {
        erpState.selectedLocation = erpState.locations.some(loc => loc.name === value) ? value : null;
    }
}

// --- Navigation & Views ---

function switchMainTab(viewId) {
    syncMaterialLocationDefaults();
    document.querySelectorAll('#content .view-section').forEach(el => el.classList.remove('active'));

    const target = document.getElementById(viewId);
    if (target) target.classList.add('active');
    shellState.activeMainTab = viewId;

    if (viewId === 'view-track-access') {
        // 프리셋 데이터는 탭 진입 시점에만 현재 UI로 다시 적용합니다.
        syncTrackPresets(true);
    } else if (viewId === 'view-vehicle-log') {
        syncVehiclePresets(true);
    } else if (viewId === 'view-erp-check') {
        // Apply User Preference for Inspector Format
        // Logic: If user has a specific preference, override the current state?
        // Or only set initial state? "Initial value" was requested.
        // Let's set it based on config every time we switch to this tab.
        const uid = selectedUserId;
        if (uid && appConfig.users && appConfig.users[uid]) {
            const pref = appConfig.users[uid].profile.erpFormat || 'summary';
            erpControlsState.isListFormat = pref === 'list';
        }
        // 통합 UI 갱신 함수 호출 (ERP 점검 포함)
        refreshUI();

        // 4차 요구사항: 탭 진입 시 자동 새로고침(AHK 갱신 요청) 수행
        requestERPRefresh();
    } else if (viewId === 'view-work-log') {
        if (!isWorkLogInitialized) {
            renderWorkLogUI();
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
        authState.view = 'users';
    } else if (viewName === 'add') {
        loginCon.style.display = 'flex';
        authState.view = 'add';
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
    flushSettingsSave();
    let passwordsChanged = false;
    const uid = selectedUserId;
    if (uid && appConfig.users && appConfig.users[uid]) {
        const profile = appConfig.users[uid].profile || {};
        if ((profile.webPW || '') !== initialPasswords.webPW ||
            (profile.pw2 || '') !== initialPasswords.pw2 ||
            (profile.sapPW || '') !== initialPasswords.sapPW) {
            passwordsChanged = true;
        }
    }
    sendMessageToAHK({ command: 'exitSettings', passwordsChanged: passwordsChanged });
    switchView('app');
}

function switchSettingsTab(tabId) {
    settingsState.activeTab = tabId;
    if (tabId === 'tab-updates') {
        requestReleaseNotes();
    }
}

function requestReleaseNotes() {
    settingsState.releases.loading = true;
    settingsState.releases.error = '';
    sendMessageToAHK({ command: 'getReleaseNotes' });
}

function renderReleaseNotes(data, error = null) {
    settingsState.releases.loading = false;
    settingsState.releases.error = error ? String(error) : '';
    settingsState.releases.items = Array.isArray(data) ? data.map(release => ({
        tagName: String(release.tag_name || ''),
        date: release.published_at ? new Date(release.published_at).toLocaleDateString('ko-KR') : '',
        body: String(release.body || '내용 없음')
    })) : [];
}

function escapeHtml(text) {
    return text.replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

// ... (Rest of Login Logic remains same) ...

// --- Login Logic ---
function renderUserList(users) {
    loginUsers = Array.isArray(users) ? users : [];
    authState.users = loginUsers.map(user => ({ ...user }));
}

function selectUser(id) {
    selectedUserId = id;
    authState.selectedId = id;
}

function moveToPasswordView() {
    if (!selectedUserId) return;
    const user = loginUsers.find(u => u.id === selectedUserId);
    if (!user) return;

    authState.loginPw = '';
    authState.view = 'password';
    Vue.nextTick(() => document.getElementById('login-pw2')?.focus());
}

function backToUserList() {
    selectedUserId = null;
    authState.selectedId = null;
    authState.loginPw = '';
    authState.view = 'users';
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

    const enteredPw = authState.loginPw;

    const user = loginUsers.find(u => u.id === selectedUserId);
    if (!user) return;

    // Local 2nd Password Check
    if (user.pw2 !== enteredPw) {
        showNativeMsgBox('2차 비밀번호가 일치하지 않습니다.', '로그인 실패');
        authState.loginPw = '';
        Vue.nextTick(() => document.getElementById('login-pw2')?.focus());
        return;
    }

    sendMessageToAHK({ command: 'tryLogin', id: selectedUserId });
}

function handleLoginSuccess(profile) {
    configReadyUserId = null;
    pendingLocationOrderCheck = null;
    if (pendingLocationOrderReady?.employeeId !== profile.id) pendingLocationOrderReady = null;
    if (pendingAutoWorkerList?.employeeId !== profile.id) pendingAutoWorkerList = null;
    switchView('app');
    shellState.title = `통합자동화 - ${profile.name}`;

    // Set selectedUserId to current logged in user to ensure settings load correct user profile
    if (workLogState.ownerId !== null && String(workLogState.ownerId) !== String(profile.id)) {
        guntaeData = null;
        guntaeStatus = 'idle';
        workLogState.guntaeStatus = 'idle';
    }
    selectedUserId = profile.id;

    // Request full config
    sendMessageToAHK({ command: 'requestConfig' });

    // Auto-init Work Log View (Work Type Auto-selection)
    // Wait for Config Load (Race Condition Fix)
    // renderWorkLogUI(); // Removed here, moved to loadConfig
    isWorkLogInitialized = false; // Add this line to ensure reset

    switchMainTab('view-work-log');

    // Missing PW Check
    if (!profile.hideMissingPwWarning) {
        let missingTypes = [];
        if (!profile.webPW) missingTypes.push('통합pw');
        if (!profile.sapPW) missingTypes.push('SAPpw');

        if (missingTypes.length > 0) {
            showMissingPwModal(missingTypes);
        }
    }
}

function showMissingPwModal(missingTypes) {
    shellState.missingPassword.types = Array.isArray(missingTypes) ? [...missingTypes] : [];
    shellState.missingPassword.hideNextTime = false;
    shellState.missingPassword.visible = true;
}

function closeMissingPwModal() {
    shellState.missingPassword.visible = false;
    if (shellState.missingPassword.hideNextTime && selectedUserId && appConfig.users && appConfig.users[selectedUserId]) {
        appConfig.users[selectedUserId].profile.hideMissingPwWarning = true;
        sendMessageToAHK({ command: 'saveConfig', data: appConfig });
    }
}

function goSetMissingPw() {
    closeMissingPwModal();
    openSettings();
}

function deleteUser() {
    if (!selectedUserId) {
        showNativeMsgBox("삭제할 유저를 선택해주세요.");
        return;
    }
    sendMessageToAHK({ command: 'deleteUser', id: selectedUserId });
}

function submitNewUser() {
    const form = authState.form;
    const name = form.name.trim();
    const id = form.id.trim();
    const dept = form.department;
    const team = form.team;
    const webpw = form.webPW;
    const webpw2 = form.webPWConfirm;
    const pw2 = form.pw2;
    const pw2_confirm = form.pw2Confirm;
    const sappw = form.sapPW;
    const sappw2 = form.sapPWConfirm;

    if (!name || !id || !dept || !webpw || !pw2) {
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

    const department = SETTINGS_DEPARTMENTS.find(item => item.value === dept);
    const deptName = department?.name || '';
    const deptCodeval = department?.codeval || '';

    const newUser = { id, name, team, arbpl: dept, department: deptName, codeval: deptCodeval, webPW: webpw, pw2: pw2, sapPW: sappw };
    sendMessageToAHK({ command: 'addUser', data: newUser, defaults: buildNewUserDefaults() });

    switchView('login');
}

// 새 사용자 등록 시 채워 넣는 설정 기본값.
// 작업자·작업장소는 로그인 후 자동 조회하므로 여기서 지정하지 않는다.
const NEW_USER_DAILY_LOG_DEFAULTS = Object.freeze({
    safety: [
        { content: '전기 및 변전설비 일일점검', start: '09:30', end: '11:30' },
        { content: '전기 및 변전설비 점검, 기지구내 도보순회 점검', start: '13:30', end: '16:00' },
        { content: '전차선로 열차 순회점검 및 모타카 점검', start: '18:30', end: '20:00' },
        { content: '전차선로 점검', start: '01:00', end: '04:00' }
    ],
    driverCheck: false,
    // 책임자는 분소마다 달라 비워 둔다.
    generalWork: [
        { workType: '주간', category: '점검업무', content: 'E/R 일일점검', manager: '', start: '09:30', end: '11:30' },
        { workType: '주간', category: '점검업무', content: '- 고압반 각종 계기류의 지시상태 : 이상없음', manager: '', start: '09:30', end: '11:30' }
    ],
    safetyEdu: {
        Mon: '안전보건 11대 수칙',
        Tue: '전기원 안전수칙',
        Wed: '일반전기작업 안전수칙',
        Thu: '시설물 안전점검 수칙',
        Fri: '작업장 내 정리 정돈 안전 수칙',
        Sat: '사다리 작업 및 정전작업 요령',
        Sun: '안전 확보 요강 및 기타 작업에 필요한 사항 등'
    },
    autoInputTime: ''
});

function buildDefaultHotkeys() {
    return defaultHotkeys.map(item => ({ action: item.action, key: item.key, desc: item.desc, enabled: true }));
}

function buildNewUserDefaults() {
    return {
        hotkeys: buildDefaultHotkeys(),
        dailyLogDefaults: JSON.parse(JSON.stringify(NEW_USER_DAILY_LOG_DEFAULTS))
    };
}

// 단축키 목록이 비어 있으면 설정 화면에는 모두 켜진 것으로 보이지만 실제로는 등록되지 않는다.
// 로그인 직후 기본 단축키를 저장해 화면 표시와 실제 동작을 맞춘다.
function ensureUserHotkeyDefaults() {
    const user = appConfig?.users?.[selectedUserId];
    if (!user || (Array.isArray(user.hotkeys) && user.hotkeys.length)) return;
    user.hotkeys = buildDefaultHotkeys();
    sendMessageToAHK({ command: 'saveConfig', data: appConfig });
}

function resetAddUserForm() {
    Object.assign(authState.form, {
        name: '', id: '', department: '', team: 'A조',
        webPW: '', webPWConfirm: '', pw2: '', pw2Confirm: '', sapPW: '', sapPWConfirm: ''
    });
}

// --- Settings Logic (Vue) ---
let saveTimeout = null;
let initialPasswords = { webPW: '', pw2: '', sapPW: '' };
const SETTINGS_SAVE_DELAY_MS = 200;
const SETTINGS_TEAMS = Object.freeze(['A조', 'B조', 'C조', 'D조', '일근']);
const SETTINGS_LOCATION_TYPES = Object.freeze(['변전소', '전기실(그룹1)', '전기실(그룹2)', '전기실(그룹3)', '기타업무']);
const SETTINGS_WORK_CATEGORIES = Object.freeze(['전체', '내부업무', '점검업무', '유지보수', '협조사항']);
const SETTINGS_DEPARTMENTS = Object.freeze([
    { value: '5111', name: '동래변전분소', codeval: '' },
    { value: '5112', name: '전포변전분소', codeval: '' },
    { value: '5113', name: '구포변전분소', codeval: '' },
    { value: '5121', name: '신평전기분소', codeval: '' },
    { value: '5122', name: '부산진전기분소', codeval: '' },
    { value: '5123', name: '동래전기분소', codeval: '' },
    { value: '5124', name: '노포전기분소', codeval: '' },
    { value: '5125', name: '장산전기분소', codeval: '' },
    { value: '5126', name: '광안전기분소', codeval: '' },
    { value: '5127', name: '전포전기분소', codeval: '' },
    { value: '5128', name: '구명전기분소', codeval: '' },
    { value: '5129', name: '호포전기분소', codeval: '2612021' },
    { value: '5130', name: '미남전기분소', codeval: '' },
    { value: '5131', name: '대저전기분소', codeval: '' },
    { value: '5132', name: '다대전기관리소', codeval: '' },
    { value: '5133', name: '남양산전기관리소', codeval: '' }
]);

const materialLocationState = Vue.reactive({
    ownerId: null, defaultLocation: '', locations: [], keyword: '', searchedKeyword: '', rows: [],
    loading: false, searched: false, error: '', requestId: '', requestOwner: null, detail: null
});

function materialDepartmentOptions() {
    return SETTINGS_DEPARTMENTS.map(option => ({
        value: option.value,
        name: option.value === '5133' ? '남양산관리' : option.value === '5132' ? '다대관리'
            : option.name.replace(/분소|관리소/g, '')
    }));
}

function syncMaterialLocationDefaults() {
    const owner = selectedUserId;
    const code = String(appConfig.users?.[owner]?.profile?.arbpl || '');
    if (materialLocationState.ownerId === owner && materialLocationState.defaultLocation === code) return;
    Object.assign(materialLocationState, {
        ownerId: owner, defaultLocation: code,
        locations: SETTINGS_DEPARTMENTS.some(option => option.value === code) ? [code] : [],
        keyword: '', searchedKeyword: '', rows: [], loading: false, searched: false, error: '', requestId: '', requestOwner: null, detail: null
    });
}

function searchMaterialLocations() {
    syncMaterialLocationDefaults();
    const state = materialLocationState;
    if (state.loading) return;
    if (!selectedUserId || configReadyUserId !== selectedUserId) {
        state.error = '로그인 정보를 불러온 후 검색해주세요.';
        return;
    }
    if (!state.locations.length) { state.error = '저장위치를 선택해주세요.'; return; }
    state.requestId = `material-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    state.requestOwner = selectedUserId;
    state.loading = true;
    state.error = '';
    state.rows = [];
    state.detail = null;
    state.searchedKeyword = state.keyword.trim();
    sendMessageToAHK({ command: 'searchMaterialLocations', employeeId: selectedUserId,
        requestId: state.requestId, locations: [...state.locations], search: state.keyword.trim() });
}

function handleMaterialLocationResult(msg) {
    const state = materialLocationState;
    if (msg.requestId !== state.requestId || msg.employeeId !== state.requestOwner
        || selectedUserId !== state.requestOwner || configReadyUserId !== state.requestOwner) return;
    state.loading = false;
    state.searched = true;
    state.requestId = '';
    state.error = msg.error || '';
    state.rows = !state.error && Array.isArray(msg.rows)
        ? msg.rows.filter(row => Number(row.newStock) + Number(row.usedStock) !== 0) : [];
    if (!state.error && !Array.isArray(msg.rows)) state.error = '조회 결과 형식이 올바르지 않습니다.';
}

function materialKeywordParts(value, query) {
    const text = String(value || ''), lower = text.toLowerCase(), ranges = [];
    const tokens = [...new Set(String(query || '').trim().toLowerCase().split(/\s+/).filter(Boolean))];
    for (const token of tokens) {
        let start = lower.indexOf(token);
        while (start >= 0) {
            ranges.push([start, start + token.length]);
            start = lower.indexOf(token, start + 1);
        }
    }
    ranges.sort((a, b) => a[0] - b[0]);
    const merged = [];
    for (const range of ranges) {
        const last = merged[merged.length - 1];
        if (last && range[0] <= last[1]) last[1] = Math.max(last[1], range[1]);
        else merged.push([...range]);
    }
    const parts = [];
    let cursor = 0;
    for (const [start, end] of merged) {
        if (start > cursor) parts.push({ text: text.slice(cursor, start), matched: false });
        parts.push({ text: text.slice(start, end), matched: true });
        cursor = end;
    }
    if (cursor < text.length || !parts.length) parts.push({ text: text.slice(cursor), matched: false });
    return parts;
}

function initMaterialLocationApp() {
    Vue.createApp({
        data() { return { state: materialLocationState, options: materialDepartmentOptions(), backdropPressed: false,
            descriptionExtraWidth: 0, tableBaseWidth: 0, columnDrag: null, columnResizeObserver: null }; },
        mounted() {
            document.addEventListener('pointerdown', this.closeLocationDropdown);
            this.columnResizeObserver = new ResizeObserver(this.updateColumnBounds);
            this.columnResizeObserver.observe(this.$refs.tableWrapper);
            this.updateColumnBounds();
        },
        beforeUnmount() {
            document.removeEventListener('pointerdown', this.closeLocationDropdown);
            this.columnResizeObserver?.disconnect();
            this.endColumnResize();
        },
        watch: {
            'state.detail'(value) { if (!value && this.$refs.detailDialog.open) this.$refs.detailDialog.close(); }
        },
        computed: {
            tableStyle() { return this.descriptionExtraWidth > 0 && this.tableBaseWidth
                ? { width: `${this.tableBaseWidth + this.descriptionExtraWidth}px` } : { width: '100%' }; },
            descriptionColumnStyle() { return this.descriptionExtraWidth > 0 && this.tableBaseWidth
                ? { width: `${this.tableBaseWidth - 352 + this.descriptionExtraWidth}px` } : {}; },
            selectionLabel() {
                const names = this.state.locations.map(code => this.options.find(option => option.value === code)?.name).filter(Boolean);
                return names.length > 1 ? `${names[0]} 외 ${names.length - 1}곳` : names[0] || '저장위치 선택';
            }
        },
        methods: {
            updateColumnBounds() {
                const width = this.$refs.tableWrapper.clientWidth;
                if (width > 352) this.tableBaseWidth = width;
            },
            startColumnResize(event) {
                if (event.button !== 0) return;
                event.preventDefault();
                this.updateColumnBounds();
                this.columnDrag = { pointerId: event.pointerId, startX: event.clientX,
                    startWidth: this.descriptionExtraWidth, target: event.currentTarget };
                event.currentTarget.setPointerCapture(event.pointerId);
                document.documentElement.classList.add('material-column-resizing');
            },
            moveColumnResize(event) {
                if (this.columnDrag?.pointerId !== event.pointerId) return;
                this.descriptionExtraWidth = Math.max(0, this.columnDrag.startWidth + event.clientX - this.columnDrag.startX);
            },
            endColumnResize(event) {
                if (event && this.columnDrag?.pointerId !== event.pointerId) return;
                const drag = this.columnDrag;
                this.columnDrag = null;
                if (drag?.target.hasPointerCapture(drag.pointerId)) drag.target.releasePointerCapture(drag.pointerId);
                document.documentElement.classList.remove('material-column-resizing');
            },
            keywordParts(value) { return materialKeywordParts(value, this.state.searchedKeyword); },
            isDetailBackdrop(event) {
                const dialog = this.$refs.detailDialog, rect = dialog.getBoundingClientRect();
                return event.target === dialog && (event.clientX < rect.left || event.clientX > rect.right
                    || event.clientY < rect.top || event.clientY > rect.bottom);
            },
            onDetailPointerDown(event) { this.backdropPressed = this.isDetailBackdrop(event); },
            closeDetailOnBackdrop(event) {
                if (this.backdropPressed && this.isDetailBackdrop(event)) this.$refs.detailDialog.close();
                this.backdropPressed = false;
            },
            closeLocationDropdown(event) {
                const dropdown = this.$refs.locationDropdown;
                if (dropdown?.open && !dropdown.contains(event.target)) dropdown.open = false;
            },
            openDetail(row) {
                this.state.detail = { ...row };
                this.$nextTick(() => { if (this.state.detail && !this.$refs.detailDialog.open) this.$refs.detailDialog.showModal(); });
            },
            search() { this.$refs.locationDropdown.open = false; searchMaterialLocations(); },
            selectAll() { this.state.locations = this.options.map(option => option.value); },
            departmentName(code) { return this.options.find(option => option.value === code)?.name || code; },
            materialNumber(value) { return String(value).replace(/^0+(?=\d)/, ''); },
            quantity(value) { return value === null || value === undefined ? '' : Number(value).toLocaleString('ko-KR', { maximumFractionDigits: 3 }); }
        }
    }).mount('#view-material-location');
}

function emptySafetyDefaults() {
    return Array.from({ length: 4 }, () => ({ content: '', start: '', end: '' }));
}

function emptySafetyEducation() {
    return { Mon: '', Tue: '', Wed: '', Thu: '', Fri: '', Sat: '', Sun: '' };
}

const settingsState = Vue.reactive({
    activeTab: 'tab-user',
    loadedUserId: null,
    profile: {
        name: '', id: '', arbpl: '', department: '', codeval: '', team: '',
        webPW: '', pw2: '', sapPW: '', erpFormat: 'summary'
    },
    workers: [],
    locations: [],
    hotkeys: [],
    dailyLogDefaults: {
        safety: emptySafetyDefaults(),
        driverCheck: false,
        generalWork: [],
        safetyEdu: emptySafetyEducation(),
        autoInputTime: ''
    },
    generalWorkShift: '주간',
    safetyEduDay: 'Mon',
    cloudBusy: false,
    sessionControlsEnabled: false,
    importingWorkers: false,
    importingLocations: false,
    releases: { loading: false, error: '', items: [] }
});

function normalizeWorker(worker = {}) {
    return {
        name: String(worker.name || ''),
        id: String(worker.id || ''),
        team: SETTINGS_TEAMS.includes(worker.team) ? worker.team : '일근',
        phone: String(worker.phone || ''),
        isManager: worker.isManager === true || worker.isManager === 1 || worker.isManager === '1',
        driverRole: ['-', '정', '부'].includes(worker.driverRole) ? worker.driverRole : '-'
    };
}

function normalizeLocation(location = {}) {
    return {
        name: String(location.name || ''),
        order: String(location.order || ''),
        type: SETTINGS_LOCATION_TYPES.includes(location.type) ? location.type : '기타업무'
    };
}

function normalizeGeneralWork(row = {}) {
    return {
        workType: row.workType === '야간' ? '야간' : '주간',
        category: SETTINGS_WORK_CATEGORIES.includes(row.category) ? row.category : '전체',
        content: String(row.content || ''),
        manager: String(row.manager || ''),
        start: String(row.start || ''),
        end: String(row.end || '')
    };
}

function getSettingsHotkeys(savedHotkeys = []) {
    return defaultHotkeys.map(def => {
        const saved = savedHotkeys.find(item => item.action === def.action);
        return {
            ...def,
            actionName: hotkeyTranslations[def.action] || def.action,
            enabled: saved ? saved.enabled !== false : true
        };
    });
}

function initSettingsApp() {
    Vue.createApp({
        data() {
            return {
                settings: settingsState,
                departments: SETTINGS_DEPARTMENTS,
                teams: SETTINGS_TEAMS,
                locationTypes: SETTINGS_LOCATION_TYPES,
                workCategories: SETTINGS_WORK_CATEGORIES
            };
        },
        computed: {
            visibleGeneralWork() {
                return this.settings.dailyLogDefaults.generalWork
                    .filter(row => row.workType === this.settings.generalWorkShift);
            }
        },
        methods: {
            close: closeSettings,
            selectTab: switchSettingsTab,
            changed: autoSaveSettings,
            importCloud: importUserSettings,
            exportCloud: exportUserSettings,
            importWorkers: importWorkersFromHeadless,
            importLocations: importLocationsFromHeadless,
            addWorker() { addWorkerRow(); },
            removeWorker(index) {
                settingsState.workers.splice(index, 1);
                autoSaveSettings();
            },
            managerChanged(index) {
                if (settingsState.workers[index]?.isManager) {
                    settingsState.workers.forEach((worker, workerIndex) => {
                        if (workerIndex !== index) worker.isManager = false;
                    });
                }
                autoSaveSettings();
            },
            driverChanged(index) {
                const worker = settingsState.workers[index];
                if (!worker || worker.driverRole === '-') {
                    autoSaveSettings();
                    return;
                }
                const duplicate = settingsState.workers.some((other, otherIndex) =>
                    otherIndex !== index && other.team === worker.team && other.driverRole === worker.driverRole);
                if (duplicate) {
                    showNativeMsgBox(`${worker.team}에는 이미 ${worker.driverRole}운전원이 있습니다.`);
                    worker.driverRole = '-';
                }
                autoSaveSettings();
            },
            workerIdInput(index, event) {
                settingsState.workers[index].id = String(event.target.value || '').replace(/\D/g, '').slice(0, 6);
                autoSaveSettings();
            },
            workerPhoneInput(index, event) {
                settingsState.workers[index].phone = formatPhoneValue(event.target.value);
                autoSaveSettings();
            },
            addLocation() { addLocationRow(); },
            removeLocation(index) {
                settingsState.locations.splice(index, 1);
                autoSaveSettings();
            },
            locationOrderInput(index, event) {
                settingsState.locations[index].order = String(event.target.value || '').replace(/\D/g, '').slice(0, 8);
                autoSaveSettings();
            },
            addGeneralWork() { addGeneralWorkRow(); },
            removeGeneralWork(row) {
                const index = settingsState.dailyLogDefaults.generalWork.indexOf(row);
                if (index >= 0) settingsState.dailyLogDefaults.generalWork.splice(index, 1);
                autoSaveSettings();
            },
            formatTime(target, field, event) {
                target[field] = formatTimeValue(event.target.value);
                autoSaveSettings();
            },
            hotkeyParts(key) {
                const parts = [];
                if (key.includes('#')) parts.push('Win');
                if (key.includes('^')) parts.push('Ctrl');
                if (key.includes('!')) parts.push('Alt');
                if (key.includes('+')) parts.push('Shift');
                parts.push(key.replace(/[#^!+]/g, '').toUpperCase());
                return parts;
            }
        }
    }).mount('#settings-view');
}

function autoSaveSettings() {
    if (saveTimeout) clearTimeout(saveTimeout);
    saveTimeout = setTimeout(() => {
        saveTimeout = null;
        saveSettings();
    }, SETTINGS_SAVE_DELAY_MS);
}

function flushSettingsSave() {
    if (!saveTimeout) return;
    clearTimeout(saveTimeout);
    saveTimeout = null;
    saveSettings();
}

function loadSettingsToUI() {
    const uid = selectedUserId;
    if (!uid || !appConfig.users) return;
    const user = appConfig.users[uid];
    if (!user) return;
    const profile = user.profile || {};
    const defaults = user.dailyLogDefaults || {};
    initialPasswords = {
        webPW: profile.webPW || '',
        pw2: profile.pw2 || '',
        sapPW: profile.sapPW || ''
    };
    const safety = emptySafetyDefaults();
    (Array.isArray(defaults.safety) ? defaults.safety : []).slice(0, 4).forEach((row, index) => {
        safety[index] = {
            content: String(row?.content || ''),
            start: String(row?.start || ''),
            end: String(row?.end || '')
        };
    });
    const safetyEdu = { ...emptySafetyEducation(), ...(defaults.safetyEdu || {}) };
    Object.assign(settingsState.profile, {
        name: String(profile.name || ''),
        id: String(profile.id || uid),
        arbpl: String(profile.arbpl || ''),
        department: String(profile.department || ''),
        codeval: String(profile.codeval || ''),
        team: SETTINGS_TEAMS.includes(profile.team) ? profile.team : '일근',
        webPW: String(profile.webPW || ''),
        pw2: String(profile.pw2 || ''),
        sapPW: String(profile.sapPW || ''),
        erpFormat: profile.erpFormat === 'list' ? 'list' : 'summary'
    });
    settingsState.workers = (appConfig.appSettings?.colleagues || []).map(normalizeWorker);
    settingsState.locations = (appConfig.appSettings?.locations || []).map(normalizeLocation);
    settingsState.hotkeys = getSettingsHotkeys(user.hotkeys || []);
    settingsState.dailyLogDefaults.safety = safety;
    settingsState.dailyLogDefaults.driverCheck = !!defaults.driverCheck;
    settingsState.dailyLogDefaults.generalWork = (defaults.generalWork || []).map(normalizeGeneralWork);
    settingsState.dailyLogDefaults.safetyEdu = safetyEdu;
    settingsState.dailyLogDefaults.autoInputTime = String(defaults.autoInputTime || '');
    settingsState.generalWorkShift = '주간';
    settingsState.safetyEduDay = 'Mon';
    settingsState.loadedUserId = uid;
    currentSafetyEduData = settingsState.dailyLogDefaults.safetyEdu;
}

function saveSettings() {
    try {
        const uid = selectedUserId;
        if (!uid || !appConfig.users || settingsState.loadedUserId !== uid) return;
        const user = appConfig.users[uid];
        if (!user) return;
        if (!user.profile) user.profile = {};
        const department = SETTINGS_DEPARTMENTS.find(item => item.value === settingsState.profile.arbpl);
        Object.assign(user.profile, {
            arbpl: settingsState.profile.arbpl,
            department: department?.name || '',
            codeval: department?.codeval || '',
            team: settingsState.profile.team,
            webPW: settingsState.profile.webPW,
            pw2: settingsState.profile.pw2,
            sapPW: settingsState.profile.sapPW,
            erpFormat: settingsState.profile.erpFormat
        });
        if (user.profile.webPW && user.profile.sapPW) user.profile.hideMissingPwWarning = false;
        if (!appConfig.appSettings) appConfig.appSettings = {};
        appConfig.appSettings.colleagues = settingsState.workers.map(worker => ({
            ...normalizeWorker(worker), isManager: worker.isManager ? 1 : 0
        }));
        appConfig.appSettings.locations = settingsState.locations.map(normalizeLocation);
        user.hotkeys = settingsState.hotkeys.map(item => ({
            action: item.action, key: item.key, desc: item.desc, enabled: !!item.enabled
        }));
        user.dailyLogDefaults = {
            safety: settingsState.dailyLogDefaults.safety.map(row => ({ ...row })),
            driverCheck: !!settingsState.dailyLogDefaults.driverCheck,
            generalWork: settingsState.dailyLogDefaults.generalWork.map(normalizeGeneralWork),
            safetyEdu: { ...settingsState.dailyLogDefaults.safetyEdu },
            autoInputTime: settingsState.dailyLogDefaults.autoInputTime
        };
        currentSafetyEduData = user.dailyLogDefaults.safetyEdu;
        sendMessageToAHK({ command: 'saveConfig', data: appConfig });
    } catch (e) {
        showNativeMsgBox('설정 저장 중 오류: ' + e.message);
    }
}

// --- Firebase User Settings Backup ---
const PRIVATE_SETTING_KEYS = new Set([
    'webPW', 'pw2', 'sapPW', 'password', 'passwordConfirm'
]);
const IDENTITY_PROFILE_KEYS = new Set(['id', 'name']);
const UNSAFE_OBJECT_KEYS = new Set(['__proto__', 'prototype', 'constructor']);

function isPrivateSettingKey(key) {
    const normalizedKey = String(key).toLowerCase();
    return PRIVATE_SETTING_KEYS.has(key)
        || normalizedKey.includes('password')
        || /pw\d*$/.test(normalizedKey);
}

function cloneExportableSettings(value) {
    if (value === undefined) {
        return null;
    }
    if (Array.isArray(value)) {
        return value.map(cloneExportableSettings);
    }
    if (!value || typeof value !== 'object') {
        return value;
    }

    const result = {};
    Object.keys(value).forEach(key => {
        if (UNSAFE_OBJECT_KEYS.has(key) || isPrivateSettingKey(key)) return;
        result[key] = cloneExportableSettings(value[key]);
    });
    return result;
}

function buildExportableUserSettings(user) {
    const settings = {};

    Object.keys(user || {}).forEach(key => {
        if (key === 'profile' || key === 'appSettings' || UNSAFE_OBJECT_KEYS.has(key) || isPrivateSettingKey(key)) return;
        // Legacy presets is separate from trackPresets; omit only its empty placeholder.
        if (key === 'presets' && user[key] && typeof user[key] === 'object' && Object.keys(user[key]).length === 0) return;
        settings[key] = cloneExportableSettings(user[key]);
    });

    const profile = {};
    Object.keys(user?.profile || {}).forEach(key => {
        if (IDENTITY_PROFILE_KEYS.has(key) || UNSAFE_OBJECT_KEYS.has(key) || isPrivateSettingKey(key)) return;
        profile[key] = cloneExportableSettings(user.profile[key]);
    });
    settings.profile = profile;

    return settings;
}

function formatSettingsExportPreview(employeeId, settings) {
    const labels = {
        profile: '개인 설정', arbpl: '분소 코드', department: '분소', codeval: '분소 부가 코드',
        team: '근무조', erpFormat: '점검자 입력 형식', autoExit: '자동 종료',
        hideMissingPwWarning: '비밀번호 누락 안내 숨김', hotkeys: '단축키 설정',
        presets: '기존 선로출입 프리셋', trackPresets: '선로출입 프리셋', vehiclePresets: '차량일지 프리셋',
        dailyLogDefaults: '일지 프리셋 기본값', safety: '안전관리 기본값', driverCheck: '운전적합성 검사 방식',
        generalWork: '일반업무 기본값', safetyEdu: '요일별 안전교육', autoInputTime: '열차순회 자동입력 예약시간',
        action: '기능', key: '단축키', desc: '설명', enabled: '활성', content: '내용',
        start: '시작', end: '종료', workType: '근무', category: '분류', manager: '책임자'
    };
    const trackLabels = {
        workType: '작업구분', content: '작업내역', workFrom: '작업구간 (From)', workTo: '작업구간 (To)',
        driverName: '운전원 이름', driverPhone: '운전원 연락처', workerName: '작업자 이름', workerPhone: '작업자 연락처',
        safetyName: '철도운행안전관리자 이름', safetyPhone: '철도운행안전관리자 연락처',
        supervisorName: '감독자 이름', supervisorId: '감독자 사번', workStart: '작업시간 시작', workEnd: '작업시간 종료',
        opStart: '운행시간 시작', opEnd: '운행시간 종료', line: '호선', trackType: '선로구분',
        trackCutoff: '선로차단여부', agreementNo: '철도운행협의서 No.', agreementNoDate: '협의번호 저장일',
        totalCount: '총원 (인원 구성)', stationInput: '출입역 입력'
    };
    const vehicleLabels = {
        driver: '운전자', point1: '작업구간 (시점)', point2: '작업구간 (종점)', trackType: '선로구분',
        startTime: '운행시간 시작', endTime: '운행시간 종료', approveNo: '승인정보 승인번호',
        dept: '승인정보 부서', approver: '승인정보 승인자', runTime: '가동시간', distance: '적산계',
        content: '내용', remarks: '비고'
    };
    const label = (key, context) => (context === 'trackPresets' ? trackLabels[key] : context === 'vehiclePresets' ? vehicleLabels[key] : null) || labels[key] || String(key);
    const escapeHtml = value => String(value).replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
    const valueText = value => {
        if (value === null || value === undefined || value === '') return '(없음)';
        if (typeof value === 'boolean') {
            return value ? '사용' : '사용 안 함';
        }
        return String(value);
    };
    const cell = value => escapeHtml(valueText(value));
    const isObject = value => value && typeof value === 'object' && !Array.isArray(value);
    const isFlatObject = value => isObject(value) && Object.values(value).every(item => !item || typeof item !== 'object');
    const selectText = (id, value) => {
        const select = document.getElementById(id);
        const option = select && [...select.options].find(item => item.value === String(value));
        return option ? option.textContent.trim() : value;
    };
    const displayValue = (value, key, context) => {
        if (context === 'trackPresets') {
            const selects = { workType: 'ta-work-type', line: 'ta-line', trackType: 'ta-track-type' };
            if (selects[key]) return selectText(selects[key], value);
        }
        return value;
    };
    const pairTable = (items, context, className = '') => {
        if (!items.length) return '<p class="empty">(없음)</p>';
        const widths = className === 'safety-edu' ? [7, 43, 7, 43]
            : className === 'vehicle-preset' ? [13, 27, 13, 47] : [20, 30, 20, 30];
        const colgroup = `<colgroup>${widths.map(width => `<col style="width:${width}%">`).join('')}</colgroup>`;
        let rows = '';
        for (let i = 0; i < items.length; i += 2) {
            const pair = items.slice(i, i + 2);
            rows += '<tr>' + pair.map(([key, value]) => `<th>${escapeHtml(label(key, context))}</th><td>${cell(displayValue(value, key, context))}</td>`).join('')
                + (pair.length === 1 ? '<th></th><td></td>' : '') + '</tr>';
        }
        return `<table class="pairs ${className}">${colgroup}<tbody>${rows}</tbody></table>`;
    };
    const trackPresetTable = preset => {
        // Match the two columns and top-to-bottom field order of the Track Access screen.
        const left = [
            ['작업구분', ['workType']], ['작업내역', ['content']], ['작업구간 (From ~ To)', ['workFrom', 'workTo']],
            ['운전원 (이름 / 연락처)', ['driverName', 'driverPhone']], ['작업자 (이름 / 연락처)', ['workerName', 'workerPhone']],
            ['철도운행안전관리자 (이름 / 연락처)', ['safetyName', 'safetyPhone']], ['감독자 (이름 / 사번)', ['supervisorName', 'supervisorId']]
        ];
        const right = [
            ['작업시간 (시작 ~ 종료)', ['workStart', 'workEnd']], ['운행시간 (시작 ~ 종료)', ['opStart', 'opEnd']],
            ['호선', ['line']], ['선로구분', ['trackType']], ['선로차단여부', ['trackCutoff']],
            ['철도운행협의서 No.', ['agreementNo']], ['총원 (인원 구성)', ['totalCount']],
            ['출입역 입력', ['stationInput']], ['협의번호 저장일', ['agreementNoDate']]
        ];
        const seen = new Set([...left, ...right].flatMap(([, keys]) => keys));
        Object.keys(preset).forEach(key => {
            if (!seen.has(key)) right.push([label(key, 'trackPresets'), [key]]);
        });
        const groupText = keys => keys.map(key => valueText(displayValue(preset[key], key, 'trackPresets'))).join(' ~ ');
        const rows = Array.from({ length: Math.max(left.length, right.length) }, (_, index) => {
            const cells = [left[index], right[index]].map(group => group
                ? `<th>${escapeHtml(group[0])}</th><td>${escapeHtml(groupText(group[1]))}</td>`
                : '<th></th><td></td>');
            return `<tr>${cells.join('')}</tr>`;
        }).join('');
        return `<table class="pairs track-preset"><colgroup><col style="width:27%"><col style="width:32%"><col style="width:19%"><col style="width:22%"></colgroup><tbody>${rows}</tbody></table>`;
    };
    const rowTable = (items, context) => {
        if (!items.length) return '<p class="empty">(없음)</p>';
        const keys = [...new Set(items.flatMap(item => Object.keys(item)))];
        const generalWorkWidths = { workType: 9, category: 10, content: 49, manager: 10, start: 8, end: 8 };
        const useGeneralWidths = context === 'generalWork' && keys.length === 6 && keys.every(key => Object.prototype.hasOwnProperty.call(generalWorkWidths, key));
        const colgroup = useGeneralWidths ? `<colgroup><col style="width:6%">${keys.map(key => `<col style="width:${generalWorkWidths[key]}%">`).join('')}</colgroup>` : '';
        return `<table class="${context === 'generalWork' ? 'general-work' : ''}">${colgroup}<thead><tr><th class="number">번호</th>${keys.map(key => `<th>${escapeHtml(label(key, context))}</th>`).join('')}</tr></thead><tbody>`
            + items.map((item, index) => `<tr><td class="number">${index + 1}</td>${keys.map(key => `<td>${cell(displayValue(item[key], key, context))}</td>`).join('')}</tr>`).join('')
            + '</tbody></table>';
    };
    const render = (value, name, context = '', depth = 0) => {
        const title = escapeHtml(label(name, context));
        const heading = depth ? 'h4' : 'h3';
        const nextContext = ['trackPresets', 'vehiclePresets', 'generalWork', 'safetyEdu'].includes(name) ? name : context;
        if (Array.isArray(value)) {
            if (value.every(isFlatObject)) return `<section><${heading}>${title}</${heading}>${rowTable(value, nextContext)}</section>`;
            if (value.every(item => !item || typeof item !== 'object')) {
                return `<section><${heading}>${title}</${heading}><table><tbody>${value.map((item, index) => `<tr><th>${index + 1}</th><td>${cell(item)}</td></tr>`).join('')}</tbody></table></section>`;
            }
            return `<section><${heading}>${title}</${heading}>${value.map((item, index) => render(item, `${index + 1}번째`, nextContext, depth + 1)).join('')}</section>`;
        }
        if (isObject(value)) {
            const entries = Object.entries(value);
            if (context === 'trackPresets' && depth > 0 && entries.every(([, item]) => !item || typeof item !== 'object')) {
                return `<section><${heading}>${title}</${heading}>${trackPresetTable(value)}</section>`;
            }
            if (entries.every(([, item]) => !item || typeof item !== 'object')) {
                const tableClass = nextContext === 'safetyEdu' ? 'safety-edu' : nextContext === 'vehiclePresets' ? 'vehicle-preset' : '';
                return `<section><${heading}>${title}</${heading}>${pairTable(entries, nextContext, tableClass)}</section>`;
            }
            return `<section><${heading}>${title}</${heading}>${entries.map(([key, item]) => render(item, key, nextContext, depth + 1)).join('')}</section>`;
        }
        return `<section><${heading}>${title}</${heading}><p>${cell(value)}</p></section>`;
    };
    const body = Object.entries(settings).map(([key, value]) => render(value, key)).join('');
    return `<!doctype html><html><head><meta charset="utf-8"><style>
        body{font-family:'Malgun Gothic',sans-serif;font-size:13px;color:#202a34;margin:14px 18px;background:#fff}
        h2{font-size:16px;margin:0 0 12px}h3{font-size:14px;margin:17px 0 7px;color:#183c59;border-bottom:1px solid #b8cbdc;padding-bottom:4px}
        h4{font-size:13px;margin:12px 0 5px;color:#34495e}section{margin-bottom:8px}section section{margin-left:10px}
        table{border-collapse:collapse;width:100%;margin:5px 0 10px;table-layout:fixed}th,td{border:1px solid #ccd5dc;padding:5px 7px;text-align:left;vertical-align:top;word-wrap:break-word}
        thead th,.pairs th{background:#eef3f7;font-weight:600}.track-preset th{font-size:12px;white-space:nowrap}
        .general-work .number{white-space:nowrap}.number{width:38px;text-align:center}.empty{color:#777;margin:5px 0}
    </style></head><body><h2>사번 : ${escapeHtml(employeeId)}</h2>${body}</body></html>`;
}

const pendingNativeConfirmations = new Map();

function requestNativeConfirmation(text, title = '확인', question = '위 내용으로 클라우드에 내보내기 하시겠습니까?') {
    const requestId = (window.crypto && typeof window.crypto.randomUUID === 'function')
        ? window.crypto.randomUUID()
        : `${Date.now()}-${Math.random()}`;
    return new Promise(resolve => {
        pendingNativeConfirmations.set(requestId, resolve);
        sendMessageToAHK({ command: 'confirmMsgbox', requestId, text, title, question });
    });
}

function requestWorkerImportConfirmation(text, title = '작업자 불러오기 확인') {
    const requestId = (window.crypto && typeof window.crypto.randomUUID === 'function')
        ? window.crypto.randomUUID()
        : `${Date.now()}-${Math.random()}`;
    return new Promise(resolve => {
        pendingNativeConfirmations.set(requestId, resolve);
        sendMessageToAHK({ command: 'confirmWorkerImport', requestId, text, title });
    });
}

function requestLocationImportConfirmation(text) {
    const requestId = `${Date.now()}-${Math.random()}`;
    return new Promise(resolve => {
        pendingNativeConfirmations.set(requestId, resolve);
        sendMessageToAHK({ command: 'confirmLocationImport', requestId, text, title: '작업장소 불러오기 확인' });
    });
}

function setSettingsCloudButtonsDisabled(disabled) {
    settingsState.cloudBusy = !!disabled;
}

function getCurrentSettingsUser() {
    const employeeId = String(selectedUserId || '').trim();
    const user = employeeId && appConfig.users ? appConfig.users[employeeId] : null;
    if (!employeeId || !user) {
        throw new Error('로그인한 사용자 정보를 확인할 수 없습니다.');
    }
    if (!/^\d+$/.test(employeeId)) {
        throw new Error('사번 형식이 올바르지 않습니다.');
    }
    return { employeeId, user };
}

async function exportUserSettings() {
    try {
        if (saveTimeout) {
            clearTimeout(saveTimeout);
            saveTimeout = null;
        }
        saveSettings();

        const { employeeId, user } = getCurrentSettingsUser();
        const settings = buildExportableUserSettings(user);
        setSettingsCloudButtonsDisabled(true);
        if (!await requestNativeConfirmation(formatSettingsExportPreview(employeeId, settings), '설정 내보내기 확인')) {
            setSettingsCloudButtonsDisabled(false);
            return;
        }

        const firestore = getFirebaseApp().firestore();
        await firestore.collection(USER_SETTINGS_COLLECTION).doc(employeeId).set({
            schemaVersion: 1,
            employeeId,
            settings,
            updatedAt: window.firebase.firestore.FieldValue.serverTimestamp()
        });
        showNativeMsgBox('사용자 설정을 Firebase에 저장했습니다.', '설정 내보내기');
    } catch (error) {
        console.error('Failed to export user settings:', error);
        showNativeMsgBox(`설정을 내보내지 못했습니다.\n${error?.message || String(error)}`, '설정 내보내기 오류');
    } finally {
        setSettingsCloudButtonsDisabled(false);
    }
}

function applyImportedUserSettings(user, importedSettings) {
    const safeSettings = cloneExportableSettings(importedSettings || {});
    const importedProfile = safeSettings.profile && typeof safeSettings.profile === 'object'
        ? safeSettings.profile
        : {};

    Object.keys(safeSettings).forEach(key => {
        if (key === 'profile' || key === 'appSettings' || UNSAFE_OBJECT_KEYS.has(key) || isPrivateSettingKey(key)) return;
        user[key] = safeSettings[key];
    });

    if (!user.profile) user.profile = {};
    Object.keys(importedProfile).forEach(key => {
        if (IDENTITY_PROFILE_KEYS.has(key) || UNSAFE_OBJECT_KEYS.has(key) || isPrivateSettingKey(key)) return;
        user.profile[key] = importedProfile[key];
    });
}

async function importUserSettings() {
    try {
        const { employeeId, user } = getCurrentSettingsUser();
        setSettingsCloudButtonsDisabled(true);

        const firestore = getFirebaseApp().firestore();
        const snapshot = await firestore.collection(USER_SETTINGS_COLLECTION).doc(employeeId).get();
        if (!snapshot.exists) {
            showNativeMsgBox('Firebase에 저장된 설정값이 없습니다. 먼저 설정 내보내기를 해주세요.', '설정 불러오기');
            return;
        }

        const stored = snapshot.data() || {};
        if (!stored.settings || typeof stored.settings !== 'object' || Array.isArray(stored.settings)) {
            throw new Error('저장된 설정 데이터 형식이 올바르지 않습니다.');
        }

        const importedSettings = buildExportableUserSettings(stored.settings);
        if (!await requestNativeConfirmation(
            formatSettingsExportPreview(employeeId, importedSettings),
            '설정 불러오기 확인',
            '위 내용으로 설정을 적용하시겠습니까?'
        )) return;

        applyImportedUserSettings(user, importedSettings);
        sendMessageToAHK({ command: 'saveConfig', data: appConfig });
        loadSettingsToUI();
        refreshUI();
        showNativeMsgBox('Firebase에서 사용자 설정을 불러와 적용했습니다.', '설정 불러오기');
    } catch (error) {
        console.error('Failed to import user settings:', error);
        showNativeMsgBox(`설정을 불러오지 못했습니다.\n${error?.message || String(error)}`, '설정 불러오기 오류');
    } finally {
        setSettingsCloudButtonsDisabled(false);
    }
}

window.exportUserSettings = exportUserSettings;
window.importUserSettings = importUserSettings;

// --- Settings Collection Logic (Vue) ---
function scrollSettingsToBottom() {
    Vue.nextTick(() => {
        const contentArea = document.querySelector('.settings-content-area');
        if (contentArea) contentArea.scrollTop = contentArea.scrollHeight;
    });
}

function addWorkerRow(data = {}) {
    settingsState.workers.push(normalizeWorker(Object.keys(data).length ? data : {}));
    autoSaveSettings();
    scrollSettingsToBottom();
}

function addLocationRow(data = {}) {
    settingsState.locations.push(normalizeLocation(Object.keys(data).length ? data : {}));
    autoSaveSettings();
    scrollSettingsToBottom();
}

const defaultHotkeys = [
    { action: 'AutoLogin', key: '#z', desc: '자동 로그인' },
    { action: 'OpenLog', key: '#!z', desc: '업무일지 실행' },
    { action: 'ConvertExcel', key: '#!a', desc: '일반업무 -> 엑셀 변환' },
    { action: 'CopyExcel', key: '#!c', desc: '엑셀 데이터 복사' },
    { action: 'PasteExcel', key: '#!v', desc: '일반업무에 붙여넣기' },
    { action: 'ForceExit', key: '^Esc', desc: '강제 종료' }
];

const hotkeyTranslations = {
    AutoLogin: '자동 로그인',
    OpenLog: '업무일지 실행',
    ConvertExcel: '엑셀 변환',
    CopyExcel: '엑셀 복사',
    PasteExcel: '붙여넣기',
    ForceExit: '강제 종료'
};

function showNativeMsgBox(text, title = "알림") {
    sendMessageToAHK({ command: 'msgbox', text: text, title: title });
}

function formatPhoneValue(rawValue) {
    let value = String(rawValue || '').replace(/[^0-9]/g, '');
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

    return formatted;
}

function formatPhone(input) {
    input.value = formatPhoneValue(input.value);

    // Auto Save is handled by global delegate on 'input' event
}

function formatTimeValue(rawValue) {
    let value = String(rawValue || '').replace(/[^0-9]/g, '');
    // Safety cut for HHmm (4 digits)
    if (value.length > 4) value = value.substr(0, 4);

    let formatted = '';
    if (value.length < 3) {
        formatted = value;
    } else {
        // HH:mm
        formatted = value.substr(0, 2) + ':' + value.substr(2);
    }
    return formatted;
}

function formatTime(input) {
    input.value = formatTimeValue(input.value);

    // Auto Save is handled by global delegate on 'input' event
}

// --- ERP Check Logic ---
const erpState = Vue.reactive({
    locations: [],
    selectedLocation: null,
    statusMap: null,
    completedOrders: null
});

function getERPWorkersForCurrentTeam() {
    const uid = selectedUserId;
    if (!uid || !appConfig.users?.[uid]) return [];
    const team = appConfig.users[uid].profile?.team || '';
    return (appConfig.appSettings?.colleagues || [])
        .filter(worker => team && worker.team === team)
        .sort((a, b) => {
            if (a.isManager !== b.isManager) return b.isManager - a.isManager;
            return a.id.localeCompare(b.id);
        });
}

function initERPIndividualApp() {
    Vue.createApp({
        data() {
            return { erp: erpState, controls: erpControlsState };
        },
        methods: {
            locationsForGroup(group) {
                return this.erp.locations.filter(loc => {
                    if (group === 'substation') return loc.type === '변전소';
                    if (loc.type === '기타업무') return group === 'etc';
                    if (!loc.type?.startsWith('전기실')) return false;
                    if (loc.type.includes('그룹1')) return group === 'group1';
                    if (loc.type.includes('그룹2')) return group === 'group2';
                    if (loc.type.includes('그룹3')) return group === 'group3';
                    return group === 'etc';
                });
            },
            selectLocation(name) {
                this.erp.selectedLocation = name;
            },
            isOrderCompleted(order) {
                return this.erp.completedOrders?.includes(String(order)) || false;
            },
            statusColor(name) {
                if (this.erp.statusMap === null) return '#FFC107';
                return this.erp.statusMap[name] ? '#4CAF50' : '#FF0000';
            },
            statusTitle(name) {
                if (this.erp.statusMap === null) return '데이터 확인 중...';
                return this.erp.statusMap[name] ? '점검 완료 (오늘)' : '점검 미완료';
            }
        }
    }).mount('#erp-individual-container');
}

function renderERPCheck() {
    const locations = appConfig.appSettings?.locations || [];
    erpState.locations = locations.map(loc => ({ ...loc }));
    if (erpState.selectedLocation && !locations.some(loc => loc.name === erpState.selectedLocation)) {
        erpState.selectedLocation = null;
    }
    updateERPBatchData();
}

function handleERPStatusUpdate(statusMap) {
    if (statusMap === null || statusMap === undefined) return;
    erpState.statusMap = { ...statusMap };
    if (erpBatchAppInstance) erpBatchAppInstance.latestStatusMap = { ...erpState.statusMap };
}

window.handleERPStatusUpdate = handleERPStatusUpdate;

function handleERPOrderListUpdate(orders) {
    if (!Array.isArray(orders)) return;
    erpState.completedOrders = orders.map(String);

    // 1. Update Timestamp
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const dd = String(now.getDate()).padStart(2, '0');
    const hh = String(now.getHours()).padStart(2, '0');
    const min = String(now.getMinutes()).padStart(2, '0');
    const timeStr = `${yyyy}-${mm}-${dd} ${hh}:${min}`;

    erpControlsState.lastUpdate = timeStr;

    if (erpBatchAppInstance) erpBatchAppInstance.completedOrders = [...erpState.completedOrders];
}

let erpBatchAppInstance = null;
let isBatchMode = false;

function setERPMode(useBatchMode) {
    isBatchMode = !!useBatchMode;
    erpControlsState.isBatchMode = isBatchMode;
    if (useBatchMode) {
        if (!erpBatchAppInstance) {
            initERPBatchApp();
        } else {
            updateERPBatchData();
        }
    }
}

function toggleERPMode() {
    setERPMode(!isBatchMode);
}

function requestERPRefresh() {
    if (!erpControlsState.sessionReady) return;
    erpControlsState.lastUpdate = '조회 중...';
    sendMessageToAHK({ command: 'refreshERPOrder' });
}

function currentLocalDateKey(date = new Date()) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
}

function collectERPRecoveryState() {
    const batchState = erpBatchAppInstance ? {
        activeRows: [...erpBatchAppInstance.activeRows],
        selections: Object.fromEntries(
            erpBatchAppInstance.activeRows.map(name => [name, [...(erpBatchAppInstance.selections[name] || [])]])
        )
    } : { activeRows: [], selections: {} };

    return {
        savedDate: currentLocalDateKey(),
        mode: isBatchMode ? 'batch' : 'individual',
        format: erpControlsState.isListFormat ? 'list' : 'summary',
        selectedLocation: erpState.selectedLocation,
        completedOrders: Array.isArray(erpState.completedOrders) ? [...erpState.completedOrders] : null,
        batch: batchState
    };
}

function restoreERPRecoveryState(savedState) {
    if (!savedState || typeof savedState !== 'object') return;

    const validLocations = new Set(erpState.locations.map(location => location.name));
    erpState.selectedLocation = validLocations.has(savedState.selectedLocation)
        ? savedState.selectedLocation
        : null;

    // 완료 오더는 날짜가 바뀌면 이전 날짜의 체크를 표시하지 않습니다.
    if (savedState.savedDate === currentLocalDateKey() && Array.isArray(savedState.completedOrders)) {
        erpState.completedOrders = savedState.completedOrders.map(String);
    }
    if (savedState.format === 'list' || savedState.format === 'summary') {
        erpControlsState.isListFormat = savedState.format === 'list';
    }

    setERPMode(savedState.mode === 'batch');
    if (!erpBatchAppInstance) return;

    const savedBatch = savedState.batch || {};
    const activeRows = Array.isArray(savedBatch.activeRows)
        ? savedBatch.activeRows.filter(name => validLocations.has(name))
        : [];
    const workerIds = new Map(erpBatchAppInstance.workers.map(worker => [String(worker.id), worker.id]));
    const selections = {};
    activeRows.forEach(name => {
        const savedWorkerIds = Array.isArray(savedBatch.selections?.[name])
            ? savedBatch.selections[name]
            : [];
        selections[name] = savedWorkerIds
            .map(id => workerIds.get(String(id)))
            .filter(id => id !== undefined);
    });

    erpBatchAppInstance.activeRows = activeRows;
    erpBatchAppInstance.selections = selections;
    if (Array.isArray(erpState.completedOrders)) {
        erpBatchAppInstance.completedOrders = [...erpState.completedOrders];
    }
}

function initERPBatchApp() {
    const { createApp } = Vue;
    erpBatchAppInstance = createApp({
        data() {
            return {
                controls: erpControlsState,
                locations: [],
                workers: [],
                activeRows: [], // Array of location names
                selections: {},  // { locName: [workerId1, workerId2, ...] }
                latestStatusMap: null,
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
                        locs.push({ name: loc.name, type: shortType, order: String(loc.order ?? '') });
                    });
                }
                this.locations = locs;

                // 로그인 유저 목록 또는 설정된 분소원 목록에서 워커 데이터 추출
                this.workers = getERPWorkersForCurrentTeam().map(w => ({ id: w.id, name: w.name }));

                // Restore results received before this Vue app was mounted.
                if (erpState.statusMap !== null) {
                    this.latestStatusMap = { ...erpState.statusMap };
                }
                if (erpState.completedOrders !== null) {
                    this.completedOrders = [...erpState.completedOrders];
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
                isOpen: false,
                batchDataRaw: [],
                batchPreviewList: [],
                isListFormat: false,
                previewMaxCount: 0
            };
        },
        methods: {
            openModal(rawData) {
                this.batchDataRaw = rawData;
                this.isListFormat = erpControlsState.isListFormat;
                this.updatePreviewList();
                this.isOpen = true;
            },
            closeModal() {
                this.isOpen = false;
            },
            toggleFormat() {
                this.isListFormat = !this.isListFormat;
                erpControlsState.isListFormat = this.isListFormat;
                saveERPFormatPreference();
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



// --- Work Log Logic (Vue) ---
const workLogState = Vue.reactive({
    workType: 'day',
    options: {
        makeLog: true,
        general: true,
        safe: true,
        driving: false,
        drink: false,
        drinkCalibration: false
    },
    safety: [
        { content: '', date: '', start: '', end: '' },
        { content: '', date: '', start: '', end: '' }
    ],
    workers: [],
    guntaeStatus: 'idle',
    ownerId: null
});

function collectWorkLogRecoveryState() {
    return {
        workType: workLogState.workType,
        options: { ...workLogState.options },
        safety: workLogState.safety.map(row => ({ ...row })),
        workers: workLogState.workers.map(worker => ({
            id: worker.id,
            attend: !!worker.attend,
            reason: worker.reason || '',
            driverRole: worker.driverRole || '',
            wasChecked: worker.wasChecked !== false
        })),
        guntaeStatus: workLogState.guntaeStatus
    };
}

function restoreWorkLogRecoveryState(savedState) {
    if (!savedState || typeof savedState !== 'object') return;
    renderWorkLogWorkerList();

    workLogState.workType = savedState.workType === 'night' ? 'night' : 'day';
    if (savedState.options) {
        Object.keys(workLogState.options).forEach(key => {
            if (savedState.options[key] !== undefined) {
                workLogState.options[key] = !!savedState.options[key];
            }
        });
    }
    if (Array.isArray(savedState.safety)) {
        workLogState.safety.forEach((row, index) => {
            const saved = savedState.safety[index] || {};
            row.content = String(saved.content || '');
            row.date = String(saved.date || '');
            row.start = String(saved.start || '');
            row.end = String(saved.end || '');
        });
    }

    const savedWorkers = new Map(
        (Array.isArray(savedState.workers) ? savedState.workers : [])
            .map(worker => [String(worker.id), worker])
    );
    workLogState.workers.forEach(worker => {
        const saved = savedWorkers.get(String(worker.id));
        if (!saved) return;
        worker.attend = !!saved.attend;
        worker.reason = String(saved.reason || '');
        worker.driverRole = String(saved.driverRole || '');
        worker.wasChecked = saved.wasChecked !== false;
    });
    updateGuntaeStatusLabel(savedState.guntaeStatus || 'idle');
}

function initWorkLogApp() {
    Vue.createApp({
        data() {
            return { work: workLogState };
        },
        methods: {
            changeWorkType() {
                handleWorkTypeChange();
            },
            toggleDrink() {
                toggleDrinkCalibration();
            },
            workerNoteChanged(worker) {
                if (worker.reason.trim() !== '' && worker.reason !== '일근') {
                    worker.attend = false;
                    worker.driverRole = '';
                }
                updateWorkerStats();
            },
            workerAttendanceChanged(worker) {
                if (!worker.attend) worker.driverRole = '';
                updateWorkerStats();
            },
            formatSafetyTime(index, field, event) {
                this.work.safety[index][field] = formatTimeValue(event.target.value);
            },
            startTask: startWorkLog
        }
    }).mount('#view-work-log');
}

function workLogDateText(date) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    return `${yyyy}${mm}${dd}`;
}

function applyWorkLogSafetyDefaults() {
    const user = appConfig.users?.[selectedUserId];
    const safetyPresets = user?.dailyLogDefaults?.safety || [];
    const now = new Date();
    const tomorrow = new Date(now);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const indexes = workLogState.workType === 'day' ? [0, 1] : [2, 3];
    const dates = workLogState.workType === 'day'
        ? [workLogDateText(now), workLogDateText(now)]
        : [workLogDateText(now), workLogDateText(tomorrow)];

    workLogState.safety.forEach((row, index) => {
        const preset = safetyPresets[indexes[index]] || {};
        row.content = String(preset.content || '');
        row.start = String(preset.start || '');
        row.end = String(preset.end || '');
        row.date = dates[index];
    });
}

function renderWorkLogUI() {
    const uid = selectedUserId;
    if (!uid || !appConfig.users?.[uid]) return;
    if (workLogState.ownerId !== uid) {
        Object.assign(workLogState.options, {
            makeLog: true,
            general: true,
            safe: true,
            driving: false,
            drink: false,
            drinkCalibration: false
        });
        workLogState.workers = [];
        updateGuntaeStatusLabel('idle');
    }
    renderWorkLogWorkerList();
    handleWorkTypeChange(false);
    isWorkLogInitialized = true;
}

function handleWorkTypeChange() {
    const isDay = workLogState.workType === 'day';
    if (isDay) {
        workLogState.options.makeLog = true;
        workLogState.options.driving = false;
        if (!workLogState.options.drink) workLogState.options.drinkCalibration = false;
    } else {
        workLogState.options.makeLog = false;
        workLogState.options.driving = true;
        workLogState.options.drink = false;
        workLogState.options.drinkCalibration = false;
    }

    updateDayShiftWorkers(isDay);
    applyWorkLogSafetyDefaults();
}

function toggleDrinkCalibration() {
    if (workLogState.workType === 'night' || !workLogState.options.drink) {
        workLogState.options.drinkCalibration = false;
    }
}

function parseGuntaeStatus(entry) {
    if (!entry) return { isAttending: false, reason: '' };
    const worktime = String(entry.WORKTIME || '').trim();
    const guntaeName = entry.GUNTAENAME || '-';
    if (guntaeName !== '-') return { isAttending: false, reason: guntaeName };
    if (worktime !== '') return { isAttending: true, reason: '' };
    return { isAttending: false, reason: '' };
}

function renderWorkLogWorkerList() {
    const uid = selectedUserId;
    const user = appConfig.users?.[uid];
    if (!uid || !user) return;

    const ownerChanged = workLogState.ownerId !== uid;
    const existing = ownerChanged
        ? new Map()
        : new Map(workLogState.workers.map(worker => [String(worker.id), worker]));
    const myTeam = user.profile?.team || '';
    const workers = (appConfig.appSettings?.colleagues || [])
        .filter(worker => worker.team === '일근' || (myTeam && worker.team === myTeam))
        .sort((a, b) => {
            if (a.isManager !== b.isManager) return b.isManager - a.isManager;
            return String(a.id).localeCompare(String(b.id));
        })
        .map(worker => {
            const saved = existing.get(String(worker.id));
            if (saved) {
                return {
                    ...saved,
                    id: worker.id,
                    name: worker.name,
                    team: worker.team,
                    isManager: worker.isManager
                };
            }

            let attend = true;
            let reason = '';
            const entry = Array.isArray(guntaeData)
                ? guntaeData.find(item => String(item.SABUN) === String(worker.id))
                : null;
            if (entry) {
                const status = parseGuntaeStatus(entry);
                attend = status.isAttending;
                reason = status.reason;
            }
            return {
                id: worker.id,
                name: worker.name,
                team: worker.team,
                isManager: worker.isManager,
                attend,
                reason,
                driverRole: worker.driverRole === '-' ? '' : String(worker.driverRole || ''),
                wasChecked: attend
            };
        });

    workLogState.workers = workers;
    workLogState.ownerId = uid;
    updateGuntaeStatusLabel(guntaeStatus);
}

function updateGuntaeStatusLabel(status) {
    if (status === true) status = 'success';
    if (status === false) status = 'idle';
    guntaeStatus = status || 'idle';
    workLogState.guntaeStatus = guntaeStatus;
}

function applyGuntaeData(data) {
    if (!Array.isArray(data)) return;
    const isDay = workLogState.workType === 'day';
    workLogState.workers.forEach(worker => {
        const entry = data.find(item => String(item.SABUN) === String(worker.id));
        if (!entry) return;
        if (!isDay && worker.team === '일근') {
            worker.attend = false;
            worker.reason = '';
            worker.driverRole = '';
        } else {
            const status = parseGuntaeStatus(entry);
            worker.attend = status.isAttending;
            worker.reason = status.reason;
            if (!worker.attend) worker.driverRole = '';
        }
    });
    updateGuntaeStatusLabel(data.length > 0 ? 'success' : 'failure');
    updateWorkerStats();
}

function enableDriverSelects() {
    // Vue 템플릿의 disabled 바인딩이 주야간 상태를 직접 반영합니다.
}

function toggleAllWorkers(mainChk) {
    const checked = !!mainChk?.checked;
    workLogState.workers.forEach(worker => {
        worker.attend = checked;
        if (!checked) worker.driverRole = '';
    });
    updateWorkerStats();
}

function updateWorkerStats() {
    // 현재 화면에는 별도 통계 표시가 없습니다.
}

function startWorkLog() {
    if (!selectedUserId) return;
    const safetyData = {
        content1: workLogState.safety[0].content,
        date1: workLogState.safety[0].date,
        start1: workLogState.safety[0].start,
        end1: workLogState.safety[0].end,
        content2: workLogState.safety[1].content,
        date2: workLogState.safety[1].date,
        start2: workLogState.safety[1].start,
        end2: workLogState.safety[1].end
    };
    const workers = workLogState.workers.map(worker => ({
        name: worker.name,
        id: worker.id,
        attend: worker.attend,
        reason: worker.reason,
        driverRole: worker.driverRole,
        team: worker.team,
        boss: worker.isManager
    }));
    sendMessageToAHK({
        command: 'runTask',
        task: 'createWorkLog',
        data: {
            workType: workLogState.workType,
            options: {
                makeLog: workLogState.options.makeLog,
                general: workLogState.options.general,
                safe: workLogState.options.safe,
                driving: workLogState.options.driving,
                drink: workLogState.options.drink
            },
            chkDrinkDetectorCalibration: workLogState.options.drinkCalibration,
            safetyData,
            safetyEdu: currentSafetyEduData || {},
            workers
        }
    });
}

// Global Exports
window.switchMainTab = switchMainTab;
window.runTask = function (task) { sendMessageToAHK({ command: 'runTask', task: task }); };

window.openSettings = openSettings;
window.closeSettings = closeSettings;
window.switchSettingsTab = switchSettingsTab;
window.tryLogin = tryLogin;
window.deleteUser = deleteUser;
window.switchView = switchView;
window.submitNewUser = submitNewUser;
window.saveSettings = saveSettings;
window.minimizeWindow = function () { sendMessageToAHK({ command: 'minimize' }); };
window.closeWindow = function () { sendMessageToAHK({ command: 'close' }); };

// --- Shift Status UI Update ---
function updateShiftUI(shiftData) {
    Object.assign(shellState.shift, {
        visible: true,
        prev: String(shiftData.prev || ''),
        current: String(shiftData.current || ''),
        next: String(shiftData.next || ''),
        prevDate: shiftData.isNight ? String(shiftData.prevDate || '') : '',
        currentDate: shiftData.isNight ? '' : String(shiftData.currDate || ''),
        nextDate: shiftData.isNight ? String(shiftData.nextDate || '') : '',
        name: String(shiftData.shiftName || ''),
        isNight: !!shiftData.isNight
    });

    const nextWorkType = shiftData.isNight ? 'night' : 'day';
    if (workLogState.workType !== nextWorkType) {
        workLogState.workType = nextWorkType;
        renderWorkLogWorkerList();
        handleWorkTypeChange();
    }
}


function updateDayShiftWorkers(isDay) {
    workLogState.workers.forEach(worker => {
        if (worker.team !== '일근') return;
        if (!isDay) {
            worker.wasChecked = worker.attend;
            worker.attend = false;
            worker.reason = '';
            worker.driverRole = '';
            return;
        }

        const entry = Array.isArray(guntaeData)
            ? guntaeData.find(item => String(item.SABUN) === String(worker.id))
            : null;
        if (entry) {
            const status = parseGuntaeStatus(entry);
            worker.attend = status.isAttending;
            worker.reason = status.reason;
        } else {
            worker.attend = worker.wasChecked !== false;
        }
    });
    updateWorkerStats();
}


// Work Log Exports
// --- ERP Worker Modal Logic ---
let erpWorkerModalAppInstance = null;

function initERPWorkerModalApp() {
    erpWorkerModalAppInstance = Vue.createApp({
        data() {
            return {
                isOpen: false,
                workers: [],
                selectedWorkerIds: [],
                isListMode: false
            };
        },
        computed: {
            selectedWorkers() {
                return this.workers.filter(worker => this.selectedWorkerIds.includes(worker.id));
            },
            buttonText() {
                const names = this.selectedWorkers.map(worker => worker.name);
                if (names.length === 0) return '(선택 없음)';
                if (this.isListMode) return names.join(', ');
                return names.length === 1 ? names[0] : `${names[0]} 외 ${names.length - 1}명`;
            }
        },
        methods: {
            openModal() {
                if (!erpState.selectedLocation) {
                    showNativeMsgBox('점검 장소를 선택해주세요.');
                    return;
                }
                this.workers = getERPWorkersForCurrentTeam();
                this.selectedWorkerIds = this.workers.map(worker => worker.id);
                this.isListMode = erpControlsState.isListFormat;
                this.isOpen = true;
            },
            closeModal() {
                this.isOpen = false;
            },
            toggleFormat() {
                this.isListMode = !this.isListMode;
                erpControlsState.isListFormat = this.isListMode;
                saveERPFormatPreference();
            },
            submitTask() {
                const members = this.selectedWorkers.map(worker => worker.name);
                if (members.length === 0) {
                    showNativeMsgBox('작업자를 한 명 이상 선택해주세요.');
                    return;
                }

                const location = erpState.locations.find(loc => loc.name === erpState.selectedLocation);
                if (!location) {
                    showNativeMsgBox('점검 장소를 다시 선택해주세요.');
                    return;
                }

                sendMessageToAHK({
                    command: 'runTask',
                    task: 'ERPCheck',
                    location: location.name,
                    targetType: location.type || '',
                    targetOrder: location.order || '',
                    members,
                    format: this.isListMode ? 'list' : 'summary'
                });
                this.closeModal();
            }
        }
    }).mount('#erp-worker-modal');
}

function openERPWorkerModal() {
    erpWorkerModalAppInstance.openModal();
}

function updateToggleStyle() {
    if (erpWorkerModalAppInstance?.isOpen) {
        erpWorkerModalAppInstance.isListMode = erpControlsState.isListFormat;
    }
}

function saveERPFormatPreference() {
    updateToggleStyle();
    const format = erpControlsState.isListFormat ? 'list' : 'summary';
    const uid = selectedUserId;
    if (!uid || !appConfig.users?.[uid]) return;
    if (!appConfig.users[uid].profile) appConfig.users[uid].profile = {};
    appConfig.users[uid].profile.erpFormat = format;
    if (settingsState.loadedUserId === uid) settingsState.profile.erpFormat = format;
    sendMessageToAHK({ command: 'saveConfig', data: appConfig });
}

// Global Exports
window.openERPWorkerModal = openERPWorkerModal;
window.handleWorkTypeChange = handleWorkTypeChange;
window.toggleAllWorkers = toggleAllWorkers;
window.updateToggleStyle = updateToggleStyle;
window.startWorkLog = startWorkLog;
window.toggleDrinkCalibration = toggleDrinkCalibration;

// --- Helper: Preset Management ---
function getUserPresetSource(type) {
    const uid = String(selectedUserId || '').trim();
    if (!uid || !appConfig.users || !appConfig.users[uid]) return null;

    const user = appConfig.users[uid];
    const presets = type === 'track'
        ? user.trackPresets
        : type === 'vehicle'
            ? user.vehiclePresets
            : null;

    return presets && typeof presets === 'object' && !Array.isArray(presets) ? presets : {};
}

function getUserPresets(type) {
    return getUserPresetSource(type) || {};
}

function saveUserPresets(type, presets) {
    const uid = selectedUserId;
    if (!uid) return;
    if (!appConfig.users[uid]) return;

    if (type === 'track') appConfig.users[uid].trackPresets = presets;
    if (type === 'vehicle') appConfig.users[uid].vehiclePresets = presets;

    sendMessageToAHK({ command: 'saveConfig', data: appConfig });
}

// --- Track Access Logic (Vue) ---
const TRACK_ACCESS_DEFAULTS = Object.freeze({
    workType: '1', content: '', workFrom: '', workTo: '',
    driverName: '', driverPhone: '', workerName: '', workerPhone: '',
    safetyName: '', safetyPhone: '', supervisorName: '', supervisorId: '',
    workStart: '', workEnd: '', opStart: '', opEnd: '',
    line: '1', trackType: '1', trackCutoff: false,
    agreementNo: '', totalCount: '', stationInput: false
});

const trackAccessState = Vue.reactive({
    presets: {},
    presetRenderKey: 0,
    selectedPreset: '__NEW__',
    form: { ...TRACK_ACCESS_DEFAULTS },
    workerPicker: { isOpen: false, role: '', selectedIndex: -1 },
    workerListRevision: 0,
    agreementLoading: false,
    agreementLookupFresh: false,
    renewalConfirmPending: false,
    hasSyncedPresets: false,
    ownerId: null
});
let trackAccessAppInstance = null;

function trackBoolean(value) {
    return value === true || value === 1 || value === '1' || value === 'Y' || value === 'true';
}

function normalizeTrackPreset(data = {}) {
    const normalized = {};
    Object.keys(TRACK_ACCESS_DEFAULTS).forEach(key => {
        if (key === 'trackCutoff' || key === 'stationInput') {
            normalized[key] = trackBoolean(data[key]);
        } else {
            const fallback = TRACK_ACCESS_DEFAULTS[key];
            normalized[key] = data[key] === undefined || data[key] === null ? fallback : String(data[key]);
        }
    });
    return normalized;
}

function clearTrackAccessForm() {
    Object.assign(trackAccessState.form, TRACK_ACCESS_DEFAULTS);
    trackAccessState.agreementLookupFresh = false;
}

function collectTrackAccessRecoveryState() {
    return {
        selectedPreset: trackAccessState.selectedPreset,
        form: { ...normalizeTrackPreset(trackAccessState.form) }
    };
}

function restoreTrackAccessRecoveryState(savedState) {
    if (!savedState?.form || typeof savedState.form !== 'object') return;

    const savedPreset = savedState.selectedPreset;
    trackAccessState.selectedPreset = savedPreset === '__NEW__' || trackAccessState.presets[savedPreset]
        ? savedPreset
        : '__NEW__';
    Object.assign(trackAccessState.form, normalizeTrackPreset(savedState.form));
}

function initTrackAccessApp() {
    trackAccessAppInstance = Vue.createApp({
        data() {
            return { track: trackAccessState };
        },
        computed: {
            availableWorkers() {
                // appConfig 자체는 Vue 반응형 객체가 아니므로 이 값을 의존성으로 사용한다.
                void this.track.workerListRevision;
                return (appConfig.appSettings?.colleagues || [])
                    .filter(person => String(person?.name || '').trim());
            },
            workerPickerTitle() {
                const labels = {
                    driver: '운전원 선택', worker: '작업자 선택',
                    safety: '철도운행안전관리자 선택', supervisor: '감독자 선택'
                };
                return labels[this.track.workerPicker.role] || '작업자 선택';
            }
        },
        methods: {
            loadPreset: loadTrackPreset,
            savePreset: saveTrackPreset,
            renamePreset: renameTrackPreset,
            deletePreset: deleteTrackPreset,
            runTask: runTrackAccessTask,
            loadAgreement: requestTrackAgreement,
            openWorkerPicker(role) {
                const roleFields = {
                    driver: 'driverName', worker: 'workerName',
                    safety: 'safetyName', supervisor: 'supervisorName'
                };
                const currentName = String(this.track.form[roleFields[role]] || '').trim();
                this.track.workerPicker.role = role;
                this.track.workerPicker.selectedIndex = this.availableWorkers
                    .findIndex(person => String(person.name || '').trim() === currentName);
                this.track.workerPicker.isOpen = true;
            },
            closeWorkerPicker() {
                this.track.workerPicker.isOpen = false;
                this.track.workerPicker.selectedIndex = -1;
            },
            applyWorkerSelection(index) {
                const selectedIndex = Number.isInteger(index) ? index : this.track.workerPicker.selectedIndex;
                const person = this.availableWorkers[selectedIndex];
                if (!person) return;

                const mappings = {
                    driver: ['driverName', 'driverPhone', 'phone'],
                    worker: ['workerName', 'workerPhone', 'phone'],
                    safety: ['safetyName', 'safetyPhone', 'phone'],
                    supervisor: ['supervisorName', 'supervisorId', 'id']
                };
                const mapping = mappings[this.track.workerPicker.role];
                if (!mapping) return;
                this.track.form[mapping[0]] = String(person.name || '').trim();
                this.track.form[mapping[1]] = mapping[2] === 'phone'
                    ? formatPhoneValue(String(person.phone || ''))
                    : String(person.id || '').replace(/\D/g, '').slice(0, 6);
                this.closeWorkerPicker();
            },
            formatField(field, event, type) {
                const formatter = type === 'phone' ? formatPhoneValue : formatTimeValue;
                this.track.form[field] = formatter(event.target.value);
            },
            digitsOnly(field, event) {
                this.track.form[field] = event.target.value.replace(/[^0-9]/g, '');
            }
        }
    }).mount('#view-track-access');
}

function syncTrackPresets(forceLoad = false) {
    const presets = getUserPresetSource('track');
    // 로그인 사용자 또는 설정이 아직 준비되지 않은 순간에는 기존 목록을 지우지 않는다.
    if (presets === null) return;

    const keys = Object.keys(presets);
    const current = trackAccessState.selectedPreset;
    const ownerId = String(selectedUserId || '').trim();
    const ownerChanged = trackAccessState.ownerId !== ownerId;
    trackAccessState.presets = { ...presets };
    // 과거 코드나 브라우저 복원 과정에서 option DOM이 직접 변경됐더라도 확실히 재구성한다.
    trackAccessState.presetRenderKey += 1;

    let next = '__NEW__';
    if (keys.length > 0) next = presets[current] ? current : keys[0];
    const shouldLoad = forceLoad || ownerChanged || !trackAccessState.hasSyncedPresets || next !== current;
    trackAccessState.selectedPreset = next;
    trackAccessState.hasSyncedPresets = true;
    trackAccessState.ownerId = ownerId;
    if (shouldLoad) loadTrackPreset();
}

function loadTrackPreset() {
    trackAccessState.agreementLookupFresh = false;
    const key = trackAccessState.selectedPreset;
    if (!key || key === '__NEW__') {
        clearTrackAccessForm();
        return;
    }
    const data = trackAccessState.presets[key] || getUserPresets('track')[key];
    if (data) Object.assign(trackAccessState.form, normalizeTrackPreset(data));
}

function getTrackPresetData() {
    return {
        ...normalizeTrackPreset(trackAccessState.form),
        agreementNoDate: new Date().toISOString().slice(0, 10)
    };
}

function saveTrackPreset() {
    let key = trackAccessState.selectedPreset;
    if (!key || key === '__NEW__') {
        key = (prompt('새 프리셋 이름을 입력하세요:') || '').trim();
        if (!key) return;
    }

    const presets = getUserPresets('track');
    presets[key] = getTrackPresetData();
    trackAccessState.selectedPreset = key;
    saveUserPresets('track', presets);
    syncTrackPresets();
    showNativeMsgBox(`'${key}' 프리셋이 저장되었습니다.`);
}

function renameTrackPreset() {
    const oldKey = trackAccessState.selectedPreset;
    if (!oldKey || oldKey === '__NEW__') {
        showNativeMsgBox('이름을 변경할 프리셋을 선택해주세요.');
        return;
    }

    const newKey = (prompt('새 이름을 입력하세요:', oldKey) || '').trim();
    if (!newKey || newKey === oldKey) return;

    const presets = getUserPresets('track');
    if (presets[newKey]) {
        showNativeMsgBox('이미 존재하는 이름입니다.');
        return;
    }

    presets[newKey] = presets[oldKey];
    delete presets[oldKey];
    trackAccessState.selectedPreset = newKey;
    saveUserPresets('track', presets);
    syncTrackPresets();
}

function deleteTrackPreset() {
    const key = trackAccessState.selectedPreset;
    if (!key || key === '__NEW__') {
        showNativeMsgBox('삭제할 프리셋을 선택해주세요.');
        return;
    }
    if (!confirm(`'${key}' 프리셋을 삭제하시겠습니까?`)) return;

    const presets = getUserPresets('track');
    delete presets[key];
    trackAccessState.selectedPreset = '__NEW__';
    saveUserPresets('track', presets);
    syncTrackPresets(true);
}

function trackAgreementNeedsRenewal(today = new Date()) {
    const day = today.getDate();
    const lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
    if (day > 15 && day !== lastDay) return false;

    const key = trackAccessState.selectedPreset;
    const preset = key && key !== '__NEW__' ? getUserPresets('track')[key] : null;
    if (!preset) return false;
    if (!preset.agreementNoDate) return !!preset.agreementNo;

    const diff = (today - new Date(preset.agreementNoDate)) / (1000 * 60 * 60 * 24);
    return diff > 15;
}

function runTrackAccessTask() {
    const needsRenewal = !trackAccessState.agreementLookupFresh && trackAgreementNeedsRenewal();
    if (needsRenewal) {
        if (trackAccessState.renewalConfirmPending) return;
        trackAccessState.renewalConfirmPending = true;
        sendMessageToAHK({ command: 'confirmTrackAgreementRenewal' });
        return;
    }

    dispatchTrackAccessTask();
}

function dispatchTrackAccessTask() {
    const data = {
        ...normalizeTrackPreset(trackAccessState.form),
        agreementNo: trackAccessState.form.agreementNo,
        needsRenewal: false
    };
    sendMessageToAHK({ command: 'runTask', task: 'TrackAccess', data });
}

function requestTrackAgreement() {
    if (trackAccessState.agreementLoading) return;

    const participants = [
        trackAccessState.form.driverName,
        trackAccessState.form.workerName,
        trackAccessState.form.safetyName,
        trackAccessState.form.supervisorName
    ].map(name => String(name || '').trim()).filter(Boolean);

    if (participants.length === 0) {
        showNativeMsgBox('운전원, 작업자, 철도운행안전관리자 또는 감독자를 먼저 입력해주세요.');
        return;
    }

    trackAccessState.agreementLoading = true;
    sendMessageToAHK({ command: 'loadTrackAgreement', participants });
}

function applyTrackAgreementNo(value) {
    if (!value) return;
    trackAccessState.form.agreementNo = value;
    trackAccessState.agreementLookupFresh = true;
}


// --- Vehicle Log Logic (Vue) ---
const VEHICLE_LOG_DEFAULTS = Object.freeze({
    driver: '', point1: '', point2: '', trackType: '상행선',
    startTime: '', endTime: '', content: '', remarks: '',
    approveNo: '', dept: '', approver: '', runTime: '', distance: ''
});

const vehicleLogState = Vue.reactive({
    presets: {},
    selectedPreset: '__NEW__',
    form: { ...VEHICLE_LOG_DEFAULTS },
    hasSyncedPresets: false,
    ownerId: null,
    sessionReady: false
});

function collectVehicleLogRecoveryState() {
    return {
        selectedPreset: vehicleLogState.selectedPreset,
        form: { ...normalizeVehiclePreset(vehicleLogState.form) }
    };
}

function restoreVehicleLogRecoveryState(savedState) {
    if (!savedState?.form || typeof savedState.form !== 'object') return;
    const savedPreset = savedState.selectedPreset;
    vehicleLogState.selectedPreset = savedPreset === '__NEW__' || vehicleLogState.presets[savedPreset]
        ? savedPreset
        : '__NEW__';
    Object.assign(vehicleLogState.form, normalizeVehiclePreset(savedState.form));
}

function normalizeVehiclePreset(data = {}) {
    const normalized = {};
    Object.keys(VEHICLE_LOG_DEFAULTS).forEach(key => {
        normalized[key] = data[key] === undefined || data[key] === null
            ? VEHICLE_LOG_DEFAULTS[key]
            : String(data[key]);
    });
    return normalized;
}

function clearVehicleLogForm() {
    Object.assign(vehicleLogState.form, VEHICLE_LOG_DEFAULTS);
}

function initVehicleLogApp() {
    Vue.createApp({
        data() {
            return { vehicle: vehicleLogState };
        },
        methods: {
            loadPreset: loadVehiclePreset,
            savePreset: saveVehiclePreset,
            renamePreset: renameVehiclePreset,
            deletePreset: deleteVehiclePreset,
            runTask: runVehicleLogTask,
            loadApproval: runBringApproved,
            formatTimeField(field, event) {
                this.vehicle.form[field] = formatTimeValue(event.target.value);
            },
            digitsOnly(field, event) {
                this.vehicle.form[field] = event.target.value.replace(/[^0-9]/g, '');
            }
        }
    }).mount('#view-vehicle-log');
}

function syncVehiclePresets(forceLoad = false) {
    const presets = getUserPresets('vehicle');
    const keys = Object.keys(presets);
    const current = vehicleLogState.selectedPreset;
    const ownerChanged = vehicleLogState.ownerId !== selectedUserId;
    vehicleLogState.presets = { ...presets };

    let next = '__NEW__';
    if (keys.length > 0) next = presets[current] ? current : keys[0];
    const shouldLoad = forceLoad || ownerChanged || !vehicleLogState.hasSyncedPresets || next !== current;
    vehicleLogState.selectedPreset = next;
    vehicleLogState.hasSyncedPresets = true;
    vehicleLogState.ownerId = selectedUserId;
    if (shouldLoad) loadVehiclePreset();
}

function loadVehiclePreset() {
    const key = vehicleLogState.selectedPreset;
    if (!key || key === '__NEW__') {
        clearVehicleLogForm();
        return;
    }
    const data = vehicleLogState.presets[key] || getUserPresets('vehicle')[key];
    if (data) Object.assign(vehicleLogState.form, normalizeVehiclePreset(data));
}

function saveVehiclePreset() {
    let key = vehicleLogState.selectedPreset;
    if (!key || key === '__NEW__') {
        key = (prompt('새 프리셋 이름을 입력하세요:') || '').trim();
        if (!key) return;
    }

    const presets = getUserPresets('vehicle');
    presets[key] = normalizeVehiclePreset(vehicleLogState.form);
    vehicleLogState.selectedPreset = key;
    saveUserPresets('vehicle', presets);
    syncVehiclePresets();
    showNativeMsgBox(`'${key}' 프리셋이 저장되었습니다.`);
}

function renameVehiclePreset() {
    const oldKey = vehicleLogState.selectedPreset;
    if (!oldKey || oldKey === '__NEW__') {
        showNativeMsgBox('이름을 변경할 프리셋을 선택해주세요.');
        return;
    }
    const newKey = (prompt('새 이름을 입력하세요:', oldKey) || '').trim();
    if (!newKey || newKey === oldKey) return;

    const presets = getUserPresets('vehicle');
    if (presets[newKey]) {
        showNativeMsgBox('이미 존재하는 이름입니다.');
        return;
    }
    presets[newKey] = presets[oldKey];
    delete presets[oldKey];
    vehicleLogState.selectedPreset = newKey;
    saveUserPresets('vehicle', presets);
    syncVehiclePresets();
}

function deleteVehiclePreset() {
    const key = vehicleLogState.selectedPreset;
    if (!key || key === '__NEW__') {
        showNativeMsgBox('삭제할 프리셋을 선택해주세요.');
        return;
    }
    if (!confirm(`'${key}' 프리셋을 삭제하시겠습니까?`)) return;

    const presets = getUserPresets('vehicle');
    delete presets[key];
    vehicleLogState.selectedPreset = '__NEW__';
    saveUserPresets('vehicle', presets);
    syncVehiclePresets(true);
}

function runVehicleLogTask() {
    const data = {
        ...normalizeVehiclePreset(vehicleLogState.form),
        department: appConfig.users[selectedUserId]?.profile?.department || ''
    };
    sendMessageToAHK({ command: 'runTask', task: 'VehicleLog', data });
}

function runBringApproved() {
    const driver = vehicleLogState.form.driver.trim();
    if (!driver) {
        showNativeMsgBox('운전자를 입력해주세요.');
        return;
    }
    sendMessageToAHK({
        command: 'runTask',
        task: 'bringApproval',
        data: { driverName: driver }
    });
}

// Helper: Init Presets after login
function initPresets() {
    syncTrackPresets();
    syncVehiclePresets();
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

function addGeneralWorkRow(data = {}) {
    const row = Object.keys(data).length ? data : {
        workType: settingsState.generalWorkShift,
        category: '전체', content: '', manager: '', start: '', end: ''
    };
    settingsState.dailyLogDefaults.generalWork.push(normalizeGeneralWork(row));
    autoSaveSettings();
    scrollSettingsToBottom();
}

function toggleGeneralWorkShift() {
    settingsState.generalWorkShift = settingsState.generalWorkShift === '주간' ? '야간' : '주간';
}

function saveAutoInputSettings() {
    flushSettingsSave();
    showNativeMsgBox(`자동입력 예약시간(${settingsState.dailyLogDefaults.autoInputTime})이 저장되었습니다.`);
}

// Global Exports
window.addGeneralWorkRow = addGeneralWorkRow;
window.saveAutoInputSettings = saveAutoInputSettings;
window.formatTime = formatTime;
window.toggleGeneralWorkShift = toggleGeneralWorkShift;

// [Refactor] 통합 UI 갱신 함수 (Hot Reload 지원)
// 설정 변경 후 호출되어 각 화면의 요소를 강제로 최신화합니다.
// 모든 UI 갱신 로직을 이 함수 하나로 통합하여 스파게티 코드 문제를 해결합니다.
function refreshUI() {
    if (!selectedUserId) return;

    // 1. 업무일지: 작업자 명단 갱신 (팀 변경, 이름 변경 등 반영)
    // renderWorkLogWorkerList는 appConfig.users[selectedUserId]와 appSettings.colleagues를 새로 읽어옵니다.
    renderWorkLogWorkerList();

    // 2. ERP 점검: 장소 목록 및 오더번호 갱신
    // 탭이 그 때 활성화되어 있지 않더라도 DOM을 갱신해두면 나중에 탭 진입 시 최신 상태가 보입니다.
    renderERPCheck();
    updateToggleStyle();

    // 3. 프리셋 목록 갱신 (선로출입/차량일지)
    // 설정에서 프리셋이 변경되었을 때 반영하기 위해 갱신합니다.
    initPresets();
}

// [New] Headless 작업자 명단 불러오기 요청
function importWorkersFromHeadless() {
    const arbpl = settingsState.profile.arbpl;
    if (!selectedUserId || !arbpl) {
        showNativeMsgBox('분소 정보가 설정되지 않았습니다. [내 정보] 탭에서 분소를 선택해주세요.');
        return;
    }
    settingsState.importingWorkers = true;
    sendMessageToAHK({ command: 'importWorkers', arbpl });
    setTimeout(() => { settingsState.importingWorkers = false; }, 10000);
}

let pendingLocationImport = null;

function importLocationsFromHeadless() {
    if (pendingLocationImport || pendingLocationOrderCheck) return;
    const arbpl = appConfig.users?.[selectedUserId]?.profile?.arbpl;
    if (!selectedUserId || !arbpl) {
        showNativeMsgBox('분소 정보가 설정되지 않았습니다. [내 정보] 탭에서 분소를 선택해주세요.');
        return;
    }
    settingsState.importingLocations = true;
    const requestId = `${Date.now()}-${Math.random()}`;
    pendingLocationImport = { requestId, employeeId: selectedUserId, arbpl };
    sendMessageToAHK({ command: 'importLocations', employeeId: selectedUserId, requestId });
}

function hasImportedWorkerChanges(newWorkers) {
    if (!Array.isArray(newWorkers) || newWorkers.length === 0) return false;
    const current = appConfig.appSettings?.colleagues || [];
    const incoming = new Map(newWorkers.map(worker => [String(worker['사번'] || '').trim(), worker])
        .filter(([id]) => id));
    if (incoming.size === 0) return false;
    const currentIds = new Set(current.map(worker => String(worker.id || '').trim()));
    if (current.length !== incoming.size || currentIds.size !== incoming.size) return true;
    const managerId = Array.from(incoming).find(([, worker]) =>
        String(worker['직책'] || '').trim() === '분소장')?.[0];
    return current.some(worker => {
        const id = String(worker.id || '').trim();
        const imported = incoming.get(id);
        if (!imported) return true;
        const name = String(imported['이름'] || '').trim();
        const team = String(imported['근무조'] || '').trim();
        const phone = String(imported['전화번호'] || '').replace(/\D/g, '');
        const currentPhone = String(worker.phone || '').replace(/\D/g, '');
        const currentManager = worker.isManager === true || worker.isManager === 1 || worker.isManager === '1';
        return (name && name !== String(worker.name || '').trim())
            || (SETTINGS_TEAMS.includes(team) && team !== worker.team)
            || (/^01\d{8,9}$/.test(phone) && phone !== currentPhone)
            || currentManager !== (id === managerId);
    });
}

async function processPendingAutoWorkerList() {
    const pending = pendingAutoWorkerList;
    if (!pending || !selectedUserId || configReadyUserId !== selectedUserId) return;
    pendingAutoWorkerList = null;
    if (pending.employeeId !== selectedUserId || !hasImportedWorkerChanges(pending.data)) return;
    const moveToSettings = await requestWorkerImportConfirmation(
        '작업자 정보 변경사항이 감지되었습니다.\n작업자 설정으로 이동하시겠습니까?',
        '작업자 설정 이동 확인'
    );
    if (!moveToSettings || selectedUserId !== pending.employeeId
        || configReadyUserId !== pending.employeeId
        || !hasImportedWorkerChanges(pending.data)) return;
    if (document.getElementById('settings-view').style.display === 'none') switchView('settings');
    switchSettingsTab('tab-worker');
    handleWorkerListUpdate(pending.data, true);
}

function importedWorkerFromRaw(raw) {
    const id = String(raw['사번'] || '').trim();
    const rawPhone = String(raw['전화번호'] || '').replace(/\D/g, '');
    const phone = /^01\d{8,9}$/.test(rawPhone) ? formatPhoneValue(rawPhone) : '';
    return {
        id,
        name: String(raw['이름'] || '').trim(),
        team: String(raw['근무조'] || '').trim(),
        phone,
        isManager: String(raw['직책'] || '').trim() === '분소장'
    };
}

async function handleWorkerListUpdate(newWorkers, silentIfUnchanged = false) {
    settingsState.importingWorkers = false;
    if (!Array.isArray(newWorkers) || newWorkers.length === 0) return;
    const existing = new Map(settingsState.workers.map(worker => [String(worker.id).trim(), worker]));
    const incoming = new Map();
    newWorkers.forEach(raw => {
        const worker = importedWorkerFromRaw(raw);
        if (worker.id) incoming.set(worker.id, worker);
    });
    if (incoming.size === 0) return;
    const managerId = Array.from(incoming.values()).find(worker => worker.isManager)?.id;
    incoming.forEach(worker => { worker.isManager = worker.id === managerId; });

    const removals = settingsState.workers.filter(worker => !incoming.has(String(worker.id).trim()));
    const updates = [];
    const additions = [];
    incoming.forEach(worker => {
        const current = existing.get(worker.id);
        if (!current) {
            additions.push({ ...worker, driverRole: '-' });
            return;
        }
        const changes = [];
        if (worker.name && current.name.trim() !== worker.name)
            changes.push({ key: 'name', value: worker.name, label: '이름', before: current.name.trim() });
        if (SETTINGS_TEAMS.includes(worker.team) && current.team !== worker.team)
            changes.push({ key: 'team', value: worker.team, label: '근무조', before: current.team });
        if (worker.phone && current.phone.replace(/\D/g, '') !== worker.phone.replace(/\D/g, ''))
            changes.push({ key: 'phone', value: worker.phone, label: '전화번호', before: current.phone.trim() });
        if (!!current.isManager !== worker.isManager)
            changes.push({ key: 'isManager', value: worker.isManager, label: '분소장', before: current.isManager ? '체크' : '해제' });
        if (changes.length) updates.push({ current, worker, changes });
    });

    if (updates.length || additions.length || removals.length) {
        const sections = [];
        if (updates.length) sections.push('[기존 작업자 변경]\n' + updates.map(({ worker, changes }) =>
            `• ${worker.name} (${worker.id})\n` + changes.map(change =>
                `  ${change.label}: ${change.before || '(없음)'} → ${change.key === 'isManager'
                    ? (change.value ? '체크' : '해제') : change.value}`).join('\n')).join('\n'));
        if (additions.length) sections.push('[추가]\n' + additions.map(worker => `• ${worker.name} (${worker.id})`).join('\n'));
        if (removals.length) sections.push('[삭제]\n' + removals.map(worker =>
            `• ${worker.name || '(이름 없음)'} (${worker.id || '사번 없음'})`).join('\n'));
        const confirmed = await requestWorkerImportConfirmation(
            `작업자 목록을 다음과 같이 변경합니다.\n\n${sections.join('\n\n')}\n\n확인을 누르면 적용합니다.`
        );
        if (!confirmed) return;
    } else {
        if (!silentIfUnchanged) showNativeMsgBox('변경할 분소원 정보가 없습니다.');
        return;
    }

    updates.forEach(({ current, changes }) => changes.forEach(change => { current[change.key] = change.value; }));
    additions.forEach(worker => settingsState.workers.push(normalizeWorker(worker)));
    const removalSet = new Set(removals);
    settingsState.workers = settingsState.workers.filter(worker => !removalSet.has(worker));
    autoSaveSettings();
    scrollSettingsToBottom();
}

const SAP_LOCATION_IDENTITY_FIELDS = [
    'ORDER_TYPE', 'SHORT_TEXT', 'FUNCLOC', 'EQUIPMENT', 'MAINTPLANT', 'MAINTLOC', 'MAINTROOM', 'PLSECTN'
];
let pendingLocationOrderReady = null;
let pendingLocationOrderCheck = null;

function getLocationsForOrderCheck() {
    const settingsOpen = document.getElementById('settings-view')?.style.display !== 'none';
    const locations = settingsOpen && settingsState.loadedUserId === selectedUserId
        ? settingsState.locations : (appConfig.appSettings?.locations || []);
    return locations.map(normalizeLocation);
}

function processPendingLocationOrderCheck() {
    const ready = pendingLocationOrderReady;
    if (!ready || !selectedUserId || configReadyUserId !== selectedUserId) return;
    if (ready.employeeId !== selectedUserId) {
        pendingLocationOrderReady = null;
        return;
    }
    // 작업자 안내·수동 불러오기와 동시에 확인창을 열지 않는다.
    if (pendingNativeConfirmations.size || pendingAutoWorkerList || pendingLocationImport || pendingLocationOrderCheck) {
        setTimeout(processPendingLocationOrderCheck, 500);
        return;
    }
    pendingLocationOrderReady = null;
    const arbpl = appConfig.users?.[selectedUserId]?.profile?.arbpl;
    const locations = getLocationsForOrderCheck();
    const orders = [...new Set(locations.map(location => locationOrderKey(location.order)))]
        .filter(order => /^\d{1,12}$/.test(order));
    if (!arbpl) return;
    if (!locations.length) {
        // 등록된 작업장소가 하나도 없으면(최초 사용) SAP에서 조회해 채운다.
        const requestId = `${Date.now()}-${Math.random()}`;
        pendingLocationImport = { requestId, employeeId: selectedUserId, arbpl, automatic: true };
        sendMessageToAHK({ command: 'importLocations', employeeId: selectedUserId, requestId, automatic: true });
        return;
    }
    if (!orders.length) return;
    const requestId = `${Date.now()}-${Math.random()}`;
    pendingLocationOrderCheck = { requestId, employeeId: selectedUserId, arbpl, snapshot: JSON.stringify(locations) };
    sendMessageToAHK({ command: 'checkLocationOrders', requestId, employeeId: selectedUserId, orders });
}

function sapLocationIdentity(row) {
    if (!row || typeof row.SHORT_TEXT !== 'string' || !row.SHORT_TEXT
        || SAP_LOCATION_IDENTITY_FIELDS.some(field => typeof row[field] !== 'string')) return null;
    // 표시 장소명의 Trim/치환 값으로 비교하지 않는다. RFC 원본 필드 전체를 정확히 대조한다.
    return JSON.stringify(SAP_LOCATION_IDENTITY_FIELDS.map(field => row[field]));
}

function sapOrderIsWorking(row) {
    return typeof row?.U_STATUS === 'string' && /(^|\s)작업(\s|$)/.test(row.U_STATUS);
}

function planLocationOrderMaintenance(locations, existingRows, currentRows, year) {
    if (!Array.isArray(existingRows) || !Array.isArray(currentRows) || !/^\d{4}$/.test(String(year)))
        throw new Error('SAP 오더 자동 확인 결과 형식이 올바르지 않습니다.');
    const existing = new Map(existingRows.map(row => [locationOrderKey(row.ORDERID), row]));
    const registered = new Set(locations.map(location => locationOrderKey(location.order)));
    const candidates = new Map();
    currentRows.forEach(row => {
        if (!['PM01', 'PM03'].includes(row.ORDER_TYPE) || !sapOrderIsWorking(row)
            || !/^\d{8}$/.test(row.START_DATE) || row.START_DATE.slice(0, 4) !== String(year)) return;
        const identity = sapLocationIdentity(row);
        if (!identity) return;
        if (!candidates.has(identity)) candidates.set(identity, new Map());
        const id = locationOrderKey(row.ORDERID);
        if (/^\d{1,12}$/.test(id)) candidates.get(identity).set(id, row);
    });
    const replacements = [], removals = [], ambiguous = [];
    locations.forEach((location, index) => {
        const oldOrder = locationOrderKey(location.order), oldRow = existing.get(oldOrder);
        if (!oldRow) return; // 미조회·권한 제한을 완료 상태로 간주하지 않는다.
        if (oldRow.ORDER_TYPE === 'PM06') {
            if (typeof oldRow.U_STATUS === 'string' && !sapOrderIsWorking(oldRow))
                removals.push({ index, location, status: oldRow.U_STATUS });
            return;
        }
        if (!['PM01', 'PM03'].includes(oldRow.ORDER_TYPE) || !/^\d{8}$/.test(oldRow.START_DATE)
            || oldRow.START_DATE === '00000000'
            || Number(oldRow.START_DATE.slice(0, 4)) >= Number(year)) return;
        const identity = sapLocationIdentity(oldRow), matches = identity && candidates.get(identity);
        if (!matches?.size) return; // 신규 생성 전이거나 원본이 달라졌으면 다음 로그인에서 재확인한다.
        if (matches.size !== 1) {
            ambiguous.push({ location, reason: '동일한 올해 오더가 여러 건' });
            return;
        }
        const newOrder = matches.keys().next().value;
        if (registered.has(newOrder)) {
            ambiguous.push({ location, reason: '새 오더번호가 이미 목록에 등록됨' });
            return;
        }
        replacements.push({ index, location, oldOrder, newOrder });
    });
    const targetCounts = new Map();
    replacements.forEach(change => targetCounts.set(change.newOrder, (targetCounts.get(change.newOrder) || 0) + 1));
    const uniqueReplacements = replacements.filter(change => {
        if (targetCounts.get(change.newOrder) === 1) return true;
        ambiguous.push({ location: change.location, reason: '여러 기존 장소가 같은 새 오더에 연결됨' });
        return false;
    });
    return { replacements: uniqueReplacements, removals, ambiguous };
}

function applyLocationOrderMaintenance(locations, plan) {
    const replacements = new Map(plan.replacements.map(change => [change.index, change.newOrder]));
    const removals = new Set(plan.removals.map(change => change.index));
    return locations.flatMap((location, index) => removals.has(index) ? []
        : [{ ...location, ...(replacements.has(index) ? { order: replacements.get(index) } : {}) }]);
}

async function handleLocationOrderMaintenance(msg) {
    const pending = pendingLocationOrderCheck;
    if (!pending || msg.requestId !== pending.requestId) return;
    try {
        if (msg.error || msg.employeeId !== pending.employeeId || selectedUserId !== pending.employeeId
            || configReadyUserId !== pending.employeeId) return;
        const isCurrent = () => selectedUserId === pending.employeeId && configReadyUserId === pending.employeeId
            && appConfig.users?.[selectedUserId]?.profile?.arbpl === pending.arbpl
            && JSON.stringify(getLocationsForOrderCheck()) === pending.snapshot;
        if (!isCurrent()) return;
        const locations = getLocationsForOrderCheck();
        const plan = planLocationOrderMaintenance(locations, msg.data.existingRows, msg.data.currentRows, msg.data.year);
        const sections = [];
        if (plan.replacements.length) sections.push('[연도별 오더번호 교체]\n' + plan.replacements.map(change =>
            `• ${change.location.name}: ${change.oldOrder} → ${change.newOrder}`).join('\n'));
        if (plan.removals.length) sections.push('[PM06 삭제: 사용자상태 작업 아님]\n' + plan.removals.map(change =>
            `• ${change.location.name} (${change.location.order}, ${change.status || '사용자상태 없음'})`).join('\n'));
        if (plan.ambiguous.length) sections.push('[직접 확인 필요: 변경하지 않음]\n' + plan.ambiguous.map(change =>
            `• ${change.location.name} (${change.location.order}): ${change.reason}`).join('\n'));
        if (!sections.length) return;
        if (!plan.replacements.length && !plan.removals.length) {
            showNativeMsgBox(sections.join('\n\n'), '점검장소 오더 확인');
            return;
        }
        const confirmed = await requestLocationImportConfirmation(
            `점검장소 변경사항이 감지되었습니다.\n\n${sections.join('\n\n')}\n\n확인을 누르면 작업장소 설정으로 이동하여 적용합니다.`
        );
        if (!confirmed || !isCurrent()) return;
        if (document.getElementById('settings-view').style.display === 'none') switchView('settings');
        switchSettingsTab('tab-location');
        settingsState.locations = applyLocationOrderMaintenance(locations, plan);
        // 로그인별 재확인은 유지하며, 반영된 번호·삭제 목록은 즉시 저장한다.
        if (saveTimeout) { clearTimeout(saveTimeout); saveTimeout = null; }
        saveSettings();
        refreshUI();
    } catch (error) {
        showNativeMsgBox(error?.message || String(error), '점검장소 오더 확인 오류');
    } finally {
        if (pendingLocationOrderCheck === pending) pendingLocationOrderCheck = null;
    }
}

function locationOrderKey(order) {
    return String(order || '').trim().replace(/^0+(?=\d)/, '');
}

function planLocationImport(currentLocations, newLocations) {
    const seen = new Set(currentLocations.map(location => locationOrderKey(location.order)));
    const additions = [];
    newLocations.forEach(raw => {
        const location = normalizeLocation(raw);
        location.order = locationOrderKey(location.order);
        if (!/^\d+$/.test(location.order)) throw new Error('조회된 오더번호가 올바르지 않습니다.');
        if (seen.has(location.order)) return;
        seen.add(location.order);
        location.name = location.name.replace(/[-_]/g, ' ').trim();
        additions.push(location);
    });
    return { next: [...currentLocations, ...additions], additions };
}

async function handleLocationListUpdate(newLocations, employeeId, requestId) {
    const pending = pendingLocationImport;
    if (!pending || pending.requestId !== requestId) return;
    try {
        if (pending.employeeId !== employeeId || selectedUserId !== employeeId) return;
        if (!Array.isArray(newLocations)) throw new Error('조회 결과 형식이 올바르지 않습니다.');
        // 자동 조회는 설정 화면이 열려 있지 않을 수 있으므로 저장된 목록을 기준으로 한다.
        const automatic = !!pending.automatic;
        const currentLocations = () => automatic ? getLocationsForOrderCheck() : settingsState.locations;
        const snapshot = JSON.stringify(currentLocations());
        const { next, additions } = planLocationImport(currentLocations(), newLocations);
        if (!additions.length) {
            if (!automatic) showNativeMsgBox('추가할 작업장소가 없습니다.', '작업장소 불러오기');
            return;
        }
        const describe = location => `${location.name || '(이름 없음)'} (${location.order || '오더번호 없음'}, ${location.type})`;
        const confirmed = await requestLocationImportConfirmation(
            `${automatic ? '등록된 작업장소가 없어 SAP에서 조회한 ' : '새 '}작업장소 ${additions.length}건을 추가합니다.\n\n[추가]\n${additions.map(location => `• ${describe(location)}`).join('\n')}\n\n확인을 누르면 적용합니다.`
        );
        if (!confirmed || selectedUserId !== employeeId) return;
        if (JSON.stringify(currentLocations()) !== snapshot
            || appConfig.users?.[employeeId]?.profile?.arbpl !== pending.arbpl) {
            showNativeMsgBox('조회 중 설정이 변경되었습니다. 다시 불러와주세요.', '작업장소 불러오기');
            return;
        }
        if (automatic) {
            // 설정 화면을 열어 목록을 불러온 뒤 반영하고 즉시 저장한다.
            if (document.getElementById('settings-view').style.display === 'none') switchView('settings');
            switchSettingsTab('tab-location');
            settingsState.locations = next;
            if (saveTimeout) { clearTimeout(saveTimeout); saveTimeout = null; }
            saveSettings();
            refreshUI();
            return;
        }
        settingsState.locations = next;
        autoSaveSettings();
        scrollSettingsToBottom();
    } catch (error) {
        showNativeMsgBox(error?.message || String(error), '작업장소 불러오기 오류');
    } finally {
        if (pendingLocationImport === pending) {
            pendingLocationImport = null;
            settingsState.importingLocations = false;
        }
    }
}

window.importWorkersFromHeadless = importWorkersFromHeadless;
window.importLocationsFromHeadless = importLocationsFromHeadless;
