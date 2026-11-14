const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '../ui/app.js'), 'utf8');
function extract(start, end) {
    return source.slice(source.indexOf(start), source.indexOf(end, source.indexOf(start)));
}

// --- 새 사용자 기본값 / 단축키 보정 ---
const defaults = vm.createContext({ sent: [], selectedUserId: 'u1', appConfig: { users: {} } });
defaults.sendMessageToAHK = message => defaults.sent.push(message);
vm.runInContext(extract('const defaultHotkeys = [', '\nconst hotkeyTranslations'), defaults);
vm.runInContext(extract('const NEW_USER_DAILY_LOG_DEFAULTS', '\nfunction resetAddUserForm('), defaults);
vm.runInContext('this.api = { buildNewUserDefaults, ensureUserHotkeyDefaults };', defaults);

const built = defaults.api.buildNewUserDefaults();
assert.equal(built.hotkeys.length, 6);
assert.ok(built.hotkeys.every(item => item.enabled === true && item.key && item.action));
assert.equal(built.dailyLogDefaults.safety.length, 4);
assert.deepEqual(Object.keys(built.dailyLogDefaults.safetyEdu).sort(),
    ['Fri', 'Mon', 'Sat', 'Sun', 'Thu', 'Tue', 'Wed']);
assert.ok(built.dailyLogDefaults.generalWork.every(row => row.manager === ''));
built.dailyLogDefaults.safety[0].content = '변경';
assert.notEqual(defaults.api.buildNewUserDefaults().dailyLogDefaults.safety[0].content, '변경');

defaults.appConfig.users.u1 = { hotkeys: [] };
defaults.api.ensureUserHotkeyDefaults();
assert.equal(defaults.appConfig.users.u1.hotkeys.length, 6);
assert.equal(defaults.sent.length, 1);
assert.equal(defaults.sent[0].command, 'saveConfig');
defaults.appConfig.users.u1.hotkeys = [{ action: 'AutoLogin', key: '#z', desc: '', enabled: false }];
defaults.api.ensureUserHotkeyDefaults();
assert.equal(defaults.appConfig.users.u1.hotkeys.length, 1);
assert.equal(defaults.sent.length, 1);

// --- 작업장소가 비어 있을 때 자동 조회 ---
const view = { style: { display: 'none' } };
const context = vm.createContext({
    normalizeLocation: value => ({ name: '', type: '기타업무', order: '', ...value }),
    document: { getElementById: () => view },
    messages: [], sent: [], saves: 0, confirmed: true, switched: [],
    showNativeMsgBox: text => context.messages.push(text),
    requestLocationImportConfirmation: async () => context.confirmed,
    sendMessageToAHK: message => context.sent.push(message),
    switchView: name => { context.switched.push(name); view.style.display = 'flex'; },
    switchSettingsTab: name => context.switched.push(name),
    saveSettings: () => context.saves++, refreshUI() {},
    autoSaveSettings() { throw new Error('자동 조회는 즉시 저장해야 한다'); }, scrollSettingsToBottom() {},
    setTimeout, clearTimeout, saveTimeout: null,
    pendingNativeConfirmations: new Map(), pendingAutoWorkerList: null, pendingLocationImport: null,
    selectedUserId: 'u1', configReadyUserId: 'u1',
    appConfig: { users: { u1: { profile: { arbpl: '5129' } } }, appSettings: { locations: [] } },
    settingsState: { locations: [], loadedUserId: null, importingLocations: false }
});
vm.runInContext(extract('let pendingLocationOrderReady', '\nfunction sapLocationIdentity('), context);
vm.runInContext(extract('function locationOrderKey(', '\nwindow.importWorkersFromHeadless'), context);
const run = code => vm.runInContext(code, context);

(async () => {
    // 목록이 비어 있으면 오더 확인 대신 자동 조회를 요청한다.
    run("pendingLocationOrderReady = { employeeId: 'u1' }; processPendingLocationOrderCheck();");
    assert.equal(context.sent.length, 1);
    assert.equal(context.sent[0].command, 'importLocations');
    assert.equal(context.sent[0].automatic, true);
    const requestId = context.sent[0].requestId;

    // 조회 결과를 확인 후 설정 화면에 반영하고 즉시 저장한다.
    await context.handleLocationListUpdate([{ order: '0001', name: '장소-A', type: '변전소' }], 'u1', requestId);
    assert.deepEqual(context.switched, ['settings', 'tab-location']);
    assert.equal(context.settingsState.locations.length, 1);
    assert.equal(context.settingsState.locations[0].order, '1');
    assert.equal(context.saves, 1);
    assert.equal(context.pendingLocationImport, null);

    // 조회 결과가 없으면 조용히 끝낸다.
    context.settingsState.locations = []; view.style.display = 'none'; context.sent.length = 0;
    run("pendingLocationOrderReady = { employeeId: 'u1' }; processPendingLocationOrderCheck();");
    await context.handleLocationListUpdate([], 'u1', context.sent[0].requestId);
    assert.equal(context.messages.length, 0);
    assert.equal(context.saves, 1);

    // 목록이 이미 있으면 기존 오더 확인 흐름을 탄다.
    context.appConfig.appSettings.locations = [{ order: '123', name: '기존', type: '변전소' }];
    context.sent.length = 0;
    run("pendingLocationOrderReady = { employeeId: 'u1' }; processPendingLocationOrderCheck();");
    assert.equal(context.sent[0].command, 'checkLocationOrders');
    console.log('new-user-defaults: ok');
})().catch(error => { console.error(error); process.exit(1); });
