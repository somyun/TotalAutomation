const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const source = fs.readFileSync(path.join(__dirname, '../ui/app.js'), 'utf8');
const section = (start, end) => source.slice(source.indexOf(start), source.indexOf(end, source.indexOf(start)));
const context = vm.createContext({
    normalizeLocation: row => ({ ...row }), messages: [], saves: 0,
    settingsState: { locations: [], loadedUserId: 'test' }, selectedUserId: 'test', configReadyUserId: 'test',
    appConfig: { users: { test: { profile: { arbpl: '5129' } } }, appSettings: { locations: [] } },
    pendingNativeConfirmations: new Map(), pendingAutoWorkerList: null, pendingLocationImport: null,
    requestLocationImportConfirmation: async () => true,
    showNativeMsgBox: text => context.messages.push(text), saveSettings: () => context.saves++, refreshUI() {},
    document: { getElementById: () => ({ style: { display: 'block' } }) },
    switchView() {}, switchSettingsTab: tab => assert.equal(tab, 'tab-location'), saveTimeout: null,
    sendMessageToAHK: msg => context.sent.push(msg), sent: [], setTimeout, clearTimeout
});
vm.runInContext(section('const SAP_LOCATION_IDENTITY_FIELDS', 'function locationOrderKey(')
    + section('function locationOrderKey(', 'function planLocationImport('), context);
const row = (id, type = 'PM01', year = '2025', status = '완료', text = '원본 작업_장소', extra = {}) => ({
    ORDERID: String(id).padStart(12, '0'), ORDER_TYPE: type, START_DATE: year + '0101', U_STATUS: status,
    SHORT_TEXT: text, FUNCLOC: '시설', EQUIPMENT: '', MAINTPLANT: '5010', MAINTLOC: '장소코드', MAINTROOM: '', PLSECTN: '', ...extra
});
const location = (id, name = '사용자 지정 이름', type = '전기실(그룹2)') => ({ order: String(id), name, type });
const old = row('1'), fresh = row('2', 'PM01', '2026', '작업');
const plan = (locations, existing = [old], current = [fresh]) =>
    context.planLocationOrderMaintenance(locations, existing, current, '2026');
let result = plan([location('1')]);
assert.equal(result.replacements.length, 1); // 전년도 완료 상태도 원본 대조 가능
let next = context.applyLocationOrderMaintenance([location('1')], result);
assert.equal(next[0].order, '2');
assert.equal(next[0].name, '사용자 지정 이름');
assert.equal(next[0].type, '전기실(그룹2)');
assert.equal(plan([location('1')], [old], [row('2', 'PM01', '2026', '작업', '원본 작업-장소')]).replacements.length, 0);
assert.equal(plan([location('1')], [old], [{ ...fresh, FUNCLOC: '다른시설' }]).replacements.length, 0);
assert.equal(plan([location('1')], [old], [{ ...fresh, MAINTLOC: '다른장소' }]).replacements.length, 0);
assert.equal(plan([location('1')], [old], [{ ...fresh, ORDER_TYPE: 'PM03' }]).replacements.length, 0);
assert.equal(plan([location('1')], [old], [{ ...fresh, U_STATUS: '완료' }]).replacements.length, 0);
assert.equal(plan([location('1')], [old], [{ ...fresh, START_DATE: '20270101' }]).replacements.length, 0);
assert.equal(plan([location('1')], [], [fresh]).replacements.length, 0);
assert.equal(plan([location('1')], [old], []).replacements.length, 0);
assert.equal(plan([location('1')], [row('1', 'PM01', '2026')], [fresh]).replacements.length, 0);
assert.equal(plan([location('1')], [{ ...old, START_DATE: '00000000' }], [fresh]).replacements.length, 0);
assert.equal(context.planLocationOrderMaintenance([location('1')], [row('1', 'PM01', '2026')],
    [row('2', 'PM01', '2027', '작업')], '2027').replacements.length, 1);
assert.equal(plan([location('1')], [row('1', 'PM01', '2024')], [fresh]).replacements.length, 1);
result = plan([location('1')], [old], [fresh, { ...fresh, ORDERID: '000000000003' }]);
assert.equal(result.replacements.length, 0);
assert.equal(result.ambiguous.length, 1);
assert.equal(plan([location('1'), location('2')]).replacements.length, 0);
assert.equal(plan([location('1'), location('1')]).replacements.length, 0);
assert.equal(plan([location('1')], [old], [{ ORDERID: '2', ORDER_TYPE: 'PM01', START_DATE: '20260101', U_STATUS: '작업', SHORT_TEXT: old.SHORT_TEXT }]).replacements.length, 0);
result = plan([location('6'), location('7'), location('8')], [row('6', 'PM06', '2026', '완료'), row('7', 'PM06', '2024', '작업'), row('8', 'PM06', '2026', '작업 보류')], []);
assert.equal(result.removals.length, 1);
assert.equal(context.applyLocationOrderMaintenance([location('6'), location('7'), location('8')], result).length, 2);
assert.equal(plan([location('6')], [], []).removals.length, 0);
assert.equal(plan([location('6')], [{ ...row('6', 'PM06'), U_STATUS: undefined }], []).removals.length, 0);
assert.equal(plan([location('1')], [old], []).removals.length, 0);
assert.equal(plan([location('3')], [row('3', 'PM03')], [row('4', 'PM03', '2026', '작업')]).replacements.length, 1);

(async () => {
    const reset = () => {
        context.settingsState.locations = [location('1')];
        context.saves = 0; context.selectedUserId = 'test'; context.messages = [];
        context.requestLocationImportConfirmation = async () => true;
        context.snapshot = JSON.stringify(context.settingsState.locations);
        vm.runInContext("pendingLocationOrderCheck = {requestId:'request', employeeId:'test', arbpl:'5129', snapshot};", context);
    };
    const msg = { employeeId: 'test', requestId: 'request', data: { year: '2026', existingRows: [old], currentRows: [fresh] } };
    reset(); await context.handleLocationOrderMaintenance(msg);
    assert.equal(context.settingsState.locations[0].order, '2'); assert.equal(context.saves, 1);
    reset(); context.requestLocationImportConfirmation = async () => false;
    await context.handleLocationOrderMaintenance(msg); assert.equal(context.saves, 0);
    assert.equal(context.settingsState.locations[0].order, '1');
    reset(); await context.handleLocationOrderMaintenance({ ...msg, error: true });
    assert.equal(context.saves, 0); assert.equal(context.messages.length, 0);
    reset(); await context.handleLocationOrderMaintenance({ ...msg, requestId: 'stale' }); assert.equal(context.saves, 0);
    reset(); context.selectedUserId = 'other'; await context.handleLocationOrderMaintenance(msg); assert.equal(context.saves, 0);
    reset(); context.requestLocationImportConfirmation = async () => { context.settingsState.locations[0].name = '편집'; return true; };
    await context.handleLocationOrderMaintenance(msg); assert.equal(context.saves, 0); assert.equal(context.settingsState.locations[0].name, '편집');
    reset();
    vm.runInContext("pendingLocationOrderCheck = null; pendingLocationOrderReady = {employeeId:'test'};", context);
    context.processPendingLocationOrderCheck();
    assert.equal(context.sent.at(-1).command, 'checkLocationOrders');
    assert.equal(context.sent.at(-1).orders[0], '1');
    assert.equal(context.sent.length, 1); context.processPendingLocationOrderCheck(); assert.equal(context.sent.length, 1);
    console.log('PASS: 완료 오더 원본 대조·PM01/PM03 번호 교체·원본 차이·연도/상태·복수 후보·미조회 보존·PM06 상태별 삭제·확인/취소/실패/사용자 전환/편집 보호·로그인 조회');
})().catch(error => { console.error(error); process.exitCode = 1; });
