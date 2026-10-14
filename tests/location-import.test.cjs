const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');

const source = fs.readFileSync(path.join(__dirname, '../ui/app.js'), 'utf8');
function extract(start, end) {
    return source.slice(source.indexOf(start), source.indexOf(end, source.indexOf(start)));
}
const context = vm.createContext({
    normalizeLocation: value => ({ name: '', type: '기타업무', order: '', ...value }),
    messages: [], saves: 0, confirmed: true,
    showNativeMsgBox: text => context.messages.push(text),
    requestLocationImportConfirmation: async () => context.confirmed,
    autoSaveSettings: () => context.saves++, scrollSettingsToBottom() {},
    selectedUserId: 'test', appConfig: { users: { test: { profile: { arbpl: '5129' } } } },
    settingsState: { locations: [], importingLocations: true },
    pendingLocationImport: { employeeId: 'test', requestId: 'request', arbpl: '5129' }
});
vm.runInContext(extract('function locationOrderKey(', '\nwindow.importWorkersFromHeadless'), context);
const row = (order, name = '장소', type = '기타업무') => ({ order, name, type });
const current = [row('0000123'), row('123'), row('999'), row('456', '기존')];
const incoming = [row('123', '새 이름', '변전소'), row('456', '변경'), row('789', '  -새_장소-  '), row('0000789')];
let plan = context.planLocationImport(current, incoming);
assert.equal(plan.next.length, 5);
current.forEach((location, index) => assert.equal(plan.next[index], location));
assert.equal(plan.additions.length, 1);
assert.equal(plan.additions[0].name, '새 장소');
assert.equal(context.planLocationImport([row('000123')], [row('123')]).additions.length, 0);
assert.equal(context.planLocationImport([row('123')], []).next.length, 1);
assert.equal(context.planLocationImport([], [row('1', '  이름  ')]).additions[0].name, '이름');
assert.throws(() => context.planLocationImport([], [row('')]), /오더번호/);

(async () => {
    const reset = () => {
        context.settingsState = { locations: [row('123', '기존')], importingLocations: true };
        context.pendingLocationImport = { employeeId: 'test', requestId: 'request', arbpl: '5129' };
        context.selectedUserId = 'test';
        context.saves = 0;
    };
    reset(); context.confirmed = false;
    await context.handleLocationListUpdate([row('456')], 'test', 'request');
    assert.equal(context.settingsState.locations[0].order, '123');
    assert.equal(context.saves, 0);
    assert.equal(context.settingsState.importingLocations, false);
    reset(); context.confirmed = true;
    await context.handleLocationListUpdate([row('456')], 'test', 'request');
    assert.equal(context.settingsState.locations[0].order, '123');
    assert.equal(context.settingsState.locations[1].order, '456');
    assert.equal(context.saves, 1);
    reset();
    await context.handleLocationListUpdate([], 'test', 'request');
    assert.equal(context.settingsState.locations.length, 1);
    assert.equal(context.saves, 0);
    reset();
    await context.handleLocationListUpdate([row('123', '다른 이름', '변전소')], 'test', 'request');
    assert.equal(context.settingsState.locations[0].name, '기존');
    assert.equal(context.settingsState.locations[0].type, '기타업무');
    assert.equal(context.saves, 0);
    reset();
    await context.handleLocationListUpdate([row('456')], 'test', 'stale');
    assert.equal(context.settingsState.locations[0].order, '123');
    assert.equal(context.pendingLocationImport.requestId, 'request');
    reset(); context.selectedUserId = 'other';
    await context.handleLocationListUpdate([row('456')], 'test', 'request');
    assert.equal(context.saves, 0);
    reset();
    context.requestLocationImportConfirmation = async () => {
        context.settingsState.locations.push(row('888'));
        return true;
    };
    await context.handleLocationListUpdate([row('456')], 'test', 'request');
    assert.equal(context.settingsState.locations.length, 2);
    assert.equal(context.saves, 0);
    console.log('PASS: 기존 항목 보존·오더번호 중복·선행 0·장소명 Trim 및 구분자 치환·추가·빈 결과·확인·취소·오래된 응답·사용자 전환·동시 편집');
})().catch(error => { console.error(error); process.exitCode = 1; });
