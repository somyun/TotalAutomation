const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync(require('node:path').join(__dirname, '../ui/app.js'), 'utf8');
const context = vm.createContext({ Vue: { reactive: value => value }, selectedUserId: 'a', configReadyUserId: 'a',
    appConfig: { users: { a: { profile: { arbpl: '5129' } }, b: { profile: { arbpl: '5111' } } } },
    SETTINGS_DEPARTMENTS: [{ value: '5129', name: '호포전기분소' }, { value: '5111', name: '동래변전분소' }, { value: '5133', name: '남양산전기관리소' }],
    sent: [], sendMessageToAHK(payload) { context.sent.push(payload); }
});
vm.runInContext(source.slice(source.indexOf('const materialLocationState ='), source.indexOf('function emptySafetyDefaults()')), context);
vm.runInContext('globalThis.state = materialLocationState;', context);
context.syncMaterialLocationDefaults();
assert.deepEqual(Array.from(context.state.locations), ['5129']);
assert.deepEqual(Array.from(context.materialDepartmentOptions(), option => option.name), ['호포전기', '동래변전', '남양산관리']);
context.state.locations = ['5129', '5111'];
context.state.keyword = '  균압선 끈끈이  ';
context.searchMaterialLocations();
assert.equal(context.sent.length, 1);
assert.equal(context.sent[0].search, '균압선 끈끈이');
assert.deepEqual(Array.from(context.sent[0].locations), ['5129', '5111']);
const first = context.sent[0];
context.searchMaterialLocations();
assert.equal(context.sent.length, 1);
context.handleMaterialLocationResult({ employeeId: 'a', requestId: 'stale', rows: [] });
assert.equal(context.state.loading, true);
context.handleMaterialLocationResult({ employeeId: 'a', requestId: first.requestId, rows: [
    { material: '001', bin: '', newStock: 0, usedStock: 2 },
    { material: '002', bin: '', newStock: 0, usedStock: 0 },
    { material: '003', bin: '', newStock: -1, usedStock: 1 }
] });
assert.equal(context.state.rows.length, 1);
assert.equal(context.state.rows[0].bin, '');
assert.equal(context.state.rows[0].usedStock, 2);
assert.equal(context.state.loading, false);
context.searchMaterialLocations();
const second = context.sent[1];
context.selectedUserId = 'b'; context.configReadyUserId = 'b';
context.syncMaterialLocationDefaults();
assert.deepEqual(Array.from(context.state.locations), ['5111']);
context.handleMaterialLocationResult({ employeeId: 'a', requestId: second.requestId, rows: [{ material: 'old' }] });
assert.equal(context.state.rows.length, 0);
context.state.locations = [];
context.searchMaterialLocations();
assert.match(context.state.error, /저장위치/);
context.state.locations = ['5111']; context.searchMaterialLocations();
context.handleMaterialLocationResult({ employeeId: 'b', requestId: context.state.requestId, error: '조회 실패' });
assert.equal(context.state.error, '조회 실패');
assert.equal(context.state.loading, false);
const html = fs.readFileSync(require('node:path').join(__dirname, '../ui/index.html'), 'utf8');
assert.match(html, /@submit.prevent="search"/);
assert.ok(html.indexOf('<span class="label">차량운행일지</span>') < html.indexOf('<span class="label">적재위치조회</span>'));
console.log('PASS: 기본 분소·복수 선택·공백 검색어 전달·중복 조회 방지·오래된 응답·사용자 전환·빈 적재위치·중고 수량·실패 복구·Enter 제출');
