// 「测试数据」场景测试：用 test/_data_testcases.json（由 python test/dump-test-data.py 导出）驱动 js/main.js
// 运行方式：node test/scenarios.test.js
// 数据来源：../shuju/测试数据/*.xlsx（01 班组填报-正常 241 条 / 02 校验失败 30 行 /
//          07 异常表-可匹配 60 条 / 12 整改表-真实自由文本 42 条 / 13 整改表-定位异常 25 条）
// 重点断言：异常匹配口径（姓名不参与定位）、整改操作必须按业务键定位而不是按 ID 定位、
//          定位异常必须进清单而不是静默跳过、空行与越界时间必须被拦。

const fs = require('fs');
const path = require('path');
const assert = require('assert');

const JSON_PATH = path.join(__dirname, '_data_testcases.json');
if (!fs.existsSync(JSON_PATH)) {
  console.log('未找到 test/_data_testcases.json，请先从主项目运行：python test/dump-test-data.py');
  process.exit(0);
}
const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));
const src = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');

// ---------- 桩（导出用捕获桩，记录生成了哪些 sheet）----------
let captured = [];
const xlsxStub = {
  SSF: { format: () => '' },
  utils: {
    book_new: () => ({}),
    aoa_to_sheet: (rows) => ({ rows }),
    book_append_sheet: (wb, ws, name) => { captured.push({ name, rows: ws.rows }); },
    sheet_to_json: () => [],
  },
  write: () => new Uint8Array(),
};
const makeElement = () => ({
  style: {}, dataset: {}, className: '', onclick: null,
  addEventListener() {}, appendChild() {}, removeChild() {}, remove() {}, click() {},
  querySelector() { return makeElement(); },
  set innerHTML(v) {}, set textContent(v) {},
});
const documentStub = {
  addEventListener() {},
  getElementById() { return makeElement(); },
  querySelectorAll() { return []; },
  createElement() { return makeElement(); },
  body: makeElement(),
};
globalThis.Blob = class { };
globalThis.URL = { createObjectURL: () => 'blob:x', revokeObjectURL() {} };

const factory = new Function('XLSX', 'document', 'console', `${src}
return {
  appState,
  processGroupWorkbook, processAbnormalWorkbook, processRectifyWorkbook,
  confirmBatch, applyBatchOperations, exportRectify, renderRectify, renderOutput,
  exportSystemData, exportShiftData, exportOperationLog, exportLocateIssues, startNewRound,
};`);
const m = factory(xlsxStub, documentStub, console);
const appState = m.appState;

let passed = 0;
const failures = [];
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log(`  PASS  ${name}`);
  } catch (err) {
    failures.push({ name, err });
    console.log(`  FAIL  ${name}\n        ${err.message}`);
  }
}
function section(title) { console.log(`\n${title}`); }

function file(name) {
  const f = data.files[name];
  assert.ok(f, `缺少测试数据文件「${name}」`);
  return f;
}
function resetAll() {
  appState.groupWorkbook = null;
  appState.groupSheets = [];
  appState.mergedRecords = [];
  appState.abnormalWorkbook = null;
  appState.abnormalRecords = [];
  appState.abnormalFailures = [];
  appState.rectifyWorkbook = null;
  appState.rectifyOperations = [];
  appState.batchConfirmed = false;
  appState.finalGenerated = false;
  appState.rounds = [];
  appState.currentRound = 0;
  appState.groupFailures = [];
  captured = [];
}
// 合并大表快照：按对象引用比对，区分「删除」与「修改」
function snap() {
  return appState.mergedRecords.map(r => ({
    ref: r, 工号: r['工号'], 姓名: r['姓名'],
    开始日期: r['加班开始日期'], 开始时间: r['加班开始时间'], 结束时间: r['加班结束时间'], 时数: r['加班时数'],
  }));
}
function diff(before, after) {
  const afterRefs = new Set(after.map(a => a.ref));
  const deleted = before.filter(b => !afterRefs.has(b.ref));
  const modified = after.filter(a => {
    const b = before.find(x => x.ref === a.ref);
    return b && (b.开始日期 !== a.开始日期 || b.开始时间 !== a.开始时间 || b.结束时间 !== a.结束时间 || b.时数 !== a.时数);
  });
  return { deleted, modified };
}
function countBy(rows, key) {
  const map = {};
  rows.forEach(r => { map[r[key]] = (map[r[key]] || 0) + 1; });
  return map;
}

console.log('总装科月度加班自动处理工具 · v2 场景测试（真实测试数据）');

// ==================== 场景 1：01-班组填报-正常 ====================
section('场景1  01-班组填报-正常.xlsx（导入与合并）');

test('17 个班组 sheet、241 条记录、无校验失败', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  assert.strictEqual(appState.groupSheets.length, 17);
  assert.strictEqual(appState.mergedRecords.length, 241);
  assert.strictEqual(appState.groupFailures.length, 0);
  assert.ok(appState.groupSheets.every(s => s.status === 'ok'));
});

// ==================== 场景 2：02-班组填报-校验失败 ====================
section('场景2  02-班组填报-校验失败.xlsx（校验与容错）');

test('30 行数据：校验失败 17 条、进入合并 11 条', () => {
  resetAll();
  m.processGroupWorkbook(file('02-班组填报-校验失败.xlsx'));
  assert.strictEqual(appState.groupFailures.length, 17);
  assert.strictEqual(appState.mergedRecords.length, 11);
});

test('完全空行不再被记成「6 项缺失」的校验失败', () => {
  resetAll();
  m.processGroupWorkbook(file('02-班组填报-校验失败.xlsx'));
  const blank = appState.groupFailures.filter(f => String(f['失败原因']) === '工号为空；姓名为空；加班开始日期为空；加班开始时间为空；加班结束日期为空；加班结束时间为空');
  // 只剩「只有序号」的两行（10 / 29），完全空行（11 / 28）已被跳过
  assert.deepStrictEqual(blank.map(f => f['行号']).sort((a, b) => a - b), [10, 29]);
});

test('时间越界（25:70 / 24:00）被拦在合并大表之外', () => {
  resetAll();
  m.processGroupWorkbook(file('02-班组填报-校验失败.xlsx'));
  assert.ok(!appState.mergedRecords.some(r => /25:70|24:00/.test(r['加班开始时间'] + r['加班结束时间'])));
  // 4 条「加班开始时间格式错误」= 原有的 2 条乱写时间（--、abc）+ 新增拦截的 25:70 / 24:00
  const timeFail = appState.groupFailures.filter(f => String(f['失败原因']).includes('加班开始时间格式错误'));
  assert.strictEqual(timeFail.length, 4);
});

// ==================== 场景 3：07-异常表-可匹配 ====================
section('场景3  07-异常表-可匹配.xlsx（异常匹配）');

test('60 条异常：已匹配 54 / 未匹配 6 / 需人工核对 4', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processAbnormalWorkbook(file('07-异常表-可匹配.xlsx'));
  const round = appState.rounds[0];
  assert.strictEqual(appState.abnormalRecords.length, 60);
  assert.deepStrictEqual(countBy(appState.abnormalRecords, '匹配状态'), { 已匹配: 54, 未匹配: 6 });
  assert.strictEqual(round.abnormalFailures.length, 6);
  assert.strictEqual(round.abnormalWarnings.length, 4);
});

test('姓名与合并大表不一致的记录仍匹配成功，只给提醒（不再判失败）', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processAbnormalWorkbook(file('07-异常表-可匹配.xlsx'));
  const round = appState.rounds[0];
  assert.ok(round.abnormalWarnings.length === 4);
  assert.ok(round.abnormalWarnings.every(w => w['定位提醒'].includes('不一致')));
  assert.ok(round.abnormalWarnings.every(w => w['匹配状态'] === '已匹配'));
  assert.ok(round.abnormalWarnings.every(w => w['系统序号']));
  const failNos = round.abnormalFailures.map(f => f['工号']);
  round.abnormalWarnings.forEach(w => assert.ok(!failNos.includes(w['工号']), `${w['工号']} 不应判为匹配失败`));
});

test('失败清单只包含工号不存在 / 日期差一的记录', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processAbnormalWorkbook(file('07-异常表-可匹配.xlsx'));
  const round = appState.rounds[0];
  assert.strictEqual(round.abnormalFailures.filter(f => String(f['工号']).startsWith('999999')).length, 3);
  assert.ok(round.abnormalFailures.every(f => f['匹配状态'] === '未匹配'));
});

test('整改表只含匹配成功的 54 条（未匹配的 6 条不出现在整改表）', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processAbnormalWorkbook(file('07-异常表-可匹配.xlsx'));
  captured = [];
  m.exportRectify();
  const dataRows = captured.reduce((n, s) => n + Math.max(0, s.rows.length - 1), 0);
  assert.strictEqual(dataRows, 54);
});


// ==================== 场景 4：12-整改表-真实自由文本 ====================
section('场景4  12-整改表-真实自由文本.xlsx（ID 来自真实校对系统）');

test('42 条操作：6 条可识别、36 条自由文本无法识别', () => {
  resetAll();
  m.processRectifyWorkbook(file('12-整改表-真实自由文本.xlsx'));
  const ops = appState.rectifyOperations;
  assert.strictEqual(ops.length, 42);
  const known = ops.filter(o => ['修改', '删除', '调班', '特殊情况', '未填写'].includes(o['操作类型']));
  assert.strictEqual(known.length, 6);
});

test('删除 / 修改 都按业务键定位到整改表里写的那个人（不再按 ID 撞上别人）', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processRectifyWorkbook(file('12-整改表-真实自由文本.xlsx'));
  const ops = appState.rectifyOperations;
  const before = snap();
  const res = m.applyBatchOperations(ops);
  const after = snap();
  const { deleted, modified } = diff(before, after);

  const delOps = ops.filter(o => o['操作类型'] === '删除');
  const shiftOps = ops.filter(o => o['操作类型'] === '调班');
  const modOps = ops.filter(o => o['操作类型'] === '修改');
  // 36 条自由文本处置方式（"删除加班""已改""已调"…）进「未识别」清单，不再静默忽略
  assert.strictEqual(res.issues.length, 36);
  assert.ok(res.issues.every(i => i['级别'] === '未识别'));
  assert.ok(res.issues.every(i => i['说明'].includes('无法识别')));
  // 删除与调班都会把记录从合并大表移除
  assert.deepStrictEqual(
    deleted.map(x => x['工号']).sort(),
    delOps.concat(shiftOps).map(o => o['工号']).sort()
  );
  assert.deepStrictEqual(modified.map(x => x['工号']).sort(), modOps.map(o => o['工号']).sort());
  assert.strictEqual(before.length - deleted.length, after.length);
  assert.strictEqual(after.length, 239);
});

// ==================== 场景 5：13-整改表-定位异常 ====================
section('场景5  13-整改表-定位异常.xlsx（ID=901… 与合并大表序号无关）');

test('24 条操作：13 条按业务键定位、1 条未填写、10 条进入定位异常清单', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processRectifyWorkbook(file('13-整改表-定位异常.xlsx'));
  const ops = appState.rectifyOperations;
  // 24 条（旧数 25）：源文件尾部那个全空行不再算成一条操作
  // 【口径变更】空行不是数据：旧行为会把它当成“未填写处置方式”并阻断整轮，提示人去找一条根本不存在的记录
  assert.strictEqual(ops.length, 24);
  assert.deepStrictEqual(countBy(ops, '操作类型'), { 修改: 20, 删除: 3, 未填写: 1 });

  const before = snap();
  const res = m.applyBatchOperations(ops);

  assert.deepStrictEqual(countBy(ops.filter(o => o['操作类型'] !== '未填写'), '定位状态'),
    { '已定位(业务键)': 13, '未定位': 7, '多条命中(2)': 3 });
  assert.strictEqual(res.issues.length, 10);
  assert.deepStrictEqual(countBy(res.issues, '级别'), { 未定位: 7, 多条命中: 3 });
  assert.ok(res.issues.every(i => i['说明']), '每条定位异常都要有可读说明');
  assert.strictEqual(appState.mergedRecords.length, before.length, '未定位的记录不能改动合并大表');
});

test('定位异常不会被改错：实际改动的工号都在整改表里', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processRectifyWorkbook(file('13-整改表-定位异常.xlsx'));
  const ops = appState.rectifyOperations;
  const before = snap();
  const res = m.applyBatchOperations(ops);
  const after = snap();
  const { deleted, modified } = diff(before, after);
  assert.strictEqual(res.issues.length, 10);
  const opNos = new Set(ops.map(o => String(o['工号'])));
  modified.forEach(x => assert.ok(opNos.has(x['工号']), `${x['工号']} 不在整改表里，说明改错了行`));
  assert.strictEqual(deleted.length, 0, '3 条删除的整改操作都未能唯一定位，不应删掉任何记录');
});

test('存在未填写处置方式时整轮阻断，并提示补充', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processRectifyWorkbook(file('13-整改表-定位异常.xlsx'));
  const before = snap();
  m.confirmBatch();
  assert.strictEqual(appState.mergedRecords.length, before.length);
  assert.notStrictEqual(appState.rounds[0].status, 'confirmed');
  assert.ok(m.renderRectify().includes('未填写处置方式'));
});

// ==================== 场景 6：全流程连通性 ====================
section('场景6  全流程连通性（导入 → 异常 → 整改 → 确认执行 → 导出）');

test('01 + 07 + 12 走完整流程：确认执行后输出页与各导出均正常', () => {
  resetAll();
  m.processGroupWorkbook(file('01-班组填报-正常.xlsx'));
  m.processAbnormalWorkbook(file('07-异常表-可匹配.xlsx'));
  m.processRectifyWorkbook(file('12-整改表-真实自由文本.xlsx'));
  // 模拟组长把「未填写」的那一条补填为删除，使本轮可以执行
  appState.rectifyOperations.filter(o => o['操作类型'] === '未填写').forEach(o => { o['操作类型'] = '删除'; });

  captured = [];
  m.confirmBatch();
  const round = appState.rounds[appState.currentRound];
  assert.strictEqual(round.status, 'confirmed', '本轮应确认完成');
  assert.ok(Array.isArray(round.locateIssues), '应写入定位异常清单字段');

  m.exportSystemData('current');
  m.exportShiftData('current');
  m.exportOperationLog();
  m.exportLocateIssues();
  assert.ok(captured.length > 0, '应生成导出文件');
  const outputPage = m.renderOutput();
  assert.ok(outputPage.includes('第 1 轮处理完成'));
});

// ==================== 汇总 ====================
console.log(`\n通过 ${passed} 项，失败 ${failures.length} 项`);
if (failures.length) {
  failures.forEach(f => console.log(`  - ${f.name}: ${f.err.message}`));
  process.exit(1);
}

