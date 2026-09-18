// 总装科月度加班自动处理工具 —— 纯函数断言测试
// 运行方式：node test/main.test.js
// 说明：js/main.js 是浏览器脚本（无模块导出），这里把源码包进一个函数并按需注入 XLSX / document /
//      Blob / URL 桩，只测试不依赖真实浏览器环境的逻辑；导出函数用「捕获桩」记录生成的工作表，
//      因此可以断言"没有数据时不会导出示例数据"这类行为。
// 覆盖重点：异常表匹配（姓名不参与定位）、批量操作定位（业务键优先 + ID 回退 + 定位异常清单）、
//          重复执行保护、导出兜底保护。

const fs = require('fs');
const path = require('path');
const assert = require('assert');

const MAIN_JS = path.join(__dirname, '..', 'js', 'main.js');
const src = fs.readFileSync(MAIN_JS, 'utf8');

// ---------- 桩 ----------
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
  toYYYYMMDD, normalizeDate, normalizeTime, padTime, parseDateParts, parseTimeParts, computeHours,
  locateMergedRecords, resolveOperationTarget, applyBatchOperations,
  processGroupWorkbook, processAbnormalWorkbook, processRectifyWorkbook, confirmBatch,
  buildSystemRecords, buildShiftRecords,
  renderImport, renderAbnormal, renderRectify, renderOutput,
  exportSystemData, exportRectify, exportShiftData, exportAbnormalFailures, exportLocateIssues,
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

function resetState() {
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

// 合并大表样例：序号 3 与「校对系统 ID 45」无关，用来验证不再按 ID 定位
function makeMergedRecords() {
  return [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 加班开始日期: '2026-08-01', 加班开始时间: '15:45', 加班结束日期: '2026-08-01', 加班结束时间: '17:35', 加班时数: 1.83 },
    { 系统序号: 2, 工号: '10010002', 姓名: '李四', 班组: '底盘一组', 加班开始日期: '2026-08-02', 加班开始时间: '07:00', 加班结束日期: '2026-08-02', 加班结束时间: '15:00', 加班时数: 8 },
    { 系统序号: 3, 工号: '10010003', 姓名: '王五', 班组: '前悬一组', 加班开始日期: '2026-08-02', 加班开始时间: '15:45', 加班结束日期: '2026-08-02', 加班结束时间: '17:35', 加班时数: 1.83 },
    { 系统序号: 4, 工号: '10010004', 姓名: '赵六', 班组: '电装一组', 加班开始日期: '2026-08-03', 加班开始时间: '20:00', 加班结束日期: '2026-08-03', 加班结束时间: '22:00', 加班时数: 2 },
  ];
}

// 造一份"异常表"解析结果（sheet_to_json 视角：第一行表头，其余为数据行）
function abnormalParsed(rows, headers = ['ID', '工号', '姓名', '科室', '开始日期', '开始时间', '结束日期', '结束时间', '上报加班时数']) {
  return { fileName: '异常表.xlsx', sheetNames: ['S1'], sheets: { S1: [headers, ...rows] } };
}
// 造一份"整改表"解析结果（整改表 = 异常列 + 处置列）
function rectifyParsed(rows) {
  const headers = ['ID', '工号', '姓名', '科室', '开始日期', '开始时间', '结束日期', '结束时间', '上报加班时数',
    '处置方式', '修改后开始日期', '修改后开始时间', '修改后结束日期', '修改后结束时间', '修改后上报加班时数', '调班日期', '调班班次', '异常说明（必填）'];
  return { fileName: '整改表.xlsx', sheetNames: ['S1'], sheets: { S1: [headers, ...rows] } };
}

console.log('总装科月度加班自动处理工具 · v2 单元测试');


// ==================== 工具函数 ====================
section('日期时间归一化');

test('padTime 把 8:15 补成 08:15（Excel h:mm 与文本 08:15 对齐）', () => {
  assert.strictEqual(m.padTime('8:15'), '08:15');
  assert.strictEqual(m.padTime('15:45:00'), '15:45');
  assert.strictEqual(m.padTime('15:45'), '15:45');
});

test('parseTimeParts 拦截越界时间', () => {
  assert.strictEqual(m.parseTimeParts('15:45').h, 15);
  assert.strictEqual(m.parseTimeParts('7:00').h, 7);
  assert.strictEqual(m.parseTimeParts('15:45:00').m, 45);
  assert.strictEqual(m.parseTimeParts('25:70'), null);
  assert.strictEqual(m.parseTimeParts('24:00'), null);
});

test('toYYYYMMDD 整串匹配，不再把 2026/8/10x 截成合法日期', () => {
  assert.strictEqual(m.toYYYYMMDD('2026-08-01'), '20260801');
  assert.strictEqual(m.toYYYYMMDD('2026/8/1'), '20260801');
  assert.strictEqual(m.toYYYYMMDD(20260801), '20260801');
  assert.ok(!/^\d{8}$/.test(m.toYYYYMMDD('2026/8/10x')), '带尾巴的日期不应被当成合法 8 位日期');
});

// ==================== 合并大表定位 ====================
section('合并大表定位（业务键）');

test('locateMergedRecords 按 工号 + 日期 唯一定位', () => {
  const recs = makeMergedRecords();
  const hits = m.locateMergedRecords('10010003', '20260802', '', recs);
  assert.strictEqual(hits.length, 1);
  assert.strictEqual(hits[0]['系统序号'], 3);
});

test('locateMergedRecords 时间比较兼容 7:00 / 07:00 / 15:45:00', () => {
  const recs = makeMergedRecords();
  assert.strictEqual(m.locateMergedRecords('10010002', '20260802', '7:00', recs).length, 1);
  assert.strictEqual(m.locateMergedRecords('10010001', '20260801', '15:45:00', recs).length, 1);
});

test('locateMergedRecords 时间对不上时返回空，不降级为同日候选', () => {
  const recs = makeMergedRecords();
  assert.strictEqual(m.locateMergedRecords('10010001', '20260801', '09:30', recs).length, 0);
});

test('locateMergedRecords 日期不匹配时不误配到其它日期', () => {
  const recs = makeMergedRecords();
  assert.strictEqual(m.locateMergedRecords('10010001', '20260805', '', recs).length, 0);
});

test('locateMergedRecords 同工号同日多条、未填时间时返回全部候选', () => {
  const recs = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '15:45' },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '20:00' },
  ];
  assert.strictEqual(m.locateMergedRecords('10010001', '20260801', '', recs).length, 2);
  assert.strictEqual(m.locateMergedRecords('10010001', '20260801', '20:00', recs).length, 1);
});

// ==================== 异常表匹配 ====================
section('异常表匹配（姓名不参与定位）');

test('姓名多一个字时仍按 工号+日期+时间 匹配成功，并给出姓名不一致提醒', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三丰', '底盘一组', 20260801, '15:45', 20260801, '18:45', 3]]));
  const rec = appState.abnormalRecords[0];
  const round = appState.rounds[0];
  assert.strictEqual(rec['匹配状态'], '已匹配');
  assert.strictEqual(rec['系统序号'], 1);
  assert.strictEqual(round.abnormalFailures.length, 0);
  assert.strictEqual(round.abnormalWarnings.length, 1);
  assert.ok(round.abnormalWarnings[0]['定位提醒'].includes('不一致'));
});

test('异常表姓名为空时仍能匹配成功', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010002', '', '底盘一组', 20260802, '07:00', 20260802, '15:00', 8]]));
  assert.strictEqual(appState.abnormalRecords[0]['匹配状态'], '已匹配');
  assert.strictEqual(appState.abnormalRecords[0]['系统序号'], 2);
});

test('工号不存在 / 日期差一天时进入失败清单', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([
    [900, '99999999', '查无此人', '底盘一组', 20260801, '15:45', 20260801, '18:45', 3],
    [901, '10010001', '张三', '底盘一组', 20260805, '15:45', 20260805, '18:45', 3],
  ]));
  assert.strictEqual(appState.abnormalRecords.length, 2);
  assert.ok(appState.abnormalRecords.every(r => r['匹配状态'] === '未匹配'));
  assert.strictEqual(appState.rounds[0].abnormalFailures.length, 2);
});

test('同工号同日多条且未填开始时间 → 标记「多条命中」并提醒', () => {
  resetState();
  appState.mergedRecords = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 加班开始日期: '2026-08-01', 加班开始时间: '15:45' },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 班组: '底盘二组', 加班开始日期: '2026-08-01', 加班开始时间: '20:00' },
  ];
  m.processAbnormalWorkbook(abnormalParsed([[7, '10010001', '张三', '底盘一组', 20260801, '', 20260801, '', 3]]));
  const round = appState.rounds[0];
  assert.strictEqual(appState.abnormalRecords[0]['匹配状态'], '多条命中(2)');
  assert.strictEqual(round.abnormalWarnings.length, 1);
  assert.ok(round.abnormalWarnings[0]['定位提醒'].includes('2 条'));
});

test('多条命中时，候选中与异常表科室一致的那条会被点名', () => {
  resetState();
  appState.mergedRecords = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 加班开始日期: '2026-08-01', 加班开始时间: '15:45' },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 班组: '前悬一组', 加班开始日期: '2026-08-01', 加班开始时间: '15:45' },
  ];
  m.processAbnormalWorkbook(abnormalParsed([[7, '10010001', '张三', '前悬一组', 20260801, '15:45', 20260801, '18:45', 3]]));
  const warn = appState.rounds[0].abnormalWarnings[0];
  assert.strictEqual(appState.abnormalRecords[0]['匹配状态'], '多条命中(2)');
  assert.ok(warn['定位提醒'].includes('与表里填写的科室/班组一致'), `提醒里应点名科室一致的那条：${warn['定位提醒']}`);
  assert.ok(warn['定位提醒'].includes('序号 2（前悬一组）'));
});

test('整改表里多条命中时同样点名与本 sheet 班组一致的候选', () => {
  resetState();
  appState.mergedRecords = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 加班开始日期: '2026-08-01', 加班开始时间: '15:45' },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 班组: '前悬一组', 加班开始日期: '2026-08-01', 加班开始时间: '20:00' },
  ];
  m.processRectifyWorkbook({
    fileName: '整改表.xlsx', sheetNames: ['前悬一组'],
    sheets: { 前悬一组: [
      ['ID', '工号', '姓名', '科室', '开始日期', '开始时间', '结束日期', '结束时间', '上报加班时数', '处置方式'],
      [2, '10010001', '张三', '前悬一组', 20260801, '', 20260801, '', 3, '修改'],
    ] },
  });
  const res = m.applyBatchOperations(appState.rectifyOperations);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '多条命中');
  assert.ok(res.issues[0]['说明'].includes('与表里填写的科室/班组一致'), res.issues[0]['说明']);
});

// ==================== 整改表解析 ====================
section('整改表解析');

test('整改表保留原始定位键（开始日期 / 开始时间）与处置方式', () => {
  resetState();
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010003', '王五', '前悬一组', 20260802, '15:45', 20260802, '17:35', 1.83, '删除', '', '', '', '', '', '', '', ''],
  ]));
  const op = appState.rectifyOperations[0];
  assert.strictEqual(op['操作类型'], '删除');
  assert.strictEqual(String(op['原开始日期']), '20260802');
  assert.strictEqual(op['原开始时间'], '15:45');
});

test('未填写处置方式标记为「未填写」', () => {
  resetState();
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010003', '王五', '前悬一组', 20260802, '15:45', 20260802, '17:35', 1.83, '', '', '', '', '', '', '', '', ''],
  ]));
  assert.strictEqual(appState.rectifyOperations[0]['操作类型'], '未填写');
});

// ==================== 批量操作定位 ====================
section('批量操作定位');

test('校对 ID 与合并大表序号不一致时，仍按业务键改对记录', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  const op = { 系统序号: 45, 工号: '10010003', 姓名: '王五', 班组: '前悬一组', 操作类型: '删除', roundNo: 1, 原开始日期: '20260802', 原开始时间: '15:45' };
  const res = m.applyBatchOperations([op]);
  assert.deepStrictEqual(res.issues, []);
  assert.strictEqual(op['定位状态'], '已定位(业务键)');
  assert.ok(!appState.mergedRecords.some(r => r['工号'] === '10010003'), '应删掉业务键命中的那一条');
  assert.strictEqual(appState.mergedRecords.length, 3);
});

test('ID 撞上别人时不删错人，按业务键定位到正确记录', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  // 序号 3 在合并大表里是王五；整改表写的 ID=3，但业务键指向李四（20260802 07:00）
  const op = { 系统序号: 3, 工号: '10010002', 姓名: '李四', 班组: '底盘一组', 操作类型: '删除', roundNo: 1, 原开始日期: '20260802', 原开始时间: '07:00' };
  const res = m.applyBatchOperations([op]);
  assert.deepStrictEqual(res.issues, []);
  assert.ok(!appState.mergedRecords.some(r => r['工号'] === '10010002'), '应删掉李四');
  assert.ok(appState.mergedRecords.some(r => r['工号'] === '10010003'), '不能误删王五');
});

test('业务键找不到时记入定位异常清单，且不改动任何记录', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  const op = { 系统序号: 900, 工号: '10010099', 姓名: '查无', 班组: '未知', 操作类型: '删除', roundNo: 1, 原开始日期: '20260801', 原开始时间: '15:45' };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '未定位');
  assert.strictEqual(res.issues[0]['操作类型'], '删除');
  assert.strictEqual(op['定位状态'], '未定位');
  assert.strictEqual(appState.mergedRecords.length, 4, '不应误删任何记录');
});

test('无原始日期时回退到 ID 等于系统序号', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  const op = { 系统序号: 2, 工号: '99999999', 姓名: '未知', 操作类型: '删除', roundNo: 1 };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 0);
  assert.strictEqual(op['定位状态'], '已定位(ID回退)');
  assert.strictEqual(appState.mergedRecords.length, 3);
});

test('多条命中时列出候选序号，并只作用于第一条', () => {
  resetState();
  appState.mergedRecords = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '15:45', 加班时数: 1.83 },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '20:00', 加班时数: 2 },
  ];
  const op = { 系统序号: 1, 工号: '10010001', 姓名: '张三', 操作类型: '删除', roundNo: 1, 原开始日期: '20260801', 原开始时间: '' };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '多条命中');
  assert.strictEqual(res.issues[0]['候选序号'], '1、2');
  assert.strictEqual(appState.mergedRecords.length, 1);
  assert.strictEqual(appState.mergedRecords[0]['加班开始时间'], '20:00', '只删掉第一条');
});

test('修改 / 删除 / 调班 / 特殊情况混合执行，全部唯一定位且无定位异常', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  const ops = [
    { 系统序号: 45, 工号: '10010001', 姓名: '张三', 操作类型: '修改', roundNo: 1, 原开始日期: '20260801', 原开始时间: '15:45', 修改后开始日期: '2026-08-01', 修改后开始时间: '18:00', 修改后结束日期: '2026-08-01', 修改后结束时间: '21:00', 修改后上报加班时数: 3 },
    { 系统序号: 46, 工号: '10010004', 姓名: '赵六', 操作类型: '删除', roundNo: 1, 原开始日期: '20260803', 原开始时间: '20:00' },
    { 系统序号: 47, 工号: '10010003', 姓名: '王五', 操作类型: '调班', roundNo: 1, 原开始日期: '20260802', 原开始时间: '15:45' },
    { 系统序号: 48, 工号: '10010002', 姓名: '李四', 操作类型: '特殊情况', roundNo: 1 },
  ];
  const res = m.applyBatchOperations(ops);
  assert.deepStrictEqual(res.issues, []);
  assert.strictEqual(appState.mergedRecords.length, 2, '删除 / 调班各移除一条');
  const zhangsan = appState.mergedRecords.find(r => r['工号'] === '10010001');
  assert.strictEqual(zhangsan['加班开始时间'], '18:00');
  assert.strictEqual(zhangsan['加班结束时间'], '21:00');
  assert.strictEqual(zhangsan['加班时数'], 3);
  assert.deepStrictEqual(appState.mergedRecords.map(r => r['系统序号']), [1, 2], '删除后重新连续编号');
  assert.strictEqual(ops[0]['定位状态'], '已定位(业务键)');
  assert.strictEqual(ops[3]['定位状态'], '无需定位');
});

test('合并大表为空时所有定位类操作进入清单，且 applied = 0', () => {
  resetState();
  appState.mergedRecords = [];
  const res = m.applyBatchOperations([{ 系统序号: 1, 工号: '10010001', 操作类型: '修改', roundNo: 1 }]);
  assert.strictEqual(res.applied, 0);
  assert.strictEqual(res.issues.length, 1);
  assert.ok(res.issues[0]['说明'].includes('合并大表为空'));
});


// ==================== 确认执行 ====================
section('确认执行批量操作');

function setupRound(ops, status = 'pending') {
  appState.rectifyOperations = ops;
  appState.rounds = [{ roundNo: 1, abnormalRecords: [], abnormalFailures: [], abnormalWarnings: [], rectifyOperations: ops, shiftRecords: [], systemRecords: [], locateIssues: [], status }];
  appState.currentRound = 0;
  appState.batchConfirmed = false;
}

test('执行后本轮标记为 confirmed，并写入定位异常清单', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  setupRound([{ 系统序号: 900, 工号: '10010099', 姓名: '查无', 操作类型: '删除', roundNo: 1, 原开始日期: '20260801', 原开始时间: '15:45' }]);
  m.confirmBatch();
  assert.strictEqual(appState.rounds[0].status, 'confirmed');
  assert.strictEqual(appState.rounds[0].locateIssues.length, 1);
});

test('批量操作不能重复执行（防止按重排后的序号删错人）', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  setupRound([{ 系统序号: 1, 工号: '10010001', 姓名: '张三', 操作类型: '删除', roundNo: 1, 原开始日期: '20260801', 原开始时间: '15:45' }]);
  m.confirmBatch();
  const afterFirst = appState.mergedRecords.map(r => r['工号']);
  assert.strictEqual(afterFirst.length, 3);
  m.confirmBatch();
  assert.deepStrictEqual(appState.mergedRecords.map(r => r['工号']), afterFirst, '第二次执行不应改动数据');
  assert.strictEqual(appState.mergedRecords.length, 3);
});

test('存在未填写处置方式时 confirmBatch 整轮阻断', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  setupRound([
    { 系统序号: 1, 工号: '10010001', 操作类型: '未填写', roundNo: 1 },
    { 系统序号: 2, 工号: '10010002', 操作类型: '删除', roundNo: 1, 原开始日期: '20260802', 原开始时间: '07:00' },
  ]);
  m.confirmBatch();
  assert.strictEqual(appState.mergedRecords.length, 4, '一条都不应执行');
  assert.notStrictEqual(appState.rounds[0].status, 'confirmed');
});

// ==================== 导入校验 ====================
section('导入校验');

test('整行空白直接跳过：不进合并表、不计入校验失败', () => {
  resetState();
  m.processGroupWorkbook({ fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    ['序号', '工号', '姓名', '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数'],
    [1, '10010001', '张三', 20260801, '15:45', 20260801, '17:35', 1.83],
    ['', '', '', '', '', '', '', ''],
    [2, '10010002', '李四', 20260802, '07:00', 20260802, '15:00', 8],
  ] } });
  assert.strictEqual(appState.mergedRecords.length, 2);
  assert.strictEqual(appState.groupFailures.length, 0, '空行不应被记成校验失败');
});

test('时间越界（25:70 / 24:00）记为校验失败，不进合并表', () => {
  resetState();
  m.processGroupWorkbook({ fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    ['序号', '工号', '姓名', '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数'],
    [1, '10010001', '张三', 20260801, '25:70', 20260801, '18:45', ''],
    [2, '10010002', '李四', 20260802, '24:00', 20260802, '15:00', ''],
  ] } });
  assert.strictEqual(appState.mergedRecords.length, 0);
  assert.strictEqual(appState.groupFailures.length, 2);
  assert.ok(appState.groupFailures.every(f => f['失败原因'].includes('时间格式错误')));
});

test('检测重复填报：同一工号同一天同一开始时间出现多次', () => {
  resetState();
  const rows = [
    ['序号', '工号', '姓名', '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数'],
    [1, '10010001', '张三', 20260801, '15:45', 20260801, '17:35', 1.83],
  ];
  m.processGroupWorkbook({
    fileName: '班组表.xlsx', sheetNames: ['一组', '二组'],
    sheets: { 一组: rows, 二组: rows.map(r => r.slice()) },
  });
  assert.strictEqual(appState.mergedRecords.length, 2);
  assert.strictEqual(appState.groupDuplicates.length, 1, '应检出 1 组重复填报');
  assert.strictEqual(appState.groupDuplicates[0]['条数'], 2);
  assert.strictEqual(appState.groupDuplicates[0]['工号'], '10010001');
  const page = m.renderImport();
  assert.ok(page.includes('重复填报'), '导入页应给出重复填报提示');
  assert.ok(page.includes('exportGroupDuplicates()'), '应提供重复清单下载入口');
});

test('同人同日但开始时间不同不算重复填报', () => {
  resetState();
  m.processGroupWorkbook({ fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    ['序号', '工号', '姓名', '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数'],
    [1, '10010001', '张三', 20260801, '15:45', 20260801, '17:35', 1.83],
    [2, '10010001', '张三', 20260801, '20:00', 20260801, '22:00', 2],
  ] } });
  assert.strictEqual(appState.mergedRecords.length, 2);
  assert.strictEqual(appState.groupDuplicates.length, 0);
});

// ==================== 页面渲染 ====================
section('页面渲染');

test('未导入班组填报表时，异常处理页提示先完成导入与合并', () => {
  resetState();
  assert.ok(m.renderAbnormal().includes('尚未导入班组填报表'));
});

test('未导入文件时页面明确标注「示例数据」', () => {
  resetState();
  assert.ok(m.renderImport().includes('示例数据'));
  assert.ok(m.renderAbnormal().includes('示例数据'));
  assert.ok(m.renderRectify().includes('示例数据'));
  assert.ok(m.renderOutput().includes('示例数据'));
});

test('存在定位异常时整改页与输出页均展示提示与下载入口', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  const issues = [{ 级别: '未定位', 系统序号: 900, 校对ID: 900, 工号: '10010099', 姓名: '查无', 班组: '未知', 操作类型: '删除', 原开始日期: '20260801', 原开始时间: '15:45', 命中数: 0, 候选序号: '', 说明: '合并大表中找不到该条加班记录' }];
  setupRound([{ 系统序号: 900, 工号: '10010099', 姓名: '查无', 班组: '未知', 操作类型: '删除', roundNo: 1, 定位状态: '未定位' }], 'confirmed');
  appState.rounds[0].locateIssues = issues;
  appState.batchConfirmed = true;
  const rectifyPage = m.renderRectify();
  const outputPage = m.renderOutput();
  assert.ok(rectifyPage.includes('未定位'), '整改页应展示定位异常级别');
  assert.ok(rectifyPage.includes('exportLocateIssues()'));
  assert.ok(outputPage.includes('需人工核对'));
  assert.ok(outputPage.includes('exportLocateIssues()'));
});

test('无法识别的处置方式记入清单，不再静默忽略', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  const op = { 系统序号: 45, 工号: '10010001', 姓名: '张三', 操作类型: '删除加班', roundNo: 1, 原开始日期: '20260801', 原开始时间: '15:45' };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '未识别');
  assert.ok(res.issues[0]['说明'].includes('无法识别'));
  assert.strictEqual(op['定位状态'], '未识别');
  assert.strictEqual(appState.mergedRecords.length, 4, '未识别的操作不应改动数据');
});

// ==================== 导出保护 ====================
section('导出保护');

test('没有数据时导出汇总 / 整改表不生成任何文件（不再导出示例数据）', () => {
  resetState();
  captured = [];
  assert.strictEqual(m.exportSystemData('final'), false);
  assert.strictEqual(m.exportSystemData('merged'), false);
  assert.strictEqual(m.exportRectify(), false);
  assert.strictEqual(captured.length, 0, '不应生成任何 sheet');
});

test('有数据时导出合并大表正常生成 sheet', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  captured = [];
  m.exportSystemData('merged');
  assert.strictEqual(captured.length, 2, '2007-加班申请 + 码表');
  assert.strictEqual(captured[0].rows.length, 5, '表头 + 4 条记录');
});

// ==================== 汇总 ====================
console.log(`\n通过 ${passed} 项，失败 ${failures.length} 项`);
if (failures.length) {
  failures.forEach(f => console.log(`  - ${f.name}: ${f.err.message}`));
  process.exit(1);
}

