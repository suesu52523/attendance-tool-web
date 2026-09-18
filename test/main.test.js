// 总装科月度加班自动处理工具 —— 纯函数断言测试
// 运行方式：node test/main.test.js
// 说明：js/main.js 是浏览器脚本（无模块导出），这里把源码包进一个函数并按需注入 XLSX / document /
//      Blob / URL 桩，只测试不依赖真实浏览器环境的逻辑；导出函数用「捕获桩」记录生成的工作表，
//      因此可以断言"没有数据时不会导出示例数据"这类行为。
// 覆盖重点：异常表匹配（姓名不参与定位）、批量操作定位（只用业务键，序号不作身份；未定位/多条命中都不改数据）、
//          重复执行保护、导出兜底保护。

const fs = require('fs');
const path = require('path');
const assert = require('assert');

const MAIN_JS = path.join(__dirname, '..', 'js', 'main.js');
const src = fs.readFileSync(MAIN_JS, 'utf8');

// ---------- 桩 ----------
let captured = [];
const xlsxStub = {
  // 只模拟测试用到的两种格式，行为对齐真实 XLSX：'h:mm'（0.5 → "12:00"，小时不补零）、
  // 'yyyy-mm-dd'（Excel 日期序列号 → 日期，1899-12-30 为 0 基准）。
  // 注意：这两种都不实现的话，相关 bug 会测不出来（假绿）。
  SSF: { format: (fmt, v) => {
    if (fmt === 'h:mm') return `${Math.floor(v * 24)}:${String(Math.round(v * 1440) % 60).padStart(2, '0')}`;
    if (fmt === 'yyyy-mm-dd') return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10);
    return '';
  } },
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
// 浏览器原生确认框的桩：默认「点确定」，用例可临时改成 false 模拟「点取消」
let confirmAnswer = true;
globalThis.window = { confirm: () => confirmAnswer };
globalThis.Blob = class { };
globalThis.URL = { createObjectURL: () => 'blob:x', revokeObjectURL() {} };

const factory = new Function('XLSX', 'document', 'console', `${src}
return {
  appState,
  toYYYYMMDD, normalizeDate, normalizeTime, padTime, parseDateParts, parseTimeParts, computeHours,
  formatCellValue,
  locateMergedRecords, resolveOperationTarget, applyBatchOperations,
  processGroupWorkbook, processAbnormalWorkbook, processRectifyWorkbook, confirmBatch,
  buildSystemRecords, buildShiftRecords,
  renderImport, renderAbnormal, renderRectify, renderOutput,
  renderTable,
  exportSystemData, exportRectify, exportShiftData, exportAbnormalFailures, exportLocateIssues, exportOperationLog,
  tableProblem, escapeHtml, startNewRound,
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

// ==================== 时数列不能被当成钟点（M1 问题 3） ====================
// 来源：docs/排查/M1-问题说明.html 的问题 3。
// 根因：「实际加班时数(未减吃饭时间)」列名里的“时间”来自“未减吃饭时间”，它其实是小时数。
section('时数列不能被当成钟点（M1 问题 3）');

test('时数列里的不足 1 小时不能被当成钟点显示', () => {
  const col = '实际加班时数(未减吃饭时间)';
  assert.strictEqual(m.formatCellValue(0.5, col), 0.5, '0.5 小时不能变成 "12:00"');
  assert.strictEqual(m.formatCellValue(0.25, col), 0.25, '0.25 小时不能变成 "6:00"');
  assert.strictEqual(m.formatCellValue(0.9, col), 0.9, '0.9 小时不能变成 "21:36"');
  assert.strictEqual(m.formatCellValue(1.5, '加班时数'), 1.5, '本来不含「时间」的时数列不受影响');
  assert.strictEqual(m.formatCellValue(8, '上报加班时数'), 8);
});

test('真正的钟点列仍按 h:mm 转换（不能顺手把它关掉）', () => {
  assert.strictEqual(m.formatCellValue(22 / 24, '加班开始时间'), '22:00');
  assert.strictEqual(m.formatCellValue(15.75 / 24, '加班结束时间'), '15:45');
  assert.strictEqual(m.formatCellValue(0, '加班开始时间'), '0:00');
});

test('发给班组的整改表里，实际加班时数那一格是 0.5（不是 12:00）', () => {
  resetState();
  // 异常表要能匹配上，否则会被排除在整改表之外
  appState.mergedRecords = [{
    系统序号: 1, 工号: '00163613', 姓名: '王义', 班组: '车门A组',
    加班开始日期: '2026-08-05', 加班开始时间: '15:45',
    加班结束日期: '2026-08-05', 加班结束时间: '18:45', 加班时数: 3,
  }];
  m.processAbnormalWorkbook({
    fileName: '异常表.xlsx', sheetNames: ['异常记录'],
    sheets: { 异常记录: [
      ['ID', '工号', '姓名', '科室', 'T0日*系统排班', '开始加班打卡时间', '结束加班打卡时间',
        '开始日期', '结束日期', '开始时间', '结束时间', '上报加班时数', '实际加班时数(未减吃饭时间)', '差异', '提醒信息'],
      [45, '00163613', '王义', '车门A组', '双班早班 2026-08-05 07:00:00~2026-08-05 15:45:00', '',
        '2026-08-05 16:56:29', 20260805, 20260805, '15:45', '18:45', 3, 0.5, 2.5, '【加班时数差异】上报与打卡不符;'],
    ] },
  });
  m.exportRectify();
  assert.strictEqual(captured.length, 1, '应导出 1 个班组 sheet');
  const headers = captured[0].rows[0];
  const row = captured[0].rows[1];
  const i = headers.indexOf('实际加班时数(未减吃饭时间)');
  assert.ok(i > -1, '整改表应带上异常表整列');
  assert.strictEqual(row[i], 0.5, '这一格必须是 0.5，不能是 "12:00"');
  assert.strictEqual(row[headers.indexOf('上报加班时数')], 3, '同一行的其它时数不能被动到');
  assert.strictEqual(row[headers.indexOf('开始时间')], '15:45', '同一行的钟点列照旧');
});
// ==================== 不存在的日期（M1 问题 2） ====================
// 来源：docs/排查/M1-问题说明.html 的问题 2。
// 口径：「格式对」不等于「日期存在」：2 月 30 日 / 13 月 / 4 月 31 日 这类必须退回，
//       不能靠 new Date 自动挪到别的日子（2026-02-30 会被挪成 2026-03-02，进 2007 表就成了 20260230）。
section('不存在的日期（M1 问题 2）');

const DATE_ROW_HEADERS = ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间',
  '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'];
function groupWorkbookWithDate(startDate, endDate) {
  return { fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    DATE_ROW_HEADERS,
    [1, '10010001', '张三', '一组', startDate, '08:00', endDate, '17:00', '', '产能爬坡', '工作日', '核准'],
  ] } };
}

test('2 月 30 日被退回，不进合并大表', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDate('2026-02-30', '2026-02-30'));
  assert.strictEqual(appState.mergedRecords.length, 0, '不存在的日期不能进大表');
  assert.strictEqual(appState.groupFailures.length, 1);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('日期无效'), appState.groupFailures[0]['失败原因']);
});

test('8 位写法 20260230 同样被退回', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDate('20260230', '20260230'));
  assert.strictEqual(appState.mergedRecords.length, 0);
  assert.strictEqual(appState.groupFailures.length, 1);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('日期无效'), appState.groupFailures[0]['失败原因']);
});

test('13 月 / 4 月 31 日被退回', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDate('2026-13-01', '2026-13-01'));
  assert.strictEqual(appState.mergedRecords.length, 0, '13 月不能进大表');
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDate('2026-04-31', '2026-04-31'));
  assert.strictEqual(appState.mergedRecords.length, 0, '4 月 31 日不能进大表');
  assert.ok(appState.groupFailures[0]['失败原因'].includes('日期无效'), appState.groupFailures[0]['失败原因']);
});

test('闰年边界：2026-02-29 退回（2026 不是闰年）、2028-02-29 通过', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDate('2026-02-29', '2026-02-29'));
  assert.strictEqual(appState.mergedRecords.length, 0, '2026 不是闰年，2 月 29 日不存在');
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDate('2028-02-29', '2028-02-29'));
  assert.strictEqual(appState.groupFailures.length, 0, '2028 是闰年，2 月 29 日必须照常通过');
  assert.strictEqual(appState.mergedRecords.length, 1);
});

test('整改阶段：修改后日期不存在 → 记一条清单且原记录不变', () => {
  resetState();
  appState.mergedRecords = [{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组',
    加班开始日期: '2026-08-01', 加班开始时间: '08:00',
    加班结束日期: '2026-08-01', 加班结束时间: '17:00', 加班时数: 9,
  }];
  const res = m.applyBatchOperations([{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 操作类型: '修改', roundNo: 1,
    原开始日期: '20260801', 原开始时间: '08:00',
    修改后开始日期: '2026-02-30', 修改后开始时间: '08:00',
    修改后结束日期: '2026-02-30', 修改后结束时间: '17:00',
  }]);
  assert.strictEqual(res.issues.length, 1, '应记一条驳回');
  assert.strictEqual(res.issues[0]['级别'], '日期不合理');
  assert.ok(res.issues[0]['说明'].includes('日期'), res.issues[0]['说明']);
  assert.strictEqual(appState.mergedRecords[0]['加班开始日期'], '2026-08-01', '原记录不能被改');
});
// ==================== 一格写两个时间（M1 问题 4） ====================
// 来源：docs/排查/M1-问题说明.html 的问题 4。
// 口径：一个格子只装一个时刻。"15:00~19:00" / "19:00（次日）" 这类以前会被当成 15:00 放行，
//       脏值一路进合并大表并被原样导出到 2007 表 → 必须退回。
section('一格写两个时间（M1 问题 4）');

const TIME_ROW_HEADERS = ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间',
  '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'];
function groupWorkbookWithTime(startTime, endTime) {
  return { fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    TIME_ROW_HEADERS,
    [1, '10010001', '张三', '一组', '2026-08-01', startTime, '2026-08-01', endTime, '', '产能爬坡', '工作日', '核准'],
  ] } };
}

test('区间写法 15:00~19:00 被退回（以前会被当成 15:00 放行）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithTime('15:00~19:00', '19:00'));
  assert.strictEqual(appState.mergedRecords.length, 0, '区间写法不能进大表');
  assert.strictEqual(appState.groupFailures.length, 1);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('加班开始时间格式错误'), appState.groupFailures[0]['失败原因']);
});

test('带备注的时间 19:00（次日）被退回', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithTime('19:00（次日）', '19:00'));
  assert.strictEqual(appState.mergedRecords.length, 0, '带备注的时间不能进大表');
  assert.ok(appState.groupFailures[0]['失败原因'].includes('加班开始时间格式错误'));
});

test('中文连接词 15:00 至 19:00 被退回', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithTime('15:00', '15:00 至 19:00'));
  assert.strictEqual(appState.mergedRecords.length, 0);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('加班结束时间格式错误'));
});

test('带秒 15:00:30 照常通过（沿用上游既有口径，与 padTime 一致）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithTime('15:00:30', '19:00'));
  assert.strictEqual(appState.groupFailures.length, 0, '带秒是合法的单个时刻，不能误拦');
  assert.strictEqual(appState.mergedRecords.length, 1);
});

test('单个时刻照常通过：8:15 / 08:15（回归）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithTime('8:15', '18:15'));
  assert.strictEqual(appState.groupFailures.length, 0, '单个时刻不能被误拦');
  assert.strictEqual(appState.mergedRecords.length, 1);
  resetState();
  m.processGroupWorkbook(groupWorkbookWithTime('08:15', '18:15'));
  assert.strictEqual(appState.mergedRecords.length, 1, '08:15 同样要放行');
});

test('整改阶段：修改后时间是区间 → 记一条清单且原记录不变', () => {
  resetState();
  appState.mergedRecords = [{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组',
    加班开始日期: '2026-08-01', 加班开始时间: '15:45',
    加班结束日期: '2026-08-01', 加班结束时间: '18:45', 加班时数: 3,
  }];
  const res = m.applyBatchOperations([{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 操作类型: '修改', roundNo: 1,
    原开始日期: '20260801', 原开始时间: '15:45',
    修改后开始日期: '2026-08-01', 修改后开始时间: '15:00~19:00',
    修改后结束日期: '2026-08-01', 修改后结束时间: '19:00',
  }]);
  assert.strictEqual(res.issues.length, 1, '应记一条驳回');
  assert.strictEqual(res.issues[0]['级别'], '时间不合理');
  assert.ok(res.issues[0]['说明'].includes('时刻'), res.issues[0]['说明']);
  assert.strictEqual(appState.mergedRecords[0]['加班开始时间'], '15:45', '原记录不能被改');
});
// ==================== 结束早于开始（M1 问题 1，业务口径 B：拦下退回） ====================
// 来源：docs/排查/M1-问题说明.html 的问题 1。
// 口径：同一天内结束时间早于开始时间 = 填错（夜班常忘了把结束日期改成次日）→ 退回班组核对；
//       不按“跨天”静默算成 4 小时，也不允许负数进合并大表。
section('结束早于开始（M1 问题 1）');

const GROUP_ROW_HEADERS = ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间',
  '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'];

test('夜班同一天 22:00→02:00 被退回，不进合并大表', () => {
  resetState();
  m.processGroupWorkbook({ fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    GROUP_ROW_HEADERS,
    [1, '10010001', '张三', '一组', '2026-08-01', '22:00', '2026-08-01', '02:00', '', '夜班生产', '工作日', '核准'],
  ] } });
  assert.strictEqual(appState.mergedRecords.length, 0, '结束早于开始的行不能进合并大表');
  assert.strictEqual(appState.groupFailures.length, 1);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('早于开始'), appState.groupFailures[0]['失败原因']);
});

test('Excel 数值形态的同日夜班同样被退回（真实 xlsx 里日期时间就是数字）', () => {
  resetState();
  const D = 46235; // 2026-08-01
  m.processGroupWorkbook({ fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    GROUP_ROW_HEADERS,
    [1, '10010001', '张三', '一组', D, 22 / 24, D, 2 / 24, '', '夜班生产', '工作日', '核准'],
  ] } });
  assert.strictEqual(appState.mergedRecords.length, 0, '真实数字形态也不能进合并大表');
  assert.strictEqual(appState.groupFailures.length, 1);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('早于开始'), appState.groupFailures[0]['失败原因']);
});

test('正常跨天不受影响：结束日期填次日仍算 4 小时', () => {
  resetState();
  m.processGroupWorkbook({ fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    GROUP_ROW_HEADERS,
    [1, '10010001', '张三', '一组', '2026-08-01', '22:00', '2026-08-02', '02:00', '', '夜班生产', '工作日', '核准'],
  ] } });
  assert.strictEqual(appState.groupFailures.length, 0, '正常跨天不能被误拦');
  assert.strictEqual(appState.mergedRecords.length, 1);
  assert.strictEqual(appState.mergedRecords[0]['加班时数'], 4);
});

test('整改阶段：修改后结束早于开始被驳回，不改动原记录', () => {
  resetState();
  appState.mergedRecords = [{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '一组',
    加班开始日期: '2026-08-01', 加班开始时间: '22:00',
    加班结束日期: '2026-08-02', 加班结束时间: '02:00', 加班时数: 4,
  }];
  const res = m.applyBatchOperations([{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '一组', 操作类型: '修改', roundNo: 1,
    原开始日期: '20260801', 原开始时间: '22:00',
    修改后开始日期: '2026-08-01', 修改后开始时间: '22:00',
    修改后结束日期: '2026-08-01', 修改后结束时间: '02:00',
  }]);
  assert.strictEqual(res.issues.length, 1, '应记一条驳回');
  assert.strictEqual(res.issues[0]['级别'], '时间不合理');
  assert.ok(res.issues[0]['说明'].includes('早于开始'), res.issues[0]['说明']);
  const rec = appState.mergedRecords[0];
  assert.strictEqual(rec['加班结束日期'], '2026-08-02', '原记录不能被改');
  assert.strictEqual(rec['加班时数'], 4, '原时数不能被改');
});
// ==================== 加班时数上限（M1 问题 5） ====================
// 业务口径（2026-09-17）：单条加班时数上限 48 小时，超过的退回班组核对。
// 背景：以前没有任何上下界，49 小时这种“跨天算错 / 多打一位”的值能直接进大表并导出。
section('加班时数上限（M1 问题 5）');

const LIMIT_HEADERS = ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间',
  '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'];
function groupWorkbookWithHours(date, startTime, endTime, hours) {
  return { fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    LIMIT_HEADERS,
    [1, '10010001', '张三', '一组', date, startTime, date, endTime, hours, '产能爬坡', '工作日', '核准'],
  ] } };
}

test('填 49 小时被退回（超过 48 上限）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '08:00', '17:00', 49));
  assert.strictEqual(appState.mergedRecords.length, 0, '超上限的行不能进大表');
  assert.strictEqual(appState.groupFailures.length, 1);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('48'), appState.groupFailures[0]['失败原因']);
});

test('填 48 小时照常通过（边界值）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '08:00', '17:00', 48));
  assert.strictEqual(appState.groupFailures.length, 0, '48 小时是边界内，不能误拦');
  assert.strictEqual(appState.mergedRecords.length, 1);
  assert.strictEqual(appState.mergedRecords[0]['加班时数'], 48);
});

test('时数留空但算出来超 48 小时，同样退回', () => {
  resetState();
  // 2026-08-01 08:00 → 2026-08-03 09:00 = 49 小时
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '08:00', '09:00', ''));
  appState.mergedRecords = [];
  resetState();
  const wb = { fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    LIMIT_HEADERS,
    [1, '10010001', '张三', '一组', '2026-08-01', '08:00', '2026-08-03', '09:00', '', '产能爬坡', '工作日', '核准'],
  ] } };
  m.processGroupWorkbook(wb);
  assert.strictEqual(appState.mergedRecords.length, 0, '算出来的 49 小时也不能进大表');
  assert.ok(appState.groupFailures[0]['失败原因'].includes('48'), appState.groupFailures[0]['失败原因']);
});

test('样本里的正常取值（8 小时）不受影响', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '08:00', '17:00', 8));
  assert.strictEqual(appState.groupFailures.length, 0);
  assert.strictEqual(appState.mergedRecords[0]['加班时数'], 8);
});

test('整改阶段：修改后时数超上限 → 记一条清单且原记录不变', () => {
  resetState();
  appState.mergedRecords = [{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组',
    加班开始日期: '2026-08-01', 加班开始时间: '08:00',
    加班结束日期: '2026-08-01', 加班结束时间: '17:00', 加班时数: 9,
  }];
  const res = m.applyBatchOperations([{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 操作类型: '修改', roundNo: 1,
    原开始日期: '20260801', 原开始时间: '08:00',
    修改后开始日期: '2026-08-01', 修改后开始时间: '08:00',
    修改后结束日期: '2026-08-01', 修改后结束时间: '17:00',
    修改后上报加班时数: 49,
  }]);
  assert.strictEqual(res.issues.length, 1, '应记一条驳回');
  assert.strictEqual(res.issues[0]['级别'], '时数不合理');
  assert.ok(res.issues[0]['说明'].includes('48'), res.issues[0]['说明']);
  assert.strictEqual(appState.mergedRecords[0]['加班时数'], 9, '原记录不能被改');
});
// ==================== 中文日期（M1-6） ====================
// 业务口径（2026-09-17）：中文日期认下来，往全了做。
// 真实来源：样本 02-班组填报-校验失败.xlsx / 内装A1组 行8 写的就是 "2026年8月10日"，
//          过去一律退回让班组重填（手工誊写到系统里平白多一步）。
// 口径：只认“年月日 / 年月号”这一种；日期不存在的（2026年2月30日）仍然退回。
section('中文日期（M1-6）');

const CN_HEADERS = ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间',
  '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'];
function groupWorkbookWithDateText(startDate, endDate) {
  return { fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    CN_HEADERS,
    [1, '10010001', '张三', '一组', startDate, '15:45', endDate, '18:45', '', '产能爬坡', '工作日', '核准'],
  ] } };
}

test('中文日期 2026年8月10日 被认下，并统一成 2026-08-10', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDateText('2026年8月10日', '2026年8月10日'));
  assert.strictEqual(appState.groupFailures.length, 0, '中文日期不该再被退回');
  assert.strictEqual(appState.mergedRecords.length, 1);
  assert.strictEqual(appState.mergedRecords[0]['加班开始日期'], '2026-08-10', '大表里要统一成短横线写法');
  assert.strictEqual(appState.mergedRecords[0]['加班时数'], 3, '算工时也要能算出来');
  assert.strictEqual(m.buildSystemRecords(appState.mergedRecords)[0]['开始日期'], '20260810', '导出的 2007 表要是 8 位');
});

test('中文日期的各种写法（往全了做）', () => {
  const forms = ['2026年08月10日', '2026年8月10号', '2026 年 8 月 10 日', '２０２６年８月１０日', '2026年8月10日 '];
  for (const f of forms) {
    resetState();
    m.processGroupWorkbook(groupWorkbookWithDateText(f, f));
    assert.strictEqual(appState.mergedRecords.length, 1, `${f} 应该被认下`);
    assert.strictEqual(appState.mergedRecords[0]['加班开始日期'], '2026-08-10', `${f} 应该归一`);
  }
});

test('中文写法但日期不存在（2026年2月30日）仍要退回', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithDateText('2026年2月30日', '2026年2月30日'));
  assert.strictEqual(appState.mergedRecords.length, 0, '不能把 2 月 30 日洗成看起来合法的日期');
  assert.strictEqual(appState.groupFailures.length, 1);
});

test('既有写法行为不变（短横线 / 8 位 / 斜线）', () => {
  for (const f of ['2026-08-10', '20260810', '2026/8/10']) {
    resetState();
    m.processGroupWorkbook(groupWorkbookWithDateText(f, f));
    assert.strictEqual(appState.mergedRecords.length, 1, `${f} 应该照旧放行`);
    assert.strictEqual(appState.mergedRecords[0]['加班开始日期'], f, `${f} 不该被改动`);
    assert.strictEqual(m.buildSystemRecords(appState.mergedRecords)[0]['开始日期'], '20260810');
  }
});

test('整改表里的中文日期同样归一（端到端）', () => {
  resetState();
  m.processRectifyWorkbook(rectifyParsed([[1, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '18:45', 3,
    '修改', '2026年8月10日', '15:45', '2026年8月10日', '18:45', '', '', '', '']]));
  assert.strictEqual(appState.rectifyOperations[0]['修改后开始日期'], '2026-08-10', '导入时就该归一');
  appState.mergedRecords = [{
    系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组',
    加班开始日期: '2026-08-01', 加班开始时间: '15:45',
    加班结束日期: '2026-08-01', 加班结束时间: '18:45', 加班时数: 3,
  }];
  const res = m.applyBatchOperations(appState.rectifyOperations);
  assert.strictEqual(res.issues.length, 0, '中文日期不该被当成异常');
  assert.strictEqual(appState.mergedRecords[0]['加班开始日期'], '2026-08-10');
  assert.strictEqual(m.buildSystemRecords(appState.mergedRecords)[0]['开始日期'], '20260810');
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
  assert.ok(!('系统序号' in rec), '异常记录不再带「系统序号」字段：序号不是身份，也不代表已处理');
  assert.strictEqual(round.abnormalFailures.length, 0);
  assert.strictEqual(round.abnormalWarnings.length, 1);
  assert.ok(round.abnormalWarnings[0]['定位提醒'].includes('不一致'));
});

test('异常表姓名为空时仍能匹配成功', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010002', '', '底盘一组', 20260802, '07:00', 20260802, '15:00', 8]]));
  assert.strictEqual(appState.abnormalRecords[0]['匹配状态'], '已匹配');
  assert.strictEqual(appState.rounds[0].abnormalWarnings.length, 0, '姓名空不产生不一致提醒');
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
  assert.ok(!round.abnormalWarnings[0]['定位提醒'].includes('已暂按'), '不能声称“已按序号处理”：多个候选时工具不会自己挑一条');
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

// 空行（Excel 尾部常见）：班组表那条路早就跳过，异常表/整改表两条路必须同口径
// 否则：整改表的空行会被当成「未填写」→ confirmBatch 整轮阻断，提示让人去补一条根本不存在的记录
test('整改表尾部空行不算「未填写」：否则整轮被假阻断', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三', '底盘一组', 20260801, '15:45', 20260801, '17:35', 1.83]]));
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010002', '李四', '底盘一组', 20260802, '07:00', 20260802, '15:00', 8, '删除', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''], // 尾部空行（真实样本 13 号表里就有）
  ]));
  assert.strictEqual(appState.rectifyOperations.length, 1, '空行不该产生操作记录');
  assert.strictEqual(appState.rectifyOperations.filter(o => o['操作类型'] === '未填写').length, 0, '空行不算「未填写」');
  m.confirmBatch();
  assert.ok(!appState.mergedRecords.some(r => r['工号'] === '10010002'), '没被空行挡住，那一行照常执行');
});

test('异常表尾部空行不算记录，也不进「未匹配」清单', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([
    [45, '10010002', '李四', '底盘一组', '20260802', '07:00', '20260802', '15:00', 8],
    ['', '', '', '', '', '', '', '', ''],
  ]));
  assert.strictEqual(appState.abnormalRecords.length, 1, '空行不算数据行');
  assert.strictEqual(appState.abnormalFailures.length, 0, '空行不该进匹配失败清单');
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

test('序号对不上人时不许动手：工号查无此人 → 进未定位清单，不动任何记录', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  // 整改表里工号 99999999 在大表中不存在，但序号 2 恰好是大表里的李四
  // 口径（2026-09-17）：内部序号每次导入都从 1 重发、删除后还会重排（见 processGroupWorkbook / 删除分支），
  // 拿它当身份会指到别人身上 —— 定位不到就进清单让人核对，不靠序号猜
  const op = { 系统序号: 2, 工号: '99999999', 姓名: '未知', 操作类型: '删除', roundNo: 1 };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1, '必须记入定位异常清单');
  assert.strictEqual(res.issues[0]['级别'], '未定位');
  assert.strictEqual(res.issues[0]['操作类型'], '删除');
  assert.strictEqual(op['定位状态'], '未定位');
  assert.strictEqual(appState.mergedRecords.length, 4, '不许删掉序号 2 那位（他不是本行要动的人）');
});

test('业务键缺日期时也不猜序号：仍然进清单，不动任何记录', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  // 工号、日期都没有可用的信息，只剩一个序号 —— 这种情况历史上“删错过人”
  const op = { 系统序号: 3, 工号: '', 姓名: '王五', 操作类型: '删除', roundNo: 1 };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '未定位');
  assert.strictEqual(appState.mergedRecords.length, 4, '不应误删任何记录');
});

test('多条命中时列出候选序号，但不许自己挑一条动手（一条都不动）', () => {
  resetState();
  appState.mergedRecords = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '15:45', 加班时数: 1.83 },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '20:00', 加班时数: 2 },
  ];
  const op = { 系统序号: 1, 工号: '10010001', 姓名: '张三', 操作类型: '删除', roundNo: 1, 原开始日期: '20260801', 原开始时间: '' };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '多条命中');
  assert.strictEqual(res.issues[0]['候选序号'], '1、2', '仍要列出候选，方便人工核对');
  assert.strictEqual(op['定位状态'], '多条命中(2)');
  assert.strictEqual(appState.mergedRecords.length, 2, '挑不准就不许动手：两条都得留着');
  assert.ok(res.issues[0]['说明'].includes('未执行'), '说明要写明“没执行”');
  assert.ok(res.issues[0]['说明'].includes('原开始时间'), '并告诉人怎么消除歧义（补填原开始时间）');
});

test('多条命中（修改）：不执行修改，两条记录一个格都不许变', () => {
  resetState();
  appState.mergedRecords = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '15:45', 加班结束日期: '2026-08-01', 加班结束时间: '18:45', 加班时数: 3 },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '20:00', 加班结束日期: '2026-08-01', 加班结束时间: '22:00', 加班时数: 2 },
  ];
  const before = JSON.stringify(appState.mergedRecords);
  const op = { 工号: '10010001', 姓名: '张三', 操作类型: '修改', roundNo: 1, 原开始日期: '20260801', 原开始时间: '',
    修改后开始日期: '2026-08-01', 修改后开始时间: '18:00', 修改后结束日期: '2026-08-01', 修改后结束时间: '21:00', 修改后上报加班时数: 3 };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '多条命中');
  assert.strictEqual(op['定位状态'], '多条命中(2)');
  assert.strictEqual(JSON.stringify(appState.mergedRecords), before, '不改任何一条（当天合计也不能变）');
});

test('多条命中（调班）：同样不执行', () => {
  resetState();
  appState.mergedRecords = [
    { 系统序号: 1, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '15:45', 加班时数: 1.83 },
    { 系统序号: 2, 工号: '10010001', 姓名: '张三', 加班开始日期: '2026-08-01', 加班开始时间: '20:00', 加班时数: 2 },
  ];
  const op = { 工号: '10010001', 姓名: '张三', 操作类型: '调班', roundNo: 1, 原开始日期: '20260801', 原开始时间: '', 调班日期: '20260801', 调班班次: 'OFF' };
  const res = m.applyBatchOperations([op]);
  assert.strictEqual(res.issues.length, 1);
  assert.strictEqual(res.issues[0]['级别'], '多条命中');
  assert.strictEqual(appState.mergedRecords.length, 2, '调班也不能挑一条动手');
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
  // abnormalRecords 放一条占位：模拟「已完成第 2 步（导过异常表）」，M4 的守卫要求本轮导过异常表
  appState.rounds = [{ roundNo: 1, abnormalRecords: [{ ID: 1 }], abnormalFailures: [], abnormalWarnings: [], rectifyOperations: ops, shiftRecords: [], systemRecords: [], locateIssues: [], status }];
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

// ==================== M2-2：序号归位（校对单号 ≠ 大表序号） ====================
section('校对单号 vs 大表序号');

// 兑一份异常表：一条能对上（李四 8-02 07:00 → 大表序号 2），一条工号查无此人
test('导出的整改表：ID 列写校对系统单号，不写大表序号', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([
    [880001, '10010002', '李四', '底盘一组', '20260802', '07:00', '20260802', '15:00', 8],
  ]));
  captured = [];
  m.exportRectify();
  const sheet = captured.find(s => s.name.includes('底盘'));
  assert.ok(sheet, '应导出底盘一组的 sheet');
  const idIdx = sheet.rows[0].indexOf('ID');
  assert.strictEqual(String(sheet.rows[1][idIdx]), '880001', 'ID 列必须还是校对系统给的 880001（旧代码这里写的是大表序号 2）');
});

test('整改表往返：定位到的操作，校对ID 仍是校对单号，系统序号 是实际改的那行', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([
    [880001, '10010002', '李四', '底盘一组', '20260802', '07:00', '20260802', '15:00', 8],
  ]));
  captured = [];
  m.exportRectify();
  const sheet = captured.find(s => s.name.includes('底盘'));
  const rows = sheet.rows.map(r => r.slice()); // 含真实表头，原样回传
  rows[1][rows[0].indexOf('处置方式')] = '删除';
  m.processRectifyWorkbook({ fileName: '整改表回传.xlsx', sheetNames: ['S1'], sheets: { S1: rows } });
  const res = m.applyBatchOperations(appState.rectifyOperations);
  const op = appState.rectifyOperations[0];
  assert.deepStrictEqual(res.issues, [], '能唯一命中就不应有清单');
  assert.strictEqual(op['定位状态'], '已定位(业务键)');
  assert.strictEqual(String(op['校对ID']), '880001', '操作对象要带着校对单号，才能和校对系统对账');
  assert.ok(!appState.mergedRecords.some(r => r['工号'] === '10010002'), '按业务键删掉李四那条');
});

test('定位异常清单：校对ID=校对单号；系统序号（执行前）只在实际有目标行时填，定位不到就留空', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  // 手工填的整改表（或班组改过工号）：ID 是校对单号 901，但大表里没有 99999999 这个人
  m.processRectifyWorkbook(rectifyParsed([
    [901, '99999999', '徐阳', '底盘一组', '20260801', '15:45', '20260801', '17:35', 2, '删除'],
    [902, '10010003', '王五', '底盘一组', '20260802', '', '20260802', '17:35', 1.83, '删除'],
  ]));
  const res = m.applyBatchOperations(appState.rectifyOperations);
  const miss = res.issues.find(i => i['工号'] === '99999999');
  assert.ok(miss, '查无此人 → 进清单');
  assert.strictEqual(String(miss['校对ID']), '901', '校对ID 列应是校对单号');
  assert.strictEqual(miss['系统序号（执行前）'], '', '未定位就没有大表行号，不能拿校对单号冒充');
  const multi = res.issues.find(i => i['工号'] === '10010003');
  if (multi) {
    assert.strictEqual(String(multi['校对ID']), '902');
    assert.strictEqual(multi['系统序号（执行前）'], 2, '多条命中时，编号应是实际作用的那一行（执行前编号）');
  }
});

test('操作执行记录：列名叫校对ID，写的是校对单号（不再冒充大表序号）', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processRectifyWorkbook(rectifyParsed([
    [901, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '删除'],
  ]));
  m.applyBatchOperations(appState.rectifyOperations);
  captured = [];
  m.exportOperationLog();
  assert.deepStrictEqual(captured[0].rows[0], ['轮次', '校对ID', '工号', '姓名', '班组', '操作类型', '定位状态', '操作详情', '备注']);
  assert.strictEqual(String(captured[0].rows[1][1]), '901');
  assert.strictEqual(captured[0].rows[1][6], '已定位(业务键)', '定位状态列要写清这条到底执行没执行');
});

test('操作执行记录：没执行的操作也必须如实标出来（不能看起来像做了）', () => {
  resetState();
  const op = { 校对ID: 5, 工号: '10010099', 姓名: '查无', 班组: '未知', 操作类型: '删除', roundNo: 1, 原开始日期: '20260801', 原开始时间: '15:45', 定位状态: '未定位' };
  appState.rounds = [{ roundNo: 1, rectifyOperations: [op] }];
  captured = [];
  m.exportOperationLog();
  assert.strictEqual(captured[0].rows[1][6], '未定位', '没执行的必须写明未定位，不能让人以为删过了');
});

test('业务键定位仍然优先于任何号码：校对单号撞上别人也不改错人', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  // 校对单号 2 恰好等于大表里张三的序号，但业务键指向王五
  m.processRectifyWorkbook(rectifyParsed([
    [2, '10010003', '王五', '底盘一组', '20260802', '15:45', '20260802', '17:35', 1.83, '删除'],
  ]));
  const res = m.applyBatchOperations(appState.rectifyOperations);
  assert.deepStrictEqual(res.issues, []);
  assert.ok(!appState.mergedRecords.some(r => r['工号'] === '10010003'), '删的是王五');
  assert.ok(appState.mergedRecords.some(r => r['工号'] === '10010001'), '不能误删张三');
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

// ==================== M3 导入校验补强（2026-09-18） ====================
section('M3 导入校验补强：时数合理性 / 表头变化提示 / 匹配失败线索');

test('加班时数填成文字「半小时」被退回，不进大表', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '15:45', '17:55', '半小时'));
  assert.strictEqual(appState.mergedRecords.length, 0, '文字时数不能进大表');
  assert.strictEqual(appState.groupFailures.length, 1);
  assert.ok(appState.groupFailures[0]['失败原因'].includes('不是数字'), appState.groupFailures[0]['失败原因']);
});

test('加班时数 0 或负数被退回', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '08:00', '10:00', 0));
  assert.strictEqual(appState.mergedRecords.length, 0, '0 小时无意义');
  assert.ok(appState.groupFailures[0]['失败原因'].includes('必须大于 0'), appState.groupFailures[0]['失败原因']);
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '08:00', '10:00', -1));
  assert.strictEqual(appState.mergedRecords.length, 0, '负数不能进大表');
  assert.ok(appState.groupFailures[0]['失败原因'].includes('必须大于 0'), appState.groupFailures[0]['失败原因']);
});

test('文本形式的数字「2.5」照常放行（不误伤 Excel 文本单元格）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '15:45', '18:15', '2.5'));
  assert.strictEqual(appState.mergedRecords.length, 1);
  assert.strictEqual(appState.groupFailures.length, 0);
});

test('整改表「修改后时数」填成文字时不执行，进定位异常清单', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processRectifyWorkbook(rectifyParsed([
    [1, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '修改',
      '20260801', '18:00', '20260801', '21:00', '半小时'],
  ]));
  const res = m.applyBatchOperations(appState.rectifyOperations);
  assert.strictEqual(res.issues.length, 1, '应进定位异常清单');
  assert.strictEqual(appState.rectifyOperations[0]['定位状态'], '时数不合理');
  const zhangsan = appState.mergedRecords.find(r => r['工号'] === '10010001');
  assert.strictEqual(zhangsan['加班时数'], 1.83, '时数不合理时不能动大表');
});

test('班组表列名被改（缺列 + 多列）会提示，不再静默丢列', () => {
  resetState();
  const headers = [...GROUP_ROW_HEADERS.slice(0, 11), '领导审核批准'];
  m.processGroupWorkbook({ fileName: '班组表.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    headers,
    [1, '10010001', '张三', '一组', '2026-08-01', '15:45', '2026-08-01', '18:45', 3, '产能爬坡', '工作日', '核准'],
  ] } });
  const sheet = appState.groupSheets[0];
  assert.strictEqual(sheet.status, 'warning', '列名不认识必须提示');
  assert.ok(sheet.note.includes('科负责人核准'), sheet.note);
  assert.ok(sheet.note.includes('领导审核批准'), sheet.note);
  assert.strictEqual(appState.mergedRecords.length, 1, '行本身照常合并');
});

test('标准表头的 sheet 仍是「正常」（不误报）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '15:45', '17:55', 2));
  assert.strictEqual(appState.groupSheets[0].status, 'ok', appState.groupSheets[0].note);
  assert.strictEqual(appState.groupSheets[0].note, '');
});

test('匹配失败给出线索：同一天大表里的记录 + 工时不一致', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([
    [7, '10010001', '张三', '底盘一组', 20260801, '07:00', 20260801, '11:00', 4],
  ]));
  const f = appState.abnormalFailures[0];
  assert.ok(f, '应当有匹配失败记录');
  assert.ok(f['线索'].includes('同一天大表里有'), f['线索']);
  assert.ok(f['线索'].includes('不一致'), f['线索']);
});

test('匹配失败且当天无记录时，线索指向最近一次', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([
    [8, '10010001', '张三', '底盘一组', 20260805, '15:45', 20260805, '17:35', 1.83],
  ]));
  const f = appState.abnormalFailures[0];
  assert.ok(f['线索'].includes('当天大表里没有'), f['线索']);
  assert.ok(f['线索'].includes('最近一次'), f['线索']);
});

test('匹配失败记录导出的表里带「线索」一列且非空', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([
    [9, '10010001', '张三', '底盘一组', 20260805, '15:45', 20260805, '17:35', 1.83],
  ]));
  captured = [];
  m.exportAbnormalFailures();
  const headers = captured[0].rows[0];
  assert.ok(headers.includes('线索'), JSON.stringify(headers));
  assert.ok(captured[0].rows[1][headers.indexOf('线索')], '线索不能是空的');
});

// ==================== M4：没执行的调班不能外发（2026-09-18） ====================
section('M4 改动执行：调班数据只收「真执行过」的操作');

test('未定位的调班不进调班数据，但要进定位异常清单', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三', '底盘一组', 20260801, '15:45', 20260801, '17:35', 1.83]]));
  m.processRectifyWorkbook(rectifyParsed([
    [11, '10010001', '张三', '底盘一组', '20260809', '15:45', '20260809', '17:35', 1.83, '不处理',
      '', '', '', '', '', '20260809', 'SF04 双班早班', '已发调班表'],
  ]));
  m.confirmBatch();
  const round = appState.rounds[appState.currentRound];
  assert.strictEqual(round.shiftRecords.length, 0, '没执行的调班不能进调班数据（否则排班与工资打架）');
  assert.strictEqual(round.locateIssues.length, 1, '未定位必须进定位异常清单');
  assert.ok(appState.mergedRecords.some(r => r['工号'] === '10010001'), '没定位到就不能动大表');
});

test('已定位的调班照常进调班数据，并把那笔加班从大表拿走', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三', '底盘一组', 20260801, '15:45', 20260801, '17:35', 1.83]]));
  m.processRectifyWorkbook(rectifyParsed([
    [12, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '不处理',
      '', '', '', '', '', '20260801', 'SF04 双班早班', '已发调班表'],
  ]));
  m.confirmBatch();
  const round = appState.rounds[appState.currentRound];
  assert.strictEqual(round.shiftRecords.length, 1, '正常调班必须还能出数据');
  assert.strictEqual(round.shiftRecords[0]['日工作计划'], 'SF04 双班早班');
  assert.strictEqual(round.shiftRecords[0]['开始日期'], '20260801');
  assert.ok(!appState.mergedRecords.some(r => r['工号'] === '10010001'), '调班要把那笔加班从大表拿走');
});

// ==================== M4：没导异常表不能执行（2026-09-18） ====================
section('M4 前置守卫：整改表必须来自本轮的异常表');

test('跳过异常表直接执行会被拦住（拿错/拿旧整改表不该动大表）', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '15:45', '17:55', 2));
  const before = appState.mergedRecords.length;
  m.processRectifyWorkbook(rectifyParsed([
    [21, '10010001', '张三', '一组', '20260801', '15:45', '20260801', '17:55', 2, '删除',
      '', '', '', '', '', '', '', '重复填报'],
  ]));
  m.confirmBatch();
  assert.strictEqual(appState.mergedRecords.length, before, '本轮没导异常表，不允许动大表');
  assert.notStrictEqual(appState.rounds[appState.currentRound].status, 'confirmed', '不能标记为已执行');
});

test('按流程走（异常表→整改表）照常执行', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '15:45', '17:55', 2));
  m.processAbnormalWorkbook(abnormalParsed([
    [45, '10010001', '张三', '一组', 20260801, '15:45', 20260801, '17:55', 2],
  ]));
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010001', '张三', '一组', '20260801', '15:45', '20260801', '17:55', 2, '删除',
      '', '', '', '', '', '', '', '重复填报'],
  ]));
  m.confirmBatch();
  assert.strictEqual(appState.mergedRecords.length, 0, '正常流程照常执行');
  assert.strictEqual(appState.rounds[appState.currentRound].status, 'confirmed');
});

// ==================== M4 边界：同一笔写了多种处置（只提示，不改执行） ====================
section('M4 边界约束：同一笔两种处置要提前提示');

test('同一笔写了「修改 + 删除」时，整改页提前提示（执行顺序不变）', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三', '底盘一组', 20260801, '15:45', 20260801, '17:35', 1.83]]));
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '修改',
      '20260801', '18:00', '20260801', '21:00', 3, '', '', '改时间'],
    [46, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '删除',
      '', '', '', '', '', '', '', '重复填报'],
  ]));
  const html = m.renderRectify();
  assert.ok(html.includes('两种以上处置'), '要提前告诉人同一笔被写了两种处置');
  assert.ok(!html.includes('系统会自动按 ID 定位'), '不能再说按 ID 定位（M2 起已改按业务键）');
  // 执行顺序与结果不变：先修改生效，删除因为原时刻找不到人而进清单
  m.confirmBatch();
  const op = appState.rectifyOperations;
  assert.strictEqual(op[0]['定位状态'], '已定位(业务键)');
  assert.strictEqual(op[1]['定位状态'], '未定位');
  assert.strictEqual(appState.rounds[appState.currentRound].locateIssues.length, 1);
});

test('正常整改表（一笔一种处置）不显示这条提示', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三', '底盘一组', 20260801, '15:45', 20260801, '17:35', 1.83]]));
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '删除',
      '', '', '', '', '', '', '', '重复填报'],
  ]));
  assert.ok(!m.renderRectify().includes('两种以上处置'));
});

// ==================== M4-4 / M4-5（2026-09-18） ====================
section('M4-4 删除前确认 · M4-5 清单编号标注为「执行前」');

test('M4-4 删除前先确认：点「取消」什么都不做，再点并确认才执行', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三', '底盘一组', 20260801, '15:45', 20260801, '17:35', 1.83]]));
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '删除', '', '', '', '', '', '', '', '重复填报'],
  ]));
  const before = appState.mergedRecords.length;
  confirmAnswer = false;
  m.confirmBatch();
  assert.strictEqual(appState.mergedRecords.length, before, '点取消后一行都不能删');
  assert.notStrictEqual(appState.rounds[appState.currentRound].status, 'confirmed', '取消后不能标记已执行');
  confirmAnswer = true;
  m.confirmBatch();
  assert.strictEqual(appState.mergedRecords.length, before - 1, '确认后照常执行');
});

test('M4-4 只有修改（没有删除/调班）时不弹确认框', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  m.processAbnormalWorkbook(abnormalParsed([[45, '10010001', '张三', '底盘一组', 20260801, '15:45', 20260801, '17:35', 1.83]]));
  m.processRectifyWorkbook(rectifyParsed([
    [45, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '17:35', 1.83, '修改',
      '20260801', '18:00', '20260801', '21:00', 3, '', '', '改时间'],
  ]));
  let asked = 0;
  const oldConfirm = globalThis.window.confirm;
  globalThis.window.confirm = () => { asked++; return true; };
  m.confirmBatch();
  globalThis.window.confirm = oldConfirm;
  assert.strictEqual(asked, 0, '不删东西就别打扰人');
  assert.strictEqual(appState.mergedRecords.find(r => r['工号'] === '10010001')['加班时数'], 3, '修改照常生效');
});

test('M4-5 定位异常清单里，编号标成「执行前」并保留校对ID', () => {
  resetState();
  m.processGroupWorkbook({ fileName: 'g.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'],
    [1, '1001', '张三', '一组', '2026-08-01', '15:45', '2026-08-01', '18:45', 3, 'x', '工作日', '核准'],
    [2, '1002', '李四', '一组', '2026-08-02', '15:45', '2026-08-02', '18:45', 3, 'x', '工作日', '核准'],
    [3, '1002', '李四', '一组', '2026-08-02', '15:45', '2026-08-02', '19:45', 4, 'x', '工作日', '核准'],
  ] } });
  m.processAbnormalWorkbook(abnormalParsed([
    [20, '1002', '李四', '一组', 20260802, '15:45', 20260802, '18:45', 3],
    [21, '1001', '张三', '一组', 20260801, '15:45', 20260801, '18:45', 3],
  ]));
  // 一条多条命中（挑不准，未执行）+ 一条删除（会重排全表编号）
  m.processRectifyWorkbook(rectifyParsed([
    [20, '1002', '李四', '一组', '20260802', '15:45', '20260802', '18:45', 3, '修改', '20260802', '19:00', '20260802', '20:00', 1, '', '', '多条记录'],
    [21, '1001', '张三', '一组', '20260801', '15:45', '20260801', '18:45', 3, '删除', '', '', '', '', '', '', '', '重复填报'],
  ]));
  m.confirmBatch();
  captured = [];
  m.exportLocateIssues();
  const headers = captured[0].rows[0];
  assert.ok(headers.includes('系统序号（执行前）'), '编号列要标明是「执行前」的：' + JSON.stringify(headers));
  assert.ok(headers.includes('校对ID'), '校对单号要保留（班组/考勤员用它对账）');
  assert.ok(!headers.includes('系统序号'), '不能再叫光秃秃的「系统序号」（删完就不是那个意思了）');
});

// ==================== M5 出文件（2026-09-18） ====================
section('M5 出文件：时间补零 / 定额量必须是数字');

test('M5-6 导出 2007 的时间补零（7:00 → 07:00），带秒的截到分', () => {
  resetState();
  m.processGroupWorkbook({ fileName: 'g.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'],
    [1, '00151177', '童辉武', '一组', '2026-08-01', '7:00', '2026-08-01', '8:45', 1.75, 'x', '工作日', '核准'],
    [2, '00151177', '童辉武', '一组', '2026-08-02', '15:45:30', '2026-08-02', '18:45', 3, 'x', '工作日', '核准'],
  ] } });
  captured = [];
  m.exportSystemData('merged');
  const rows = captured[0].rows;
  assert.strictEqual(rows[1][6], '07:00', '开始时间要两位小时（真实上传文件 13265 行全两位）');
  assert.strictEqual(rows[1][7], '08:45', '结束时间要两位小时');
  assert.strictEqual(rows[2][6], '15:45', '带秒的截到分（M1 记过的升级路径）');
  assert.strictEqual(rows[2][7], '18:45');
});

test('M5-2 文本形式的数字时数，导出时归一成数字（真实文件那格是数字）', () => {
  resetState();
  m.processGroupWorkbook({ fileName: 'g.xlsx', sheetNames: ['一组'], sheets: { 一组: [
    ['序号', '工号', '姓名', '班组', '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数', '加班原因', '加班类别', '科负责人核准'],
    [1, '00151177', '童辉武', '一组', '2026-08-01', '15:45', '2026-08-01', '17:55', '2.5', 'x', '工作日', '核准'],
    [2, '00151177', '童辉武', '一组', '2026-08-02', '15:45', '2026-08-02', '18:45', 3, 'x', '工作日', '核准'],
  ] } });
  captured = [];
  m.exportSystemData('merged');
  const rows = captured[0].rows;
  assert.strictEqual(typeof rows[1][8], 'number', '文本数字要变成数字：' + JSON.stringify(rows[1][8]));
  assert.strictEqual(rows[1][8], 2.5);
  assert.strictEqual(typeof rows[2][8], 'number', '本来就是数字的照旧');
  assert.strictEqual(rows[2][8], 3);
});

// ==================== M6/M7/M8 评审跟进（2026-09-18） ====================
section('M6-1 投错区拒绝 · M7-1/M7-2 重置确认与完整重置 · M8-1 转义 · M8-2 文案');

test('M6-1 投错区会被拒绝（班组表投进异常表区不再静默"已匹配"）', () => {
  const groupLike = { fileName: 'g', sheetNames: ['一组'], sheets: { '一组': [GROUP_ROW_HEADERS,
    [1, '10010001', '张三', '一组', '2026-08-01', '15:45', '2026-08-01', '18:45', 3, 'x', '工作日', '核准']] } };
  const rectifyLike = rectifyParsed([[1, '10010001', '张三', '底盘一组', '20260801', '15:45', '20260801', '18:45', 1.83, '修改']]);
  assert.ok(m.tableProblem('dropZoneAbnormal', groupLike).includes('异常表'), '班组表投进异常表区要拒绝');
  assert.strictEqual(m.tableProblem('dropZoneAbnormal', rectifyLike), '', '整改表含异常表全部列，允许当异常表用');
  assert.strictEqual(m.tableProblem('dropZone', groupLike), '', '自家表放行');
  assert.ok(m.tableProblem('dropZone', rectifyLike).includes('班组填报表'), '整改表投进班组表区要拒绝');
  assert.strictEqual(m.tableProblem('dropZoneRectify', rectifyLike), '', '整改表放行');
});

test('M7-1「重新开始」会先确认：点取消不清空，确认后才清空', () => {
  resetState();
  m.processGroupWorkbook(groupWorkbookWithHours('2026-08-01', '15:45', '17:55', 2));
  const before = appState.mergedRecords.length;
  confirmAnswer = false;
  m.startNewRound();
  assert.strictEqual(appState.mergedRecords.length, before, '点取消不能清空');
  confirmAnswer = true;
  m.startNewRound();
  assert.strictEqual(appState.mergedRecords.length, 0, '确认后才清空');
});

test('M7-2 重置会把调班视图与解析中标志一起复位', () => {
  resetState();
  appState.outputShiftView = 'all';
  appState.isParsing = true;
  confirmAnswer = true;
  m.startNewRound();
  assert.strictEqual(appState.outputShiftView, 'current', '视图选择要回到默认');
  assert.strictEqual(appState.isParsing, false, '解析中标志要复位');
});

test('M8-1 单元格里的 < > & 会被转义（不再当 HTML 渲染）', () => {
  const html = m.renderTable([{ 工号: '1001', 姓名: '<b>粗体</b>', 加班原因: 'A&B<C>' }]);
  assert.ok(!html.includes('<b>粗体</b>'), '标签不能被当 HTML 渲染');
  assert.ok(html.includes('&lt;b&gt;粗体&lt;/b&gt;'), '要原样转义显示');
  assert.ok(html.includes('A&amp;B&lt;C&gt;'), '& 与 < 都要转义');
  assert.strictEqual(m.escapeHtml('a"b'), 'a&quot;b', '引号也要转义');
});

test('M8-2 界面上不再出现「组长」（业务口径是「班组考勤员」）', () => {
  resetState();
  appState.mergedRecords = makeMergedRecords();
  const pages = [m.renderImport(), m.renderRectify(), m.renderOutput()].join('\n');
  assert.ok(!pages.includes('组长'), '界面文案应统一为「班组考勤员」；仍出现的位置：' + (pages.match(/.{0,20}组长.{0,20}/) || [''])[0]);
});

// ==================== 汇总 ====================
console.log(`\n通过 ${passed} 项，失败 ${failures.length} 项`);
if (failures.length) {
  failures.forEach(f => console.log(`  - ${f.name}: ${f.err.message}`));
  process.exit(1);
}

