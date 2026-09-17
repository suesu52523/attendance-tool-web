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
  // 只模拟测试用到的那一种格式：'h:mm' 按真实 XLSX 的行为来（0.5 → "12:00"，小时不补零），
  // 其余格式返回空串。注意：如果这里永远返回空串，「时数列被当成钟点」的 bug 测不出来（假绿）。
  SSF: { format: (fmt, v) => (fmt === 'h:mm'
    ? `${Math.floor(v * 24)}:${String(Math.round(v * 1440) % 60).padStart(2, '0')}`
    : '') },
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
  formatCellValue,
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
    '处置方式', '修改后开始日期', '修改后开始时间', '修改后结束日期', '修改后结束时间', '修改后上报加班时数', '调班日期', '调班班次', '特殊情况说明'];
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

