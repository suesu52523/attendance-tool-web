// 月度数据全流程测试：七月 + 八月
// 用工具的 js/main.js 无头执行，模拟浏览器导入三张表 → 确认 → 导出
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const SRC_DIR = path.join(__dirname, '..');
const MAIN_JS = path.join(SRC_DIR, 'js', 'main.js');
const DATA_DIR = path.join(SRC_DIR, 'data');

// ---------- 加载 main.js（注入浏览器桩） ----------
const src = fs.readFileSync(MAIN_JS, 'utf8');

let captured = [];
const xlsxStub = {
  SSF: XLSX.SSF,
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
globalThis.window = { confirm: () => true };
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

// ---------- 用 real XLSX 解析 Excel → 工具所期望的格式 ----------
function parseExcelFile(filePath) {
  const workbook = XLSX.readFile(filePath);
  const result = {
    fileName: path.basename(filePath),
    sheetNames: workbook.SheetNames,
    sheets: {},
  };
  workbook.SheetNames.forEach(name => {
    const ws = workbook.Sheets[name];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    result.sheets[name] = json;
  });
  return result;
}

// ---------- 重置 ----------
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
  appState.groupDuplicates = [];
  captured = [];
}

// ---------- 统计辅助 ----------
function countBy(rows, key) {
  const map = {};
  rows.forEach(r => {
    const v = String(r[key] || '');
    map[v] = (map[v] || 0) + 1;
  });
  return map;
}

// ---------- 测试一个月的数据 ----------
function testMonth(monthName, dataDir) {
  const groupFile = path.join(dataDir, '班组填报表.xlsx');
  const abnormalFile = path.join(dataDir, '异常表.xlsx');
  const rectifyFile = path.join(dataDir, '整改表.xlsx');

  const result = {
    month: monthName,
    files: { group: groupFile, abnormal: abnormalFile, rectify: rectifyFile },
    steps: {},
    errors: [],
    warnings: [],
  };

  // 步骤1：导入与合并
  try {
    resetAll();
    const groupData = parseExcelFile(groupFile);
    m.processGroupWorkbook(groupData);

    const sheetCount = appState.groupSheets.length;
    const mergedCount = appState.mergedRecords.length;
    const failCount = appState.groupFailures.length;
    const dupCount = (appState.groupDuplicates || []).length;

    result.steps.import = {
      status: 'pass',
      sheetCount,
      mergedCount,
      failCount,
      dupCount,
      sheets: appState.groupSheets.map(s => ({ name: s.name, rowCount: s.rowCount, status: s.status, note: s.note })),
      failures: appState.groupFailures.slice(0, 20),
      duplicates: (appState.groupDuplicates || []).slice(0, 10),
    };

    if (failCount > 0) result.warnings.push(`导入校验失败 ${failCount} 条`);
    if (dupCount > 0) result.warnings.push(`重复填报 ${dupCount} 组`);
  } catch (e) {
    result.steps.import = { status: 'fail', error: e.message };
    result.errors.push(`导入失败: ${e.message}`);
    return result;
  }

  // 步骤2：异常处理
  try {
    const abnormalData = parseExcelFile(abnormalFile);
    m.processAbnormalWorkbook(abnormalData);

    const round = appState.rounds[0];
    const abnormalCount = appState.abnormalRecords.length;
    const matched = appState.abnormalRecords.filter(r => r['匹配状态'] === '已匹配').length;
    const unmatched = appState.abnormalRecords.filter(r => r['匹配状态'] === '未匹配').length;
    const multiHit = appState.abnormalRecords.filter(r => String(r['匹配状态']).includes('多条命中')).length;
    const failCount = round.abnormalFailures.length;
    const warnCount = round.abnormalWarnings.length;

    result.steps.abnormal = {
      status: failCount > 0 ? 'warning' : 'pass',
      abnormalCount,
      matched,
      unmatched,
      multiHit,
      failCount,
      warnCount,
      matchStats: countBy(appState.abnormalRecords, '匹配状态'),
      failures: round.abnormalFailures.slice(0, 20),
      warnings: round.abnormalWarnings.slice(0, 10),
    };

    if (failCount > 0) result.warnings.push(`异常匹配失败 ${failCount} 条`);
  } catch (e) {
    result.steps.abnormal = { status: 'fail', error: e.message };
    result.errors.push(`异常处理失败: ${e.message}`);
    return result;
  }

  // 步骤3：整改与批量操作
  try {
    const rectifyData = parseExcelFile(rectifyFile);
    m.processRectifyWorkbook(rectifyData);

    const ops = appState.rectifyOperations;
    const opStats = countBy(ops, '操作类型');

    // 确认执行
    m.confirmBatch();
    const round = appState.rounds[appState.currentRound];
    const confirmed = round.status === 'confirmed';
    const locateIssues = round.locateIssues || [];
    const finalCount = appState.mergedRecords.length;

    result.steps.rectify = {
      status: 'pass',
      opCount: ops.length,
      opStats,
      confirmed,
      locateIssueCount: locateIssues.length,
      locateIssueStats: countBy(locateIssues, '级别'),
      finalCount,
      locateIssues: locateIssues.slice(0, 20),
    };

    if (locateIssues.length > 0) result.warnings.push(`定位异常 ${locateIssues.length} 条`);
  } catch (e) {
    result.steps.rectify = { status: 'fail', error: e.message };
    result.errors.push(`整改失败: ${e.message}`);
    return result;
  }

  // 步骤4：导出
  try {
    captured = [];
    m.exportSystemData('current');
    m.exportShiftData('current');
    m.exportOperationLog();
    m.exportLocateIssues();
    const outputPage = m.renderOutput();

    result.steps.output = {
      status: 'pass',
      exportedSheets: captured.length,
      outputPageContains: outputPage.includes('轮处理完成') ? '完成标记' : '无完成标记',
      capturedSheets: captured.map(c => c.name),
    };
  } catch (e) {
    result.steps.output = { status: 'fail', error: e.message };
    result.errors.push(`导出失败: ${e.message}`);
  }

  return result;
}

// ---------- 执行测试 ----------
const testResults = {};

// 七月
console.log('\n========== 七月数据测试 ==========');
const julyDir = path.join(DATA_DIR, '七月模拟数据');
testResults.july = testMonth('七月', julyDir);

// 八月
console.log('\n========== 八月数据测试 ==========');
const augDir = path.join(DATA_DIR, '八月模拟数据');
testResults.august = testMonth('八月', augDir);

// 输出摘要
console.log('\n========== 测试摘要 ==========');
['july', 'august'].forEach(key => {
  const r = testResults[key];
  console.log(`\n${r.month}:`);
  console.log(`  错误: ${r.errors.length}`);
  console.log(`  警告: ${r.warnings.length}`);
  Object.keys(r.steps).forEach(step => {
    const s = r.steps[step];
    console.log(`  ${step}: ${s.status}${s.error ? ' (' + s.error + ')' : ''}`);
    if (s.mergedCount !== undefined) console.log(`    合并记录: ${s.mergedCount}`);
    if (s.abnormalCount !== undefined) console.log(`    异常记录: ${s.abnormalCount}`);
    if (s.opCount !== undefined) console.log(`    整改操作: ${s.opCount}`);
    if (s.finalCount !== undefined) console.log(`    最终记录: ${s.finalCount}`);
  });
});

// 写入 JSON 结果供报告生成使用
const outputPath = path.join(__dirname, '_monthly_test_results.json');
fs.writeFileSync(outputPath, JSON.stringify(testResults, null, 2), 'utf8');
console.log(`\n详细结果已写入: ${outputPath}`);
