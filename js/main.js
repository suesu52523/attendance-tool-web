// 总装科月度加班自动处理工具 - 前端交互
// 集成 SheetJS 实现本地 Excel 解析与导出

const steps = [
  { id: 'overview', label: '流程总览', subtitle: '查看本月加班数据处理全流程', icon: 'ph-squares-four' },
  { id: 'import', label: '导入与合并', subtitle: '读取班组填报表并生成合并大表', icon: 'ph-upload-simple' },
  { id: 'abnormal', label: '异常处理', subtitle: '导入校对异常表并生成整改表', icon: 'ph-warning-circle' },
  { id: 'rectify', label: '整改与批量操作', subtitle: '导入整改表并执行批量操作', icon: 'ph-arrows-left-right' },
  { id: 'output', label: '输出与审计', subtitle: '导出最终文件并查看操作记录', icon: 'ph-files' },
];

let currentStep = 0;

// 全局应用状态
const appState = {
  groupWorkbook: null,
  groupSheets: [],
  // 合并大表：随每轮批量操作逐步更新，始终保持最新状态
  mergedRecords: [],

  // 当前轮次的异常/整改数据（兼容单轮快速访问）
  abnormalWorkbook: null,
  abnormalRecords: [],
  rectifyWorkbook: null,
  rectifyOperations: [],

  // 步骤4「确认执行批量」后才会生成最终输出
  batchConfirmed: false,

  // 最终生成：点击「最终生成」后，异常处理/整改/输出才视为已完成
  finalGenerated: false,

  // 轮次管理：每月可能有 2-3 轮异常处理
  // 每轮独立保存异常、整改、调班、操作、失败记录及本轮处理后的快照
  rounds: [],
  currentRound: 0, // 当前正在处理的轮次索引

  // 失败数据记录（校验失败、匹配失败）
  groupFailures: [],
  abnormalFailures: [],

  // 重复填报检测（同工号 + 同一天 + 同一开始时间出现多次）
  groupDuplicates: [],

  fileName: '',

  // 输出页「调班记录」查看模式：'current' 只看本轮 / 'all' 累计全部
  outputShiftView: 'current',

  isParsing: false,
};

// 获取当前轮次对象；不存在时自动创建
function getCurrentRound() {
  if (!appState.rounds[appState.currentRound]) {
    appState.rounds[appState.currentRound] = createRound(appState.currentRound + 1);
  }
  return appState.rounds[appState.currentRound];
}

function createRound(roundNo) {
  return {
    roundNo,
    abnormalWorkbook: null,
    abnormalRecords: [],
    rectifyWorkbook: null,
    rectifyOperations: [],
    shiftRecords: [],
    systemRecords: [],
    abnormalFailures: [], // 本轮异常匹配失败记录
    abnormalWarnings: [], // 本轮异常匹配提醒（多条命中 / 姓名不一致）
    locateIssues: [], // 本轮批量操作定位异常清单（未定位 / 多条命中）
    status: 'pending', // pending -> processing -> confirmed
    confirmedAt: null,
  };
}

// 获取所有已完成轮次
function getConfirmedRounds() {
  return appState.rounds.filter(r => r.status === 'confirmed');
}

// 累计全部轮次的调班数据（用于最终输出）
function getAllShiftRecords() {
  return appState.rounds.flatMap(r => r.shiftRecords || []);
}

// 累计全部轮次的操作记录（用于最终输出）
function getAllOperations() {
  return appState.rounds.flatMap(r => r.rectifyOperations || []);
}

// 根据实际导入/处理数据计算工作状态与步骤完成状态
// 主菜单图标应基于 workflow status，不随当前查看步骤回退
function getWorkflowStatus() {
  const round = getCurrentRound();
  const hasGroup = !!appState.groupWorkbook && appState.mergedRecords.length > 0;
  const hasAbnormal = appState.rounds.some(r => r.abnormalRecords && r.abnormalRecords.length > 0)
    || (round.abnormalRecords && round.abnormalRecords.length > 0);
  const hasRectify = appState.rounds.some(r => r.rectifyOperations && r.rectifyOperations.length > 0)
    || (round.rectifyOperations && round.rectifyOperations.length > 0);
  const hasConfirmed = appState.rounds.some(r => r.status === 'confirmed');
  const finalGenerated = appState.finalGenerated;

  let workingStatus = '待导入班组表';
  let workingStep = 0;
  let continueStep = 1;

  if (!hasGroup) {
    workingStatus = '待导入班组表';
    workingStep = 0;
    continueStep = 1;
  } else if (!hasAbnormal) {
    workingStatus = '导入与合并';
    workingStep = 1;
    continueStep = 2;
  } else if (!hasRectify) {
    workingStatus = `异常处理 第${round.roundNo}轮`;
    workingStep = 2;
    continueStep = 3;
  } else if (!hasConfirmed && round.status !== 'confirmed') {
    workingStatus = `整改与批量操作 第${round.roundNo}轮`;
    workingStep = 3;
    continueStep = 4;
  } else {
    workingStatus = `输出与审计 第${round.roundNo}轮`;
    workingStep = 4;
    continueStep = 4;
  }

  return {
    hasGroup,
    hasAbnormal,
    hasRectify,
    hasConfirmed,
    finalGenerated,
    workingStatus,
    workingStep,
    continueStep,
    stepCompleted: {
      overview: true,
      import: hasGroup,
      // 异常处理 / 整改与批量操作 / 输出 只有在点击「最终生成」后才显示已完成
      abnormal: finalGenerated,
      rectify: finalGenerated,
      output: finalGenerated,
    },
    stepSubtitle: {
      overview: '流程指引',
      import: hasGroup ? '已完成' : '待处理',
      abnormal: finalGenerated ? '已完成' : (hasAbnormal ? `第${round.roundNo}轮` : '待处理'),
      rectify: finalGenerated ? '已完成' : (hasRectify ? `第${round.roundNo}轮` : '待处理'),
      output: finalGenerated ? '已完成' : (hasConfirmed ? `第${round.roundNo}轮` : '待处理'),
    },
  };
}

// ==================== 流程约束：只有无异常时才允许继续 ====================

// 步骤1校验：班组填报表导入后是否有校验失败或重复填报
function getImportBlockers() {
  const blockers = [];
  if (!appState.groupWorkbook) {
    blockers.push('尚未导入班组填报表');
    return blockers;
  }
  const failures = appState.groupFailures || [];
  if (failures.length) {
    blockers.push(`有 ${failures.length} 条校验失败记录（工号/姓名/日期/时间为空或格式错误），请修正后重新导入`);
  }
  const duplicates = appState.groupDuplicates || [];
  if (duplicates.length) {
    blockers.push(`有 ${duplicates.length} 组重复填报（同一工号同一天同一开始时间出现多次），请清理后重新导入`);
  }
  return blockers;
}

// 步骤2校验：异常表导入后是否有匹配失败或需人工核对的异常
function getAbnormalBlockers() {
  const blockers = [];
  const round = getCurrentRound();
  if (!round.abnormalRecords.length) {
    blockers.push('尚未导入异常表');
    return blockers;
  }
  const failures = round.abnormalFailures || appState.abnormalFailures || [];
  if (failures.length) {
    blockers.push(`有 ${failures.length} 条匹配失败记录（无法在合并大表中定位），请核对工号/日期/时间后修正异常表并重新导入`);
  }
  const warnings = round.abnormalWarnings || [];
  if (warnings.length) {
    blockers.push(`有 ${warnings.length} 条需人工核对（多条命中或姓名不一致），请在异常表补填「开始时间」或清理合并大表后重新导入`);
  }
  // 检查异常表内部数据是否有结束时间早于开始时间等数据问题
  const dateIssues = (round.abnormalRecords || []).filter(r => {
    const st = String(r['开始时间'] || '').trim();
    const et = String(r['结束时间'] || '').trim();
    if (!st || !et) return false;
    const stParts = parseTimeParts(st);
    const etParts = parseTimeParts(et);
    if (!stParts || !etParts) return false;
    return stParts.h > etParts.h || (stParts.h === etParts.h && stParts.m > etParts.m);
  });
  if (dateIssues.length) {
    blockers.push(`有 ${dateIssues.length} 条记录的结束时间早于开始时间，请修正后重新导入`);
  }
  return blockers;
}

// 步骤3校验：整改表导入后是否有未填写处置方式或定位异常
function getRectifyBlockers() {
  const blockers = [];
  const round = getCurrentRound();
  if (!round.rectifyOperations.length) {
    blockers.push('尚未导入整改表');
    return blockers;
  }
  const ops = round.rectifyOperations || appState.rectifyOperations || [];
  const unfilled = ops.filter(o => o['操作类型'] === '未填写').length;
  if (unfilled > 0) {
    blockers.push(`有 ${unfilled} 条记录未填写处置方式，请在整改表中补充填写后重新导入`);
  }
  // 检查修改类操作的修改后数据是否有结束时间早于开始时间
  const modifyIssues = ops.filter(op => {
    if (op['操作类型'] !== '修改') return false;
    const st = String(op['修改后开始时间'] || '').trim();
    const et = String(op['修改后结束时间'] || '').trim();
    if (!st || !et) return false;
    const stParts = parseTimeParts(st);
    const etParts = parseTimeParts(et);
    if (!stParts || !etParts) return false;
    return stParts.h > etParts.h || (stParts.h === etParts.h && stParts.m > etParts.m);
  });
  if (modifyIssues.length) {
    blockers.push(`有 ${modifyIssues.length} 条修改操作的结束时间早于开始时间，请修正后重新导入`);
  }
  return blockers;
}

// 综合校验：判断从当前步骤是否可以继续到下一步
function canProceedToStep(targetStep) {
  // 允许回退到之前的步骤
  if (targetStep <= currentStep) return { ok: true, blockers: [] };

  const blockers = [];
  // 从当前步骤到目标步骤之间的每一步都必须通过校验
  for (let s = currentStep; s < targetStep; s++) {
    if (s === 1 && getImportBlockers().length) {
      blockers.push(...getImportBlockers());
    }
    if (s === 2 && getAbnormalBlockers().length) {
      blockers.push(...getAbnormalBlockers());
    }
    if (s === 3 && getRectifyBlockers().length) {
      blockers.push(...getRectifyBlockers());
    }
  }
  return { ok: blockers.length === 0, blockers };
}
const GROUP_HEADERS = [
  '序号', '工号', '姓名', '班组',
  '加班开始日期', '加班开始时间', '加班结束日期', '加班结束时间', '加班时数',
  '加班原因', '加班类别', '科负责人核准'
];

const ABNORMAL_HEADERS = [
  'ID', '工号', '姓名', '科室',
  'T-1日*系统排班', 'T0日*系统排班', 'T+1日*系统排班',
  '开始加班打卡时间', '结束加班打卡时间',
  '开始日期', '结束日期', '开始时间', '结束时间', '转换加班区间',
  '上报加班时数', '实际加班时数(未减吃饭时间)', '差异',
  '提醒信息', '就餐信息'
];

const RECTIFY_ACTION_HEADERS = [
  '处置方式', '修改后开始日期', '修改后开始时间', '修改后结束日期', '修改后结束时间',
  '修改后上报加班时数', '调班日期', '调班班次', '异常说明（必填）'
];

const SYSTEM_OUTPUT_HEADERS = [
  '中文名称', '工号', '姓名', '开始日期', '结束日期',
  '类型', '开始时间', '结束时间', '定额量', '加班报酬类型', '加班原因'
];

// 加班时数上限（业务口径：2026-09-17 由考勤业务方定为 48 小时）
// 用途：合表导入与整改执行时拦下明显异常的时数（如 49 小时这类跨天算错 / 多打一位的值）
// 为什么是 48：单条加班最长按“连续两个整天”估；再高基本是填错，宁可退回班组也不往上传
// 要调这个数字：只改这一处，改完重跑 `node test/main.test.js` 与 `node test/scenarios.test.js`
// 0 与负数现已按「不合理」拦下；若业务上确需「0 小时占位」，只改 overtimeHoursProblem 里的 n <= 0
const MAX_OVERTIME_HOURS = 48;

// 加班时数校验（合表导入与整改执行共用）：不合理返回原因文本，合理返回空串
// 允许数字，也允许「文本形式的数字」（Excel 里被存成文本的 "2.5"）；"半小时" 这类文字一律退回
function overtimeHoursProblem(hours) {
  if (hours === '' || hours === undefined || hours === null) return '';
  const n = typeof hours === 'number' ? hours : Number(String(hours).trim());
  if (!Number.isFinite(n)) return `加班时数「${hours}」不是数字，请填小时数（如 2 或 2.5）`;
  if (n <= 0) return `加班时数 ${n} 小时不合理（必须大于 0）`;
  if (n > MAX_OVERTIME_HOURS) return `加班时数 ${n} 小时超过 ${MAX_OVERTIME_HOURS} 小时上限`;
  return '';
}

// 导出用的定额量：必须是数字（真实上传文件 13265 行该格全是数字）
// 文本形式的数字（"2.5"）归一成数字；非数字原样保留（这类值 M3 导入校验已拦过）
function toHourNumber(hours) {
  if (hours === '' || hours === undefined || hours === null) return 0;
  if (typeof hours === 'number') return hours;
  const n = Number(String(hours).trim());
  return Number.isFinite(n) ? n : hours;
}

const SHIFT_MAIN_HEADERS = ['中文名称', '工号', '姓名', '开始日期', '结束日期', '日工作计划'];
const SHIFT_SHEET2_HEADERS = ['中文名称', '工号', '姓名', '开始日期', '结束日期', '日工作计划', '出勤项目分类', '备注'];

const SHIFT_TIME_MAP = {
  'SF04 双班早班': '7:00-15:45',
  'SF17 固定班': '8:45-17:30',
  'SF05 双班中班': '15:45-00:20',
  'SF10 二线中班1545': '15:45-00:30',
  'SF11 二线中班1645': '16:45-1:30',
  'SF12 二线中班1755': '17:55-2:40',
  'SF13 二线中班1845': '18:45-3:45',
  'OFF 休息': '-',
  'NS 未排班': '-',
};

// 示例数据
const demoSheets = [
  { name: '底盘一组', rows: 12, status: 'ok' },
  { name: '底盘二组', rows: 10, status: 'ok' },
  { name: '前悬一组', rows: 15, status: 'ok' },
  { name: '前悬二组', rows: 11, status: 'warning' },
  { name: '车门一组', rows: 13, status: 'ok' },
];

const demoRecords = [
  { 系统序号: 1, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 加班开始日期: '2026-08-01', 加班开始时间: '15:45', 加班结束日期: '2026-08-01', 加班结束时间: '17:35', 加班时数: 1.83, 加班原因: '产能爬坡', 加班类别: '工作日', 科负责人核准: '核准' },
  { 系统序号: 2, 工号: '10010002', 姓名: '李四', 班组: '底盘一组', 加班开始日期: '2026-08-02', 加班开始时间: '07:00', 加班结束日期: '2026-08-02', 加班结束时间: '15:45', 加班时数: 8.75, 加班原因: '设备检修', 加班类别: '休息日', 科负责人核准: '核准' },
  { 系统序号: 3, 工号: '10010003', 姓名: '王五', 班组: '前悬一组', 加班开始日期: '2026-08-02', 加班开始时间: '15:45', 加班结束日期: '2026-08-02', 加班结束时间: '18:45', 加班时数: 3, 加班原因: '产能爬坡', 加班类别: '工作日', 科负责人核准: '核准' },
];

const demoAbnormal = [
  { ID: 45, 工号: '10010001', 姓名: '张三', 科室: '底盘一组', 'T-1日*系统排班': '双班中班 2026-07-31 15:45:00~2026-08-01 00:20:00', 'T0日*系统排班': '双班早班 2026-08-01 07:00:00~2026-08-01 15:45:00', 'T+1日*系统排班': '固定班 2026-08-02 08:45:00~2026-08-02 17:30:00', 开始加班打卡时间: '', 结束加班打卡时间: '2026-08-01 16:56:29', 开始日期: '20260801', 结束日期: '20260801', 开始时间: '15:45', 结束时间: '18:45', '转换加班区间': '15:45-18:45', 上报加班时数: 3, '实际加班时数(未减吃饭时间)': 1, 差异: 2, 提醒信息: '【加班结束卡】请确认加班结束时间', 就餐信息: '无', 处置状态: '待处理' },
];

const demoOperations = [
  { 系统序号: 45, 工号: '10010001', 姓名: '张三', 班组: '底盘一组', 操作类型: '修改', 操作详情: '修改后：20260801 18:00-21:00，3h' },
  { 系统序号: 46, 工号: '10010002', 姓名: '李四', 班组: '前悬一组', 操作类型: '删除', 操作详情: '班组考勤员确认重复填报，执行删除' },
  { 系统序号: 47, 工号: '10010003', 姓名: '王五', 班组: '车门一组', 操作类型: '调班', 操作详情: '调班处理：导出至调班模板' },
  { 系统序号: 48, 工号: '10010004', 姓名: '赵六', 班组: '电装一组', 操作类型: '特殊情况', 操作详情: '已口头报备，不做处理' },
];

function init() {
  renderNav();
  renderContent();
  updateStats();
}

function renderNav() {
  const nav = document.getElementById('stepNav');
  const wf = getWorkflowStatus();
  const stepStatusKeys = ['overview', 'import', 'abnormal', 'rectify', 'output'];

  nav.innerHTML = steps.map((step, index) => {
    const isActive = index === currentStep;
    const completed = wf.stepCompleted[stepStatusKeys[index]];
    const stepNo = index + 1;

    // 图标状态：已完成显示绿色对勾；当前步骤显示原图标蓝色背景；未开始显示原图标灰色背景
    let iconHtml;
    if (completed && !isActive) {
      iconHtml = `<div class="w-8 h-8 rounded-xl flex items-center justify-center text-lg bg-apple-green/10 text-apple-green"><i class="ph ph-check-circle"></i></div>`;
    } else if (isActive) {
      iconHtml = `<div class="w-8 h-8 rounded-xl flex items-center justify-center text-lg bg-apple-blue text-white"><i class="ph ${step.icon}"></i></div>`;
    } else {
      iconHtml = `<div class="w-8 h-8 rounded-xl flex items-center justify-center text-lg bg-apple-gray text-apple-muted"><i class="ph ${step.icon}"></i></div>`;
    }

    return `
      <button
        onclick="goToStep(${index})"
        class="step-btn w-full flex items-center gap-3 px-4 py-3.5 rounded-2xl text-left transition-all duration-200 ${
          isActive
            ? 'bg-apple-blue/10 text-apple-blue active'
            : 'text-apple-muted hover:bg-apple-gray hover:text-apple-text'
        }"
      >
        ${iconHtml}
      <div class="flex-1 min-w-0">
        <div class="text-sm font-semibold leading-tight">${step.label}</div>
        <div class="text-xs mt-0.5 truncate ${isActive ? 'text-apple-blue/70' : 'text-apple-muted'}">
          ${wf.stepSubtitle[stepStatusKeys[index]]}
        </div>
      </div>
      </button>
    `;
  }).join('');
}

function goToStep(index) {
  // 流程约束：前进到后续步骤时，检查中间步骤是否全部通过校验
  if (index > currentStep) {
    const { ok, blockers } = canProceedToStep(index);
    if (!ok) {
      showToast(`无法继续，请先解决以下问题：\n${blockers.join('\n')}`, 'error');
      return;
    }
  }
  currentStep = index;
  renderNav();
  renderContent();
}

function renderContent() {
  const step = steps[currentStep];
  document.getElementById('pageTitle').textContent = step.label;
  document.getElementById('pageSubtitle').textContent = step.subtitle;

  const contentArea = document.getElementById('contentArea');
  contentArea.innerHTML = `<div class="step-content max-w-7xl mx-auto">${getStepHtml(step.id)}</div>`;

  // 每次内容渲染时同步更新左侧菜单状态与会话统计
  renderNav();
  updateStats();
  bindDropZones();
  bindHiddenFileInput();
}

function getStepHtml(id) {
  switch (id) {
    case 'overview': return renderOverview();
    case 'import': return renderImport();
    case 'abnormal': return renderAbnormal();
    case 'rectify': return renderRectify();
    case 'output': return renderOutput();
    default: return '';
  }
}

// ==================== 通用表格渲染 ====================

// M8-1：把数据拼进 HTML 前一律转义（表格 / 文件名 / 清单说明等）。
// 真数据里暂时没有 < > &（扫过 17 个输入文件 0 格），但值是人工填的，不能靠运气
function escapeHtml(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function renderTable(rows, options = {}) {
  const { maxRows = 20, emptyText = '暂无数据' } = options;
  if (!rows || rows.length === 0) {
    return `<div class="flex flex-col items-center justify-center py-16 text-apple-muted">
      <i class="ph ph-table text-4xl mb-3 opacity-30"></i>
      <span class="text-sm">${emptyText}</span>
    </div>`;
  }

  const headers = Object.keys(rows[0]);
  const displayRows = rows.slice(0, maxRows);

  return `
    <div class="overflow-x-auto">
      <table class="data-table bg-white min-w-full">
        <thead class="bg-apple-gray/50">
          <tr>
            ${headers.map(h => `<th>${h}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
          ${displayRows.map((row, idx) => `
            <tr>
              ${headers.map(h => {
                const val = row[h];
                if (h === '状态' || h === '异常状态' || h === '处置状态') {
                  return `<td>${renderStatusBadge(val)}</td>`;
                }
                if (h === '操作类型') {
                  return `<td>${renderOpBadge(val)}</td>`;
                }
                return `<td class="${h === '系统序号' || h === '工号' || h === 'ID' ? 'font-medium' : ''}">${val === undefined || val === null ? '-' : escapeHtml(val)}</td>`;
              }).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
    ${rows.length > maxRows ? `<div class="mt-3 text-xs text-apple-muted text-center">共 ${rows.length} 条，显示前 ${maxRows} 条</div>` : ''}
  `;
}

function renderStatusBadge(status) {
  const s = String(status || '');
  if (s.includes('正常') || s === 'ok') return '<span class="badge badge-success">正常</span>';
  if (s.includes('异常') || s.includes('待处理') || s === 'warning') return '<span class="badge badge-warning">待处理</span>';
  if (s.includes('删除')) return '<span class="badge badge-danger">删除</span>';
  if (s.includes('修改')) return '<span class="badge badge-info">修改</span>';
  if (s.includes('调班')) return '<span class="badge badge-warning">调班</span>';
  if (s.includes('特殊情况')) return '<span class="badge badge-muted">特殊情况</span>';
  return `<span class="badge badge-muted">${escapeHtml(s)}</span>`;
}

function renderOpBadge(type) {
  const cls = getOpBadgeClass(type);
  return `<span class="badge ${cls}">${escapeHtml(type)}</span>`;
}

// 未导入文件时页面展示的是内置示例数据，明确标注避免误读为已导入的真实数据
function renderDemoBadge(isDemo) {
  return isDemo
    ? '<span class="badge badge-warning">示例数据 · 未导入文件</span>'
    : '';
}

// ==================== Excel 解析与工具函数 ====================

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const result = {
          fileName: file.name,
          sheetNames: workbook.SheetNames,
          sheets: {},
        };
        workbook.SheetNames.forEach(name => {
          const worksheet = workbook.Sheets[name];
          const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          result.sheets[name] = json;
        });
        resolve(result);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function sheetToObjects(rows) {
  if (!rows || rows.length < 2) return [];
  const headers = rows[0].map(h => String(h || '').trim());
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i] === undefined || row[i] === null ? '' : row[i];
    });
    return obj;
  });
}

function normalizeTime(str) {
  if (str === undefined || str === null) return '';
  return String(str).replace(/[\uFF1A]/g, ':').replace(/[\uFF0D]/g, '-').trim();
}

// 将时间统一规范为 HH:MM，避免 Excel h:mm（如 8:15）与文本 08:15 不匹配
function padTime(str) {
  const t = normalizeTime(str);
  if (!t) return '';
  const m = t.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!m) return t;
  return `${String(parseInt(m[1], 10)).padStart(2, '0')}:${m[2]}`;
}

function normalizeDate(str) {
  if (str === undefined || str === null) return '';
  let s = String(str).trim();
  if (!s) return '';
  // 处理 Excel 日期对象（如 Sat Aug 01 2026 ...）
  if (s.includes('GMT') || s.includes('UTC')) {
    const d = new Date(s);
    if (!isNaN(d)) {
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    }
  }
  // 中文日期写法统一成 2026-08-10：2026年8月10日 / 2026 年 8 月 10 号 / 全角数字
  // 这里只管“把写法统一”，日期到底存不存在由 parseDateParts 的日历校验把关（2 月 30 日不转）
  const cn = s.normalize('NFKC').match(/^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]$/);
  if (cn) {
    const y = +cn[1], mo = +cn[2], d = +cn[3];
    const rt = new Date(y, mo - 1, d);
    if (rt.getFullYear() === y && rt.getMonth() === mo - 1 && rt.getDate() === d) {
      return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }
  }
  return s;
}

function toYYYYMMDD(dateVal) {
  if (!dateVal && dateVal !== 0) return '';
  const s = String(dateVal).trim();
  if (/^\d{8}$/.test(s)) return s;
  // 2026-08-01 / 2026/08/01（必须整串匹配，避免 "2026/8/10x" 被截成合法日期）
  const m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) {
    return `${m[1]}${String(m[2]).padStart(2, '0')}${String(m[3]).padStart(2, '0')}`;
  }
  // 尝试 Date 对象
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}${mo}${day}`;
  }
  return s;
}

function parseTimeParts(timeStr) {
  const t = normalizeTime(timeStr);
  // 必须整串就是一个时刻：拦下 "15:00~19:00"（一格写两个时间）、"19:00（次日）" 这类
  // （以前没锁头尾，"15:00~19:00" 会被当成 15:00 用，脏值一路进大表并原样导出）
  // 口径跟 padTime 一致：h:mm，可带秒
  // ponytail: 带秒的值会原样进导出；若校对系统不认 h:mm:ss，再在导出前截断到分
  const m = t.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (!m) return null;
  const h = parseInt(m[1], 10);
  const mi = parseInt(m[2], 10);
  // 时/分超出范围视为非法时间（拦截 25:70、24:00 这类脏数据）
  if (h > 23 || mi > 59) return null;
  return { h, m: mi };
}

function computeHours(startDate, startTime, endDate, endTime) {
  const sd = normalizeDate(startDate);
  const st = normalizeTime(startTime);
  const ed = normalizeDate(endDate);
  const et = normalizeTime(endTime);
  if (!sd || !st || !ed || !et) return '';

  const sdParts = parseDateParts(sd);
  const stParts = parseTimeParts(st);
  const edParts = parseDateParts(ed);
  const etParts = parseTimeParts(et);

  if (!sdParts || !stParts || !edParts || !etParts) return '';

  const start = new Date(sdParts.y, sdParts.m - 1, sdParts.d, stParts.h, stParts.m);
  const end = new Date(edParts.y, edParts.m - 1, edParts.d, etParts.h, etParts.m);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return '';

  const diff = (end - start) / 3600000;
  // 结束早于开始**不再自动补一天**：业务口径是「填错了」，由调用方拦下退回人工核对
  // （合表校验 processGroupWorkbook / 整改执行 applyBatchOperations）
  // ponytail: 这里用本地时间做差；中国无夏令时，夏令时地区会差 1 小时。升级路径 = 换时区感知的日期库
  return parseFloat(diff.toFixed(2));
}

function parseDateParts(dateStr) {
  const s = String(dateStr).trim();
  let parts = null;
  if (/^\d{8}$/.test(s)) {
    parts = { y: parseInt(s.slice(0, 4), 10), m: parseInt(s.slice(4, 6), 10), d: parseInt(s.slice(6, 8), 10) };
  } else {
    const m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (m) {
      parts = { y: parseInt(m[1], 10), m: parseInt(m[2], 10), d: parseInt(m[3], 10) };
    }
  }
  if (parts) {
    // 格式对不代表日期存在：用 Date 往返比对拦下 2 月 30 日 / 13 月 / 4 月 31 日
    // （不拦的话 new Date(2026,1,30) 会锚静挪到 2026-03-02，进 2007 表就成了 20260230）
    // ponytail: 往返校验靠本地时间；中国无夏令时，夏令时地区会差 1 小时。升级路径 = Temporal.PlainDate
    const rt = new Date(parts.y, parts.m - 1, parts.d);
    const exists = rt.getFullYear() === parts.y && rt.getMonth() === parts.m - 1 && rt.getDate() === parts.d;
    return exists ? parts : null;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return { y: d.getFullYear(), m: d.getMonth() + 1, d: d.getDate() };
  }
  return null;
}

function inferShiftCode(startTime, endTime, schedule) {
  // 简单推断：按实际加班起止时间匹配常见班次
  const st = normalizeTime(startTime);
  const et = normalizeTime(endTime);
  if (schedule && String(schedule).includes('固定班')) return 'SF17 固定班';
  if (st === '07:00' || st === '7:00') return 'SF04 双班早班';
  if (st === '15:45') return 'SF05 双班中班';
  if (st === '16:45') return 'SF11 二线中班1645';
  if (st === '17:55') return 'SF12 二线中班1755';
  if (st === '18:45') return 'SF13 二线中班1845';
  if (st === '20:00' || et === '04:45') return 'SF29 一部改造期中班';
  return 'OFF 休息';
}

// Excel 日期/时间序列号格式化为可读字符串
function excelDateToString(serial) {
  if (typeof serial !== 'number' || isNaN(serial)) return '';
  // Excel 日期序列号大致范围：1 = 1900-01-01，45400 ≈ 2024-04
  if (serial < 30000 || serial > 60000) return '';
  try {
    return XLSX.SSF.format('yyyy-mm-dd', serial);
  } catch (e) {
    return '';
  }
}

function excelTimeToString(serial) {
  if (typeof serial !== 'number' || isNaN(serial)) return '';
  if (serial < 0 || serial >= 1) return '';
  try {
    return XLSX.SSF.format('h:mm', serial);
  } catch (e) {
    return '';
  }
}

// 根据表头判断并转换单元格值（日期/时间序列号）
function formatCellValue(value, header) {
  const h = String(header || '');
  if (h.includes('日期') && typeof value === 'number') {
    const formatted = excelDateToString(value);
    if (formatted) return formatted;
  }
  // 「时数」列存的是小时数（0.5 = 半小时），不是钟点 —— 不能按 h:mm 格式化。
  // 典型陷阱：「实际加班时数(未减吃饭时间)」，列名里的“时间”来自“未减吃饭时间”。
  // ponytail: 此处靠列名约定判断；一旦出现名字里不含「时数」的小时列就会退化。
  //            升级路径 = 读单元格自带的数字格式（SheetJS cellNF + SSF.is_date），不再猜列名。
  if (h.includes('时间') && !h.includes('时数') && typeof value === 'number') {
    const formatted = excelTimeToString(value);
    if (formatted) return formatted;
  }
  // 日期列里的中文写法也统一（三张导入表都过这个函数，转一次，下游的校验、匹配、导出就全对齐了）
  // 只对含“年月日/号”的文本下手，其他写法原样不动（不扩大改动面）
  if (h.includes('日期') && typeof value === 'string' && /[年月日号]/.test(value)) {
    const normalized = normalizeDate(value);
    if (normalized !== value) return normalized;
  }
  return value;
}

function computeHoursRaw(startDate, startTime, endDate, endTime) {
  const sdNum = typeof startDate === 'number' && startDate > 30000 && startDate < 60000 ? startDate : null;
  const edNum = typeof endDate === 'number' && endDate > 30000 && endDate < 60000 ? endDate : null;
  const stNum = typeof startTime === 'number' && startTime >= 0 && startTime < 1 ? startTime : null;
  const etNum = typeof endTime === 'number' && endTime >= 0 && endTime < 1 ? endTime : null;

  if (sdNum !== null && stNum !== null && edNum !== null && etNum !== null) {
    const diff = ((edNum + etNum) - (sdNum + stNum)) * 24;
    return parseFloat(diff.toFixed(2));
  }

  return computeHours(
    normalizeDate(startDate),
    normalizeTime(startTime),
    normalizeDate(endDate),
    normalizeTime(endTime)
  );
}

// 业务键定位：工号 + 开始日期（+开始时间）
// 返回全部候选记录，不做"取第一条"的隐式截断，由调用方决定如何处理多条命中
// 注意：不依赖校对系统 ID（ID 与合并大表序号无关），避免改错/删错行
function locateMergedRecords(empNo, startDate, startTime, records) {
  const pool = records || appState.mergedRecords || [];
  const no = String(empNo || '').trim();
  if (!no) return [];

  let candidates = pool.filter(r => String(r['工号']).trim() === no);
  if (!candidates.length) return [];

  const date = toYYYYMMDD(startDate);
  if (date) {
    const byDate = candidates.filter(r => toYYYYMMDD(r['加班开始日期']) === date);
    if (!byDate.length) return [];
    candidates = byDate;
  }

  // 提供了开始时间就必须匹配上（按 时:分 比较，兼容 15:45 / 15:45:00 / 8:15）
  const time = padTime(startTime);
  if (time) {
    return candidates.filter(r => !!padTime(r['加班开始时间']) && padTime(r['加班开始时间']) === time);
  }

  // 未提供开始时间时保留同日全部候选，由调用方按"多条命中"提示人工确认
  return candidates;
}

// 为一条整改操作定位合并大表中的目标记录（定位异常清单里，“校对ID”是校对系统的单号，“系统序号”是大表行号，两者不是一回事）
// 只用业务键（工号 + 原开始日期 + 原开始时间）：定位不到就返回未定位，由调用方写进《定位异常清单》让人核对
// 为什么不回退到「序号」：大表的系统序号每次导入都从 1 重发、删除后还会重排
//   （见 processGroupWorkbook 的 systemNo、applyBatchOperations 删除分支的重排序号）
//   —— 它只是当次会话的排号，不是身份，旧文件里的号码会指到别人身上（历史上“删错人”就是这么来的）
// ponytail: 若将来确实需要回填旧清单，用稳定的业务键（工号 + 日期）而不是序号
// 返回 { target, method, hits, candidates }，target 为 null 表示未定位到
function resolveOperationTarget(op) {
  const byKey = locateMergedRecords(op['工号'], op['原开始日期'], op['原开始时间']);
  if (byKey.length) return { target: byKey[0], method: '业务键', hits: byKey.length, candidates: byKey };

  return { target: null, method: '未定位', hits: 0, candidates: [] };
}

// 候选记录摘要：序号(班组) 前 5 条，超过 5 条补"等 N 条"，便于一眼看出重复数据来自哪些班组
function describeCandidates(candidates) {
  const list = candidates.slice(0, 5)
    .map(r => `${r['系统序号']}${r['班组'] ? `(${r['班组']})` : ''}`)
    .join('、');
  return candidates.length > 5 ? `${list} 等 ${candidates.length} 条` : list;
}

// 多条命中时的辅助提示：若候选里恰好有一条的班组与异常表/整改表填写的科室(班组)一致，直接指出来
function buildDeptHint(candidates, dept) {
  const d = String(dept || '').trim();
  if (!d) return '';
  const same = candidates.filter(r => String(r['班组'] || '').trim() === d);
  return same.length === 1
    ? `；其中序号 ${same[0]['系统序号']}（${d}）与表里填写的科室/班组一致，建议优先核对这一条`
    : '';
}

function buildLocateIssue(op, resolved, level, message) {
  const candidates = (resolved && resolved.candidates) || [];
  return {
    级别: level,
    // 标明是「执行前」的编号：删除会把全表序号重排成 1..N，删完之后这个号可能指向别人
    '系统序号（执行前）': (resolved && resolved.target) ? resolved.target['系统序号'] : '',
    校对ID: op['校对ID'] || '',
    工号: op['工号'],
    姓名: op['姓名'],
    班组: op['班组'],
    操作类型: op['操作类型'],
    原开始日期: op['原开始日期'] || '',
    原开始时间: op['原开始时间'] || '',
    命中数: candidates.length,
    候选序号: describeCandidates(candidates),
    说明: message,
  };
}

// 日期换算成「天序号」，用于算两条记录相差几天（null = 日期不可用）
function dayNumber(dateStr) {
  const p = parseDateParts(normalizeDate(dateStr));
  return p ? Date.UTC(p.y, p.m - 1, p.d) / 86400000 : null;
}

// 匹配失败时的核对线索：同一天大表里有什么 / 最近一次是什么
// 为什么需要：校对系统记的是「班次起点」、班组表记的是「加班起点」，两边时刻不同就会匹配失败；
//   把大表里当天的记录（含工时是否一致）直接写进清单，人一眼判断「不用动」还是「要改」
function buildMatchHint(empNo, startDate, reportedHours) {
  const mine = appState.mergedRecords.filter(r => String(r['工号']).trim() === empNo);
  if (!mine.length) return '该工号在大表里没有任何记录（可能本月没有上报）';
  const span = r => `${r['加班开始时间']}-${r['加班结束时间']}（${r['加班时数']}h）`;
  const key = toYYYYMMDD(startDate);
  const sameDay = key ? mine.filter(r => toYYYYMMDD(r['加班开始日期']) === key) : [];
  if (sameDay.length) {
    const hasHours = reportedHours !== '' && reportedHours !== undefined && reportedHours !== null;
    const hit = hasHours && sameDay.some(r => Number(r['加班时数']) === Number(reportedHours));
    const tail = hasHours ? `；与上报 ${reportedHours}h ${hit ? '一致' : '不一致，请核对'}` : '';
    return `同一天大表里有：${sameDay.map(span).join('、')}${tail}`;
  }
  const t = dayNumber(startDate);
  let best = null;
  for (const r of mine) {
    const n = dayNumber(r['加班开始日期']);
    if (n === null || t === null) continue;
    const gap = Math.round(n - t);
    if (!best || Math.abs(gap) < Math.abs(best.gap)) best = { gap, text: `${r['加班开始日期']} ${span(r)}` };
  }
  if (!best) return '当天大表里没有，也找不到可比的记录';
  return `当天大表里没有；最近一次是 ${best.text}（${best.gap > 0 ? '+' : ''}${best.gap} 天）`;
}

// 边界检查（只提示，不改执行）：同一笔加班（工号 + 原开始日期 + 原开始时间）在整改表里被写了多种处置
// 背景：校对系统可能对同一笔加班连报两条异常（提醒「多条记录」），班组若分别填了「改」和「删」就会出现
// 执行顺序仍是先修改后删除；删不掉的那条进《定位异常清单》。这里只是提前把这种输入摆到人眼前
function findDispositionConflicts(operations) {
  const byKey = new Map();
  (operations || []).forEach(op => {
    const k = [op['工号'], op['原开始日期'], op['原开始时间']].join('|');
    if (!byKey.has(k)) byKey.set(k, []);
    byKey.get(k).push(op);
  });
  return [...byKey.values()]
    .filter(list => list.length > 1 && new Set(list.map(o => o['操作类型'])).size > 1)
    .map(list => ({
      工号: list[0]['工号'],
      姓名: list[0]['姓名'],
      加班日期: `${list[0]['原开始日期']} ${list[0]['原开始时间']}`,
      处置: list.map(o => o['操作类型']).join(' + '),
    }));
}

// ==================== 步骤 1：导入与合并 ====================

function processGroupWorkbook(parsed) {
  const sheets = [];
  let totalRecords = 0;
  const failures = [];

  parsed.sheetNames.forEach(name => {
    const rows = parsed.sheets[name];
    const dataRows = rows.slice(1).filter(r => r.some(cell => cell !== '' && cell !== undefined));
    const headers = (rows[0] || []).map(h => String(h).trim());
    const rowCount = dataRows.length;
    totalRecords += rowCount;

    const hasRequired = headers.includes('工号') && headers.includes('姓名') && headers.includes('加班开始日期');
    // 表头比对：缺了认识的列 / 出现不认识的列，都提示出来
    // （班组把「科负责人核准」改成别的写法时，旧行为是那列静默变空，页面上看不出任何异常）
    const missingCols = GROUP_HEADERS.filter(h => !headers.includes(h));
    const unknownCols = headers.filter(h => h && !GROUP_HEADERS.includes(h));
    const note = [
      missingCols.length ? `缺少列：${missingCols.join('、')}` : '',
      unknownCols.length ? `不认识的列：${unknownCols.join('、')}（该列不会被读取）` : '',
    ].filter(Boolean).join('；');
    const status = (hasRequired && !note) ? 'ok' : 'warning';

    sheets.push({
      name,
      rowCount,
      status,
      note,
      headers,
      sample: dataRows.slice(0, 3),
    });
  });

  appState.groupWorkbook = parsed;
  appState.groupSheets = sheets;

  let systemNo = 1;
  const merged = [];
  parsed.sheetNames.forEach(name => {
    const rawObjs = sheetToObjects(parsed.sheets[name]);
    const objs = rawObjs.map(obj => {
      const formatted = {};
      Object.keys(obj).forEach(k => { formatted[k] = formatCellValue(obj[k], k); });
      return { raw: obj, formatted };
    });

    objs.forEach(({ raw, formatted }, idx) => {
      const rowNum = idx + 2; // Excel 行号

      // 整行空白：Excel 里常见的空行，直接跳过（既不进合并表，也不计入校验失败）
      if (isBlankRow(formatted)) return;

      const startDate = formatted['加班开始日期'];
      const startTime = normalizeTime(formatted['加班开始时间']);
      const endDate = formatted['加班结束日期'];
      const endTime = normalizeTime(formatted['加班结束时间']);
      const empNo = String(formatted['工号'] || '').trim();
      const empName = formatted['姓名'] || '';

      // 校验必填与格式
      const rowFailures = [];
      if (!empNo) rowFailures.push('工号为空');
      if (!empName) rowFailures.push('姓名为空');
      if (!startDate) rowFailures.push('加班开始日期为空');
      if (!startTime) rowFailures.push('加班开始时间为空');
      if (!endDate) rowFailures.push('加班结束日期为空');
      if (!endTime) rowFailures.push('加班结束时间为空');
      if (startDate && endDate && startTime && endTime) {
        const sd = parseDateParts(startDate);
        const ed = parseDateParts(endDate);
        const st = parseTimeParts(startTime);
        const et = parseTimeParts(endTime);
        if (!sd) rowFailures.push('加班开始日期无效（日期不存在或格式不对）');
        if (!ed) rowFailures.push('加班结束日期无效（日期不存在或格式不对）');
        if (!st) rowFailures.push('加班开始时间格式错误（只能填一个时刻，如 8:15 或 08:15）');
        if (!et) rowFailures.push('加班结束时间格式错误（只能填一个时刻，如 8:15 或 08:15）');
        // 结束早于开始 = 填错（夜班常忘了把结束日期改成次日）→ 退回班组核对，不进合并大表
        // 口径见 docs/排查/M1-问题说明.html 问题 1：不按跨天静默算，也不允许负数进大表
        const span = computeHours(startDate, startTime, endDate, endTime);
        if (typeof span === 'number' && span < 0) {
          rowFailures.push('加班结束时间早于开始时间（请核对结束日期，夜班通常应填次日）');
        }
      }

      // 时数：留空则按时长算出来（要放在失败判定之前 —— 算出来才能一并校验上限）
      let hours = formatted['加班时数'];
      if ((hours === '' || hours === undefined || hours === null) && startDate && startTime && endDate && endTime) {
        hours = computeHoursRaw(raw['加班开始日期'], raw['加班开始时间'], raw['加班结束日期'], raw['加班结束时间']);
      }
      // 时数合理性（数字 / >0 / ≤48，见 overtimeHoursProblem）：不合理退回班组核对，不进大表
      const hoursProblem = overtimeHoursProblem(hours);
      if (hoursProblem) rowFailures.push(hoursProblem);

      if (rowFailures.length) {
        failures.push({
          sheet: name,
          行号: rowNum,
          工号: empNo,
          姓名: empName,
          班组: name,
          失败原因: rowFailures.join('；'),
        });
        return;
      }

      merged.push({
        系统序号: systemNo++,
        原始序号: formatted['序号'] || idx + 1,
        工号: empNo,
        姓名: empName,
        班组: name,
        加班开始日期: startDate,
        加班开始时间: startTime,
        加班结束日期: endDate,
        加班结束时间: endTime,
        加班时数: hours,
        加班原因: formatted['加班原因'] || '',
        加班类别: formatted['加班类别'] || '',
        科负责人核准: formatted['科负责人核准'] || '',
      });
    });
  });

  appState.mergedRecords = merged;
  appState.groupFailures = failures;

  // 重复填报检测：同工号 + 同一天 + 同一开始时间出现多次
  // 常见成因：同一个人被填进了多个班组 sheet，或同一条加班被重复抄录
  // 这类重复会让后续异常/整改的定位出现"多条命中"，越早发现越好
  const dupMap = new Map();
  merged.forEach(r => {
    const k = `${String(r['工号']).trim()}|${toYYYYMMDD(r['加班开始日期'])}|${padTime(r['加班开始时间'])}`;
    if (!dupMap.has(k)) dupMap.set(k, []);
    dupMap.get(k).push(r);
  });
  appState.groupDuplicates = [...dupMap.values()]
    .filter(list => list.length > 1)
    .map(list => ({
      工号: list[0]['工号'],
      姓名: list[0]['姓名'],
      加班开始日期: list[0]['加班开始日期'],
      加班开始时间: list[0]['加班开始时间'],
      条数: list.length,
      系统序号: list.map(r => r['系统序号']).join('、'),
      班组: list.map(r => r['班组']).join('、'),
    }));
  // 重新导入班组表后，之前的批量确认与最终生成状态失效
  appState.batchConfirmed = false;
  appState.finalGenerated = false;
  updateStats();
}

// ==================== 步骤 2：异常处理 ====================

// 整行空白：Excel 的 used range 常带尾部空行（导出商、空格、格式残留都会造成）
// 三条导入路径（班组表 / 异常表 / 整改表）统一用这一条：整行都没内容就当它不存在
// 不跳过的后果：异常表里变成一条“未匹配”，整改表里变成一条“未填写”
//   —— 而 confirmBatch 遇到“未填写”会整轮阻断，提示人去找一条根本不存在的记录
function isBlankRow(rec) {
  return Object.values(rec).every(v => String(v === undefined || v === null ? '' : v).trim() === '');
}

function processAbnormalWorkbook(parsed) {
  // 若当前轮次已确认，自动进入下一轮处理新的异常表
  const current = getCurrentRound();
  if (current.status === 'confirmed') {
    appState.currentRound++;
  }
  const round = getCurrentRound();

  appState.abnormalWorkbook = parsed;
  round.abnormalWorkbook = parsed;
  const records = [];
  const failures = [];
  const warnings = [];

  parsed.sheetNames.forEach(name => {
    const objs = sheetToObjects(parsed.sheets[name]);
    objs.forEach((obj, idx) => {
      const rec = {};
      Object.keys(obj).forEach(h => {
        rec[h] = formatCellValue(obj[h], h);
      });
      if (isBlankRow(rec)) return; // 尾部空行：不算记录，也不进“未匹配”清单
      ABNORMAL_HEADERS.forEach(h => {
        if (!(h in rec)) rec[h] = '';
      });
      rec['班组'] = rec['科室'] || '';
      rec['处置状态'] = '待处理';

      // 匹配校验：定位键 = 工号 + 开始日期（+开始时间，若异常表提供）
      // 姓名不参与定位（异常表姓名可能少字/多字/空白），只在匹配后做一致性提醒
      // 不做"取第一条"式静默匹配：0 条进失败清单，多条进提醒清单，均可在页面与导出文件中核对
      const empNo = String(rec['工号'] || '').trim();
      const empName = String(rec['姓名'] || '').trim();
      const hits = locateMergedRecords(empNo, rec['开始日期'], rec['开始时间']);

      rec['匹配状态'] = hits.length === 0
        ? '未匹配'
        : (hits.length > 1 ? `多条命中(${hits.length})` : '已匹配');

      if (!hits.length) {
        failures.push({
          ...rec,
          行号: idx + 2,
          线索: buildMatchHint(empNo, rec['开始日期'], rec['上报加班时数']),
          失败原因: '无法在合并大表中匹配到对应记录（请核对工号、开始日期/时间）',
        });
      } else {
        if (hits.length > 1) {
          warnings.push({
            ...rec,
            行号: idx + 2,
            定位提醒: `合并大表中存在 ${hits.length} 条同工号同日期记录：${describeCandidates(hits)}，尚未处理${buildDeptHint(hits, rec['科室'])}。若是同一条加班被重复填报到多个班组（一个人只应属于一个班组），请先清理合并大表；若确实是同一天两次加班，请在异常表补填「开始时间」以便唯一定位`,
          });
        }
        if (empName && String(hits[0]['姓名'] || '').trim() !== empName) {
          warnings.push({
            ...rec,
            行号: idx + 2,
            定位提醒: `工号 ${empNo} 在合并大表中的姓名为「${hits[0]['姓名']}」，与异常表填写的「${empName}」不一致，请核对（已按工号+日期+时间匹配成功）`,
          });
        }
      }
      records.push(rec);
    });
  });

  appState.abnormalRecords = records;
  appState.abnormalFailures = failures;
  round.abnormalRecords = records;
  round.abnormalFailures = failures;
  round.abnormalWarnings = warnings;
  round.status = 'processing';
  // 重新导入异常表后，之前的批量确认与最终生成状态失效
  appState.batchConfirmed = false;
  appState.finalGenerated = false;
}

// ==================== 步骤 3：整改与批量操作 ====================

function processRectifyWorkbook(parsed) {
  const round = getCurrentRound();
  appState.rectifyWorkbook = parsed;
  round.rectifyWorkbook = parsed;
  const operations = [];

  parsed.sheetNames.forEach(name => {
    const objs = sheetToObjects(parsed.sheets[name]).map(obj => {
      const rec = {};
      Object.keys(obj).forEach(h => { rec[h] = formatCellValue(obj[h], h); });
      return rec;
    });
    objs.forEach(obj => {
      if (isBlankRow(obj)) return; // 尾部空行：不算操作（否则会被当成“未填写”，整轮被假阻断）
      const type = String(obj['处置方式'] || '').trim();
      let detail = '';
      let opType = type;

      if (!type) {
        // 未填写处置方式：标记为未填写，提示用户补充，不参与批量执行
        opType = '未填写';
        detail = '未填写处置方式，请在整改表中填写后重新导入';
      } else if (type === '修改') {
        detail = `修改后：${obj['修改后开始日期']} ${normalizeTime(obj['修改后开始时间'])}-${normalizeTime(obj['修改后结束时间'])}，${obj['修改后上报加班时数']}h`;
      } else if (type === '删除') {
        detail = '从合并大表中删除该记录';
      } else if (type.includes('不处理')) {
        const shift = obj['调班班次'] || '';
        if (shift) {
          opType = '调班';
          detail = `调班处理：${obj['调班日期']} 导出至 ${shift}`;
        } else {
          opType = '特殊情况';
          detail = `特殊情况不处理：${obj['异常说明（必填）'] || ''}`;
        }
      }

      operations.push({
        // obj['ID'] 是校对系统给的单号（导出整改表时原样带出去、原样带回来），不是大表行号
        校对ID: obj['ID'] || '',
        工号: String(obj['工号'] || '').trim(),
        姓名: obj['姓名'] || '',
        班组: name,
        操作类型: opType,
        操作详情: detail,
        备注: obj['异常说明（必填）'] || '',
        roundNo: round.roundNo,
        // 定位键：整改表自带的原始加班信息（不依赖校对系统 ID）
        原开始日期: obj['开始日期'] || '',
        原开始时间: obj['开始时间'] || '',
        原结束日期: obj['结束日期'] || '',
        原结束时间: obj['结束时间'] || '',
        定位状态: '待定位',
        // 修改后的字段，用于 applyBatchOperations 更新合并大表
        修改后开始日期: obj['修改后开始日期'] || '',
        修改后开始时间: obj['修改后开始时间'] || '',
        修改后结束日期: obj['修改后结束日期'] || '',
        修改后结束时间: obj['修改后结束时间'] || '',
        修改后上报加班时数: obj['修改后上报加班时数'] || '',
        // 调班字段，用于生成本轮调班数据
        调班日期: obj['调班日期'] || '',
        调班班次: obj['调班班次'] || '',
      });
    });
  });

  appState.rectifyOperations = operations;
  round.rectifyOperations = operations;
  // 每次导入新的整改表后，需要重新确认执行批量与最终生成
  appState.batchConfirmed = false;
  appState.finalGenerated = false;
}

// ==================== 页面渲染 ====================

function renderOverview() {
  const wf = getWorkflowStatus();
  const round = getCurrentRound();

  // 流程指引卡片使用实际数据状态，不依赖 finalGenerated
  const flowStates = [
    { key: 'import', label: '导入与合并', subtitle: '读取班组填报表并生成合并大表', done: wf.hasGroup, badge: wf.hasGroup ? '已完成' : '待处理', blockers: wf.hasGroup ? getImportBlockers() : [] },
    { key: 'abnormal', label: '异常处理', subtitle: '导入校对异常表并生成整改表', done: wf.hasAbnormal, badge: wf.hasAbnormal ? `第${round.roundNo}轮` : '待处理', blockers: wf.hasAbnormal ? getAbnormalBlockers() : [] },
    { key: 'rectify', label: '整改与批量操作', subtitle: '导入整改表并执行批量操作', done: wf.hasRectify, badge: wf.hasRectify ? `第${round.roundNo}轮` : '待处理', blockers: wf.hasRectify ? getRectifyBlockers() : [] },
    { key: 'output', label: '输出与审计', subtitle: '导出最终文件并查看操作记录', done: wf.finalGenerated, badge: wf.finalGenerated ? '已完成' : '待生成', blockers: [] },
  ];

  return `
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <div class="lg:col-span-2 bg-apple-card rounded-3xl p-8 shadow-card">
        <h3 class="text-xl font-semibold mb-2 tracking-tight">处理流程</h3>
        <p class="text-sm text-apple-muted mb-6">按以下步骤完成本月加班数据处理，支持多轮异常迭代</p>
        <div class="space-y-4">
          ${flowStates.map((flow, index) => {
            const stepIndex = index + 1;
            const hasBlockers = flow.blockers && flow.blockers.length > 0;
            return `
              <div class="flex items-start gap-4 p-4 rounded-2xl border ${flow.done && !hasBlockers ? 'border-apple-green/20 bg-apple-green/5' : hasBlockers ? 'border-apple-red/20 bg-apple-red/5' : 'border-apple-border bg-apple-gray/30'} hover:bg-apple-gray/50 transition-colors cursor-pointer" onclick="goToStep(${stepIndex})">
                <div class="w-10 h-10 rounded-xl ${flow.done && !hasBlockers ? 'bg-apple-green text-white' : hasBlockers ? 'bg-apple-red text-white' : 'bg-apple-card border border-apple-border text-apple-blue'} flex items-center justify-center shrink-0 font-semibold">
                  ${flow.done && !hasBlockers ? '<i class="ph ph-check"></i>' : hasBlockers ? '<i class="ph ph-warning"></i>' : stepIndex}
                </div>
                <div class="flex-1 min-w-0">
                  <div class="font-semibold flex items-center gap-2">
                    ${flow.label}
                    <span class="badge ${flow.done && !hasBlockers ? 'badge-success' : hasBlockers ? 'badge-danger' : 'badge-muted'} text-xs">${hasBlockers ? '校验未通过' : flow.badge}</span>
                  </div>
                  <div class="text-sm text-apple-muted mt-0.5">${flow.subtitle}</div>
                  ${hasBlockers ? `<div class="mt-2 text-xs text-apple-red space-y-1">${flow.blockers.map(b => `<div>· ${b}</div>`).join('')}</div>` : ''}
                </div>
                <i class="ph ph-caret-right text-apple-muted text-xl"></i>
              </div>
            `;
          }).join('')}
        </div>
      </div>

      <div class="space-y-6">
        <div class="bg-gradient-to-br from-apple-blue to-blue-500 rounded-3xl p-8 text-white shadow-card">
          <div class="text-sm font-medium opacity-90 mb-2">当前工作状态</div>
          <div class="text-2xl font-semibold mb-4">${wf.workingStatus}</div>
          <button onclick="goToStep(${wf.continueStep})" class="w-full h-11 rounded-full bg-white/20 hover:bg-white/30 text-white text-sm font-medium transition-colors backdrop-blur-sm inline-flex items-center justify-center gap-2">
            ${wf.hasGroup ? '继续当前工作' : '开始处理'}
            <i class="ph ph-arrow-right"></i>
          </button>
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-lg font-semibold mb-4 tracking-tight">本月处理规则</h3>
          <div class="space-y-4 text-sm text-apple-muted">
            <div class="flex items-start gap-3">
              <div class="w-8 h-8 rounded-lg bg-apple-blue/10 text-apple-blue flex items-center justify-center shrink-0"><i class="ph ph-number-circle-one text-lg"></i></div>
              <div>
                <div class="font-medium text-apple-text">导入与合并</div>
                <div>以 sheet 名为班组，允许新增 sheet，合并后导出系统格式大表</div>
              </div>
            </div>
            <div class="flex items-start gap-3">
              <div class="w-8 h-8 rounded-lg bg-apple-orange/10 text-apple-orange flex items-center justify-center shrink-0"><i class="ph ph-number-circle-two text-lg"></i></div>
              <div>
                <div class="font-medium text-apple-text">异常处理</div>
                <div>每月可能进行 2-3 轮，每轮导入新的异常表并生成整改表</div>
              </div>
            </div>
            <div class="flex items-start gap-3">
              <div class="w-8 h-8 rounded-lg bg-apple-green/10 text-apple-green flex items-center justify-center shrink-0"><i class="ph ph-number-circle-three text-lg"></i></div>
              <div>
                <div class="font-medium text-apple-text">整改与输出</div>
                <div>支持中间处理输出与最终生成；调班数据在同一会话内按轮次累计（点「重新开始」会清空本次数据）</div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- 本月概览数据看板 -->
    <div class="mt-6 bg-apple-card rounded-3xl p-8 shadow-card">
      <div class="flex items-center justify-between mb-6">
        <h3 class="text-xl font-semibold tracking-tight">本月概览</h3>
        <span class="text-sm text-apple-muted">数据看板</span>
      </div>
      <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-blue/10 text-apple-blue flex items-center justify-center"><i class="ph ph-users text-lg"></i></div>
            <span class="text-sm text-apple-muted">班组数量</span>
          </div>
          <div class="text-3xl font-semibold">${appState.groupSheets.length || 0}</div>
        </div>
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-green/10 text-apple-green flex items-center justify-center"><i class="ph ph-check-circle text-lg"></i></div>
            <span class="text-sm text-apple-muted">总记录</span>
          </div>
          <div class="text-3xl font-semibold">${appState.mergedRecords.length || 0}</div>
        </div>
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-orange/10 text-apple-orange flex items-center justify-center"><i class="ph ph-warning text-lg"></i></div>
            <span class="text-sm text-apple-muted">异常记录</span>
          </div>
          <div class="text-3xl font-semibold">${getAllOperations().length || 0}</div>
        </div>
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-blue/10 text-apple-blue flex items-center justify-center"><i class="ph ph-arrows-clockwise text-lg"></i></div>
            <span class="text-sm text-apple-muted">已处理轮次</span>
          </div>
          <div class="text-3xl font-semibold">${appState.rounds.filter(r => r.status === 'confirmed').length}</div>
        </div>
      </div>
      <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mt-4">
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-blue/10 text-apple-blue flex items-center justify-center"><i class="ph ph-pencil-simple text-lg"></i></div>
            <span class="text-sm text-apple-muted">累计修改</span>
          </div>
          <div class="text-3xl font-semibold">${getAllOperations().filter(o => o['操作类型'] === '修改').length}</div>
        </div>
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-red/10 text-apple-red flex items-center justify-center"><i class="ph ph-trash text-lg"></i></div>
            <span class="text-sm text-apple-muted">累计删除</span>
          </div>
          <div class="text-3xl font-semibold">${getAllOperations().filter(o => o['操作类型'] === '删除').length}</div>
        </div>
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-orange/10 text-apple-orange flex items-center justify-center"><i class="ph ph-calendar-check text-lg"></i></div>
            <span class="text-sm text-apple-muted">累计调班</span>
          </div>
          <div class="text-3xl font-semibold">${getAllShiftRecords().length}</div>
        </div>
        <div class="p-5 rounded-2xl bg-apple-gray/50">
          <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-xl bg-apple-green/10 text-apple-green flex items-center justify-center"><i class="ph ph-info text-lg"></i></div>
            <span class="text-sm text-apple-muted">特殊情况</span>
          </div>
          <div class="text-3xl font-semibold">${getAllOperations().filter(o => o['操作类型'] === '特殊情况').length}</div>
        </div>
      </div>
    </div>
  `;
}

function renderImport() {
  const sheets = appState.groupSheets.length ? appState.groupSheets : demoSheets.map(s => ({ ...s, status: 'ok' }));
  const records = appState.mergedRecords.length ? appState.mergedRecords : demoRecords;
  const hasFile = !!appState.groupWorkbook;
  const isDemo = !hasFile;
  const failureCount = appState.groupFailures.length;
  const duplicates = appState.groupDuplicates || [];

  return `
    ${duplicates.length ? `<div class="mb-6 rounded-3xl p-6 bg-apple-orange/5 border border-apple-orange/20">
      <div class="flex items-start justify-between gap-4">
        <div class="flex items-start gap-3 text-sm text-apple-orange">
          <i class="ph ph-warning-circle mt-0.5 text-lg"></i>
          <div>
            <div class="font-medium">发现 ${duplicates.length} 组重复填报：同一工号同一天同一开始时间有多条记录</div>
            <div class="text-apple-muted mt-1">
              常见原因是同一个人被填进了多个班组（一个人只能属于一个班组），或同一条加班被重复抄录。
              这类重复会让后续「异常处理 / 整改」按工号+日期+时间定位时出现「多条命中」而需要人工核对，建议先在班组填报表里核对并清理。
            </div>
            <div class="mt-2 space-y-1">
              ${duplicates.slice(0, 3).map(d => `<div class="text-xs text-apple-muted">· ${escapeHtml(d['工号'])} ${escapeHtml(d['姓名'])} ${escapeHtml(d['加班开始日期'])} ${escapeHtml(d['加班开始时间'])}：共 ${escapeHtml(d['条数'])} 条（班组：${escapeHtml(d['班组'])}）</div>`).join('')}
              ${duplicates.length > 3 ? `<div class="text-xs text-apple-muted">· 共 ${duplicates.length} 组，其余见下载清单</div>` : ''}
            </div>
          </div>
        </div>
        <button onclick="exportGroupDuplicates()" class="h-9 px-4 rounded-full bg-apple-orange/10 text-apple-orange text-xs font-medium hover:bg-apple-orange/20 transition-colors shrink-0">下载清单</button>
      </div>
    </div>` : ''}
    <div class="grid grid-cols-1 xl:grid-cols-3 gap-6">
      <div class="xl:col-span-1 space-y-6">
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-xl font-semibold mb-2 tracking-tight">上传班组填报表</h3>
          <p class="text-sm text-apple-muted mb-6">支持 .xlsx / .xls 格式，自动识别多个 sheet</p>
          <div class="drop-zone border-2 border-dashed border-apple-border rounded-3xl p-8 text-center cursor-pointer bg-apple-gray/30 hover:bg-apple-gray/50" id="dropZone" data-type="group">
            <input type="file" class="hidden file-input" accept=".xlsx,.xls" />
            <div class="zone-content">
              <div class="w-14 h-14 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center mx-auto mb-4">
                <i class="ph ph-upload-simple text-2xl"></i>
              </div>
              <div class="text-sm font-medium mb-1">拖拽文件到此处</div>
              <div class="text-xs text-apple-muted">或点击选择文件</div>
              ${hasFile ? `<div class="mt-3 text-xs text-apple-green font-medium">已加载：${escapeHtml(appState.fileName)}</div>` : ''}
            </div>
          </div>

          ${hasFile ? `
            <div class="mt-5 space-y-3">
              ${(() => {
                const blockers = getImportBlockers();
                if (blockers.length) {
                  return `
                    <div class="p-4 rounded-2xl bg-apple-red/5 border border-apple-red/20">
                      <div class="flex items-start gap-2 text-sm text-apple-red">
                        <i class="ph ph-warning-circle mt-0.5 text-lg shrink-0"></i>
                        <div>
                          <div class="font-medium">数据校验未通过，无法继续</div>
                          <ul class="mt-1 space-y-1 text-xs">
                            ${blockers.map(b => `<li>· ${b}</li>`).join('')}
                          </ul>
                          <div class="mt-2 text-apple-muted">请修正数据后重新导入班组填报表，全部校验通过后才能继续下一步。</div>
                        </div>
                      </div>
                    </div>
                    <button disabled class="w-full h-11 rounded-full bg-apple-gray text-apple-muted text-sm font-medium cursor-not-allowed inline-flex items-center justify-center gap-2">
                      <i class="ph ph-lock"></i>
                      导出合并大表并继续（需先修正异常）
                    </button>
                  `;
                }
                return `
                  <button onclick="exportMergedAndContinue()" class="w-full h-11 rounded-full bg-apple-blue text-white text-sm font-medium hover:bg-apple-blue-hover transition-colors shadow-sm inline-flex items-center justify-center gap-2">
                    <i class="ph ph-download-simple"></i>
                    导出合并大表并继续
                  </button>
                `;
              })()}
              ${failureCount > 0 ? `
                <button onclick="exportGroupFailures()" class="w-full h-11 rounded-full bg-apple-red/10 text-apple-red text-sm font-medium hover:bg-apple-red/20 transition-colors inline-flex items-center justify-center gap-2">
                  <i class="ph ph-warning"></i>
                  下载校验失败记录 (${failureCount})
                </button>
              ` : ''}
              ${duplicates.length > 0 ? `
                <button onclick="exportGroupDuplicates()" class="w-full h-11 rounded-full bg-apple-orange/10 text-apple-orange text-sm font-medium hover:bg-apple-orange/20 transition-colors inline-flex items-center justify-center gap-2">
                  <i class="ph ph-copy"></i>
                  下载重复填报清单 (${duplicates.length})
                </button>
              ` : ''}
            </div>
          ` : ''}

          <div class="mt-6 flex items-center gap-2 text-xs text-apple-muted">
            <i class="ph ph-info"></i>
            <span>班组以 sheet 名为准，允许新增 sheet</span>
          </div>
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-lg font-semibold mb-4 tracking-tight">校验结果</h3>
          <div class="space-y-3">
            <div class="flex items-center justify-between text-sm">
              <span class="text-apple-muted">识别 sheet</span>
              <span class="font-medium">${sheets.length} 个</span>
            </div>
            <div class="flex items-center justify-between text-sm">
              <span class="text-apple-muted">格式正确</span>
              <span class="font-medium text-apple-green">${sheets.filter(s => s.status === 'ok').length} 个</span>
            </div>
            <div class="flex items-center justify-between text-sm">
              <span class="text-apple-muted">需核对</span>
              <span class="font-medium text-apple-orange">${sheets.filter(s => s.status === 'warning').length} 个</span>
            </div>
            <div class="flex items-center justify-between text-sm">
              <span class="text-apple-muted">重复填报</span>
              <span class="font-medium ${duplicates.length ? 'text-apple-red' : 'text-apple-green'}">${duplicates.length} 组</span>
            </div>
          </div>
        </div>
      </div>

      <div class="xl:col-span-2 space-y-6">
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-6">
            <h3 class="text-xl font-semibold tracking-tight">Sheet 识别预览</h3>
            <div class="flex items-center gap-2">
              ${renderDemoBadge(isDemo)}
              <a href="templates/班组填报表模板.xlsx" download class="h-9 px-4 rounded-full bg-apple-gray text-sm font-medium hover:bg-gray-200 transition-colors inline-flex items-center gap-2">
                <i class="ph ph-download-simple"></i>
                空白模板
              </a>
            </div>
          </div>
          <div class="overflow-hidden rounded-2xl border border-apple-border">
            <table class="data-table bg-white">
              <thead class="bg-apple-gray/50">
                <tr>
                  <th>班组</th>
                  <th>记录数</th>
                  <th>状态</th>
                  <th>字段数</th>
                </tr>
              </thead>
              <tbody>
                ${sheets.map(sheet => `
                  <tr>
                    <td class="font-medium">${sheet.name}</td>
                    <td>${sheet.rowCount ?? sheet.rows}</td>
                    <td><span class="badge ${sheet.status === 'ok' ? 'badge-success' : 'badge-warning'}">${sheet.status === 'ok' ? '正常' : '需核对'}</span>${sheet.note ? `<div class="text-xs text-apple-muted mt-1">${sheet.note}</div>` : ''}</td>
                    <td class="text-apple-muted">${sheet.headers ? sheet.headers.length : '-'}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-6">
            <h3 class="text-xl font-semibold tracking-tight">合并大表预览</h3>
            ${isDemo ? renderDemoBadge(true) : '<span class="text-sm text-apple-muted">前 20 条</span>'}
          </div>
          <div class="overflow-hidden rounded-2xl border border-apple-border">
            ${renderTable(records)}
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderAbnormal() {
  const round = getCurrentRound();
  const records = appState.abnormalRecords.length ? appState.abnormalRecords : demoAbnormal;
  const hasFile = !!appState.abnormalWorkbook;
  const isDemo = !hasFile;
  const failureCount = (round.abnormalFailures || appState.abnormalFailures || []).length;
  const warnings = round.abnormalWarnings || [];
  const noMerged = !appState.mergedRecords.length;

  return `
    ${noMerged ? `<div class="mb-6 rounded-3xl p-6 bg-apple-red/5 border border-apple-red/20">
      <div class="flex items-start gap-3 text-sm text-apple-red">
        <i class="ph ph-warning-circle mt-0.5 text-lg"></i>
        <div>
          <div class="font-medium">尚未导入班组填报表</div>
          <div class="text-apple-muted mt-1">异常记录需要在「合并大表」中按工号+日期+时间定位对应加班记录。请先完成「导入与合并」，再回到本步骤，否则所有记录都会显示为匹配失败。</div>
          <button onclick="goToStep(1)" class="mt-3 h-9 px-4 rounded-full bg-apple-blue text-white text-xs font-medium hover:bg-apple-blue-hover transition-colors">返回步骤 2 导入班组合并</button>
        </div>
      </div>
    </div>` : ''}
    ${warnings.length ? `<div class="mb-6 rounded-3xl p-6 bg-apple-orange/5 border border-apple-orange/20">
      <div class="flex items-start justify-between gap-4">
        <div class="flex items-start gap-3 text-sm text-apple-orange">
          <i class="ph ph-info mt-0.5 text-lg"></i>
          <div>
            <div class="font-medium">有 ${warnings.length} 条记录需要人工核对</div>
            <div class="text-apple-muted mt-1">存在同一工号同一天多条加班（需要确认是哪一条），或异常表姓名与合并大表不一致。请在下方「异常记录清单」核对「匹配状态」列，必要时下载定位提醒交由班组考勤员确认。</div>
          </div>
        </div>
        <button onclick="exportAbnormalWarnings()" class="h-9 px-4 rounded-full bg-apple-orange/10 text-apple-orange text-xs font-medium hover:bg-apple-orange/20 transition-colors shrink-0">下载定位提醒</button>
      </div>
    </div>` : ''}
    <div class="grid grid-cols-1 xl:grid-cols-3 gap-6">
      <div class="xl:col-span-1 space-y-6">
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-2">
            <h3 class="text-xl font-semibold tracking-tight">导入异常表</h3>
            <span class="badge badge-info">第 ${round.roundNo} 轮</span>
          </div>
          <p class="text-sm text-apple-muted mb-6">从校对系统导出的异常记录</p>
          <div class="drop-zone border-2 border-dashed border-apple-border rounded-3xl p-8 text-center cursor-pointer bg-apple-gray/30 hover:bg-apple-gray/50" id="dropZoneAbnormal" data-type="abnormal">
            <input type="file" class="hidden file-input" accept=".xlsx,.xls" />
            <div class="zone-content">
              <div class="w-14 h-14 rounded-2xl bg-apple-orange/10 text-apple-orange flex items-center justify-center mx-auto mb-4">
                <i class="ph ph-warning text-2xl"></i>
              </div>
              <div class="text-sm font-medium mb-1">拖拽异常表到此处</div>
              <div class="text-xs text-apple-muted">或点击选择文件</div>
              ${hasFile ? `<div class="mt-3 text-xs text-apple-green font-medium">已加载：${escapeHtml(appState.abnormalWorkbook.fileName)}</div>` : ''}
            </div>
          </div>

          ${hasFile ? `
            <div class="mt-5 space-y-3">
              ${(() => {
                const blockers = getAbnormalBlockers();
                if (blockers.length) {
                  return `
                    <div class="p-4 rounded-2xl bg-apple-red/5 border border-apple-red/20">
                      <div class="flex items-start gap-2 text-sm text-apple-red">
                        <i class="ph ph-warning-circle mt-0.5 text-lg shrink-0"></i>
                        <div>
                          <div class="font-medium">数据校验未通过，无法继续</div>
                          <ul class="mt-1 space-y-1 text-xs">
                            ${blockers.map(b => `<li>· ${b}</li>`).join('')}
                          </ul>
                          <div class="mt-2 text-apple-muted">请修正异常表数据后重新导入，全部校验通过后才能生成整改表。</div>
                        </div>
                      </div>
                    </div>
                    <button disabled class="w-full h-11 rounded-full bg-apple-gray text-apple-muted text-sm font-medium cursor-not-allowed inline-flex items-center justify-center gap-2">
                      <i class="ph ph-lock"></i>
                      生成整改表（需先修正异常）
                    </button>
                  `;
                }
                return `
                  <button onclick="exportRectify()" class="w-full h-11 rounded-full bg-apple-blue text-white text-sm font-medium hover:bg-apple-blue-hover transition-colors shadow-sm inline-flex items-center justify-center gap-2">
                    <i class="ph ph-file-plus"></i>
                    生成整改表
                  </button>
                `;
              })()}
              ${failureCount > 0 ? `
                <button onclick="exportAbnormalFailures()" class="w-full h-11 rounded-full bg-apple-red/10 text-apple-red text-sm font-medium hover:bg-apple-red/20 transition-colors inline-flex items-center justify-center gap-2">
                  <i class="ph ph-warning"></i>
                  下载匹配失败记录 (${failureCount})
                </button>
              ` : ''}
              ${warnings.length > 0 ? `
                <button onclick="exportAbnormalWarnings()" class="w-full h-11 rounded-full bg-apple-orange/10 text-apple-orange text-sm font-medium hover:bg-apple-orange/20 transition-colors inline-flex items-center justify-center gap-2">
                  <i class="ph ph-info"></i>
                  下载定位提醒 (${warnings.length})
                </button>
              ` : ''}
            </div>
          ` : ''}

          <div class="mt-5 flex items-center gap-2 text-xs text-apple-muted">
            <i class="ph ph-info"></i>
            <a href="templates/异常表模板.xlsx" download class="text-apple-blue hover:underline">下载空白异常表模板</a>
          </div>
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-lg font-semibold mb-4 tracking-tight">匹配统计</h3>
          <div class="space-y-4">
            <div class="flex items-center justify-between p-3 rounded-xl bg-apple-gray/50">
              <span class="text-sm text-apple-muted">异常总数</span>
              <span class="font-semibold">${records.length}</span>
            </div>
            <div class="flex items-center justify-between p-3 rounded-xl bg-apple-gray/50">
              <span class="text-sm text-apple-muted">成功匹配</span>
              <span class="font-semibold text-apple-green">${records.length - failureCount}</span>
            </div>
            <div class="flex items-center justify-between p-3 rounded-xl bg-apple-gray/50">
              <span class="text-sm text-apple-muted">匹配失败</span>
              <span class="font-semibold text-apple-red">${failureCount}</span>
            </div>
            <div class="flex items-center justify-between p-3 rounded-xl bg-apple-gray/50">
              <span class="text-sm text-apple-muted">需人工核对</span>
              <span class="font-semibold text-apple-orange">${warnings.length}</span>
            </div>
          </div>
        </div>
      </div>

      <div class="xl:col-span-2 space-y-6">
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-6">
            <h3 class="text-xl font-semibold tracking-tight">异常记录清单</h3>
            <div class="flex items-center gap-2">
              ${renderDemoBadge(isDemo)}
              <button class="h-9 px-4 rounded-full bg-apple-orange/10 text-apple-orange text-sm font-medium hover:bg-apple-orange/20 transition-colors">${records.length} 条待处理</button>
            </div>
          </div>
          <div class="overflow-hidden rounded-2xl border border-apple-border">
            ${renderTable(records)}
          </div>
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-xl font-semibold mb-4 tracking-tight">整改表生成设置</h3>
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div class="p-4 rounded-2xl bg-apple-gray/50">
              <div class="text-sm text-apple-muted mb-1">拆分方式</div>
              <div class="font-medium">按班组（科室）拆分为多个 sheet</div>
            </div>
            <div class="p-4 rounded-2xl bg-apple-gray/50">
              <div class="text-sm text-apple-muted mb-1">下发方式</div>
              <div class="font-medium">整份文件由班组考勤员自行查找</div>
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderRectify() {
  const round = getCurrentRound();
  const operations = appState.rectifyOperations.length ? appState.rectifyOperations : demoOperations;
  const hasFile = !!appState.rectifyWorkbook;
  const isDemo = !hasFile;
  const locateIssues = round.locateIssues || [];
  const conflicts = findDispositionConflicts(operations);
  const stats = { 修改: 0, 删除: 0, 调班: 0, 特殊情况: 0, 未填写: 0 };
  operations.forEach(op => {
    const t = op['操作类型'];
    if (t === '修改') stats['修改']++;
    else if (t === '删除') stats['删除']++;
    else if (t === '调班') stats['调班']++;
    else if (t === '特殊情况') stats['特殊情况']++;
    else if (t === '未填写') stats['未填写']++;
  });

  return `
    <div class="grid grid-cols-1 xl:grid-cols-3 gap-6">
      <div class="xl:col-span-1 space-y-6">
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-2">
            <h3 class="text-xl font-semibold tracking-tight">导入整改表</h3>
            <span class="badge badge-info">第 ${round.roundNo} 轮</span>
          </div>
          <p class="text-sm text-apple-muted mb-6">班组考勤员填写后发回的整改文件</p>
          <div class="drop-zone border-2 border-dashed border-apple-border rounded-3xl p-8 text-center cursor-pointer bg-apple-gray/30 hover:bg-apple-gray/50" id="dropZoneRectify" data-type="rectify">
            <input type="file" class="hidden file-input" accept=".xlsx,.xls" />
            <div class="zone-content">
              <div class="w-14 h-14 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center mx-auto mb-4">
                <i class="ph ph-download-simple text-2xl"></i>
              </div>
              <div class="text-sm font-medium mb-1">拖拽整改表到此处</div>
              <div class="text-xs text-apple-muted">或点击选择文件</div>
              ${hasFile ? `<div class="mt-3 text-xs text-apple-green font-medium">已加载：${escapeHtml(appState.rectifyWorkbook.fileName)}</div>` : ''}
            </div>
          </div>

          ${hasFile ? `
            <div class="mt-5 space-y-3">
              ${(() => {
                const blockers = getRectifyBlockers();
                if (blockers.length) {
                  return `
                    <div class="p-4 rounded-2xl bg-apple-red/5 border border-apple-red/20">
                      <div class="flex items-start gap-2 text-sm text-apple-red">
                        <i class="ph ph-warning-circle mt-0.5 text-lg shrink-0"></i>
                        <div>
                          <div class="font-medium">数据校验未通过，无法执行</div>
                          <ul class="mt-1 space-y-1 text-xs">
                            ${blockers.map(b => `<li>· ${b}</li>`).join('')}
                          </ul>
                          <div class="mt-2 text-apple-muted">请修正整改表数据后重新导入，全部校验通过后才能执行批量操作。</div>
                        </div>
                      </div>
                    </div>
                    <button disabled class="w-full h-11 rounded-full bg-apple-gray text-apple-muted text-sm font-medium cursor-not-allowed inline-flex items-center justify-center gap-2">
                      <i class="ph ph-lock"></i>
                      确认执行第 ${round.roundNo} 轮批量操作（需先修正异常）
                    </button>
                  `;
                }
                return `
                  <button onclick="confirmBatch()" class="w-full h-11 rounded-full bg-apple-blue text-white text-sm font-medium hover:bg-apple-blue-hover transition-colors shadow-sm inline-flex items-center justify-center gap-2">
                    <i class="ph ph-check-circle"></i>
                    确认执行第 ${round.roundNo} 轮批量操作
                  </button>
                `;
              })()}
            </div>
          ` : ''}

          <div class="mt-5 p-4 rounded-2xl bg-apple-green/5 border border-apple-green/10">
            <div class="flex items-start gap-2 text-sm text-apple-green">
              <i class="ph ph-check-circle mt-0.5"></i>
              <span>无需同步导入合并大表，系统按「工号 + 原开始日期 + 原开始时间」定位。</span>
            </div>
          </div>

          ${stats['未填写'] > 0 ? `
            <div class="mt-4 p-4 rounded-2xl bg-apple-red/5 border border-apple-red/20">
              <div class="flex items-start gap-2 text-sm text-apple-red">
                <i class="ph ph-warning-circle mt-0.5"></i>
                <span>有 ${stats['未填写']} 条记录未填写处置方式，请在整改表中补充填写后重新导入，否则无法执行批量操作。</span>
              </div>
            </div>
          ` : ''}

          ${conflicts.length ? `
            <div class="mt-4 p-4 rounded-2xl bg-apple-orange/5 border border-apple-orange/20">
              <div class="flex items-start gap-2 text-sm text-apple-orange">
                <i class="ph ph-warning mt-0.5"></i>
                <div>
                  <div>有 ${conflicts.length} 笔加班写了<b>两种以上处置</b>：执行时仍按「先修改、再删除」的顺序走，删不掉的那条会进入定位异常清单，请先确认是不是同一笔被重复填报。</div>
                  <div class="text-xs text-apple-muted mt-1">
                    ${conflicts.slice(0, 3).map(c => `· ${c['工号']} ${c['姓名']} ${c['加班日期']}：${c['处置']}`).join('<br>')}
                    ${conflicts.length > 3 ? `<br>· 共 ${conflicts.length} 笔，其余略` : ''}
                  </div>
                </div>
              </div>
            </div>
          ` : ''}
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-4">
            <h3 class="text-lg font-semibold tracking-tight">批量操作定位结果</h3>
            ${round.status === 'confirmed'
              ? (locateIssues.length ? '<span class="badge badge-warning">需人工核对</span>' : '<span class="badge badge-success">全部定位成功</span>')
              : '<span class="badge badge-muted">待确认执行</span>'}
          </div>
          ${round.status !== 'confirmed'
            ? '<div class="text-sm text-apple-muted">尚未执行批量操作。点击上方「确认执行」后，这里会显示每条改动的定位结果。</div>'
            : (locateIssues.length
              ? `<div class="space-y-3">
                  <div class="text-sm text-apple-red">有 ${locateIssues.length} 条操作未执行或未能唯一确定目标行（含无法识别的处置方式），请人工核对后再使用导出文件：</div>
                  ${locateIssues.slice(0, 5).map(i => `
                    <div class="p-3 rounded-2xl bg-apple-red/5 border border-apple-red/10 text-xs">
                      <div class="font-medium text-apple-red">${escapeHtml(i['级别'])} · ${escapeHtml(i['操作类型'])} · 工号 ${escapeHtml(i['工号'])} ${escapeHtml(i['姓名'])}</div>
                      <div class="text-apple-muted mt-1">${escapeHtml(i['说明'])}</div>
                    </div>
                  `).join('')}
                  ${locateIssues.length > 5 ? `<div class="text-xs text-apple-muted">共 ${locateIssues.length} 条，仅显示前 5 条</div>` : ''}
                  <button onclick="exportLocateIssues()" class="w-full h-10 rounded-full bg-apple-red/10 text-apple-red text-sm font-medium hover:bg-apple-red/20 transition-colors inline-flex items-center justify-center gap-2">
                    <i class="ph ph-download-simple"></i>
                    下载定位异常清单 (${locateIssues.length})
                  </button>
                </div>`
              : '<div class="flex items-start gap-2 text-sm text-apple-green"><i class="ph ph-check-circle mt-0.5"></i><span>全部操作均已按「工号 + 开始日期 + 开始时间」唯一定位并执行。</span></div>')}
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-lg font-semibold mb-4 tracking-tight">操作统计</h3>
          <div class="grid grid-cols-2 gap-3">
            <div class="p-4 rounded-2xl bg-apple-gray/50 text-center">
              <div class="text-2xl font-semibold">${stats['修改']}</div>
              <div class="text-xs text-apple-muted mt-1">修改</div>
            </div>
            <div class="p-4 rounded-2xl bg-apple-gray/50 text-center">
              <div class="text-2xl font-semibold">${stats['删除']}</div>
              <div class="text-xs text-apple-muted mt-1">删除</div>
            </div>
            <div class="p-4 rounded-2xl bg-apple-gray/50 text-center">
              <div class="text-2xl font-semibold">${stats['调班']}</div>
              <div class="text-xs text-apple-muted mt-1">调班</div>
            </div>
            <div class="p-4 rounded-2xl bg-apple-gray/50 text-center">
              <div class="text-2xl font-semibold">${stats['特殊情况']}</div>
              <div class="text-xs text-apple-muted mt-1">特殊情况</div>
            </div>
            <div class="p-4 rounded-2xl ${stats['未填写'] > 0 ? 'bg-apple-red/10' : 'bg-apple-gray/50'} text-center">
              <div class="text-2xl font-semibold ${stats['未填写'] > 0 ? 'text-apple-red' : ''}">${stats['未填写']}</div>
              <div class="text-xs ${stats['未填写'] > 0 ? 'text-apple-red' : 'text-apple-muted'} mt-1">未填写</div>
            </div>
          </div>
        </div>
      </div>

      <div class="xl:col-span-2 space-y-6">
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-6">
            <h3 class="text-xl font-semibold tracking-tight">待执行操作预览</h3>
            ${isDemo ? renderDemoBadge(true) : '<span class="text-sm text-apple-muted">前 20 条</span>'}
          </div>
          <div class="overflow-hidden rounded-2xl border border-apple-border">
            ${renderTable(operations)}
          </div>
        </div>

        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-xl font-semibold mb-4 tracking-tight">执行顺序</h3>
          <div class="flex items-center gap-3 text-sm flex-wrap">
            <div class="flex items-center gap-2 px-4 py-2 rounded-full bg-apple-blue/10 text-apple-blue font-medium">
              <span class="w-5 h-5 rounded-full bg-apple-blue text-white text-xs flex items-center justify-center">1</span>
              修改
            </div>
            <i class="ph ph-arrow-right text-apple-muted"></i>
            <div class="flex items-center gap-2 px-4 py-2 rounded-full bg-apple-blue/10 text-apple-blue font-medium">
              <span class="w-5 h-5 rounded-full bg-apple-blue text-white text-xs flex items-center justify-center">2</span>
              删除
            </div>
            <i class="ph ph-arrow-right text-apple-muted"></i>
            <div class="flex items-center gap-2 px-4 py-2 rounded-full bg-apple-blue/10 text-apple-blue font-medium">
              <span class="w-5 h-5 rounded-full bg-apple-blue text-white text-xs flex items-center justify-center">3</span>
              调班导出
            </div>
            <i class="ph ph-arrow-right text-apple-muted"></i>
            <div class="flex items-center gap-2 px-4 py-2 rounded-full bg-apple-muted/10 text-apple-muted font-medium">
              <span class="w-5 h-5 rounded-full bg-apple-muted text-white text-xs flex items-center justify-center">4</span>
              重新编排序号
            </div>
          </div>
        </div>
      </div>
    </div>
  `;
}

function renderOutput() {
  const round = getCurrentRound();
  const confirmed = appState.batchConfirmed && round.status === 'confirmed';
  const isDemo = !round.rectifyOperations.length && !appState.rectifyOperations.length;
  const locateIssues = round.locateIssues || [];

  // 当前轮次统计
  const currentOps = round.rectifyOperations.length ? round.rectifyOperations : (appState.rectifyOperations.length ? appState.rectifyOperations : demoOperations);
  const modify = currentOps.filter(o => o['操作类型'] === '修改').length;
  const del = currentOps.filter(o => o['操作类型'] === '删除').length;
  const shift = currentOps.filter(o => o['操作类型'] === '调班').length;
  const special = currentOps.filter(o => o['操作类型'] === '特殊情况').length;

  // 累计统计
  const allOps = getAllOperations();
  const allShift = getAllShiftRecords();

  // 调班展示：默认只看当前轮次，最终输出时看全部
  const viewShiftMode = appState.outputShiftView || 'current';
  const shiftRecordsToShow = viewShiftMode === 'all'
    ? allShift
    : (round.shiftRecords.length ? round.shiftRecords : buildShiftRecords(currentOps.filter(o => o['操作类型'] === '调班')));

  // 轮次标签
  const roundTabs = appState.rounds
    .map((r, idx) => {
      const active = idx === appState.currentRound;
      const statusIcon = r.status === 'confirmed' ? 'ph-check-circle' : 'ph-circle';
      const statusColor = r.status === 'confirmed' ? 'text-apple-green' : 'text-apple-muted';
      return `
        <button onclick="switchRound(${idx})" class="h-9 px-4 rounded-full text-sm font-medium transition-colors inline-flex items-center gap-2 ${active ? 'bg-apple-blue text-white' : 'bg-apple-gray text-apple-muted hover:text-apple-text'}">
          第${r.roundNo}轮
          <i class="ph ${statusIcon} ${active ? 'text-white' : statusColor}"></i>
        </button>
      `;
    })
    .join('');

  const intermediateDisabled = !confirmed;
  const intermediateTooltip = intermediateDisabled
    ? '需先在步骤4确认执行本轮批量操作'
    : '导出本轮处理后的加班汇总与本轮新增调班数据';

  return `
    <div class="grid grid-cols-1 xl:grid-cols-3 gap-6">
      <div class="xl:col-span-1 space-y-6">
        <!-- 当前轮次信息 -->
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-6">
            <div>
              <h3 class="text-xl font-semibold tracking-tight">第 ${round.roundNo} 轮处理</h3>
              <p class="text-sm text-apple-muted mt-0.5">${confirmed ? '已完成批量操作' : '等待确认执行批量'}</p>
            </div>
            ${confirmed
              ? '<span class="badge badge-success">已确认</span>'
              : '<span class="badge badge-warning">待确认</span>'}
          </div>

          <div class="flex flex-wrap gap-2 mb-6">
            ${roundTabs}
          </div>

          <div class="space-y-3">
            <div class="flex justify-between text-sm p-3 rounded-2xl bg-apple-gray/50">
              <span class="text-apple-muted">修改</span>
              <span class="font-medium text-apple-blue">${modify}</span>
            </div>
            <div class="flex justify-between text-sm p-3 rounded-2xl bg-apple-gray/50">
              <span class="text-apple-muted">删除</span>
              <span class="font-medium text-apple-red">${del}</span>
            </div>
            <div class="flex justify-between text-sm p-3 rounded-2xl bg-apple-gray/50">
              <span class="text-apple-muted">调班</span>
              <span class="font-medium text-apple-orange">${shift}</span>
            </div>
            <div class="flex justify-between text-sm p-3 rounded-2xl bg-apple-gray/50">
              <span class="text-apple-muted">特殊情况</span>
              <span class="font-medium text-apple-green">${special}</span>
            </div>
          </div>
        </div>

        <!-- 输出控制 -->
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-xl font-semibold mb-2 tracking-tight">输出控制</h3>
          <p class="text-sm text-apple-muted mb-6">选择本轮中间输出或全部完成后的最终生成</p>

          <div class="space-y-4">
            <!-- 中间处理输出 -->
            <div class="p-5 rounded-2xl border ${intermediateDisabled ? 'border-apple-border bg-apple-gray/30' : 'border-apple-blue/20 bg-apple-blue/5'} transition-colors">
              <div class="flex items-start gap-3">
                <div class="w-10 h-10 rounded-xl ${intermediateDisabled ? 'bg-apple-gray text-apple-muted' : 'bg-apple-blue text-white'} flex items-center justify-center shrink-0">
                  <i class="ph ph-export text-lg"></i>
                </div>
                <div class="flex-1 min-w-0">
                  <div class="text-sm font-medium">中间处理输出</div>
                  <div class="text-xs text-apple-muted mt-0.5">${intermediateTooltip}</div>
                  <div class="mt-3 flex flex-wrap gap-2">
                    <button onclick="exportIntermediateOutput()" ${intermediateDisabled ? 'disabled' : ''} class="h-9 px-4 rounded-full ${intermediateDisabled ? 'bg-apple-gray text-apple-muted cursor-not-allowed' : 'bg-apple-blue text-white hover:bg-apple-blue-hover'} text-xs font-medium transition-colors">
                      导出本轮加班汇总
                    </button>
                    <button onclick="exportShiftData('current')" ${intermediateDisabled ? 'disabled' : ''} class="h-9 px-4 rounded-full ${intermediateDisabled ? 'bg-apple-gray text-apple-muted cursor-not-allowed' : 'bg-apple-blue/10 text-apple-blue hover:bg-apple-blue/20'} text-xs font-medium transition-colors">
                      导出本轮调班数据
                    </button>
                  </div>
                </div>
              </div>
            </div>

            <!-- 最终生成 -->
            <div class="p-5 rounded-2xl border border-apple-green/20 bg-apple-green/5 transition-colors">
              <div class="flex items-start gap-3">
                <div class="w-10 h-10 rounded-xl bg-apple-green text-white flex items-center justify-center shrink-0">
                  <i class="ph ph-flag-banner-fold text-lg"></i>
                </div>
                <div class="flex-1 min-w-0">
                  <div class="text-sm font-medium">最终生成</div>
                  <div class="text-xs text-apple-muted mt-0.5">导出全部处理后的最终文件组合</div>
                  <div class="mt-3 flex flex-wrap gap-2">
                    <button onclick="exportFinalAll()" class="h-9 px-4 rounded-full bg-apple-green text-white text-xs font-medium hover:bg-green-600 transition-colors inline-flex items-center gap-1.5">
                      <i class="ph ph-download-simple"></i>
                      最终生成
                    </button>
                    <button onclick="exportSystemData('final')" class="h-9 px-4 rounded-full bg-apple-green/10 text-apple-green text-xs font-medium hover:bg-apple-green/20 transition-colors">
                      最终加班汇总
                    </button>
                    <button onclick="exportShiftData('final')" class="h-9 px-4 rounded-full bg-apple-green/10 text-apple-green text-xs font-medium hover:bg-apple-green/20 transition-colors">
                      累计调班数据
                    </button>
                    <button onclick="exportOperationLog()" class="h-9 px-4 rounded-full bg-apple-green/10 text-apple-green text-xs font-medium hover:bg-apple-green/20 transition-colors">
                      操作执行记录
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- 输出文件清单 -->
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <h3 class="text-lg font-semibold mb-4 tracking-tight">输出文件说明</h3>
          <div class="space-y-4 text-sm">
            <div class="p-4 rounded-2xl bg-apple-gray/50">
              <div class="font-medium mb-1">中间处理输出</div>
              <ul class="text-apple-muted space-y-1 text-xs">
                <li>· 总装科加班汇总（第N轮）：本轮处理后的 2007-加班申请</li>
                <li>· 调班数据导入模板（第N轮）：仅含本轮新增调班记录</li>
              </ul>
            </div>
            <div class="p-4 rounded-2xl bg-apple-gray/50">
              <div class="font-medium mb-1">最终生成</div>
              <ul class="text-apple-muted space-y-1 text-xs">
                <li>· 总装科月度加班汇总（最终）：全部轮次处理后的结果</li>
                <li>· 调班数据导入模板（最终）：累计全部轮次调班记录</li>
                <li>· 操作执行记录：全部轮次操作日志</li>
              </ul>
            </div>
          </div>
          <button onclick="startNewRound()" class="w-full mt-6 h-11 rounded-full bg-apple-gray text-sm font-medium hover:bg-gray-200 transition-colors">
            重新开始（清空本次数据）
          </button>
        </div>
      </div>

      <div class="xl:col-span-2 space-y-6">
        <!-- 状态提示 -->
        ${confirmed
          ? `<div class="bg-gradient-to-r from-apple-green/5 to-apple-blue/5 rounded-3xl p-8 border border-apple-border">
              <div class="flex items-start gap-4">
                <div class="w-12 h-12 rounded-2xl bg-apple-green/10 text-apple-green flex items-center justify-center shrink-0">
                  <i class="ph ph-check-circle text-2xl"></i>
                </div>
                <div>
                  <h3 class="text-lg font-semibold mb-1">第 ${round.roundNo} 轮处理完成</h3>
                  <p class="text-sm text-apple-muted leading-relaxed">
                    本轮批量操作已执行。可选择「中间处理输出」进入下一轮迭代，或在全部轮次完成后点击「最终生成」。
                  </p>
                </div>
              </div>
            </div>`
          : `<div class="bg-gradient-to-r from-apple-orange/5 to-apple-red/5 rounded-3xl p-8 border border-apple-border">
              <div class="flex items-start gap-4">
                <div class="w-12 h-12 rounded-2xl bg-apple-orange/10 text-apple-orange flex items-center justify-center shrink-0">
                  <i class="ph ph-hourglass text-2xl"></i>
                </div>
                <div>
                  <h3 class="text-lg font-semibold mb-1">第 ${round.roundNo} 轮处理未完成</h3>
                  <p class="text-sm text-apple-muted leading-relaxed">
                    请返回「整改与批量操作」步骤，点击「确认执行批量操作」后，本轮输出文件才允许导出。
                  </p>
                  <button onclick="goToStep(3)" class="mt-4 h-10 px-5 rounded-full bg-apple-blue text-white text-sm font-medium hover:bg-apple-blue-hover transition-colors shadow-sm">
                    返回步骤4
                  </button>
                </div>
              </div>
            </div>`}

        <!-- 定位异常提示：批量操作未能唯一确定目标行的记录 -->
        ${confirmed && locateIssues.length ? `
        <div class="bg-apple-card rounded-3xl p-8 shadow-card border border-apple-red/20">
          <div class="flex items-start justify-between gap-4 mb-4">
            <div class="flex items-start gap-3">
              <div class="w-10 h-10 rounded-xl bg-apple-red/10 text-apple-red flex items-center justify-center shrink-0">
                <i class="ph ph-warning-octagon text-lg"></i>
              </div>
              <div>
                <h3 class="text-lg font-semibold tracking-tight">有 ${locateIssues.length} 条操作未执行或需人工核对</h3>
                <p class="text-sm text-apple-muted mt-0.5">这些改动可能未生效或作用于错误行，请核对后再使用导出文件。</p>
              </div>
            </div>
            <button onclick="exportLocateIssues()" class="h-9 px-4 rounded-full bg-apple-red/10 text-apple-red text-xs font-medium hover:bg-apple-red/20 transition-colors shrink-0">
              下载清单
            </button>
          </div>
          <div class="space-y-2">
            ${locateIssues.slice(0, 5).map(i => `
              <div class="p-3 rounded-2xl bg-apple-red/5 text-xs">
                <div class="font-medium text-apple-red">${escapeHtml(i['级别'])} · ${escapeHtml(i['操作类型'])} · 工号 ${escapeHtml(i['工号'])} ${escapeHtml(i['姓名'])}（校对ID ${escapeHtml(i['校对ID'] || '-')}）</div>
                <div class="text-apple-muted mt-1">${escapeHtml(i['说明'])}</div>
              </div>
            `).join('')}
            ${locateIssues.length > 5 ? `<div class="text-xs text-apple-muted">共 ${locateIssues.length} 条，仅显示前 5 条</div>` : ''}
          </div>
        </div>` : ''}

        <!-- 操作执行记录 -->
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-6">
            <h3 class="text-xl font-semibold tracking-tight">操作执行记录</h3>
            <div class="flex items-center gap-2">
              ${renderDemoBadge(isDemo)}
              <button onclick="setOutputShiftView('current')" class="h-9 px-4 rounded-full ${viewShiftMode === 'current' ? 'bg-apple-blue/10 text-apple-blue' : 'bg-apple-gray text-apple-muted hover:text-apple-text'} text-sm font-medium transition-colors">只看本轮</button>
              <button onclick="setOutputShiftView('all')" class="h-9 px-4 rounded-full ${viewShiftMode === 'all' ? 'bg-apple-blue/10 text-apple-blue' : 'bg-apple-gray text-apple-muted hover:text-apple-text'} text-sm font-medium transition-colors">累计全部</button>
            </div>
          </div>
          <div class="overflow-hidden rounded-2xl border border-apple-border">
            ${renderTable(currentOps)}
          </div>
        </div>

        <!-- 调班记录 -->
        <div class="bg-apple-card rounded-3xl p-8 shadow-card">
          <div class="flex items-center justify-between mb-6">
            <h3 class="text-xl font-semibold tracking-tight">调班记录</h3>
            <span class="text-sm text-apple-muted">${shiftRecordsToShow.length} 条 · ${viewShiftMode === 'all' ? '累计全部' : '仅本轮'}</span>
          </div>
          <div class="overflow-hidden rounded-2xl border border-apple-border">
            ${shiftRecordsToShow.length
              ? renderTable(shiftRecordsToShow)
              : `<div class="p-12 text-center text-sm text-apple-muted">暂无调班记录</div>`}
          </div>
        </div>
      </div>

    </div>
  `;
}

function switchRound(index) {
  appState.currentRound = index;
  const round = appState.rounds[index];
  appState.abnormalRecords = round ? round.abnormalRecords : [];
  appState.rectifyOperations = round ? round.rectifyOperations : [];
  appState.abnormalWorkbook = round ? round.abnormalWorkbook : null;
  appState.rectifyWorkbook = round ? round.rectifyWorkbook : null;
  appState.abnormalFailures = round ? (round.abnormalFailures || []) : [];
  appState.batchConfirmed = round ? round.status === 'confirmed' : false;
  renderContent();
}

function setOutputShiftView(mode) {
  appState.outputShiftView = mode;
  renderContent();
}

// 中间处理输出：同时导出本轮加班汇总 + 本轮调班数据
function exportIntermediateOutput() {
  exportSystemData('current');
  setTimeout(() => exportShiftData('current'), 600);
}

function getOpBadgeClass(type) {
  switch (type) {
    case '修改': return 'badge-info';
    case '删除': return 'badge-danger';
    case '导出调班':
    case '调班': return 'badge-warning';
    case '特殊情况': return 'badge-muted';
    default: return 'badge-info';
  }
}

// 步骤4：确认执行批量操作，之后步骤5的最终输出文件才视为已生成
function confirmBatch() {
  const round = getCurrentRound();
  // 本轮已执行过：必须禁止重复执行
  // 同一批操作做第二遍会写在已经被改过的数据上（删除会重排大表序号、修改会改掉原业务键），结果正确性无法保证
  // —— 宁可拦住，让人确认后再开新一轮
  if (round.status === 'confirmed' || appState.batchConfirmed) {
    // 提示语要与 startNewRound 的真实行为一致：它是整场重置（大表、异常表、整改表、轮次历史都会清空）
    showToast('本轮批量操作已执行过，不能重复执行。如需重做，请先点「重新开始（清空本次数据）」，再按 班组表 → 异常表 → 整改表 重新导入', 'error');
    return;
  }

  // 本轮还没导入异常表：《整改表》是从异常表那一步导出来的，跳过它执行 = 拿一份来源不明的表改大表
  // 真实风险：拿错 / 拿上个月的整改表直接套到本月大表上（改删都不可逆，且无提示）
  if (!round.abnormalRecords.length) {
    showToast('本轮还没导入异常表，不能执行。请先完成第 2 步：导入异常表并生成整改表', 'error');
    return;
  }

  // 没有导入整改表时不允许执行（避免把页面上的示例操作当成真实操作执行）
  if (!appState.rectifyOperations.length) {
    showToast('尚未导入整改表，请先导入班组考勤员填好并发回的文件', 'error');
    return;
  }

  const ops = appState.rectifyOperations;

  // 流程约束：整改表有未填写处置方式或修改后时间异常时阻止执行
  const blockers = getRectifyBlockers();
  if (blockers.length) {
    showToast(`数据校验未通过，无法执行批量操作：\n${blockers.join('\n')}`, 'error');
    return;
  }

  // 存在未填写处置方式的记录时，阻止执行并提示补充
  const unfilled = ops.filter(o => o['操作类型'] === '未填写').length;
  if (unfilled > 0) {
    showToast(`有 ${unfilled} 条记录未填写处置方式，请补充填写后重新导入整改表`, 'error');
    return;
  }

  // M4-4：真要动删除/调班之前，给一次后悔机会（原生确认框；点「取消」什么都不做）
  // 只有修改时不打扰（修改可再改回来，删除不可逆）
  const removeCount = ops.filter(o => o['操作类型'] === '删除' || o['操作类型'] === '调班').length;
  if (removeCount > 0 && typeof window !== 'undefined' && typeof window.confirm === 'function') {
    const modifyCount = ops.filter(o => o['操作类型'] === '修改').length;
    const ok = window.confirm(
      `本轮将执行：修改 ${modifyCount} 行，删除 / 调班 ${removeCount} 行。\n删除不可撤销（删了要重新导入班组表重建），确定执行？`
    );
    if (!ok) {
      showToast('已取消执行，数据未改动。可再核对一遍整改表后重新点执行', 'info');
      return;
    }
  }

  // 执行实际的批量操作：修改/删除/调班
  // 未定位/多条命中的操作不再静默跳过，而是进入本轮定位异常清单
  const result = applyBatchOperations(ops);
  round.locateIssues = result.issues;

  // 生成本轮处理后的系统数据快照
  round.systemRecords = buildSystemRecords(appState.mergedRecords);

  // 生成本轮调班数据
  round.shiftRecords = buildShiftRecords(ops.filter(o => o['操作类型'] === '调班'));

  round.status = 'confirmed';
  round.confirmedAt = new Date().toISOString();
  appState.batchConfirmed = true;

  if (result.issues.length) {
    showToast(`第 ${round.roundNo} 轮已执行，但有 ${result.issues.length} 条操作未能唯一定位，请在「输出与审计」核对并下载定位异常清单`, 'warning');
  } else {
    showToast(`第 ${round.roundNo} 轮批量操作已执行，正在生成输出文件`, 'success');
  }
  goToStep(4);
}

// 将整改操作应用到合并大表
// 定位只用业务键（工号 + 原开始日期 + 原开始时间），见 resolveOperationTarget：序号不是身份，定位不到就进清单
// 返回 { applied, issues }；未定位/多条命中的操作会记录到 issues，不再静默跳过
function applyBatchOperations(operations) {
  const issues = [];
  const needLocate = operations.filter(o => o['操作类型'] === '修改' || o['操作类型'] === '删除' || o['操作类型'] === '调班');

  if (!appState.mergedRecords.length) {
    needLocate.forEach(op => {
      op['定位状态'] = '未定位';
      issues.push(buildLocateIssue(op, null, '未定位', '合并大表为空，操作未执行。请先完成步骤 2 的导入与合并'));
    });
    return { applied: 0, issues };
  }

  // 先处理修改，再处理删除/调班
  const modifyOps = operations.filter(o => o['操作类型'] === '修改');
  const removeOps = operations.filter(o => o['操作类型'] === '删除' || o['操作类型'] === '调班');

  modifyOps.forEach(op => {
    const resolved = resolveOperationTarget(op);
    if (!resolved.target) {
      op['定位状态'] = '未定位';
      issues.push(buildLocateIssue(op, resolved, '未定位', '合并大表中找不到该条加班记录（工号+开始日期+开始时间），修改未执行。请核对该行原始信息，或确认该记录是否已被本批次前一条操作改动'));
      return;
    }
    if (resolved.hits > 1) {
      // 同工号同日期多条 = 挑不准是哪一条：宁可不改（改错行会算错当天工时），退回让人补「原开始时间」
      op['定位状态'] = `多条命中(${resolved.hits})`;
      issues.push(buildLocateIssue(op, resolved, '多条命中',
        `合并大表中存在 ${resolved.hits} 条同工号同日期记录：${describeCandidates(resolved.candidates)}，修改未执行。请补填「原开始时间」以唯一定位后重新导入${buildDeptHint(resolved.candidates, op['班组'])}`));
      return;
    }
    op['定位状态'] = `已定位(${resolved.method})`;

    const target = resolved.target;
    const startDate = normalizeDate(op['修改后开始日期']);
    const startTime = normalizeTime(op['修改后开始时间']);
    const endDate = normalizeDate(op['修改后结束日期']);
    const endTime = normalizeTime(op['修改后结束时间']);
    // 日期必须是日历上真存在的那天（2 月 30 日这类会被 JS 悄悄挪走）→ 不执行，退回人工核对
    if ((startDate && !parseDateParts(startDate)) || (endDate && !parseDateParts(endDate))) {
      op['定位状态'] = '日期不合理';
      issues.push(buildLocateIssue(op, resolved, '日期不合理',
        `修改后日期不存在（开始 ${startDate || '-'}，结束 ${endDate || '-'}），修改未执行。请核对日期`));
      return;
    }
    // 时间必须是单个时刻（一格写两个时间/带备注都不认）→ 不执行，退回人工核对
    if ((startTime && !parseTimeParts(startTime)) || (endTime && !parseTimeParts(endTime))) {
      op['定位状态'] = '时间不合理';
      issues.push(buildLocateIssue(op, resolved, '时间不合理',
        `修改后时间不是一个时刻（开始 ${startTime || '-'}，结束 ${endTime || '-'}），修改未执行。请一格只填一个时刻，如 8:15`));
      return;
    }
    // 结束早于开始 = 填错 → 不执行这条修改，退回人工核对（与「未定位」同一套处理）
    const span = (startDate && startTime && endDate && endTime)
      ? computeHours(startDate, startTime, endDate, endTime) : '';
    if (typeof span === 'number' && span < 0) {
      op['定位状态'] = '时间不合理';
      issues.push(buildLocateIssue(op, resolved, '时间不合理',
        `修改后结束时间早于开始时间（${endDate} ${endTime} 早于 ${startDate} ${startTime}），修改未执行。请核对结束日期，夜班通常应填次日`));
      return;
    }
    let hours = op['修改后上报加班时数'];
    if ((hours === '' || hours === undefined || hours === null) && span !== '') {
      hours = span;
    }
    // 时数合理性（数字 / >0 / ≤48）：不合理就不执行这条修改，退回人工核对
    const hoursProblem = overtimeHoursProblem(hours);
    if (hoursProblem) {
      op['定位状态'] = '时数不合理';
      issues.push(buildLocateIssue(op, resolved, '时数不合理', `修改后${hoursProblem}，修改未执行。请核对是否填错`));
      return;
    }
    if (startDate) target['加班开始日期'] = startDate;
    if (startTime) target['加班开始时间'] = startTime;
    if (endDate) target['加班结束日期'] = endDate;
    if (endTime) target['加班结束时间'] = endTime;
    if (hours !== '' && hours !== undefined && hours !== null) target['加班时数'] = hours;
    target['操作标记'] = `第 ${op.roundNo || appState.currentRound + 1} 轮修改`;
  });

  const removeTargets = new Set();
  removeOps.forEach(op => {
    const resolved = resolveOperationTarget(op);
    if (!resolved.target) {
      op['定位状态'] = '未定位';
      issues.push(buildLocateIssue(op, resolved, '未定位', `合并大表中找不到该条加班记录（工号+开始日期+开始时间），${op['操作类型']}未执行。请核对该行原始信息，或确认该记录是否已被本批次前一条操作改动`));
      return;
    }
    if (resolved.hits > 1) {
      // 同上：删错行不可逆，宁可不删；人补齐「原开始时间」后重导自然能唯一定位
      op['定位状态'] = `多条命中(${resolved.hits})`;
      issues.push(buildLocateIssue(op, resolved, '多条命中',
        `合并大表中存在 ${resolved.hits} 条同工号同日期记录：${describeCandidates(resolved.candidates)}，${op['操作类型']}未执行。请补填「原开始时间」以唯一定位后重新导入${buildDeptHint(resolved.candidates, op['班组'])}`));
      return;
    }
    op['定位状态'] = `已定位(${resolved.method})`;
    removeTargets.add(resolved.target);
  });

  if (removeTargets.size) {
    appState.mergedRecords = appState.mergedRecords.filter(r => !removeTargets.has(r));
    // 重新编排序号
    appState.mergedRecords.forEach((r, idx) => {
      r['系统序号'] = idx + 1;
    });
  }

  // 特殊情况不涉及合并大表
  operations.forEach(op => {
    if (op['操作类型'] === '特殊情况') op['定位状态'] = '无需定位';
  });

  // 无法识别的处置方式（班组考勤员写的自由文本，如"删除加班""已改""已调"）不再静默忽略
  const KNOWN_OPS = ['修改', '删除', '调班', '特殊情况', '未填写'];
  operations.forEach(op => {
    if (KNOWN_OPS.includes(op['操作类型'])) return;
    op['定位状态'] = '未识别';
    issues.push(buildLocateIssue(op, null, '未识别',
      `处置方式「${op['操作类型']}」无法识别，未执行。请在整改表里改成 修改 / 删除 / 调班 / 不处理 之一后重新导入`));
  });

  return { applied: operations.length, issues };
}

// 根据合并大表构建系统输出格式的记录
function buildSystemRecords(records) {
  return records.map((r, i) => {
    const start = toYYYYMMDD(r['加班开始日期']);
    const end = toYYYYMMDD(r['加班结束日期']);
    return {
      中文名称: i + 1,
      工号: r['工号'],
      姓名: r['姓名'],
      开始日期: start,
      结束日期: end,
      类型: '10 已核准的加班',
      // M5-6：时间补零（真实上传文件 13265 行全是两位小时）；带秒的值顺带截到分（M1 记过的升级路径）
      开始时间: padTime(r['加班开始时间']),
      结束时间: padTime(r['加班结束时间']),
      // M5-2：定额量归一成数字（文本形式的数字也变数字）
      定额量: toHourNumber(r['加班时数']),
      加班报酬类型: '1 支付加班费',
      加班原因: r['加班原因'] || '',
    };
  });
}

// 根据调班操作构建调班数据记录
function buildShiftRecords(shiftOps) {
  // 只导出「真执行过」的调班：定位状态以「已定位」开头（口径同《操作执行记录》，见 exportOperationLog 注释）
  // 没执行的（未定位 / 多条命中）不能外发：否则调班表说"这天调走了"、大表里那笔加班还在，两份输出互相打架
  return shiftOps.filter(op => String(op['定位状态'] || '').startsWith('已定位')).map((op, i) => {
    const detail = op['操作详情'] || '';
    const dateMatch = detail.match(/(\d{8})/);
    const codeMatch = detail.match(/(SF\w+|OFF|NS)/);
    const date = normalizeDate(op['调班日期']) || (dateMatch ? dateMatch[1] : '20260801');
    const code = op['调班班次'] || (codeMatch ? `${codeMatch[1]} 调班` : 'OFF 休息');
    return {
      序号: i + 1,
      工号: op['工号'],
      姓名: op['姓名'],
      开始日期: toYYYYMMDD(date),
      结束日期: toYYYYMMDD(date),
      日工作计划: code,
      出勤项目分类: '已反馈出勤计划',
      备注: detail,
      roundNo: op.roundNo || appState.currentRound + 1,
    };
  });
}

// 开始新一轮处理，重置会话状态
function startNewRound() {
  // M7-1：这不是“再来一轮”，是把本次会话整个清空（大表、异常表、整改表、轮次记录）—— 先让人知情
  if (typeof window !== 'undefined' && typeof window.confirm === 'function'
    && !window.confirm('这会清空本次已导入的全部数据（班组表、合并大表、异常表、整改表、轮次记录），确定重新开始？')) {
    showToast('已取消，数据未改动', 'info');
    return;
  }
  appState.groupWorkbook = null;
  appState.groupSheets = [];
  appState.mergedRecords = [];
  appState.abnormalWorkbook = null;
  appState.abnormalRecords = [];
  appState.rectifyWorkbook = null;
  appState.rectifyOperations = [];
  appState.batchConfirmed = false;
  appState.finalGenerated = false;
  appState.rounds = [];
  appState.currentRound = 0;
  appState.groupFailures = [];
  appState.abnormalFailures = [];
  appState.groupDuplicates = [];
  appState.outputShiftView = 'current';   // M7-2：重置漏了这个（调班记录的本轮/累计视图）
  appState.isParsing = false;             // M7-2：重置漏了这个（解析中标志）
  appState.fileName = '';
  updateStats();
  goToStep(0);
  showToast('已重置会话，可开始新一轮处理', 'info');
}

// ==================== 客户端导出 ====================

function downloadWorkbook(wb, filename) {
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  showToast(`已导出 ${filename}`, 'success');
}

// mode: 'current' 导出本轮处理后的汇总；'final' 导出全部处理后的最终汇总
function exportSystemData(mode = 'current') {
  const round = getCurrentRound();
  if (!appState.mergedRecords.length) {
    showToast('暂无数据，请先导入班组填报表', 'error');
    return false;
  }
  let records;
  let filename;

  if (mode === 'final') {
    records = appState.mergedRecords;
    filename = '总装科月度加班汇总_最终.xlsx';
  } else if (mode === 'merged') {
    // 步骤2：直接导出当前合并大表，尚未经过任何异常处理
    records = appState.mergedRecords;
    filename = '总装科月度加班汇总_合并大表.xlsx';
  } else {
    records = round.systemRecords.length
      ? round.systemRecords.map(r => ({
          工号: r['工号'],
          姓名: r['姓名'],
          加班开始日期: r['开始日期'],
          加班结束日期: r['结束日期'],
          加班开始时间: r['开始时间'],
          加班结束时间: r['结束时间'],
          加班时数: r['定额量'],
          加班原因: r['加班原因'],
        }))
      : appState.mergedRecords;
    filename = `总装科月度加班汇总_第${round.roundNo}轮.xlsx`;
  }

  // 2007 的行只由 buildSystemRecords 一个地方生产（快照与导出共用）
  // —— 以前这里另抄了一份，导致「时间补零 / 定额量归一」只改到快照、没改到导出（M5-2 / M5-6 就是这么漏的）
  const rows = buildSystemRecords(records).map(r => SYSTEM_OUTPUT_HEADERS.map(h => (r[h] === undefined ? '' : r[h])));

  // 前 6 行表头结构（严格匹配上传系统模板格式）
  const headerRows = [
    SYSTEM_OUTPUT_HEADERS,                                                         // 第 1 行：中文名称
    ['选项值', '', '', '', '', 'T556P', '', '', '', 'T555R', ''],                 // 第 2 行：选项值
    ['是否必填', '', '', 'X', 'X', 'X', 'X', 'X', 'X', 'X', 'X'],                 // 第 3 行：是否必填
    ['填报说明', '使用新方案的8位数字工号',
     '此字段作为校验字段，如与个人基本信息内姓名不一致，导入会报错',
     'YYYYMMDD格式', 'YYYYMMDD格式', '', 'HHMM格式', 'HHMM格式',
     '时数', '选项', '自由填写'],                                                    // 第 4 行：填报说明
    ['取值举例', '163161', '谭文杰', '20200101', '20200101',
     '', '0845', '1730', '5', '3 补休', '测试'],                                    // 第 5 行：取值举例
    ['序列号', 'RP50G-PERNR', '', 'P2007-BEGDA', 'P2007-ENDDA',
     'P2007-KTART', 'P2007-BEGUZ', 'P2007-ENDUZ',
     'P2007-ANZHL', 'P2007-VERSL', 'RESON']                                        // 第 6 行：序列号
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([...headerRows, ...rows]);
  XLSX.utils.book_append_sheet(wb, ws, '2007-加班申请');

  const codeSheet = XLSX.utils.aoa_to_sheet([
    ['类型码', '加班报酬类型码'],
    ['10 已核准的加班', '1 支付加班费'],
    ['', '3 补休']
  ]);
  XLSX.utils.book_append_sheet(wb, codeSheet, '码表');

  downloadWorkbook(wb, filename);
}

// 步骤2：导出合并大表（系统模板格式）并继续到异常处理
function exportMergedAndContinue() {
  // 流程约束：有校验失败或重复填报时不允许继续
  const blockers = getImportBlockers();
  if (blockers.length) {
    showToast(`数据校验未通过，无法继续：\n${blockers.join('\n')}`, 'error');
    return;
  }
  if (!appState.mergedRecords.length) {
    showToast('请先导入班组填报表', 'error');
    return;
  }
  exportSystemData('merged');
  setTimeout(() => {
    goToStep(2);
    showToast('合并大表已导出，进入异常处理', 'success');
  }, 400);
}

function exportRectify() {
  // 流程约束：有匹配失败或需人工核对的异常时不允许生成整改表
  const blockers = getAbnormalBlockers();
  if (blockers.length) {
    showToast(`数据校验未通过，无法生成整改表：\n${blockers.join('\n')}`, 'error');
    return false;
  }
  if (!appState.abnormalRecords.length) {
    showToast('尚无异常记录，请先导入异常表', 'error');
    return false;
  }
  const allRecords = appState.abnormalRecords;
  // 匹配失败的记录没有对应系统序号，无法执行整改，不进入整改表
  const records = allRecords.filter(r => r['匹配状态'] !== '未匹配');
  const excluded = allRecords.length - records.length;
  if (excluded > 0) {
    showToast(`已排除 ${excluded} 条匹配失败记录（明细见「下载匹配失败记录」按钮），整改表仅包含匹配成功的记录`, 'warning');
  }
  // 按班组/科室分组
  const groups = {};
  records.forEach(r => {
    const group = r['科室'] || r['班组'] || '未分组';
    if (!groups[group]) groups[group] = [];
    groups[group].push(r);
  });

  const headers = [...ABNORMAL_HEADERS, ...RECTIFY_ACTION_HEADERS];
  const wb = XLSX.utils.book_new();
  Object.keys(groups).forEach(group => {
    const rows = groups[group].map(r => {
      return headers.map(h => {
        // ID 列原样导出校对系统单号：不要拿大表系统序号覆盖它
        // （序号只是当次会话的排号，每次导入重发、删除后重排；覆盖掉单号后班组与考勤员就没法跟校对系统对账）
        if (ABNORMAL_HEADERS.includes(h)) return r[h] !== undefined ? r[h] : '';
        return '';
      });
    });
    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    XLSX.utils.book_append_sheet(wb, ws, group.slice(0, 31));
  });

  downloadWorkbook(wb, '异常整改表.xlsx');
}

// mode: 'current' 导出本轮调班数据；'final' 导出累计全部调班数据
function exportShiftData(mode = 'current') {
  let shiftRecords;
  let filename;

  if (mode === 'final') {
    shiftRecords = getAllShiftRecords();
    filename = '调班数据导入模板_最终.xlsx';
  } else {
    const round = getCurrentRound();
    shiftRecords = round.shiftRecords.length
      ? round.shiftRecords
      : buildShiftRecords(appState.rectifyOperations.filter(o => o['操作类型'] === '调班'));
    filename = `调班数据导入模板_第${round.roundNo}轮.xlsx`;
  }

  if (!shiftRecords.length) {
    showToast(mode === 'final' ? '暂无累计调班记录' : '本轮没有调班记录', 'error');
    return false;
  }

  const mainRows = [];
  const sheet2Rows = [];
  const names = {};

  shiftRecords.forEach((rec, i) => {
    const seq = i + 1;
    mainRows.push([seq, rec['工号'], rec['姓名'], rec['开始日期'], rec['结束日期'], rec['日工作计划']]);
    sheet2Rows.push([seq, rec['工号'], rec['姓名'], rec['开始日期'], rec['结束日期'], rec['日工作计划'], rec['出勤项目分类'] || '已反馈出勤计划', rec['备注'] || '']);
    names[rec['工号']] = rec['姓名'];
  });

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([SHIFT_MAIN_HEADERS, ...mainRows]), '2003-调班信息');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([SHIFT_SHEET2_HEADERS, ...sheet2Rows]), 'Sheet2');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['班次代码', '上班时间', '备注'],
    ['SF04 双班早班', '7:00-15:45', ''],
    ['SF17 固定班', '8:45-17:30', ''],
    ['SF05 双班中班', '15:45-00:20', ''],
    ['SF10 二线中班1545', '15:45-00:30', ''],
    ['SF11 二线中班1645', '16:45-1:30', ''],
    ['SF12 二线中班1755', '17:55-2:40', '2:40-3:55'],
    ['SF13 二线中班1845', '18:45-3:45', '']
  ]), '上班时间段');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['工号', '姓名'],
    ...Object.keys(names).map(id => [id, names[id]])
  ]), '名单');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), 'Sheet1');

  const codeWs = XLSX.utils.aoa_to_sheet([
    ['班次代码'],
    ['NS 未排班'],
    ['OFF 休息'],
    ['SF01 正常班'],
    ['SF02 动总二厂正常班'],
    ['SF04 双班早班'],
    ['SF05 双班中班'],
    ['SF17 固定班'],
    ['SF28 一部改造期早班'],
    ['SF29 一部改造期中班'],
    ['X260 自定义班次1']
  ]);
  XLSX.utils.book_append_sheet(wb, codeWs, '码表');

  downloadWorkbook(wb, filename);
}

function exportOperationLog() {
  const ops = getAllOperations();
  if (!ops.length) {
    showToast('暂无操作记录，请先执行批量操作', 'error');
    return false;
  }
  // 定位状态列：只有「已定位(…)」是真执行过的；未定位 / 多条命中 / 未识别 / 待定位 都是"没动手"
  // （M2-3 起多条命中不再挑一条执行，所以这份记录必须能区分"记了"和"做了"）
  const headers = ['轮次', '校对ID', '工号', '姓名', '班组', '操作类型', '定位状态', '操作详情', '备注'];
  const rows = ops.map(op => [
    op['roundNo'] || 1, op['校对ID'], op['工号'], op['姓名'], op['班组'], op['操作类型'], op['定位状态'] || '', op['操作详情'], op['备注'] || ''
  ]);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), '操作执行记录');
  downloadWorkbook(wb, '操作执行记录.xlsx');
}

// 最终生成：一次性导出最终加班汇总、累计调班数据、操作执行记录，并标记最终生成完成
function exportFinalAll() {
  exportSystemData('final');
  setTimeout(() => exportShiftData('final'), 600);
  setTimeout(() => exportOperationLog(), 1200);

  appState.finalGenerated = true;
  renderNav();
  updateStats();
  showToast('最终生成完成，已导出全部最终文件', 'success');
}

// 导出导入与合并步骤的校验失败记录
function exportGroupFailures() {
  const failures = appState.groupFailures || [];
  if (!failures.length) {
    showToast('暂无校验失败记录', 'info');
    return;
  }
  const headers = ['sheet', '行号', '工号', '姓名', '班组', '失败原因'];
  const rows = failures.map(f => [f['sheet'], f['行号'], f['工号'], f['姓名'], f['班组'], f['失败原因']]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), '校验失败记录');
  downloadWorkbook(wb, '班组填报校验失败记录.xlsx');
}

// 导出重复填报清单（同工号 + 同一天 + 同一开始时间出现多次）
function exportGroupDuplicates() {
  const duplicates = appState.groupDuplicates || [];
  if (!duplicates.length) {
    showToast('暂无重复填报记录', 'info');
    return;
  }
  const headers = ['工号', '姓名', '加班开始日期', '加班开始时间', '条数', '系统序号', '班组'];
  const rows = duplicates.map(d => headers.map(h => d[h] !== undefined ? d[h] : ''));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), '重复填报记录');
  downloadWorkbook(wb, '班组填报重复记录.xlsx');
}

// 导出异常处理步骤的匹配失败记录
function exportAbnormalFailures() {
  const round = getCurrentRound();
  const failures = round.abnormalFailures || appState.abnormalFailures || [];
  if (!failures.length) {
    showToast('暂无匹配失败记录', 'info');
    return;
  }
  const headers = [...ABNORMAL_HEADERS, '行号', '线索', '失败原因'];
  const rows = failures.map(f => headers.map(h => f[h] !== undefined ? f[h] : ''));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), '匹配失败记录');
  downloadWorkbook(wb, '异常匹配失败记录.xlsx');
}

// 导出异常处理步骤的定位提醒（多条命中 / 姓名不一致）
function exportAbnormalWarnings() {
  const round = getCurrentRound();
  const warnings = round.abnormalWarnings || [];
  if (!warnings.length) {
    showToast('暂无定位提醒', 'info');
    return;
  }
  const headers = [...ABNORMAL_HEADERS, '行号', '定位提醒'];
  const rows = warnings.map(w => headers.map(h => w[h] !== undefined ? w[h] : ''));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), '定位提醒');
  downloadWorkbook(wb, '异常定位提醒.xlsx');
}

// 导出批量操作的定位异常清单（未定位 / 多条命中）
function exportLocateIssues() {
  const round = getCurrentRound();
  const issues = round.locateIssues || [];
  if (!issues.length) {
    showToast('暂无定位异常记录', 'info');
    return;
  }
  const headers = ['级别', '操作类型', '系统序号（执行前）', '校对ID', '工号', '姓名', '班组', '原开始日期', '原开始时间', '命中数', '候选序号', '说明'];
  const rows = issues.map(i => headers.map(h => i[h] !== undefined ? i[h] : ''));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([headers, ...rows]), '定位异常清单');
  downloadWorkbook(wb, '批量操作定位异常清单.xlsx');
}

// ==================== 文件上传绑定 ====================

function bindHiddenFileInput() {
  const input = document.getElementById('hiddenFileInput');
  input.onchange = (e) => {
    const file = e.target.files[0];
    const targetId = input.dataset.targetZone;
    if (file && targetId) {
      handleFile(file, targetId);
    }
    input.value = '';
    input.dataset.targetZone = '';
  };
}

function bindDropZones() {
  const zones = document.querySelectorAll('.drop-zone');
  zones.forEach(zone => {
    const fileInput = zone.querySelector('.file-input');
    const zoneId = zone.id;

    zone.addEventListener('dragover', (e) => {
      e.preventDefault();
      zone.classList.add('dragover');
    });
    zone.addEventListener('dragleave', () => {
      zone.classList.remove('dragover');
    });
    zone.addEventListener('drop', (e) => {
      e.preventDefault();
      zone.classList.remove('dragover');
      const files = e.dataTransfer.files;
      if (files.length > 0) {
        handleFile(files[0], zoneId);
      }
    });

    zone.addEventListener('click', () => {
      const hidden = document.getElementById('hiddenFileInput');
      hidden.dataset.targetZone = zoneId;
      hidden.click();
    });

    if (fileInput) {
      fileInput.addEventListener('click', (e) => {
        e.stopPropagation();
      });
      fileInput.addEventListener('change', (e) => {
        e.stopPropagation();
        if (e.target.files.length > 0) {
          handleFile(e.target.files[0], zoneId);
        }
      });
    }
  });
}

// M6-1：按区校验“这张表是不是这个区的表”。投错区时直接拒绝 —— 旧行为是取不到定位列后退化成「只按工号匹配」，
// 结果全部显示「已匹配/多条命中」、失败 0 条，用户以为导入成功了
const ZONE_TABLE_REQUIREMENTS = {
  dropZone: { name: '班组填报表', headers: ['工号', '姓名', '加班开始日期', '加班开始时间'] },
  dropZoneAbnormal: { name: '异常表', headers: ['工号', '开始日期', '开始时间', '上报加班时数'] },
  dropZoneRectify: { name: '整改表', headers: ['工号', '开始日期', '处置方式'] },
};

function tableProblem(zoneId, parsed) {
  const req = ZONE_TABLE_REQUIREMENTS[zoneId];
  if (!req) return '';
  const first = (parsed.sheetNames || [])[0];
  const rows = (parsed.sheets && parsed.sheets[first]) || [];
  const headers = (rows[0] || []).map(h => String(h === undefined || h === null ? '' : h).trim());
  const missing = req.headers.filter(h => !headers.includes(h));
  return missing.length ? `这看起来不是${req.name}（缺少列：${missing.join('、')}）。请确认拖对了文件` : '';
}

async function handleFile(file, zoneId) {
  const zone = document.getElementById(zoneId);
  const content = zone.querySelector('.zone-content');

  content.innerHTML = `
    <div class="w-14 h-14 rounded-2xl bg-apple-blue/10 text-apple-blue flex items-center justify-center mx-auto mb-4 animate-pulse">
      <i class="ph ph-spinner animate-spin text-2xl"></i>
    </div>
    <div class="text-sm font-medium text-apple-text">正在解析 ${file.name}</div>
    <div class="text-xs text-apple-muted mt-1">请稍候...</div>
  `;

  try {
    const parsed = await parseExcel(file);
    const problem = tableProblem(zoneId, parsed);
    if (problem) {
      content.innerHTML = `
        <div class="w-14 h-14 rounded-2xl bg-apple-red/10 text-apple-red flex items-center justify-center mx-auto mb-4">
          <i class="ph ph-warning-octagon text-2xl"></i>
        </div>
        <div class="text-sm font-medium text-apple-red">文件不对</div>
        <div class="text-xs text-apple-muted mt-1">${escapeHtml(problem)}</div>
      `;
      showToast(problem, 'error');
      return;
    }
    appState.fileName = file.name;

    if (zoneId === 'dropZone') {
      processGroupWorkbook(parsed);
    } else if (zoneId === 'dropZoneAbnormal') {
      processAbnormalWorkbook(parsed);
    } else if (zoneId === 'dropZoneRectify') {
      processRectifyWorkbook(parsed);
    }

    renderContent();
    showToast(`成功解析 ${file.name}`, 'success');
  } catch (err) {
    console.error(err);
    content.innerHTML = `
      <div class="w-14 h-14 rounded-2xl bg-apple-red/10 text-apple-red flex items-center justify-center mx-auto mb-4">
        <i class="ph ph-x-circle text-2xl"></i>
      </div>
      <div class="text-sm font-medium text-apple-red">解析失败</div>
      <div class="text-xs text-apple-muted mt-1">${escapeHtml(err.message || '请检查文件格式')}</div>
    `;
    showToast('文件解析失败', 'error');
  }
}

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  const colors = {
    success: 'bg-apple-green text-white',
    error: 'bg-apple-red text-white',
    warning: 'bg-apple-orange text-white',
    info: 'bg-apple-blue text-white',
  };
  toast.className = `fixed bottom-8 right-8 px-6 py-3 rounded-2xl shadow-lg text-sm font-medium z-50 ${colors[type] || colors.info} animate-[fadeIn_0.3s_ease-out]`;
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(8px)';
    toast.style.transition = 'all 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

function updateStats() {
  const total = appState.mergedRecords.length || 0;
  const allOps = getAllOperations();
  const cumulativeDelete = allOps.filter(o => o['操作类型'] === '删除').length;
  const cumulativeShift = getAllShiftRecords().length;
  const wf = getWorkflowStatus();

  const statTotal = document.getElementById('statTotal');
  const statCumulativeDelete = document.getElementById('statCumulativeDelete');
  const statCumulativeShift = document.getElementById('statCumulativeShift');
  const statWorkingStatus = document.getElementById('statWorkingStatus');

  if (statTotal) statTotal.textContent = total;
  if (statCumulativeDelete) statCumulativeDelete.textContent = cumulativeDelete;
  if (statCumulativeShift) statCumulativeShift.textContent = cumulativeShift;
  if (statWorkingStatus) statWorkingStatus.textContent = wf.workingStatus;
}

// 初始化
document.addEventListener('DOMContentLoaded', init);
