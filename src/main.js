import * as XLSX from 'xlsx-js-style';
import initialRules from './rules.json';
import './styles.css';

const app = document.querySelector('#root');
const levelLabels = { 1: '低', 2: '中', 3: '高', 4: '嚴重' };
const REFERENCE_BASE_COLUMNS = ['序號', '性別', '年齡', '糖尿病', '吸菸', '檳榔', '身高', '體重', 'BMI', '腰圍', '收縮壓', '舒張壓', '矯正右', '矯正左', '辨色力', '白血球', '血紅素', 'SGPT', '肌酸酐', '總膽固醇', '三酸甘油脂', '高密度膽固醇', '低密度膽固醇', '飯前血糖', '糖化血色素', '尿酸', '尿蛋白', '尿潛血', '胸部X光', '640切心臟'];
const REFERENCE_GRADE_COLUMNS = ['既往病史', '家族史', ...REFERENCE_BASE_COLUMNS.slice(8)];
const REFERENCE_COLUMNS = [...REFERENCE_BASE_COLUMNS, ...REFERENCE_GRADE_COLUMNS.map((name) => `(分級)${name}`), '第4級', '第3級', '第2級', '第1級', '最大級別', '10年內腦心血管風險'];
const EXPORT_BASE_COLUMNS = ['序號', '性別', '年齡', '糖尿病', '吸菸', '高血壓用藥', '身高', '體重'];
const RESULT_ITEMS = [
  ['BMI', '身體質量指數'], ['腰圍', '腰圍'], ['收縮壓', '收縮壓'], ['舒張壓', '舒張壓'], ['矯正(右眼)', '矯正(右眼)'], ['矯正(左眼)', '矯正(左眼)'], ['辨色力', '辨色力'], ['白血球', '白血球'], ['血紅素', '血紅素'], ['丙胺酸轉胺酶', '丙胺酸轉胺酶'], ['肌酸酐', '肌酸酐'], ['總膽固醇', '總膽固醇'], ['三酸甘油脂', '三酸甘油脂'], ['高密度脂蛋白', '高密度脂蛋白'], ['低密度脂蛋白', '低密度脂蛋白'], ['飯前血糖', '空腹血糖'], ['糖化血色素', '糖化血色素'], ['尿蛋白', '尿蛋白'], ['尿潛血', '尿潛血'], ['胸部X光', '胸部X光']
];
const RESULT_COLUMNS = [...EXPORT_BASE_COLUMNS, ...RESULT_ITEMS.flatMap(([label]) => [label, '分級']), '最大級別', '10年內腦心血管風險'];
const IMPORT_FIELD_ALIASES = { BMI: ['BMI', '身體質量指數'], 矯正左: ['矯正左', '視力(左)', '矯正(左眼)'], 矯正右: ['矯正右', '視力(右)', '矯正(右眼)'], 血紅素: ['血紅素', '血色素'], SGPT: ['SGPT', '丙胺酸轉胺酶'], 高密度膽固醇: ['高密度膽固醇', '高密度脂蛋白'], 高血壓用藥: ['高血壓用藥', '服用高血壓藥物', '使用高血壓藥物'], 胸部X光: ['胸部X光', 'X光', 'X-ray', 'Xray'] };
const IMPORT_TEMPLATE_COLUMNS = ['序號(僅供填寫數字、英文)', '性別(男、女)', '年齡(數字)', '糖尿病(有、無)', '吸菸(有、無)', '高血壓用藥(有、無)', '身高', '體重', '身體質量指數', '腰圍', '收縮壓', '舒張壓', '矯正(右眼)', '矯正(左眼)', '辨色力', '白血球', '血紅素', '丙胺酸轉胺酶', '肌酸酐', '總膽固醇', '三酸甘油脂', '高密度脂蛋白', '低密度脂蛋白', '空腹血糖', '糖化血色素', '尿蛋白', '尿潛血', '胸部X光'];
const REQUIRED_IMPORT_COLUMNS = ['序號', '性別', '年齡', '糖尿病', '吸菸', '身高', '體重', '高血壓用藥'];
const NON_GRADE_IMPORT_COLUMNS = ['身高', '體重', '使用高血壓藥物', '服用高血壓藥物', '高血壓用藥', '降壓藥', '血壓用藥', '640切心臟'];
const RULES_STORAGE_KEY = 'healthgrade.rules.v3';
const RULES_DEFAULT_VERSION = '2026-11-03-user-backup-v1';
const DEFAULT_RULE_ALIASES = {"身體質量指數": ["BMI", "body mass index"], "腰圍": ["waist circumference", "waist"], "收縮壓": ["SBP", "systolic blood pressure"], "舒張壓": ["DBP", "diastolic blood pressure"], "矯正(右眼)": ["矯正右", "視力(右)", "視力(右眼)", "corrected right vision"], "矯正(左眼)": ["矯正左", "視力(左)", "視力(左眼)", "corrected left vision"], "辨色力": ["色覺", "color vision", "color blindness"], "白血球": ["WBC", "white blood cell"], "血紅素": ["Hb", "hemoglobin", "血色素"], "丙胺酸轉胺酶": ["SGPT", "ALT", "GPT", "alanine aminotransferase", "肝功能"], "肌酸酐": ["Cr", "creatinine"], "總膽固醇": ["TC", "total cholesterol"], "三酸甘油脂": ["TG", "triglyceride", "三酸甘油酯"], "高密度脂蛋白": ["HDL", "high-density lipoprotein", "高密度膽固醇"], "低密度脂蛋白": ["LDL", "low-density lipoprotein", "低密度膽固醇"], "空腹血糖": ["FBS", "fasting blood sugar", "飯前血糖"], "糖化血色素": ["HbA1c", "A1C", "glycated hemoglobin"], "尿蛋白": ["pro", "Urine protein", "U/pro", "Protein"], "尿潛血": ["U/OB", "Occult blood", "OB"], "胸部X光": ["X-ray", "Xray", "X光"], "尿酸": ["UA", "uric acid"], "既往病史": ["過去疾病", "過去病史", "過往疾病", "既(過)往病史"], "促甲狀腺素": ["TSH", "TSH甲狀腺", "甲狀腺TSH", "甲狀腺"], "家族史": ["家庭病史", "家族病史"], "主訴": ["主要訴求"]};
const HIDDEN_RULES_COPY = { guidance: '每個項目名稱集中管理分級條件與可替代欄名；點選「分級規則」即可直接修改並儲存。主清單可刪除整個項目。文字條件不可完全相同，但「結節」、「大結節」、「小結節」等不同字詞可以並存。', source: '還原依據：職健導航_分級規則備份.xlsx', format: '數字：數值~數值；文字：完全相同不可重複' };
const BRAND_LOGO_URL = '/zhijian-logo.png';
const state = {
  consentAccepted: false,
  view: 'overview',
  rules: initialRules.map((rule, index) => ({ ...rule, order: index })),
  ruleAliases: Object.fromEntries(Object.entries(DEFAULT_RULE_ALIASES).map(([name, aliases]) => [name, [...aliases]])),
  records: [],
  results: [],
  sourceRows: [],
  sourceHeaders: [],
  reportMode: 'long',
  uploadStatus: 'empty',
  fileName: '',
  uploadedAt: '',
  riskResults: [],
  ruleGroupName: '',
  ruleGroupOriginalName: '',
  ruleAliasDraft: '',
  ruleGroupEdit: false,
  ruleGroupEditing: false,
  ruleEditDraft: [],
  ruleAddOpen: false,
  notice: '',
  ruleDraft: { name: '', aliases: '', type: 'number', levelCount: 4, conditions: [{ condition: '', level: 1 }, { condition: '', level: 2 }, { condition: '', level: 3 }, { condition: '', level: 4 }] },
  risk: { age: '', sex: '女性', smoking: false, diabetes: false, bpMedication: false, bpMedicationProvided: true, systolic: '', totalCholesterol: '', hdl: '', status: '待輸入資料', result: null }
};

const icon = {
  shield: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 20 6v5c0 5.1-3.4 8.7-8 10-4.6-1.3-8-4.9-8-10V6l8-3Z"/><path d="m8.5 12 2.2 2.2 4.8-5"/></svg>',
  upload: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5M5 14v5h14v-5"/></svg>',
  sliders: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h16"/><circle cx="9" cy="6" r="2"/><circle cx="15" cy="12" r="2"/><circle cx="11" cy="18" r="2"/></svg>',
  chart: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 19V5m0 14h16"/><path d="m7 15 3-3 3 2 5-6"/></svg>',
  download: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 4v11m0 0 4-4m-4 4-4-4M5 20h14"/></svg>',
  trash: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 7h14M10 11v5m4-5v5M8 7l1 13h6l1-13M9 7l1-3h4l1 3"/></svg>',
  spark: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m12 3 1.6 5.4L19 10l-5.4 1.6L12 17l-1.6-5.4L5 10l5.4-1.6L12 3ZM19 16l.7 2.3L22 19l-2.3.7L19 22l-.7-2.3L16 19l2.3-.7L19 16Z"/>',
  warning: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 22 20H2L12 3Z"/><path d="M12 9v5M12 17h.01"/></svg>',
  file: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 3h8l4 4v14H6z"/><path d="M14 3v5h5M9 13h6M9 17h6"/></svg>'
};

function cleanAliasList(values, canonical = '') {
  return [...new Set(String(values || '').split(/[,，\n]/).map((value) => value.trim()).filter((value) => value && value !== canonical))];
}

function saveRuleConfiguration() {
  // 自訂規則只在目前開啟的工作階段有效；跨次使用請以 Excel 備份匯入。
}

function loadRuleConfiguration() {
  try {
    const saved = JSON.parse(localStorage.getItem(RULES_STORAGE_KEY) || 'null');
    if (!saved?.rules?.length) return;
    if (saved.version !== RULES_DEFAULT_VERSION) {
      state.rules = initialRules.map((rule, index) => ({ ...rule, order: index }));
      state.ruleAliases = Object.fromEntries(Object.entries(DEFAULT_RULE_ALIASES).map(([name, aliases]) => [name, [...aliases]]));
      saveRuleConfiguration();
      return;
    }
    state.rules = saved.rules.map((rule, index) => ({ ...rule, order: Number.isFinite(rule.order) ? rule.order : index }));
    state.ruleAliases = Object.fromEntries([...new Set([...Object.keys(DEFAULT_RULE_ALIASES), ...Object.keys(saved.aliases || {})])].map((name) => [name, [...new Set([...(DEFAULT_RULE_ALIASES[name] || []), ...(saved.aliases?.[name] || [])].filter((alias) => alias && alias !== name))]]));
  } catch (error) { /* ignore invalid or unavailable local configuration */ }
}

function rulesFromBackupRows(rows) {
  const header = (rows[0] || []).map((value) => String(value ?? '').trim());
  const indexes = { name: header.indexOf('項目名稱'), aliases: header.indexOf('可替代名稱'), condition: header.indexOf('數值/文字'), level: header.indexOf('分級'), sex: header.indexOf('性別') };
  if ([indexes.name, indexes.aliases, indexes.condition, indexes.level].some((index) => index < 0)) throw new Error('備份檔第一排必須包含：項目名稱、可替代名稱、數值/文字、分級。');
  const rules = []; const aliases = {}; let currentName = '';
  rows.slice(1).forEach((row) => {
    const nameCell = String(row[indexes.name] ?? '').trim(); const aliasCell = String(row[indexes.aliases] ?? '').trim(); const condition = String(row[indexes.condition] ?? '').trim(); const level = Number(row[indexes.level]);
    if (nameCell) currentName = nameCell;
    if (!currentName || !condition || !Number.isFinite(level) || level < 1 || level > 4) return;
    if (aliasCell) aliases[currentName] = cleanAliasList(aliasCell, currentName);
    const normalizedCondition = condition.replace(/[–—〜－]/g, '~');
    const sexText = indexes.sex >= 0 ? String(row[indexes.sex] ?? '').trim() : '';
    const sex = sexText === '男性' ? 'male' : sexText === '女性' ? 'female' : 'all';
    rules.push({ id: `rule-restored-${Date.now()}-${rules.length}`, name: currentName, condition: normalizedCondition, level, sex, enabled: true, type: parseRange(normalizedCondition) ? 'number' : 'text', order: rules.length });
  });
  if (!rules.length) throw new Error('備份檔沒有可用的分級規則。');
  return { rules, aliases };
}

async function downloadRuleBackup() {
  const rows = [['項目名稱', '可替代名稱', '數值/文字', '分級', '性別']];
  [...new Set(state.rules.map((rule) => rule.name.trim()))].forEach((name) => {
    const group = state.rules.filter((rule) => rule.name.trim() === name); const aliasText = (state.ruleAliases[name] || []).join('、');
    group.forEach((rule, index) => rows.push([index === 0 ? name : '', index === 0 ? aliasText : '', displayCondition(rule.condition), rule.level, ruleSexLabel(rule.sex)]));
  });
  const sheet = XLSX.utils.aoa_to_sheet(rows); sheet['!cols'] = [{ wch: 20 }, { wch: 42 }, { wch: 24 }, { wch: 10 }, { wch: 12 }];
  const headerStyle = { fill: { fgColor: { rgb: '0F776E' } }, font: { color: { rgb: 'FFFFFF' }, bold: true }, alignment: { horizontal: 'center' } };
  rows[0].forEach((_, columnIndex) => { const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: columnIndex })]; if (cell) cell.s = headerStyle; });
  const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, sheet, '分級規則'); await saveWorkbookWithPicker(workbook, '職健導航_分級規則備份.xlsx');
}

async function importRuleBackup(file) {
  if (!file) return;
  try {
    const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: '', blankrows: false });
    const payload = rulesFromBackupRows(rows);
    state.rules = payload.rules; state.ruleAliases = Object.fromEntries([...new Set([...Object.keys(DEFAULT_RULE_ALIASES), ...Object.keys(payload.aliases || {})])].map((name) => [name, [...new Set([...(DEFAULT_RULE_ALIASES[name] || []), ...(payload.aliases?.[name] || [])].filter((alias) => alias && alias !== name))]]));
    saveRuleConfiguration(); render(); window.alert('分級規則備份已匯入並保存於此瀏覽器。');
  } catch (error) { window.alert(`規則備份匯入失敗：${error.message || '請選擇職健導航匯出的 Excel 備份檔。'}`); }
}

try { localStorage.removeItem(RULES_STORAGE_KEY); } catch (error) { /* localStorage may be blocked by private browsing */ }

function esc(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' }[char]));
}

function ruleType(rule) {
  return rule.type || (parseRange(rule.condition) ? 'number' : 'text');
}

function displayCondition(condition) {
  const single = String(condition ?? '').replace(/[．。]/g, '.').trim().match(/^(<=|>=|<|>)\s*(-?\d+(?:\.\d+)?)$/);
  if (single) return `${single[1]}${single[2]}`;
  const range = parseRange(condition);
  return range ? `${range.min}~${range.max}` : String(condition ?? '');
}

function successCount() { return state.results.filter((item) => item.matched).length; }
function failureCount() { return state.results.filter((item) => !item.matched).length; }
function showNotice(message) { state.notice = message; render(); }
function uploadStatusText() { return ({ empty: '尚無檔案', uploading: '上傳中', success: '上傳成功', failed: '上傳失敗' })[state.uploadStatus] || '尚無檔案'; }
function resultStatusText() { return state.uploadStatus === 'uploading' ? '分級中' : (state.uploadStatus === 'success' ? '完成分級' : '尚無檔案'); }
function statusClass(status) { return `nav-state-${status}`; }
function currentViewLabel() { return ({ overview: '總表', import: '匯入健檢資料', results: '分級結果', rules: '分級規則' })[state.view] || '總表'; }
function currentClock() { return new Intl.DateTimeFormat('zh-TW', { hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false }).format(new Date()); }
function renderLandingPage() {
  return `<main class="landing-page"><section class="landing-card"><div class="landing-brand"><div class="landing-logo"><img src="${BRAND_LOGO_URL}" alt="職健導航 Logo" /></div><div><div class="landing-kicker">OCCUPATIONAL HEALTH NAVIGATION TOOL</div><h1>職健導航</h1><p>本機健檢分級工具</p></div></div><div class="landing-hero"><h2>開始前，請先閱讀使用聲明</h2><p>本工具協助職護人員在本機完成健檢資料分級、統計與結果匯出。</p></div><div class="landing-sections"><article><h3>個資與資料處理聲明</h3><ul><li>您匯入的健檢 Excel 僅在目前瀏覽器記憶體中讀取、分析與匯出。</li><li>本工具不主動上傳健檢資料、不建立健檢資料庫，也不要求姓名、身分證字號或工號等直接識別資訊。</li><li>請先完成個資去識別化，再上傳檔案至平台進行分級。</li><li>重新整理或關閉頁面後，本次健檢資料不會由本工具永久保存；自訂規則請另行下載 Excel 備份。</li></ul></article><article><h3>使用規則與重要提醒</h3><ul><li>匯入資料前請確認您已取得組織授權，並遵守個人資料保護及相關醫療、職業安全衛生規範。</li><li>分級規則與腦心血管風險結果僅供篩檢與工作輔助，不能取代醫師診斷、正式病歷或醫療建議。</li><li>所有分級、風險模型與匯出結果，均須由合格職護／職醫人工覆核。</li><li>使用者須自行確認資料正確性、規則適用性與匯入內容是否符合工作需求。</li></ul></article></div><div class="landing-consent"><label><input type="checkbox" id="landingConsent" /><span>我已閱讀並同意上述個資處理聲明與使用規則</span></label><button class="btn btn-primary" id="enterTool" disabled>同意並進入職健導航</button></div>${renderAdSlot('landing')}<p class="landing-footnote">本工具不提供醫療診斷；使用前請確認已完成必要的個資告知與組織授權。</p></section></main>`;
}
function render() {
  if (!state.consentAccepted) {
    app.innerHTML = renderLandingPage();
    const consent = document.querySelector('#landingConsent');
    const enter = document.querySelector('#enterTool');
    consent?.addEventListener('change', () => { enter.disabled = !consent.checked; });
    enter?.addEventListener('click', () => { if (!consent.checked) return; state.consentAccepted = true; render(); });
    return;
  }
  const conflictCount = detectConflicts(state.rules).length;
  const levelCount = [1, 2, 3, 4].map((level) => state.results.filter((item) => item.level === level).length);
  app.innerHTML = `
    <div class="shell">
      <aside class="sidebar">
        <div class="brand"><div class="brand-mark"><img src="${BRAND_LOGO_URL}" alt="職健導航 Logo" /></div><div><div class="brand-name">職健導航</div><div class="brand-caption">本機健檢分級工具</div></div></div>
        <div class="privacy-chip"><span class="status-dot"></span><span>本機模式已啟用</span></div>
        <nav class="nav"><a class="nav-item nav-overview ${state.view === 'overview' ? 'active' : ''}" data-view="overview" href="#overview">${icon.chart}<span class="nav-label">總表</span></a><div class="nav-divider"><span>工作流程</span></div><a class="nav-item nav-step-item ${state.view === 'rules' ? 'active' : ''}" data-view="rules" href="#rules">${icon.sliders}<span class="nav-step-no">Step 1.</span><span class="nav-label">分級規則</span></a><a class="nav-item nav-step-item ${state.view === 'import' ? 'active' : ''}" data-view="import" href="#import">${icon.upload}<span class="nav-step-no">Step 2.</span><span class="nav-copy"><span class="nav-label">匯入健檢報告</span><span class="nav-state ${statusClass(state.uploadStatus)}">${uploadStatusText()}</span></span></a><a class="nav-item nav-step-item ${state.view === 'results' ? 'active' : ''}" data-view="results" href="#results">${icon.chart}<span class="nav-step-no">Step 3.</span><span class="nav-copy"><span class="nav-label">分級結果</span><span class="nav-state ${statusClass(state.uploadStatus === 'uploading' ? 'uploading' : (state.uploadStatus === 'success' ? 'success' : 'empty'))}">${resultStatusText()}</span></span></a></nav>
        <div class="sidebar-note"><div class="note-icon">${icon.shield}</div><strong>資料不離開您的裝置</strong><p>報告只在瀏覽器記憶體中處理，不上傳、不建檔。</p></div>
        ${renderAdSlot('sidebar')}
        <div class="sidebar-footer">v0.3 本機原型<br /><span>規則由使用者自行確認</span><button class="privacy-link" id="privacyPolicy">隱私權與使用說明</button></div>
      </aside>
      <main class="main">
        <header class="topbar"><div class="topbar-breadcrumb"><span>職健導航</span></div><div class="topbar-main"><div><div class="eyebrow">OCCUPATIONAL HEALTH MANAGEMENT</div><h1>${currentViewLabel()}</h1><p class="subtitle">把報告轉成可追蹤的分級結果，保留規則調整權，資料全程留在本機。</p></div><div class="top-actions"><span class="topbar-date">2026 年 10 月 15 日　週四　<span id="liveClock">${currentClock()}</span></span><span class="secure-label">${icon.shield} 不連線處理</span></div></div></header>
        ${renderView(conflictCount, levelCount)}
        <footer class="page-footer"><span>${icon.shield} 本工具不提供醫療診斷，分級結果須由合格職護／職醫覆核。</span><button class="privacy-footer-link" id="privacyPolicyFooter">隱私權與使用說明</button><span>健檢資料僅存在目前瀏覽器；自訂規則只在本次開啟期間生效</span></footer>
      </main>
      ${renderPrivacyModal()}${renderNoticeModal()}
    </div>`;
  bindEvents();
}

function renderView(conflictCount, levelCount) {
  if (state.view === 'overview') return renderDashboard(levelCount);
  if (state.view === 'import') return renderImportPanel();
  if (state.view === 'rules') return `${renderRulesPanel(conflictCount)}${renderRiskPanel()}`;
  return renderResultsPanel(levelCount);
}

function importedRowsWithSerial() { return state.sourceRows.filter((row) => String(row['序號'] ?? '').trim() !== ''); }
function serialKey(row) { return String(row['序號'] ?? '').trim().replace(/\s+/g, '').toUpperCase(); }
function importedSerialCount() { return new Set(importedRowsWithSerial().map(serialKey)).size; }

function renderDashboard(levelCount) {
  const importedCount = importedSerialCount();
  return `<section class="dashboard-section" id="overview"><div class="dashboard-heading"><div><div class="section-kicker">SUMMARY / DASHBOARD</div><p>快速掌握本次匯入資料的主要異常項目。</p></div><div class="dashboard-file">${state.fileName ? `${icon.file} ${esc(state.fileName)}` : '尚未匯入報告'}</div></div><div class="hero-stats"><div class="stat-card accent"><div class="stat-label">目前匯入</div><div class="stat-value">${importedCount || '—'}</div><div class="stat-foot">${importedCount ? '筆檢查資料（依序號統計）' : '尚未匯入 Excel'}</div></div></div>${renderAnomalyChart(levelCount)}${renderAdSlot('dashboard')}</section>`;
}

function renderPrivacyModal() {
  return `<div class="privacy-modal" id="privacyModal" hidden><div class="privacy-backdrop" data-close-privacy></div><section class="privacy-dialog" role="dialog" aria-modal="true" aria-labelledby="privacyTitle"><button class="privacy-close" data-close-privacy aria-label="關閉隱私權說明">×</button><div class="section-kicker">PRIVACY & USE</div><h2 id="privacyTitle">隱私權與使用說明</h2><p class="privacy-updated">最後更新：2026 年 11 月 5 日</p><div class="privacy-content"><h3>一、適用範圍與服務定位</h3><p>「職健導航」是提供職護人員使用的純前端健檢資料分級、統計與 Excel 匯出輔助工具。本工具不是醫療機構、電子病歷系統、正式健檢資料庫或醫療診斷服務。使用本工具即表示您已閱讀本說明，並同意依組織規範與適用法令使用。</p><h3>二、健檢資料如何處理</h3><p>您選取的 Excel 會由目前瀏覽器在本機記憶體中讀取、分析、分級與匯出。本工具不主動將 Excel 內容、姓名、身分證字號、工號、健檢數值或分級結果上傳至本工具的應用程式伺服器，也不建立健檢資料庫。請先完成個資去識別化，再上傳檔案至平台進行分級。</p><h3>三、資料保存與刪除</h3><p>本次健檢資料只在目前瀏覽器工作階段中使用；重新上傳檔案會取代目前分析結果，關閉或重新整理頁面後，本工具不提供本次資料的永久保存或復原功能。請使用者自行管理下載到電腦的 Excel 檔案、瀏覽器下載紀錄、暫存檔、備份檔與螢幕截圖。分級規則修改、新增或刪除後不會自動保存；需要保留時，請自行下載 Excel 備份。</p><h3>四、網站託管與技術性紀錄</h3><p>雖然健檢 Excel 由瀏覽器本機處理，公開網站的託管、CDN、DNS、安全防護或錯誤監控服務仍可能依其政策產生一般技術性紀錄，例如 IP 位址、連線時間、瀏覽器類型、裝置資訊、請求網址與安全事件紀錄。這些網站連線紀錄不等同於本工具主動收集健檢內容；實際保存期間與處理方式仍以相關託管服務商政策為準。</p><h3>五、Cookie、廣告與第三方服務</h3><p>目前頁面僅保留非干擾性廣告版位，未在本工具程式中放入健檢資料追蹤機制。若日後啟用 Google AdSense、分析工具、Cookie 或其他第三方服務，該服務可能依其自身政策處理 IP、Cookie、裝置、瀏覽與廣告成效資料；啟用前應更新本說明、提供必要的告知與選擇機制，並依服務地區要求完成設定。第三方服務不應取得使用者匯入的 Excel 內容。</p><h3>六、使用者義務與誤傳個資</h3><p>使用者應在匯入前完成去識別化、確認已取得組織授權，並確認檔案不含不必要的姓名、身分證字號、工號、電話、地址、病歷號或其他可直接識別個人的資訊。若使用者自行匯入、下載、保存、分享或公開含有真實個資的檔案，使用者仍須依其身分、組織政策及適用法令負責處理；首頁告知並不等於對使用者違法或不當處理個資的豁免。若發生誤傳，請立即停止分享、刪除相關檔案與下載副本，並依所屬組織的個資事件通報流程處理。</p><h3>七、分級與醫療責任限制</h3><p>分級規則、公開門檻、腦心血管風險模型與統計結果可能因資料品質、欄位名稱、規則設定、版本或模型限制而產生差異。本工具僅供篩檢與工作流程輔助，不提供疾病診斷、治療建議或正式醫療判定；所有結果均須由合格職護、職醫或相關專業人員查核，不得直接作為醫療決策、雇用決策或其他重大人身決策的唯一依據。</p><h3>八、資訊安全與使用限制</h3><p>本工具採瀏覽器本機處理，但使用者仍須自行維護裝置、瀏覽器、下載資料夾、備份檔與帳號環境的安全。請勿在共用電腦留下未加密的健檢檔案，也請勿將去識別化不足的檔案寄送或分享給未授權人員。網站提供者不保證所有瀏覽器、檔案格式、網路環境或第三方託管服務永遠不中斷或完全無錯誤。</p><h3>九、規則與說明更新</h3><p>網站可能因功能、分級規則、公開健康門檻、風險模型、託管服務或法規需求而更新。更新後會調整版本內容與本頁最後更新日期；使用者在正式工作流程中採用結果前，應確認目前版本是否經組織核准與專業覆核。</p><h3>十、聯絡與法律檢視</h3><p>若您代表公司、醫療院所或健檢機構使用，請由組織指定窗口處理個資、資安與醫療責任問題。本頁是網站公開使用說明，不構成法律意見，也不取代個資盤點、委外契約、資安評估、事件通報計畫或正式隱私權政策；正式對外營運前，建議依服務地區、託管商、廣告設定與實際營運主體，請個資或法律專業人士檢視。</p><h3>十一、免費服務與使用條款</h3><p>本工具目前提供免費使用，網站可能透過廣告版位維持營運；免費使用不代表網站提供任何醫療保證、資料保存、特定可用性或不間斷服務承諾。使用者不得以本工具取代專業判斷，不得故意上傳、公開、分享未完成去識別化的資料，不得以干擾網站、繞過安全機制、濫用自動化請求或其他違法方式使用本服務。網站提供者得因維護、資安、法規、第三方服務或功能更新而暫停、調整或停止部分功能。使用者應自行確認匯入資料、規則設定、結果解讀與下載檔案的正確性，並承擔其組織授權、資料管理及工作決策責任。</p><p class="privacy-disclaimer">使用前請確認：已完成個資去識別化、具備必要授權、理解本工具僅供工作輔助，並會由合格職護／職醫人工覆核結果。</p></div></section></div>`;
}
function renderNoticeModal() {
  if (!state.notice) return '';
  return `<div class="notice-modal"><div class="notice-backdrop" data-close-notice></div><section class="notice-dialog" role="dialog" aria-modal="true"><div class="notice-icon">✓</div><h2>${esc(state.notice)}</h2><button class="btn btn-primary" id="closeNotice">確定</button></section></div>`;
}

function renderAdSlot(place) {
  return `<div class="ad-slot ad-slot-${place}" data-ad-placement="${place}" aria-label="廣告"><span class="ad-label">廣告</span><span class="ad-placeholder">保留清爽版位，不干擾健檢資料操作</span></div>`;
}

function metricValue(row, keys) { const raw = keys.find((key) => row[key] !== undefined && String(row[key]).trim() !== ''); return raw ? parseNumber(row[raw]) : null; }
function rowFlag(row, keys) { const raw = keys.find((key) => row[key] !== undefined && String(row[key]).trim() !== ''); return raw ? flagValue(row[raw]) : false; }
function rowSex(row) { const raw = String(row['性別'] ?? '').trim().toLowerCase(); return ['男', '男性', 'm', 'male'].includes(raw) ? '男性' : ['女', '女性', 'f', 'female'].includes(raw) ? '女性' : ''; }
function ruleSexLabel(sex) { return sex === 'male' ? '男性' : sex === 'female' ? '女性' : '不限性別'; }
function ruleSexMatches(rule, record) { return !rule.sex || rule.sex === 'all' || (rule.sex === 'male' && rowSex(record) === '男性') || (rule.sex === 'female' && rowSex(record) === '女性'); }
function ruleSexesOverlap(left, right) { return !left || left === 'all' || !right || right === 'all' || left === right; }

function healthConditionFlags(row) {
  const sex = rowSex(row);
  const bmi = metricValue(row, ['BMI', '身體質量指數', 'body mass index']); const waist = metricValue(row, ['腰圍']); const sbp = metricValue(row, ['收縮壓']); const dbp = metricValue(row, ['舒張壓']);
  const totalCholesterol = metricValue(row, ['總膽固醇', '總膽固醇(TC)']); const ldl = metricValue(row, ['低密度膽固醇', '低密度脂蛋白']); const hdl = metricValue(row, ['高密度膽固醇', '高密度脂蛋白', '高密度脂蛋白膽固醇']); const triglyceride = metricValue(row, ['三酸甘油脂']); const fastingGlucose = metricValue(row, ['飯前血糖', '空腹血糖', 'FBS', 'fasting blood sugar', 'glucose']); const hba1c = metricValue(row, ['糖化血色素', 'HbA1c', 'A1C', 'glycated hemoglobin']); const uricAcid = metricValue(row, ['尿酸']); const alt = metricValue(row, ['丙胺酸轉胺酶', 'SGPT', 'ALT']);
  const diabetes = rowFlag(row, ['糖尿病']); const diabetesMedication = rowFlag(row, ['使用糖尿病藥物', '糖尿病用藥', '降血糖藥', '降血糖藥物', '糖尿病藥物']); const bpMedication = rowFlag(row, ['使用高血壓藥物', '高血壓用藥', '降壓藥', '血壓用藥']); const lipidMedication = rowFlag(row, ['使用降血脂藥物', '降血脂藥', '降血脂藥物', '血脂用藥', '降膽固醇藥']);
  const bpElevated = sbp !== null && sbp >= 120 && sbp <= 129 && dbp !== null && dbp < 80;
  const bpStage1 = Boolean((sbp !== null && sbp >= 130 && sbp <= 139) || (dbp !== null && dbp >= 80 && dbp <= 89));
  const bpStage2 = Boolean((sbp !== null && sbp >= 140) || (dbp !== null && dbp >= 90));
  const highBloodPressure = Boolean((sbp !== null && sbp >= 140) || (dbp !== null && dbp >= 90) || bpMedication);
  const lowHdl = hdl !== null && ((sex === '男性' && hdl < 40) || (sex === '女性' && hdl < 50));
  const metabolicAbdominal = waist !== null && ((sex === '男性' && waist >= 90) || (sex === '女性' && waist >= 80));
  const metabolicBloodPressure = Boolean((sbp !== null && sbp >= 130) || (dbp !== null && dbp >= 85) || bpMedication);
  const metabolicGlucose = Boolean((fastingGlucose !== null && fastingGlucose >= 100) || diabetesMedication);
  const metabolicTriglyceride = Boolean((triglyceride !== null && triglyceride >= 150) || lipidMedication);
  const metabolicComponents = [metabolicAbdominal, metabolicBloodPressure, metabolicGlucose, metabolicTriglyceride, lowHdl].filter(Boolean).length;
  return {
    hypertension: Boolean(highBloodPressure || bpElevated || bpStage1),
    // 血壓異常總人數只取畫面列出的三個血壓數值細項聯集；
    // 高血壓用藥僅屬代謝症候群判定條件，不併入此項，避免總數被多算。
    bloodPressure: Boolean(bpElevated || bpStage1 || bpStage2),
    bpElevated, bpStage1, bpStage2,
    dyslipidaemia: Boolean((totalCholesterol !== null && totalCholesterol >= 200) || (ldl !== null && ldl >= 130) || (triglyceride !== null && triglyceride >= 150) || lowHdl),
    totalCholesterol: totalCholesterol !== null && totalCholesterol >= 200,
    ldl: ldl !== null && ldl >= 130,
    triglyceride: triglyceride !== null && triglyceride >= 150,
    lowHdl,
    obesity: bmi !== null && bmi >= 24,
    bmiOverweight: bmi !== null && bmi >= 24 && bmi < 27,
    bmiMildObesity: bmi !== null && bmi >= 27 && bmi < 30,
    bmiModerateObesity: bmi !== null && bmi >= 30 && bmi < 35,
    bmiSevereObesity: bmi !== null && bmi >= 35,
    metabolicSyndrome: metabolicComponents >= 3,
    metabolicAbdominal, metabolicBloodPressure, metabolicGlucose, metabolicTriglyceride, metabolicHdl: lowHdl,
    hyperuricaemia: uricAcid !== null && ((sex === '男性' && uricAcid > 7) || (sex === '女性' && uricAcid > 6)),
    hyperglycaemia: Boolean((fastingGlucose !== null && fastingGlucose >= 126) || (hba1c !== null && hba1c >= 6.5)),
    diabetesLevel: fastingGlucose !== null && fastingGlucose >= 126,
    diabetesA1c: hba1c !== null && hba1c >= 6.5,
    liver: alt !== null && alt > 40
  };
}

function renderAnomalyChart() {
  const rows = importedRowsWithSerial();
  const total = rows.length;
  const countBy = (key) => new Set(rows.filter((row) => healthConditionFlags(row)[key]).map(serialKey)).size;
  const countUnionBy = (keys) => new Set(rows.filter((row) => keys.some((key) => healthConditionFlags(row)[key])).map(serialKey)).size;
  const definitions = [
    ['bloodPressure', '血壓異常', '收縮壓／舒張壓依公開成人血壓分級門檻統計', [
      ['bpElevated', '血壓偏高（警戒期）：收縮壓 120–129 mmHg 且舒張壓 <80 mmHg'],
      ['bpStage1', '第一期高血壓：收縮壓 130–139 mmHg 或舒張壓 80–89 mmHg'],
      ['bpStage2', '第二期高血壓：收縮壓 ≥140 mmHg 或舒張壓 ≥90 mmHg']
    ]],
    ['liver', '肝功能異常', 'SGPT／ALT >40 U/L', [['liver', 'SGPT／ALT >40 U/L']]],
    ['hyperglycaemia', '血糖異常', '空腹血糖 ≥126 mg/dL 或糖化血色素 ≥6.5%', [
      ['diabetesLevel', '空腹血糖 ≥126 mg/dL'],
      ['diabetesA1c', '糖化血色素 ≥6.5%']
    ]],
    ['dyslipidaemia', '血脂異常', '總膽固醇、LDL、三酸甘油脂或 HDL 達異常門檻', [
      ['totalCholesterol', '總膽固醇 ≥200 mg/dL'],
      ['ldl', 'LDL ≥130 mg/dL'],
      ['triglyceride', '三酸甘油脂 ≥150 mg/dL'],
      ['lowHdl', 'HDL 偏低：男性 <40、女性 <50 mg/dL']
    ]],
    ['obesity', 'BMI 異常', '亞洲成人 BMI ≥24 視為異常提示', [
      ['bmiOverweight', '體重過重：24 ≤ BMI < 27'],
      ['bmiMildObesity', '輕度肥胖：27 ≤ BMI < 30'],
      ['bmiModerateObesity', '中度肥胖：30 ≤ BMI < 35'],
      ['bmiSevereObesity', '重度肥胖：BMI ≥35']
    ]],
    ['hyperuricaemia', '尿酸異常', '男性 >7、女性 >6 mg/dL', [['hyperuricaemia', '高尿酸：男性 >7、女性 >6 mg/dL']]],
    ['metabolicSyndrome', '代謝症候群風險', '五項指標中符合三項（含）以上', [
      ['metabolicAbdominal', '腹部肥胖：男性腰圍 ≥90 公分、女性腰圍 ≥80 公分'],
      ['metabolicBloodPressure', '血壓偏高：收縮壓 ≥130、舒張壓 ≥85 mmHg，或服用高血壓藥物'],
      ['metabolicGlucose', '空腹血糖偏高：空腹血糖 ≥100 mg/dL，或服用糖尿病藥物'],
      ['metabolicTriglyceride', '三酸甘油酯偏高：≥150 mg/dL，或服用降血脂藥物'],
      ['metabolicHdl', 'HDL-C 偏低：男性 <40、女性 <50 mg/dL']
    ]]
  ];
  const categories = definitions.map(([key, label, note, details]) => ({
    key, label, note, details,
    count: key === 'dyslipidaemia' ? countUnionBy(['totalCholesterol', 'ldl', 'triglyceride', 'lowHdl']) : countBy(key),
    detailCounts: details.map(([detailKey, detailLabel]) => ({ label: detailLabel, count: countBy(detailKey) }))
  })).sort((a, b) => b.count - a.count || a.label.localeCompare(b.label, 'zh-Hant'));
  const max = Math.max(...categories.map((item) => item.count), 1);
  const content = total ? categories.map((item) => {
    const percent = Math.round(item.count / total * 100);
    const width = item.count ? Math.max(6, item.count / max * 100) : 2;
    const detailMarkup = item.detailCounts.map((detail) => `<li><span>${esc(detail.label)}</span><b>${detail.count} 人</b></li>`).join('');
    return `<div class="anomaly-row" title="${esc(item.note)}"><div class="anomaly-label"><strong>${esc(item.label)}</strong><span>${item.count} 人／${percent}%</span></div><div class="anomaly-track"><i style="width:${width}%"></i></div><ul class="anomaly-subdetails">${detailMarkup}</ul></div>`;
  }).join('') : `<div class="chart-empty"><div class="empty-chart">${icon.chart}</div><strong>匯入報告後顯示疾病統計</strong><span>依本頁公開門檻計算異常人數</span></div>`;
  const footnotes = total ? `<div class="anomaly-footnotes"><span>統計母數：${total} 位不重複序號</span><span>柱狀圖依異常人數由高至低排列</span><span>圖表為篩檢提示，不等同醫療診斷</span></div>` : '';
  return `<div class="anomaly-panel"><div class="anomaly-heading"><div><h3>主要異常項目</h3><p>依匯入健檢數值與 WHO／國際公開健康門檻統計異常人數</p><small class="anomaly-disclaimer">本網站僅統計部分異常，如：代謝症候群、肝功能異常、血糖異常、BMI異常、血脂異常、尿酸異常……等。</small></div><span class="chart-legend"><i></i>異常人數／占比</span></div><div class="anomaly-chart">${content}</div>${footnotes}</div>`;
}

function renderImportPanel() {
  const fileClass = state.uploadStatus === 'success' ? 'ok' : state.uploadStatus === 'failed' ? 'error' : 'loading';
  const successBanner = state.uploadStatus === 'success' ? `<div class="upload-success-banner"><span class="upload-success-icon">✓</span><div><strong>上傳成功</strong><span>${esc(state.fileName)} 已完成本機讀取，請點選左側「分級結果」。</span></div></div>` : '';
  return `<section class="panel import-panel" id="import"><div class="section-heading"><div><div class="section-kicker">STEP 02 / IMPORT</div><p>上傳 Excel 後，請點選左側欄位「分級結果」。</p></div></div><label class="dropzone" for="excelInput"><div class="drop-icon">${icon.upload}</div><div class="drop-copy"><strong>拖曳 Excel 至此，或點擊選取檔案</strong><span>支援 .xlsx、.xls、.csv；資料只在本機瀏覽器讀取</span></div><span class="drop-browse-button">瀏覽檔案</span><input id="excelInput" type="file" accept=".xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv" /></label>${successBanner}<div class="upload-note">${icon.warning}<div>1. 匯入Excel檔案，上方第一排欄位項目名稱必須包含：「序號、性別、年齡、糖尿病、吸菸、身高、體重、服用高血壓藥物」，序號可使用英文、數字或英數混合。</div><div>2. 若您匯入多次失敗，請下載匯入範例參考，謝謝。</div></div><div class="import-lower"><div class="file-status">${state.fileName ? `<span class="file-${fileClass}">${uploadStatusText()}</span> ${esc(state.fileName)}` : '尚無檔案'}</div><button class="btn btn-secondary" id="downloadTemplate">${icon.download} 下載匯入範例</button></div></section>`;
}

function renderRulesPanel(conflictCount) {
  const names = [...new Set(state.rules.map((rule) => rule.name.trim()).filter(Boolean))];
  return `<section class="panel rules-panel" id="rules"><div class="section-heading"><div><div class="section-kicker">STEP 04 / RULES</div></div><div class="rule-summary"><span class="summary-number">${names.length}</span><span>個項目／${state.rules.length} 條規則</span></div></div><div class="rule-toolbar"><div class="search-wrap"><span>⌕</span><input id="ruleSearch" placeholder="搜尋項目或替代名稱，例如：Hb" /></div><button class="btn btn-secondary" id="addRule">＋ 新增規則</button><button class="btn btn-secondary" id="resetRules">還原網站規則</button><button class="btn btn-secondary" id="downloadRulesBackup">下載 Excel 備份</button><label class="btn btn-secondary backup-upload-label" for="rulesBackupInput">匯入 Excel 備份<input id="rulesBackupInput" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" hidden /></label></div><div class="rule-persistence-note">分級規則修改或新增或刪除後，皆不會自動保存，因此若關閉網站後，仍要保留您設定的分級規則，請先下載 Excel 備份檔案；下次開啟網站時，再匯入備份檔案即可。</div>${conflictCount ? `<div class="conflict-banner">${icon.warning}<div><strong>目前有 ${conflictCount} 組規則重複或數值重疊</strong><span>請按同名規則後直接修改；系統不會讓新的重複條件儲存。</span></div></div>` : ''}<div class="rule-table-wrap"><table class="rule-table rule-index-table"><thead><tr><th>項目名稱</th><th>檢視／修改</th><th>刪除</th></tr></thead><tbody>${names.map((name, index) => renderRuleRow(name, index)).join('')}</tbody></table></div><div class="table-footer"><span>顯示 ${names.length} 個分級項目，共 ${state.rules.length} 條規則</span></div></section>${renderRuleGroupModal()}${renderAddRuleModal()}`;
}

function renderRuleRow(name, index) {
  const sameName = state.rules.filter((rule) => rule.name.trim() === name.trim());
  const conflict = sameName.some((rule) => detectConflicts(state.rules).some((pair) => pair.includes(rule.id)));
  const aliases = state.ruleAliases[name] || [];
  return `<tr class="${conflict ? 'conflict-row' : ''}" data-rule-name="${esc(name)}"><td><strong class="rule-index-name">${esc(name)}</strong><small>${sameName.length} 條分級條件</small>${aliases.length ? `<em class="rule-alias-preview">可替代：${esc(aliases.join('、'))}</em>` : ''}</td><td><button class="btn btn-secondary group-rule" title="開啟並修改 ${esc(name)} 的分級規則">分級規則</button></td><td><button class="btn btn-danger delete-rule-group" title="刪除 ${esc(name)} 的全部規則">刪除</button></td></tr>`;
}

function renderRuleGroupModal() {
  if (!state.ruleGroupName) return '';
  const sourceRules = state.ruleEditDraft;
  const editing = state.ruleGroupEditing;
  const field = (value, className = '', placeholder = '') => editing ? `<input class="${className} edit-white-field" value="${esc(value)}" placeholder="${placeholder}" />` : `<span class="rule-readonly-value ${className}">${esc(value)}</span>`;
  const conditionRows = sourceRules.map((rule, index) => { const type = ruleType(rule); const sex = rule.sex || 'all'; if (!editing) return `<div class="rule-group-item rule-group-readonly"><span>${esc(displayCondition(rule.condition))}</span><span>${type === 'number' ? '數字' : '文字'}</span><span>${levelLabels[rule.level] ? `${rule.level} · ${levelLabels[rule.level]}` : rule.level}</span><span>${ruleSexLabel(sex)}</span></div>`; return `<div class="rule-group-item" data-group-rule-id="${esc(rule.id)}"><input class="cell-input group-condition edit-white-field" value="${esc(displayCondition(rule.condition))}" placeholder="${type === 'number' ? '130~139、<0.0001 或 >2' : '例如異常、正常、小結節'}" /><select class="level-select group-type edit-white-field"><option value="number" ${type === 'number' ? 'selected' : ''}>數字</option><option value="text" ${type === 'text' ? 'selected' : ''}>文字</option></select><select class="level-select group-level edit-white-field"><option value="1" ${rule.level === 1 ? 'selected' : ''}>1 · 低</option><option value="2" ${rule.level === 2 ? 'selected' : ''}>2 · 中</option><option value="3" ${rule.level === 3 ? 'selected' : ''}>3 · 高</option><option value="4" ${rule.level === 4 ? 'selected' : ''}>4 · 嚴重</option></select><select class="level-select group-sex edit-white-field"><option value="all" ${sex === 'all' ? 'selected' : ''}>不限性別</option><option value="male" ${sex === 'male' ? 'selected' : ''}>男性</option><option value="female" ${sex === 'female' ? 'selected' : ''}>女性</option></select><div class="group-row-actions"><button type="button" class="count-btn group-row-adjust" data-adjust-group-row="${index},-1" ${sourceRules.length <= 1 ? 'disabled' : ''}>−</button><button type="button" class="count-btn group-row-adjust" data-adjust-group-row="${index},1" ${sourceRules.length >= 4 ? 'disabled' : ''}>＋</button></div></div>`; }).join('');
  const header = editing ? `<label class="rule-group-name-editor">項目名稱<input id="editRuleGroupName" class="edit-white-field" value="${esc(state.ruleGroupName)}" placeholder="請輸入項目名稱" /><small>編輯中</small></label>` : `<div class="rule-view-title"><strong>${esc(state.ruleGroupName)}</strong><small>檢視分級規則</small></div>`;
  const alias = editing ? `<label class="rule-alias-editor">可替代名稱<small>請以逗號分隔，例如：Hb, hemoglobin, 血色素</small><input id="editRuleAliases" class="edit-white-field" value="${esc(state.ruleAliasDraft)}" placeholder="例如 Hb, hemoglobin, 血色素" /></label>` : `<div class="rule-view-alias"><strong>可替代名稱</strong><span>${esc(state.ruleAliasDraft || '尚未設定')}</span></div>`;
  const actions = editing ? `<div class="rule-group-actions"><button class="btn btn-ghost" id="cancelRuleGroupEdit">取消</button><button class="btn btn-primary" id="saveRuleGroup">儲存</button></div>` : `<div class="rule-group-actions"><button class="btn btn-primary" id="startRuleGroupEdit">修改</button></div>`;
  return `<div class="rule-group-modal"><div class="rule-group-backdrop" data-close-rule-group></div><section class="rule-group-dialog" role="dialog" aria-modal="true"><button class="privacy-close" data-close-rule-group aria-label="關閉分級規則視窗">×</button><div class="section-kicker">RULE GROUP</div><div class="rule-group-title-row"><div>${header}</div>${actions}</div>${alias}<p>${editing ? '數字條件可填寫「數值~數值」、<數值或>數值；文字條件可填寫文字。修改完成後請按「儲存」。' : '此為目前分級規則內容；若要調整項目名稱、替代名稱、條件或性別，請按「修改」。'}</p><div class="rule-group-list">${conditionRows}</div></section></div>`;
}
function renderAddRuleModal() {
  if (!state.ruleAddOpen) return '';
  const draft = state.ruleDraft;
  const levels = [1, 2, 3, 4];
  return `<div class="rule-group-modal"><div class="rule-group-backdrop" data-cancel-add-rule></div><section class="rule-group-dialog add-rule-dialog" role="dialog" aria-modal="true"><button class="privacy-close" data-cancel-add-rule aria-label="取消新增規則">×</button><div class="section-kicker">NEW RULE</div><h2>新增分級規則</h2><div class="add-rule-form"><label>項目名稱<input id="newRuleName" value="${esc(draft.name)}" placeholder="例如 收縮壓" /></label><label>條件類型<select id="newRuleType"><option value="number" ${draft.type === 'number' ? 'selected' : ''}>數字</option><option value="text" ${draft.type === 'text' ? 'selected' : ''}>文字</option></select></label></div><label class="rule-alias-editor new-rule-alias-editor">可替代名稱<small>請以逗號分隔，例如：Hb, hemoglobin, 血色素。匯入時會自動對應為此項目。</small><input id="newRuleAliases" class="edit-white-field" value="${esc(draft.aliases)}" placeholder="例如 Hb, hemoglobin, 血色素" /></label><div class="new-rule-hint">每列可選擇「不限性別、男性或女性」；選擇男性／女性後，分級時會先配對匯入資料的性別。項目名稱僅可使用中文、英文或數字；文字條件不限制字數，例如 異常、正常、小結節，但完全相同的文字不可重複。數字可填寫「數值~數值」、&lt;數值或&gt;數值，例如 130~139、&lt;0.0001、&gt;2。</div><div class="new-rule-count"><span>分級條件（${draft.levelCount} 個）</span></div><div class="new-rule-rows">${draft.conditions.slice(0, draft.levelCount).map((item, index) => `<div class="new-rule-row"><input class="new-rule-condition" data-index="${index}" value="${esc(item.condition)}" placeholder="${draft.type === 'number' ? `例如 130~139、<0.0001 或 >2` : `例如 ${['異常', '正常', '小結節', '其他'][index] || '其他'}`}" /><select class="new-rule-level" data-index="${index}">${levels.map((level) => `<option value="${level}" ${item.level === level ? 'selected' : ''}>${level} · ${levelLabels[level]}</option>`).join('')}</select><select class="new-rule-sex" data-index="${index}"><option value="all" ${(!item.sex || item.sex === 'all') ? 'selected' : ''}>不限性別</option><option value="male" ${item.sex === 'male' ? 'selected' : ''}>男性</option><option value="female" ${item.sex === 'female' ? 'selected' : ''}>女性</option></select><div class="new-rule-row-actions"><button type="button" class="count-btn" data-adjust-rule-row="${index},-1" ${draft.levelCount <= 1 ? 'disabled' : ''} aria-label="刪除第${index + 1}項條件">−</button><button type="button" class="count-btn" data-adjust-rule-row="${index},1" ${draft.levelCount >= 4 ? 'disabled' : ''} aria-label="在第${index + 1}項後新增條件">＋</button></div></div>`).join('')}</div><div class="add-rule-actions"><button class="btn btn-ghost" data-cancel-add-rule>取消</button><button class="btn btn-primary" id="confirmAddRule">確認新增</button></div></section></div>`;
}

const NTUH_FRAMINGHAM = {
  男性: { age: 3.06117, totalCholesterol: 1.12370, hdl: -0.93263, systolicUntreated: 1.93303, systolicTreated: 1.99881, smoking: 0.65451, diabetes: 0.57367, survival: 0.88936, mean: 23.9802 },
  女性: { age: 2.32888, totalCholesterol: 1.20904, hdl: -0.70833, systolicUntreated: 2.76157, systolicTreated: 2.82263, smoking: 0.52873, diabetes: 0.69154, survival: 0.95012, mean: 26.1931 }
};

function riskCategory(percent, diabetes = false) {
  if (diabetes || percent > 20) return '高危險';
  if (percent >= 10) return '中度危險';
  return '低危險';
}

function calculateFramingham(r) {
  const age = Number(r.age); const total = Number(r.totalCholesterol); const hdl = Number(r.hdl); const systolic = Number(r.systolic);
  const blank = (value) => value === undefined || value === null || String(value).trim() === '';
  const missing = [];
  if (!['女性', '男性'].includes(r.sex)) missing.push('性別');
  if (blank(r.age) || !Number.isFinite(age)) missing.push('年齡');
  if (blank(r.systolic) || !Number.isFinite(systolic)) missing.push('收縮壓');
  if (blank(r.totalCholesterol) || !Number.isFinite(total)) missing.push('總膽固醇');
  if (blank(r.hdl) || !Number.isFinite(hdl)) missing.push('高密度膽固醇（HDL）');
  if (missing.length) return { error: `缺少：${missing.join('、')}。` };
  if (r.bpMedicationProvided === false) return { error: '缺少：高血壓用藥（請填寫有或無）。' };
  if (age < 30 || age > 74) return { error: 'Framingham 心血管風險模型適用年齡為 30～74 歲。' };
  if (total <= 0 || hdl <= 0 || systolic <= 0) return { error: '總膽固醇、HDL 與收縮壓必須是大於 0 的數值。' };
  const c = NTUH_FRAMINGHAM[r.sex];
  const components = { age: Math.log(age) * c.age, totalCholesterol: Math.log(total) * c.totalCholesterol, hdl: Math.log(hdl) * c.hdl, systolic: Math.log(systolic) * (r.bpMedication ? c.systolicTreated : c.systolicUntreated), smoking: r.smoking ? c.smoking : 0, diabetes: r.diabetes ? c.diabetes : 0 };
  const linearPredictor = Object.values(components).reduce((sum, value) => sum + value, 0);
  const percent = (1 - Math.pow(c.survival, Math.exp(linearPredictor - c.mean))) * 100;
  return { percent, risk: `${percent.toFixed(1)}%`, category: riskCategory(percent, r.diabetes), components, model: 'NTUH-Framingham-log-coefficient' };
}

function flagValue(value) { return /^(y|yes|1|有|是|true|陽性|使用)$/i.test(String(value ?? '').trim()); }
function firstValue(row, keys) { const key = keys.find((name) => row[name] !== undefined && String(row[name]).trim() !== ''); return key ? row[key] : ''; }
function calculateImportedRisk(row, personIndex) {
  const rawSex = String(row['性別'] ?? '').trim(); const sex = ['女', '女性', 'F', 'Female'].includes(rawSex) ? '女性' : ['男', '男性', 'M', 'Male'].includes(rawSex) ? '男性' : '未提供';
  const medicationKeys = ['高血壓用藥', '服用高血壓藥物', '使用高血壓藥物', '降壓藥', '血壓用藥']; const medicationRaw = firstValue(row, medicationKeys); const medicationProvided = medicationKeys.some((key) => Object.prototype.hasOwnProperty.call(row, key)) && String(medicationRaw).trim() !== '';
  const input = { age: row['年齡'], sex, totalCholesterol: firstValue(row, ['總膽固醇', '總膽固醇(TC)']), hdl: firstValue(row, ['高密度脂蛋白', '高密度膽固醇', '高密度脂蛋白膽固醇', 'HDL']), systolic: row['收縮壓'], bpMedication: flagValue(medicationRaw), bpMedicationProvided: medicationProvided, smoking: flagValue(row['吸菸']), diabetes: flagValue(row['糖尿病']) };
  return { personIndex, result: calculateFramingham(input) };
}
function getImportedRuleName(header) {
  const rule = state.rules.find((item) => headersMatch(header, item.name) || (state.ruleAliases[item.name] || []).some((alias) => headersMatch(header, alias)));
  return rule?.name.trim() || null;
}
function isExportBaseHeader(header) {
  return EXPORT_BASE_COLUMNS.some((column) => headersMatch(header, column)) || NON_GRADE_IMPORT_COLUMNS.some((column) => headersMatch(header, column));
}
function getImportedUnruledHeaders() {
  return state.sourceHeaders.filter((header) => header && !isExportBaseHeader(header) && !getImportedRuleName(header));
}
function getUnruledHeaders(headers) {
  return headers.filter((header) => header && !isExportBaseHeader(header) && !getImportedRuleName(header));
}
function getExportBaseColumns() {
  const extras = state.sourceHeaders.filter((header) => header && isExportBaseHeader(header) && !EXPORT_BASE_COLUMNS.some((column) => headersMatch(header, column)));
  return [...EXPORT_BASE_COLUMNS, ...extras];
}
function renderRiskPanel() {
  return `<section class="panel risk-panel" id="risk"><div class="section-heading"><div><div class="section-kicker">OPTIONAL / FRAMINGHAM CVD</div><h2>腦心血管風險試算</h2><p>依 Framingham 心血管疾病模型計算，結果須由職護／職醫覆核。</p></div></div></section>`;
}

function buildPersonSummaries() {
  const count = state.reportMode === 'wide' ? state.sourceRows.length : Math.max(...state.results.map((item) => item.personIndex ?? 0), 0) + 1;
  return Array.from({ length: count }, (_, personIndex) => {
    const personResults = state.results.filter((item) => (item.personIndex ?? 0) === personIndex);
    const levels = personResults.map((item) => item.level).filter(Boolean);
    const sourceSerial = state.sourceRows[personIndex]?.['序號'];
    const importedRisk = state.riskResults.find((item) => item.personIndex === personIndex)?.result;
    const source = state.sourceRows[personIndex] || {};
    const ruleNames = [...new Set(state.rules.map((rule) => rule.name.trim()).filter(Boolean))];
    const missingItems = ruleNames.filter((name) => Object.prototype.hasOwnProperty.call(source, name) && !hasInputValue(source[name]));
    const unruledItems = getImportedUnruledHeaders().filter((header) => Object.prototype.hasOwnProperty.call(source, header));
    return { serial: sourceSerial || personIndex + 1, maxLevel: levels.length ? Math.max(...levels) : null, missingItems, unruledItems, cardiovascularRisk: importedRisk?.error ? `待補資料：${importedRisk.error.replace(/^缺少：/, '').replace(/。$/, '')}` : (importedRisk?.risk || '尚未計算') };
  });
}

function renderResultsPanel() {
  const summaries = buildPersonSummaries();
  const resultRows = summaries.length ? summaries.map((item) => `<tr><td><strong>${esc(item.serial)}</strong></td><td>${item.maxLevel ? `<span class="level-pill level-${item.maxLevel}">${item.maxLevel} ${levelLabels[item.maxLevel]}</span>` : '<span class="level-pill level-ungraded">未分級</span>'}</td><td>${item.missingItems.length ? `<span class="missing-pill">未提供數值：${esc(item.missingItems.join('、'))}</span>` : '<span class="complete-pill">無</span>'}</td><td><span class="risk-pill">${esc(item.cardiovascularRisk)}</span></td></tr>`).join('') : `<tr><td colspan="4"><div class="empty-results"><div class="empty-chart">${icon.chart}</div><strong>尚未產生分級結果</strong><span>請先上傳包含必要欄位的健檢報告。</span></div></td></tr>`;
  const currentFile = state.fileName ? `<span class="current-source-file">本次匯入：${esc(state.fileName)}</span>` : '';
  return `<section class="panel results-panel" id="results"><div class="section-heading"><div><div class="section-kicker">STEP 03 / RESULTS</div><p>顯示每位受檢者的序號、分級最大值、未提供數值與腦心血管風險；匯出檔會依分級結果範本排列。</p>${currentFile}</div><button class="btn btn-primary" id="exportExcel" ${state.sourceRows.length || state.results.length ? '' : 'disabled'}>${icon.download} 匯出 Excel</button></div><div class="result-table-wrap"><table class="result-table result-summary-table"><thead><tr><th>序號</th><th>分級最大值</th><th>未提供數值</th><th>腦心血管風險</th></tr></thead><tbody>${resultRows}</tbody></table></div></section>`;
}

function normalizeHeader(header) { return String(header || '').replace(/\s/g, '').toLowerCase(); }
function headersMatch(left, right) {
  const a = normalizeHeader(left); const b = normalizeHeader(right);
  if (!a || !b) return false;
  return a === b || a.startsWith(b) || (a.length >= 3 && b.startsWith(a));
}
function findColumn(headers, names) { return headers.find((header) => names.some((name) => normalizeHeader(header).includes(normalizeHeader(name)))); }

function parseRows(rows, headerRow = []) {
  const headers = headerRow.map((header) => String(header ?? '').trim());
  const nonEmptyHeaders = headers.filter(Boolean);
  const aliasPairs = Object.entries(state.ruleAliases || {}).flatMap(([canonical, aliases]) => (aliases || []).map((alias) => ({ alias, canonical })));
  const allKnownColumns = [...new Set([...REFERENCE_BASE_COLUMNS, ...REQUIRED_IMPORT_COLUMNS, ...NON_GRADE_IMPORT_COLUMNS, ...Object.values(IMPORT_FIELD_ALIASES).flat(), ...state.rules.map((rule) => rule.name.trim()), ...aliasPairs.map((item) => item.alias)])];
  const canonicalForHeader = (header) => state.rules.map((rule) => rule.name.trim()).find((column) => headersMatch(header, column)) || aliasPairs.find((item) => headersMatch(header, item.alias))?.canonical || Object.entries(IMPORT_FIELD_ALIASES).find(([, aliases]) => aliases.some((alias) => headersMatch(header, alias)))?.[0] || allKnownColumns.find((column) => headersMatch(header, column)) || header;
  const canonicalHeaders = headers.map(canonicalForHeader);
  const firstRowMatches = REQUIRED_IMPORT_COLUMNS.every((column) => canonicalHeaders.includes(column)) && canonicalHeaders[0] === '序號';
  if (!firstRowMatches) throw new Error('上傳失敗：Excel 第一排必須包含「序號、性別、年齡、糖尿病、吸菸、身高、體重、服用高血壓藥物」欄位；欄位順序可調整，但最左側第一欄必須是「序號」。');
  const normalizedRows = rows.map((row) => {
    const next = { ...row };
    headers.forEach((header) => { const canonical = canonicalForHeader(header); if (next[canonical] === undefined) next[canonical] = row[header]; });
    return next;
  });
  if (!rows.length) return { records: [], sourceRows: [], sourceHeaders: nonEmptyHeaders, mode: 'wide' };
  const invalidSerial = normalizedRows.find((row) => !/^[A-Za-z0-9]+$/.test(String(row['序號'] ?? '').trim()));
  if (invalidSerial) throw new Error('上傳失敗：最左側「序號」欄位每筆資料都必須填寫英文、數字或英數混合內容。');
  const incompleteBody = normalizedRows.find((row) => !String(row['身高'] ?? '').trim() || !String(row['體重'] ?? '').trim());
  if (incompleteBody) throw new Error('上傳失敗：每筆資料都必須填寫「身高」與「體重」；這兩欄只作為基本資料，不列入分級。');
  const ruleNames = new Set(state.rules.map((rule) => rule.name.trim()).filter(Boolean));
  const aliasNames = new Set([...Object.values(IMPORT_FIELD_ALIASES).flat(), ...aliasPairs.map((item) => item.alias)]);
  // 非分級規則欄位允許保留在匯入資料中；這些欄位不會觸發分級，也不會阻擋整份報告上傳。
  const configuredGradeColumns = [...REFERENCE_BASE_COLUMNS.slice(8), ...state.rules.map((rule) => rule.name.trim()).filter(Boolean)];
  const gradeColumns = [...new Set(configuredGradeColumns.map((name) => state.rules.find((rule) => headersMatch(rule.name, name) || (state.ruleAliases[rule.name] || []).some((alias) => headersMatch(alias, name)))?.name || name))];
  const wideFieldMap = Object.fromEntries(gradeColumns.map((name) => {
    const aliases = [...new Set([name, ...(IMPORT_FIELD_ALIASES[name] || []), ...(state.ruleAliases[name] || [])])];
    const header = headers.find((header, index) => canonicalHeaders[index] === name) || headers.find((header) => aliases.some((alias) => headersMatch(header, alias)));
    return [name, header];
  }).filter(([, header]) => header));
  const records = [];
  normalizedRows.forEach((row, personIndex) => Object.keys(wideFieldMap).forEach((name) => {
    const value = String(row[wideFieldMap[name]] ?? '').trim();
    if (value && value !== '-') records.push({ name, value, unit: '', personIndex });
  }));
  return { records, sourceRows: normalizedRows, sourceHeaders: nonEmptyHeaders, mode: 'wide' };
}

function parseNumber(value) { const match = String(value).replace(/,/g, '').replace(/[．。]/g, '.').match(/-?\d+(?:\.\d+)?/); return match ? Number(match[0]) : null; }
function parseRange(condition) {
  const clean = String(condition).replace(/[＊*]/g, '').replace(/[．。]/g, '.').trim();
  const match = clean.match(/^(-?\d+(?:\.\d+)?)\s*[~～\-至]\s*(-?\d+(?:\.\d+)?)$/);
  if (match) return { min: Number(match[1]), max: Number(match[2]), minInclusive: true, maxInclusive: true };
  const single = clean.match(/^(<=|>=|<|>)\s*(-?\d+(?:\.\d+)?)$/);
  if (!single) return null;
  const value = Number(single[2]);
  if (single[1] === '<' || single[1] === '<=') return { min: -Infinity, max: value, minInclusive: true, maxInclusive: single[1] === '<=' };
  return { min: value, max: Infinity, minInclusive: single[1] === '>=', maxInclusive: true };
}

function rangesOverlap(left, right) {
  const lower = Math.max(left.min, right.min); const upper = Math.min(left.max, right.max);
  if (lower < upper) return true;
  if (lower > upper) return false;
  const leftIncludes = (lower !== left.min || left.minInclusive) && (lower !== left.max || left.maxInclusive);
  const rightIncludes = (lower !== right.min || right.minInclusive) && (lower !== right.max || right.maxInclusive);
  return leftIncludes && rightIncludes;
}

function validNumericCondition(condition) {
  const range = parseRange(condition);
  return Boolean(range && range.min < range.max);
}

function ruleMatches(rule, record) {
  if (rule.name.trim() !== record.name.trim()) return false;
  if (!ruleSexMatches(rule, record)) return false;
  if (ruleType(rule) === 'number') { const range = parseRange(rule.condition); const parsed = parseNumber(record.value); const value = parsed === null ? null : Math.round((parsed + Number.EPSILON) * 10) / 10; return Boolean(range && value !== null && (range.minInclusive ? value >= range.min : value > range.min) && (range.maxInclusive ? value <= range.max : value < range.max)); }
  const keyword = String(rule.condition).replace(/[＊*]/g, '').trim().toLowerCase();
  return Boolean(keyword && record.value.toLowerCase().includes(keyword));
}

function gradeRecords(records) {
  return records.map((record) => {
    const match = state.rules.find((rule) => rule.enabled && ruleMatches(rule, record));
    return { ...record, matched: Boolean(match), level: match?.level ?? null, rule: match ? `${match.name} ${displayCondition(match.condition)}` : '無對應規則（請新增或確認規則）', review: Boolean(!match || match.level >= 3) };
  });
}

function detectConflicts(rules) {
  const conflicts = [];
  for (let i = 0; i < rules.length; i += 1) {
    const left = rules[i]; const leftRange = ruleType(left) === 'number' ? parseRange(left.condition) : null;
    for (let j = i + 1; j < rules.length; j += 1) {
      const right = rules[j]; if (left.name.trim() !== right.name.trim() || !ruleSexesOverlap(left.sex, right.sex)) continue;
      if (ruleType(left) === 'text' && ruleType(right) === 'text') {
        const leftText = String(left.condition ?? '').trim().toLocaleLowerCase(); const rightText = String(right.condition ?? '').trim().toLocaleLowerCase();
        if (leftText && rightText && leftText === rightText) conflicts.push([left.id, right.id]);
        continue;
      }
      if (!leftRange) continue;
      const rightRange = parseRange(right.condition);
      if (rightRange && rangesOverlap(leftRange, rightRange)) conflicts.push([rules[i].id, rules[j].id]);
    }
  }
  return conflicts;
}

function tryUpdateRule(rule, patch) {
  const candidate = { ...rule, ...patch };
  if (candidate.type === 'number') {
    const range = parseRange(candidate.condition);
    if (!validNumericCondition(candidate.condition)) { window.alert('無法儲存數字規則：請填寫 130~139、<0.0001 或 >2 等有效條件。'); render(); scrollToRules(); return false; }
  }
  const nextRules = state.rules.map((item) => item.id === rule.id ? candidate : item);
  if (detectConflicts(nextRules).some((pair) => pair.includes(rule.id))) {
    window.alert(`無法儲存「${candidate.name}」：數字條件與同名規則重疊，請改成不重疊的區間，例如 130 ~ 139、140 ~ 159。`);
    render(); scrollToRules(); return false;
  }
  Object.assign(rule, patch); saveRuleConfiguration(); render(); scrollToRules(); return true;
}

function resizeRuleDraft(levelCount) {
  const current = state.ruleDraft.conditions;
  state.ruleDraft.levelCount = Number(levelCount);
  state.ruleDraft.conditions = Array.from({ length: state.ruleDraft.levelCount }, (_, index) => current[index] || { condition: '', level: index + 1 }).map((item, index) => ({ ...item, level: Number(item.level) || index + 1 }));
}

function adjustRuleDraftCount(delta) {
  resizeRuleDraft(Math.max(1, Math.min(4, state.ruleDraft.levelCount + delta)));
  render();
}

function adjustRuleGroupRow(index, delta) {
  const nextCount = Math.max(1, Math.min(4, state.ruleEditDraft.length + delta));
  if (nextCount === state.ruleEditDraft.length) return;
  if (delta < 0) state.ruleEditDraft.splice(index, 1);
  else state.ruleEditDraft.splice(index + 1, 0, { id: `rule-edit-new-${Date.now()}-${index}`, name: state.ruleGroupName, condition: '', type: 'number', level: Math.min(index + 2, nextCount), sex: 'all', enabled: true, order: index + 1 });
  state.ruleEditDraft = state.ruleEditDraft.map((rule, rowIndex) => ({ ...rule, order: rowIndex }));
  render();
}

function adjustRuleDraftRow(index, delta) {
  const nextCount = Math.max(1, Math.min(4, state.ruleDraft.levelCount + delta));
  if (nextCount === state.ruleDraft.levelCount) return;
  const conditions = [...state.ruleDraft.conditions];
  if (delta < 0) conditions.splice(index, 1);
  else conditions.splice(index + 1, 0, { condition: '', level: Math.min(index + 2, nextCount), sex: 'all' });
  state.ruleDraft.levelCount = nextCount;
  state.ruleDraft.conditions = conditions.map((item, rowIndex) => ({ ...item, level: Number(item.level) || rowIndex + 1 }));
  render();
}

function confirmAddRule() {
  const draft = state.ruleDraft;
  if (!/^[\u4e00-\u9fffA-Za-z0-9]+$/.test(draft.name.trim())) { window.alert('項目名稱只能使用中文、英文或數字，且不可留白。'); return; }
  if (draft.conditions.some((item) => !String(item.condition).trim())) { window.alert('請填寫每一級的分級條件。'); return; }
  if (draft.type === 'number' && draft.conditions.some((item) => !validNumericCondition(item.condition))) { window.alert('數字條件格式錯誤，請使用例如 130~139、<0.0001 或 >2。'); return; }
  const newRules = draft.conditions.map((item, index) => ({ id: `rule-new-${Date.now()}-${index}`, name: draft.name.trim(), condition: draft.type === 'number' ? displayCondition(item.condition) : item.condition.trim(), type: draft.type, level: Number(item.level), sex: item.sex || 'all', enabled: true, order: index }));
  const conflicts = detectConflicts([...newRules, ...state.rules]);
  const newIds = new Set(newRules.map((rule) => rule.id));
  if (conflicts.some((pair) => pair.some((id) => newIds.has(id)))) { window.alert(draft.type === 'text' ? '無法新增：此項目已有完全相同的文字條件。相同字眼不可重複，但「結節」、「大結節」、「小結節」可以並存。' : '無法新增：此項目名稱已存在相同或重疊的數字區間。若要調整既有項目，請先使用「檢視分級規則」→「修改」，完成後按「儲存」。'); return; }
  state.rules = [...newRules, ...state.rules]; state.ruleAliases[draft.name.trim()] = cleanAliasList(draft.aliases, draft.name.trim()); saveRuleConfiguration(); state.ruleAddOpen = false; state.ruleDraft = { name: '', aliases: '', type: 'number', levelCount: 4, conditions: [1, 2, 3, 4].map((level) => ({ condition: '', level })) }; state.notice = '新增分級規則完成'; render(); scrollToRules();
}

function saveRuleGroup() {
  const draft = state.ruleEditDraft;
  if (!draft.length) { window.alert('至少保留一條分級規則，若要全部刪除請回到項目清單刪除。'); return; }
  const nextName = String(state.ruleGroupName || '').trim();
  if (!/^[\u4e00-\u9fffA-Za-z0-9()（）_-]+$/.test(nextName)) { window.alert('項目名稱只能使用中文、英文、數字、括號、底線或連字號，且不可留白。'); return; }
  const originalName = state.ruleGroupOriginalName.trim();
  const aliases = cleanAliasList(state.ruleAliasDraft, nextName);
  const otherNames = new Set(state.rules.filter((rule) => rule.name.trim() !== originalName).map((rule) => rule.name.trim()));
  if (otherNames.has(nextName)) { window.alert(`無法儲存：項目名稱「${nextName}」已存在，請使用其他名稱。`); return; }
  draft.forEach((rule) => { rule.name = nextName; });
  for (const rule of draft) {
    if (rule.type === 'number' && !validNumericCondition(rule.condition)) { window.alert(`「${rule.condition}」不是有效數字條件，請使用例如 130~139、<0.0001 或 >2。`); return; }
  }
  const others = state.rules.filter((rule) => rule.name.trim() !== originalName);
  const conflicts = detectConflicts([...draft, ...others]);
  if (conflicts.some((pair) => pair.some((id) => draft.some((rule) => rule.id === id)))) { window.alert(draft.some((rule) => rule.type === 'text') ? '無法儲存：同一項目內有完全相同的文字條件。相同字眼不可重複，但「結節」、「大結節」、「小結節」可以並存。' : '無法儲存：此項目的數值區間仍有重疊，請修改為例如 130~139、140~149。'); return; }
  state.rules = [...draft, ...others];
  if (originalName && originalName !== nextName) delete state.ruleAliases[originalName];
  state.ruleAliases[nextName] = aliases;
  saveRuleConfiguration(); state.ruleGroupEdit = false; state.ruleGroupEditing = false; state.ruleEditDraft = []; state.ruleGroupName = ''; state.ruleGroupOriginalName = ''; state.ruleAliasDraft = ''; state.notice = '分級規則儲存成功'; render(); scrollToRules();
}

function moveRule(id, direction) { const index = state.rules.findIndex((rule) => rule.id === id); const target = index + direction; if (index < 0 || target < 0 || target >= state.rules.length) return; [state.rules[index], state.rules[target]] = [state.rules[target], state.rules[index]]; saveRuleConfiguration(); render(); scrollToRules(); }
function scrollToRules() { document.querySelector('#rules')?.scrollIntoView({ behavior: 'smooth', block: 'start' }); }

function downloadTemplate() {
  const link = document.createElement('a');
  link.href = '/下載參考.xlsx';
  link.download = '職健導航_匯入健檢報告範例.xlsx';
  document.body.appendChild(link);
  link.click();
  link.remove();
}

async function handleUpload(file) {
  if (!file) return;
  const previousFileName = state.fileName;
  const previousStatus = state.uploadStatus;
  state.uploadStatus = 'uploading'; state.fileName = file.name; render();
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const headerRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true });
    const headerRow = headerRows[0] || [];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', blankrows: false });
    const parsed = parseRows(rows, headerRow);
    const unruledHeaders = getUnruledHeaders(parsed.sourceHeaders);
    if (unruledHeaders.length) {
      const proceed = window.confirm(`匯入資料包含未新增分級規則的項目：\n${unruledHeaders.join('、')}\n\n這些項目仍會保留在匯入資料與匯出結果中，但不會進行分級。\n\n是否仍要進行分級？\n\n按「確定」繼續，按「取消」停止本次匯入。`);
      if (!proceed) {
        state.fileName = previousFileName;
        state.uploadStatus = previousStatus;
        render();
        return;
      }
    }
    state.records = parsed.records; state.sourceRows = parsed.sourceRows; state.sourceHeaders = parsed.sourceHeaders; state.reportMode = parsed.mode; state.results = gradeRecords(state.records); state.riskResults = parsed.sourceRows.map((row, personIndex) => calculateImportedRisk(row, personIndex)); state.uploadedAt = new Date().toISOString(); state.uploadStatus = 'success'; render();
  } catch (error) {
    state.records = []; state.results = []; state.riskResults = []; state.sourceRows = []; state.sourceHeaders = []; state.reportMode = 'long'; state.uploadedAt = ''; state.uploadStatus = 'failed'; render();
    window.alert(error.message || '上傳失敗：請確認 Excel 第一排欄位。');
  }
}

function bindEvents() {
  if (window.__healthgradeClockTimer) clearInterval(window.__healthgradeClockTimer);
  const liveClock = document.querySelector('#liveClock');
  if (liveClock) window.__healthgradeClockTimer = setInterval(() => { liveClock.textContent = currentClock(); }, 1000);
  const privacyModal = document.querySelector('#privacyModal');
  const openPrivacy = () => { if (privacyModal) privacyModal.hidden = false; };
  document.querySelector('#privacyPolicy')?.addEventListener('click', openPrivacy);
  document.querySelector('#privacyPolicyFooter')?.addEventListener('click', openPrivacy);
  privacyModal?.querySelectorAll('[data-close-privacy]').forEach((button) => button.addEventListener('click', () => { privacyModal.hidden = true; }));
  document.querySelectorAll('[data-close-notice]').forEach((button) => button.addEventListener('click', () => { state.notice = ''; render(); }));
  document.querySelector('#closeNotice')?.addEventListener('click', () => { state.notice = ''; render(); });
  document.querySelectorAll('[data-view]').forEach((link) => link.addEventListener('click', (event) => {
    event.preventDefault();
    state.view = link.dataset.view;
    render();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }));
  document.querySelector('#excelInput')?.addEventListener('change', (event) => { const file = event.target.files[0]; event.target.value = ''; handleUpload(file); });
  const dropzone = document.querySelector('.dropzone');
  dropzone?.addEventListener('dragover', (event) => { event.preventDefault(); event.stopPropagation(); dropzone.classList.add('is-dragging'); });
  dropzone?.addEventListener('dragleave', (event) => { event.preventDefault(); event.stopPropagation(); dropzone.classList.remove('is-dragging'); });
  dropzone?.addEventListener('drop', (event) => { event.preventDefault(); event.stopPropagation(); dropzone.classList.remove('is-dragging'); handleUpload(event.dataTransfer.files[0]); });
  document.querySelector('#downloadTemplate')?.addEventListener('click', downloadTemplate);
  document.querySelector('#resetRules')?.addEventListener('click', () => { if (!window.confirm('確定要還原網站規則嗎？\n系統將以「職健導航_分級規則備份.xlsx」的 107 條規則取代目前自行調整、新增或刪除的規則。')) return; state.rules = initialRules.map((rule, index) => ({ ...rule, order: index })); state.ruleAliases = Object.fromEntries(Object.entries(DEFAULT_RULE_ALIASES).map(([name, aliases]) => [name, [...aliases]])); state.ruleGroupName = ''; state.ruleGroupOriginalName = ''; state.ruleGroupEdit = false; saveRuleConfiguration(); render(); scrollToRules(); });
  document.querySelector('#addRule')?.addEventListener('click', () => { state.ruleAddOpen = true; state.ruleGroupName = ''; state.ruleGroupOriginalName = ''; state.ruleGroupEdit = false; render(); });
  document.querySelector('#ruleSearch')?.addEventListener('input', (event) => { const query = event.target.value.toLowerCase(); document.querySelectorAll('.rule-table tbody tr').forEach((row) => { row.hidden = query && !row.textContent.toLowerCase().includes(query); }); });
  document.querySelectorAll('[data-risk]').forEach((input) => input.addEventListener('input', (event) => { const key = event.target.dataset.risk; state.risk[key] = event.target.type === 'checkbox' ? event.target.checked : event.target.value; }));
  document.querySelector('#calculateRisk')?.addEventListener('click', () => { state.risk.result = calculateFramingham(state.risk); state.risk.status = state.risk.result.error ? '資料不足' : '已計算'; render(); });
  document.querySelector('#exportExcel')?.addEventListener('click', exportResults);
  document.querySelectorAll('.group-rule').forEach((button) => button.addEventListener('click', () => { const row = button.closest('tr'); if (row?.dataset.ruleName) { state.ruleGroupName = row.dataset.ruleName; state.ruleGroupOriginalName = row.dataset.ruleName; state.ruleAliasDraft = (state.ruleAliases[row.dataset.ruleName] || []).join(', '); state.ruleGroupEdit = true; state.ruleGroupEditing = false; state.ruleEditDraft = state.rules.filter((rule) => rule.name.trim() === row.dataset.ruleName.trim()).map((rule) => ({ ...rule })); state.ruleAddOpen = false; render(); } }));
  document.querySelectorAll('.delete-rule-group').forEach((button) => button.addEventListener('click', () => { const row = button.closest('tr'); const name = row?.dataset.ruleName; if (!name || !window.confirm(`確定要刪除「${name}」的全部分級規則嗎？`)) return; state.rules = state.rules.filter((item) => item.name.trim() !== name.trim()); delete state.ruleAliases[name]; saveRuleConfiguration(); render(); scrollToRules(); }));
  document.querySelector('#startRuleGroupEdit')?.addEventListener('click', () => { state.ruleGroupEditing = true; render(); });
  document.querySelectorAll('[data-close-rule-group]').forEach((button) => button.addEventListener('click', () => { if (state.ruleGroupEditing && !window.confirm('尚未儲存的修改將會取消，確定要關閉嗎？')) return; state.ruleGroupName = ''; state.ruleGroupOriginalName = ''; state.ruleGroupEdit = false; state.ruleGroupEditing = false; state.ruleEditDraft = []; render(); }));
  document.querySelector('#saveRuleGroup')?.addEventListener('click', saveRuleGroup);
  document.querySelector('#editRuleGroupName')?.addEventListener('input', (event) => { state.ruleGroupName = event.target.value; });
  document.querySelector('#editRuleAliases')?.addEventListener('input', (event) => { state.ruleAliasDraft = event.target.value; });
  document.querySelector('#cancelRuleGroupEdit')?.addEventListener('click', () => { state.ruleGroupEditing = false; state.ruleEditDraft = state.rules.filter((rule) => rule.name.trim() === state.ruleGroupOriginalName.trim()).map((rule) => ({ ...rule })); state.ruleGroupName = state.ruleGroupOriginalName; state.ruleAliasDraft = (state.ruleAliases[state.ruleGroupOriginalName] || []).join(', '); render(); });
  document.querySelectorAll('[data-cancel-add-rule]').forEach((button) => button.addEventListener('click', () => { state.ruleAddOpen = false; render(); }));
  document.querySelector('#newRuleName')?.addEventListener('input', (event) => { state.ruleDraft.name = event.target.value; });
  document.querySelector('#newRuleAliases')?.addEventListener('input', (event) => { state.ruleDraft.aliases = event.target.value; });
  document.querySelector('#newRuleType')?.addEventListener('change', (event) => { state.ruleDraft.type = event.target.value; render(); });
  document.querySelectorAll('[data-adjust-rule-count]').forEach((button) => button.addEventListener('click', () => { adjustRuleDraftCount(Number(button.dataset.adjustRuleCount)); }));
  document.querySelectorAll('[data-adjust-rule-row]').forEach((button) => button.addEventListener('click', () => { const [index, delta] = button.dataset.adjustRuleRow.split(',').map(Number); adjustRuleDraftRow(index, delta); }));
  document.querySelectorAll('.new-rule-condition').forEach((input) => input.addEventListener('input', (event) => { state.ruleDraft.conditions[Number(event.target.dataset.index)].condition = event.target.value; }));
  document.querySelectorAll('.new-rule-level').forEach((input) => input.addEventListener('change', (event) => { state.ruleDraft.conditions[Number(event.target.dataset.index)].level = Number(event.target.value); }));
  document.querySelectorAll('.new-rule-sex').forEach((input) => input.addEventListener('change', (event) => { state.ruleDraft.conditions[Number(event.target.dataset.index)].sex = event.target.value; }));
  document.querySelector('#confirmAddRule')?.addEventListener('click', confirmAddRule);
  document.querySelectorAll('.group-condition').forEach((input) => input.addEventListener('input', (event) => { const rule = state.ruleEditDraft.find((item) => item.id === event.target.closest('[data-group-rule-id]')?.dataset.groupRuleId); if (rule) rule.condition = event.target.value; }));
  document.querySelectorAll('.group-type').forEach((input) => input.addEventListener('change', (event) => { const rule = state.ruleEditDraft.find((item) => item.id === event.target.closest('[data-group-rule-id]')?.dataset.groupRuleId); if (rule) rule.type = event.target.value; }));
  document.querySelectorAll('.group-level').forEach((input) => input.addEventListener('change', (event) => { const rule = state.ruleEditDraft.find((item) => item.id === event.target.closest('[data-group-rule-id]')?.dataset.groupRuleId); if (rule) rule.level = Number(event.target.value); }));
  document.querySelectorAll('.group-sex').forEach((input) => input.addEventListener('change', (event) => { const rule = state.ruleEditDraft.find((item) => item.id === event.target.closest('[data-group-rule-id]')?.dataset.groupRuleId); if (rule) rule.sex = event.target.value; }));
  document.querySelectorAll('.group-row-adjust').forEach((button) => button.addEventListener('click', () => { const [index, delta] = button.dataset.adjustGroupRow.split(',').map(Number); adjustRuleGroupRow(index, delta); }));
  document.querySelector('#downloadRulesBackup')?.addEventListener('click', downloadRuleBackup);
  document.querySelector('#rulesBackupInput')?.addEventListener('change', (event) => importRuleBackup(event.target.files[0]));
}

function buildReferenceRows() {
  const peopleCount = state.reportMode === 'wide' ? state.sourceRows.length : Math.max(...state.records.map((item) => item.personIndex ?? 0), 0) + 1;
  return Array.from({ length: peopleCount }, (_, personIndex) => {
    const row = {};
    const source = state.reportMode === 'wide' ? (state.sourceRows[personIndex] || {}) : {};
    REFERENCE_BASE_COLUMNS.forEach((column) => { row[column] = source[column] ?? ''; });
    if (state.reportMode === 'long') row['序號'] = personIndex + 1;
    const personResults = state.results.filter((item) => (item.personIndex ?? 0) === personIndex);
    const byName = new Map(personResults.map((item) => [item.name, item]));
    personResults.forEach((item) => { if (!row[item.name]) row[item.name] = item.value; });
    REFERENCE_GRADE_COLUMNS.forEach((name) => { row[`(分級)${name}`] = byName.get(name)?.level ?? source[`(分級)${name}`] ?? '-'; });
    const levels = personResults.map((item) => item.level).filter(Boolean);
    row['第4級'] = levels.filter((level) => level === 4).length;
    row['第3級'] = levels.filter((level) => level === 3).length;
    row['第2級'] = levels.filter((level) => level === 2).length;
    row['第1級'] = levels.filter((level) => level === 1).length;
    row['最大級別'] = levels.length ? Math.max(...levels) : (source['最大級別'] ?? '-');
    const importedRisk = state.riskResults.find((item) => item.personIndex === personIndex)?.result;
    row['10年內腦心血管風險'] = importedRisk?.error ? `待補資料：${importedRisk.error.replace(/^缺少：/, '').replace(/。$/, '')}` : (importedRisk?.risk || state.risk.result?.risk || source['10年內腦心血管風險'] || '尚未計算');
    return row;
  });
}

function hasInputValue(value) { return value !== undefined && value !== null && String(value).trim() !== '' && String(value).trim() !== '-'; }
function findSourceHeaderForRule(ruleName) {
  const candidates = [...new Set([ruleName, ...(IMPORT_FIELD_ALIASES[ruleName] || []), ...(state.ruleAliases[ruleName] || [])])].filter(Boolean);
  return state.sourceHeaders.find((header) => candidates.some((candidate) => headersMatch(header, candidate))) || null;
}

function getExportResultItems() {
  return state.sourceHeaders.filter((header) => header && !isExportBaseHeader(header)).map((header) => [header, getImportedRuleName(header)]);
}

function buildGradingResultRows() {
  const peopleCount = state.reportMode === 'wide' ? state.sourceRows.length : Math.max(...state.records.map((item) => item.personIndex ?? 0), 0) + 1;
  const exportItems = getExportResultItems();
  return Array.from({ length: peopleCount }, (_, personIndex) => {
    const source = state.reportMode === 'wide' ? (state.sourceRows[personIndex] || {}) : {};
    const personResults = state.results.filter((item) => (item.personIndex ?? 0) === personIndex);
    const row = {};
    getExportBaseColumns().forEach((column) => { row[column] = source[column] ?? source[state.sourceHeaders.find((header) => headersMatch(header, column))] ?? ''; });
    const gradedLevels = [];
    exportItems.forEach(([label, ruleName]) => {
      const aliases = [...new Set([label, ruleName, ...(IMPORT_FIELD_ALIASES[ruleName] || []), ...(state.ruleAliases[ruleName] || [])].filter(Boolean))];
      const sourceKey = aliases.find((key) => Object.prototype.hasOwnProperty.call(source, key));
      const matched = personResults.find((item) => item.name === ruleName);
      const rawValue = sourceKey ? source[sourceKey] : matched?.value ?? '';
      row[label] = rawValue ?? '';
      const ruleExists = Boolean(ruleName) && state.rules.some((rule) => rule.name === ruleName);
      if (!ruleExists) row[`__grade_${label}`] = '未新增分級規則';
      else if (!hasInputValue(rawValue)) row[`__grade_${label}`] = '無數值';
      else if (ruleExists && matched?.level) { row[`__grade_${label}`] = matched.level; gradedLevels.push(matched.level); }
      else row[`__grade_${label}`] = '-';
    });
    row['最大級別'] = gradedLevels.length ? Math.max(...gradedLevels) : '-';
    const importedRisk = state.riskResults.find((item) => item.personIndex === personIndex)?.result;
    row['10年內腦心血管風險'] = importedRisk?.error ? `待補資料：${importedRisk.error.replace(/^缺少：/, '').replace(/。$/, '')}` : (importedRisk?.risk || '尚未計算');
    return row;
  });
}

async function saveWorkbookWithPicker(workbook, filename) {
  if (typeof window.showSaveFilePicker === 'function') {
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: filename,
        types: [{ description: 'Excel 活頁簿', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
      });
      const writable = await handle.createWritable();
      await writable.write(XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }));
      await writable.close();
      return;
    } catch (error) {
      if (error?.name === 'AbortError') return;
    }
  }
  XLSX.writeFile(workbook, filename);
}

async function exportResults() {
  if (!state.sourceRows.length && !state.results.length) return;
  const rows = buildGradingResultRows();
  const exportItems = getExportResultItems();
  const resultColumns = [...getExportBaseColumns(), ...exportItems.flatMap(([label]) => [label, '分級']), '最大級別', '10年內腦心血管風險'];
  const sheetRows = [resultColumns, ...rows.map((row) => {
    let itemIndex = 0;
    return resultColumns.map((column) => {
      if (column === '分級') { const label = exportItems[itemIndex][0]; itemIndex += 1; return row[`__grade_${label}`]; }
      return row[column] ?? '';
    });
  })];
  const worksheet = XLSX.utils.aoa_to_sheet(sheetRows);
  worksheet['!cols'] = resultColumns.map((column) => ({ wch: column === '分級' ? 9 : (column.length > 8 ? 16 : 11) }));
  worksheet['!autofilter'] = { ref: `A1:${XLSX.utils.encode_col(resultColumns.length - 1)}${rows.length + 1}` };
  const gradeStyles = {
    '1': { font: { color: { rgb: '237765' }, bold: true }, fill: { fgColor: { rgb: 'E3F4ED' } } },
    '2': { font: { color: { rgb: '8B691F' }, bold: true }, fill: { fgColor: { rgb: 'FFF2D2' } } },
    '3': { font: { color: { rgb: 'B45B2D' }, bold: true }, fill: { fgColor: { rgb: 'FFE8D8' } } },
    '4': { font: { color: { rgb: 'A33F3F' }, bold: true }, fill: { fgColor: { rgb: 'FDE0E0' } } }
  };
  resultColumns.forEach((column, columnIndex) => {
    if (column !== '分級' && column !== '最大級別') return;
    for (let rowIndex = 2; rowIndex <= rows.length + 1; rowIndex += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex - 1, c: columnIndex })];
      const style = cell ? gradeStyles[String(cell.v)] : null;
      if (cell && style) cell.s = style;
    }
  });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '分級結果');
  const sourceBase = (state.fileName || '健檢報告').replace(/\.[^.]+$/, '').replace(/[^A-Za-z0-9\u4e00-\u9fff_-]+/g, '_').slice(0, 60) || '健檢報告';
  const exportStamp = (state.uploadedAt || new Date().toISOString()).replace(/[-:T]/g, '').slice(0, 14);
  await saveWorkbookWithPicker(workbook, `職健導航_分級結果_${sourceBase}_${exportStamp}.xlsx`);
}

render();
