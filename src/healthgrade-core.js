export function normalizeHeader(header) {
  return String(header || '').replace(/\s/g, '').toLowerCase();
}

export function headersMatch(left, right) {
  const a = normalizeHeader(left); const b = normalizeHeader(right);
  if (!a || !b) return false;
  return a === b || a.startsWith(b) || (a.length >= 3 && b.startsWith(a));
}

export function serialKey(row) {
  return String(row?.['序號'] ?? '').trim().replace(/\s+/g, '').toUpperCase();
}

export function parseNumber(value) {
  const match = String(value).replace(/,/g, '').replace(/[．。]/g, '.').match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : null;
}

export function roundOneDecimal(value) {
  return Math.round((Number(value) + Number.EPSILON) * 10) / 10;
}

export function parseRange(condition) {
  const clean = String(condition).replace(/[＊*]/g, '').replace(/[．。]/g, '.').trim();
  const match = clean.match(/^(-?\d+(?:\.\d+)?)\s*[~～\-至]\s*(-?\d+(?:\.\d+)?)$/);
  if (match) return { min: Number(match[1]), max: Number(match[2]), minInclusive: true, maxInclusive: true };
  const single = clean.match(/^(<=|>=|<|>)\s*(-?\d+(?:\.\d+)?)$/);
  if (!single) return null;
  const value = Number(single[2]);
  if (single[1] === '<' || single[1] === '<=') return { min: -Infinity, max: value, minInclusive: true, maxInclusive: single[1] === '<=' };
  return { min: value, max: Infinity, minInclusive: single[1] === '>=', maxInclusive: true };
}

export function rangesOverlap(left, right) {
  const lower = Math.max(left.min, right.min); const upper = Math.min(left.max, right.max);
  if (lower < upper) return true;
  if (lower > upper) return false;
  const leftIncludes = (lower !== left.min || left.minInclusive) && (lower !== left.max || left.maxInclusive);
  const rightIncludes = (lower !== right.min || right.minInclusive) && (lower !== right.max || right.maxInclusive);
  return leftIncludes && rightIncludes;
}

export function validNumericCondition(condition) {
  const range = parseRange(condition);
  return Boolean(range && range.min < range.max);
}

export function uniqueSerialCount(rows) {
  return new Set(rows.filter((row) => String(row?.['序號'] ?? '').trim() !== '').map(serialKey)).size;
}

export const NTUH_FRAMINGHAM = {
  男性: { age: 3.06117, totalCholesterol: 1.12370, hdl: -0.93263, systolicUntreated: 1.93303, systolicTreated: 1.99881, smoking: 0.65451, diabetes: 0.57367, survival: 0.88936, mean: 23.9802 },
  女性: { age: 2.32888, totalCholesterol: 1.20904, hdl: -0.70833, systolicUntreated: 2.76157, systolicTreated: 2.82263, smoking: 0.52873, diabetes: 0.69154, survival: 0.95012, mean: 26.1931 },
};

export function riskCategory(percent, diabetes = false) {
  if (diabetes || percent > 20) return '高危險';
  if (percent >= 10) return '中度危險';
  return '低危險';
}

export function calculateFramingham(r) {
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
