import test from 'node:test';
import assert from 'node:assert/strict';
import {
  headersMatch,
  normalizeHeader,
  parseNumber,
  parseRange,
  rangesOverlap,
  roundOneDecimal,
  serialKey,
  uniqueSerialCount,
  validNumericCondition,
} from '../src/healthgrade-core.js';

test('數值解析可處理千分位與全形小數點', () => {
  assert.equal(parseNumber('1,234.56 mg/dL'), 1234.56);
  assert.equal(parseNumber('BMI 24．96'), 24.96);
  assert.equal(parseNumber('無數值'), null);
});

test('小數點兩位以上四捨五入至一位', () => {
  assert.equal(roundOneDecimal(24.96), 25.0);
  assert.equal(roundOneDecimal(24.94), 24.9);
});

test('可解析包含端點的數值區間', () => {
  assert.deepEqual(parseRange('130~139'), { min: 130, max: 139, minInclusive: true, maxInclusive: true });
  assert.equal(validNumericCondition('130~139'), true);
  assert.equal(validNumericCondition('139~130'), false);
});

test('可解析單邊條件並正確處理邊界', () => {
  assert.deepEqual(parseRange('<0.0001'), { min: -Infinity, max: 0.0001, minInclusive: true, maxInclusive: false });
  assert.deepEqual(parseRange('>2'), { min: 2, max: Infinity, minInclusive: false, maxInclusive: true });
  const less = parseRange('<0.0001');
  const greater = parseRange('>2');
  assert.equal(0.0001 < less.max, false);
  assert.equal(2 > greater.min, false);
});

test('相鄰區間不重疊，真正交集才算重疊', () => {
  assert.equal(rangesOverlap(parseRange('130~139'), parseRange('140~159')), false);
  assert.equal(rangesOverlap(parseRange('130~140'), parseRange('140~159')), true);
});

test('欄名正規化與模糊對應支援別名', () => {
  assert.equal(normalizeHeader(' Hb '), 'hb');
  assert.equal(headersMatch('BMI(%)', 'BMI'), true);
  assert.equal(headersMatch('TSH甲狀腺', 'TSH'), true);
  assert.equal(headersMatch('家庭病史', '血糖'), false);
});

test('序號去重不受空白與英文字母大小寫影響', () => {
  assert.equal(serialKey({ 序號: ' ab 01 ' }), 'AB01');
  assert.equal(uniqueSerialCount([{ 序號: 'A01' }, { 序號: ' a 01 ' }, { 序號: 'B02' }, { 序號: '' }]), 2);
});

test('序號可接受英文、數字與英數混合，空白序號不計入人數', () => {
  const rows = [{ 序號: '001' }, { 序號: 'EMPA2' }, { 序號: 'Z9' }, { 序號: ' ' }];
  assert.equal(uniqueSerialCount(rows), 3);
});


test('Framingham 缺少必要欄位時回傳明確提示', async () => {
  const { calculateFramingham } = await import('../src/healthgrade-core.js');
  const result = calculateFramingham({ sex: '男性', age: 45, systolic: 130, totalCholesterol: 200 });
  assert.match(result.error, /高密度膽固醇/);
});

test('Framingham 接受 30～74 歲完整資料並回傳百分比與模型名稱', async () => {
  const { calculateFramingham } = await import('../src/healthgrade-core.js');
  const result = calculateFramingham({ sex: '男性', age: 45, systolic: 130, totalCholesterol: 200, hdl: 50, bpMedicationProvided: true, bpMedication: false, smoking: false, diabetes: false });
  assert.equal(result.model, 'NTUH-Framingham-log-coefficient');
  assert.equal(typeof result.percent, 'number');
  assert.match(result.risk, /^\d+\.\d%$/);
});

test('Framingham 拒絕模型範圍外年齡與缺少用藥欄位', async () => {
  const { calculateFramingham } = await import('../src/healthgrade-core.js');
  const base = { sex: '女性', systolic: 120, totalCholesterol: 180, hdl: 60, bpMedication: false, smoking: false, diabetes: false };
  assert.match(calculateFramingham({ ...base, age: 29, bpMedicationProvided: true }).error, /30～74/);
  assert.match(calculateFramingham({ ...base, age: 50, bpMedicationProvided: false }).error, /高血壓用藥/);
});
