const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { test } = require('node:test');
const XLSX = require('../vendor/xlsx.full.min.js');

function loadApp() {
  const elements = new Map();
  const context = vm.createContext({
    console,
    document: { querySelector(selector) {
      if (!elements.has(selector)) elements.set(selector, {});
      return elements.get(selector);
    } },
    window: { XLSX: { ...XLSX } },
  });
  const source = fs.readFileSync(path.join(__dirname, '../file-library.js'), 'utf8');
  vm.runInContext(source.replace(/\ninit\(\)\.catch\([\s\S]*$/, ''), context);
  return { context, elements, run: (code) => vm.runInContext(code, context) };
}

test('duplicate quantity headers select demand detail, never SKU digits or transfer quantity', () => {
  const { context, run } = loadApp();
  context.headers = ['创建人', '数量', '品名', '数量', '要求货好时间', 'sku', '识别码'];
  context.row = ['测试申请人', 99, '测试产品', '7', '2026-09-25', 'YT02-WT-C', 'TEST-001'];
  assert.equal(run('findDemandQuantityColumnIndex(headers, { requiredReadyDate: 4 })'), 3);
  assert.equal(run('getDemandRowQuantity(row, { quantity: 3, materialCode: 6 })'), 7);
  assert.equal(run('findDemandQuantityColumnIndex(["识别码", "数量"], {})'), 1);
  assert.equal(run('findDemandQuantityColumnIndex(["数量", "数量", "识别码"], {})'), undefined);
  assert.equal(run('findDemandQuantityColumnIndex(["识别码", "sku"], {})'), undefined);
});

test('missing and invalid quantities never fall back; zero and valid signed decimals survive', () => {
  const { context, run } = loadApp();
  for (const [input, expected] of [[0, 0], ['0', 0], [' 7 ', 7], ['1,234.5', 1234.5], [-2, -2], ['+0.5', 0.5], [Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER]]) {
    context.input = input;
    assert.equal(run('getDemandRowQuantity([input, "YT02-WT-C", "2026-09-25"], { quantity: 0 })'), expected);
  }
  for (const input of ['', ' ', null, undefined, NaN, Infinity, 'YT02-WT-C', '2026-09-25', '7个', '1,2', '#N/A', 'Infinity']) {
    context.input = input;
    assert(Number.isNaN(run('getDemandRowQuantity([input, 7, "YT02-WT-C"], { quantity: 0 })')));
  }
  assert(Number.isNaN(run('getDemandRowQuantity([7, "YT02-WT-C"], {})')));
});

test('workbook import, rows, metrics, minimum-order comparison and export use detail quantities', async () => {
  const { context, elements, run } = loadApp();
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
    ['创建人', '数量', '品名', '数量', '要求货好时间', 'sku', '识别码'],
    ['测试申请人', 99, '产品A', 7, '2026-09-25', 'YT02-WT-C', 'TEST-001'],
    ['测试申请人', 99, '产品B', 3, '2026-09-25', 'YT02-WT-PRO-C', 'TEST-002'],
    ['测试申请人', 99, '产品C', 0, '2026-09-25', 'SKU09', 'TEST-003'],
  ]), '需要备货');
  const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
  context.records = { demand: { name: 'synthetic.xlsx', file: { name: 'synthetic.xlsx', arrayBuffer: async () => bytes } } };
  const rows = await run('buildDemandAllocationRows(records)');
  assert.deepEqual(Array.from(rows, row => row.quantityNumber), [7, 3, 0]);
  context.rows = rows;
  run('state.filteredRows = rows; updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '10');
  assert.equal(run('getMinimumOrderStatus(rows[0].quantityNumber, 5)'), '满足');
  assert.match(run('renderDemandTableRow(rows[0])'), /<td>7<\/td>/);
  let exported;
  context.window.XLSX.writeFile = (output) => { exported = output; };
  await run('downloadDetailWorkbook()');
  const output = XLSX.utils.sheet_to_json(exported.Sheets[exported.SheetNames[0]], { header: 1 });
  const index = output[0].indexOf('数量');
  assert.deepEqual(output.slice(1).map(row => row[index]), ['7', '3', '0']);
  run('state.filteredRows = [{ quantityNumber: NaN }]; updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '待核对');
  assert.equal(elements.get('#quantityTotal').textContent, '数量合计：0（1 条数量待核对）');
  run('state.filteredRows = []; updateDemandMetrics()');
  assert.equal(elements.get('#beihuoTotal').textContent, '0');
});
