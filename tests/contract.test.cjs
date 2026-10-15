const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

function loadScript(overrides = {}) {
  const source = fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8');
  const context = vm.createContext({ ...overrides });
  vm.runInContext(source, context);
  return context;
}

function makeTable(rows) {
  const table = {
    rows: rows.map((values) => {
      const row = {
        cells: [],
        getNumCells() { return this.cells.length; },
        getCell(index) { return this.cells[index]; }
      };
      row.cells = values.map((value) => ({
        text: value,
        setWidth(width) { this.width = width; },
        merge() {
          const index = row.cells.indexOf(this);
          row.cells[index - 1].text += this.text;
          row.cells.splice(index, 1);
        },
        clear() { this.text = ''; },
        appendParagraph(value) { this.text += value; }
      }));
      return row;
    }),
    getNumRows() { return this.rows.length; },
    getRow(index) { return this.rows[index]; }
  };
  return table;
}

test('one equipment item uses only two columns within table width', () => {
  const context = loadScript();
  context.styleContractTable_ = () => {};
  let table;
  context.body = { appendTable(rows) { table = makeTable(rows); return table; } };
  context.rowData = { '租借設備清單': 'vivo X300 Ultra 12/256GB' };
  vm.runInContext('appendContractEquipmentTable_(body, rowData)', context);

  assert.deepEqual(Array.from(table.rows, (row) => row.getNumCells()), [2, 2]);
  assert.equal(table.getRow(0).getCell(0).width, 38);
  assert.equal(table.getRow(0).getCell(1).width, 482);
  assert.equal(table.getRow(1).getCell(1).text, 'vivo X300 Ultra 12/256GB');
});

test('odd final equipment item spans the remaining table width', () => {
  const context = loadScript();
  context.styleContractTable_ = () => {};
  let table;
  context.body = { appendTable(rows) { table = makeTable(rows); return table; } };
  context.rowData = { '租借設備清單': 'A\nB\nC' };
  vm.runInContext('appendContractEquipmentTable_(body, rowData)', context);

  assert.deepEqual(Array.from(table.rows, (row) => row.getNumCells()), [4, 4, 2]);
  assert.equal(table.getRow(2).getCell(1).text, 'C');
  assert.equal(table.getRow(2).getCell(0).width, 38);
  assert.equal(table.getRow(2).getCell(1).width, undefined);
});

test('existing contract files block generation even if status was changed', () => {
  const context = loadScript();
  const rowData = { '預約編號': 'G1234501', '合約狀態': '草稿', '合約PDF': 'https://example.com/contract.pdf' };
  let createCount = 0;
  let releaseCount = 0;
  context.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => { releaseCount += 1; } }) };
  context.getRowContext_ = () => ({ headers: [], rowData });
  context.validateContractDetail_ = () => {};
  context.createContractFiles_ = () => { createCount += 1; return {}; };
  context.writeContractResult_ = () => '';
  context.sheet = {};

  assert.throws(() => vm.runInContext('generateContractForRow_(sheet, 2)', context), /已有合約/);
  assert.equal(createCount, 0);
  assert.equal(releaseCount, 1);
});

test('a second generation request cannot replace the first contract', () => {
  const context = loadScript();
  const rowData = { '預約編號': 'G1234501', '合約狀態': '草稿' };
  let createCount = 0;
  context.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
  context.getRowContext_ = () => ({ headers: [], rowData });
  context.validateContractDetail_ = () => {};
  context.createContractFiles_ = () => { createCount += 1; return { pdfUrl: 'https://example.com/first.pdf' }; };
  context.writeContractResult_ = () => {
    rowData['合約狀態'] = '已產生';
    rowData['合約PDF'] = 'https://example.com/first.pdf';
    return 'https://example.com/sign';
  };
  context.sheet = {};

  vm.runInContext('generateContractForRow_(sheet, 2)', context);
  assert.throws(() => vm.runInContext('generateContractForRow_(sheet, 2)', context), /已有合約/);
  assert.equal(createCount, 1);
  assert.equal(rowData['合約PDF'], 'https://example.com/first.pdf');
});

test('draft refresh preserves existing formulas and manual values', () => {
  const context = loadScript();
  context.findRowByHeaderValue_ = () => 2;
  context.ensureSheetSize_ = () => {};
  context.formatContractSheet_ = () => {};
  const headers = ['預約編號', '合約狀態', '出租人姓名', '簽名連結', '租借設備清單'];
  const values = ['G1234501', '草稿', '手動姓名', '點我簽名', '舊設備'];
  const formulas = ['', '', '', '=HYPERLINK("https://example.com/sign","點我簽名")', ''];
  const writes = [];
  const sheet = {
    getLastRow: () => 2,
    getMaxRows: () => 2,
    getRange(row, column, height, width) {
      if (width) {
        return { getDisplayValues: () => [values], getFormulas: () => [formulas] };
      }
      return { setValue(value) { writes.push([column, value]); values[column - 1] = value; } };
    }
  };
  context.sheet = sheet;
  context.headers = headers;
  context.detail = { '預約編號': 'G1234501', '合約狀態': '草稿', '出租人姓名': '預設姓名', '租借設備清單': '新設備' };
  vm.runInContext('upsertContractDetail_(sheet, headers, detail)', context);

  assert.deepEqual(writes, [[5, '新設備']]);
  assert.equal(values[2], '手動姓名');
  assert.match(formulas[3], /^=HYPERLINK/);
});

test('generated contract detail is not refreshed over its linked files', () => {
  const context = loadScript();
  context.findRowByHeaderValue_ = () => 2;
  context.ensureSheetSize_ = () => {};
  context.formatContractSheet_ = () => {};
  const headers = ['預約編號', '合約狀態', '合約文件', '租借設備清單'];
  const values = ['G1234501', '已產生', 'https://example.com/first', '原設備'];
  const writes = [];
  context.sheet = {
    getLastRow: () => 2,
    getMaxRows: () => 2,
    getRange(row, column, height, width) {
      if (width) {
        return { getDisplayValues: () => [values], getFormulas: () => [[...headers.map(() => '')]] };
      }
      return { setValue(value) { writes.push([column, value]); } };
    }
  };
  context.headers = headers;
  context.detail = { '預約編號': 'G1234501', '租借設備清單': '新設備' };

  assert.throws(() => vm.runInContext('upsertContractDetail_(sheet, headers, detail)', context), /已有合約/);
  assert.deepEqual(writes, []);
});
