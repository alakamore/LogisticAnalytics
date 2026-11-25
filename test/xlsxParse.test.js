import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const commonCode = fs.readFileSync('./shared/js/common.js', 'utf8');
const xlsxCode = fs.readFileSync('./shared/js-lib/xlsx.full.min.js', 'utf8');

const xlsxSandbox = {};
const XLSX = new Function(
    'self',
    'window',
    'globalThis',
    'module',
    'exports',
    `${xlsxCode}\nreturn (typeof XLSX !== 'undefined' ? XLSX : self.XLSX);`
)(xlsxSandbox, xlsxSandbox, xlsxSandbox, undefined, undefined);

const noop = () => {};
const fakeDocument = {
    readyState: 'loading',
    getElementById: () => null,
    querySelectorAll: () => [],
    addEventListener: noop,
};
const fakeWindow = {
    addEventListener: noop,
    removeEventListener: noop,
    dispatchEvent: () => true,
};

/** Подгружает common.js с указанной реализацией XLSX и FileReader. */
function loadCommon(realXlsx, fileResult) {
    class FakeFileReader {
        readAsArrayBuffer() {
            queueMicrotask(() => this.onload({ target: { result: fileResult } }));
        }
    }
    const wrapper = new Function(
        'window',
        'document',
        'XLSX',
        'Chart',
        'console',
        'alert',
        'navigator',
        'FileReader',
        `${commonCode}\nreturn { expandSheetRange, parseXLSXFile };`
    );
    return wrapper(fakeWindow, fakeDocument, realXlsx, function () {}, console, noop, {}, FakeFileReader);
}

const common = loadCommon(XLSX, new ArrayBuffer(0));

/* ---------- expandSheetRange ---------- */

test('expandSheetRange: расширяет занижённый !ref по фактическим ячейкам', () => {
    const sheet = XLSX.utils.aoa_to_sheet([
        ['Код', 'Документ', 'Сумма'],
        ['A1', 'D1', 10],
        ['A2', 'D2', 20],
        ['A3', 'D3', 30],
    ]);
    sheet['!ref'] = 'A1:B1';

    common.expandSheetRange(sheet);

    assert.strictEqual(sheet['!ref'], 'A1:C4');
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    assert.strictEqual(rows.length, 4);
    assert.deepStrictEqual(rows[3], ['A3', 'D3', 30]);
});

test('expandSheetRange: не сужает корректный !ref и учитывает пустые строки в конце', () => {
    const sheet = XLSX.utils.aoa_to_sheet([['A'], ['B']]);
    sheet['!ref'] = 'A1:C10';

    common.expandSheetRange(sheet);

    assert.strictEqual(sheet['!ref'], 'A1:C10');
});

test('expandSheetRange: лист без ячеек оставляет !ref как есть', () => {
    const sheet = { '!ref': 'A1:B2' };

    common.expandSheetRange(sheet);

    assert.strictEqual(sheet['!ref'], 'A1:B2');
});

/* ---------- parseXLSXFile ---------- */

test('parseXLSXFile: возвращает все строки, если файл сохранён с занижённым !ref', async () => {
    const data = [['Код аналитики', 'Документ', 'Сумма']];
    for (let i = 1; i <= 50; i++) data.push([`PMG${i}`, `DOC${i}`, i]);
    const sheet = XLSX.utils.aoa_to_sheet(data);
    const fullRows = data.length;
    sheet['!ref'] = 'A1:S5'; // как в файле, сохранённом сразу после экспорта

    const brokenXlsx = {
        ...XLSX,
        read: () => ({ SheetNames: ['Лист1'], Sheets: { Лист1: sheet } }),
    };
    const parser = loadCommon(brokenXlsx, new ArrayBuffer(8));
    const file = { name: 'Сохранён сразу.xlsx' };

    const rows = await parser.parseXLSXFile(file);

    assert.strictEqual(rows.length, fullRows);
    assert.deepStrictEqual(rows[0].slice(0, 3), ['Код аналитики', 'Документ', 'Сумма']);
    assert.deepStrictEqual(rows[fullRows - 1].slice(0, 3), ['PMG50', 'DOC50', 50]);
});

test('parseXLSXFile: отклоняет файл неподдерживаемого формата', async () => {
    await assert.rejects(common.parseXLSXFile({ name: 'report.csv' }), /\.xlsx или \.xls/);
});

