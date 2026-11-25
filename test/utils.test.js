import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const code = fs.readFileSync('./shared/js/common.js', 'utf8');

const globalWindow = {
    addEventListener: () => {},
    removeEventListener: () => {},
    dispatchEvent: () => true,
};
const globalDocument = {
    getElementById: () => null,
    querySelectorAll: () => [],
    addEventListener: () => {},
};
const lastExport = { fileName: null, sheetName: null, rows: null };
const globalXLSX = {
    read: () => ({}),
    writeFile: (wb, fileName) => {
        lastExport.fileName = fileName;
    },
    utils: {
        sheet_to_json: () => [],
        aoa_to_sheet: (rows) => ({ rows }),
        book_new: () => ({}),
        book_append_sheet: (wb, ws, sheetName) => {
            lastExport.sheetName = sheetName;
            lastExport.rows = ws.rows;
        },
    },
};
const globalChart = function () {};
const alerts = [];
const globalAlert = (message) => alerts.push(message);

const wrapper = new Function(
    'window',
    'document',
    'XLSX',
    'Chart',
    'console',
    'alert',
    `${code}\nreturn { parseNumber, parseNullableNumber, parseDateCell, parseTimeCell, findColumnIndex, cleanValue, sortData, debounce, escapeHtml, validateXLSXHeaders, exportRowsToXLSX, getPageCount, getPageSlice };`
);

const {
    parseNumber,
    parseNullableNumber,
    parseDateCell,
    parseTimeCell,
    findColumnIndex,
    cleanValue,
    sortData,
    debounce,
    escapeHtml,
    validateXLSXHeaders,
    exportRowsToXLSX,
    getPageCount,
    getPageSlice,
} = wrapper(
    globalWindow,
    globalDocument,
    globalXLSX,
    globalChart,
    console,
    globalAlert
);

/* ---------- parseNumber ---------- */

test('parseNumber: число без изменений', () => {
    assert.strictEqual(parseNumber(42), 42);
});

test('parseNumber: строка с пробелами в тысячах', () => {
    assert.strictEqual(parseNumber('1 234 567'), 1234567);
});

test('parseNumber: строка с запятой как разделителем', () => {
    assert.strictEqual(parseNumber('1,5'), 1.5);
});

test('parseNumber: пустая строка', () => {
    assert.strictEqual(parseNumber(''), 0);
});

test('parseNumber: null / undefined', () => {
    assert.strictEqual(parseNumber(null), 0);
    assert.strictEqual(parseNumber(undefined), 0);
});

test('parseNumber: нечисловая строка', () => {
    assert.strictEqual(parseNumber('abc'), 0);
});

test('parseNumber: строка с пробелами и запятой', () => {
    assert.strictEqual(parseNumber('1 234,56'), 1234.56);
});

/* ---------- parseNullableNumber ---------- */

test('parseNullableNumber: пустые значения дают null, а не 0', () => {
    assert.strictEqual(parseNullableNumber(''), null);
    assert.strictEqual(parseNullableNumber('   '), null);
    assert.strictEqual(parseNullableNumber(null), null);
    assert.strictEqual(parseNullableNumber(undefined), null);
});

test('parseNullableNumber: нечисловая строка даёт null', () => {
    assert.strictEqual(parseNullableNumber('abc'), null);
    assert.strictEqual(parseNullableNumber('12 кг'), null);
});

test('parseNullableNumber: числа и числовые строки', () => {
    assert.strictEqual(parseNullableNumber(0), 0);
    assert.strictEqual(parseNullableNumber(42), 42);
    assert.strictEqual(parseNullableNumber('1 234,5'), 1234.5);
    assert.strictEqual(parseNullableNumber('-7,25'), -7.25);
    assert.strictEqual(parseNullableNumber('+3'), 3);
});

test('parseNullableNumber: нечисловые числа (NaN, Infinity)', () => {
    assert.strictEqual(parseNullableNumber(NaN), null);
    assert.strictEqual(parseNullableNumber(Infinity), null);
});

/* ---------- parseDateCell ---------- */

test('parseDateCell: объект Date возвращается как есть', () => {
    const date = new Date(2024, 2, 15, 10, 30);
    assert.strictEqual(parseDateCell(date), date);
});

test('parseDateCell: dd.mm.yyyy', () => {
    const date = parseDateCell('15.03.2024');
    assert.ok(date instanceof Date);
    assert.strictEqual(date.getFullYear(), 2024);
    assert.strictEqual(date.getMonth(), 2);
    assert.strictEqual(date.getDate(), 15);
});

test('parseDateCell: dd.mm.yy', () => {
    const date = parseDateCell('15.03.24');
    assert.strictEqual(date.getFullYear(), 2024);
});

test('parseDateCell: yyyy-mm-dd', () => {
    const date = parseDateCell('2024-03-15');
    assert.strictEqual(date.getFullYear(), 2024);
    assert.strictEqual(date.getMonth(), 2);
    assert.strictEqual(date.getDate(), 15);
});

test('parseDateCell: ISO-строка', () => {
    const date = parseDateCell('2024-03-15T10:30:00');
    assert.strictEqual(date.getFullYear(), 2024);
    assert.strictEqual(date.getMonth(), 2);
    assert.strictEqual(date.getDate(), 15);
});

test('parseDateCell: Excel serial 45292 - это 1 января 2024', () => {
    const date = parseDateCell(45292);
    assert.strictEqual(date.getFullYear(), 2024);
    assert.strictEqual(date.getMonth(), 0);
    assert.strictEqual(date.getDate(), 1);
});

test('parseDateCell: некорректные даты и пустые значения', () => {
    assert.strictEqual(parseDateCell(''), null);
    assert.strictEqual(parseDateCell(null), null);
    assert.strictEqual(parseDateCell('не дата'), null);
    assert.strictEqual(parseDateCell('32.01.2024'), null);
    assert.strictEqual(parseDateCell('31.02.2024'), null);
    assert.strictEqual(parseDateCell(0), null);
    assert.strictEqual(parseDateCell(-1), null);
});

/* ---------- parseTimeCell ---------- */

test('parseTimeCell: строка ЧЧ:ММ', () => {
    assert.strictEqual(parseTimeCell('13:45'), 13 * 60 + 45);
    assert.strictEqual(parseTimeCell('7:30'), 7 * 60 + 30);
    assert.strictEqual(parseTimeCell('0:0'), 0);
    assert.strictEqual(parseTimeCell('23:59'), 23 * 60 + 59);
});

test('parseTimeCell: строка с секундами', () => {
    assert.strictEqual(parseTimeCell('08:15:30'), 8 * 60 + 15);
});

test('parseTimeCell: Excel-доля суток', () => {
    assert.strictEqual(parseTimeCell(0.5), 12 * 60);
    assert.strictEqual(parseTimeCell(0.25), 6 * 60);
    assert.strictEqual(parseTimeCell(0), 0);
});

test('parseTimeCell: объект Date', () => {
    assert.strictEqual(parseTimeCell(new Date(2024, 0, 1, 13, 45)), 13 * 60 + 45);
});

test('parseTimeCell: некорректные значения', () => {
    assert.strictEqual(parseTimeCell(''), null);
    assert.strictEqual(parseTimeCell(null), null);
    assert.strictEqual(parseTimeCell('24:00'), null);
    assert.strictEqual(parseTimeCell('13:60'), null);
    assert.strictEqual(parseTimeCell('abc'), null);
});

/* ---------- findColumnIndex ---------- */

test('findColumnIndex: точное совпадение', () => {
    const headers = ['АДРЕС', 'ЗОНА', 'LM Код'];
    assert.strictEqual(findColumnIndex(headers, ['АДРЕС']), 0);
    assert.strictEqual(findColumnIndex(headers, ['ЗОНА']), 1);
    assert.strictEqual(findColumnIndex(headers, ['LM Код']), 2);
});

test('findColumnIndex: частичное совпадение', () => {
    const headers = ['Адрес (код)', 'Зона хранения', 'Код ЛМ'];
    assert.strictEqual(findColumnIndex(headers, ['АДРЕС']), 0);
    assert.strictEqual(findColumnIndex(headers, ['ЗОНА']), 1);
    assert.strictEqual(findColumnIndex(headers, ['ЛМ']), 2);
});

test('findColumnIndex: регистронезависимый поиск', () => {
    const headers = ['адрес', 'зона'];
    assert.strictEqual(findColumnIndex(headers, ['АДРЕС']), 0);
    assert.strictEqual(findColumnIndex(headers, ['ЗОНА']), 1);
});

test('findColumnIndex: не найдено', () => {
    const headers = ['АДРЕС', 'ЗОНА'];
    assert.strictEqual(findColumnIndex(headers, ['ЛМ Код']), -1);
});

test('findColumnIndex: пустой заголовок', () => {
    assert.strictEqual(findColumnIndex([], ['АДРЕС']), -1);
});

/* ---------- cleanValue ---------- */

test('cleanValue: строка с пробелами по краям', () => {
    assert.strictEqual(cleanValue('  hello  '), 'hello');
});

test('cleanValue: null / undefined', () => {
    assert.strictEqual(cleanValue(null), '');
    assert.strictEqual(cleanValue(undefined), '');
});

test('cleanValue: число', () => {
    assert.strictEqual(cleanValue(123), '123');
});

/* ---------- sortData ---------- */

test('sortData: строки по возрастанию', () => {
    const data = [
        { name: 'Б', val: 2 },
        { name: 'А', val: 1 },
        { name: 'В', val: 3 },
    ];
    const res = sortData(data, 'name', null, false);
    assert.strictEqual(res.sortColumn, 'name');
    assert.strictEqual(res.sortDirection, false);
    assert.deepStrictEqual(res.data.map((d) => d.name), ['А', 'Б', 'В']);
});

test('sortData: строки по убыванию при повторном клике', () => {
    const data = [
        { name: 'Б', val: 2 },
        { name: 'А', val: 1 },
    ];
    const first = sortData(data, 'name', null, false);
    const second = sortData(first.data, 'name', first.sortColumn, first.sortDirection);
    assert.deepStrictEqual(second.data.map((d) => d.name), ['Б', 'А']);
});

test('sortData: числа sortData', () => {
    const data = [
        { val: 10 },
        { val: 2 },
        { val: 100 },
    ];
    const res = sortData(data, 'val', null, false);
    assert.deepStrictEqual(res.data.map((d) => d.val), [2, 10, 100]);
});

test('sortData: пустой массив', () => {
    const res = sortData([], 'name', null, false);
    assert.deepStrictEqual(res.data, []);
});

/* ---------- debounce ---------- */

test('debounce: вызывает функцию один раз после задержки', async () => {
    let calls = 0;
    const fn = debounce(() => {
        calls++;
    }, 50);
    fn();
    fn();
    fn();
    assert.strictEqual(calls, 0);
    await new Promise((r) => setTimeout(r, 100));
    assert.strictEqual(calls, 1);
});

/* ---------- escapeHtml ---------- */

test('escapeHtml: экранирует спецсимволы', () => {
    assert.strictEqual(escapeHtml('<script>alert("x")</script>'), '&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;');
});

test('escapeHtml: пустая строка', () => {
    assert.strictEqual(escapeHtml(''), '');
});

/* ---------- validateXLSXHeaders ---------- */

test('validateXLSXHeaders: все колонки найдены', () => {
    const json = [
        ['Дата', 'Время', 'Операция', 'Кол-во штук'],
        ['2024-01-01', '10:00', 'Приёмка', 10],
    ];
    assert.strictEqual(validateXLSXHeaders(json, ['Дата', 'Операция']), true);
});

test('validateXLSXHeaders: группы альтернатив', () => {
    const json = [['Кол-во доступное', 'Макс. Запас']];
    assert.strictEqual(
        validateXLSXHeaders(json, [['Кол-во физическое', 'Кол-во доступное'], 'Макс. Запас']),
        true
    );
});

test('validateXLSXHeaders: короткое имя «Код» принимается для паллет', () => {
    const json = [['Код', 'Отдел Код', 'Описание', 'Хранение (дней)', 'Ячейка', 'Посл. операция']];
    assert.strictEqual(
        validateXLSXHeaders(json, [
            ['Код паллеты', 'Код паллета', 'Код паллет', 'Код'],
            ['Отдел Код', 'Код отдела'],
            'Описание',
            ['Хранение (дней)', 'Срок хранения', 'Дней хранения'],
            'Ячейка',
            ['Посл. операция', 'Последняя операция']
        ]),
        true
    );
});

test('findColumnIndex: точное «Код» важнее частичного «Отдел Код»', () => {
    const headers = ['Отдел Код', 'Код', 'Описание'];
    assert.strictEqual(findColumnIndex(headers, ['Код паллеты', 'Код паллета', 'Код паллет', 'Код']), 1);
});

test('validateXLSXHeaders: сообщает об отсутствующих колонках', () => {
    const json = [['Дата', 'Операция']];
    assert.throws(
        () => validateXLSXHeaders(json, ['Дата', 'ФИО']),
        /ФИО/
    );
});

test('validateXLSXHeaders: учитывает индекс строки заголовков', () => {
    const json = [
        ['Отчёт за период', '', ''],
        ['', '', ''],
        ['Дата', 'Операция', 'Кол-во штук'],
        ['2024-01-01', 'Приёмка', 10],
    ];
    assert.strictEqual(validateXLSXHeaders(json, ['Дата', 'Операция'], 2), true);
    assert.throws(() => validateXLSXHeaders(json, ['Дата', 'Операция']), /Дата/);
});

test('validateXLSXHeaders: пустой файл и пустые заголовки', () => {
    assert.throws(() => validateXLSXHeaders([], ['Дата']), /не содержит данных/);
    assert.throws(() => validateXLSXHeaders([[]], ['Дата']), /не содержит заголовков/);
    assert.throws(() => validateXLSXHeaders(null, ['Дата']), /не содержит данных/);
});

/* ---------- exportRowsToXLSX ---------- */

test('exportRowsToXLSX: пишет лист с данными и датой в имени файла', () => {
    lastExport.fileName = null;
    lastExport.sheetName = null;
    lastExport.rows = null;
    exportRowsToXLSX([['А', 'Б'], [1, 2]], 'Тест', 'отчёт');
    assert.deepStrictEqual(lastExport.rows, [['А', 'Б'], [1, 2]]);
    assert.strictEqual(lastExport.sheetName, 'Тест');
    assert.match(lastExport.fileName, /^отчёт_\d{4}-\d{2}-\d{2}_\d{2}-\d{2}\.xlsx$/);
});

test('exportRowsToXLSX: пустые данные не пишутся, показывается предупреждение', () => {
    alerts.length = 0;
    lastExport.fileName = null;
    exportRowsToXLSX([], 'Тест', 'отчёт');
    assert.strictEqual(lastExport.fileName, null);
    assert.deepStrictEqual(alerts, ['Нет данных для экспорта']);
});

/* ---------- Пагинация ---------- */

test('getPageCount: считает страницы и обрабатывает пустые данные', () => {
    assert.strictEqual(getPageCount(0, 50), 0);
    assert.strictEqual(getPageCount(100, 50), 2);
    assert.strictEqual(getPageCount(101, 50), 3);
    assert.strictEqual(getPageCount(10, 0), 0);
});

test('getPageSlice: выдаёт нужную страницу и ограничивает номер', () => {
    const items = Array.from({ length: 120 }, (_, i) => i + 1);
    assert.deepStrictEqual(getPageSlice(items, 2, 50), items.slice(50, 100));
    assert.deepStrictEqual(getPageSlice(items, 99, 50), items.slice(100, 120));
    assert.deepStrictEqual(getPageSlice(items, 'не число', 50), items.slice(0, 50));
});

/* ---------- Хранилище и состояние фильтров ---------- */

/**
 * Загружает общие помощники в изолированное окружение с подставными DOM и localStorage.
 * @param {Object} elements - Элементы по ID.
 * @param {Object} localStorageStub - Подставное localStorage.
 * @returns {Object} Доступные функции.
 */
function loadCommonWithEnv(elements, localStorageStub) {
    const documentStub = {
        readyState: 'complete',
        getElementById: (id) => elements[id] || null,
        querySelectorAll: () => [],
        addEventListener: () => {},
    };
    return new Function(
        'window',
        'document',
        'XLSX',
        'Chart',
        'console',
        'alert',
        'localStorage',
        `${code}\nreturn { createReportStorage };`
    )(globalWindow, documentStub, globalXLSX, globalChart, console, globalAlert, localStorageStub);
}

test('createReportStorage: сохраняет и загружает обёрнутый формат', async () => {
    const elements = {};
    const db = new Map();
    let deleted = false;
    const idb = {
        get: async (key) => (db.has(key) ? db.get(key) : null),
        set: async (key, value) => { db.set(key, value); return true; },
        deleteDB: async () => { deleted = true; db.clear(); return true; },
    };
    const { createReportStorage } = loadCommonWithEnv(elements, {});

    const wrapped = createReportStorage({ idb, key: 'wrapped', wrapped: true });
    await wrapped.save([{ id: 2 }], 123);
    assert.deepStrictEqual(db.get('wrapped'), { records: [{ id: 2 }], timestamp: 123 });
    assert.deepStrictEqual(await wrapped.load(), { records: [{ id: 2 }], timestamp: 123 });

    const withMeta = createReportStorage({ idb, key: 'meta', wrapped: true });
    await withMeta.save([{ id: 3 }], 456, { zones: ['A', 'B'] });
    assert.deepStrictEqual(db.get('meta'), {
        records: [{ id: 3 }],
        timestamp: 456,
        meta: { zones: ['A', 'B'] },
    });
    assert.deepStrictEqual(await withMeta.load(), {
        records: [{ id: 3 }],
        timestamp: 456,
        meta: { zones: ['A', 'B'] },
    });

    await withMeta.save([{ id: 4 }], 789, {});
    assert.deepStrictEqual(db.get('meta'), { records: [{ id: 4 }], timestamp: 789 });

    await withMeta.clear();
    assert.strictEqual(deleted, true);
    assert.strictEqual(db.size, 0);
});
