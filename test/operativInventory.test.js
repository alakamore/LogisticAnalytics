import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const SERVICE_PATH = './Services/operativInventory/index.html';
const SAMPLE_PATH = './not indexed/samples_XLSX/OperativInventory.xlsx';
const serviceHtml = fs.readFileSync(SERVICE_PATH, 'utf8');
const commonCode = fs.readFileSync('./shared/js/common.js', 'utf8');
const xlsxCode = fs.readFileSync('./shared/js-lib/xlsx.full.min.js', 'utf8');

const scriptMatch = serviceHtml.match(/<script>([\s\S]*?)<\/script>/);
assert.ok(scriptMatch, 'в сервисе не найден встроенный скрипт');

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
    addEventListener: noop
};
const fakeWindow = {
    addEventListener: noop,
    removeEventListener: noop,
    dispatchEvent: () => true
};
const alerts = [];

const wrapper = new Function(
    'window',
    'document',
    'XLSX',
    'Chart',
    'console',
    'alert',
    'navigator',
    'setTimeout',
    'clearTimeout',
    `${commonCode}\n${scriptMatch[1]}\nreturn {
        parseOperativInventory,
        validateInventoryHeaders,
        filterInventoryStatements,
        sortInventoryStatements,
        computeInventoryKpi,
        buildInventoryDailySeries,
        buildInventoryEmployeeStats,
        buildInventoryDiscrepancyBreakdown,
        buildInventoryDurationBuckets,
        buildTopInventoryStatements,
        buildTopInventoryItems,
        buildInventoryItemStats,
        buildInventoryExportRows,
        getInventoryFilterDescription,
        getInventoryDateRangeKeys,
        formatInventoryDuration,
        formatInventoryQty,
        formatInventoryDelta,
        formatInventoryDate,
        formatInventoryTime,
        parseInventoryDateKey,
        pluralizeInventory,
        roundInventoryNumber
    };`
);

const service = wrapper(
    fakeWindow,
    fakeDocument,
    XLSX,
    function Chart() {},
    console,
    message => alerts.push(message),
    { clipboard: null },
    setTimeout,
    clearTimeout
);

const workbook = XLSX.read(fs.readFileSync(SAMPLE_PATH), { type: 'buffer', cellDates: false });
const jsonData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], {
    header: 1,
    defval: '',
    raw: true
});
const parsed = service.parseOperativInventory(jsonData);
const statements = parsed.statements;
const kpi = service.computeInventoryKpi(statements);

function round(value) {
    return value === null || value === undefined ? null : Math.round(value);
}

/* ---------- Заголовки ---------- */

test('operativInventory: обязательные колонки исходного отчёта распознаются', () => {
    const validation = service.validateInventoryHeaders(jsonData[0]);
    assert.strictEqual(validation.ok, true);
    assert.deepEqual(validation.missing, []);
    assert.strictEqual(validation.indexes.doc, 0);
    assert.strictEqual(validation.indexes.status, 17);
});

test('operativInventory: отсутствующая колонка возвращается в списке missing', () => {
    const headers = [...jsonData[0]];
    const deltaIndex = headers.indexOf('Дельта');
    headers.splice(deltaIndex, 1);
    const validation = service.validateInventoryHeaders(headers);
    assert.strictEqual(validation.ok, false);
    assert.ok(validation.missing.includes('Дельта'));
});

test('operativInventory: исправленное название «Тип инвентаризации» тоже принимается', () => {
    const headers = [...jsonData[0]];
    const index = headers.indexOf('Тип инвентризации');
    headers[index] = 'Тип инвентаризации';
    const validation = service.validateInventoryHeaders(headers);
    assert.strictEqual(validation.ok, true);
});

/* ---------- Разбор ---------- */

test('operativInventory: строки и ведомости разбираются без потерь', () => {
    assert.strictEqual(parsed.lines.length, 2345);
    assert.strictEqual(statements.length, 237);
});

test('operativInventory: статусы ведомостей - 237 «Учтено», незавершённых нет', () => {
    assert.strictEqual(kpi.completed, 237);
    assert.strictEqual(kpi.progress, 0);
});

test('operativInventory: уникальные ячейки, паллеты, товары и сотрудники', () => {
    assert.strictEqual(kpi.cells, 369);
    assert.strictEqual(kpi.pallets, 962);
    assert.strictEqual(kpi.items, 1575);
    assert.strictEqual(kpi.employees, 14);
});

test('operativInventory: внутри ведомости один статус, сотрудник, тип задания и тип инвентаризации', () => {
    const groups = new Map();
    for (const line of parsed.lines) {
        if (!groups.has(line.doc)) groups.set(line.doc, []);
        groups.get(line.doc).push(line);
    }

    for (const [doc, lines] of groups) {
        assert.strictEqual(new Set(lines.map(line => line.status)).size, 1, `статус в ${doc}`);
        assert.strictEqual(new Set(lines.map(line => line.fio)).size, 1, `сотрудник в ${doc}`);
        assert.strictEqual(new Set(lines.map(line => line.task)).size, 1, `тип задания в ${doc}`);
        assert.strictEqual(new Set(lines.map(line => line.invType)).size, 1, `тип инвентаризации в ${doc}`);
    }
});

/* ---------- KPI ---------- */

test('operativInventory: KPI по расхождениям сходится с исходным отчётом', () => {
    assert.strictEqual(kpi.totalStatements, 237);
    assert.strictEqual(kpi.mismatchStatements, 47);
    assert.strictEqual(kpi.mismatchLines, 136);
    assert.strictEqual(kpi.shortageLines, 63);
    assert.strictEqual(kpi.surplusLines, 73);
    assert.strictEqual(kpi.shortage, -1026.5);
    assert.strictEqual(kpi.surplus, 1306);
    assert.strictEqual(kpi.net, 279.5);
    assert.strictEqual(kpi.gross, 2332.5);
    assert.ok(Math.abs(kpi.before - 54603.135) < 0.000001);
    assert.ok(Math.abs(kpi.result - 54882.635) < 0.000001);
});

test('operativInventory: структура расхождений - 190 / 18 / 18 / 11 / 0', () => {
    assert.deepEqual(kpi.categories, {
        clean: 190,
        shortage: 18,
        surplus: 18,
        both: 11,
        progress: 0
    });
    assert.strictEqual(kpi.cleanCompleted, 190);
    assert.strictEqual(kpi.days, 18);
    assert.strictEqual(kpi.dateFrom, '2026-11-01');
    assert.strictEqual(kpi.dateTo, '2026-11-19');
});

test('operativInventory: медиана и среднее время считаются по 237 завершённым ведомостям', () => {
    assert.strictEqual(kpi.completedDurations, 237);
    assert.strictEqual(round(kpi.medianDuration), 19);
    assert.strictEqual(round(kpi.averageDuration), 42);
});

test('operativInventory: интервалы длительности - 97 / 68 / 38 / 34', () => {
    const buckets = service.buildInventoryDurationBuckets(statements);
    assert.deepEqual({
        lt15: buckets.lt15,
        from15to30: buckets.from15to30,
        from30to60: buckets.from30to60,
        over60: buckets.over60
    }, {
        lt15: 97,
        from15to30: 68,
        from30to60: 38,
        over60: 34
    });
    assert.strictEqual(buckets.total, 237);
    assert.strictEqual(round(buckets.median), 19);
    assert.strictEqual(round(buckets.average), 42);
});

/* ---------- Динамика ---------- */

test('operativInventory: динамика по дням строится по дате учёта и заполняет 11.11 нулём', () => {
    const daily = service.buildInventoryDailySeries(statements);
    assert.strictEqual(daily.length, 19);

    const first = daily.find(day => day.key === '2026-11-01');
    assert.deepEqual({
        statements: first.statements,
        clean: first.clean,
        mismatch: first.mismatch,
        progress: first.progress,
        rows: first.rows
    }, {
        statements: 43,
        clean: 39,
        mismatch: 4,
        progress: 0,
        rows: 245
    });

    const second = daily.find(day => day.key === '2026-11-02');
    assert.deepEqual({
        statements: second.statements,
        clean: second.clean,
        mismatch: second.mismatch,
        rows: second.rows
    }, {
        statements: 2,
        clean: 0,
        mismatch: 2,
        rows: 13
    });

    const gap = daily.find(day => day.key === '2026-11-11');
    assert.strictEqual(gap.statements, 0);
    assert.strictEqual(gap.rows, 0);

    const peak = daily.find(day => day.key === '2026-11-14');
    assert.deepEqual({
        statements: peak.statements,
        clean: peak.clean,
        mismatch: peak.mismatch,
        rows: peak.rows
    }, {
        statements: 20,
        clean: 10,
        mismatch: 10,
        rows: 974
    });

    const last = daily.find(day => day.key === '2026-11-15');
    assert.deepEqual({
        statements: last.statements,
        clean: last.clean,
        mismatch: last.mismatch,
        progress: last.progress,
        rows: last.rows
    }, {
        statements: 14,
        clean: 13,
        mismatch: 1,
        progress: 0,
        rows: 37
    });
});

test('operativInventory: сумма дней сходится с общим количеством ведомостей', () => {
    const daily = service.buildInventoryDailySeries(statements);
    const sum = key => daily.reduce((total, day) => total + day[key], 0);
    assert.strictEqual(sum('statements'), 237);
    assert.strictEqual(sum('clean'), 190);
    assert.strictEqual(sum('mismatch'), 47);
    assert.strictEqual(sum('progress'), 0);
});

/* ---------- ТОП-10 ---------- */

test('operativInventory: ТОП-10 ведомостей по недостачам возглавляет RB043', () => {
    const top = service.buildTopInventoryStatements(statements, 10, 'shortage');
    assert.strictEqual(top.length, 10);

    const first = top[0];
    assert.strictEqual(first.doc, 'I261114.165244.602');
    assert.strictEqual(first.fio, 'Игорь Брагарь');
    assert.strictEqual(first.gross, 1281);
    assert.strictEqual(first.net, 141);
    assert.strictEqual(first.shortage, -570);
    assert.strictEqual(first.surplus, 711);
    assert.strictEqual(first.mismatchLines.length, 32);
    assert.strictEqual(first.lines.length, 58);
    assert.strictEqual(first.cells.length, 1);

    const second = top[1];
    assert.strictEqual(second.doc, 'I261101.090333.95');
    assert.strictEqual(second.shortage, -84);
});

test('operativInventory: ТОП-10 ведомостей по излишкам сортируется по излишку', () => {
    const top = service.buildTopInventoryStatements(statements, 10, 'surplus');
    assert.strictEqual(top.length, 10);

    assert.strictEqual(top[0].doc, 'I261114.165244.602');
    assert.strictEqual(top[0].surplus, 711);

    assert.strictEqual(top[1].doc, 'I261116.140246.192');
    assert.strictEqual(top[1].surplus, 120);

    assert.ok(top.every(statement => statement.surplus > 0));
});

test('operativInventory: ТОП-10 артикулов по недостачам возглавляет 87103810', () => {
    const top = service.buildTopInventoryItems(statements, 10, 'shortage');
    assert.strictEqual(top.length, 10);

    const first = top[0];
    assert.strictEqual(first.item, '87103810');
    assert.strictEqual(first.shortage, -160);
    assert.strictEqual(first.surplus, 0);
    assert.strictEqual(first.gross, 160);
    assert.strictEqual(first.lines, 1);

    const second = top[1];
    assert.strictEqual(second.item, '89404145');
    assert.strictEqual(second.shortage, -104);
    assert.strictEqual(second.gross, 164);
    assert.strictEqual(second.lines, 3);
    assert.ok(top.every(item => item.shortage < 0));
});

test('operativInventory: ТОП-10 артикулов по излишкам сортируется по излишку', () => {
    const top = service.buildTopInventoryItems(statements, 10, 'surplus');
    assert.strictEqual(top.length, 10);

    assert.strictEqual(top[0].item, '84408504');
    assert.strictEqual(top[0].surplus, 120);
    assert.strictEqual(top[1].item, '87103792');
    assert.strictEqual(top[1].surplus, 120);

    const item = top.find(entry => entry.item === '89404145');
    assert.ok(item);
    assert.strictEqual(item.surplus, 60);
    assert.ok(top.every(entry => entry.surplus > 0));
});

test('operativInventory: карточка артикула 89404145 собирает строки из трёх ведомостей', () => {
    const item = service.buildInventoryItemStats(statements, '89404145');
    assert.strictEqual(item.gross, 164);
    assert.strictEqual(item.net, -44);
    assert.strictEqual(item.lines.length, 3);
    assert.strictEqual(item.statements, 3);
    assert.strictEqual(item.cells, 2);
    assert.strictEqual(item.fios, 2);
});

test('operativInventory: самая большая ведомость - 209 строк, 15 ячеек и 207 паллет', () => {
    const largest = statements.find(statement => statement.doc === 'I261114.091845.782');
    assert.ok(largest);
    assert.strictEqual(largest.lines.length, 209);
    assert.strictEqual(largest.cells.length, 15);
    assert.strictEqual(largest.pallets.length, 207);
});

/* ---------- Фильтры ---------- */

test('operativInventory: фильтры по статусу и расхождениям', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, { status: 'В процессе' }).length, 0);
    assert.strictEqual(service.filterInventoryStatements(statements, { status: 'Учтено' }).length, 237);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'yes' }).length, 47);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'no' }).length, 190);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'shortage' }).length, 18);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'surplus' }).length, 18);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'both' }).length, 11);
});

test('operativInventory: варианты фильтра расхождений взаимоисключающие и покрывают все ведомости', () => {
    const select = serviceHtml.match(/<select id="discrepancyFilter">([\s\S]*?)<\/select>/);
    assert.ok(select, 'не найден фильтр расхождений');
    const values = [...select[1].matchAll(/<option value="([^"]*)"/g)].map(match => match[1]);
    assert.deepEqual(values, ['', 'no', 'shortage', 'surplus', 'both']);

    const count = value => service.filterInventoryStatements(statements, { discrepancy: value }).length;
    assert.strictEqual(count('no') + count('shortage') + count('surplus') + count('both'), statements.length);
});

test('operativInventory: у незавершённых ведомостей расхождений нет', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, {
        status: 'В процессе',
        discrepancy: 'shortage'
    }).length, 0);
    assert.strictEqual(service.filterInventoryStatements(statements, {
        status: 'В процессе',
        discrepancy: 'no'
    }).length, 0);
});

test('operativInventory: фильтр по сотруднику и дате', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, { employee: 'Игорь Брагарь' }).length, 11);
    assert.strictEqual(service.filterInventoryStatements(statements, {
        dateFrom: '2026-11-14',
        dateTo: '2026-11-14'
    }).length, 20);
});

test('operativInventory: поиск работает по артикулу, ячейке и нескольким словам', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, { search: '87103810' }).length, 1);
    assert.strictEqual(service.filterInventoryStatements(statements, { search: 'RB043' }).length, 1);
    assert.strictEqual(service.filterInventoryStatements(statements, { search: 'Игорь RB043' }).length, 1);
});

/* ---------- Сортировка ---------- */

test('operativInventory: сортировка по строкам и дельте', () => {
    const byRows = service.sortInventoryStatements(statements, 'rows', 'desc');
    assert.strictEqual(byRows[0].doc, 'I261114.091845.782');
    assert.strictEqual(byRows[0].lines.length, 209);

    const byNet = service.sortInventoryStatements(statements, 'net', 'desc');
    assert.strictEqual(byNet[0].doc, 'I261114.165244.602');
    assert.strictEqual(byNet[0].net, 141);
});

/* ---------- Экспорт ---------- */

test('operativInventory: экспортные листы содержат все данные и учитывают фильтры', () => {
    const sheets = service.buildInventoryExportRows(statements);
    assert.strictEqual(sheets.statements.length, 238);
    assert.strictEqual(sheets.lines.length, 2346);
    assert.strictEqual(sheets.employees.length, 15);
    assert.ok(sheets.summary.length > 50);

    assert.ok(sheets.statements[0].includes('№ Ведомости'));
    assert.ok(sheets.statements[0].includes('Оборот, шт.'));
    assert.ok(sheets.lines[0].includes('Дельта, шт.'));
    assert.ok(sheets.lines[0].includes('Статус'));
    assert.ok(sheets.employees[0].includes('Медиана, мин'));

    const grossRow = sheets.summary.find(row => row[0] === 'Оборот расхождений, шт.');
    assert.ok(grossRow);
    assert.strictEqual(grossRow[1], 2332.5);

    const deltaRow = sheets.summary.find(row => row[0] === 'Дельта, шт.');
    assert.ok(deltaRow);
    assert.strictEqual(deltaRow[1], 279.5);

    const titles = sheets.summary.map(row => row[0]);
    assert.ok(titles.includes('ТОП-10 ведомостей по недостачам'));
    assert.ok(titles.includes('ТОП-10 ведомостей по излишкам'));
    assert.ok(titles.includes('ТОП-10 артикулов по недостачам'));
    assert.ok(titles.includes('ТОП-10 артикулов по излишкам'));

    assert.ok(sheets.statements[0].includes('Дельта, шт.'));
    const dailyHeader = sheets.summary.find(row => row[0] === 'Дата');
    assert.ok(dailyHeader.includes('Дельта, шт.'));
});

test('operativInventory: описание фильтров попадает в сводку', () => {
    const description = service.getInventoryFilterDescription(
        { employee: 'Игорь Брагарь', status: 'Учтено' },
        { min: '2026-11-01', max: '2026-11-19' }
    );
    assert.ok(description.includes('сотрудник: Игорь Брагарь'));
    assert.ok(description.includes('статус: Учтено'));
    assert.strictEqual(
        service.getInventoryFilterDescription({}, { min: '2026-11-01', max: '2026-11-19' }),
        'Все данные'
    );
});

/* ---------- Форматирование ---------- */

test('operativInventory: длительность форматируется в минуты и часы', () => {
    assert.strictEqual(service.formatInventoryDuration(21.533), '22 мин');
    assert.strictEqual(service.formatInventoryDuration(65), '1 ч 05 мин');
    assert.strictEqual(service.formatInventoryDuration(1398.916), '23 ч 19 мин');
    assert.strictEqual(service.formatInventoryDuration(null), '-');
});

test('operativInventory: дельта форматируется со знаком и типографским минусом', () => {
    assert.strictEqual(service.formatInventoryDelta(-926), '−926');
    assert.strictEqual(service.formatInventoryDelta(141), '+141');
    assert.strictEqual(service.formatInventoryDelta(0), '0');
});

test('operativInventory: дата и время форматируются по данным WMS', () => {
    assert.strictEqual(service.formatInventoryDate('01.11.26'), '01.11.26');
    assert.strictEqual(service.formatInventoryDate('01.11.2026'), '01.11.26');
    assert.strictEqual(service.formatInventoryTime('9:56:31'), '09:56');
    assert.strictEqual(service.formatInventoryTime('9:56:31', true), '09:56:31');
    assert.strictEqual(service.parseInventoryDateKey('01.11.26'), '2026-11-01');
});

test('operativInventory: количества форматируются с пробелами и запятой', () => {
    const formatted = service.formatInventoryQty(54603.135);
    assert.ok(formatted.includes('54'));
    assert.ok(formatted.includes('603,135'));
});

test('operativInventory: диапазон дат включает пропущенный день', () => {
    const keys = service.getInventoryDateRangeKeys('2026-11-01', '2026-11-19');
    assert.strictEqual(keys.length, 19);
    assert.ok(keys.includes('2026-11-11'));
});

test('operativInventory: экспорт округляет агрегаты без хвостов плавающей точки', () => {
    const sheets = service.buildInventoryExportRows(statements);
    const beforeRow = sheets.summary.find(row => row[0] === 'На начало подсчёта, шт.');
    const resultRow = sheets.summary.find(row => row[0] === 'Результат подсчёта, шт.');
    assert.strictEqual(beforeRow[1], 54603.135);
    assert.strictEqual(resultRow[1], 54882.635);
});

test('operativInventory: числительные склоняются по правилам русского языка', () => {
    const forms = ['строка', 'строки', 'строк'];
    assert.strictEqual(service.pluralizeInventory(1, forms), 'строка');
    assert.strictEqual(service.pluralizeInventory(2, forms), 'строки');
    assert.strictEqual(service.pluralizeInventory(5, forms), 'строк');
    assert.strictEqual(service.pluralizeInventory(11, forms), 'строк');
    assert.strictEqual(service.pluralizeInventory(21, forms), 'строка');
    assert.strictEqual(service.pluralizeInventory(122, forms), 'строки');
    assert.strictEqual(service.pluralizeInventory(0, forms), 'строк');
});

/* ---------- Ошибки ---------- */

test('operativInventory: пустой файл и отсутствие колонок дают понятную ошибку', () => {
    assert.throws(() => service.parseOperativInventory([]), /не содержит данных/i);

    const headers = [...jsonData[0]];
    headers.splice(headers.indexOf('Дельта'), 1);
    assert.throws(
        () => service.parseOperativInventory([headers]),
        /обязательные колонки/i
    );
});
