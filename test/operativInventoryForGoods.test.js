import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const SERVICE_PATH = './Services/operativInventoryForGoods/index.html';
const SAMPLE_PATH = './not indexed/samples_XLSX/operativInventoryForGoods.xlsx';
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
        buildInventoryZoneStats,
        buildInventoryProductStats,
        buildInventoryDiscrepancyBreakdown,
        buildTopInventoryStatements,
        buildTopInventoryItems,
        buildInventoryItemStats,
        buildInventoryExportRows,
        getInventoryFilterDescription,
        getInventoryDateRangeKeys,
        formatInventoryDuration,
        formatInventoryQty,
        formatInventoryDelta,
        formatInventoryMoney,
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

test('operativInventoryForGoods: обязательные колонки исходного отчёта распознаются', () => {
    const validation = service.validateInventoryHeaders(jsonData[0]);
    assert.strictEqual(validation.ok, true);
    assert.deepEqual(validation.missing, []);
    assert.strictEqual(validation.indexes.doc, 0);
    assert.strictEqual(validation.indexes.item, 1);
    assert.strictEqual(validation.indexes.deltaQty, 12);
    assert.strictEqual(validation.indexes.deltaSum, 15);
    assert.strictEqual(validation.indexes.activatedBy, 27);
});

test('operativInventoryForGoods: отсутствующая колонка возвращается в списке missing', () => {
    const headers = [...jsonData[0]];
    headers.splice(headers.indexOf('Количество Дельта'), 1);
    const validation = service.validateInventoryHeaders(headers);
    assert.strictEqual(validation.ok, false);
    assert.ok(validation.missing.includes('Количество Дельта'));
});

test('operativInventoryForGoods: необязательные колонки не ломают разбор', () => {
    const headers = [...jsonData[0]];
    headers.splice(headers.indexOf('Владелец Товара'), 1);
    headers.splice(headers.indexOf('Код завершения'), 1);
    headers.splice(headers.indexOf('Номер документа коррекции'), 1);
    const validation = service.validateInventoryHeaders(headers);
    assert.strictEqual(validation.ok, true);
    assert.deepEqual(validation.missing, []);
});

/* ---------- Разбор ---------- */

test('operativInventoryForGoods: строки и ведомости разбираются без потерь', () => {
    assert.strictEqual(parsed.lines.length, 1113);
    assert.strictEqual(statements.length, 27);
    assert.deepEqual(parsed.meta, { rowCount: 1113, statementCount: 27 });
});

test('operativInventoryForGoods: уникальные товары, ячейки, паллеты, задания и сотрудники', () => {
    assert.strictEqual(kpi.items, 276);
    assert.strictEqual(kpi.cells, 252);
    assert.strictEqual(kpi.pallets, 752);
    assert.strictEqual(kpi.tasks, 901);
    assert.strictEqual(kpi.zones, 5);
    assert.strictEqual(kpi.employees, 14);
});

test('operativInventoryForGoods: все ведомости завершены, незавершённых нет', () => {
    assert.strictEqual(kpi.completed, 27);
    assert.strictEqual(kpi.progress, 0);
    assert.ok(statements.every(statement => statement.isCompleted));
    assert.ok(statements.every(statement => statement.completeAt > 0));
});

test('operativInventoryForGoods: первая строка разобрана в поля товарной строки', () => {
    const line = parsed.lines[0];
    assert.strictEqual(line.doc, '1092');
    assert.strictEqual(line.item, '84133815');
    assert.strictEqual(line.task, 'ED383 ->');
    assert.strictEqual(line.zone, 'EM');
    assert.strictEqual(line.cell, 'ED383');
    assert.strictEqual(line.pallet, '');
    assert.strictEqual(line.uom, 'ШТ');
    assert.strictEqual(line.calcQty, 32);
    assert.strictEqual(line.factQty, 32);
    assert.strictEqual(line.deltaQty, 0);
    assert.strictEqual(line.fio, 'Данил Рогожников');
    assert.strictEqual(line.ldap, '60031507');
    assert.strictEqual(line.activatedBy, 'Сергей Лосев');
    assert.ok(line.desc.startsWith('ПЛ. ОБЛ.НЕВАДА'));
    assert.strictEqual(line.effectiveDateKey, '2026-11-01');
});

test('operativInventoryForGoods: суммы и дельта суммы разбираются из отчёта', () => {
    const row = parsed.lines.find(line => line.doc === '1101' && line.deltaQty !== 0);
    assert.ok(row);
    assert.strictEqual(row.calcQty, 254);
    assert.strictEqual(row.factQty, 15);
    assert.strictEqual(row.deltaQty, -239);
    assert.strictEqual(row.calcSum, 54610);
    assert.strictEqual(row.factSum, 3225);
    assert.strictEqual(row.deltaSum, -51385);
    assert.strictEqual(row.correction, '05112026.00124');
    assert.strictEqual(row.secondQty, 15);
});

test('operativInventoryForGoods: коды завершения и коррекции считаются по строкам', () => {
    const codes = parsed.lines.filter(line => line.code !== '');
    assert.strictEqual(codes.length, 5);
    assert.ok(codes.every(line => line.code === 'SSCC НЕТ В АДРЕСЕ'));

    const corrections = parsed.lines.filter(line => line.correction !== '');
    assert.strictEqual(corrections.length, 59);
    assert.strictEqual(kpi.corrections, 59);
});

/* ---------- KPI ---------- */

test('operativInventoryForGoods: KPI по количеству сходится с исходным отчётом', () => {
    assert.strictEqual(kpi.totalStatements, 27);
    assert.strictEqual(kpi.mismatchStatements, 16);
    assert.strictEqual(kpi.mismatchLines, 64);
    assert.strictEqual(kpi.shortageLines, 42);
    assert.strictEqual(kpi.surplusLines, 22);
    assert.strictEqual(kpi.shortageQty, -942);
    assert.strictEqual(kpi.surplusQty, 688);
    assert.strictEqual(kpi.deltaQty, -254);
    assert.strictEqual(kpi.grossQty, 1630);
    assert.strictEqual(kpi.calcQty, 107803);
    assert.strictEqual(kpi.factQty, 107549);
});

test('operativInventoryForGoods: KPI по суммам сходится с исходным отчётом', () => {
    assert.ok(Math.abs(kpi.calcSum - 25682023.94) < 0.001);
    assert.ok(Math.abs(kpi.factSum - 25563997.53) < 0.001);
    assert.ok(Math.abs(kpi.deltaSum - (-118026.41)) < 0.001);
    assert.ok(Math.abs(kpi.shortageSum - (-266321.64)) < 0.001);
    assert.ok(Math.abs(kpi.surplusSum - 148295.23) < 0.001);
    assert.ok(Math.abs(kpi.grossSum - 414616.87) < 0.001);
});

test('operativInventoryForGoods: структура расхождений - 11 / 8 / 0 / 8 / 0', () => {
    assert.deepEqual(kpi.categories, {
        clean: 11,
        shortage: 8,
        surplus: 0,
        both: 8,
        progress: 0
    });
    assert.strictEqual(kpi.cleanCompleted, 11);
});

test('operativInventoryForGoods: период и длительность считаются по завершённым ведомостям', () => {
    assert.strictEqual(kpi.days, 17);
    assert.strictEqual(kpi.dateFrom, '2026-11-01');
    assert.strictEqual(kpi.dateTo, '2026-11-19');
    assert.strictEqual(kpi.completedDurations, 27);
    assert.strictEqual(round(kpi.medianDuration), 6);
    assert.strictEqual(round(kpi.averageDuration), 50);
});

test('operativInventoryForGoods: коррекции и вторые подсчёты посчитаны по строкам', () => {
    assert.strictEqual(kpi.secondRows, 63);
    assert.strictEqual(kpi.codeRows, 5);
});

/* ---------- Динамика ---------- */

test('operativInventoryForGoods: динамика по дням заполняет 06.11 и 15.11 нулями', () => {
    const daily = service.buildInventoryDailySeries(statements);
    assert.strictEqual(daily.length, 19);

    const first = daily.find(day => day.key === '2026-11-01');
    assert.deepEqual({
        statements: first.statements,
        clean: first.clean,
        mismatch: first.mismatch,
        rows: first.rows,
        goods: first.goods,
        deltaQty: first.deltaQty
    }, {
        statements: 1,
        clean: 0,
        mismatch: 1,
        rows: 26,
        goods: 11,
        deltaQty: 4
    });

    const gap = daily.find(day => day.key === '2026-11-06');
    assert.strictEqual(gap.statements, 0);
    assert.strictEqual(gap.rows, 0);

    const peak = daily.find(day => day.key === '2026-11-11');
    assert.deepEqual({
        statements: peak.statements,
        rows: peak.rows,
        goods: peak.goods
    }, {
        statements: 2,
        rows: 662,
        goods: 132
    });
});

test('operativInventoryForGoods: сумма дней сходится с общим количеством ведомостей и строк', () => {
    const daily = service.buildInventoryDailySeries(statements);
    const sum = key => daily.reduce((total, day) => total + day[key], 0);
    assert.strictEqual(sum('statements'), 27);
    assert.strictEqual(sum('clean'), 11);
    assert.strictEqual(sum('mismatch'), 16);
    assert.strictEqual(sum('progress'), 0);
    assert.strictEqual(sum('rows'), 1113);
    assert.strictEqual(sum('deltaQty'), -254);
});

/* ---------- Сотрудники и зоны ---------- */

test('operativInventoryForGoods: статистика сотрудников считается по строкам', () => {
    const employees = service.buildInventoryEmployeeStats(statements);
    assert.strictEqual(employees.length, 14);

    const top = employees[0];
    assert.strictEqual(top.fio, 'Богдан Леснов');
    assert.strictEqual(top.rows, 555);
    assert.strictEqual(top.statements, 14);
    assert.strictEqual(top.goods, 160);
    assert.strictEqual(top.cells, 163);
    assert.strictEqual(top.pallets, 440);
    assert.strictEqual(top.cleanRows, 550);
    assert.strictEqual(top.mismatchRows, 5);
    assert.strictEqual(top.deltaQty, -69);
    assert.ok(Math.abs(top.deltaSum - (-12376.12)) < 0.001);

    assert.strictEqual(employees[1].fio, 'Андрей Бокарев');
    assert.strictEqual(employees[1].rows, 268);
});

test('operativInventoryForGoods: статистика зон отделяет недостачу и излишек в рублях', () => {
    const zones = service.buildInventoryZoneStats(statements);
    assert.deepEqual(zones.map(zone => zone.zone), ['EM', 'P', 'RM', 'U', 'Y']);

    const em = zones.find(zone => zone.zone === 'EM');
    assert.strictEqual(em.statements, 14);
    assert.strictEqual(em.rows, 172);
    assert.strictEqual(em.goods, 80);
    assert.strictEqual(em.deltaQty, -69);
    assert.ok(Math.abs(em.deltaSum - (-8921)) < 0.001);
    assert.ok(Math.abs(em.shortageSum - (-45934.47)) < 0.001);
    assert.ok(Math.abs(em.surplusSum - 37013.47) < 0.001);

    const empty = zones.find(zone => zone.zone === 'P');
    assert.strictEqual(empty.rows, 8);
    assert.strictEqual(empty.deltaSum, 0);
});

test('operativInventoryForGoods: товарная статистика содержит все товары', () => {
    const products = service.buildInventoryProductStats(statements);
    assert.strictEqual(products.length, 276);

    const first = products[0];
    assert.strictEqual(first.item, '10072681');
    assert.strictEqual(first.rows, 3);
    assert.strictEqual(first.statements, 1);
    assert.strictEqual(first.calcQty, 168);
    assert.strictEqual(first.deltaQty, 0);

    assert.strictEqual(products[1].item, '10072745');
    assert.strictEqual(products[1].rows, 8);
    assert.strictEqual(products[products.length - 1].item, '89475714');
    assert.ok(products.every(item => item.rows > 0));
});

/* ---------- ТОП-10 ---------- */

test('operativInventoryForGoods: ТОП-10 ведомостей по недостачам возглавляет 1101', () => {
    const top = service.buildTopInventoryStatements(statements, 10, 'shortage');
    assert.strictEqual(top.length, 10);

    const first = top[0];
    assert.strictEqual(first.doc, '1101');
    assert.strictEqual(first.shortageQty, -239);
    assert.strictEqual(first.surplusQty, 0);
    assert.strictEqual(first.grossQty, 239);
    assert.strictEqual(first.mismatchLines.length, 1);
    assert.strictEqual(first.lines.length, 2);
    assert.strictEqual(first.cells.length, 2);
    assert.deepEqual(first.fios, ['Ян Ашрафов']);

    const second = top[1];
    assert.strictEqual(second.doc, '1100');
    assert.strictEqual(second.shortageQty, -238);
    assert.strictEqual(second.surplusQty, 591);
    assert.ok(top.every(statement => statement.shortageQty < 0));
});

test('operativInventoryForGoods: ТОП-10 ведомостей по излишкам возглавляет 1100', () => {
    const top = service.buildTopInventoryStatements(statements, 10, 'surplus');
    assert.strictEqual(top.length, 8);

    assert.strictEqual(top[0].doc, '1100');
    assert.strictEqual(top[0].surplusQty, 591);
    assert.ok(Math.abs(top[0].deltaSum - 54164.82) < 0.001);

    assert.strictEqual(top[1].doc, '1118');
    assert.strictEqual(top[1].surplusQty, 54);
    assert.ok(top.every(statement => statement.surplusQty > 0));
});

test('operativInventoryForGoods: ТОП-10 товаров по недостачам возглавляет 82308013', () => {
    const top = service.buildTopInventoryItems(statements, 10, 'shortage');
    assert.strictEqual(top.length, 10);

    const first = top[0];
    assert.strictEqual(first.item, '82308013');
    assert.strictEqual(first.shortageQty, -249);
    assert.strictEqual(first.surplusQty, 232);
    assert.strictEqual(first.grossQty, 481);
    assert.strictEqual(first.lines, 3);
    assert.strictEqual(first.statements, 2);
    assert.strictEqual(first.cells, 2);

    assert.strictEqual(top[1].item, '18527151');
    assert.strictEqual(top[1].shortageQty, -156);
    assert.ok(top.every(item => item.shortageQty < 0));
});

test('operativInventoryForGoods: ТОП-10 товаров по излишкам сортируется по излишку', () => {
    const top = service.buildTopInventoryItems(statements, 10, 'surplus');
    assert.strictEqual(top.length, 10);

    assert.strictEqual(top[0].item, '82308013');
    assert.strictEqual(top[0].surplusQty, 232);

    assert.strictEqual(top[1].item, '82308012');
    assert.strictEqual(top[1].surplusQty, 93);
    assert.ok(top.every(item => item.surplusQty > 0));
});

test('operativInventoryForGoods: карточка товара 82308013 собирает строки из двух ведомостей', () => {
    const item = service.buildInventoryItemStats(statements, '82308013');
    assert.strictEqual(item.grossQty, 481);
    assert.strictEqual(item.deltaQty, -17);
    assert.strictEqual(item.shortageQty, -249);
    assert.strictEqual(item.surplusQty, 232);
    assert.strictEqual(item.lines.length, 3);
    assert.strictEqual(item.statements, 2);
    assert.strictEqual(item.cells, 2);
    assert.strictEqual(item.pallets, 0);
    assert.strictEqual(item.zones, 1);
    assert.strictEqual(item.fios, 1);
    assert.strictEqual(item.secondRows, 2);
    assert.ok(Math.abs(item.deltaSum - (-3655)) < 0.001);
});

test('operativInventoryForGoods: карточка товара без расхождений пуста', () => {
    const item = service.buildInventoryItemStats(statements, '84133815');
    assert.strictEqual(item.lines.length, 0);
    assert.strictEqual(item.grossQty, 0);
    assert.strictEqual(item.deltaSum, 0);
});

test('operativInventoryForGoods: самая большая ведомость - 1111, 659 строк и 640 паллет', () => {
    const largest = statements.find(statement => statement.doc === '1111');
    assert.ok(largest);
    assert.strictEqual(largest.lines.length, 659);
    assert.strictEqual(largest.items.length, 129);
    assert.strictEqual(largest.cells.length, 62);
    assert.strictEqual(largest.pallets.length, 640);
    assert.strictEqual(largest.category, 'clean');
});

/* ---------- Фильтры ---------- */

test('operativInventoryForGoods: фильтры по расхождениям', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'yes' }).length, 16);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'no' }).length, 11);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'shortage' }).length, 8);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'surplus' }).length, 0);
    assert.strictEqual(service.filterInventoryStatements(statements, { discrepancy: 'both' }).length, 8);
});

test('operativInventoryForGoods: варианты фильтра расхождений взаимоисключающие и покрывают все ведомости', () => {
    const select = serviceHtml.match(/<select id="discrepancyFilter">([\s\S]*?)<\/select>/);
    assert.ok(select, 'не найден фильтр расхождений');
    const values = [...select[1].matchAll(/<option value="([^"]*)"/g)].map(match => match[1]);
    assert.deepEqual(values, ['', 'no', 'shortage', 'surplus', 'both']);

    const count = value => service.filterInventoryStatements(statements, { discrepancy: value }).length;
    assert.strictEqual(count('no') + count('shortage') + count('surplus') + count('both'), statements.length);
});

test('operativInventoryForGoods: фильтр по коррекциям', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, { correction: 'yes' }).length, 14);
    assert.strictEqual(service.filterInventoryStatements(statements, { correction: 'no' }).length, 13);
});

test('operativInventoryForGoods: фильтр по сотруднику и зоне', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, { employee: 'Богдан Леснов' }).length, 14);
    assert.strictEqual(service.filterInventoryStatements(statements, { employee: 'Макар Макаров' }).length, 1);
    assert.strictEqual(service.filterInventoryStatements(statements, { zone: 'RM' }).length, 13);
    assert.strictEqual(service.filterInventoryStatements(statements, { zone: 'P' }).length, 4);
});

test('operativInventoryForGoods: фильтр по дате', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, {
        dateFrom: '2026-11-01',
        dateTo: '2026-11-01'
    }).length, 1);
    assert.strictEqual(service.filterInventoryStatements(statements, {
        dateFrom: '2026-11-05',
        dateTo: '2026-11-05'
    }).length, 3);
});

test('operativInventoryForGoods: поиск работает по ведомости, товару, ячейке и сотруднику', () => {
    assert.strictEqual(service.filterInventoryStatements(statements, { search: '84133815' }).length, 1);
    assert.strictEqual(service.filterInventoryStatements(statements, { search: 'ED383' }).length, 1);
    assert.strictEqual(service.filterInventoryStatements(statements, { search: 'Данил 1092' }).length, 1);
});

/* ---------- Сортировка ---------- */

test('operativInventoryForGoods: сортировка по строкам, товарам и дельте суммы', () => {
    const byRows = service.sortInventoryStatements(statements, 'rows', 'desc');
    assert.strictEqual(byRows[0].doc, '1111');
    assert.strictEqual(byRows[0].lines.length, 659);

    const byGoods = service.sortInventoryStatements(statements, 'goods', 'desc');
    assert.strictEqual(byGoods[0].doc, '1111');
    assert.strictEqual(byGoods[0].items.length, 129);

    const byDeltaSum = service.sortInventoryStatements(statements, 'deltaSum', 'asc');
    assert.strictEqual(byDeltaSum[0].doc, '1119');
    assert.ok(Math.abs(byDeltaSum[0].deltaSum - (-65168.48)) < 0.001);
});

/* ---------- Экспорт ---------- */

test('operativInventoryForGoods: экспортные листы содержат все данные', () => {
    const sheets = service.buildInventoryExportRows(statements);
    assert.strictEqual(sheets.statements.length, 28);
    assert.strictEqual(sheets.lines.length, 1114);
    assert.strictEqual(sheets.products.length, 277);
    assert.strictEqual(sheets.employees.length, 15);
    assert.ok(sheets.summary.length > 100);
});

test('operativInventoryForGoods: заголовки экспортных листов описывают товарный контекст', () => {
    const sheets = service.buildInventoryExportRows(statements);

    assert.ok(sheets.statements[0].includes('№ Ведомости'));
    assert.ok(sheets.statements[0].includes('Товаров'));
    assert.ok(sheets.statements[0].includes('Дельта суммы, ₽'));
    assert.ok(sheets.statements[0].includes('Строк с коррекцией'));

    assert.ok(sheets.lines[0].includes('Товар'));
    assert.ok(sheets.lines[0].includes('Номер коррекции'));
    assert.ok(sheets.lines[0].includes('Дельта суммы, ₽'));

    assert.ok(sheets.products[0].includes('Оборот, шт.'));
    assert.ok(sheets.products[0].includes('Дельта суммы, ₽'));

    assert.ok(sheets.employees[0].includes('Строк со 2-м подсчётом'));
});

test('operativInventoryForGoods: сводка содержит ключевые показатели и ТОП-10', () => {
    const sheets = service.buildInventoryExportRows(statements);
    const find = name => sheets.summary.find(row => row[0] === name);

    assert.strictEqual(find('Ведомостей')[1], 27);
    assert.strictEqual(find('Товаров')[1], 276);
    assert.strictEqual(find('Ячеек')[1], 252);
    assert.strictEqual(find('Паллет')[1], 752);
    assert.strictEqual(find('Заданий')[1], 901);
    assert.strictEqual(find('Строк подсчёта')[1], 1113);
    assert.strictEqual(find('Ведомостей с расхождениями')[1], 16);
    assert.strictEqual(find('Строк с расхождениями')[1], 64);
    assert.strictEqual(find('Строк с коррекцией')[1], 59);
    assert.strictEqual(find('Строк со 2-м подсчётом')[1], 63);
    assert.strictEqual(find('Строк с кодом завершения')[1], 5);
    assert.strictEqual(find('Расчётное кол-во, шт.')[1], 107803);
    assert.strictEqual(find('Фактическое кол-во, шт.')[1], 107549);
    assert.strictEqual(find('Дельта кол-ва, шт.')[1], -254);
    assert.strictEqual(find('Недостача, шт.')[1], -942);
    assert.strictEqual(find('Излишек, шт.')[1], 688);
    assert.strictEqual(find('Оборот кол-ва, шт.')[1], 1630);
    assert.ok(Math.abs(find('Расчётная сумма, ₽')[1] - 25682023.94) < 0.001);
    assert.ok(Math.abs(find('Дельта суммы, ₽')[1] - (-118026.41)) < 0.001);
    assert.ok(Math.abs(find('Оборот суммы, ₽')[1] - 414616.87) < 0.001);
    assert.strictEqual(find('Медиана длительности, мин')[1], 6);
    assert.strictEqual(find('Среднее длительности, мин')[1], 50);

    const titles = sheets.summary.map(row => row[0]);
    assert.ok(titles.includes('ТОП-10 ведомостей по недостачам'));
    assert.ok(titles.includes('ТОП-10 ведомостей по излишкам'));
    assert.ok(titles.includes('ТОП-10 товаров по недостачам'));
    assert.ok(titles.includes('ТОП-10 товаров по излишкам'));
    assert.ok(titles.includes('Зоны'));

    const dailyHeader = sheets.summary.find(row => row[0] === 'Дата');
    assert.ok(dailyHeader.includes('Дельта суммы, ₽'));
});

test('operativInventoryForGoods: экспорт учитывает активные фильтры', () => {
    const filtered = service.filterInventoryStatements(statements, { employee: 'Богдан Леснов' });
    assert.strictEqual(filtered.length, 14);

    const sheets = service.buildInventoryExportRows(filtered, { filtersDescription: 'сотрудник: Богдан Леснов' });
    assert.strictEqual(sheets.statements.length, 15);
    const filterRow = sheets.summary.find(row => row[0] === 'Фильтры');
    assert.strictEqual(filterRow[1], 'сотрудник: Богдан Леснов');
});

test('operativInventoryForGoods: описание фильтров попадает в сводку', () => {
    const description = service.getInventoryFilterDescription(
        { employee: 'Богдан Леснов', zone: 'U', correction: 'yes' },
        { min: '2026-11-01', max: '2026-11-19' }
    );
    assert.ok(description.includes('сотрудник: Богдан Леснов'));
    assert.ok(description.includes('зона: U'));
    assert.ok(description.includes('коррекции: есть'));
    assert.strictEqual(
        service.getInventoryFilterDescription({}, { min: '2026-11-01', max: '2026-11-19' }),
        'Все данные'
    );
});

/* ---------- Форматирование ---------- */

test('operativInventoryForGoods: длительность форматируется в минуты и часы', () => {
    assert.strictEqual(service.formatInventoryDuration(9), '9 мин');
    assert.strictEqual(service.formatInventoryDuration(65), '1 ч 05 мин');
    assert.strictEqual(service.formatInventoryDuration(1398.916), '23 ч 19 мин');
    assert.strictEqual(service.formatInventoryDuration(null), '-');
});

test('operativInventoryForGoods: дельта форматируется со знаком и типографским минусом', () => {
    assert.strictEqual(service.formatInventoryDelta(-700), '−700');
    assert.strictEqual(service.formatInventoryDelta(623), '+623');
    assert.strictEqual(service.formatInventoryDelta(0), '0');
});

test('operativInventoryForGoods: деньги форматируются в рублях с копейками', () => {
    const formatted = service.formatInventoryMoney(25682023.94);
    assert.ok(formatted.startsWith('25'));
    assert.ok(formatted.includes(',94'));
    assert.ok(formatted.includes('₽'));

    const negative = service.formatInventoryMoney(-118026.41, { sign: true });
    assert.ok(negative.startsWith('−'));
    assert.ok(negative.includes('118'));
    assert.ok(negative.includes('₽'));
});

test('operativInventoryForGoods: количества форматируются с пробелами', () => {
    const formatted = service.formatInventoryQty(107803);
    assert.ok(formatted.includes('107'));
    assert.ok(formatted.includes('803'));
});

test('operativInventoryForGoods: диапазон дат включает пропущенные дни', () => {
    const keys = service.getInventoryDateRangeKeys('2026-11-01', '2026-11-19');
    assert.strictEqual(keys.length, 19);
    assert.ok(keys.includes('2026-11-06'));
    assert.ok(keys.includes('2026-11-15'));
});

test('operativInventoryForGoods: числительные склоняются по правилам русского языка', () => {
    const forms = ['строка', 'строки', 'строк'];
    assert.strictEqual(service.pluralizeInventory(1, forms), 'строка');
    assert.strictEqual(service.pluralizeInventory(2, forms), 'строки');
    assert.strictEqual(service.pluralizeInventory(5, forms), 'строк');
    assert.strictEqual(service.pluralizeInventory(11, forms), 'строк');
    assert.strictEqual(service.pluralizeInventory(21, forms), 'строка');
    assert.strictEqual(service.pluralizeInventory(1113, forms), 'строк');
    assert.strictEqual(service.pluralizeInventory(2, ['товар', 'товара', 'товаров']), 'товара');
});

/* ---------- Ошибки ---------- */

test('operativInventoryForGoods: пустой файл и отсутствие колонок дают понятную ошибку', () => {
    assert.throws(() => service.parseOperativInventory([]), /не содержит данных/i);
    assert.throws(() => service.parseOperativInventory(null), /не содержит данных/i);

    const headers = [...jsonData[0]];
    headers.splice(headers.indexOf('Количество Дельта'), 1);
    assert.throws(
        () => service.parseOperativInventory([headers]),
        /обязательные колонки/i
    );
});
