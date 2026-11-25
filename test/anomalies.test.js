import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const SERVICE_PATH = './Services/anomalies/index.html';
const serviceHtml = fs.readFileSync(SERVICE_PATH, 'utf8');
const commonCode = fs.readFileSync('./shared/js/common.js', 'utf8');

const scriptMatch = serviceHtml.match(/<script>([\s\S]*?)<\/script>/);
assert.ok(scriptMatch, 'в сервисе не найден встроенный скрипт');

const noop = () => {};
const fakeDocument = {
    readyState: 'loading',
    getElementById: () => null,
    querySelectorAll: () => [],
    addEventListener: noop,
};
const fakeWindow = { addEventListener: noop, removeEventListener: noop, dispatchEvent: () => true };
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
        parseNum,
        parseAnomalyReport,
        parseWarehouseReport,
        buildIndexes,
        getWarehouseFor,
        anomalyWarehouseStats,
        formatExcelDate,
        get anomalyRows() { return anomalyRows; },
        set anomalyRows(value) { anomalyRows = value; },
        get warehouseRows() { return warehouseRows; },
        set warehouseRows(value) { warehouseRows = value; }
    };`
);

const service = wrapper(
    fakeWindow,
    fakeDocument,
    { read: () => ({}), utils: { sheet_to_json: () => [] } },
    function Chart() {},
    console,
    (message) => alerts.push(message),
    { clipboard: null },
    noop,
    noop
);

/* ---------- Разбор чисел ---------- */

test('anomalies: отрицательные значения разбираются с разными знаками', () => {
    assert.strictEqual(service.parseNum('-47'), -47);
    assert.strictEqual(service.parseNum('−47'), -47, 'типографский минус должен читаться как минус');
    assert.strictEqual(service.parseNum('-0.001'), -0.001);
});

test('anomalies: дробные и тысячные значения разбираются верно', () => {
    assert.strictEqual(service.parseNum('103.5'), 103.5);
    assert.strictEqual(service.parseNum('1,5'), 1.5, 'запятая как десятичный разделитель');
    assert.strictEqual(service.parseNum('431,873.44'), 431873.44, 'запятая как разделитель тысяч');
    assert.strictEqual(service.parseNum('31,260'), 31260, 'трёхзначная группа после запятой - тысячи');
    assert.strictEqual(service.parseNum('0,001'), 0.001, 'короткая дробная часть остаётся дробью');
    assert.strictEqual(service.parseNum('31 260,00'), 31260);
});

test('anomalies: пустые и некорректные значения дают null', () => {
    assert.strictEqual(service.parseNum(''), null);
    assert.strictEqual(service.parseNum(null), null);
    assert.strictEqual(service.parseNum(undefined), null);
    assert.strictEqual(service.parseNum('н/д'), null);
});

test('anomalies: ячейка-дата с числовым форматом возвращается как число', () => {
    // В выгрузке встречается ячейка «Сумма» 31260, оформленная как дата (31.12.1985)
    const asDate = new Date(Date.UTC(1985, 6, 31, 20, 0, 0));
    assert.strictEqual(service.parseNum(asDate), 31260);
});

test('anomalies: серийный номер даты форматируется', () => {
    assert.strictEqual(service.formatExcelDate(44512), new Date(Math.round((44512 - 25569) * 86400000)).toLocaleDateString('ru-RU'));
    assert.strictEqual(service.formatExcelDate(''), '');
});

/* ---------- Разбор отчёта аномалий ---------- */

const anomalyTable = [
    ['Примененные фильтры:', 'id_location является 14'],
    [],
    ['Магазин', 'Отдел', 'Артикул', 'Гамма', 'Наименование', 'Теор. запас', 'Доступно для продажи',
        'В торговом зале', 'Резерв для клиентов', 'Резерв для возврата', 'На RM', 'На ЕМ', 'На RD',
        'буфер склада', 'C3', 'Экспо'],
    ['014 Ростов', '9', '84009732', 'S', 'УКРАШЕНИЕ', '5', '5', '-47', '0', '0', '52', '0', '0', '0', '-56', '0'],
    ['014 Ростов', '8', '89436895', 'A0', 'ТРУБА', '15', '13', '-45', '2', '0', '60', '0', '0', '0', '-9', '0'],
    ['014 Ростов', '7', '89427429', 'A', 'ЛИН', '99.999', '99.999', '-0.001', '0', '0', '0', '0', '0', '0', '0', '0'],
    ['014 Ростов', '7', '83515685', 'S', 'РАКОВИНА', '7', '7', '-12', '0', '0', '0', '0', '0', '0', '0', '0'],
];

test('anomalies: отчёт аномалий разбирается с мета-строкой и пустой строкой', () => {
    const parsed = service.parseAnomalyReport(anomalyTable);
    assert.strictEqual(parsed.rows.length, 4);
    assert.strictEqual(parsed.meta.length, 1);
    assert.ok(parsed.meta[0].includes('id_location'));
});

test('anomalies: дефицит считается как модуль «В торговом зале»', () => {
    const parsed = service.parseAnomalyReport(anomalyTable);
    assert.strictEqual(parsed.rows[0].deficit, 47);
    assert.strictEqual(parsed.rows[1].deficit, 45);
    assert.strictEqual(parsed.rows[2].deficit, 0.001);
    // C3 и экспо в дефицит не входят
    assert.strictEqual(parsed.rows[0].c3Deficit, 56);
    assert.strictEqual(parsed.rows[0].expoSurplus, 0);
});

test('anomalies: отчёт без колонки «Артикул» отклоняется', () => {
    assert.throws(() => service.parseAnomalyReport([['Магазин', 'Отдел'], ['x', '1']]), /Артикул/);
});

/* ---------- Разбор отчёта склада ---------- */

const warehouseTable = [
    ['Picking Box', 'Осн. адрес', 'ЗОНА', 'ЯЧЕЙКА', 'ЯЧЕЙКА ТИП', 'ОТДЕЛ', 'ОПИСАНИЕ', 'LM Код',
        'ЛОГ. ТОП', 'ТОП', 'Гамма', 'AVS', 'Дата AVS', 'ЕИ', 'СВОБОДНО', 'НАЛИЧИЕ', 'В ПОДБОРЕ',
        'ЗАБЛОКИРОВАНО', 'SSCC', 'МОНО?', 'Сумма', 'Комментарий к паллете'],
    ['', '', 'EM', 'EA000', 'ХРАНЕНИЕ ПАЛ', '8', 'РАДИАТОР', '12876560', 'EM', '0', 'B', 'Нет', '', 'ШТ', '34', '34', '', '', '14X0196688', 'Нет', '431,873.44', ''],
    ['', '', 'RM', 'RQ071', 'ХРАНЕНИЕ ТОВ', '7', 'РАКОВИНА', '83515685', '', '0', 'S', 'Нет', '', 'ШТ', '7', '7', '', '', 'Всего в 2 паллетах.', 'Нет', '35,084.00', ''],
];

test('anomalies: отчёт склада разбирается, суммы с запятой читаются верно', () => {
    const parsed = service.parseWarehouseReport(warehouseTable);
    assert.strictEqual(parsed.rows.length, 2);
    assert.strictEqual(parsed.rows[0].sum, 431873.44);
    assert.strictEqual(parsed.rows[1].sum, 35084);
    assert.strictEqual(parsed.rows[0].qty, 34);
    assert.strictEqual(parsed.rows[0].zone, 'EM');
});

test('anomalies: текст в SSCC переносится в комментарий', () => {
    const parsed = service.parseWarehouseReport(warehouseTable);
    assert.strictEqual(parsed.rows[1].sscc, '');
    assert.strictEqual(parsed.rows[1].comment, 'Всего в 2 паллетах.');
});

test('anomalies: отчёт склада без колонки «LM Код» отклоняется', () => {
    assert.throws(() => service.parseWarehouseReport([['ЗОНА', 'ЯЧЕЙКА'], ['EM', 'EA000']]), /LM Код/);
});

/* ---------- Слияние отчётов ---------- */

test('anomalies: артикулы сопоставляются с адресами по коду', () => {
    const anomalies = service.parseAnomalyReport(anomalyTable);
    const warehouse = service.parseWarehouseReport(warehouseTable);
    service.anomalyRows = anomalies.rows;
    service.warehouseRows = warehouse.rows;
    service.buildIndexes();

    assert.strictEqual(service.getWarehouseFor('83515685').length, 1);
    assert.strictEqual(service.getWarehouseFor('84009732').length, 0, 'артикула без склада быть не должно');
    const stats = service.anomalyWarehouseStats(anomalies.rows[0]);
    assert.strictEqual(stats.rows.length, 0);
    assert.strictEqual(stats.qty, 0);
});

test('anomalies: итоги по найденному артикулу считаются по адресам', () => {
    const anomalies = service.parseAnomalyReport(anomalyTable);
    const warehouse = service.parseWarehouseReport(warehouseTable);
    service.anomalyRows = anomalies.rows;
    service.warehouseRows = warehouse.rows;
    service.buildIndexes();

    const stats = service.anomalyWarehouseStats(anomalies.rows[3]);
    assert.strictEqual(stats.rows.length, 1);
    assert.strictEqual(stats.qty, 7);
    assert.strictEqual(stats.sum, 35084);
    assert.strictEqual(stats.cells, 1);
});

/* ---------- Структура сервиса ---------- */

test('anomalies: сервис зарегистрирован в конфигурации портала', () => {
    const config = fs.readFileSync('./shared/js/config.js', 'utf8');
    assert.ok(config.includes("id: 'anomalies'"));
    assert.ok(config.includes("url: 'Services/anomalies/index.html'"));
});

test('anomalies: оболочка сервиса добавлена в precache Service Worker', () => {
    const sw = fs.readFileSync('./sw.js', 'utf8');
    assert.ok(sw.includes("'./Services/anomalies/index.html'"));
});

test('anomalies: ??????? ??????? ????? ? ????? samples_XLSX', () => {
    assert.ok(fs.existsSync('./not indexed/samples_XLSX/anomalies_Список аномалий запаса.xlsx'));
    assert.ok(fs.existsSync('./not indexed/samples_XLSX/anomalies_содержимое ячеек.xlsx'));
});

test('anomalies: сервис собран на одном экране: показатели, графики, список и детализация', () => {
    assert.ok(serviceHtml.includes('id="statsGrid"'), 'нет показателей');
    assert.ok(serviceHtml.includes('id="anomalyList"'), 'нет списка аномалий');
    assert.ok(serviceHtml.includes('id="detailBody"'), 'нет детализации артикула');
    assert.ok(serviceHtml.includes('deptChart') && serviceHtml.includes('zoneChart') && serviceHtml.includes('cellChart'), 'нет графиков');
    assert.ok(!serviceHtml.includes('coverageChart'), 'лишний график покрытия');
    assert.ok(!serviceHtml.includes('structureChart'), 'лишний график состава отклонений');
    assert.ok(!serviceHtml.includes('topChart'), 'лишний график топ-10 артикулов');
    assert.strictEqual((serviceHtml.match(/class="charts-row"/g) || []).length, 1, 'графики должны быть в одном ряду');
    assert.ok(!serviceHtml.includes('data-tab="'), 'в сервисе не должно быть вкладок');
    assert.ok(serviceHtml.includes('LM Код'), 'нет ключа сопоставления с отчётом склада');
    assert.ok(serviceHtml.includes('В торговом зале'), 'нет колонки с аномальным количеством');
    assert.ok(serviceHtml.includes('copyAddresses'), 'нет копирования адресов');
    assert.ok(serviceHtml.includes('openTaskModal'), 'нет задания для сборщика');
    assert.ok(!serviceHtml.includes('exportAnomalies'), 'лишняя выгрузка аномалий');
    assert.ok(!serviceHtml.includes('exportAddresses'), 'лишняя выгрузка адресов');
    assert.ok(!serviceHtml.includes('exportMissing'), 'лишняя выгрузка ненайденных');
});

test('anomalies: из сервиса убраны фильтры и отдельный блок склада', () => {
    ['searchInput', 'deptFilter', 'gammaFilter', 'statusFilter', 'zoneFilter', 'anomalyFilters',
        'warehouseDetails', 'warehouseTree', 'warehouseTable', 'whSearch'].forEach(id => {
        assert.ok(!serviceHtml.includes(id), `лишний блок остался: ${id}`);
    });
});

test('anomalies: в показателях остались только три карточки', () => {
    const stats = serviceHtml.match(/class="stat-label">([^<]+)</g) || [];
    const labels = stats.map(item => item.replace(/class="stat-label">/, '').replace(/<$/, '').trim());
    assert.deepEqual(labels, ['Аномальных артикулов', 'Найдено на складе', 'На складе, шт.']);
});

test('anomalies: оформление и инфо-окно собраны по образцу «AVS Контроль»', () => {
    ['avs-critical', 'avs-normal', 'avs-warning', 'instruction-section', 'instruction-steps',
        'info-cards', 'warning-box', 'important-header'].forEach(cls => {
        assert.ok(serviceHtml.includes(cls), `нет элемента оформления AVS: ${cls}`);
    });
    assert.ok(serviceHtml.includes('repeat(3, minmax(0, 1fr))'), 'графики должны стоять по три в ряд');
    assert.ok(serviceHtml.includes('Найдены на складе') || serviceHtml.includes('Не найдены на складе'),
        'в инфо-окне потерян контекст аномалий');
});

test('anomalies: графики считают SKU и стоят над списком аномалий', () => {
    assert.ok(serviceHtml.includes('Аномальные артикулы по отделам, SKU'), 'нет графика отделов в SKU');
    assert.ok(serviceHtml.includes('Аномальные артикулы по зонам склада, SKU'), 'нет графика зон в SKU');
    assert.ok(serviceHtml.includes('Топ-10 ячеек по аномальным артикулам, SKU'), 'нет графика ячеек в SKU');
    assert.ok(!/дефицит/i.test(serviceHtml), 'термин «дефицит» убран как избыточный');
    const chartsIndex = serviceHtml.indexOf('class="charts-row"');
    const splitIndex = serviceHtml.indexOf('class="split-layout"');
    assert.ok(chartsIndex !== -1 && splitIndex !== -1 && chartsIndex < splitIndex,
        'графики должны стоять в зоне над списком аномалий');
});

test('anomalies: пустой экран не показывает служебные сообщения', () => {
    assert.ok(serviceHtml.includes('.session-warning {'), 'нет правил для предупреждения о режиме сессии');
    assert.ok(serviceHtml.includes('.session-warning.active'), 'предупреждение о режиме сессии не управляется классом');
});

test('anomalies: анализ запускается только после загрузки обоих отчётов', () => {
    assert.ok(serviceHtml.includes('afterFileLoaded'), 'нет общей точки проверки загрузки отчётов');
    assert.ok(!serviceHtml.includes('afterDataChanged'), 'данные обрабатываются до загрузки второго отчёта');
    assert.ok(serviceHtml.includes('uploadWaiting'), 'нет подсказки об ожидании второго отчёта');
    assert.ok(serviceHtml.includes('hasAnomalies && hasWarehouse'), 'нет условия запуска по двум отчётам');
});

test('anomalies: приоритет - позиции, найденные на складе', () => {
    assert.ok(serviceHtml.includes("listMode = 'found'"), 'нет режима «только найденные на складе»');
    assert.ok(serviceHtml.includes('showAllAnomalies'), 'нет переключателя показа всех позиций');
    assert.ok(serviceHtml.includes('foundToggle'), 'нет кнопки переключателя в шапке списка');
});

test('anomalies: фильтры списка и круговой график зон', () => {
    assert.ok(serviceHtml.includes('id="foundOnlyBtn"'), 'нет кнопки «Только найденные»');
    assert.ok(serviceHtml.includes('id="missingOnlyBtn"'), 'нет кнопки «Нет на складе»');
    assert.ok(serviceHtml.includes("setListMode('missing')"), 'фильтр «Нет на складе» не подключён');
    assert.ok(serviceHtml.includes('listFiltersShown'), 'кнопки фильтров не раскрываются после «Показать всё»');
    assert.ok(serviceHtml.includes("type: 'pie'"), 'график зон должен быть круговой диаграммой');
    assert.match(serviceHtml, /charts\.zone = createChart\('zoneChart', \{\s*type: 'pie'/, 'график зон остался столбчатым');
});

test('anomalies: у ненайденных артикулов остатки зон перечёркнуты', () => {
    ['На RM', 'На ЕМ', 'На RD', 'Буфер склада'].forEach(label => {
        assert.ok(serviceHtml.includes(`zoneMetric('${label}'`), `нет разметки остатка: ${label}`);
    });
    assert.ok(serviceHtml.includes('mb-gone'), 'нет перечёркивания остатков');
    assert.ok(serviceHtml.includes('искать физически негде'), 'нет пояснения о фактическом отсутствии на складе');
});

test('anomalies: в списке вместо гаммы показывается количество на складе', () => {
    assert.ok(serviceHtml.includes('badge-qty'), 'нет шильдика количества на складе');
    assert.ok(serviceHtml.includes('formatCount(stats.qty)'), 'количество не берётся из остатков склада');
    assert.ok(!serviceHtml.includes('escapeHtml(record.gamma)'), 'гамма осталась в карточке списка');
});

test('anomalies: шильдики списка стоят в одну строку без денежной суммы', () => {
    assert.match(serviceHtml, /\.art-item \.ai-meta \{[^}]*flex-wrap: nowrap/, 'шильдики переносятся на новую строку');
    assert.ok(!serviceHtml.includes('badge-ok'), 'денежный шильдик остался в списке');
    assert.ok(!serviceHtml.includes('formatMoneyShort(stats.sum)'), 'сумма всё ещё выводится в карточке списка');
});

test('anomalies: правая панель детализации по высоте равна списку', () => {
    assert.ok(serviceHtml.includes('setupSplitLayoutSync'), 'нет синхронизации высоты панелей');
    assert.ok(serviceHtml.includes('ResizeObserver'), 'высота не обновляется при изменении списка');
    assert.match(serviceHtml, /\.split-layout > \.panel > \.panel-body \{[^}]*overflow-y: auto/, 'длинная детализация не прокручивается внутри панели');
});
