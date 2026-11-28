import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

const SERVICE_PATH = './Services/analyseLogisticOperation/index.html';
const SAMPLE_PATH = './not indexed/samples_XLSX/analyseLogisticOperation.xlsx';

const serviceHtml = fs.readFileSync(SERVICE_PATH, 'utf8');
const xlsxCode = fs.readFileSync('./shared/js-lib/xlsx.full.min.js', 'utf8');
const idbCode = fs.readFileSync('./shared/js/idb.js', 'utf8');
const commonCode = fs.readFileSync('./shared/js/common.js', 'utf8');

const scriptMatch = serviceHtml.match(/<script>([\s\S]*?)<\/script>/);
assert.ok(scriptMatch, 'в сервисе не найден встроенный скрипт');

const xlsxSandbox = {};
const XLSX = new Function(
    'self', 'window', 'globalThis', 'module', 'exports',
    `${xlsxCode}\nreturn (typeof XLSX !== 'undefined' ? XLSX : self.XLSX);`
)(xlsxSandbox, xlsxSandbox, xlsxSandbox, undefined, undefined);

const noop = () => {};
const fakeDocument = {
    readyState: 'loading',
    addEventListener: noop,
    getElementById: () => null,
    querySelectorAll: () => [],
    querySelector: () => null,
    body: { dataset: {}, classList: { add: noop, remove: noop, toggle: noop } },
    createElement: () => ({
        style: {},
        classList: { add: noop, remove: noop, toggle: noop },
        appendChild: noop,
        addEventListener: noop,
    }),
};
const fakeWindow = {
    addEventListener: noop,
    removeEventListener: noop,
    dispatchEvent: () => true,
    localStorage: { getItem: () => null, setItem: noop },
};
const config = {
    PORTAL_CONFIG: { services: [] },
    PORTAL_VERSION: '1.0.0',
    PORTAL_RELEASE_DATE: '25.11.2026',
};

const factory = new Function(
    'self', 'window', 'document', 'globalThis', 'module', 'exports', 'XLSX', 'Chart',
    `${idbCode}\n${commonCode}\n${scriptMatch[1]}\nreturn {
        anaParseWorkbook, anaHydrate, anaWeight, anaFilterRows,
        anaAggEmployees, anaAggDays, anaAggHours, anaTotals, anaShiftStats, anaShiftOfSlot,
        anaFilterBySlots, anaRankComparator, anaTableColumns,
        anaFilterOptionValues, anaSlotsToHours, anaRate, anaFmtRate, anaFmtOpt,
        anaProdComparator, anaProdSortLabel, ANA_PROD_SORT_OPTIONS,
        ANA_OP_CAT, ANA_WEIGHTS, ANA_TABS, ANA_CAT_ORDER, ANA_OPS_BY_CAT, ANA_CATEGORY_STATS,
        ANA_SLOT_FILTER_OPTIONS, ANA_SLOT_FILTER_DEFAULT,
        ANA_FILTER_FIELDS, ANA_RANK_SORT_DEFAULT, ANA_RATE_MIN_SLOTS
    };`
);
const api = factory(config, fakeWindow, fakeDocument, config, undefined, undefined, XLSX, undefined);

const wb = XLSX.read(fs.readFileSync(SAMPLE_PATH), { type: 'buffer', cellDates: true });
const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '', raw: true });
const parsed = api.anaParseWorkbook(json);
const rows = api.anaHydrate(parsed.records);
const validRows = rows.filter(r => r.ldap && r.fio && r.ldap !== 'нет данных' && r.fio !== 'нет данных');

const sum = (list, fn) => list.reduce((total, item) => total + fn(item), 0);
const close = (actual, expected, epsilon = 1e-6) => Math.abs(actual - expected) <= epsilon;

/* Синтетические строки: 6 слотов в день. */
function slotRows(ldap, cat, slots, ops, opts = {}) {
    const ctx = opts.ctx || 'sku';
    const qty = opts.qty || 0;
    const op = opts.op || api.ANA_OPS_BY_CAT[cat][0];
    const rows = [];
    let left = slots;
    let day = 1;
    while (left > 0) {
        const n = Math.min(6, left);
        for (let s = 0; s < n; s++) {
            const slot = s * 10;
            rows.push({
                date: `2026-11-${String(day).padStart(2, '0')}`,
                slot,
                hour: Math.floor(slot / 60),
                op,
                ctx,
                ops,
                qty,
                ldap,
                fio: `Сотрудник ${ldap}`,
                pos: '',
                dept: 'тест',
                deptType: 'логистика',
                cat,
                w: ops * api.anaWeight(op, ctx),
            });
        }
        left -= n;
        day++;
    }
    return rows;
}

/* ---------- Структура сервиса ---------- */

test('analyseLogisticOperation: пять вкладок, единая палитра и диаграммы', () => {
    const labels = ['📊 Продуктивность', '📦 Контекстная продуктивность', '📅 Динамика по дням', '🕒 Динамика по часам', '🌗 День и ночь'];
    labels.forEach(label => assert.ok(serviceHtml.includes(label), `нет вкладки ${label}`));
    assert.deepStrictEqual(api.ANA_TABS.map(t => t.key), ['rating', 'productivity', 'daily', 'hourly', 'shifts']);
    ['🏆 Рейтинг', '📦 Пополнение', '🚚 ОЗК', '🏭 Складские операции', '👥 Внешний персонал', '🎯 Операции с достоверностью'].forEach(label => {
        assert.ok(!serviceHtml.includes(label), `осталась вкладка ${label}`);
    });
    assert.ok(serviceHtml.includes('data-theme="emerald"'));
    assert.ok(serviceHtml.includes('--accent: #c97b4a'));
    assert.ok(!serviceHtml.includes('id="themeSelect"'), 'переключатель палитры не удалён');
    assert.ok(!serviceHtml.includes('>📖 Методика</button>'), 'кнопка «Методика» не удалена');
    assert.ok(!serviceHtml.includes('Единый период · 8 вкладок'), 'пояснительная надпись не удалена');
    assert.ok(!serviceHtml.includes('openAnaInfo'), 'остался вызов openAnaInfo');
    assert.ok(!serviceHtml.includes('Мин. слотов'), 'остался фильтр по слотам');
    assert.ok(!serviceHtml.includes('anaSetMinSlots'), 'осталась функция anaSetMinSlots');
    assert.ok(serviceHtml.includes('anaSlotFilter-${k}'), 'нет фильтра по слотам в блоке топ-10');
    assert.ok(serviceHtml.includes('anaSetRankSort'), 'нет сортировки топов');
    assert.ok(serviceHtml.includes("['w', 'По взвеш. операциям']"), 'нет сортировки по взвешенным операциям');
    assert.ok(serviceHtml.includes('anaSetSlotFilter'), 'нет обработчика порога слотов');
    assert.ok(!serviceHtml.includes('anaSetCohort'), 'остался переключатель среза');
    assert.ok(!serviceHtml.includes('ANA_COHORTS'), 'остался список срезов');
    assert.ok(!serviceHtml.includes('cohort-btn'), 'остались кнопки среза');
    assert.ok(!serviceHtml.includes('anaShareBarChart'), 'остались процентные бары');
    assert.strictEqual((serviceHtml.match(/anaDoughnutChart\(`chart-\$\{key\}-/g) || []).length, 3, 'не все графики сотрудников стали диаграммами');
    assert.ok(serviceHtml.includes('Операции по категориям, %'), 'нет графика долей по категориям');
    assert.ok(serviceHtml.includes('Структура по категориям, % (взвеш. операции)'), 'не уточнена единица структуры категорий');
    assert.ok(serviceHtml.includes('Контексты операций, %'), 'нет графика долей по контекстам');
    assert.ok(!serviceHtml.includes('Топ-12 по КПД'), 'остался график топ-12 по КПД');
    assert.ok(!serviceHtml.includes('anaRankMetric'), 'осталась функция anaRankMetric');
    assert.ok(!serviceHtml.includes('Рейтинг, %'), 'осталась колонка рейтинга');
    assert.ok(!serviceHtml.includes('anaApplyRatings'), 'остался расчёт рейтинга');
    assert.ok(!serviceHtml.includes('anaRatingFallback'), 'остался fallback рейтинга');
    assert.ok(!serviceHtml.includes('Топ-5 дней по операциям'), 'остался топ дней');
    assert.ok(!serviceHtml.includes('Топ-5 часов по операциям'), 'остался топ часов');
});

test('analyseLogisticOperation: тепловые карты операций и категорий по часам', () => {
    assert.ok(serviceHtml.includes('id="heat-hourly-ops"'));
    assert.ok(serviceHtml.includes('id="heat-hourly-cats"'));
    assert.ok(serviceHtml.includes('Операции × часы'));
    assert.ok(serviceHtml.includes('Категории операций × часы'));
    assert.ok(serviceHtml.includes('function anaRenderHourHeatmaps'));
    assert.ok(serviceHtml.includes('anaRenderHourHeatmaps(rows)'));
    assert.ok(!serviceHtml.includes('heat-total-ops'), 'осталась сумма операций в тепловой карте');
    assert.ok(!serviceHtml.includes('heat-total-cats'), 'осталась сумма категорий в тепловой карте');
    assert.ok(!serviceHtml.includes('<th>Σ</th>'), 'осталась колонка сумм в тепловой карте');
    assert.ok(!serviceHtml.includes('heat-total-cell'), 'осталась ячейка суммы в тепловой карте');
    assert.ok(!/class="heat-total"/.test(serviceHtml), 'остался бейдж суммы у тепловой карты');
});

test('analyseLogisticOperation: дни и часы без таблиц, с новыми блоками', () => {
    assert.ok(serviceHtml.includes('id="daily-context-chart"'), 'нет графика контекстов по дням');
    assert.ok(serviceHtml.includes('id="daily-weekday-qty-chart"'), 'нет графика среднего объёма по дням недели');
    assert.ok(serviceHtml.includes('id="daily-cat-period-chart"'), 'нет графика операций по категориям за период');
    assert.ok(serviceHtml.includes('id="daily-weekday-ops-chart"'), 'нет графика средних операций по дням недели');
    assert.ok(serviceHtml.includes('Динамика по дням • разрез по контексту операций'), 'нет заголовка блока дней');
    assert.ok(serviceHtml.includes('function anaRenderDailyContextChart'), 'нет рендера контекстов по дням');
    assert.ok(serviceHtml.includes('anaRenderDailyContextChart(rows)'), 'рендер контекстов по дням не вызывается');
    assert.ok(serviceHtml.includes('function anaRenderDailyWeekdayCharts'), 'нет рендера дней недели');
    assert.ok(serviceHtml.includes('anaRenderDailyWeekdayCharts(rows)'), 'рендер дней недели не вызывается');
    assert.ok(serviceHtml.includes('function anaRenderDailyCatPeriodChart'), 'нет рендера категорий за период');
    assert.ok(serviceHtml.includes('anaRenderDailyCatPeriodChart(rows)'), 'рендер категорий за период не вызывается');
    assert.strictEqual((serviceHtml.match(/charts-row wide-left/g) || []).length >= 2, true, 'блок дней должен быть в две строки по два графика');
    assert.ok(serviceHtml.includes('id="hourly-cat-charts"'), 'нет контейнера часовых графиков категорий');
    assert.ok(serviceHtml.includes('function anaRenderHourlyCategoryCharts'), 'нет рендера часовых категорий');
    assert.ok(serviceHtml.includes('anaRenderHourlyCategoryCharts(rows)'), 'рендер часовых категорий не вызывается');
    assert.ok(!serviceHtml.includes('id="chart-hourly-1"'), 'остался график «Операции по часам»');
    assert.ok(!serviceHtml.includes('id="chart-hourly-2"'), 'остался график «Профиль 10-минутных слотов»');
    assert.ok(!serviceHtml.includes('id="chart-hourly-3"'), 'остался график «Взвешенные операции по часам»');
    assert.ok(!serviceHtml.includes('Операции по часам'), 'остался заголовок графика «Операции по часам»');
    assert.ok(!serviceHtml.includes('Профиль 10-минутных слотов'), 'остался заголовок профиля слотов');
    assert.ok(!serviceHtml.includes('Взвешенные операции по часам'), 'остался заголовок взвешенных по часам');
    assert.ok(serviceHtml.includes("if (tab.kind === 'days' || tab.kind === 'hours' || tab.kind === 'shifts') return;"), 'нет раннего выхода из таблицы');
    assert.ok(serviceHtml.includes("const noTable = tab.kind === 'days' || tab.kind === 'hours' || tab.kind === 'shifts';"), 'нет флага отсутствия таблицы');
});

test('analyseLogisticOperation: часовая вкладка - блоки, подписи и статистика', () => {
    assert.ok(serviceHtml.includes('class="hourly-block"'), 'нет блока часовой вкладки');
    assert.ok(serviceHtml.includes('hourly-block-title'), 'нет заголовка блока');
    assert.ok(serviceHtml.includes('hourly-block-hint'), 'нет пояснения блока');
    assert.ok(serviceHtml.includes('Операции по категориям за 24 часа'), 'нет подписи блока категорий');
    assert.ok(serviceHtml.includes('Тепловые карты'), 'нет подписи блока тепловых карт');
    assert.ok(serviceHtml.includes("isHours ? 'Динамика по часам'"), 'нет заголовка вкладки');
    assert.ok(serviceHtml.includes("isHours ? '24 часа'"), 'нет бейджа «24 часа»');
    assert.ok(!serviceHtml.includes('<div class="section-header"><div class="sec-icon">⏰</div>'), 'остался вложенный заголовок в карточке');
    assert.ok(serviceHtml.includes('Пиковые часы ${peakLabels[cat]}'), 'нет карточки пиковых часов категории');
    ['приёмки', 'пополнения', 'складских операций', 'достоверности']
        .forEach(label => assert.ok(serviceHtml.includes(`'${label}'`), `нет подписи «${label}»`));
    assert.ok(serviceHtml.includes("orders: 'ОЗК'"), 'нет подписи ОЗК');
    assert.ok(serviceHtml.includes('peakLabels'), 'нет расчёта пиковых часов категорий');
    assert.ok(serviceHtml.includes('slice(0, 3)'), 'учитывается не три пиковых часа');
    assert.ok(!serviceHtml.includes("'Пиковый слот'"), 'осталась карточка пикового слота');
    assert.ok(serviceHtml.includes('maxRotation = 0'), 'подписи времени всё ещё поворачиваются');
    assert.ok(serviceHtml.includes('maxTicksLimit = 9'), 'нет ограничения числа подписей');
    assert.ok(serviceHtml.includes('ticks.padding = 6'), 'нет отступа подписей');
    assert.ok(serviceHtml.includes('function anaWrapCanvas'), 'нет обёртки канваса');
    assert.ok(serviceHtml.includes("box.className = 'chart-box'"), 'нет контейнера канваса');
    assert.ok(serviceHtml.includes('flex: 1 1 auto'), 'обёртка канваса не растягивается');
    assert.ok(serviceHtml.includes('anaCharts[id].resize()'), 'графики профиля не пересчитываются после открытия');
    assert.ok(serviceHtml.includes('max-height: none'), 'тепловые карты всё ещё прокручиваются по высоте');
    assert.ok(serviceHtml.includes('ANA_HEAT_GREEN'), 'нет контрастной палитры тепловых карт');
    assert.ok(serviceHtml.includes('ANA_HEAT_COPPER'), 'нет медного акцента для самых горячих ячеек');
    assert.ok(serviceHtml.includes('lum > 145'), 'цвет текста не подстраивается под фон');
});

test('analyseLogisticOperation: вкладка сотрудников без срезов, с карточками категорий', () => {
    assert.ok(!serviceHtml.includes('ANA_DEPTS'), 'остался список подразделений-вкладок');
    assert.ok(!serviceHtml.includes('ANA_TOP_ONLY_KEYS'), 'остался список вкладок-гистограмм');
    assert.ok(!serviceHtml.includes('anaIsTopOnlyTab'), 'осталась проверка вкладки-гистограммы');
    assert.ok(!serviceHtml.includes('anaRenderDeptTopCharts'), 'остался рендер гистограмм подразделений');
    ['topqty-', 'topops-', 'topw-'].forEach(id => {
        assert.ok(!serviceHtml.includes(id), `осталась гистограмма ${id}`);
    });
    assert.ok(serviceHtml.includes('catFilter: true'), 'у вкладки сотрудников нет фильтра категорий');
    assert.ok(!serviceHtml.includes('block-tabs'), 'остались бейджи-вкладки категорий над таблицей');
    assert.ok(!serviceHtml.includes('anaSetCatTab'), 'остался переключатель вкладок категорий');
    assert.ok(!serviceHtml.includes('anaRenderBlockTabs'), 'остался рендер бейджей категорий');
    assert.ok(!serviceHtml.includes('anaCohort'), 'осталось состояние среза');
    assert.ok(!serviceHtml.includes('Достоверностью'), 'осталась подпись среза');
    assert.ok(serviceHtml.includes('ANA_CATEGORY_STATS'), 'нет списка категорий для карточек');
    ['Внешний персонал', 'Пополнение', 'Достоверность', 'Приёмка', 'Складские операции', 'Сборка и выдача заказов']
        .forEach(label => assert.ok(serviceHtml.includes(`label: '${label}'`), `нет карточки «${label}»`));
    assert.ok(serviceHtml.includes('countWhere'), 'нет подсчёта сотрудников по категориям');
    assert.ok(serviceHtml.includes('задействовано ·'), 'в карточках не уточнено, что показаны задействованные сотрудники');
    assert.ok(serviceHtml.includes('anaRenderUniversality'), 'нет карты универсальности');
    assert.ok(serviceHtml.includes('Карта универсальности по сотрудникам'), 'нет заголовка карты универсальности');
    assert.ok(serviceHtml.includes('sec-matrix'), 'карта универсальности не вынесена в отдельный блок');
    assert.ok(serviceHtml.includes('uni-toolbar'), 'нет панели фильтра по универсальности');
    assert.ok(serviceHtml.includes('anaSetUniLevel'), 'фильтр по универсальности не подключён');
    assert.ok(!serviceHtml.includes('anaUniFilters'), 'у карты остались отдельные фильтры');
    assert.ok(serviceHtml.includes("anaRenderUniversality(key);"), 'карта не рендерится вместе с вкладкой');
});

test('analyseLogisticOperation: единый период и раздельные фильтры', () => {
    assert.ok(serviceHtml.includes('id="periodFrom"'));
    assert.ok(serviceHtml.includes('id="periodTo"'));
    assert.ok(serviceHtml.includes('function anaSelectPeriodDate'), 'нет выбора дня кликом по гистограмме периода');
    assert.ok(serviceHtml.includes('anaSelectPeriodDate(bar.dataset.idx)'), 'клик по столбцу периода не подключён');
    assert.ok(serviceHtml.includes('role="button" tabindex="0"'), 'столбцы периода недоступны с клавиатуры');
    assert.ok(serviceHtml.includes('function anaPeriodLabel'), 'нет единой подписи периода');
    assert.ok(serviceHtml.includes('anaSetChartFilter'));
    assert.ok(serviceHtml.includes('anaSetTableFilter'));
    assert.ok(serviceHtml.includes("anaFilterSelect('🗂️ Категория'"), 'нет фильтра категорий в таблице');
    assert.ok(serviceHtml.includes('anaResetFilters'));
    assert.ok(serviceHtml.includes('При переходе на вкладку фильтры всегда начинаются с чистого листа'), 'фильтры не сбрасываются при переключении вкладок');
    assert.ok(serviceHtml.includes('anaResetTabState(key)'), 'нет сброса состояния вкладки при переключении');
});

test('analyseLogisticOperation: карточка сотрудника и профиль', () => {
    assert.ok(serviceHtml.includes('openAnaProfile'));
    assert.ok(serviceHtml.includes('https://intraru.lemanapro.ru/profile/'));
    assert.ok(serviceHtml.includes('id="profileModal"'));
    assert.ok(serviceHtml.includes('id="profileDayChart"'));
    assert.ok(serviceHtml.includes('id="profileHourChart"'));
    /* Открытие карточки не должно прокручивать страницу к началу: body
       блокируется только через overflow, без position: fixed. */
    assert.ok(!/body\.no-scroll\s*\{[^}]*position:\s*fixed/.test(serviceHtml), 'body.no-scroll фиксируется и страница прыгает вверх');
    assert.ok(serviceHtml.includes("document.body.style.paddingRight = `${pad + sbw}px`"), 'нет компенсации исчезнувшего скроллбара');
    assert.ok(serviceHtml.includes("document.body.style.removeProperty('padding-right')"), 'компенсация скроллбара не снимается при закрытии');
});

test('analyseLogisticOperation: IndexedDB и версия сервиса', () => {
    assert.ok(serviceHtml.includes('lemanapro_analyse_logistic_operation'));
    assert.ok(serviceHtml.includes('analyse_logistic_operation_report'));
    assert.ok(serviceHtml.includes('data-version-service="analyse_logistic_operation"'));
    assert.ok(!serviceHtml.includes('__ANA_SCRIPT__'));
    assert.ok(!serviceHtml.includes('__ANA_PART_'));
});

test('analyseLogisticOperation: 22 операции распределены по 5 категориям', () => {
    assert.strictEqual(Object.keys(api.ANA_OP_CAT).length, 22);
    assert.strictEqual(api.ANA_CAT_ORDER.length, 5);
    for (const cat of api.ANA_CAT_ORDER) {
        assert.ok(api.ANA_OPS_BY_CAT[cat].length > 0, `пустая категория ${cat}`);
        for (const op of api.ANA_OPS_BY_CAT[cat]) assert.strictEqual(api.ANA_OP_CAT[op], cat);
    }
});

/* ---------- Парсинг файла ---------- */

test('analyseLogisticOperation: заголовки находятся на третьей строке, 33 124 строки', () => {
    assert.strictEqual(parsed.headerRow, 2);
    assert.strictEqual(parsed.records.length, 33124);
    assert.strictEqual(rows.length, 33124);
    assert.strictEqual(sum(rows, r => r.ops), 112340);
    assert.ok(close(sum(rows, r => r.qty), 4438252.233, 1e-3));
    assert.ok(close(sum(rows, r => r.w), 198706, 1e-3));
});

test('analyseLogisticOperation: даты приведены к 01.11-16.11.2026', () => {
    const dates = [...new Set(rows.map(r => r.date))].sort();
    assert.strictEqual(dates.length, 16);
    assert.strictEqual(dates[0], '2026-11-01');
    assert.strictEqual(dates[dates.length - 1], '2026-11-16');
    assert.ok(rows.every(r => r.slot >= 0 && r.slot < 1440 && r.slot % 10 === 0));
});

test('analyseLogisticOperation: контексты SKU / PALLET / ORDER', () => {
    const ctx = new Map();
    rows.forEach(r => {
        if (!ctx.has(r.ctx)) ctx.set(r.ctx, { rows: 0, ops: 0 });
        const item = ctx.get(r.ctx);
        item.rows++;
        item.ops += r.ops;
    });
    assert.deepStrictEqual([...ctx.keys()].sort(), ['order', 'pallet', 'sku']);
    assert.deepStrictEqual(ctx.get('sku'), { rows: 21876, ops: 86366 });
    assert.deepStrictEqual(ctx.get('pallet'), { rows: 5252, ops: 15834 });
    assert.deepStrictEqual(ctx.get('order'), { rows: 5996, ops: 10140 });
});

test('analyseLogisticOperation: категории и трудоёмкость', () => {
    const cats = new Map();
    rows.forEach(r => {
        if (!cats.has(r.cat)) cats.set(r.cat, { rows: 0, ops: 0, w: 0 });
        const item = cats.get(r.cat);
        item.rows++;
        item.ops += r.ops;
        item.w += r.w;
    });
    const expected = {
        replenishment: { rows: 13812, ops: 53585, w: 107170 },
        warehouse: { rows: 7127, ops: 16766, w: 23650 },
        orders: { rows: 10280, ops: 26890, w: 43640 },
        reception: { rows: 597, ops: 5953, w: 5954 },
        accuracy: { rows: 268, ops: 785, w: 1570 },
        other: { rows: 1040, ops: 8361, w: 16722 },
    };
    for (const [cat, value] of Object.entries(expected)) {
        const actual = cats.get(cat);
        assert.strictEqual(actual.rows, value.rows, `${cat}: строки`);
        assert.strictEqual(actual.ops, value.ops, `${cat}: операции`);
        assert.ok(close(actual.w, value.w, 1e-3), `${cat}: трудоёмкость ${actual.w}`);
    }
});

test('analyseLogisticOperation: простые веса по контексту', () => {
    assert.deepStrictEqual({ ...api.ANA_WEIGHTS }, { sku: 2, pallet: 1, order: 1 });
    for (const op of Object.keys(api.ANA_OP_CAT)) {
        assert.strictEqual(api.anaWeight(op, 'sku'), 2, `${op}: SKU`);
        assert.strictEqual(api.anaWeight(op, 'pallet'), 1, `${op}: PALLET`);
        assert.strictEqual(api.anaWeight(op, 'order'), 1, `${op}: ORDER`);
    }
});

test('analyseLogisticOperation: фильтрация строк по контексту', () => {
    const pallet = api.anaFilterRows(rows, { ctx: 'pallet' });
    assert.strictEqual(pallet.length, 5252);
    assert.strictEqual(sum(pallet, r => r.ops), 15834);
});

test('analyseLogisticOperation: списки фильтров не содержат недоступных значений', () => {
    const empty = { search: '', pos: '', ctx: '', type: '', dept: '', cat: '', op: '' };

    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, empty, 'ctx'), ['sku', 'pallet', 'order']);
    assert.strictEqual(api.anaFilterOptionValues(validRows, empty, 'type').length, 4);
    assert.strictEqual(api.anaFilterOptionValues(validRows, empty, 'dept').length, 24);
    assert.strictEqual(api.anaFilterOptionValues(validRows, empty, 'pos').length, 14);
    assert.strictEqual(api.anaFilterOptionValues(validRows, empty, 'op').length, 23);
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, empty, 'cat'), api.ANA_CAT_ORDER);

    const pallet = { ...empty, ctx: 'pallet' };
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, pallet, 'type'), ['логистика']);
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, pallet, 'op'),
        ['1я приемка', 'Перемещение в ТЗ', 'Подбор', 'Размещение', 'Размещение с приемки']);
    api.anaFilterOptionValues(validRows, pallet, 'dept').forEach(dept => {
        assert.ok(validRows.some(r => r.ctx === 'pallet' && r.dept === dept), `лишнее подразделение: ${dept}`);
    });

    const dir = { ...empty, type: 'дирекция' };
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, dir, 'ctx'), ['sku']);
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, dir, 'dept'), ['дирекция']);
    assert.strictEqual(api.anaFilterOptionValues(validRows, dir, 'op').length, 8);
    assert.strictEqual(api.anaFilterOptionValues(validRows, dir, 'type').length, 4, 'тип не должен зависеть от подразделения');

    const acc = { ...empty, op: 'Достоверность' };
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, acc, 'ctx'), ['sku']);
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, acc, 'type'), ['дирекция', 'коммерция', 'логистика']);
    assert.deepStrictEqual(api.anaFilterOptionValues(validRows, acc, 'cat'), ['accuracy']);
});

test('analyseLogisticOperation: фильтр «Должность» работает и сужается вместе с другими фильтрами', () => {
    assert.ok(api.ANA_FILTER_FIELDS.includes('pos'));
    assert.ok(serviceHtml.includes("anaFilterSelect('💼 Должность'"), 'нет фильтра «Должность»');

    const external = api.anaFilterRows(rows, { pos: 'внешний персонал' });
    assert.strictEqual(external.length, 9021);
    assert.strictEqual(sum(external, r => r.ops), 25674);

    const searched = api.anaFilterRows(rows, { search: 'штабелера' });
    assert.strictEqual(searched.length, 1379, 'поиск не учитывает должность');

    const empty = { search: '', pos: '', ctx: '', type: '', dept: '', cat: '', op: '' };
    const dirPositions = api.anaFilterOptionValues(validRows, { ...empty, type: 'дирекция' }, 'pos');
    assert.deepStrictEqual(dirPositions, [
        'руководитель отдела по предотвращению внештатных ситуаций (ПВС)',
        'руководитель торгового сектора',
    ]);
});

/* ---------- Агрегации ---------- */

test('analyseLogisticOperation: 218 сотрудников с корректными ЛДАП и ФИО', () => {
    const employees = api.anaAggEmployees(validRows);
    const totals = api.anaTotals(validRows);
    assert.strictEqual(employees.length, 218);
    assert.strictEqual(totals.ops, 112340);
    assert.ok(close(totals.qty, 4438252.233, 1e-3));
    assert.ok(close(totals.w, 198706, 1e-3));
    assert.strictEqual(totals.slots, 27593);
    assert.ok(close(totals.kpd, 7.2013191751531185, 1e-9));
    const top = employees.slice().sort((a, b) => b.kpd - a.kpd)[0];
    assert.ok(top.kpd > 90 && top.kpd < 100, `неожиданный топ-КПД: ${top.kpd}`);
});

test('analyseLogisticOperation: население подразделений доступно через фильтры', () => {
    const count = list => new Set(validRows.filter(r => list.includes(r.dept)).map(r => r.ldap)).size;
    assert.strictEqual(count(['отдел пополнения цепи поставок']), 22);
    assert.strictEqual(count(['отдел выдачи и отгрузки заказов клиента', 'отдел обслуживания заказов клиентов цепи поставок']), 23);
    assert.strictEqual(count(['отдел складских операций цепи поставок']), 30);
    assert.strictEqual(count(['внешний персонал']), 49);
});

test('analyseLogisticOperation: достоверность - 55 сотрудников и 785 операций', () => {
    const employees = api.anaAggEmployees(validRows);
    const acc = employees.filter(e => e.accOps > 0);
    assert.strictEqual(acc.length, 55);
    assert.strictEqual(sum(acc, e => e.accOps), 785);
    assert.ok(close(sum(acc, e => e.accQty), 3534.91, 1e-6));
    assert.ok(acc.every(e => e.deptType === 'коммерция' || e.deptType === 'логистика' || e.deptType === 'дирекция'));
});

test('analyseLogisticOperation: карточки категорий считают уникальных сотрудников', () => {
    const groups = api.ANA_CATEGORY_STATS;
    assert.deepStrictEqual(groups.map(g => g.key), ['external', 'replenishment', 'accuracy', 'reception', 'warehouse', 'orders']);
    const count = g => new Set(validRows.filter(r => g.cat ? r.cat === g.cat : r.dept === g.dept).map(r => r.ldap)).size;
    assert.strictEqual(count(groups[0]), 49);
    assert.strictEqual(count(groups[1]), 176);
    assert.strictEqual(count(groups[2]), 55);
    assert.strictEqual(count(groups[3]), 20);
    assert.strictEqual(count(groups[4]), 49);
    assert.strictEqual(count(groups[5]), 47);
});

test('analyseLogisticOperation: динамика по дням и часам', () => {
    const days = api.anaAggDays(rows);
    assert.strictEqual(days.length, 16);
    assert.strictEqual(days[0].date, '2026-11-01');
    assert.strictEqual(days[0].ops, 5662);
    assert.strictEqual(days[days.length - 1].date, '2026-11-16');
    assert.strictEqual(days[days.length - 1].ops, 7431);
    const hours = api.anaAggHours(rows);
    assert.strictEqual(hours.length, 24);
    assert.strictEqual(sum(hours, h => h.ops), 112340);
    assert.ok(hours.some(h => h.ops > 0));
});

test('analyseLogisticOperation: смены день и ночь', () => {
    assert.ok(serviceHtml.includes('⚖️ Взвешенные операции: день и ночь'), 'нет диаграммы взвешенных операций по сменам');
    assert.ok(serviceHtml.includes('📦 Штуки: день и ночь (в общем)'), 'нет диаграммы штук по сменам');
    assert.ok(!serviceHtml.includes('⚡ КПД: день и ночь'), 'остался график КПД по сменам');
    assert.ok(!serviceHtml.includes('🗂️ Категории: день и ночь'), 'остался график категорий по сменам');
    assert.ok(serviceHtml.includes('s.day.w, s.night.w'), 'взвешенные операции не участвуют в диаграмме смен');
    assert.ok(serviceHtml.includes('s.day.qty, s.night.qty'), 'штуки не участвуют в диаграмме смен');
    assert.strictEqual(api.anaShiftOfSlot(449), 'night');
    assert.strictEqual(api.anaShiftOfSlot(450), 'day');
    assert.strictEqual(api.anaShiftOfSlot(1169), 'day');
    assert.strictEqual(api.anaShiftOfSlot(1170), 'night');
    const s = api.anaShiftStats(rows);
    assert.strictEqual(s.day.ops + s.night.ops, sum(rows, r => r.ops));
    assert.strictEqual(s.day.ops, 65284);
    assert.strictEqual(s.night.ops, 47056);
    assert.strictEqual(s.day.slots, 17960);
    assert.strictEqual(s.night.slots, 9633);
    assert.ok(close(s.day.kpd, s.day.w / s.day.slots));
    assert.ok(close(s.night.kpd, s.night.w / s.night.slots));
    assert.strictEqual(s.day.ops + s.night.ops, sum(Array.from(s.day.cats.values()), c => c.ops) + sum(Array.from(s.night.cats.values()), c => c.ops));
    assert.ok(close(s.day.qty + s.night.qty, sum(rows, r => r.qty)));
    assert.ok(close(s.day.qty, sum(Array.from(s.day.cats.values()), c => c.qty)));
    assert.ok(close(s.night.qty, sum(Array.from(s.night.cats.values()), c => c.qty)));
});

test('analyseLogisticOperation: парсер сообщает об отсутствии обязательных колонок', () => {
    assert.throws(() => api.anaParseWorkbook([['Примененные фильтры'], [], ['Дата']]), /обязательные колонки/);
});

/* ---------- КПД ---------- */

test('analyseLogisticOperation: КПД остаётся взвешенными операциями на слот', () => {
    const rows = slotRows('ONE', 'replenishment', 30, 3);
    const employee = api.anaAggEmployees(rows)[0];
    assert.strictEqual(employee.slots, 30);
    assert.ok(close(employee.w, 180));
    assert.ok(close(employee.kpd, 6), `КПД ${employee.kpd}`);
});

test('analyseLogisticOperation: порог слотов фильтрует сотрудников', () => {
    const list = [
        { ldap: 'A', slots: 5 },
        { ldap: 'B', slots: 29 },
        { ldap: 'C', slots: 30 },
        { ldap: 'D', slots: 120 },
    ];
    assert.deepStrictEqual(api.anaFilterBySlots(list, 30).map(e => e.ldap), ['C', 'D']);
    assert.deepStrictEqual(api.anaFilterBySlots(list, 0).map(e => e.ldap), ['A', 'B', 'C', 'D']);
    assert.deepStrictEqual(api.anaFilterBySlots(list, 5).map(e => e.ldap), ['A', 'B', 'C', 'D']);
    assert.deepStrictEqual(api.anaFilterBySlots(list, 100).map(e => e.ldap), ['D']);
});

test('analyseLogisticOperation: сортировка топов по КПД, взвешенным операциям, слотам и операциям', () => {
    const list = [
        { fio: 'A', kpd: 3, slots: 10, ops: 100, w: 150 },
        { fio: 'B', kpd: 5, slots: 5, ops: 300, w: 300 },
        { fio: 'C', kpd: 1, slots: 50, ops: 200, w: 900 },
    ];
    assert.deepStrictEqual(list.slice().sort(api.anaRankComparator('kpd')).map(e => e.fio), ['B', 'A', 'C']);
    assert.deepStrictEqual(list.slice().sort(api.anaRankComparator('w')).map(e => e.fio), ['C', 'B', 'A']);
    assert.deepStrictEqual(list.slice().sort(api.anaRankComparator('slots')).map(e => e.fio), ['C', 'A', 'B']);
    assert.deepStrictEqual(list.slice().sort(api.anaRankComparator('ops')).map(e => e.fio), ['B', 'C', 'A']);
    assert.deepStrictEqual(list.slice().sort(api.anaRankComparator()).map(e => e.fio), ['B', 'A', 'C']);
    assert.strictEqual(api.ANA_RANK_SORT_DEFAULT, 'kpd');
    assert.deepStrictEqual([...api.ANA_SLOT_FILTER_OPTIONS], [0, 5, 10, 20, 30, 50, 100]);
    assert.strictEqual(api.ANA_SLOT_FILTER_DEFAULT, 0);
});

/* ---------- Вкладка «Продуктивность» ---------- */

test('analyseLogisticOperation: часы из слотов и ставка без данных дают «—»', () => {
    assert.strictEqual(api.anaSlotsToHours(6), 1);
    assert.strictEqual(api.anaSlotsToHours(30), 5);
    assert.strictEqual(api.anaRate(180, 5), 36);
    assert.strictEqual(api.anaRate(10, 0), null);
    assert.strictEqual(api.anaFmtRate(null), '—');
    assert.strictEqual(api.anaFmtRate(21.0419), '21,04');
    assert.strictEqual(api.ANA_RATE_MIN_SLOTS, 3);
});

test('analyseLogisticOperation: продуктивность по контексту pallet и sku на сотрудника', () => {
    const sku = api.anaAggEmployees(slotRows('SKU', 'replenishment', 30, 6, { ctx: 'sku', qty: 30 }))[0];
    assert.strictEqual(sku.skuSlots, 30);
    assert.strictEqual(sku.skuOps, 180);
    assert.ok(close(sku.skuQty, 900));
    assert.ok(close(sku.skuHours, 5));
    assert.ok(close(sku.piecesPerHour, 180), `штук/ч ${sku.piecesPerHour}`);
    assert.ok(close(sku.skuPerHour, 36), `SKU/ч ${sku.skuPerHour}`);
    assert.strictEqual(sku.palletsPerHour, null);

    const pallet = api.anaAggEmployees(slotRows('PALLET', 'warehouse', 12, 2, { ctx: 'pallet' }))[0];
    assert.strictEqual(pallet.palletSlots, 12);
    assert.strictEqual(pallet.palletOps, 24);
    assert.ok(close(pallet.palletHours, 2));
    assert.ok(close(pallet.palletsPerHour, 12), `паллет/ч ${pallet.palletsPerHour}`);
    assert.strictEqual(pallet.piecesPerHour, null);
});

test('analyseLogisticOperation: мало слотов контекста - «—», а не 0', () => {
    const employee = api.anaAggEmployees(slotRows('MIN', 'replenishment', 2, 1, { ctx: 'sku', qty: 10 }))[0];
    assert.strictEqual(employee.skuSlots, 2);
    assert.strictEqual(employee.piecesPerHour, null);
    assert.strictEqual(employee.skuPerHour, null);
    assert.strictEqual(employee.palletsPerHour, null);
    assert.strictEqual(employee.assembliesPerHour, null);
});

test('analyseLogisticOperation: сборки и сборки/час в контекстной продуктивности', () => {
    const build = api.anaAggEmployees(slotRows('BUILD', 'orders', 30, 4, { ctx: 'sku', op: 'Сборка в заказы' }))[0];
    assert.strictEqual(build.assemblySlots, 30);
    assert.strictEqual(build.assemblyOps, 120);
    assert.ok(close(build.assemblyHours, 5));
    assert.ok(close(build.assembliesPerHour, 24), `сборок/ч ${build.assembliesPerHour}`);
    assert.strictEqual(build.palletsPerHour, null);
    /* Сборка в заказы идёт в контексте sku, поэтому её часы - часть SKU-часов. */
    assert.strictEqual(build.skuSlots, 30);

    /* Другие операции контекста order в сборки не попадают. */
    const other = api.anaAggEmployees(slotRows('OTHER', 'orders', 30, 4, { ctx: 'order' }))[0];
    assert.strictEqual(other.assemblyOps, 0);
    assert.strictEqual(other.assembliesPerHour, null);
    /* Но сам контекст order считается отдельно: заказы и заказы/ч. */
    assert.strictEqual(other.orderSlots, 30);
    assert.strictEqual(other.orderOps, 120);
    assert.ok(close(other.orderHours, 5));
    assert.ok(close(other.ordersPerHour, 24), `заказов/ч ${other.ordersPerHour}`);
    assert.strictEqual(other.palletsPerHour, null);

    const few = api.anaAggEmployees(slotRows('BUILD2', 'orders', 2, 4, { ctx: 'sku', op: 'Сборка в заказы' }))[0];
    assert.strictEqual(few.assembliesPerHour, null);

    const fewOrder = api.anaAggEmployees(slotRows('ORDER2', 'orders', 2, 4, { ctx: 'order' }))[0];
    assert.strictEqual(fewOrder.ordersPerHour, null);

    const t = api.anaTotals(validRows);
    assert.strictEqual(t.assemblyOps, 16750);
    assert.ok(close(t.assemblyHours, 714), `часы сборок ${t.assemblyHours}`);
    assert.ok(close(t.assembliesPerHour, 23.4594, 0.001), `сборок/ч ${t.assembliesPerHour}`);
    assert.ok(t.assemblyHours < t.activeHours, 'часы сборок - часть активных часов');

    assert.strictEqual(t.orderOps, 10140);
    assert.ok(close(t.orderHours, 975.833, 0.01), `часы заказов ${t.orderHours}`);
    assert.ok(close(t.ordersPerHour, 10.391, 0.001), `заказов/ч ${t.ordersPerHour}`);
    assert.ok(t.orderHours < t.activeHours, 'часы заказов - часть активных часов');
});

test('analyseLogisticOperation: часы контекстов пересекаются и не суммируются', () => {
    const rows = [
        { date: '2026-11-01', slot: 0, hour: 0, op: 'Подбор', ctx: 'pallet', ops: 1, qty: 1, ldap: 'BOTH', fio: 'Оба', pos: '', dept: 'тест', deptType: 'логистика', cat: 'warehouse', w: 1 },
        { date: '2026-11-01', slot: 0, hour: 0, op: 'Подбор', ctx: 'sku', ops: 1, qty: 3, ldap: 'BOTH', fio: 'Оба', pos: '', dept: 'тест', deptType: 'логистика', cat: 'warehouse', w: 2 },
        { date: '2026-11-01', slot: 10, hour: 0, op: 'Подбор', ctx: 'sku', ops: 1, qty: 4, ldap: 'BOTH', fio: 'Оба', pos: '', dept: 'тест', deptType: 'логистика', cat: 'warehouse', w: 2 },
    ];
    const employee = api.anaAggEmployees(rows)[0];
    assert.strictEqual(employee.slots, 2);
    assert.strictEqual(employee.palletSlots, 1);
    assert.strictEqual(employee.skuSlots, 2);
    assert.ok(close(employee.activeHours, 2 / 6));
    assert.ok(close(employee.palletHours, 1 / 6));
    assert.ok(close(employee.skuHours, 2 / 6));
    assert.ok(employee.palletHours + employee.skuHours > employee.activeHours);
});

test('analyseLogisticOperation: ставки и занятость на образце', () => {
    const t = api.anaTotals(validRows);
    assert.strictEqual(t.palletOps, 15834);
    assert.ok(close(t.palletHours, 752.5));
    assert.ok(close(t.palletsPerHour, 21.0419, 0.001), `паллет/ч ${t.palletsPerHour}`);
    assert.ok(close(t.skuQty, 1485151.243, 0.01));
    assert.ok(close(t.skuOps, 86366));
    assert.ok(close(t.skuHours, 3205.833, 0.01));
    assert.ok(close(t.piecesPerHour, 463.26, 0.01), `штук/ч ${t.piecesPerHour}`);
    assert.ok(close(t.skuPerHour, 26.94, 0.01), `SKU/ч ${t.skuPerHour}`);
    assert.ok(close(t.activeHours, 4598.833, 0.01));
    assert.strictEqual(t.employeeDays, 1637);
    assert.ok(close(t.avgHoursPerDay, 2.8093, 0.001));
    assert.ok(t.palletHours + t.skuHours < t.activeHours, 'контекстные часы не равны общим активным часам');
});

test('analyseLogisticOperation: отдельная вкладка контекстной продуктивности со своей таблицей', () => {
    const tab = api.ANA_TABS.find(t => t.key === 'productivity');
    assert.ok(tab, 'нет вкладки «Контекстная продуктивность»');
    assert.strictEqual(tab.kind, 'productivity');
    assert.strictEqual(tab.label, '📦 Контекстная продуктивность');
    assert.strictEqual(api.ANA_TABS.find(t => t.key === 'rating').label, '📊 Продуктивность');
    const keys = api.anaTableColumns('productivity').map(c => c.key);
    assert.deepStrictEqual(keys, [
        'fio', 'dept', 'palletOps', 'palletsPerHour',
        'skuQty', 'piecesPerHour', 'skuPerHour',
        'assemblyOps', 'assembliesPerHour',
        'orderOps', 'ordersPerHour', 'activeHours',
    ]);
    const cols = api.anaTableColumns('productivity');
    assert.strictEqual(cols.find(c => c.key === 'palletOps').emptyOn, 'palletSlots');
    assert.strictEqual(cols.find(c => c.key === 'skuQty').emptyOn, 'skuSlots');
    assert.strictEqual(cols.find(c => c.key === 'assemblyOps').emptyOn, 'assemblySlots');
    assert.strictEqual(cols.find(c => c.key === 'orderOps').emptyOn, 'orderSlots');
    assert.ok(!cols.some(c => c.key === 'palletHours'), 'остался столбец часов pallet');
    assert.ok(!cols.some(c => c.key === 'skuHours'), 'остался столбец часов SKU');
    assert.ok(!cols.some(c => c.key === 'avgHoursPerDay'), 'остался столбец Ч/день');
    assert.ok(cols.filter(c => c.rate).length === 5, 'должны быть пять ставок');
    assert.ok(serviceHtml.includes('Таблица по контекстам'), 'нет собственной таблицы вкладки');
    assert.ok(serviceHtml.includes('Паллет/час = Паллеты / Активные часы паллетных строк'), 'нет формулы паллет/час');
    assert.ok(serviceHtml.includes('Штук SKU/час = Штуки SKU / Активные часы SKU-строк'), 'нет формулы штук SKU/час');
    assert.ok(serviceHtml.includes('Сборки/час = Сборки / Активные часы строк «Сборка в заказы»'), 'нет формулы сборок/час');
    assert.ok(serviceHtml.includes('Заказы/час = Заказы / Активные часы строк контекста order'), 'нет формулы заказов/час');
});

test('analyseLogisticOperation: топ-10 и сортировки на вкладке контекстной продуктивности', () => {
    assert.deepStrictEqual(api.ANA_PROD_SORT_OPTIONS.map(o => o[0]), [
        'palletsPerHour', 'skuPerHour', 'piecesPerHour', 'palletOps', 'skuQty',
        'assemblyOps', 'assembliesPerHour', 'orderOps', 'ordersPerHour',
    ]);
    assert.strictEqual(api.anaProdSortLabel('palletsPerHour'), 'паллетам/ч');
    assert.strictEqual(api.anaProdSortLabel('piecesPerHour', true), 'Шт. SKU/ч');
    assert.strictEqual(api.anaProdSortLabel('assembliesPerHour'), 'сборкам/ч');
    assert.strictEqual(api.anaProdSortLabel('ordersPerHour'), 'заказам/ч');
    const list = [
        { fio: 'A', palletsPerHour: 10, piecesPerHour: 100, skuPerHour: 5, palletOps: 100, skuQty: 1000, assemblyOps: 50, assembliesPerHour: 12, orderOps: 40, ordersPerHour: 8, activeHours: 10 },
        { fio: 'B', palletsPerHour: null, piecesPerHour: null, skuPerHour: null, palletOps: 0, skuQty: 0, assemblyOps: 0, assembliesPerHour: null, orderOps: 0, ordersPerHour: null, activeHours: 0 },
        { fio: 'C', palletsPerHour: 20, piecesPerHour: 50, skuPerHour: 8, palletOps: 50, skuQty: 2000, assemblyOps: 80, assembliesPerHour: 6, orderOps: 90, ordersPerHour: 11, activeHours: 5 },
    ];
    const sorted = mode => list.slice().sort(api.anaProdComparator(mode)).map(e => e.fio);
    assert.deepStrictEqual(sorted('palletsPerHour'), ['C', 'A', 'B']);
    assert.deepStrictEqual(sorted('skuPerHour'), ['C', 'A', 'B']);
    assert.deepStrictEqual(sorted('piecesPerHour'), ['A', 'C', 'B']);
    assert.deepStrictEqual(sorted('palletOps'), ['A', 'C', 'B']);
    assert.deepStrictEqual(sorted('skuQty'), ['C', 'A', 'B']);
    assert.deepStrictEqual(sorted('assemblyOps'), ['C', 'A', 'B']);
    assert.deepStrictEqual(sorted('assembliesPerHour'), ['A', 'C', 'B']);
    assert.deepStrictEqual(sorted('orderOps'), ['C', 'A', 'B']);
    assert.deepStrictEqual(sorted('ordersPerHour'), ['C', 'A', 'B']);
    assert.ok(serviceHtml.includes('anaRenderProdTopCards'), 'нет топ-10 на вкладке');
    assert.ok(serviceHtml.includes('data-prod-sort'), 'нет кнопок сортировки топ-10');
    /* Порог слотов из топ-10 вкладки «Продуктивность» доступен и здесь. */
    assert.ok(/const slotFilterBlock = \(isEmployee \|\| isProductivity\)/.test(serviceHtml), 'нет порога слотов в топ-10 контекстной продуктивности');
    assert.ok(/tab\.key === 'productivity'[\s\S]{0,200}anaRenderProdTopFromFilters/.test(serviceHtml), 'порог слотов не перерисовывает топ-10');
    /* Фильтры топ-10 - тот же набор, что над топ-10 вкладки «Продуктивность». */
    assert.ok(serviceHtml.includes('function anaTopFiltersHtml'), 'нет панели фильтров в топ-10');
    assert.ok(serviceHtml.includes('tp-${tab.key}-${f}'), 'у фильтров топ-10 нет своего префикса');
    assert.ok(serviceHtml.includes('function anaSetTopFilter'), 'фильтры топ-10 не подключены');
    assert.ok(serviceHtml.includes('function anaResetTopFilters'), 'нет сброса фильтров топ-10');
    assert.ok(/\$\{anaTopFiltersHtml\(tab\)\}/.test(serviceHtml), 'панель фильтров не выведена в блоке топ-10');
    /* Топ-10 и таблица фильтруются независимо. */
    assert.ok(serviceHtml.includes('let anaTopFilters'), 'нет отдельного состояния фильтров топ-10');
    assert.ok(/function anaSetTopFilter[\s\S]{0,300}anaTopFilters\[key\]\[field\]/.test(serviceHtml), 'фильтры топ-10 пишут не в своё состояние');
    assert.ok(!/function anaSetTopFilter\(key, field, value\) \{\s*anaSetTableFilter/.test(serviceHtml), 'фильтры топ-10 всё ещё общие с таблицей');
    assert.ok(/anaRefreshFilterOptions\(key, 'top'\)/.test(serviceHtml), 'опции фильтров топ-10 не обновляются');
    assert.ok(/const isTop = kind === 'top'/.test(serviceHtml), 'нет режима top в anaRefreshFilterOptions');
    assert.ok(/function anaRenderTab[\s\S]{0,400}anaRenderProdTopFromFilters/.test(serviceHtml), 'топ-10 не рендерится при открытии вкладки');
    assert.ok(!/function anaRenderTable[\s\S]{0,400}anaRenderProdTopCards/.test(serviceHtml), 'таблица всё ещё перерисовывает топ-10');
    assert.ok(serviceHtml.includes('Динамика по дням по категориям операций'), 'нет гистограммы категорий в карточке');
    assert.ok(/stack:\s*'ops'/.test(serviceHtml), 'динамика по дням не разбита по категориям');
});

test('analyseLogisticOperation: медианы удалены из сервиса и справки', () => {
    assert.ok(!/медиан/i.test(serviceHtml), 'в сервисе осталось упоминание медианы');
    assert.ok(!serviceHtml.includes('anaMedian'), 'остался расчёт медианы');
    assert.ok(!serviceHtml.includes('anaKpdMedians'), 'остался расчёт медиан КПД');
    assert.ok(!serviceHtml.includes('anaApplyKpdReference'), 'остался справочный ориентир');
    assert.ok(serviceHtml.includes('<h4>🧮 Формулы</h4>'), 'формулы не вынесены в отдельный блок');
    assert.ok(serviceHtml.includes('📖 Как работает сервис'), 'справка не переименована');
});
