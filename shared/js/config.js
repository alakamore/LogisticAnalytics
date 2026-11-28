/* =====================================================
   ЛОГИСТИКА ANALYTICS - единая конфигурация портала
   ------------------------------------------------------------
   Этот файл - единственный источник версии портала и версий
   сервисов. Он используется главной страницей, сервисами и
   Service Worker. При выпуске обновляется только PORTAL_CONFIG.
   ===================================================== */

self.PORTAL_CONFIG = Object.freeze({
    version: '1.0.8',
    releaseDate: '28.11.2026',
    services: Object.freeze([
        Object.freeze({
            id: 'avs',
            title: 'AVS Контроль',
            version: '1.0.0',
            icon: '🚨',
            url: 'Services/avs_control/index.html',
            desc: 'Анализ avs-товаров для вывода из складских зон хранения: просроченные, критичные и нормальные остатки.',
            accent: '#ed8936',
            grad: 'linear-gradient(135deg,#f6ad55,#ed8936)',
            glow: 'rgba(237,137,54,.5)',
        }),
        Object.freeze({
            id: 'restock',
            title: 'Пополнение Основных Адресов',
            version: '1.0.0',
            icon: '📋',
            url: 'Services/main_address/index.html',
            desc: 'Анализ потребности пополнения основных адресов из других мест хранения, поиск источников пополнения.',
            accent: '#d69e2e',
            grad: 'linear-gradient(135deg,#f6e05e,#d69e2e)',
            glow: 'rgba(214,158,46,.5)',
        }),
        Object.freeze({
            id: 'pallet_rotation',
            title: 'Анализ Ротации Паллет',
            version: '1.0.0',
            icon: '📆',
            url: 'Services/pallet_rotation/index.html',
            desc: 'Выявление паллет с нарушением ротации: определение долгосрочного хранения на складе.',
            accent: '#b45309',
            grad: 'linear-gradient(135deg,#f6ad55,#b45309)',
            glow: 'rgba(180,83,9,.5)',
        }),
        Object.freeze({
            id: 'cells',
            title: 'Анализ Ячеек Склада',
            version: '1.0.0',
            icon: '📦️',
            url: 'Services/analyse_cells/index.html',
            desc: 'Мониторинг статусов, блокировок и доступности адресов хранения, коэффициент использования зон.',
            accent: '#38a169',
            grad: 'linear-gradient(135deg,#68d391,#38a169)',
            glow: 'rgba(56,161,105,.5)',
        }),
        Object.freeze({
            id: 'corrections',
            title: 'Анализ Коррекций',
            version: '1.0.1',
            icon: '🧾',
            url: 'Services/corrections/index.html',
            desc: 'Анализ списаний под разными кодами аналитики, учёта инвентаризаций и выявление аномалий в проведённых ведомостях.',
            accent: '#c53030',
            grad: 'linear-gradient(135deg,#fc8181,#c53030)',
            glow: 'rgba(197,48,48,.5)',
        }),
        Object.freeze({
            id: 'anomalies',
            title: 'Аномалии склада',
            version: '1.0.0',
            icon: '⚠️',
            url: 'Services/anomalies/index.html',
            desc: 'Слияние отчёта «Аномалии запаса» с фактическим наличием на складе для оперативного поиска и анализа.',
            accent: '#d97706',
            grad: 'linear-gradient(135deg,#f6ad55,#d97706)',
            glow: 'rgba(217,119,6,.5)',
        }),
        Object.freeze({
            id: 'operativInventory',
            title: 'Оперативные инвентаризации по адресу',
            version: '1.0.0',
            icon: '🧮',
            url: 'Services/operativInventory/index.html',
            desc: 'Контроль оперативных инвентаризаций по адресу: статусы ведомостей, недостачи и излишки, длительность',
            accent: '#ca8a04',
            grad: 'linear-gradient(135deg,#fde047,#ca8a04)',
            glow: 'rgba(202,138,4,.5)',
        }),
        Object.freeze({
            id: 'operativInventoryForGoods',
            title: 'Оперативные инвентаризации по товару',
            version: '1.0.0',
            icon: '📝',
            url: 'Services/operativInventoryForGoods/index.html',
            desc: 'Контроль инвентаризаций по товару: количество и суммы, недостачи и излишки, зоны, сотрудники и коррекции.',
            accent: '#a16207',
            grad: 'linear-gradient(135deg,#facc15,#a16207)',
            glow: 'rgba(161,98,7,.5)',
        }),
        Object.freeze({
            id: 'analyse_logistic_operation',
            title: 'Аналитика логистических операций',
            version: '1.0.7',
            icon: '📈',
            url: 'Services/analyseLogisticOperation/index.html',
            desc: 'Динамика операций по дням и часам, продуктивность по паллетам и штукам SKU, КПД, топ-10, доли по категориям и контекстам.',
            accent: '#14b8a6',
            grad: 'linear-gradient(135deg,#5eead4,#0f8f7a)',
            glow: 'rgba(20,184,166,.5)',
        }),
    ]),
});

self.PORTAL_VERSION = self.PORTAL_CONFIG.version;
self.PORTAL_RELEASE_DATE = self.PORTAL_CONFIG.releaseDate;
