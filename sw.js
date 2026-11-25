/* Service worker портала: офлайн-работа, быстрый запуск и управление обновлениями */

/* Единый источник версии - общая конфигурация портала.
   Имя кэша выводится из версии, поэтому выпуск новой версии
   автоматически создаёт новый кэш, а старый удаляется при активации. */
importScripts('./shared/js/config.js');

/** Текущая версия, с которой работает Service Worker. */
let activeVersion = self.PORTAL_VERSION;

/** Имя основного кэша оболочки. */
let CACHE = cacheName(activeVersion);

/** Имя runtime-кэша для динамических ресурсов. */
let RUNTIME = runtimeName(activeVersion);

/** Промис текущей проверки версии (защита от параллельных запросов). */
let versionCheckPromise = null;

/** Метка времени последней проверки версии. */
let lastVersionCheck = 0;

/** Интервал между проверками версии, мс. */
const VERSION_CHECK_INTERVAL = 60 * 1000;

/**
 * Собирает имя основного кэша из версии.
 * @param {string} version - Версия портала.
 * @returns {string} Имя кэша.
 */
function cacheName(version) {
    return `lemanapro-portal-${String(version).replace(/\./g, '-')}`;
}

/**
 * Собирает имя runtime-кэша из версии.
 * @param {string} version - Версия портала.
 * @returns {string} Имя кэша.
 */
function runtimeName(version) {
    return `lemanapro-portal-runtime-${String(version).replace(/\./g, '-')}`;
}

// Оболочка (кэшируется при установке)
const SHELL = [
    './',
    './index.html',
    './manifest.webmanifest',
    './icon.svg',
    './icon-192x192.png',
    './icon-512x512.png',
    // Общие модули (единые для всех сервисов)
    './shared/css/common.css',
    './shared/js/config.js',
    './shared/js/pwa-update.js',
    './shared/js/idb.js',
    './shared/js/common.js',
    './shared/js-lib/xlsx.full.min.js',
    './shared/js-lib/chart.umd.min.js',
    // Сервисы (оболочки)
    './Services/avs_control/index.html',
    './Services/main_address/index.html',
    './Services/analyseLogisticOperation/index.html',
    './Services/pallet_rotation/index.html',
    './Services/analyse_cells/index.html',
    './Services/corrections/index.html',
    './Services/anomalies/index.html',
    './Services/operativInventory/index.html',
    './Services/operativInventoryForGoods/index.html'
];

// Паттерны для динамического кэширования (runtime)
const RUNTIME_PATTERNS = [
    /\/shared\//, // всё из shared/ (css, js, js-lib)
];

/**
 * Кэширует оболочку. Запросы идут с cache: 'reload', чтобы не получить
 * 304 Not Modified: cache.add()/cache.put() принимают только полный ответ.
 * @returns {Promise<void>}
 */
async function cacheShell() {
    const cache = await caches.open(CACHE);
    await Promise.allSettled(SHELL.map(async u => {
        try {
            const res = await fetch(u, { cache: 'reload' });
            if (!res.ok) throw new Error('HTTP ' + res.status);
            await cache.put(u, res);
        } catch (err) {
            console.warn('[SW] Не удалось закэшировать:', u, err);
        }
    }));
}

/**
 * Сообщает открытым страницам, что доступна новая версия.
 * Страница покажет предложение обновиться; переключение Worker'а
 * выполняется только после подтверждения пользователя.
 * @param {string} [version] - Версия, которая стала доступна.
 * @returns {Promise<void>}
 */
async function notifyClientsUpdate(version) {
    const clients = await self.clients.matchAll({ includeUncontrolled: true, type: 'window' });
    clients.forEach(client => {
        client.postMessage({ type: 'UPDATE_READY', version: version || activeVersion });
    });
}

/**
 * Читает актуальную версию портала из конфигурации на сервере.
 * @returns {Promise<string|null>} Версия или null, если определить не удалось.
 */
async function readRemoteVersion() {
    try {
        const res = await fetch('./shared/js/config.js', { cache: 'no-store' });
        if (!res || !res.ok) return null;
        const text = await res.text();
        const match = text.match(/version:\s*'([^']+)'/);
        return match ? match[1] : null;
    } catch (err) {
        return null;
    }
}

/**
 * Проверяет, не появилась ли новая версия портала.
 * При обнаружении переключает кэши на новую версию и уведомляет страницы.
 * @param {boolean} [force=false] - Проверить без учёта интервала.
 * @returns {Promise<string|null>} Новая версия или null.
 */
async function checkForNewVersion(force = false) {
    const now = Date.now();
    if (!force && now - lastVersionCheck < VERSION_CHECK_INTERVAL) return null;
    lastVersionCheck = now;
    if (versionCheckPromise) return versionCheckPromise;

    versionCheckPromise = (async () => {
        try {
            const remote = await readRemoteVersion();
            if (!remote || remote === activeVersion) return null;

            const previousCaches = [CACHE, RUNTIME];
            activeVersion = remote;
            CACHE = cacheName(activeVersion);
            RUNTIME = runtimeName(activeVersion);

            await cacheShell();
            const keys = await caches.keys();
            await Promise.all(
                keys
                    .filter(key => key !== CACHE && key !== RUNTIME)
                    .map(key => caches.delete(key))
            );
            await notifyClientsUpdate(remote);
            return remote;
        } catch (err) {
            console.warn('[SW] Проверка версии не удалась:', err);
            return null;
        } finally {
            versionCheckPromise = null;
        }
    })();

    return versionCheckPromise;
}

self.addEventListener('install', e => {
    // Новый Worker не активируется автоматически: переключение происходит
    // после подтверждения пользователя на странице.
    e.waitUntil(cacheShell().then(() => notifyClientsUpdate(activeVersion)));
});

self.addEventListener('activate', e => {
    e.waitUntil(
        caches.keys()
            .then(keys => Promise.all(keys.filter(k => k !== CACHE && k !== RUNTIME).map(k => caches.delete(k))))
            .then(() => self.clients.claim())
            .then(() => checkForNewVersion(true))
    );
});

self.addEventListener('message', event => {
    const data = event.data || {};
    if (data.type === 'SKIP_WAITING') {
        self.skipWaiting();
        return;
    }
    if (data.type === 'GET_VERSION' && event.source) {
        event.source.postMessage({ type: 'VERSION', version: activeVersion });
        return;
    }
    if (data.type === 'APPLY_UPDATE') {
        event.waitUntil(checkForNewVersion(true).then(version => {
            if (event.source) event.source.postMessage({ type: 'APPLIED', version: version || activeVersion });
        }));
        return;
    }
    if (data.type === 'CHECK_VERSION') {
        event.waitUntil(checkForNewVersion(true));
    }
});

function isRuntime(url) {
    return RUNTIME_PATTERNS.some(p => p.test(url));
}

/**
 * Ищет в кэше страницу для навигации: сначала точный URL (без учёта query),
 * затем index.html того же сервиса.
 * @param {Request} req - Навигационный запрос.
 * @returns {Promise<Response|undefined>}
 */
async function matchCachedNavigation(req) {
    const exact = await caches.match(req, { ignoreSearch: true });
    if (exact) return exact;

    try {
        const url = new URL(req.url);
        // Отбрасываем пустые сегменты, чтобы /Services/имя/ и /Services/имя
        // давали один и тот же путь.
        const parts = url.pathname.split('/').filter(Boolean);
        const servicesIndex = parts.indexOf('Services');
        const serviceName = parts[servicesIndex + 1];
        if (servicesIndex === -1 || !serviceName) return undefined;

        // /Services/<имя>/любой/путь → /Services/<имя>/index.html
        const serviceUrl = new URL(`/Services/${serviceName}/index.html`, url.origin).toString();
        return await caches.match(serviceUrl, { ignoreSearch: true });
    } catch (err) {
        return undefined;
    }
}

self.addEventListener('fetch', e => {
    const req = e.request;
    const url = req.url;

    if (req.method !== 'GET') return;
    // Только same-origin
    try {
        if (new URL(url).origin !== location.origin) return;
    } catch (err) { return; }

    // Навигация: сеть → кэш (точная страница или index.html сервиса) → портал
    if (req.mode === 'navigate') {
        e.respondWith(
            Promise.resolve(checkForNewVersion())
                .then(() => fetch(req, { cache: 'reload' }))
                .then(res => {
                    if (!res || !res.ok) throw new Error('Сеть вернула ошибку ' + (res ? res.status : ''));
                    const copy = res.clone();
                    caches.open(CACHE).then(c => c.put(req, copy));
                    return res;
                })
                .catch(async () => {
                    const cached = await matchCachedNavigation(req);
                    if (cached) return cached;
                    return (await caches.match('./index.html')) || Response.error();
                })
        );
        return;
    }

    // Статические ресурсы shared/: сеть → кэш (network-first с докешированием)
    if (isRuntime(url)) {
        e.respondWith(
            fetch(req, { cache: 'reload' })
                .then(res => {
                    if (!res || !res.ok) throw new Error('Сеть вернула ошибку ' + (res ? res.status : ''));
                    const copy = res.clone();
                    caches.open(RUNTIME).then(c => c.put(req, copy));
                    return res;
                })
                .catch(() => caches.match(req).then(hit => hit || new Response('', { status: 504, statusText: 'Offline' })))
        );
        return;
    }

    // Остальное: кэш → сеть с докешированием успешных ответов
    e.respondWith(
        caches.match(req).then(hit => hit || fetch(req, { cache: 'reload' }).then(res => {
            if (res && res.ok) {
                const copy = res.clone();
                caches.open(CACHE).then(c => c.put(req, copy));
            }
            return res;
        }))
    );
});
