/* =====================================================
   ЛОГИСТИКА ANALYTICS - управление обновлениями PWA
   ------------------------------------------------------------
   Единый модуль для главной страницы и всех сервисов.
   Показывает уведомление «Доступна новая версия», применяет
   обновление по подтверждению пользователя и перезагружает
   страницу после смены Service Worker.
   ===================================================== */

(function () {
    'use strict';

    const STYLES = `
        #pwa-update {
            display: none;
            position: fixed;
            top: 20px;
            right: 20px;
            align-items: center;
            gap: 12px;
            max-width: min(420px, calc(100vw - 32px));
            padding: 12px 14px;
            border: 1px solid rgba(255,193,7,.65);
            border-radius: 10px;
            background: #1a202c;
            color: #fff;
            box-shadow: 0 10px 30px rgba(0,0,0,.3);
            z-index: 1100;
            font-size: .88em;
        }
        #pwa-update.show { display: flex; }
        #pwa-update-text { flex: 1; line-height: 1.4; }
        #pwa-update-actions { display: flex; gap: 8px; flex-shrink: 0; }
        #pwa-update-actions button {
            border: 0;
            border-radius: 6px;
            padding: 7px 10px;
            cursor: pointer;
            font-weight: 600;
            font-family: inherit;
        }
        #pwa-update-now { background: #ffc107; color: #1a202c; }
        #pwa-update-later { background: transparent; color: #cbd5e0; border: 1px solid #4a5568 !important; }
        @media (max-width: 768px) {
            #pwa-update { top: 12px; right: 12px; left: 12px; max-width: none; }
        }
    `;

    /**
     * Сравнивает две версии в формате semver.
     * @param {string} a - Первая версия.
     * @param {string} b - Вторая версия.
     * @returns {number} -1, если a < b; 0, если равны; 1, если a > b.
     */
    function compareVersions(a, b) {
        const pa = String(a || '').split('.').map(Number);
        const pb = String(b || '').split('.').map(Number);
        for (let i = 0; i < 3; i++) {
            const x = pa[i] || 0;
            const y = pb[i] || 0;
            if (x !== y) return x < y ? -1 : 1;
        }
        return 0;
    }

    /**
     * Форматирует версию для показа: 9.10.0 → v9.10.
     * @param {string} version - Версия в формате semver.
     * @returns {string} Подпись версии.
     */
    function formatVersion(version) {
        return `v${String(version || '').split('.').slice(0, 2).join('.')}`;
    }

    /** @type {string} Путь к этому скрипту на момент загрузки. */
    const SCRIPT_SRC = document.currentScript ? document.currentScript.src : '';

    /**
     * Определяет корневой путь портала относительно расположения этого скрипта.
     * @returns {string} Путь к корню портала с завершающим слэшем.
     */
    function portalRoot() {
        if (SCRIPT_SRC) {
            const path = new URL(SCRIPT_SRC).pathname;
            return path.replace(/shared\/js\/pwa-update\.js$/, '');
        }
        return '/';
    }

    /**
     * Добавляет стили уведомления (один раз на страницу).
     */
    function ensureStyles() {
        if (document.getElementById('pwa-update-styles')) return;
        const style = document.createElement('style');
        style.id = 'pwa-update-styles';
        style.textContent = STYLES;
        document.head.appendChild(style);
    }

    /**
     * Создаёт блок уведомления об обновлении.
     * @returns {HTMLElement|null} Контейнер уведомления.
     */
    function ensureBanner() {
        let banner = document.getElementById('pwa-update');
        if (banner) return banner;

        banner = document.createElement('div');
        banner.id = 'pwa-update';
        banner.setAttribute('role', 'status');
        banner.setAttribute('aria-live', 'polite');
        banner.innerHTML = `
            <span id="pwa-update-text">Доступна новая версия портала. Обновить сейчас?</span>
            <div id="pwa-update-actions">
                <button type="button" id="pwa-update-later">Позже</button>
                <button type="button" id="pwa-update-now">Обновить</button>
            </div>`;
        document.body.appendChild(banner);
        return banner;
    }

    /** @type {boolean} Флаг: модуль уже инициализирован на этой странице. */
    let initialized = false;

    /**
     * Инициализирует регистрацию Service Worker и управление обновлениями.
     */
    function initPwaUpdate() {
        if (initialized) return;
        if (!('serviceWorker' in navigator)) return;
        initialized = true;

        ensureStyles();
        const banner = ensureBanner();
        if (!banner) return;

        const updateText = document.getElementById('pwa-update-text');
        const updateNowButton = document.getElementById('pwa-update-now');
        const updateLaterButton = document.getElementById('pwa-update-later');

        let serviceWorkerRegistration = null;
        let waitingWorker = null;
        let updateAccepted = false;
        let reloadScheduled = false;

        /**
         * Показывает уведомление о доступном обновлении.
         * @param {string} [version] - Версия, которая установилась.
         */
        function showUpdateNotice(version) {
            if (!version) {
                banner.classList.add('show');
                return;
            }
            const comparison = compareVersions(version, self.PORTAL_VERSION);
            if (comparison === 0) {
                // Версии совпадают - обновление не требуется.
                banner.classList.remove('show');
                return;
            }
            if (updateText) {
                updateText.textContent = comparison < 0
                    ? 'Офлайн-версия портала устарела. Обновить и перезагрузить?'
                    : `Доступна новая версия портала ${formatVersion(version)}. Обновить сейчас?`;
            }
            banner.classList.add('show');
        }

        /**
         * Запрашивает у ожидающего Worker его версию.
         */
        function requestWaitingVersion() {
            const target = waitingWorker
                || (serviceWorkerRegistration && serviceWorkerRegistration.waiting);
            if (target) target.postMessage({ type: 'GET_VERSION' });
        }

        /**
         * Следит за установкой нового Worker.
         * @param {ServiceWorker} worker - Устанавливаемый Worker.
         */
        function trackInstallingWorker(worker) {
            if (!worker) return;
            worker.addEventListener('statechange', () => {
                if (worker.state === 'installed' && navigator.serviceWorker.controller) {
                    waitingWorker = worker;
                    showUpdateNotice();
                    requestWaitingVersion();
                }
            });
        }

        /**
         * Активирует ожидающий Worker по подтверждению пользователя.
         * Если нового Worker нет (изменилась только версия в конфигурации),
         * Worker обновляет кэши, а страница перезагружается.
         */
        function activateWaitingWorker() {
            const target = waitingWorker
                || (serviceWorkerRegistration && serviceWorkerRegistration.waiting)
                || (serviceWorkerRegistration && serviceWorkerRegistration.installing);
            updateAccepted = true;
            banner.classList.remove('show');

            if (target) {
                target.postMessage({ type: 'SKIP_WAITING' });
                return;
            }

            if (navigator.serviceWorker.controller) {
                navigator.serviceWorker.controller.postMessage({ type: 'APPLY_UPDATE' });
            }
            setTimeout(() => window.location.reload(), 1500);
        }

        if (updateNowButton) updateNowButton.addEventListener('click', activateWaitingWorker);
        if (updateLaterButton) {
            updateLaterButton.addEventListener('click', () => banner.classList.remove('show'));
        }

        navigator.serviceWorker.addEventListener('message', async event => {
            if (!event.data) return;
            if (event.data.type === 'VERSION') {
                if (event.data.version) showUpdateNotice(event.data.version);
                return;
            }
            if (event.data.type === 'APPLIED') {
                if (updateAccepted) window.location.reload();
                return;
            }
            if (event.data.type !== 'UPDATE_READY') return;
            if (updateAccepted) return; // Обновление уже подтверждено пользователем.
            if (!navigator.serviceWorker.controller) return; // Первая установка не требует обновления страницы.
            serviceWorkerRegistration = await navigator.serviceWorker.getRegistration();
            // waitingWorker может быть null: изменилась только версия в конфигурации.
            waitingWorker = serviceWorkerRegistration ? serviceWorkerRegistration.waiting : null;
            showUpdateNotice(event.data.version);
        });

        navigator.serviceWorker.addEventListener('controllerchange', () => {
            if (!updateAccepted) return;
            if (reloadScheduled) return;
            reloadScheduled = true;
            window.location.reload();
        });

        /**
         * Проверяет конфигурацию на сервере: если версия портала новее
         * загруженной, сообщает об обновлении.
         * @returns {Promise<void>}
         */
        async function checkPortalVersion() {
            try {
                const res = await fetch(`${portalRoot()}shared/js/config.js`, { cache: 'no-store' });
                if (!res || !res.ok) return;
                const text = await res.text();
                const match = text.match(/version:\s*'([^']+)'/);
                if (match && match[1] !== self.PORTAL_VERSION) {
                    showUpdateNotice(match[1]);
                }
            } catch (e) {
                // Офлайн или недоступный сервер - обновление не требуется.
            }
        }

        /**
         * Регистрирует Service Worker и включает периодическую проверку обновлений.
         */
        async function registerServiceWorker() {
            try {
                // URL Service Worker не меняется: обновление приходит через
                // проверку версии в конфигурации, а не через смену адреса.
                serviceWorkerRegistration = await navigator.serviceWorker.register(`${portalRoot()}sw.js`);

                if (serviceWorkerRegistration.waiting && navigator.serviceWorker.controller) {
                    waitingWorker = serviceWorkerRegistration.waiting;
                    showUpdateNotice();
                    requestWaitingVersion();
                }

                serviceWorkerRegistration.addEventListener('updatefound', () => {
                    trackInstallingWorker(serviceWorkerRegistration.installing);
                });

                const checkForUpdates = () => {
                    if (document.visibilityState !== 'visible') return;
                    serviceWorkerRegistration.update().catch(() => {});
                    if (navigator.serviceWorker.controller) {
                        navigator.serviceWorker.controller.postMessage({ type: 'GET_VERSION' });
                        navigator.serviceWorker.controller.postMessage({ type: 'CHECK_VERSION' });
                    }
                    checkPortalVersion();
                };
                setInterval(checkForUpdates, 15 * 60 * 1000);
                document.addEventListener('visibilitychange', checkForUpdates);
                checkPortalVersion();

                // Уже установленный Worker сообщает свою версию и проверяет новую.
                if (navigator.serviceWorker.controller) {
                    navigator.serviceWorker.controller.postMessage({ type: 'GET_VERSION' });
                    navigator.serviceWorker.controller.postMessage({ type: 'CHECK_VERSION' });
                }
            } catch (e) {
                console.warn('SW недоступен (нужен HTTPS или localhost):', e);
            }
        }

        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', registerServiceWorker);
        } else {
            registerServiceWorker();
        }
    }

    window.initPwaUpdate = initPwaUpdate;

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initPwaUpdate);
    } else {
        initPwaUpdate();
    }
})();
