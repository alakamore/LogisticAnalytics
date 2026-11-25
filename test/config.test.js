import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

/**
 * Загружает shared/js/config.js в изолированном окружении.
 * @returns {Object} Объект self с PORTAL_CONFIG и PORTAL_VERSION.
 */
function loadConfig() {
    const code = fs.readFileSync('./shared/js/config.js', 'utf8');
    const sandbox = {};
    new Function('self', code)(sandbox);
    return sandbox;
}

const config = loadConfig();
const portalConfig = config.PORTAL_CONFIG;
const packageJson = JSON.parse(fs.readFileSync('./package.json', 'utf8'));

/* ---------- Единая версия ---------- */

test('config: версия портала совпадает с версией в package.json', () => {
    assert.strictEqual(portalConfig.version, packageJson.version);
});

test('config: версия портала задана в формате semver', () => {
    assert.match(portalConfig.version, /^\d+\.\d+\.\d+$/);
});

test('config: глобальные константы PORTAL_VERSION и PORTAL_RELEASE_DATE доступны', () => {
    assert.strictEqual(config.PORTAL_VERSION, portalConfig.version);
    assert.strictEqual(config.PORTAL_RELEASE_DATE, portalConfig.releaseDate);
});

test('config: дата выпуска задана в формате ДД.ММ.ГГГГ', () => {
    assert.match(portalConfig.releaseDate, /^\d{2}\.\d{2}\.\d{4}$/);
});

/* ---------- Реестр сервисов ---------- */

test('config: реестр сервисов не пуст', () => {
    assert.ok(Array.isArray(portalConfig.services));
    assert.ok(portalConfig.services.length > 0);
});

test('config: идентификаторы сервисов уникальны', () => {
    const ids = portalConfig.services.map(service => service.id);
    assert.strictEqual(new Set(ids).size, ids.length);
});

test('config: у каждого сервиса заполнены обязательные поля', () => {
    for (const service of portalConfig.services) {
        assert.ok(service.id, 'не задан id');
        assert.ok(service.title, `не задан title у ${service.id}`);
        assert.ok(service.icon, `не задан icon у ${service.id}`);
        assert.ok(service.url, `не задан url у ${service.id}`);
        assert.ok(service.desc, `не задан desc у ${service.id}`);
        assert.match(service.version, /^\d+\.\d+\.\d+$/, `неверная версия сервиса ${service.id}`);
    }
});

test('config: страницы сервисов существуют на диске', () => {
    for (const service of portalConfig.services) {
        assert.ok(fs.existsSync(path.join('.', service.url)), `нет файла ${service.url}`);
    }
});

test('config: каждая папка Services с index.html зарегистрирована и добавлена в SHELL', () => {
    const servicePages = fs.readdirSync('./Services', { withFileTypes: true })
        .filter(entry => entry.isDirectory())
        .map(entry => `Services/${entry.name}/index.html`)
        .filter(page => fs.existsSync(page));
    const registered = new Set(portalConfig.services.map(service => service.url));
    const sw = fs.readFileSync('./sw.js', 'utf8');

    for (const page of servicePages) {
        assert.ok(registered.has(page), `папка с index.html не зарегистрирована в config.js: ${page}`);
        assert.ok(sw.includes(`'./${page}'`), `нет в precache Service Worker: ${page}`);
    }
    for (const url of registered) {
        assert.ok(servicePages.includes(url), `в config.js указана несуществующая страница: ${url}`);
    }
});

/* ---------- Подключение конфигурации ---------- */

test('index.html: подключает config.js и pwa-update.js', () => {
    const html = fs.readFileSync('./index.html', 'utf8');
    assert.ok(html.includes('shared/js/config.js'));
    assert.ok(html.includes('shared/js/pwa-update.js'));
});

test('index.html: реестр карточек берётся из конфигурации', () => {
    const html = fs.readFileSync('./index.html', 'utf8');
    assert.ok(html.includes('self.PORTAL_CONFIG.services'));
    assert.ok(!html.includes('const SERVICES = ['), 'реестр сервисов продублирован в index.html');
});

test('index.html: версия портала подставляется в подвал', () => {
    const html = fs.readFileSync('./index.html', 'utf8');
    assert.ok(html.includes('id="portalVersion"'));
    assert.ok(!/·\s*v\d+\.\d+\s*\(PWA\)/.test(html), 'в подвале осталась захардкоженная версия');
    assert.ok(!/<span id="portalVersion">[^<]+<\/span>/.test(html),
        'в подвале осталась захардкоженная версия внутри span#portalVersion');
});

test('сервисы: подключают config.js и pwa-update.js', () => {
    for (const service of portalConfig.services) {
        const html = fs.readFileSync(path.join('.', service.url), 'utf8');
        assert.ok(html.includes('shared/js/config.js'), `${service.id}: не подключён config.js`);
        assert.ok(html.includes('shared/js/pwa-update.js'), `${service.id}: не подключён pwa-update.js`);
    }
});

test('сервисы: шапка и подвал связаны с записью конфигурации', () => {
    for (const service of portalConfig.services) {
        const html = fs.readFileSync(path.join('.', service.url), 'utf8');
        const headerMatch = html.match(/data-service-header[^>]*data-version-service="([^"]+)"/);
        const footerMatch = html.match(/data-service-footer[^>]*data-version-service="([^"]+)"/);
        assert.ok(headerMatch, `${service.id}: у шапки нет data-version-service`);
        assert.ok(footerMatch, `${service.id}: у подвала нет data-version-service`);
        assert.strictEqual(headerMatch[1], service.id);
        assert.strictEqual(footerMatch[1], service.id);
    }
});

test('сервисы: в разметке нет захардкоженных версий', () => {
    const pattern = /v\d+\.\d+/;
    for (const service of portalConfig.services) {
        const html = fs.readFileSync(path.join('.', service.url), 'utf8');
        const withoutConfig = html.replace(/data-version-service="[^"]+"/g, '');
        const match = withoutConfig.match(pattern);
        assert.strictEqual(match, null, `${service.id}: версия продублирована в разметке (${match && match[0]})`);
    }
});

/* ---------- Service Worker ---------- */

test('sw.js: импортирует единую конфигурацию', () => {
    const sw = fs.readFileSync('./sw.js', 'utf8');
    assert.ok(sw.includes("importScripts('./shared/js/config.js')"));
});

test('sw.js: имя кэша выводится из версии портала', () => {
    const sw = fs.readFileSync('./sw.js', 'utf8');
    assert.ok(sw.includes('self.PORTAL_VERSION'));
    assert.ok(!/lemanapro-portal-v\d+/.test(sw), 'в sw.js остался захардкоженный номер кэша');
});

test('sw.js: поддерживает сообщения SKIP_WAITING и GET_VERSION', () => {
    const sw = fs.readFileSync('./sw.js', 'utf8');
    assert.ok(sw.includes('SKIP_WAITING'));
    assert.ok(sw.includes('GET_VERSION'));
    assert.ok(sw.includes("type: 'UPDATE_READY'"));
});

test('sw.js: не активируется молча - skipWaiting только по запросу страницы', () => {
    const sw = fs.readFileSync('./sw.js', 'utf8');
    const installBlock = sw.slice(sw.indexOf("addEventListener('install'"), sw.indexOf("addEventListener('activate'"));
    assert.ok(!installBlock.includes('skipWaiting'), 'в install запрещено автоматическое skipWaiting');
    assert.ok(sw.includes("data.type === 'SKIP_WAITING'"));
});

test('sw.js: сверяет версию с сервером и переключает кэш', () => {
    const sw = fs.readFileSync('./sw.js', 'utf8');
    assert.ok(sw.includes('checkForNewVersion'));
    assert.ok(sw.includes("cache: 'no-store'"));
    assert.ok(sw.includes('APPLY_UPDATE'));
});

test('config.js: первая версия в файле - версия портала (её читает Service Worker)', () => {
    const code = fs.readFileSync('./shared/js/config.js', 'utf8');
    const match = code.match(/version:\s*'([^']+)'/);
    assert.ok(match);
    assert.strictEqual(match[1], portalConfig.version);
});

test('pwa-update.js: регистрирует Service Worker по постоянному адресу', () => {
    const module = fs.readFileSync('./shared/js/pwa-update.js', 'utf8');
    assert.ok(module.includes("register(`${portalRoot()}sw.js`)"));
    assert.ok(!/register\([^)]*\?v=/.test(module), 'адрес Service Worker не должен содержать версию');
});

test('pwa-update.js: поддерживает подтверждение обновления и перезагрузку', () => {
    const module = fs.readFileSync('./shared/js/pwa-update.js', 'utf8');
    assert.ok(module.includes('SKIP_WAITING'));
    assert.ok(module.includes('APPLY_UPDATE'));
    assert.ok(module.includes('checkPortalVersion'));
    assert.ok(module.includes('window.location.reload()'));
});

test('pwa-update.js: не перезагружает страницу без подтверждения пользователя', () => {
    const module = fs.readFileSync('./shared/js/pwa-update.js', 'utf8');
    const controllerBlock = module.slice(module.indexOf("addEventListener('controllerchange'"));
    assert.ok(controllerBlock.includes('if (!updateAccepted) return;'));
});

test('sw.js: оболочка включает общие модули конфигурации и обновлений', () => {
    const sw = fs.readFileSync('./sw.js', 'utf8');
    assert.ok(sw.includes("'./shared/js/config.js'"));
    assert.ok(sw.includes("'./shared/js/pwa-update.js'"));
});

/* ---------- Общие утилиты версий ---------- */

test('common.js: formatVersion сокращает semver до подписи в интерфейсе', () => {
    const commonCode = fs.readFileSync('./shared/js/common.js', 'utf8');
    const wrapper = new Function(
        'window', 'document', 'self', 'XLSX', 'Chart', 'console', 'alert',
        `${commonCode}\nreturn { formatVersion, getServiceConfig };`
    );
    const noop = () => {};
    const env = wrapper(
        { addEventListener: noop, removeEventListener: noop, dispatchEvent: () => true },
        { getElementById: () => null, querySelectorAll: () => [], addEventListener: noop },
        config,
        { utils: {} },
        function () {},
        console,
        noop
    );

    assert.strictEqual(env.formatVersion('2.5.9'), 'v2.5');
    assert.strictEqual(env.formatVersion('1.0.0'), 'v1.0');
    assert.strictEqual(env.getServiceConfig('avs').version, '1.0.0');
    assert.strictEqual(env.getServiceConfig('unknown-service'), null);
});
