import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';

/**
 * Загружает shared/js/config.js в изолированном окружении.
 * @returns {Object} Объект self с PORTAL_CONFIG.
 */
function loadConfig() {
    const code = fs.readFileSync('./shared/js/config.js', 'utf8');
    const sandbox = {};
    new Function('self', code)(sandbox);
    return sandbox;
}

const portalConfig = loadConfig().PORTAL_CONFIG;
const notFound = fs.readFileSync('./404.html', 'utf8');
const serviceWorker = fs.readFileSync('./sw.js', 'utf8');

/* ---------- Страница 404 ---------- */

test('404: страница существует и объясняет ситуацию', () => {
    assert.ok(notFound.includes('Такой страницы нет'));
    assert.ok(notFound.includes('>404<'));
});

test('404: есть ссылка на главную страницу портала', () => {
    assert.ok(notFound.includes('href="/"'), 'нет ссылки на главную');
});

test('404: перечислены все отчёты портала', () => {
    for (const service of portalConfig.services) {
        assert.ok(notFound.includes(`href="/${service.url}"`),
            `на странице 404 нет ссылки на ${service.title}`);
    }
});

/* ---------- Service Worker ---------- */

test('404: страница добавлена в precache Service Worker', () => {
    assert.ok(serviceWorker.includes("'./404.html'"), 'нет ./404.html в SHELL');
});

test('404: Service Worker отдаёт статус 404, а не главную страницу', () => {
    assert.ok(serviceWorker.includes('res.status === 404'),
        'навигация не передаёт серверный ответ 404 как есть');
    assert.ok(serviceWorker.includes('cachedNotFound'),
        'нет офлайн-страницы 404');
    assert.ok(!serviceWorker.includes("caches.match('./index.html')"),
        'остался фолбэк на главную страницу для несуществующих адресов');
});
