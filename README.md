# Портал отчётов · Логистика Analytics

Короткий указатель. Канонический README проекта со всеми деталями:

**[LogisticAnalytics/README.md](LogisticAnalytics/README.md)**

Кратко: клиентский PWA-портал складской аналитики на чистом HTML/CSS/JavaScript, 9 сервисов, версия портала **1.0.0** от **25.11.2026**. Сервер и сборка не нужны.

- Запуск: откройте папку проекта через локальный статический сервер (`localhost` или HTTPS - для PWA и Service Worker).
- Тесты: `node --test` из этой папки. Сейчас 218 из 218 проходят, 0 пропущено. Все примеры XLSX лежат в `not indexed/samples_XLSX/`; имена файлов канонические, без суффиксов даты. Примеры XLSX и `opencode.json` не входят в git, поэтому тесты запускаются только локально. Подробности - в каноническом README.
- Деплой: GitHub Pages через GitHub Actions, публикуются только файлы сайта (`index.html`, `Services/`, `shared/`, `sw.js`, `manifest.webmanifest`, иконки). Тесты, `not indexed/`, XLSX и `opencode.json` на сайт не попадают.
