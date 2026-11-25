/* =====================================================
   ЛОГИСТИКА ANALYTICS - общие утилиты для сервисов
   Работает совместно с shared/js/idb.js.
   Подклюжать после idb.js и после библиотек XLSX / Chart.js.
   ===================================================== */

/* ---------- XLSX / парсинг ---------- */

/**
 * Находит индекс колонки по списку возможных имён (точное или частичное совпадение).
 * @param {string[]} headers - Массив имён колонок (заголовков).
 * @param {string[]} possibleNames - Допустимые имена для поиска.
 * @returns {number} Индекс колонки или -1, если не найдена.
 */
function findColumnIndex(headers, possibleNames) {
    for (let name of possibleNames) {
        const exact = headers.findIndex(h => h && h.trim().toLowerCase() === name.toLowerCase());
        if (exact >= 0) return exact;
    }
    for (let name of possibleNames) {
        const idx = headers.findIndex(h => h && h.toLowerCase().includes(name.toLowerCase()));
        if (idx >= 0) return idx;
    }
    return -1;
}

/**
 * Приводит значение к чистой строке.
 * @param {*} val - Исходное значение.
 * @returns {string} Строка без пробелов по краям или пустая строка.
 */
function cleanValue(val) {
    if (val === null || val === undefined) return '';
    return String(val).trim();
}

/**
 * Парсит число из значения, сохраняя отличие «пусто» от нуля.
 * Поддерживает пробелы в тысячах и запятую как десятичный разделитель.
 * @param {*} val - Исходное значение.
 * @returns {number|null} Число или null, если значение пустое/нечисловое.
 */
function parseNullableNumber(val) {
    if (val === null || val === undefined) return null;
    if (typeof val === 'number') return Number.isFinite(val) ? val : null;
    const str = String(val).trim();
    if (!str) return null;
    const normalized = str.replace(/\s/g, '').replace(',', '.');
    if (!/^[+-]?(?:\d+\.?\d*|\.\d+)$/.test(normalized)) return null;
    const num = parseFloat(normalized);
    return Number.isFinite(num) ? num : null;
}

/**
 * Парсит число из значения (поддерживает пробелы в тысячах и запятую как разделитель).
 * @param {*} val - Исходное значение.
 * @returns {number} Число или 0, если не удалось распарсить.
 */
function parseNumber(val) {
    const num = parseNullableNumber(val);
    return num === null ? 0 : num;
}

/**
 * Парсит дату из ячейки XLSX/строки. Понимает Date, Excel-serial, dd.mm.yyyy,
 * dd.mm.yy, yyyy-mm-dd и ISO-строки. Неоднозначные MM/DD намеренно не угадывает.
 * @param {*} val - Исходное значение.
 * @returns {Date|null} Дата или null, если распарсить не удалось.
 */
function parseDateCell(val) {
    if (val === null || val === undefined || val === '') return null;
    if (val instanceof Date) return isNaN(val.getTime()) ? null : val;

    if (typeof val === 'number') {
        if (!Number.isFinite(val) || val <= 0) return null;
        // Excel serial date (1900 date system). Дробную часть (время) отбрасываем.
        const ms = Math.round((Math.floor(val) - 25569) * 86400000);
        const utc = new Date(ms);
        if (isNaN(utc.getTime())) return null;
        return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
    }

    const str = String(val).trim();
    if (!str) return null;

    // dd.mm.yyyy / dd.mm.yy (допускается время после даты)
    let match = str.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})(?:[ T].*)?$/);
    if (match) {
        const day = parseInt(match[1], 10);
        const month = parseInt(match[2], 10);
        let year = parseInt(match[3], 10);
        if (year < 100) year += year >= 70 ? 1900 : 2000;
        const date = new Date(year, month - 1, day);
        return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
            ? date
            : null;
    }

    // yyyy-mm-dd (допускается время после даты)
    match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T].*)?$/);
    if (match) {
        const year = parseInt(match[1], 10);
        const month = parseInt(match[2], 10);
        const day = parseInt(match[3], 10);
        const date = new Date(year, month - 1, day);
        return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day
            ? date
            : null;
    }

    const parsed = new Date(str);
    return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Парсит время из ячейки XLSX/строки.
 * @param {*} val - Date, число (доля суток Excel) или строка вида «13:45».
 * @returns {number|null} Минуты от начала суток (0-1439) или null.
 */
function parseTimeCell(val) {
    if (val === null || val === undefined || val === '') return null;
    if (val instanceof Date) {
        return isNaN(val.getTime()) ? null : val.getHours() * 60 + val.getMinutes();
    }
    if (typeof val === 'number') {
        if (!Number.isFinite(val)) return null;
        const minutes = Math.round(val * 24 * 60) % (24 * 60);
        return ((minutes % 1440) + 1440) % 1440;
    }
    const str = String(val).trim();
    if (!str) return null;
    const match = str.match(/^(\d{1,2}):(\d{1,2})(?::\d{1,2})?$/);
    if (match) {
        const hours = parseInt(match[1], 10);
        const minutes = parseInt(match[2], 10);
        return hours < 24 && minutes < 60 ? hours * 60 + minutes : null;
    }
    const numeric = parseNullableNumber(str);
    return numeric === null ? null : parseTimeCell(numeric);
}

/**
 * Экранирует HTML-спецсимволы в строке.
 * @param {string} str - Исходная строка.
 * @returns {string} Экранированная строка.
 */
function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, s => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[s]));
}

/**
 * Создаёт debounce-обёртку над функцией.
 * @param {Function} fn - Функция для обёртки.
 * @param {number} delay - Задержка в мс.
 * @returns {Function} Debounced функция.
 */
function debounce(fn, delay) {
    let timer;
    return function (...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}

/* ---------- Пагинация ---------- */

/**
 * Считает количество страниц.
 * @param {number} totalItems - Всего записей.
 * @param {number} perPage - Записей на страницу.
 * @returns {number} Количество страниц (0, если данных нет).
 */
function getPageCount(totalItems, perPage) {
    const total = Number(totalItems) || 0;
    const size = Number(perPage) || 0;
    if (total <= 0 || size <= 0) return 0;
    return Math.ceil(total / size);
}

/**
 * Возвращает срез данных для страницы, не выходя за границы.
 * @param {Array} items - Все записи.
 * @param {number} page - Номер страницы (1-based).
 * @param {number} perPage - Записей на страницу.
 * @returns {Array} Записи выбранной страницы.
 */
function getPageSlice(items, page, perPage) {
    const list = Array.isArray(items) ? items : [];
    const size = Math.max(1, Number(perPage) || 0);
    const totalPages = Math.max(1, getPageCount(list.length, size));
    const current = Math.min(Math.max(1, parseInt(page, 10) || 1), totalPages);
    const start = (current - 1) * size;
    return list.slice(start, start + size);
}

/**
 * Настраивает зону загрузки файла: клик, drag-and-drop, клавиатура, input[type=file].
 * @param {string} areaId - ID элемента зоны загрузки.
 * @param {string} inputId - ID скрытого <input type="file">.
 * @param {Function} onFile - Коллбэк с объектом File при выборе файла.
 */
function setupUploadArea(areaId, inputId, onFile) {
    const area = document.getElementById(areaId);
    const input = document.getElementById(inputId);
    if (!area || !input) return;

    area.addEventListener('click', () => input.click());
    area.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            input.click();
        }
    });
    area.addEventListener('dragover', (e) => {
        e.preventDefault();
        area.classList.add('dragover');
    });
    area.addEventListener('dragleave', () => area.classList.remove('dragover'));
    area.addEventListener('drop', (e) => {
        e.preventDefault();
        area.classList.remove('dragover');
        if (e.dataTransfer.files.length > 0) onFile(e.dataTransfer.files[0]);
    });
    input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) onFile(e.target.files[0]);
    });
}

/**
 * Расширяет диапазон листа (`!ref`) по фактически заполненным ячейкам.
 * Некоторые программы сохраняют XLSX с занижённым `!ref` (например, `A1:S5` при 2 165 строках),
 * из-за чего `sheet_to_json` молча обрезает данные. Диапазон только расширяется и никогда не сужается.
 * @param {Object} sheet - Лист SheetJS.
 */
function expandSheetRange(sheet) {
    if (!sheet) return;
    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -1;
    let maxCol = -1;

    for (const key of Object.keys(sheet)) {
        if (key.charCodeAt(0) === 33) continue; // служебные ключи SheetJS начинаются с «!»
        if (!/^[A-Z]+[0-9]+$/.test(key)) continue;
        const cell = sheet[key];
        const hasValue = cell && (
            (cell.v !== undefined && cell.v !== null && cell.v !== '') ||
            cell.f !== undefined ||
            cell.w !== undefined
        );
        if (!hasValue) continue;
        const addr = XLSX.utils.decode_cell(key);
        if (!Number.isFinite(addr.r) || !Number.isFinite(addr.c)) continue;
        if (addr.r < minRow) minRow = addr.r;
        if (addr.c < minCol) minCol = addr.c;
        if (addr.r > maxRow) maxRow = addr.r;
        if (addr.c > maxCol) maxCol = addr.c;
    }

    if (maxRow < 0 || maxCol < 0) return;

    const current = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    const start = current ? current.s : { r: minRow, c: minCol };
    const end = current ? current.e : { r: maxRow, c: maxCol };
    sheet['!ref'] = XLSX.utils.encode_range({
        s: { r: Math.min(start.r, minRow), c: Math.min(start.c, minCol) },
        e: { r: Math.max(end.r, maxRow), c: Math.max(end.c, maxCol) }
    });
}

/**
 * Читает .xlsx/.xls файл через FileReader + SheetJS и возвращает массив массивов (header: 1).
 * @param {File} file - Файл Excel.
 * @returns {Promise<Array<Array<*>>>} Данные первого листа в виде массива строк/ячеек.
 * @rejects {Error} Ошибка формата, чтения или отсутствия листов.
 */
function parseXLSXFile(file) {
    return new Promise((resolve, reject) => {
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
            reject(new Error('Загрузите файл формата .xlsx или .xls'));
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', cellDates: true });
                if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
                    reject(new Error('Файл не содержит листов'));
                    return;
                }
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                expandSheetRange(sheet);
                const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
                resolve(json);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = () => reject(new Error('Ошибка чтения файла'));
        reader.readAsArrayBuffer(file);
    });
}

/* ---------- Валидация XLSX ---------- */

/**
 * Проверяет наличие обязательных колонок в заголовке XLSX (частичное совпадение имён).
 * Каждый элемент requiredColumns может быть строкой (одно имя) или массивом строк
 * (группа альтернатив - достаточно, чтобы нашлась хотя бы одна из них).
 * @param {Array<Array<*>>} jsonData - Данные листа.
 * @param {(string|string[])[]} requiredColumns - Обязательные имена или группы альтернатив.
 * @param {number} [headerRowIndex=0] - Индекс строки с заголовками.
 * @returns {boolean} true, если все колонки найдены.
 * @throws {Error} Если данных нет, нет заголовков или отсутствуют обязательные колонки.
 */
function validateXLSXHeaders(jsonData, requiredColumns, headerRowIndex = 0) {
    if (!jsonData || !Array.isArray(jsonData) || jsonData.length === 0) {
        throw new Error('Файл не содержит данных');
    }
    const rowIndex = Number.isInteger(headerRowIndex) && headerRowIndex >= 0 ? headerRowIndex : 0;
    const headers = jsonData[rowIndex];
    if (!Array.isArray(headers) || headers.length === 0) {
        throw new Error('Файл не содержит заголовков колонок');
    }
    const missing = [];
    for (const col of requiredColumns) {
        const alternatives = Array.isArray(col) ? col : [col];
        const found = alternatives.some((alt) => {
            const lowerAlt = String(alt).toLowerCase();
            return headers.some((h) => h && String(h).toLowerCase().includes(lowerAlt));
        });
        if (!found) missing.push(alternatives.join(' / '));
    }
    if (missing.length > 0) {
        throw new Error('В файле не найдены обязательные колонки: ' + missing.join(', '));
    }
    return true;
}

/* ---------- Хранилище отчёта (IndexedDB) ---------- */

/**
 * Создаёт обёртку для загрузки/сохранения данных отчёта в IndexedDB.
 *
 * Поддерживаются два формата хранения:
 * - массив записей (wrapped: false);
 * - объект { records, timestamp } (wrapped: true).
 *
 * @param {Object} options
 * @param {IDBStore} options.idb - Экземпляр хранилища.
 * @param {string} options.key - Ключ записи в IndexedDB.
 * @param {boolean} [options.wrapped=false] - Хранить вместе с датой загрузки.
 * @returns {{load: Function, save: Function, clear: Function}} Методы работы с хранилищем.
 */
function createReportStorage({ idb, key, wrapped = false }) {
    /**
     * Приводит сохранённое значение к формату { records, timestamp }.
     * @param {*} payload - Значение из IndexedDB.
     * @returns {{records: Array, timestamp: number|null}|null}
     */
    function unwrapPayload(payload) {
        if (!payload) return null;
        if (Array.isArray(payload)) return { records: payload, timestamp: null };
        if (Array.isArray(payload.records)) {
            const result = { records: payload.records, timestamp: payload.timestamp || null };
            if (payload.meta && typeof payload.meta === 'object') result.meta = payload.meta;
            return result;
        }
        return null;
    }

    return {
        /**
         * Загружает данные из IndexedDB.
         * @returns {Promise<{records: Array, timestamp: number|null, meta?: Object}|null>}
         */
        async load() {
            const stored = unwrapPayload(await idb.get(key));
            if (stored && stored.records.length > 0) return stored;
            return null;
        },

        /**
         * Сохраняет данные отчёта.
         * @param {Array} records - Записи.
         * @param {number} [timestamp=Date.now()] - Дата загрузки.
         * @param {Object|null} [meta=null] - Дополнительные данные отчёта (только для wrapped).
         * @returns {Promise<number>} Сохранённая дата загрузки.
         */
        async save(records, timestamp = Date.now(), meta = null) {
            const payload = wrapped ? { records, timestamp } : records;
            if (wrapped && meta && typeof meta === 'object' && Object.keys(meta).length > 0) {
                payload.meta = meta;
            }
            await idb.set(key, payload);
            return timestamp;
        },

        /**
         * Удаляет базу отчёта.
         * @returns {Promise<void>}
         */
        async clear() {
            await idb.deleteDB();
        },
    };
}

/* ---------- UI state ---------- */

/**
 * Закрывает модалку по Escape, если она открыта.
 * @param {string} modalId - ID модалки.
 * @param {Function} closeFn - Функция закрытия.
 * @param {Function} [onEscape] - Дополнительный обработчик Escape.
 */
function setupEscapeToClose(modalId, closeFn, onEscape) {
    if (typeof document === 'undefined') return;
    document.addEventListener('keydown', (e) => {
        if (e.key !== 'Escape') return;
        if (modalId) {
            const modal = document.getElementById(modalId);
            if (modal && (modal.classList.contains('active') || modal.classList.contains('open'))) {
                if (typeof closeFn === 'function') closeFn();
            }
        }
        if (typeof onEscape === 'function') onEscape();
    });
}

/**
 * Показывает или скрывает элемент по ID.
 * @param {string} id - ID элемента.
 * @param {boolean} show - true - показать, false - скрыть.
 */
function showElement(id, show) {
    const el = document.getElementById(id);
    if (el) el.style.display = show ? '' : 'none';
}

/**
 * Показывает/скрывает индикатор загрузки (#loading).
 * @param {boolean} show - true - показать, false - скрыть.
 */
function showLoading(show) {
    showElement('loading', show);
}

/**
 * Показывает/скрывает секцию загрузки файла (#uploadSection).
 * @param {boolean} show - true - показать, false - скрыть.
 */
function showUpload(show) {
    showElement('uploadSection', show);
}

/**
 * Показывает/скрывает дашборд (#dashboard) через класс .active.
 * @param {boolean} show - true - показать, false - скрыть.
 */
function showDashboard(show) {
    const el = document.getElementById('dashboard');
    if (el) el.classList.toggle('active', show);
}

/* ---------- Data notice ---------- */

/**
 * Обновляет блок уведомления о данных (сохранено/сессия).
 * @param {string} noticeId - ID контейнера уведомления.
 * @param {string} titleId - ID элемента заголовка.
 * @param {string} infoId - ID элемента дополнительной информации.
 * @param {string} mode - 'saved' или 'session'.
 * @param {number} count - Количество записей.
 * @param {Date} date - Дата загрузки.
 */
function updateDataNotice(noticeId, titleId, infoId, mode, count, date) {
    const notice = document.getElementById(noticeId);
    const title = document.getElementById(titleId);
    const info = document.getElementById(infoId);
    if (!notice || !title || !info) return;

    notice.classList.add('active');
    notice.classList.remove('session');

    if (mode === 'saved') {
        title.textContent = '💾 Данные сохранены (IndexedDB)!';
        info.innerHTML = `Загружено: <strong>${count.toLocaleString('ru-RU')}</strong> | Дата: <strong>${date.toLocaleString('ru-RU')}</strong>`;
    } else {
        notice.classList.add('session');
        title.textContent = '⚠️ Режим сессии (без сохранения)!';
        info.innerHTML = `Загружено: <strong>${count.toLocaleString('ru-RU')}</strong> | Данные будут удалены при закрытии страницы`;
    }
}

/* ---------- Confirm modal (inject once) ---------- */

/**
 * Внедряет в DOM модалку подтверждения очистки данных (пересоздаёт при каждом вызове,
 * чтобы гарантированно использовать актуальные классы стилей).
 */
function ensureConfirmModal() {
    const existing = document.getElementById('confirmModal');
    if (existing) existing.remove();
    const html = `
    <div class="confirm-modal-overlay" id="confirmModal" role="dialog" aria-modal="true">
        <div class="confirm-modal">
            <div class="confirm-modal-header">
                <h3>⚠️ Подтвердите действие</h3>
                <button class="confirm-modal-close" onclick="closeConfirmModal()" aria-label="Закрыть">✕</button>
            </div>
            <div class="confirm-modal-body">
                <p class="confirm-modal-text" id="confirmText">
                    Все сохранённые данные будут <strong>безвозвратно удалены</strong> из IndexedDB браузера. Продолжить?
                </p>
            </div>
            <div class="confirm-modal-footer">
                <button class="confirm-btn-cancel" onclick="closeConfirmModal()">Отмена</button>
                <button class="confirm-btn-clear" onclick="doClearData()">🗑️ Удалить данные</button>
            </div>
        </div>
    </div>`;
    document.body.insertAdjacentHTML('beforeend', html);
}

/**
 * Закрывает модалку подтверждения.
 */
function closeConfirmModal() {
    const m = document.getElementById('confirmModal');
    if (m) { m.classList.remove('open'); }
}

// Глобальный обработчик doClearData переопределяется в каждом сервисе
window.doClearData = function () {
    closeConfirmModal();
};

/**
 * Показывает модалку подтверждения очистки данных.
 * @param {string} [confirmMessage] - Текст подтверждения (по умолчанию стандартный).
 * @param {Function} onConfirmed - Коллбэк после подтверждения пользователем.
 */
function requestClearData(confirmMessage, onConfirmed) {
    ensureConfirmModal();
    const text = document.getElementById('confirmText');
    if (text) text.innerHTML = confirmMessage || 'Все сохранённые данные будут <strong>безвозвратно удалены</strong> из IndexedDB браузера. Продолжить?';
    const m = document.getElementById('confirmModal');
    if (m) { m.classList.add('open'); }

    window.doClearData = function () {
        closeConfirmModal();
        if (typeof onConfirmed === 'function') onConfirmed();
    };
}

/* ---------- Table rendering ---------- */

/**
 * Рендерит строки таблицы в tbody.
 * @param {string} tbodyId - ID элемента <tbody>.
 * @param {Array} pageData - Массив данных страницы.
 * @param {Function} rowHtmlFn - Функция(item, index) -> HTML строки.
 */
function renderTable(tbodyId, pageData, rowHtmlFn) {
    const tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    tbody.innerHTML = '';
    pageData.forEach((item, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = rowHtmlFn(item, index);
        tbody.appendChild(tr);
    });
}

/**
 * Включает одиночное выделение строк таблицы: клик выделяет строку и снимает
 * выделение с предыдущей, повторный клик по той же строке снимает выделение.
 * Клики по ссылкам, кнопкам и полям внутри строки игнорируются.
 * @param {HTMLElement} tbody - Контейнер со строками (<tbody>).
 * @param {string} [rowSelector] - Селектор строк внутри контейнера.
 */
function enableSingleRowSelection(tbody, rowSelector = 'tr') {
    if (!tbody) return;
    const rows = tbody.matches && tbody.matches(rowSelector)
        ? [tbody]
        : Array.from(tbody.querySelectorAll(rowSelector));
    rows.forEach(row => {
        if (row.dataset.selectionBound === '1') return;
        row.dataset.selectionBound = '1';
        row.addEventListener('click', event => {
            if (event.target.closest('a, button, input, select, textarea, label')) return;
            const wasSelected = row.classList.contains('row-selected');
            rows.forEach(r => r.classList.remove('row-selected'));
            if (!wasSelected) row.classList.add('row-selected');
        });
    });
}

/**
 * Рендерит пагинацию с кнопками страницами.
 * @param {string} containerId - ID контейнера пагинации.
 * @param {number} totalPages - Всего страниц.
 * @param {number} currentPage - Текущая страница (1-based).
 * @param {string} onPageClick - Имя глобальной функции-обработчика клика (вставляется в onclick).
 */
function updatePagination(containerId, totalPages, currentPage, onPageClick) {
    const p = document.getElementById(containerId);
    if (!p) return;
    if (totalPages <= 1) {
        p.innerHTML = '';
        return;
    }
    let html = `<button onclick="${onPageClick}(1)" ${currentPage === 1 ? 'disabled' : ''} aria-label="Первая страница">«</button>`;
    for (let i = Math.max(1, currentPage - 2); i <= Math.min(totalPages, currentPage + 2); i++) {
        html += `<button onclick="${onPageClick}(${i})" class="${i === currentPage ? 'active' : ''}" aria-label="Страница ${i}">${i}</button>`;
    }
    html += `<button onclick="${onPageClick}(${totalPages})" ${currentPage === totalPages ? 'disabled' : ''} aria-label="Последняя страница">»</button>`;
    p.innerHTML = html;
}

/* ---------- Sorting ---------- */

/**
 * Сортирует массив объектов по колонке.
 * @param {Array} data - Исходные данные.
 * @param {string} column - Имя колонки для сортировки.
 * @param {string|null} sortColumn - Текущая отсортированная колонка.
 * @param {boolean} sortDirection - Направление (true = по убыванию).
 * @returns {{data: Array, sortColumn: string, sortDirection: boolean}} Результат сортировки.
 */
function sortData(data, column, sortColumn, sortDirection) {
    if (sortColumn === column) sortDirection = !sortDirection;
    else { sortColumn = column; sortDirection = false; }
    const sorted = [...data].sort((a, b) => {
        let va = a[column] || '';
        let vb = b[column] || '';
        const na = parseFloat(va);
        const nb = parseFloat(vb);
        if (!isNaN(na) && !isNaN(nb) && va !== '' && vb !== '') { va = na; vb = nb; }
        if (va === vb) return 0;
        return sortDirection ? (va > vb ? -1 : 1) : (va < vb ? -1 : 1);
    });
    return { data: sorted, sortColumn, sortDirection };
}

/**
 * Устанавливает визуальное состояние индикатора сортировки на одном заголовке.
 * @param {HTMLElement} th - Элемент <th>.
 * @param {boolean} isActive - Активна ли колонка сортировки.
 * @param {boolean} direction - Направление сортировки (true = убывание).
 */
function setSortIndicator(th, isActive, direction) {
    th.dataset.sortActive = isActive ? 'true' : null;
    th.setAttribute('aria-sort', isActive ? (direction ? 'descending' : 'ascending') : 'none');
    const indicator = th.querySelector('.sort-indicator');
    if (indicator) {
        indicator.textContent = isActive ? (direction ? '▲' : '▼') : '⬍';
    }
}

/**
 * Обновляет индикаторы сортировки всех заголовков таблицы через data-атрибуты.
 * @param {string|null} sortColumn - Активная колонка сортировки.
 * @param {boolean} sortDirection - Направление сортировки.
 */
function updateSortIndicators(sortColumn, sortDirection) {
    document.querySelectorAll('th[data-sort-col]').forEach((th) => {
        const isActive = th.dataset.sortCol === sortColumn;
        setSortIndicator(th, isActive, sortDirection);
    });
}

/* ---------- Charts ---------- */

/**
 * Уничтожает все графики Chart.js в объекте, освобождая память.
 * @param {Object} chartsObj - Объект с экземплярами Chart по ключам.
 */
function destroyCharts(chartsObj) {
    for (const key in chartsObj) {
        if (chartsObj[key]) {
            chartsObj[key].destroy();
            chartsObj[key] = null;
        }
    }
}

/* ---------- Export ---------- */

/**
 * Экспортирует произвольные строки в XLSX-файл через SheetJS.
 * Имя не пересекается с локальными функциями экспорта сервисов.
 * @param {Array<Array<*>>} wsData - Данные листа (массив строк).
 * @param {string} [sheetName='Отчёт'] - Имя листа.
 * @param {string} [fileName='отчёт'] - Базовое имя файла (добавляется дата и время).
 */
function exportRowsToXLSX(wsData, sheetName, fileName) {
    if (!wsData || wsData.length === 0) {
        alert('Нет данных для экспорта');
        return;
    }
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Отчёт');
    const dateStr = new Date().toISOString().slice(0, 10);
    const timeStr = new Date().toTimeString().slice(0, 5).replace(':', '-');
    XLSX.writeFile(wb, `${fileName || 'отчёт'}_${dateStr}_${timeStr}.xlsx`);
}

/* ---------- Общая шапка и подвал сервисов ---------- */

const SERVICE_SUPPORT_EMAIL = '60041915@lemanapro.ru';

/**
 * Форматирует семантическую версию для интерфейса: 9.10.0 → v9.10.
 * @param {string} version - Версия в формате semver.
 * @returns {string} Подпись версии.
 */
function formatVersion(version) {
    const parts = String(version || '').split('.');
    return `v${parts.slice(0, 2).join('.')}`;
}

/**
 * Возвращает конфигурацию сервиса по его идентификатору.
 * @param {string} serviceId - Идентификатор сервиса.
 * @returns {Object|null} Конфигурация сервиса или null.
 */
function getServiceConfig(serviceId) {
    const config = self.PORTAL_CONFIG;
    if (!config || !Array.isArray(config.services)) return null;
    return config.services.find(service => service.id === serviceId) || null;
}

/**
 * Рендерит стандартную шапку сервиса из data-атрибутов:
 * data-icon, data-title, data-accent, data-subtitle, data-info="false".
 * data-version-service связывает страницу с записью единой конфигурации.
 * @param {HTMLElement} el - Контейнер шапки.
 */
function renderServiceHeader(el) {
    if (!el || !el.dataset) return;
    const d = el.dataset;
    const icon = d.icon ? escapeHtml(d.icon) + ' ' : '';
    const accent = d.accent ? '<span>&nbsp;' + escapeHtml(d.accent) + '</span>' : '';
    const service = getServiceConfig(d.versionService);
    const version = service ? ` <span class="service-version">${escapeHtml(formatVersion(service.version))}</span>` : '';
    const infoButton = d.info === 'false'
        ? ''
        : '<button class="info-btn" onclick="toggleModal()" title="Информация" aria-label="Информация">ℹ️</button>';

    if (service) document.title = `${service.title} ${formatVersion(service.version)}`;

    el.innerHTML = `
        <a href="../../" class="btn-back" title="На портал" aria-label="На портал">←</a>
        <div class="header-left">
            <h1>${icon}${escapeHtml(d.title || '')}${accent}${version}</h1>
            <p>${escapeHtml(d.subtitle || '')}</p>
        </div>
        ${infoButton}`;
}

/**
 * Рендерит стандартный подвал сервиса из data-атрибутов:
 * data-title, data-app, data-mail-subject.
 * @param {HTMLElement} el - Контейнер подвала.
 */
function renderServiceFooter(el) {
    if (!el || !el.dataset) return;
    const d = el.dataset;
    const app = escapeHtml(d.app || 'Логистика Analytics');
    const subject = escapeHtml(encodeURIComponent(d.mailSubject || 'Портал Логистика Analytics'));
    const service = getServiceConfig(d.versionService);
    const version = service ? ` ${escapeHtml(formatVersion(service.version))}` : '';

    el.innerHTML = `
        <div class="footer-tools">
            <span>${escapeHtml(d.title || '')}${version}</span>
            <span>|</span>
            <span>${app}</span>
        </div>
        <div>
            Поддержка:
            <a href="mailto:${SERVICE_SUPPORT_EMAIL}?subject=${subject}">${SERVICE_SUPPORT_EMAIL}</a>
        </div>`;
}

/**
 * Наполняет все шапки и подвалы с data-атрибутами на странице.
 */
function initServiceLayout() {
    if (typeof document === 'undefined') return;
    document.querySelectorAll('[data-service-header]').forEach(renderServiceHeader);
    document.querySelectorAll('[data-service-footer]').forEach(renderServiceFooter);
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', initServiceLayout);
    } else {
        initServiceLayout();
    }
}

/* ---------- Глобальный перехват ошибок ---------- */

/**
 * Устанавливает глобальные обработчики ошибок JS и непойманных промисов.
 */
function installGlobalErrorHandler() {
    window.addEventListener('error', (e) => {
        console.error('[Портал] Глобальная ошибка JS:', e.message,
            e.filename ? `(${e.filename}:${e.lineno}:${e.colno})` : '');
        if (e.error) {
            console.error(e.error);
        }
    });

    window.addEventListener('unhandledrejection', (e) => {
        console.error('[Портал] Необработанный промис:', e.reason);
    });
}

installGlobalErrorHandler();
