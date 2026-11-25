/* =====================================================
   ЛОГИСТИКА ANALYTICS - IndexedDB хранилище
   Стандарт для всех сервисов портала (по образцу «Анализ Ячеек Склада»).
   Каждый сервис создаёт свой экземпляр с уникальным dbName.
   Соединение открывается один раз и переиспользуется.
   ===================================================== */

/**
 * Обёртка над IndexedDB для хранения пар ключ-значение.
 * Каждый сервис создаёт свой экземпляр с уникальным dbName.
 */
class IDBStore {
    /**
     * @param {Object} options
     * @param {string} options.dbName - Имя базы IndexedDB.
     * @param {string} [options.storeName='kv'] - Имя хранилища (objectStore).
     * @param {number} [options.version=1] - Версия базы.
     */
    constructor({ dbName, storeName = 'kv', version = 1 } = {}) {
        this.dbName = dbName;
        this.storeName = storeName;
        this.version = version;
        /** @type {Promise<IDBDatabase>|null} Кэш открытого соединения. */
        this._dbPromise = null;
    }

    /**
     * Возвращает фабрику IndexedDB или null, если API недоступно.
     * @returns {IDBFactory|null}
     */
    _factory() {
        return typeof indexedDB !== 'undefined' ? indexedDB : null;
    }

    /**
     * Открывает соединение с БД (один раз) и создаёт хранилище при необходимости.
     * Повторные вызовы возвращают то же соединение.
     * @returns {Promise<IDBDatabase>} Соединение с базой.
     * @rejects {Error} Если IndexedDB недоступна, ошибка открытия или блокировка вкладками.
     */
    open() {
        if (this._dbPromise) return this._dbPromise;

        const factory = this._factory();
        if (!factory) {
            return Promise.reject(new Error('IndexedDB недоступна в этом браузере'));
        }

        this._dbPromise = new Promise((resolve, reject) => {
            const req = factory.open(this.dbName, this.version);

            req.onupgradeneeded = (e) => {
                const db = e.target.result;
                if (!db.objectStoreNames.contains(this.storeName)) {
                    db.createObjectStore(this.storeName, { keyPath: 'key' });
                }
            };

            req.onsuccess = () => {
                const db = req.result;
                // Другая вкладка обновила версию БД - отпускаем соединение,
                // чтобы она не осталась заблокированной.
                db.onversionchange = () => {
                    db.close();
                    this._dbPromise = null;
                };
                db.onclose = () => {
                    this._dbPromise = null;
                };
                resolve(db);
            };

            req.onerror = () => {
                this._dbPromise = null;
                reject(req.error || new Error('Ошибка открытия IndexedDB'));
            };

            req.onblocked = () => {
                this._dbPromise = null;
                reject(new Error('Открытие БД заблокировано другими вкладками'));
            };
        });

        return this._dbPromise;
    }

    /**
     * Создаёт транзакцию, сбрасывая кэш соединения при ошибке.
     * @param {IDBDatabase} db - Открытая база.
     * @param {IDBTransactionMode} mode - Режим транзакции.
     * @param {Function} reject - Коллбэк отклонения промиса.
     * @returns {IDBTransaction|null} Транзакция или null при ошибке.
     */
    _transaction(db, mode, reject) {
        try {
            return db.transaction(this.storeName, mode);
        } catch (e) {
            this._dbPromise = null;
            reject(e);
            return null;
        }
    }

    /**
     * Читает значение по ключу.
     * @param {string} key - Ключ записи.
     * @returns {Promise<*>} Значение или null, если ключ не найден.
     * @rejects {Error} При ошибке чтения или блокировке БД.
     */
    async get(key) {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const tx = this._transaction(db, 'readonly', reject);
            if (!tx) return;
            const r = tx.objectStore(this.storeName).get(key);
            r.onsuccess = () => { resolve(r.result ? r.result.value : null); };
            r.onerror = () => { reject(r.error); };
        });
    }

    /**
     * Сохраняет значение по ключу (создаёт или обновляет).
     * @param {string} key - Ключ записи.
     * @param {*} value - Значение.
     * @returns {Promise<boolean>} true при успешном сохранении.
     * @rejects {Error} При ошибке транзакции или блокировке БД.
     */
    async set(key, value) {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const tx = this._transaction(db, 'readwrite', reject);
            if (!tx) return;
            tx.objectStore(this.storeName).put({ key: key, value: value });
            tx.oncomplete = () => { resolve(true); };
            tx.onerror = () => { reject(tx.error); };
            tx.onabort = () => { reject(tx.error || new Error('Транзакция прервана')); };
        });
    }

    /**
     * Удаляет запись по ключу.
     * @param {string} key - Ключ записи.
     * @returns {Promise<boolean>} true при успешном удалении.
     * @rejects {Error} При ошибке транзакции.
     */
    async delete(key) {
        const db = await this.open();
        return new Promise((resolve, reject) => {
            const tx = this._transaction(db, 'readwrite', reject);
            if (!tx) return;
            tx.objectStore(this.storeName).delete(key);
            tx.oncomplete = () => { resolve(true); };
            tx.onerror = () => { reject(tx.error); };
            tx.onabort = () => { reject(tx.error || new Error('Транзакция прервана')); };
        });
    }

    /**
     * Закрывает открытое соединение (если оно есть).
     * @returns {Promise<void>}
     */
    async close() {
        const promise = this._dbPromise;
        this._dbPromise = null;
        if (!promise) return;
        try {
            const db = await promise;
            db.close();
        } catch (e) {
            // Соединение не было открыто - закрывать нечего.
        }
    }

    /**
     * Удаляет всю базу данных. Сначала закрывает текущее соединение,
     * чтобы удаление не блокировалось этой же вкладкой.
     * @returns {Promise<boolean>} true при успешном удалении.
     * @rejects {Error} Если удаление заблокировано открытыми вкладками.
     */
    async deleteDB() {
        await this.close();

        const factory = this._factory();
        if (!factory) throw new Error('IndexedDB недоступна в этом браузере');

        return new Promise((resolve, reject) => {
            const rq = factory.deleteDatabase(this.dbName);
            rq.onsuccess = () => { resolve(true); };
            rq.onerror = () => { reject(rq.error); };
            rq.onblocked = () => { reject(new Error('Удаление заблокировано открытыми вкладками - закройте другие вкладки с отчётом')); };
        });
    }
}
