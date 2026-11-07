/* File-backed record sync. OneDrive transports files; browser data stays local. */
(function (global) {
    'use strict';
    const FORMAT = 'ielts-record-sync-v1';
    const clone = value => JSON.parse(JSON.stringify(value));
    function stable(value) {
        if (Array.isArray(value)) return '[' + value.map(stable).join(',') + ']';
        if (value && typeof value === 'object') return '{' + Object.keys(value).sort()
            .map(key => JSON.stringify(key) + ':' + stable(value[key])).join(',') + '}';
        return JSON.stringify(value);
    }
    function index(records) {
        if (!Array.isArray(records)) throw new Error('记录格式无效，已停止同步');
        const result = new Map();
        for (const record of records) {
            if (!record || record.id == null || !String(record.id)) throw new Error('记录缺少 ID，已停止同步');
            result.set(String(record.id), record);
        }
        return result;
    }
    function validate(payload) {
        if (!payload || payload.format !== FORMAT || !Array.isArray(payload.entries)) {
            throw new Error('同步文件格式不正确，未覆盖本地记录');
        }
        for (const e of payload.entries) {
            if (!e || typeof e.id !== 'string' || !e.id || !Number.isSafeInteger(e.clock) || e.clock < 1 ||
                typeof e.device !== 'string' || !e.device || typeof e.deleted !== 'boolean' ||
                (!e.deleted && (!e.record || String(e.record.id) !== e.id))) {
                throw new Error('同步文件记录损坏，未覆盖本地记录');
            }
        }
        return payload.entries;
    }
    // Lamport counters avoid depending on the computers' wall clocks.
    function merge(...groups) {
        const result = new Map();
        for (const entries of groups) for (const e of entries) {
            const old = result.get(e.id);
            if (!old || e.clock > old.clock || (e.clock === old.clock &&
                (e.device > old.device || (e.device === old.device && stable(e) > stable(old))))) result.set(e.id, e);
        }
        return [...result.values()].sort((a, b) => a.id.localeCompare(b.id));
    }
    function observe(state, records, device) {
        const previous = index(state.baseline || []);
        const current = index(records);
        let clock = state.entries.reduce((max, e) => Math.max(max, e.clock), 0);
        const changes = [];
        for (const [id, record] of current) {
            if (!previous.has(id) || stable(previous.get(id)) !== stable(record)) {
                changes.push({ id, clock: ++clock, device, deleted: false, record: clone(record) });
            }
        }
        for (const id of previous.keys()) if (!current.has(id)) {
            changes.push({ id, clock: ++clock, device, deleted: true });
        }
        return merge(state.entries, changes);
    }
    class OneDriveRecordSync {
        constructor() {
            this.config = null;
            this.running = null;
            this.status('尚未连接。请在每台电脑选择同一个 OneDrive 文件夹。');
        }
        status(message) {
            const el = global.document.getElementById('onedrive-sync-status');
            if (el) el.textContent = message;
        }
        async database() {
            if (!this.dbPromise) this.dbPromise = new Promise((resolve, reject) => {
                const request = global.indexedDB.open('IELTSOneDriveRecordSync', 1);
                request.onupgradeneeded = () => request.result.createObjectStore('config');
                request.onsuccess = () => resolve(request.result);
                request.onerror = () => reject(request.error);
            });
            return this.dbPromise;
        }
        async configIO(value) {
            const db = await this.database();
            return new Promise((resolve, reject) => {
                const tx = db.transaction('config', value === undefined ? 'readonly' : 'readwrite');
                const store = tx.objectStore('config');
                const req = value === undefined ? store.get('connection') : store.put(value, 'connection');
                tx.oncomplete = () => resolve(req.result);
                tx.onerror = () => reject(tx.error);
                tx.onabort = () => reject(tx.error || new Error('无法保存同步连接'));
            });
        }
        async restore() {
            this.config = await this.configIO();
            if (this.config) await this.sync();
        }
        async connect() {
            if (!global.showDirectoryPicker) throw new Error('当前浏览器不支持文件夹同步。请用最新版 Edge 或 Chrome 打开；也可使用原有导出/导入功能。');
            // Picker must be the first asynchronous operation in the click handler.
            const folder = await global.showDirectoryPicker({ mode: 'readwrite', id: 'ielts-record-sync' });
            if (this.running) await this.running;
            const old = this.config;
            if (old && await folder.isSameEntry(old.folder)) {
                this.config = old;
            } else {
                this.config = { folder, device: global.crypto.randomUUID(), entries: [], baseline: [] };
            }
            await this.configIO(this.config);
            await this.sync();
        }
        async resume() {
            if (!this.config) return this.connect();
            if (await this.config.folder.requestPermission({ mode: 'readwrite' }) !== 'granted') {
                throw new Error('未获得文件夹读写权限，做题记录仍保存在本机');
            }
            return this.sync();
        }
        sync() {
            if (this.running) return this.running;
            const work = () => this.performSync();
            this.running = (global.navigator.locks
                ? global.navigator.locks.request('ielts-onedrive-record-sync', work) : work())
                .finally(() => { this.running = null; });
            return this.running;
        }
        async performSync() {
            // Reload inside the cross-tab lock; another tab may have advanced the baseline.
            this.config = await this.configIO();
            const config = this.config;
            if (!config) return;
            if (await config.folder.queryPermission({ mode: 'readwrite' }) !== 'granted') {
                this.status('连接已记住，等待授权。请点击“立即同步 / 重新授权”。');
                return;
            }
            if (!global.dataRepositories || !global.dataRepositories.practice) {
                this.status('正在等待记录系统就绪，将自动重试。');
                return;
            }
            this.status('正在读取 OneDrive 本地文件并合并记录…');
            const folder = await config.folder.getDirectoryHandle('IELTS-Record-Sync', { create: true });
            let remote = [];
            for await (const [name, handle] of folder.entries()) {
                if (handle.kind !== 'file' || !name.startsWith('records-') || !name.endsWith('.json')) continue;
                try {
                    const text = await (await handle.getFile()).text();
                    // A failed first write can leave our own newly created file empty.
                    // Its complete pending state is still in IndexedDB and can be retried.
                    if (!text && name === 'records-' + config.device + '.json') continue;
                    remote = merge(remote, validate(JSON.parse(text)));
                } catch (error) {
                    throw new Error('无法读取 ' + name + '；未覆盖本地记录。' + error.message);
                }
            }
            const repo = global.dataRepositories.practice;
            let records;
            let changed = false;
            // Use the repository's queue so a submission cannot be overwritten while files are read.
            await repo.dataSource.runTransaction(async tx => {
                const current = await tx.get('practice_records', []);
                const local = observe(config, current, config.device);
                config.entries = merge(local, remote);
                records = config.entries.filter(e => !e.deleted).map(e => clone(e.record))
                    .sort((a, b) => String(b.date || '').localeCompare(String(a.date || '')));
                if (stable(current) !== stable(records)) {
                    changed = true;
                    // Preserve a pre-merge recovery copy before the first remote change in this connection.
                    if (!config.recovery) config.recovery = clone(current);
                    await this.configIO(config);
                    tx.set('practice_records', records);
                }
            });
            // Arrays are stored verbatim by StorageManager. Use the committed snapshot,
            // not a later read that could include a newly submitted, unsynced record.
            config.baseline = clone(records);
            await this.configIO(config);
            const payload = JSON.stringify({ format: FORMAT, entries: config.entries }, null, 2);
            const file = await folder.getFileHandle('records-' + config.device + '.json', { create: true });
            // Do not rewrite unchanged files on every polling cycle.
            if (await (await file.getFile()).text() !== payload) {
                const writable = await file.createWritable();
                try { await writable.write(payload); await writable.close(); }
                catch (error) { try { await writable.abort(); } catch (_) {} throw error; }
            }
            if (changed) {
                if (typeof global.syncPracticeRecords === 'function') await global.syncPracticeRecords({ forceRender: true });
                else {
                    const latest = await repo.list();
                    if (typeof global.setPracticeRecordsState === 'function') global.setPracticeRecordsState(latest);
                    if (global.app?.state?.practice) global.app.state.practice.records = latest.slice();
                }
                if (global.practiceRecorder?.scoreStorage?.recalculateUserStats) {
                    await global.practiceRecorder.scoreStorage.recalculateUserStats();
                }
            }
            this.status('已读写本机 OneDrive 文件夹“' + config.folder.name + '”，共 ' + config.baseline.length +
                ' 条记录（' + new Date().toLocaleTimeString() + '）。云端传输进度请查看 OneDrive。');
        }
        async disconnect() {
            if (this.running) await this.running;
            await this.configIO(null);
            this.config = null;
            this.status('已断开。现有本地记录和 OneDrive 文件均保留。');
        }
        async downloadRecovery() {
            const config = await this.configIO();
            if (!config?.recovery) throw new Error('尚未产生同步前恢复副本');
            const data = { exportInfo: { version: '1.0.0', timestamp: new Date().toISOString() }, practiceRecords: config.recovery };
            const url = URL.createObjectURL(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }));
            const link = global.document.createElement('a');
            link.href = url; link.download = 'ielts-before-onedrive-sync.json'; link.click();
            setTimeout(() => URL.revokeObjectURL(url), 1000);
        }
    }
    global.OneDriveRecordSync = OneDriveRecordSync;
    global.OneDriveRecordSyncModel = { merge, observe, validate, stable, FORMAT };
    function boot() {
        if (!global.document.getElementById('onedrive-connect-btn')) return;
        const service = global.oneDriveRecordSync = new OneDriveRecordSync();
        const report = error => {
            if (error.name !== 'AbortError') service.status('同步未完成：' + error.message + '。本机做题仍可使用。');
        };
        for (const [id, action] of [
            ['onedrive-connect-btn', () => service.connect()],
            ['onedrive-sync-btn', () => service.resume()],
            ['onedrive-disconnect-btn', () => service.disconnect()],
            ['onedrive-recovery-btn', () => service.downloadRecovery()]
        ]) global.document.getElementById(id).addEventListener('click', () => action().catch(report));
        let timer;
        const schedule = () => {
            clearTimeout(timer);
            timer = setTimeout(() => { if (service.config) service.sync().catch(report); }, 1500);
        };
        global.addEventListener('storage-sync', event => {
            if (String(event.detail?.key || '').endsWith('practice_records')) schedule();
        });
        global.addEventListener('focus', schedule);
        setInterval(() => { if (service.config && !global.document.hidden) service.sync().catch(report); }, 15000);
        service.restore().catch(report);
    }
    if (global.document.readyState === 'loading') global.document.addEventListener('DOMContentLoaded', boot);
    else boot();
})(window);
