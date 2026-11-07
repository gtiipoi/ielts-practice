import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
import { test } from 'node:test';

const source = fs.readFileSync(new URL('../../../js/utils/oneDriveRecordSync.js', import.meta.url), 'utf8');
const copy = x => JSON.parse(JSON.stringify(x));
function folder() {
    const files = new Map();
    let writes = 0;
    let allowed = true;
    let failWrite = false;
    const dir = {
        name: 'OneDrive-test', files,
        queryPermission: async () => allowed ? 'granted' : 'prompt',
        getDirectoryHandle: async () => dir,
        async *entries() {
            for (const name of files.keys()) yield [name, await dir.getFileHandle(name)];
        },
        async getFileHandle(name) {
            if (!files.has(name)) files.set(name, '');
            return {
                kind: 'file', getFile: async () => ({ text: async () => files.get(name) }),
                createWritable: async () => {
                    let content;
                    return {
                        write: async text => { if (failWrite) throw new Error('disk full'); content = text; },
                        close: async () => { files.set(name, content); writes++; }, abort: async () => {}
                    };
                }
            };
        },
        get writes() { return writes; },
        deny() { allowed = false; }, grant() { allowed = true; },
        fail(value) { failWrite = value; }
    };
    return dir;
}
function harness(dir, device, initial = []) {
    let records = copy(initial);
    let connection = { folder: dir, device, entries: [], baseline: [] };
    const statuses = [];
    const context = {
        console, setTimeout, clearTimeout, setInterval: () => 0, Date,
        navigator: {}, document: { readyState: 'loading', addEventListener() {},
            getElementById: () => ({ set textContent(value) { statuses.push(value); } }) },
        dataRepositories: { practice: {
            list: async () => copy(records),
            dataSource: { runTransaction: async fn => {
                let pending;
                await fn({ get: async () => copy(records), set: (_, value) => { pending = copy(value); } });
                if (pending) records = pending;
            } }
        } }
    };
    context.window = context;
    vm.runInNewContext(source, context);
    const service = new context.OneDriveRecordSync();
    service.configIO = async value => {
        if (value !== undefined) connection = value === null ? null : { ...copy({ ...value, folder: null }), folder: dir };
        return connection === null ? null : { ...copy({ ...connection, folder: null }), folder: dir };
    };
    service.config = connection;
    return { service, model: context.OneDriveRecordSyncModel, statuses,
        records: () => copy(records), set: value => { records = copy(value); },
        config: () => connection };
}
const record = (id, score = 1) => ({ id, score, date: '2026-11-07', type: 'reading',
    realData: { answers: { q1: 'answer' }, highlights: ['preserved'] } });

test('two devices merge existing history and full answer data without duplicates', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('a')]), b = harness(dir, 'b', [record('b')]);
    await a.service.sync(); await b.service.sync(); await a.service.sync();
    assert.deepEqual(a.records(), b.records());
    assert.equal(a.records().length, 2);
    assert.deepEqual(a.records()[0].realData.highlights, ['preserved']);
    const writes = dir.writes;
    await a.service.sync(); await b.service.sync(); await a.service.sync();
    assert.equal(a.records().length, 2);
    assert.equal(dir.writes, writes);
});
test('new records on both devices survive alternating synchronization', async () => {
    const dir = folder(), a = harness(dir, 'a'), b = harness(dir, 'b');
    await a.service.sync(); await b.service.sync();
    a.set([record('a')]); b.set([record('b')]);
    await b.service.sync(); await a.service.sync(); await b.service.sync();
    assert.deepEqual(a.records(), b.records()); assert.equal(a.records().length, 2);
});
test('deletion propagates and an unchanged stale device cannot resurrect a record', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('old')]), b = harness(dir, 'b');
    await a.service.sync(); await b.service.sync();
    a.set([]); await a.service.sync(); await b.service.sync(); await a.service.sync();
    assert.deepEqual(b.records(), []); assert.deepEqual(a.records(), []);
});
test('record edit propagates with the same ID', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('x')]), b = harness(dir, 'b');
    await a.service.sync(); await b.service.sync();
    b.set([record('x', 9)]); await b.service.sync(); await a.service.sync();
    assert.equal(a.records()[0].score, 9);
});
test('concurrent conflicting edits converge deterministically', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('x')]), b = harness(dir, 'b');
    await a.service.sync(); await b.service.sync();
    a.set([record('x', 3)]); b.set([record('x', 7)]);
    await a.service.sync(); await b.service.sync(); await a.service.sync();
    assert.deepEqual(a.records(), b.records());
});
test('corrupt file stops writes and preserves existing browser records', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('safe')]);
    dir.files.set('records-broken.json', '{');
    await assert.rejects(a.service.sync(), /records-broken/);
    assert.equal(dir.writes, 0); assert.deepEqual(a.records(), [record('safe')]);
});
test('revoked permission leaves local records untouched and can be retried', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('safe')]); dir.deny();
    await a.service.sync(); assert.equal(dir.writes, 0); assert.equal(a.records().length, 1);
    dir.grant(); await a.service.sync(); assert.equal(dir.writes, 1);
});
test('failed file write retains changes for retry', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('safe')]); dir.fail(true);
    await assert.rejects(a.service.sync(), /disk full/);
    assert.equal(a.records().length, 1);
    dir.fail(false); await a.service.sync();
    const b = harness(dir, 'b'); await b.service.sync(); assert.equal(b.records().length, 1);
});
test('recovery copy preserves pre-merge local data; disconnect retains both copies', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('a')]), b = harness(dir, 'b', [record('b')]);
    await a.service.sync(); await b.service.sync();
    assert.deepEqual(b.config().recovery, [record('b')]);
    await b.service.disconnect(); assert.equal(b.config(), null);
    assert.equal(b.records().length, 2); assert.equal(dir.files.size, 2);
});
test('invalid schema rejected, merge is order-independent and idempotent', () => {
    const h = harness(folder(), 'a'), m = h.model;
    assert.throws(() => m.validate({ format: m.FORMAT, entries: [{ id: 'x' }] }));
    const a = [{ id: 'x', clock: 2, device: 'a', deleted: true }];
    const b = [{ id: 'x', clock: 2, device: 'b', deleted: false, record: record('x') }];
    assert.equal(m.stable(m.merge(a, b)), m.stable(m.merge(b, a)));
    assert.equal(m.stable(m.merge(a, a)), m.stable(a));
});

test('a submission arriving while remote files are read is included in the merge', async () => {
    const dir = folder(), a = harness(dir, 'a'), b = harness(dir, 'b', [record('remote')]);
    await b.service.sync();
    const original = dir.entries;
    dir.entries = async function* () {
        a.set([record('during-read')]);
        yield* original();
    };
    await a.service.sync();
    assert.deepEqual(a.records().map(r => r.id).sort(), ['during-read', 'remote']);
});

test('a submission just after commit remains a new change on the next sync', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('first')]);
    const configIO = a.service.configIO;
    let inserted = false;
    a.service.configIO = async value => {
        if (!inserted && value?.baseline?.length) {
            inserted = true;
            a.set([...a.records(), record('after-commit')]);
        }
        return configIO(value);
    };
    await a.service.sync(); await a.service.sync();
    const b = harness(dir, 'b'); await b.service.sync();
    assert.deepEqual(b.records().map(r => r.id).sort(), ['after-commit', 'first']);
});

test('restart can restore the persisted baseline and detect offline deletions', async () => {
    const dir = folder(), a = harness(dir, 'a', [record('x')]);
    await a.service.sync();
    const restarted = harness(dir, 'a', []);
    restarted.service.configIO = a.service.configIO;
    await restarted.service.restore();
    const b = harness(dir, 'b'); await b.service.sync();
    assert.deepEqual(b.records(), []);
});
