import { mkdtemp, readFile, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { MutationStore } from '../src/safety/mutation-store.js';

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map(path => rm(path, { recursive: true, force: true })));
});

async function createStore(): Promise<MutationStore> {
  const directory = await mkdtemp(join(tmpdir(), 'webmail-mutations-test-'));
  temporaryDirectories.push(directory);
  return new MutationStore(join(directory, 'mutations.json'), join(directory, 'audit.jsonl'));
}

describe('MutationStore', () => {
  it('preserves different request records started concurrently', async () => {
    const store = await createStore();
    await Promise.all([
      store.begin('request-001', 'move', store.payloadHash('move', { id: '1' }), 'm_aaaaaaaaaaaaaaaaaaaa'),
      store.begin('request-002', 'delete', store.payloadHash('delete', { id: '2' }), 'm_bbbbbbbbbbbbbbbbbbbb'),
    ]);

    const state = JSON.parse(await readFile(store.path, 'utf8')) as { records: Record<string, unknown> };
    expect(Object.keys(state.records).sort()).toEqual(['request-001', 'request-002']);
  });

  it('refuses to retry a request that is still pending', async () => {
    const store = await createStore();
    const hash = store.payloadHash('delete', { id: '1' });
    await store.begin('request-pending', 'delete', hash, 'm_aaaaaaaaaaaaaaaaaaaa');

    await expect(store.prior('request-pending', 'delete', hash, '1'))
      .rejects.toMatchObject({ code: 'OPERATION_UNKNOWN' });
  });

  it('preserves an uncertain result and forbids automatic retry', async () => {
    const store = await createStore();
    const hash = store.payloadHash('delete', { id: '1' });
    await store.begin('request-unknown', 'delete', hash, 'm_aaaaaaaaaaaaaaaaaaaa');
    await store.uncertain('request-unknown', 'PLAYWRIGHT_TIMEOUT', { action: 'delete', mailId: 'm_aaaaaaaaaaaaaaaaaaaa' });
    await expect(store.prior('request-unknown', 'delete', hash, '1'))
      .rejects.toMatchObject({ code: 'OPERATION_UNKNOWN', message: expect.stringContaining('禁止') });
  });
});
