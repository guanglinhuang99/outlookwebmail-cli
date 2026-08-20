import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { SessionStore, type MailSession } from '../src/session/session-store.js';

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map(path => rm(path, { recursive: true, force: true })));
});

async function createStore(): Promise<{ directory: string; store: SessionStore }> {
  const directory = await mkdtemp(join(tmpdir(), 'webmail-cli-test-'));
  temporaryDirectories.push(directory);
  return { directory, store: new SessionStore(join(directory, 'state', 'session.json')) };
}

const session: MailSession = {
  version: 1,
  updatedAt: '2026-08-20T01:00:00.000Z',
  source: 'inbox',
  messages: {
    '1': {
      subject: '测试邮件',
      senderName: '张三',
      senderAddress: null,
      receivedAt: null,
      receivedAtText: '09:00',
      preview: '测试预览',
      hasAttachments: null,
    },
  },
};

describe('SessionStore', () => {
  it('returns null when no session exists', async () => {
    const { store } = await createStore();
    await expect(store.read()).resolves.toBeNull();
  });

  it('writes atomically and reads a validated session', async () => {
    const { store } = await createStore();
    await store.write(session);

    await expect(store.read()).resolves.toEqual(session);
    expect(JSON.parse(await readFile(store.path, 'utf8'))).toEqual(session);
    expect((await stat(store.path)).mode & 0o777).toBe(0o600);
  });

  it('rejects invalid session JSON', async () => {
    const { directory, store } = await createStore();
    await writeFile(join(directory, 'invalid.json'), '{}');
    const invalidStore = new SessionStore(join(directory, 'invalid.json'));

    await expect(invalidStore.read()).rejects.toMatchObject({ code: 'SESSION_INVALID' });
    expect(store.path).not.toBe(invalidStore.path);
  });

  it('clears an existing session idempotently', async () => {
    const { store } = await createStore();
    await store.write(session);
    await store.clear();
    await store.clear();
    await expect(store.read()).resolves.toBeNull();
  });
});
