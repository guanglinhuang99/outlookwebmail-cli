import { mkdtemp, readFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { DatedMailListResult } from '../src/outlook/service.js';
import type { MailSummary } from '../src/types/mail.js';
import { watchMail } from '../src/watch/mail-watcher.js';

function message(stableId: string, subject: string): MailSummary {
  return {
    id: stableId === 'm_old00000000000000000' ? '1' : '2', stableId, subject,
    sender: { name: 'Sender', address: 'sender@example.com' },
    receivedAt: '2026-08-20T08:00:00.000Z', receivedAtText: '今天 16:00', preview: null,
    unread: true, hasAttachments: false,
  };
}

function result(messages: MailSummary[]): DatedMailListResult {
  return {
    date: '2026-08-20', fromDate: '2026-08-20', toDate: '2026-08-20',
    directory: { name: '收件箱', path: '收件箱', level: 0, expanded: true },
    messages, nextCursor: null, hasMore: false,
    filters: { sender: null, subject: null, unread: false, hasAttachments: false },
  };
}

describe('watchMail', () => {
  it('首次建立基线，后续只输出新增邮件并持久化状态', async () => {
    const root = await mkdtemp(join(tmpdir(), 'webmail-watch-'));
    const statePath = join(root, 'watch.json');
    const old = message('m_old00000000000000000', 'old');
    const fresh = message('m_new00000000000000000', 'new');
    let poll = 0;
    const events: unknown[] = [];
    const watchResult = await watchMail({
      listByDate: async () => result(poll++ === 0 ? [old] : [fresh, old]),
    }, {
      iterations: 2,
      intervalSeconds: 5,
      statePath,
      sleep: async () => undefined,
      now: () => new Date('2026-08-20T09:00:00.000Z'),
      onEvent: event => { events.push(event); },
    });

    expect(watchResult).toMatchObject({ polls: 2, emitted: 1 });
    expect(events).toMatchObject([{ type: 'message.new', message: { stableId: fresh.stableId } }]);
    expect(JSON.parse(await readFile(statePath, 'utf8')).scopes['收件箱\u001f2026-08-20']).toEqual([fresh.stableId, old.stableId]);
  });

  it('emitExisting 会在首次轮询输出已有邮件', async () => {
    const root = await mkdtemp(join(tmpdir(), 'webmail-watch-existing-'));
    const events: unknown[] = [];
    await watchMail({ listByDate: async () => result([message('m_old00000000000000000', 'old')]) }, {
      iterations: 1, intervalSeconds: 5, emitExisting: true, statePath: join(root, 'watch.json'),
      onEvent: event => { events.push(event); },
    });
    expect(events).toHaveLength(1);
  });
});
