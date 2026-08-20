import { appendFile, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import type { DatedMailListResult } from '../src/outlook/service.js';
import type { MailMessage, MailSummary } from '../src/types/mail.js';
import { syncObsidian, type ObsidianSyncSource } from '../src/sync/obsidian-sync.js';

const stableId = 'm_12345678901234567890';

function summary(): MailSummary {
  return {
    id: '1', stableId, sender: { name: 'Alice', address: 'alice@example.com' }, subject: 'P2 report',
    receivedAt: '2026-08-20T08:00:00.000Z', receivedAtText: '今天 16:00', preview: 'body', unread: false,
    hasAttachments: true,
  };
}

function listResult(): DatedMailListResult {
  return {
    date: '2026-08-20', fromDate: '2026-08-20', toDate: '2026-08-20',
    directory: { name: '收件箱', path: '收件箱', level: 0, expanded: true },
    messages: [summary()], hasMore: false, nextCursor: null,
    filters: { sender: null, subject: null, unread: false, hasAttachments: false },
  };
}

function mail(bodyText: string): MailMessage {
  return {
    id: stableId, stableId, subject: 'P2 report', from: { name: 'Alice', address: 'alice@example.com' },
    to: [{ name: 'Bob', address: 'bob@example.com' }], cc: [],
    receivedAt: '2026-08-20T08:00:00.000Z', receivedAtText: '今天 16:00', bodyText,
    attachments: [{ id: '1', filename: 'chart.png', sizeText: '3 B' }],
  };
}

describe('syncObsidian', () => {
  it('创建、去重、更新 Markdown，保留用户尾注并生成图片和附件索引', async () => {
    const root = await mkdtemp(join(tmpdir(), 'webmail-sync-'));
    let body = 'first body';
    const downloadAttachment = vi.fn(async (_id: string, _attachmentId: string, outputDirectory: string) => {
      const path = join(outputDirectory, 'chart.png');
      await writeFile(path, 'png');
      return { matchCount: 1, status: 'performed' as const, performed: true, verified: true, filename: 'chart.png', path, bytes: 3 };
    });
    const source: ObsidianSyncSource = {
      listByDate: async () => listResult(),
      read: async () => mail(body),
      downloadAttachment,
    };

    const created = await syncObsidian(source, {}, root);
    expect(created).toMatchObject({ created: 1, updated: 0, unchanged: 0 });
    expect(await readFile(created.items[0]!.markdownPath, 'utf8')).toContain('![chart.png]');
    expect(await readFile(created.attachmentIndexPath, 'utf8')).toContain('chart.png');

    const unchanged = await syncObsidian(source, {}, root);
    expect(unchanged).toMatchObject({ created: 0, updated: 0, unchanged: 1 });
    expect(downloadAttachment).toHaveBeenCalledTimes(1);

    await appendFile(created.items[0]!.markdownPath, '\n用户笔记\n');
    body = 'updated body';
    const updated = await syncObsidian(source, {}, root);
    expect(updated).toMatchObject({ created: 0, updated: 1, unchanged: 0 });
    const markdown = await readFile(updated.items[0]!.markdownPath, 'utf8');
    expect(markdown).toContain('updated body');
    expect(markdown).toContain('用户笔记');
    expect(downloadAttachment).toHaveBeenCalledTimes(2);
  });
});
