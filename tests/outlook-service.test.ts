import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { BrowserBackend } from '../src/browser/backend.js';
import type { InboxParser } from '../src/outlook/inbox-parser.js';
import type { MessageParser } from '../src/outlook/message-parser.js';
import { OutlookService } from '../src/outlook/service.js';
import { SessionStore } from '../src/session/session-store.js';
import type { RawMessageRow } from '../src/types/mail.js';

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map(path => rm(path, { recursive: true, force: true })));
  vi.useRealTimers();
});

function createBackend(searchValue = ''): BrowserBackend {
  return {
    status: vi.fn().mockResolvedValue({
      connected: true,
      taskSpaceId: 1,
      url: 'https://partner.outlook.cn/mail/',
      title: 'Outlook',
      page: { url: 'https://partner.outlook.cn/mail/', title: 'Outlook' },
      snapshot: '收件箱',
    }),
    eval: vi.fn().mockResolvedValue({ count: 1, value: searchValue }),
    clickAndWait: vi.fn().mockResolvedValue(undefined),
    fill: vi.fn().mockResolvedValue(undefined),
    fillAndPress: vi.fn().mockResolvedValue(undefined),
    press: vi.fn().mockResolvedValue(undefined),
    wheel: vi.fn().mockResolvedValue(undefined),
    wait: vi.fn().mockResolvedValue(undefined),
    listInboxFolders: vi.fn().mockResolvedValue({
      accountCount: 1,
      inboxCount: 1,
      complete: true,
      folders: [{ name: '投后', path: '收件箱/投后', level: 3, expanded: null }],
    }),
    selectInboxFolder: vi.fn().mockImplementation(async (directory: string | null) => ({
      count: 1,
      selected: true,
      folder: directory
        ? { name: directory.split('/').at(-1)!, path: directory.includes('/') ? directory : `收件箱/${directory}`, level: 3, expanded: null }
        : { name: '收件箱', path: '收件箱', level: 2, expanded: true },
    })),
    deleteMessage: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true,
    }),
    moveMessage: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true,
    }),
    downloadAttachment: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true,
      attachmentCount: 1, attachmentId: '1', filename: 'report.xlsx', path: '/tmp/report.xlsx', bytes: 12,
    }),
  } as unknown as BrowserBackend;
}

function createParser(rows: RawMessageRow[]): InboxParser {
  return {
    extract: vi.fn().mockResolvedValue(rows),
    resetAndExtract: vi.fn().mockResolvedValue(rows),
    scrollAndExtract: vi.fn().mockResolvedValue(rows),
  };
}

const rows: RawMessageRow[] = [
  {
    stableHint: null,
    senderName: '已读用户',
    senderAddress: 'read@example.com',
    subject: '已读邮件',
    receivedAt: '2026-08-20T09:00:00+08:00',
    receivedAtText: '9:00',
    preview: '已读预览',
    unread: false,
    hasAttachments: false,
  },
  {
    stableHint: null,
    senderName: '未读用户',
    senderAddress: 'unread@example.com',
    subject: '未读邮件',
    receivedAt: '2026-08-20T09:30:00+08:00',
    receivedAtText: '9:30',
    preview: '未读预览',
    unread: true,
    hasAttachments: true,
  },
];

async function createService(searchValue = ''): Promise<{
  backend: BrowserBackend;
  parser: InboxParser;
  store: SessionStore;
  service: OutlookService;
}> {
  const directory = await mkdtemp(join(tmpdir(), 'webmail-service-test-'));
  temporaryDirectories.push(directory);
  const backend = createBackend(searchValue);
  const parser = createParser(rows);
  const store = new SessionStore(join(directory, 'session.json'));
  return { backend, parser, store, service: new OutlookService(backend, store, parser) };
}

describe('OutlookService mail listing', () => {
  it('lists every message received today and writes a today session', async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-08-20T02:00:00.000Z'));
    const { service, store } = await createService();

    const result = await service.today();

    expect(result.messages).toHaveLength(2);
    expect(result.directory).toMatchObject({ path: '收件箱' });
    await expect(store.read()).resolves.toMatchObject({ source: 'today:收件箱:2026-08-20' });
  });

  it('lists today messages from an explicitly selected Inbox child directory', async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-08-20T02:00:00.000Z'));
    const { service, backend, store } = await createService();

    const result = await service.today('收件箱/投后');

    expect(backend.selectInboxFolder).toHaveBeenCalledWith('收件箱/投后');
    expect(result.directory).toMatchObject({ path: '收件箱/投后' });
    await expect(store.read()).resolves.toMatchObject({ source: 'today:收件箱/投后:2026-08-20' });
  });

  it('returns validated Outlook folders', async () => {
    const { service, backend } = await createService();

    await expect(service.folders()).resolves.toEqual({
      root: '收件箱',
      folders: [{ name: '投后', path: '收件箱/投后', level: 3, expanded: null }],
    });
    expect(backend.listInboxFolders).toHaveBeenCalledOnce();
  });

  it('filters unread rows, assigns short IDs and writes the session', async () => {
    const { service, store } = await createService();
    const result = await service.inbox({ limit: 10, unreadOnly: true });

    expect(result.messages).toHaveLength(1);
    expect(result.directory).toMatchObject({ path: '收件箱' });
    expect(result.messages[0]).toMatchObject({ id: '1', subject: '未读邮件', unread: true });
    await expect(store.read()).resolves.toMatchObject({ source: 'folder:收件箱' });
  });

  it('selects an Inbox child directory before listing its messages', async () => {
    const { service, backend, store } = await createService();
    const result = await service.inbox({ limit: 10, directory: '投后' });

    expect(backend.selectInboxFolder).toHaveBeenCalledWith('投后');
    expect(result.directory).toMatchObject({ name: '投后', path: '收件箱/投后' });
    await expect(store.read()).resolves.toMatchObject({ source: 'folder:收件箱/投后' });
  });

  it('treats an empty directory argument as Inbox', async () => {
    const { service, backend } = await createService();
    const result = await service.inbox({ limit: 1, directory: '   ' });

    expect(backend.selectInboxFolder).toHaveBeenCalledWith(null);
    expect(result.directory).toMatchObject({ path: '收件箱' });
  });

  it('rejects an ambiguous folder name and requests a full path', async () => {
    const { service, backend } = await createService();
    vi.mocked(backend.selectInboxFolder).mockResolvedValueOnce({ count: 2, selected: false, folder: null });

    await expect(service.inbox({ limit: 10, directory: '重复目录' }))
      .rejects.toMatchObject({ code: 'AMBIGUOUS_FOLDER' });
  });

  it('submits the original normalized query to Outlook search', async () => {
    const { service, backend, store } = await createService();
    await service.search('  风险报告  ', { limit: 1 });

    expect(backend.fillAndPress).toHaveBeenCalledWith(expect.stringContaining('combobox'), '风险报告', 'Enter', 1_500);
    await expect(store.read()).resolves.toMatchObject({ source: 'search:风险报告' });
  });

  it('exits an active Outlook search before listing the inbox', async () => {
    const { service, backend } = await createService('风险报告');
    vi.mocked(backend.eval)
      .mockResolvedValueOnce({ count: 1, value: '风险报告' })
      .mockResolvedValueOnce({ count: 1, value: '' });

    await service.inbox({ limit: 1 });
    expect(backend.clickAndWait).toHaveBeenCalledWith(expect.stringContaining('退出搜索'), 1_000);
  });

  it('reads one unique session message and numbers its attachments', async () => {
    const { backend, parser, store } = await createService();
    await store.write({
      version: 1,
      updatedAt: '2026-08-20T01:00:00.000Z',
      source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: '张三', senderAddress: 'zhangsan@example.com',
          receivedAt: '2026-08-20T09:30:00+08:00',
          receivedAtText: '9:30', preview: '测试预览', hasAttachments: true,
        },
      },
    });
    const messageParser: MessageParser = {
      openAndExtract: vi.fn().mockResolvedValue({
        matchCount: 1,
        message: {
          subject: '测试主题', fromName: '张三', fromAddress: 'zhangsan@example.com',
          to: [], cc: [], receivedAt: '2026-08-20T09:30:00+08:00',
          receivedAtText: '周四 2026/8/20 9:30', bodyText: '测试正文',
          attachments: [{ filename: 'report.xlsx', sizeText: '12 KB' }],
        },
      }),
    };
    const service = new OutlookService(backend, store, parser, messageParser);

    await expect(service.read('1')).resolves.toMatchObject({
      id: '1', subject: '测试主题', attachments: [{ id: '1', filename: 'report.xlsx', sizeText: '12 KB' }],
    });
  });

  it('rejects multiple distinct candidates as ambiguous', async () => {
    const { backend, parser, store } = await createService();
    await store.write({
      version: 1,
      updatedAt: '2026-08-20T01:00:00.000Z',
      source: 'inbox',
      messages: {
        '1': {
          subject: '重复主题', senderName: null, senderAddress: null, receivedAt: null,
          receivedAtText: null, preview: null, hasAttachments: null,
        },
      },
    });
    const messageParser: MessageParser = {
      openAndExtract: vi.fn().mockResolvedValue({ matchCount: 2, message: null }),
    };
    const service = new OutlookService(backend, store, parser, messageParser);

    await expect(service.read('1')).rejects.toMatchObject({ code: 'AMBIGUOUS_MESSAGE' });
  });

  it('requires explicit confirmation before deleting or moving', async () => {
    const { service, backend } = await createService();
    await expect(service.delete('1', false)).rejects.toMatchObject({ code: 'CONFIRMATION_REQUIRED' });
    await expect(service.move('1', '投后', false)).rejects.toMatchObject({ code: 'CONFIRMATION_REQUIRED' });
    expect(backend.deleteMessage).not.toHaveBeenCalled();
    expect(backend.moveMessage).not.toHaveBeenCalled();
  });

  it('moves a session message only after confirmation', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: null, senderAddress: null, receivedAt: null,
          receivedAtText: null, preview: null, hasAttachments: false,
        },
      },
    });

    await expect(service.move('1', ' 投后 ', true)).resolves.toMatchObject({ moved: true, folder: '投后' });
    expect(backend.moveMessage).toHaveBeenCalledWith(expect.objectContaining({ subject: '测试主题' }), '投后');
    await expect(store.read()).resolves.toBeNull();
  });

  it('returns the verified local attachment download result', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: null, senderAddress: null, receivedAt: null,
          receivedAtText: null, preview: null, hasAttachments: true,
        },
      },
    });
    const directory = temporaryDirectories[temporaryDirectories.length - 1]!;

    await expect(service.downloadAttachment('1', '1', directory)).resolves.toMatchObject({
      status: 'performed', filename: 'report.xlsx', verified: true,
    });
    expect(backend.downloadAttachment).toHaveBeenCalledWith(expect.any(Object), 0, directory);
  });
});
