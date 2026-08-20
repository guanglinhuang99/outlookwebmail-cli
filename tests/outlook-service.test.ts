import { createHash } from 'node:crypto';
import { mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { BrowserBackend } from '../src/browser/backend.js';
import type { InboxParser } from '../src/outlook/inbox-parser.js';
import type { MessageParser } from '../src/outlook/message-parser.js';
import { OutlookService } from '../src/outlook/service.js';
import { SessionStore } from '../src/session/session-store.js';
import { MutationStore } from '../src/safety/mutation-store.js';
import type { RawMessageRow } from '../src/types/mail.js';

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map(path => rm(path, { recursive: true, force: true })));
  vi.useRealTimers();
});

function createBackend(searchValue = ''): BrowserBackend {
  let currentSearch = searchValue;
  const backend = {
    status: vi.fn().mockResolvedValue({
      connected: true,
      taskSpaceId: 1,
      url: 'https://partner.outlook.cn/mail/',
      title: 'Outlook',
      page: { url: 'https://partner.outlook.cn/mail/', title: 'Outlook' },
      snapshot: '收件箱',
    }),
    handoffForLogin: vi.fn().mockResolvedValue({
      taskSpaceId: 1,
      url: 'https://partner.outlook.cn/mail/',
      handedOff: true,
    }),
    waitUntilMailReady: vi.fn().mockResolvedValue({
      ready: true, url: 'https://partner.outlook.cn/mail/', title: 'Outlook',
      searchInputs: 1, inboxFolders: 1, mailLists: 1, loginFrames: 0, busy: false,
    }),
    eval: vi.fn().mockImplementation(async (script: string) => script.includes('firstKey')
      ? { count: 1, value: currentSearch, listCount: 1, firstKey: 'row-1', busy: false, empty: false }
      : { count: 1, value: currentSearch }),
    clickAndWait: vi.fn().mockImplementation(async () => { currentSearch = ''; }),
    fill: vi.fn().mockResolvedValue(undefined),
    fillAndPress: vi.fn().mockImplementation(async (_selector: string, value: string) => { currentSearch = value; }),
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
    replyMessage: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'draft_ready', performed: true, verified: true,
      draft: true, replyAll: false, handedOff: true,
    }),
    composeMessage: vi.fn().mockResolvedValue({
      status: 'draft_ready', performed: true, verified: true, draft: true,
      handedOff: true, attachmentCount: 0,
    }),
    forwardMessage: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'draft_ready', performed: true, verified: true,
      draft: true, handedOff: true, attachmentCount: 0,
    }),
    selectSystemFolder: vi.fn().mockResolvedValue({
      count: 1, selected: true,
      folder: { name: '草稿', path: '草稿', level: 2, expanded: null },
    }),
    openDraft: vi.fn().mockResolvedValue({
      matchCount: 1,
      closed: true,
      draft: {
        to: [{ name: null, address: 'to@example.com' }], cc: [], bcc: [],
        subject: '草稿主题', bodyText: '草稿正文', attachments: [],
      },
    }),
    updateDraft: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'draft_ready', performed: true, verified: true,
      draft: true, handedOff: false, attachmentCount: 0,
    }),
    sendDraft: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'sent', performed: true, verified: true,
      draft: false, handedOff: false, attachmentCount: 0,
    }),
    discardDraft: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true,
    }),
    setReadState: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true, changed: true,
    }),
    setFlagState: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true, changed: true,
    }),
    setCategoryState: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true, changed: true,
    }),
    getConversation: vi.fn().mockResolvedValue({
      matchCount: 1, complete: true,
      messages: [{
        subject: '测试主题', fromName: '张三', fromAddress: 'zhangsan@example.com',
        to: [], cc: [], receivedAt: null, receivedAtText: null,
        bodyText: '会话正文', attachments: [],
      }],
    }),
    downloadAttachment: vi.fn().mockResolvedValue({
      matchCount: 1, status: 'performed', performed: true, verified: true,
      attachmentCount: 1, attachmentId: '1', filename: 'report.xlsx', path: '/tmp/report.xlsx', bytes: 12,
    }),
  } as unknown as BrowserBackend;
  return backend;
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
  mutationStore: MutationStore;
  service: OutlookService;
}> {
  const directory = await mkdtemp(join(tmpdir(), 'webmail-service-test-'));
  temporaryDirectories.push(directory);
  const backend = createBackend(searchValue);
  const parser = createParser(rows);
  const store = new SessionStore(join(directory, 'session.json'));
  const mutationStore = new MutationStore(join(directory, 'mutations.json'), join(directory, 'audit.jsonl'));
  return { backend, parser, store, mutationStore, service: new OutlookService(backend, store, parser, undefined, mutationStore) };
}

describe('OutlookService mail listing', () => {
  it('returns status without handing off when Outlook is already signed in', async () => {
    const { service, backend } = await createService();

    await expect(service.status()).resolves.toMatchObject({ state: 'INBOX', mailReady: true });
    expect(backend.handoffForLogin).not.toHaveBeenCalled();
  });

  it('does not report an authenticated shell as mail-ready', async () => {
    const { service, backend } = await createService();
    vi.mocked(backend.waitUntilMailReady).mockResolvedValueOnce({
      ready: false, url: 'https://partner.outlook.cn/mail/', title: 'Outlook',
      searchInputs: 0, inboxFolders: 0, mailLists: 0, loginFrames: 1, busy: true,
    });

    await expect(service.status()).rejects.toMatchObject({
      code: 'OUTLOOK_NOT_READY',
      message: expect.stringContaining('登录 iframe=1'),
    });
  });

  it('turns a redirect to the login host during readiness waiting into AUTH_REQUIRED', async () => {
    const { service, backend } = await createService();
    vi.mocked(backend.waitUntilMailReady).mockResolvedValueOnce({
      ready: false, url: 'https://login.partner.microsoftonline.cn/authorize?login_hint=secret', title: 'Sign in',
      searchInputs: 0, inboxFolders: 0, mailLists: 0, loginFrames: 0, busy: false,
    });
    await expect(service.status()).rejects.toMatchObject({ code: 'AUTH_REQUIRED' });
    expect(backend.handoffForLogin).toHaveBeenCalledOnce();
  });

  it('opens and hands off Outlook when authentication is required', async () => {
    const { service, backend } = await createService();
    vi.mocked(backend.status).mockResolvedValueOnce({
      connected: true,
      taskSpaceId: 1,
      url: 'https://login.partner.microsoftonline.cn/',
      title: 'Sign in',
      page: { url: 'https://login.partner.microsoftonline.cn/', title: 'Sign in' },
      snapshot: 'Sign in Password',
    });

    await expect(service.status()).rejects.toMatchObject({
      code: 'AUTH_REQUIRED',
      message: expect.stringContaining('将页面交给你'),
    });
    expect(backend.handoffForLogin).toHaveBeenCalledOnce();
  });

  it('keeps AUTH_REQUIRED when login handoff itself fails', async () => {
    const { service, backend } = await createService();
    vi.mocked(backend.status).mockResolvedValueOnce({
      connected: true,
      taskSpaceId: 1,
      url: 'https://login.partner.microsoftonline.cn/',
      title: 'Sign in',
      page: { url: 'https://login.partner.microsoftonline.cn/', title: 'Sign in' },
      snapshot: '登录 密码',
    });
    vi.mocked(backend.handoffForLogin).mockRejectedValueOnce(new Error('handoff unavailable'));

    await expect(service.status()).rejects.toMatchObject({
      code: 'AUTH_REQUIRED',
      message: expect.stringContaining('控制权交接失败'),
    });
  });

  it('lists every message received today and writes a today session', async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-08-20T02:00:00.000Z'));
    const { service, store } = await createService();

    const result = await service.today();

    expect(result.messages).toHaveLength(2);
    expect(result.date).toBe('2026-08-20');
    expect(result).toMatchObject({ searchQuery: 'received:2026-08-20', scope: 'selected-folder' });
    expect(result.directory).toMatchObject({ path: '收件箱' });
    await expect(store.read()).resolves.toMatchObject({ source: 'date:收件箱:2026-08-20' });
  });

  it('preserves visually identical rows and assigns collision-safe stable IDs', async () => {
    const { service, parser } = await createService();
    const duplicate: RawMessageRow = {
      stableHint: null, senderName: '机器人', senderAddress: 'robot@example.com', subject: '重复告警',
      receivedAt: '2026-08-20T09:00:00+08:00', receivedAtText: '9:00', preview: '相同内容',
      unread: true, hasAttachments: false,
    };
    vi.mocked(parser.resetAndExtract).mockResolvedValueOnce([duplicate, { ...duplicate }]);
    vi.mocked(parser.scrollAndExtract).mockResolvedValue([duplicate, { ...duplicate }]);

    const result = await service.listByDate({ date: '2026-08-20', limit: 10 });
    expect(result.messages).toHaveLength(2);
    expect(new Set(result.messages.map(message => message.stableId))).toHaveProperty('size', 2);
  });

  it('lists today messages from an explicitly selected Inbox child directory', async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-08-20T02:00:00.000Z'));
    const { service, backend, store } = await createService();

    const result = await service.today('收件箱/投后');

    expect(backend.selectInboxFolder).toHaveBeenCalledWith('收件箱/投后');
    expect(result.directory).toMatchObject({ path: '收件箱/投后' });
    await expect(store.read()).resolves.toMatchObject({ source: 'date:收件箱/投后:2026-08-20' });
  });

  it('lists a specified date from a specified Inbox child directory', async () => {
    const { service, backend, parser, store } = await createService();
    const previousDay: RawMessageRow = {
      ...rows[0]!, subject: '昨日邮件', receivedAt: '2026-08-19T16:00:00+08:00', receivedAtText: '昨日 16:00',
    };
    const older: RawMessageRow = {
      ...rows[0]!, subject: '更早邮件', receivedAt: '2026-08-18T16:00:00+08:00', receivedAtText: '周二 16:00',
    };
    vi.mocked(parser.resetAndExtract).mockResolvedValueOnce(rows);
    vi.mocked(parser.scrollAndExtract).mockResolvedValueOnce([previousDay, older]);

    const result = await service.listByDate({ date: '2026-08-19', directory: '收件箱/投后' });

    expect(result).toMatchObject({ date: '2026-08-19', directory: { path: '收件箱/投后' } });
    expect(result.messages.map(message => message.subject)).toEqual(['昨日邮件']);
    expect(backend.selectInboxFolder).toHaveBeenCalledWith('收件箱/投后');
    await expect(store.read()).resolves.toMatchObject({ source: 'date:收件箱/投后:2026-08-19' });
  });

  it('treats empty date and directory options as today and Inbox', async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-08-20T02:00:00.000Z'));
    const { service, backend } = await createService();

    const result = await service.listByDate({ date: '   ', directory: '   ' });

    expect(result.date).toBe('2026-08-20');
    expect(result.directory).toMatchObject({ path: '收件箱' });
    expect(backend.selectInboxFolder).toHaveBeenCalledWith(null);
  });

  it('rejects malformed and impossible dates before opening Outlook', async () => {
    const { service, backend } = await createService();

    await expect(service.listByDate({ date: '20-08-2026' }))
      .rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    await expect(service.listByDate({ date: '2026-02-30' }))
      .rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    expect(backend.status).not.toHaveBeenCalled();
  });

  it('paginates a filtered date range with an opaque cursor and absolute short IDs', async () => {
    const { service, parser } = await createService();
    const anotherUnread: RawMessageRow = {
      ...rows[1]!,
      subject: '另一封未读报告',
      receivedAt: '2026-08-19T12:00:00+08:00',
      receivedAtText: '昨日 12:00',
      preview: '第二页',
    };
    vi.mocked(parser.resetAndExtract).mockResolvedValue([rows[0]!, rows[1]!, anotherUnread]);
    vi.mocked(parser.scrollAndExtract).mockResolvedValue([rows[0]!, rows[1]!, anotherUnread]);

    const first = await service.listByDate({
      fromDate: '2026-08-19', toDate: '2026-08-20', sender: 'unread@example.com',
      unread: true, hasAttachments: true, limit: 1,
    });
    expect(first).toMatchObject({ date: null, fromDate: '2026-08-19', toDate: '2026-08-20', hasMore: true });
    expect(first.messages).toHaveLength(1);
    expect(first.nextCursor).toMatch(/^[A-Za-z0-9_-]+$/);

    const second = await service.listByDate({
      fromDate: '2026-08-19', toDate: '2026-08-20', sender: 'unread@example.com',
      unread: true, hasAttachments: true, limit: 1, cursor: first.nextCursor,
    });
    expect(second.messages).toHaveLength(1);
    expect(second.messages[0]?.id).toBe('2');
    expect(second.hasMore).toBe(false);
    expect(second.nextCursor).toBeNull();
  });

  it('rejects a cursor when its directory or filters do not match', async () => {
    const { service } = await createService();
    const first = await service.listByDate({ date: '2026-08-20', limit: 1 });

    await expect(service.listByDate({ date: '2026-08-20', limit: 1, unread: true, cursor: first.nextCursor }))
      .rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
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
    expect(result.messages[0]?.stableId).toMatch(/^m_[A-Za-z0-9_-]{20}$/);
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

    expect(backend.fillAndPress).toHaveBeenCalledWith(expect.stringContaining('combobox'), '风险报告', 'Enter', 0);
    await expect(store.read()).resolves.toMatchObject({ source: 'search:风险报告' });
  });

  it('exits an active Outlook search before listing the inbox', async () => {
    const { service, backend } = await createService('风险报告');
    vi.mocked(backend.eval)
      .mockResolvedValueOnce({ count: 1, value: '风险报告' })
      .mockResolvedValueOnce({ count: 1, value: '' });

    await service.inbox({ limit: 1 });
    expect(backend.clickAndWait).toHaveBeenCalledWith(expect.stringContaining('退出搜索'), 0);
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

  it('restores unread state after reading an unread session message', async () => {
    const { backend, parser, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'search:report',
      messages: {
        '1': {
          stableHint: 'item-1', subject: '未读报告', senderName: '张三', senderAddress: 'zhangsan@example.com',
          receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30', preview: '预览',
          hasAttachments: false, unread: true,
        },
      },
    });
    const messageParser: MessageParser = {
      openAndExtract: vi.fn().mockResolvedValue({
        matchCount: 1,
        message: {
          subject: '未读报告', fromName: '张三', fromAddress: 'zhangsan@example.com', to: [], cc: [],
          receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30', bodyText: '正文', attachments: [],
        },
      }),
    };
    const service = new OutlookService(backend, store, parser, messageParser);

    await expect(service.read('1')).resolves.toMatchObject({ unreadRestored: true });
    expect(backend.setReadState).toHaveBeenCalledWith(expect.objectContaining({ stableHint: 'item-1', unread: true }), true);
  });

  it('keeps stable IDs resolvable across list refreshes', async () => {
    const { service, parser, backend } = await createService();
    const first = await service.inbox({ limit: 2 });
    const stableId = first.messages[0]!.stableId;
    const replacement: RawMessageRow = {
      ...rows[0]!, subject: '刷新后的邮件', receivedAt: '2026-08-20T10:00:00+08:00', receivedAtText: '10:00',
    };
    vi.mocked(parser.resetAndExtract).mockResolvedValueOnce([replacement]);
    await service.inbox({ limit: 1 });
    const directory = temporaryDirectories[temporaryDirectories.length - 1]!;

    await service.downloadAttachment(stableId, '1', directory);
    expect(backend.downloadAttachment).toHaveBeenCalledWith(
      expect.objectContaining({ subject: '已读邮件' }), 0, directory,
    );
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

    await expect(service.move('1', ' 投后 ', true, 'move-test-001')).resolves.toMatchObject({
      moved: true, folder: '投后', requestId: 'move-test-001', deduplicated: false,
    });
    expect(backend.moveMessage).toHaveBeenCalledWith(expect.objectContaining({ subject: '测试主题' }), '投后');
    await expect(store.read()).resolves.toBeNull();
  });

  it('deduplicates a repeated move request and writes a redacted audit', async () => {
    const { service, backend, store, mutationStore } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '敏感测试主题', senderName: '张三', senderAddress: 'secret@example.com', receivedAt: null,
          receivedAtText: null, preview: '敏感正文预览', hasAttachments: false,
        },
      },
    });

    const first = await service.move('1', '投后', true, 'move-dedupe-001');
    const second = await service.move('1', '投后', true, 'move-dedupe-001');

    expect(first.deduplicated).toBe(false);
    expect(second.deduplicated).toBe(true);
    expect(backend.moveMessage).toHaveBeenCalledOnce();
    const audit = await readFile(mutationStore.auditPath, 'utf8');
    expect(audit).toContain('m_');
    expect(audit).not.toContain('敏感测试主题');
    expect(audit).not.toContain('secret@example.com');
    expect(audit).not.toContain('敏感正文预览');
    expect((await stat(mutationStore.path)).mode & 0o777).toBe(0o600);
    expect((await stat(mutationStore.auditPath)).mode & 0o777).toBe(0o600);
  });

  it('rejects reusing a request ID with different mutation parameters', async () => {
    const { service, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: null, senderAddress: null, receivedAt: null,
          receivedAtText: null, preview: null, hasAttachments: false,
        },
      },
    });
    await service.move('1', '投后', true, 'move-conflict-001');

    await expect(service.move('1', '归档', true, 'move-conflict-001'))
      .rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
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

  it('creates a reply draft by default and hands it to the user', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: '张三', senderAddress: 'zhangsan@example.com',
          receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30',
          preview: '测试预览', hasAttachments: false,
        },
      },
    });

    await expect(service.reply('1', '收到，谢谢。')).resolves.toEqual({
      id: '1', draft: true, replyAll: false, sent: false,
      requiresManualSend: true, handedOff: true, verified: true,
    });
    expect(backend.replyMessage).toHaveBeenCalledWith(
      expect.objectContaining({ subject: '测试主题' }),
      '收到，谢谢。',
      true,
      false,
    );
  });

  it('automatically sends a reply-all only when draft is explicitly false', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: '张三', senderAddress: 'zhangsan@example.com',
          receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30',
          preview: '测试预览', hasAttachments: false,
        },
      },
    });
    vi.mocked(backend.replyMessage).mockResolvedValueOnce({
      matchCount: 1, status: 'sent', performed: true, verified: true,
      draft: false, replyAll: true, handedOff: false,
    });

    await expect(service.reply('1', '请大家查收。', false, true, 'reply-test-001')).resolves.toEqual({
      id: '1', draft: false, replyAll: true, sent: true,
      requiresManualSend: false, handedOff: false, verified: true,
      requestId: 'reply-test-001', deduplicated: false,
    });
    expect(backend.replyMessage).toHaveBeenCalledWith(expect.any(Object), '请大家查收。', false, true);
  });

  it('records an unverified automatic send as unknown and blocks retries', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: '张三', senderAddress: 'zhangsan@example.com',
          receivedAt: null, receivedAtText: null, preview: null, hasAttachments: false,
        },
      },
    });
    vi.mocked(backend.replyMessage).mockResolvedValueOnce({
      matchCount: 1, status: 'send_not_verified', performed: true, verified: false,
      draft: false, replyAll: false, handedOff: false,
    });

    await expect(service.reply('1', '自动发送', false, false, 'reply-unknown-001'))
      .rejects.toMatchObject({ code: 'OPERATION_UNKNOWN' });
    await expect(service.reply('1', '自动发送', false, false, 'reply-unknown-001'))
      .rejects.toMatchObject({ code: 'OPERATION_UNKNOWN' });
    expect(backend.replyMessage).toHaveBeenCalledOnce();
  });

  it('deduplicates automatic replies by request ID', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: '张三', senderAddress: 'zhangsan@example.com',
          receivedAt: null, receivedAtText: null, preview: null, hasAttachments: false,
        },
      },
    });
    vi.mocked(backend.replyMessage).mockResolvedValue({
      matchCount: 1, status: 'sent', performed: true, verified: true,
      draft: false, replyAll: false, handedOff: false,
    });

    await service.reply('1', '同一内容', false, false, 'reply-dedupe-001');
    const repeated = await service.reply('1', '同一内容', false, false, 'reply-dedupe-001');

    expect(repeated.deduplicated).toBe(true);
    expect(backend.replyMessage).toHaveBeenCalledOnce();
  });

  it('reports doctor checks without mutating the mailbox', async () => {
    const { service, backend } = await createService();
    vi.mocked(backend.eval).mockResolvedValueOnce({ searchInputs: 1, mailLists: 1, readingPanes: 0 });

    const result = await service.doctor();

    expect(result.ok).toBe(Number(process.versions.node.split('.')[0]) >= 24);
    expect(result.checks.map(check => check.name)).toEqual(['node', 'ego-lite', 'authentication', 'dom']);
    expect(backend.deleteMessage).not.toHaveBeenCalled();
    expect(backend.moveMessage).not.toHaveBeenCalled();
    expect(backend.replyMessage).not.toHaveBeenCalled();
  });

  it('rejects an empty reply before opening Outlook', async () => {
    const { service, backend } = await createService();

    await expect(service.reply('1', '   ')).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    expect(backend.status).not.toHaveBeenCalled();
    expect(backend.replyMessage).not.toHaveBeenCalled();
  });

  it('does not report success when reply content cannot be verified', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: null, senderAddress: null,
          receivedAt: null, receivedAtText: null, preview: null, hasAttachments: false,
        },
      },
    });
    vi.mocked(backend.replyMessage).mockResolvedValueOnce({
      matchCount: 1, status: 'content_not_verified', performed: true, verified: false,
      draft: true, replyAll: false, handedOff: true,
    });

    await expect(service.reply('1', '测试内容')).rejects.toMatchObject({ code: 'OPERATION_FAILED' });
  });

  it('exports one message and downloaded attachments as Obsidian Markdown', async () => {
    const { backend, parser, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: {
        '1': {
          subject: '测试主题', senderName: '张三', senderAddress: 'zhangsan@example.com',
          receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30',
          preview: '测试预览', hasAttachments: true,
        },
      },
    });
    const messageParser: MessageParser = {
      openAndExtract: vi.fn().mockResolvedValue({
        matchCount: 1,
        message: {
          subject: '测试主题', fromName: '张三', fromAddress: 'zhangsan@example.com',
          to: [{ name: '李四', address: 'lisi@example.com' }], cc: [],
          receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '2026/8/20 9:30',
          bodyText: '测试正文', attachments: [{ filename: '附件 报告.xlsx', sizeText: '12 KB' }],
        },
      }),
    };
    vi.mocked(backend.downloadAttachment).mockImplementation(async (_locator, _index, outputDirectory) => {
      const path = join(outputDirectory, '附件 报告.xlsx');
      await writeFile(path, 'attachment');
      return {
        matchCount: 1, status: 'performed', performed: true, verified: true,
        attachmentCount: 1, attachmentId: '1', filename: '附件 报告.xlsx', path, bytes: 10,
      };
    });
    const service = new OutlookService(backend, store, parser, messageParser);
    const outputDirectory = temporaryDirectories[temporaryDirectories.length - 1]!;

    const result = await service.exportObsidian('1', outputDirectory);
    const markdown = await readFile(result.markdownPath, 'utf8');

    expect(result.attachments).toHaveLength(1);
    expect(result.attachments[0]?.link).toMatch(/^attachments\/.+\/%E9%99%84%E4%BB%B6%20%E6%8A%A5%E5%91%8A\.xlsx$/);
    expect(markdown).toContain('# 测试主题');
    expect(markdown).toContain('## 正文\n\n测试正文');
    expect(markdown).toContain(`](${result.attachments[0]?.link})`);
    await expect(stat(result.attachments[0]!.path)).resolves.toMatchObject({ size: 10 });
  });

  it('creates a new mail draft by default without requiring a request ID', async () => {
    const { service, backend } = await createService();

    await expect(service.compose({
      to: [' Alice@example.com ', 'alice@example.com'], cc: [], bcc: [],
      subject: ' 测试新邮件 ', content: '正文', attachments: [], draft: true,
    })).resolves.toMatchObject({
      draft: true, sent: false, requiresManualSend: true, handedOff: true, verified: true,
    });
    expect(backend.composeMessage).toHaveBeenCalledWith(expect.objectContaining({
      to: ['Alice@example.com'], subject: '测试新邮件', draft: true,
    }));
  });

  it('requires and deduplicates request IDs for automatic new-mail sending', async () => {
    const { service, backend } = await createService();
    const options = {
      to: ['to@example.com'], cc: [], bcc: [], subject: '自动发送',
      content: '正文', attachments: [], draft: false,
    };
    vi.mocked(backend.composeMessage).mockResolvedValue({
      status: 'sent', performed: true, verified: true, draft: false,
      handedOff: false, attachmentCount: 0,
    });

    await expect(service.compose(options)).rejects.toMatchObject({ code: 'INVALID_ARGUMENT' });
    await service.compose(options, 'compose-send-001');
    await expect(service.compose(options, 'compose-send-001')).resolves.toMatchObject({
      sent: true, deduplicated: true,
    });
    expect(backend.composeMessage).toHaveBeenCalledOnce();
  });

  it('creates a forward draft while preserving original attachments', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: { '1': { ...rows[0]!, preview: rows[0]!.preview } },
    });

    await expect(service.forward('1', {
      to: ['to@example.com'], cc: [], bcc: [], content: '请查收', attachments: [], draft: true,
    })).resolves.toMatchObject({
      id: '1', draft: true, originalAttachmentsPreserved: true, handedOff: true,
    });
    expect(backend.forwardMessage).toHaveBeenCalledWith(
      expect.objectContaining({ subject: '已读邮件' }), expect.objectContaining({ content: '请查收' }),
    );
  });

  it('lists, reads, updates, sends and discards drafts through verified controls', async () => {
    const { service, backend, store } = await createService();
    const listed = await service.drafts({ limit: 1 });
    expect(listed.directory).toMatchObject({ path: '草稿' });
    expect(backend.selectSystemFolder).toHaveBeenCalledWith('草稿');

    await expect(service.readDraft('1')).resolves.toMatchObject({
      subject: '草稿主题', to: [{ address: 'to@example.com' }],
    });
    await expect(service.updateDraft('1', {
      subject: '更新主题', content: '更新正文', attachments: [],
    })).resolves.toMatchObject({ updated: true, sessionInvalidated: true });
    expect(backend.updateDraft).toHaveBeenCalledWith(
      expect.any(Object), expect.objectContaining({ subject: '更新主题', content: '更新正文' }),
    );

    await service.drafts({ limit: 1 });
    await expect(service.sendDraft('1', 'draft-send-001')).resolves.toMatchObject({ sent: true });
    await service.drafts({ limit: 1 });
    await expect(service.discardDraft('1', false, 'draft-discard-001'))
      .rejects.toMatchObject({ code: 'CONFIRMATION_REQUIRED' });
    await expect(service.discardDraft('1', true, 'draft-discard-001'))
      .resolves.toMatchObject({ discarded: true });
    await expect(store.read()).resolves.toBeNull();
  });

  it('changes read, flag and category state and returns a parsed conversation', async () => {
    const { service, backend, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: { '1': { ...rows[0]!, preview: rows[0]!.preview } },
    });

    await expect(service.markRead('1', false)).resolves.toMatchObject({ unread: false, changed: true });
    await expect(service.flag('1', true)).resolves.toMatchObject({ flagged: true, changed: true });
    await expect(service.categorize('1', ' 项目A ', true)).resolves.toMatchObject({
      category: '项目A', applied: true, changed: true,
    });
    await expect(service.conversation('1')).resolves.toMatchObject({
      complete: true, messages: [{ index: 1, bodyText: '会话正文' }],
    });
    expect(backend.setReadState).toHaveBeenCalledWith(expect.any(Object), false);
    expect(backend.setFlagState).toHaveBeenCalledWith(expect.any(Object), true);
    expect(backend.setCategoryState).toHaveBeenCalledWith(expect.any(Object), '项目A', true);
  });

  it('downloads every attachment and reports a SHA-256 digest', async () => {
    const { backend, parser, store } = await createService();
    await store.write({
      version: 1, updatedAt: '2026-08-20T01:00:00.000Z', source: 'inbox',
      messages: { '1': { ...rows[1]!, preview: rows[1]!.preview } },
    });
    const messageParser: MessageParser = {
      openAndExtract: vi.fn().mockResolvedValue({
        matchCount: 1,
        message: {
          subject: '未读邮件', fromName: null, fromAddress: 'unread@example.com', to: [], cc: [],
          receivedAt: null, receivedAtText: null, bodyText: '正文',
          attachments: [{ filename: 'report.xlsx', sizeText: '10 B' }],
        },
      }),
    };
    vi.mocked(backend.downloadAttachment).mockImplementation(async (_locator, _index, outputDirectory) => {
      const path = join(outputDirectory, 'report.xlsx');
      await writeFile(path, 'attachment');
      return {
        matchCount: 1, status: 'performed', performed: true, verified: true,
        filename: 'report.xlsx', path, bytes: 10,
      };
    });
    const service = new OutlookService(backend, store, parser, messageParser);
    const outputDirectory = temporaryDirectories[temporaryDirectories.length - 1]!;
    await writeFile(join(outputDirectory, 'report.xlsx'), 'existing');

    await expect(service.downloadAll('1', outputDirectory)).resolves.toMatchObject({
      attachments: [{
        id: '1', filename: 'report (2).xlsx', bytes: 10,
        sha256: createHash('sha256').update('attachment').digest('hex'),
      }],
    });
    await expect(readFile(join(outputDirectory, 'report.xlsx'), 'utf8')).resolves.toBe('existing');
  });
});
