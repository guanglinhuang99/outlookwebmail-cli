import { describe, expect, it, vi } from 'vitest';
import type { BrowserBackend } from '../src/browser/backend.js';
import { FolderParser } from '../src/outlook/folder-parser.js';

describe('FolderParser', () => {
  it('returns validated folder paths', async () => {
    const backend = {
      listInboxFolders: vi.fn().mockResolvedValue({
        accountCount: 1,
        inboxCount: 1,
        complete: true,
        folders: [
          { name: '投后', path: '收件箱/投后', level: 3, expanded: null },
          { name: '股票入池', path: '收件箱/投资审批/股票入池', level: 4, expanded: null },
        ],
      }),
    } as unknown as BrowserBackend;
    await expect(new FolderParser(backend).extract()).resolves.toHaveLength(2);
  });

  it('rejects an incomplete Inbox subtree', async () => {
    const backend = {
      listInboxFolders: vi.fn().mockResolvedValue({
        accountCount: 1, inboxCount: 1, complete: false, folders: [],
      }),
    } as unknown as BrowserBackend;
    await expect(new FolderParser(backend).extract()).rejects.toMatchObject({ code: 'OPERATION_FAILED' });
  });
});
