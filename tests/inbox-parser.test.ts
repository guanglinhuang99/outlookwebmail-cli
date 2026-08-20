import { describe, expect, it, vi } from 'vitest';
import type { BrowserBackend } from '../src/browser/backend.js';
import { EgoInboxParser } from '../src/outlook/inbox-parser.js';

function backendWithEval(value: unknown): BrowserBackend {
  return {
    eval: vi.fn().mockResolvedValue(value),
    wheel: vi.fn().mockResolvedValue(undefined),
    wheelAndEval: vi.fn().mockResolvedValue(value),
  } as unknown as BrowserBackend;
}

const parsedRow = {
  stableHint: null,
  senderName: '张三',
  senderAddress: 'zhangsan@example.com',
  subject: '测试主题',
  receivedAt: '2026-08-20T09:35:00+08:00',
  receivedAtText: '9:35',
  preview: '测试预览',
  unread: true,
  hasAttachments: false,
};

describe('EgoInboxParser', () => {
  it('validates and returns extracted rows', async () => {
    const parser = new EgoInboxParser(backendWithEval({ listCount: 1, rowCount: 1, rows: [parsedRow] }));
    await expect(parser.extract()).resolves.toEqual([parsedRow]);
  });

  it('reports UI_CHANGED when the list root is missing', async () => {
    const parser = new EgoInboxParser(backendWithEval({ listCount: 0, rowCount: 0, rows: [] }));
    await expect(parser.extract()).rejects.toMatchObject({ code: 'UI_CHANGED' });
  });

  it('reports UI_CHANGED when rows exist but none parse', async () => {
    const parser = new EgoInboxParser(backendWithEval({ listCount: 1, rowCount: 4, rows: [] }));
    await expect(parser.extract()).rejects.toMatchObject({ code: 'UI_CHANGED' });
  });

  it('keeps wheel, wait and extraction in one virtual-list transaction', async () => {
    const value = { listCount: 1, rowCount: 0, rows: [] };
    const backend = backendWithEval(value);
    const parser = new EgoInboxParser(backend);
    await expect(parser.resetAndExtract()).resolves.toEqual([]);
    await expect(parser.scrollAndExtract()).resolves.toEqual([]);
    expect(backend.wheelAndEval).toHaveBeenNthCalledWith(1, expect.stringContaining('listbox'), -900, 1, 400, expect.any(String));
    expect(backend.wheelAndEval).toHaveBeenLastCalledWith(expect.stringContaining('listbox'), 900, 1, 500, expect.any(String));
  });
});
