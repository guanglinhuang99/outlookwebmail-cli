import { describe, expect, it, vi } from 'vitest';
import type { BrowserBackend } from '../src/browser/backend.js';
import { EgoMessageParser } from '../src/outlook/message-parser.js';

const locator = {
  subject: '测试主题',
  senderName: '张三',
  senderAddress: 'zhangsan@example.com',
  receivedAt: '2026-08-20T09:30:00+08:00',
  receivedAtText: '9:30',
  preview: '测试预览',
  hasAttachments: true,
};

const message = {
  subject: '测试主题',
  fromName: '张三',
  fromAddress: 'zhangsan@example.com',
  to: [{ name: '李四', address: null }],
  cc: [],
  receivedAt: '2026-08-20T09:30:00+08:00',
  receivedAtText: '周四 2026/8/20 9:30',
  bodyText: '测试正文',
  attachments: [{ filename: 'report.xlsx', sizeText: '12 KB' }],
};

describe('EgoMessageParser', () => {
  it('validates a uniquely opened message', async () => {
    const backend = {
      openAndExtractMessage: vi.fn().mockResolvedValue({ matchCount: 1, message }),
    } as unknown as BrowserBackend;
    const parser = new EgoMessageParser(backend);
    await expect(parser.openAndExtract(locator)).resolves.toEqual({ matchCount: 1, message });
  });

  it('accepts a not-found result without fabricated message data', async () => {
    const backend = {
      openAndExtractMessage: vi.fn().mockResolvedValue({ matchCount: 0, message: null }),
    } as unknown as BrowserBackend;
    const parser = new EgoMessageParser(backend);
    await expect(parser.openAndExtract(locator)).resolves.toEqual({ matchCount: 0, message: null });
  });
});
