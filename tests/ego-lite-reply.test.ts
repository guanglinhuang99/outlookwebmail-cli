import { describe, expect, it, vi } from 'vitest';
import ts from 'typescript';
import { EgoLiteBackend } from '../src/browser/ego-lite.js';
import type { EgoRunner } from '../src/browser/ego-runner.js';
import type { MessageLocator, RawMessage, ReplyActionResult } from '../src/types/mail.js';

const locator: MessageLocator = {
  subject: '测试主题',
  senderName: '张三',
  senderAddress: 'zhangsan@example.com',
  receivedAt: '2026-08-20T09:30:00+08:00',
  receivedAtText: '9:30',
  preview: '测试预览',
  hasAttachments: false,
};

const message: RawMessage = {
  subject: '测试主题',
  fromName: '张三',
  fromAddress: 'zhangsan@example.com',
  to: [],
  cc: [],
  receivedAt: '2026-08-20T09:30:00+08:00',
  receivedAtText: '9:30',
  bodyText: '测试正文',
  attachments: [],
};

function expectValidBrowserScript(script: string): void {
  const diagnostics = ts.transpileModule(script, {
    reportDiagnostics: true,
    compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2023 },
  }).diagnostics ?? [];
  expect(diagnostics.map(diagnostic => diagnostic.messageText)).toEqual([]);
}

describe('EgoLiteBackend reply', () => {
  it('builds a reply-all draft script that inserts content and hands off control', async () => {
    const replyResult: ReplyActionResult = {
      matchCount: 1,
      status: 'draft_ready',
      performed: true,
      verified: true,
      draft: true,
      replyAll: true,
      handedOff: true,
    };
    const run = vi.fn()
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: { matchCount: 1, message } })
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: replyResult });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await expect(backend.replyMessage(locator, '第一行\n第二行', true, true)).resolves.toEqual(replyResult);

    const script = run.mock.calls[1]?.[0] as string;
    expectValidBrowserScript(script);
    expect(script).toContain('const replyAll = true;');
    expect(script).toContain('全部答复|reply all');
    expect(script).toContain("Input.insertText");
    expect(script).toContain('第一行\\n第二行');
    expect(script).toContain('handOffTaskSpace(task.id)');
    expect(script).toContain("status: 'draft_ready'");
  });

  it('requires the editor to close before reporting an automatic send', async () => {
    const replyResult: ReplyActionResult = {
      matchCount: 1,
      status: 'sent',
      performed: true,
      verified: true,
      draft: false,
      replyAll: false,
      handedOff: false,
    };
    const run = vi.fn()
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: { matchCount: 1, message } })
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: replyResult });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await expect(backend.replyMessage(locator, '收到', false, false)).resolves.toEqual(replyResult);

    const script = run.mock.calls[1]?.[0] as string;
    expectValidBrowserScript(script);
    expect(script).toContain('const replyAll = false;');
    expect(script).toContain('答复|reply');
    expect(script).toContain("status: 'sent'");
    expect(script).toContain('if (sent) break');
    expect(script).toContain("status: 'send_not_verified'");
  });
});
