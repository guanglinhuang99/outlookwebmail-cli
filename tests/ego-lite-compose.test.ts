import { describe, expect, it, vi } from 'vitest';
import ts from 'typescript';
import { EgoLiteBackend } from '../src/browser/ego-lite.js';
import type { EgoRunner } from '../src/browser/ego-runner.js';
import type { ComposeActionResult, MessageLocator, RawMessage } from '../src/types/mail.js';

const locator: MessageLocator = {
  subject: '原邮件', senderName: '张三', senderAddress: 'zhangsan@example.com',
  receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30',
  preview: '原正文', hasAttachments: true,
};

const message: RawMessage = {
  subject: '原邮件', fromName: '张三', fromAddress: 'zhangsan@example.com',
  to: [], cc: [], receivedAt: locator.receivedAt, receivedAtText: locator.receivedAtText,
  bodyText: '原正文', attachments: [{ filename: 'original.pdf', sizeText: '1 MB' }],
};

function expectValidBrowserScript(script: string): void {
  const diagnostics = ts.transpileModule(script, {
    reportDiagnostics: true,
    compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2023 },
  }).diagnostics ?? [];
  expect(diagnostics.map(diagnostic => diagnostic.messageText)).toEqual([]);
}

describe('EgoLiteBackend compose and forward', () => {
  it('builds a verified draft script with recipients, BCC and local attachments', async () => {
    const result: ComposeActionResult = {
      status: 'draft_ready', performed: true, verified: true, draft: true,
      handedOff: true, attachmentCount: 1,
    };
    const run = vi.fn().mockResolvedValue({ stdout: '', stderr: '', value: result });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await expect(backend.composeMessage({
      to: ['to@example.com'], cc: ['cc@example.com'], bcc: ['bcc@example.com'],
      subject: '新邮件', content: '第一行\n第二行', attachments: ['/tmp/report.xlsx'], draft: true,
    })).resolves.toEqual(result);

    const script = run.mock.calls[0]?.[0] as string;
    expectValidBrowserScript(script);
    expect(script).toContain('密件抄送');
    expect(script).toContain("Input.insertText");
    expect(script).toContain("uploadFile('input[type=\"file\"]");
    expect(script).toContain('/tmp/report.xlsx');
    expect(script).toContain('!verified.subject');
    expect(script).toContain('保存草稿|save draft');
    expect(script).toContain('handOffTaskSpace(task.id)');
  });

  it('opens one source message and verifies editor closure before reporting a forward send', async () => {
    const result: ComposeActionResult = {
      matchCount: 1, status: 'sent', performed: true, verified: true,
      draft: false, handedOff: false, attachmentCount: 0,
    };
    const run = vi.fn()
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: { matchCount: 1, message } })
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: result });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await expect(backend.forwardMessage(locator, {
      to: ['to@example.com'], cc: [], bcc: [], content: '转发说明', attachments: [], draft: false,
    })).resolves.toEqual(result);

    const script = run.mock.calls[1]?.[0] as string;
    expectValidBrowserScript(script);
    expect(script).toContain('^(转发|forward)$');
    expect(script).toContain('original.pdf');
    expect(script).toContain("closed ? 'sent' : 'send_not_verified'");
    expect(script).toContain('if (closed) break');
  });

  it('verifies every explicitly updated draft field after saving', async () => {
    const result: ComposeActionResult = {
      matchCount: 1, status: 'draft_ready', performed: true, verified: true,
      draft: true, handedOff: false, attachmentCount: 1,
    };
    const run = vi.fn()
      .mockResolvedValueOnce({
        stdout: '', stderr: '',
        value: {
          matchCount: 1,
          draft: { to: [], cc: [], bcc: [], subject: '旧主题', bodyText: '旧正文', attachments: [] },
        },
      })
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: result })
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: null })
      .mockResolvedValueOnce({
        stdout: '', stderr: '',
        value: {
          matchCount: 1, closed: true,
          draft: {
            to: [{ name: null, address: 'new@example.com' }], cc: [], bcc: [],
            subject: '新主题', bodyText: '新正文', attachments: [{ filename: 'new.pdf', sizeText: null }],
          },
        },
      });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await expect(backend.updateDraft(locator, {
      to: ['new@example.com'], cc: [], subject: '新主题', content: '新正文',
      attachments: ['/tmp/new.pdf'],
    })).resolves.toEqual(result);

    const openScript = run.mock.calls[0]?.[0] as string;
    expectValidBrowserScript(openScript);
    expect(openScript).toContain('await fillInput');
    expect(openScript).toContain('const seenMatches = new Map()');
    expect(openScript).toContain('pageIndex < 20');
    const script = run.mock.calls[1]?.[0] as string;
    expectValidBrowserScript(script);
    expect(script).toContain('new@example.com');
    expect(script).toContain('新主题');
    expect(script).toContain('new.pdf');
    expect(script).toContain('document.activeElement===field');
    expect(script).toContain('verification.recipients');
    expect(script).toContain("verified ? 'draft_ready' : 'content_not_verified'");
    expect(run.mock.calls[2]?.[0]).toContain("Page.reload");
    expect(run).toHaveBeenCalledTimes(4);
  });
});
