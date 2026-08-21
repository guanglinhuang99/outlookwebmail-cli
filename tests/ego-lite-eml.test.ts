import { describe, expect, it, vi } from 'vitest';
import ts from 'typescript';
import { EgoLiteBackend } from '../src/browser/ego-lite.js';
import type { EgoRunner } from '../src/browser/ego-runner.js';
import type { MessageLocator, RawMessage } from '../src/types/mail.js';

const locator: MessageLocator = {
  subject: '测试主题', senderName: '张三', senderAddress: 'zhangsan@example.com',
  receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30',
  preview: '测试预览', hasAttachments: true,
};

const message: RawMessage = {
  subject: '测试主题', fromName: '张三', fromAddress: 'zhangsan@example.com', to: [], cc: [],
  receivedAt: '2026-08-20T09:30:00+08:00', receivedAtText: '9:30', bodyText: '测试正文',
  attachments: [{ filename: '附件.txt', sizeText: '10 B' }],
};

function expectValidBrowserScript(script: string): void {
  const diagnostics = ts.transpileModule(script, {
    reportDiagnostics: true,
    compilerOptions: { module: ts.ModuleKind.ESNext, target: ts.ScriptTarget.ES2023 },
  }).diagnostics ?? [];
  expect(diagnostics.map(diagnostic => diagnostic.messageText)).toEqual([]);
}

describe('EgoLiteBackend native EML download', () => {
  it('uses Outlook File, Download, Download as EML menus and verifies the downloaded file', async () => {
    const result = {
      matchCount: 1, status: 'performed' as const, performed: true, verified: true,
      attachmentCount: 1, filename: 'message.eml', path: '/tmp/message.eml', bytes: 100,
    };
    const run = vi.fn()
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: { matchCount: 1, message } })
      .mockResolvedValueOnce({ stdout: '', stderr: '', value: result });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await expect(backend.downloadMessageAsEml(locator, '/tmp/eml')).resolves.toEqual(result);

    const script = run.mock.calls[1]?.[0] as string;
    expectValidBrowserScript(script);
    expect(script).toContain('Page.setDownloadBehavior');
    expect(script).toContain('/^(文件|file)(?:\\\\s*(菜单|menu))?$/i');
    expect(script).toContain('/^(下载|download)$/i');
    expect(script).toContain('下载为\\\\s*\\\\.?eml(?:\\\\s*文件)?|download as\\\\s*\\\\.?eml(?:\\\\s*file)?');
    expect(script).toContain("!/\\.eml$/i.test(name)");
    expect(script).toContain('info.isFile() && info.size > 0');
    expect(script).toContain("status: 'download_failed'");
  });
});
