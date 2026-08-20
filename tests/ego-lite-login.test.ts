import { describe, expect, it, vi } from 'vitest';
import { EgoLiteBackend } from '../src/browser/ego-lite.js';
import type { EgoRunner } from '../src/browser/ego-runner.js';

describe('EgoLiteBackend login interaction', () => {
  it('opens Outlook automatically while checking status', async () => {
    const run = vi.fn().mockResolvedValueOnce({
      stdout: '',
      stderr: '',
      value: {
        connected: true,
        taskSpaceId: 7,
        url: 'https://partner.outlook.cn/mail/',
        title: 'Outlook',
        page: { url: 'https://partner.outlook.cn/mail/', title: 'Outlook' },
        snapshot: '收件箱',
      },
    });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await backend.status();

    const script = run.mock.calls[0]?.[0] as string;
    expect(script).toContain("useOrCreateTaskSpace(\"webmail-cli-production\")");
    expect(script).toContain("openOrReuseTab(\"https://partner.outlook.cn/mail/\"");
  });

  it('reuses the current tab and hands it to the user for login', async () => {
    const run = vi.fn().mockResolvedValueOnce({
      stdout: '',
      stderr: '',
      value: { taskSpaceId: 7, url: 'https://partner.outlook.cn/mail/', handedOff: true },
    });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);

    await expect(backend.handoffForLogin()).resolves.toMatchObject({ handedOff: true });

    const script = run.mock.calls[0]?.[0] as string;
    expect(script).not.toContain('await ensureRealTab()');
    expect(script).toContain('const tab = await openOrReuseTab');
    expect(script).toContain('await handOffTaskSpace(task.id)');
  });

  it('uses a separate configurable task space', async () => {
    const previous = process.env.WEBMAIL_EGO_TASK_SPACE;
    process.env.WEBMAIL_EGO_TASK_SPACE = 'webmail-cli-test-fixture';
    try {
      const run = vi.fn().mockResolvedValueOnce({ stdout: '', stderr: '', value: {} });
      const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);
      await backend.status();
      expect(run.mock.calls[0]?.[0]).toContain('useOrCreateTaskSpace("webmail-cli-test-fixture")');
    } finally {
      if (previous === undefined) delete process.env.WEBMAIL_EGO_TASK_SPACE;
      else process.env.WEBMAIL_EGO_TASK_SPACE = previous;
    }
  });

  it('checks the current stable Outlook row before falling back to a subject search', async () => {
    const run = vi.fn().mockResolvedValueOnce({
      stdout: '', stderr: '', value: { matchCount: 0, message: null },
    });
    const backend = new EgoLiteBackend({ run } as unknown as EgoRunner);
    await backend.openAndExtractMessage({
      stableHint: 'item-1', subject: '同主题', senderName: '张三', senderAddress: 'a@example.com',
      receivedAt: null, receivedAtText: '9:30', preview: '唯一预览', hasAttachments: false, unread: false,
    });
    const script = run.mock.calls[0]?.[0] as string;
    expect(script).toContain('const currentStableMatch');
    expect(script).toContain('if (!currentStableMatch)');
  });
});
