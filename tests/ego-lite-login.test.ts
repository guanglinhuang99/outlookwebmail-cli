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
    expect(script).toContain("useOrCreateTaskSpace(\"webmail-cli\")");
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
    expect(script).toContain('await ensureRealTab()');
    expect(script).toContain('if (!tab) tab = await openOrReuseTab');
    expect(script).toContain('await handOffTaskSpace(task.id)');
  });
});
