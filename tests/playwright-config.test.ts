import { describe, expect, it, vi } from 'vitest';
import { AppError } from '../src/util/errors.js';
import {
  loadPlaywrightConfig,
  resolveBrowserExecutable,
} from '../src/browser/playwright-config.js';

describe('Playwright configuration', () => {
  it('defaults to automatic Playwright-first backend selection', () => {
    const config = loadPlaywrightConfig({ env: {}, platform: 'win32', cwd: '/work/repo', home: '/users/test' });
    expect(config).toMatchObject({
      backend: 'auto',
      browser: 'auto',
      headless: false,
      timeoutMs: 30_000,
      outlookUrl: 'https://partner.outlook.cn/mail/',
    });
    expect(config.profileDir).toContain('webmail-cli');
  });

  it('rejects conflicting CDP and executable configuration', () => {
    expect(() => loadPlaywrightConfig({
      env: { WEBMAIL_CDP_ENDPOINT: 'http://127.0.0.1:9222', WEBMAIL_EXECUTABLE_PATH: '/browser' },
      cwd: '/work/repo',
      home: '/users/test',
    })).toThrow(AppError);
  });

  it('rejects a profile inside the repository', () => {
    expect(() => loadPlaywrightConfig({
      env: { WEBMAIL_PROFILE_DIR: '/work/repo/playwright-profile' },
      cwd: '/work/repo',
      home: '/users/test',
    })).toThrowError(/不能位于代码仓库/);
  });

  it('discovers Edge from Windows installation roots', async () => {
    const config = loadPlaywrightConfig({
      env: { PROGRAMFILES: 'C:\\Program Files', WEBMAIL_BROWSER: 'edge' },
      platform: 'win32',
      cwd: '/work/repo',
      home: '/users/test',
    });
    const accessFn = vi.fn(async (path: string) => {
      if (!path.endsWith('msedge.exe')) throw new Error('missing');
    });
    const result = await resolveBrowserExecutable(config, {
      env: { PROGRAMFILES: 'C:\\Program Files' },
      platform: 'win32',
      accessFn,
    });
    expect(result.name).toBe('edge');
    expect(result.path).toMatch(/msedge\.exe$/);
  });

  it('discovers Chromium commands from PATH on Linux', async () => {
    const config = loadPlaywrightConfig({
      env: { PATH: '/opt/bin', WEBMAIL_BROWSER: 'chromium' },
      platform: 'linux',
      cwd: '/work/repo',
      home: '/users/test',
    });
    const result = await resolveBrowserExecutable(config, {
      env: { PATH: '/opt/bin' },
      platform: 'linux',
      accessFn: async path => {
        if (path !== '/opt/bin/chromium') throw new Error('missing');
      },
    });
    expect(result).toEqual({ name: 'chromium', path: '/opt/bin/chromium' });
  });

  it('returns BROWSER_NOT_FOUND when discovery fails', async () => {
    const config = loadPlaywrightConfig({ env: {}, platform: 'linux', cwd: '/work/repo', home: '/users/test' });
    await expect(resolveBrowserExecutable(config, {
      env: {},
      platform: 'linux',
      accessFn: async () => { throw new Error('missing'); },
    })).rejects.toMatchObject({ code: 'BROWSER_NOT_FOUND' });
  });
});
