import { describe, expect, it, vi } from 'vitest';
import { AppError } from '../src/util/errors.js';
import {
  loadPlaywrightConfig,
  prepareProfileDirectory,
  resolveBrowserExecutable,
} from '../src/browser/playwright-config.js';

describe('Playwright configuration', () => {
  it('defaults to automatic Playwright-first backend selection', () => {
    const config = loadPlaywrightConfig({ env: {}, platform: 'win32', cwd: '/work/repo', home: '/users/test' });
    expect(config).toMatchObject({
      mode: 'default',
      backend: 'auto',
      browser: 'auto',
      headless: false,
      shareEdge: false,
      timeoutMs: 30_000,
      outlookUrl: 'https://partner.outlook.cn/mail/',
    });
    expect(config.profileDir).toContain('webmail-cli');
  });

  it('forces the complete Ego Lite path when WEBMAIL_MODE=egolite', () => {
    const config = loadPlaywrightConfig({
      env: {
        WEBMAIL_MODE: 'egolite',
        WEBMAIL_BACKEND: 'unsupported',
        WEBMAIL_SHARE_EDGE: 'true',
        WEBMAIL_BROWSER: 'chrome',
        WEBMAIL_CDP_ENDPOINT: 'http://127.0.0.1:57652',
        WEBMAIL_HEADLESS: 'true',
      },
      platform: 'darwin', cwd: '/work/repo', home: '/users/test',
    });
    expect(config).toMatchObject({
      mode: 'egolite', backend: 'ego-lite', browser: 'auto', shareEdge: false,
      cdpEndpoint: null, executablePath: null, headless: false,
    });
  });

  it('rejects unsupported WEBMAIL_MODE values', () => {
    expect(() => loadPlaywrightConfig({
      env: { WEBMAIL_MODE: 'edge' }, platform: 'darwin', cwd: '/work/repo', home: '/users/test',
    })).toThrowError(/WEBMAIL_MODE 必须是 default、egolite/);
  });

  it('enables explicit shared Edge mode', () => {
    const config = loadPlaywrightConfig({
      env: { WEBMAIL_SHARE_EDGE: 'true' }, platform: 'darwin', cwd: '/work/repo', home: '/users/test',
    });
    expect(config).toMatchObject({ shareEdge: true, browser: 'auto', backend: 'auto', cdpEndpoint: null });
  });

  it('rejects conflicting shared Edge settings', () => {
    const common = { platform: 'darwin' as const, cwd: '/work/repo', home: '/users/test' };
    expect(() => loadPlaywrightConfig({ ...common, env: {
      WEBMAIL_SHARE_EDGE: 'true', WEBMAIL_CDP_ENDPOINT: 'http://127.0.0.1:57652',
    } })).toThrowError(/不能同时设置 WEBMAIL_CDP_ENDPOINT/);
    expect(() => loadPlaywrightConfig({ ...common, env: {
      WEBMAIL_SHARE_EDGE: 'true', WEBMAIL_BROWSER: 'chrome',
    } })).toThrowError(/只支持 WEBMAIL_BROWSER=auto 或 edge/);
    expect(() => loadPlaywrightConfig({ ...common, env: {
      WEBMAIL_SHARE_EDGE: 'true', WEBMAIL_BACKEND: 'ego-lite',
    } })).toThrowError(/不能与 WEBMAIL_BACKEND=ego-lite/);
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

  it('accepts an unrelated absolute profile with a coincidental path length', () => {
    expect(loadPlaywrightConfig({
      env: { WEBMAIL_PROFILE_DIR: '/abcde/profile' },
      platform: 'linux', cwd: '/repo1', home: '/users/test',
    }).profileDir).toBe('/abcde/profile');
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

  it('reports a clear profile path when the directory is not writable', async () => {
    const config = loadPlaywrightConfig({
      env: { WEBMAIL_PROFILE_DIR: '/profiles/webmail' }, platform: 'linux', cwd: '/repo', home: '/users/test',
    });
    const denied = Object.assign(new Error('denied'), { code: 'EACCES' });
    await expect(prepareProfileDirectory(config, {
      mkdirFn: vi.fn().mockRejectedValue(denied),
      cwd: '/repo',
      platform: 'linux',
    })).rejects.toMatchObject({
      code: 'PROFILE_ACCESS_DENIED',
      message: expect.stringContaining('/profiles/webmail'),
    });
  });
});
