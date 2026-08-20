import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { managedBrowserArgs, readManagedCdpEndpoint } from '../src/browser/managed-cdp.js';
import type { PlaywrightConfig } from '../src/browser/playwright-config.js';

const directories: string[] = [];

afterEach(async () => {
  await Promise.all(directories.splice(0).map(path => rm(path, { recursive: true, force: true })));
});

function config(profileDir: string): PlaywrightConfig {
  return {
    backend: 'playwright', browser: 'edge', executablePath: null, profileDir,
    headless: false, cdpEndpoint: null, timeoutMs: 30_000,
    outlookUrl: 'https://partner.outlook.cn/mail/',
  };
}

describe('managed Playwright CDP browser', () => {
  it('launches a dedicated browser profile with an ephemeral local CDP port', () => {
    const args = managedBrowserArgs(config('/tmp/webmail-profile'));
    expect(args).toContain('--user-data-dir=/tmp/webmail-profile');
    expect(args).toContain('--remote-debugging-address=127.0.0.1');
    expect(args).toContain('--remote-debugging-port=0');
    expect(args.at(-1)).toBe('https://partner.outlook.cn/mail/');
  });

  it('discovers a reusable endpoint from the dedicated profile', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'webmail-cdp-test-'));
    directories.push(directory);
    await mkdir(directory, { recursive: true });
    await writeFile(join(directory, 'DevToolsActivePort'), '43123\n/devtools/browser/test\n');
    await expect(readManagedCdpEndpoint(directory)).resolves.toBe('http://127.0.0.1:43123');
  });

  it('rejects malformed or absent endpoint files', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'webmail-cdp-test-'));
    directories.push(directory);
    await expect(readManagedCdpEndpoint(directory)).resolves.toBeNull();
    await writeFile(join(directory, 'DevToolsActivePort'), 'not-a-port\n');
    await expect(readManagedCdpEndpoint(directory)).resolves.toBeNull();
  });
});
