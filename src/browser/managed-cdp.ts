import { open, readFile, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { spawn } from 'node:child_process';
import { AppError } from '../util/errors.js';
import type { PlaywrightConfig } from './playwright-config.js';

const ACTIVE_PORT_FILE = 'DevToolsActivePort';
const LAUNCH_LOCK_FILE = '.webmail-browser-launch.lock';

function delay(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

export function managedBrowserArgs(config: PlaywrightConfig): string[] {
  const args = [
    `--user-data-dir=${config.profileDir}`,
    '--remote-debugging-address=127.0.0.1',
    '--remote-debugging-port=0',
    '--no-first-run',
    '--no-default-browser-check',
  ];
  if (config.headless) args.push('--headless=new');
  args.push(config.outlookUrl);
  return args;
}

export async function readManagedCdpEndpoint(profileDir: string): Promise<string | null> {
  try {
    const content = await readFile(join(profileDir, ACTIVE_PORT_FILE), 'utf8');
    const port = Number(content.split(/\r?\n/, 1)[0]);
    return Number.isInteger(port) && port >= 1 && port <= 65_535
      ? `http://127.0.0.1:${port}`
      : null;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null;
    throw error;
  }
}

async function waitForEndpoint(profileDir: string, timeoutMs: number): Promise<string> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const endpoint = await readManagedCdpEndpoint(profileDir);
    if (endpoint) return endpoint;
    await delay(100);
  }
  throw new AppError(
    'PLAYWRIGHT_TIMEOUT',
    `浏览器已启动，但 ${timeoutMs}ms 内未建立 CDP 会话。Profile：${profileDir}。该目录也可能正被其他浏览器进程占用。`,
  );
}

export async function clearManagedCdpEndpoint(profileDir: string): Promise<void> {
  await unlink(join(profileDir, ACTIVE_PORT_FILE)).catch((error: NodeJS.ErrnoException) => {
    if (error.code !== 'ENOENT') throw error;
  });
}

export async function launchManagedBrowser(
  executablePath: string,
  config: PlaywrightConfig,
): Promise<string> {
  const lockPath = join(config.profileDir, LAUNCH_LOCK_FILE);
  let lock: Awaited<ReturnType<typeof open>> | null = null;
  try {
    try {
      lock = await open(lockPath, 'wx', 0o600);
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
      try {
        return await waitForEndpoint(config.profileDir, config.timeoutMs);
      } catch (waitError) {
        throw new AppError(
          'PROFILE_LOCKED',
          `Playwright 专用 Profile 可能被另一条 webmail 命令占用：${config.profileDir}。若没有浏览器正在启动，请删除 ${lockPath} 后重试。`,
          { cause: waitError },
        );
      }
    }

    await clearManagedCdpEndpoint(config.profileDir);
    const child = spawn(executablePath, managedBrowserArgs(config), {
      detached: true,
      stdio: 'ignore',
      windowsHide: false,
    });
    child.once('error', () => undefined);
    child.unref();
    return await waitForEndpoint(config.profileDir, config.timeoutMs);
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    if (code === 'EACCES' || code === 'EPERM') {
      throw new AppError(
        'PROFILE_ACCESS_DENIED',
        `无法使用 Playwright 专用 Profile：${config.profileDir}。请检查目录权限，或设置 WEBMAIL_PROFILE_DIR。`,
        { cause: error },
      );
    }
    throw error;
  } finally {
    await lock?.close().catch(() => undefined);
    if (lock) await unlink(lockPath).catch(() => undefined);
  }
}
