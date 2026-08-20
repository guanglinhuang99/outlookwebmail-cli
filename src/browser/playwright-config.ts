import { constants } from 'node:fs';
import { access, chmod, mkdir, realpath } from 'node:fs/promises';
import { homedir } from 'node:os';
import { delimiter, isAbsolute, join, resolve } from 'node:path';
import { AppError } from '../util/errors.js';

export type BrowserChoice = 'auto' | 'edge' | 'chrome' | 'chromium';
export type BackendChoice = 'auto' | 'playwright' | 'ego-lite';

export interface PlaywrightConfig {
  backend: BackendChoice;
  browser: BrowserChoice;
  executablePath: string | null;
  profileDir: string;
  headless: boolean;
  cdpEndpoint: string | null;
  timeoutMs: number;
  outlookUrl: string;
}

export interface BrowserExecutable {
  name: Exclude<BrowserChoice, 'auto'>;
  path: string;
}

export interface ConfigOptions {
  env?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  cwd?: string;
  home?: string;
}

const OUTLOOK_HOSTS = new Set(['partner.outlook.cn']);

function enumValue<T extends string>(value: string | undefined, allowed: readonly T[], fallback: T, name: string): T {
  const normalized = value?.normalize('NFKC').trim().toLowerCase() || fallback;
  if (allowed.includes(normalized as T)) return normalized as T;
  throw new AppError('INVALID_ARGUMENT', `${name} 必须是 ${allowed.join('、')} 之一。`);
}

function booleanValue(value: string | undefined, fallback: boolean, name: string): boolean {
  if (value === undefined || value.trim() === '') return fallback;
  const normalized = value.normalize('NFKC').trim().toLowerCase();
  if (normalized === 'true') return true;
  if (normalized === 'false') return false;
  throw new AppError('INVALID_ARGUMENT', `${name} 必须是 true 或 false。`);
}

function timeoutValue(value: string | undefined): number {
  if (value === undefined || value.trim() === '') return 30_000;
  const timeout = Number(value);
  if (!Number.isInteger(timeout) || timeout < 1_000 || timeout > 300_000) {
    throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_BROWSER_TIMEOUT_MS 必须是 1000 到 300000 之间的整数。');
  }
  return timeout;
}

function defaultProfileDir(platform: NodeJS.Platform, env: NodeJS.ProcessEnv, home: string): string {
  if (platform === 'win32') {
    return join(env.LOCALAPPDATA || join(home, 'AppData', 'Local'), 'webmail-cli', 'playwright-profile');
  }
  if (platform === 'darwin') return join(home, 'Library', 'Application Support', 'webmail-cli', 'playwright-profile');
  return join(env.XDG_DATA_HOME || join(home, '.local', 'share'), 'webmail-cli', 'playwright-profile');
}

function assertOutlookUrl(rawUrl: string): string {
  let url: URL;
  try {
    url = new URL(rawUrl);
  } catch (error) {
    throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_URL 不是有效 URL。', { cause: error });
  }
  if (url.protocol !== 'https:' || !OUTLOOK_HOSTS.has(url.hostname.toLowerCase())) {
    throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_URL 只允许使用 https://partner.outlook.cn。');
  }
  return url.href;
}

function inside(candidate: string, parent: string): boolean {
  const relative = resolve(candidate).slice(resolve(parent).length);
  return resolve(candidate) === resolve(parent) || (relative.startsWith('/') || relative.startsWith('\\'));
}

function validateProfilePath(profileDir: string, cwd: string, platform: NodeJS.Platform): void {
  if (!isAbsolute(profileDir)) throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_PROFILE_DIR 必须是绝对路径。');
  if (inside(profileDir, cwd)) throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_PROFILE_DIR 不能位于代码仓库中。');

  const normalized = resolve(profileDir).toLowerCase();
  const forbiddenSuffixes = platform === 'win32'
    ? ['google\\chrome\\user data', 'microsoft\\edge\\user data']
    : platform === 'darwin'
      ? ['library/application support/google/chrome', 'library/application support/microsoft edge']
      : ['.config/google-chrome', '.config/chromium', '.config/microsoft-edge'];
  if (forbiddenSuffixes.some(suffix => normalized.endsWith(suffix))) {
    throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_PROFILE_DIR 不能指向浏览器的默认用户数据目录。');
  }
}

export function loadPlaywrightConfig(options: ConfigOptions = {}): PlaywrightConfig {
  const env = options.env ?? process.env;
  const platform = options.platform ?? process.platform;
  const cwd = resolve(options.cwd ?? process.cwd());
  const home = options.home ?? homedir();
  const backend = enumValue(env.WEBMAIL_BACKEND, ['auto', 'playwright', 'ego-lite'] as const, 'auto', 'WEBMAIL_BACKEND');
  const browser = enumValue(env.WEBMAIL_BROWSER, ['auto', 'edge', 'chrome', 'chromium'] as const, 'auto', 'WEBMAIL_BROWSER');
  const executablePath = env.WEBMAIL_EXECUTABLE_PATH?.trim() || null;
  const cdpEndpoint = env.WEBMAIL_CDP_ENDPOINT?.trim() || null;
  if (cdpEndpoint && executablePath) {
    throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_CDP_ENDPOINT 与 WEBMAIL_EXECUTABLE_PATH 不能同时设置。');
  }
  const profileDir = resolve(env.WEBMAIL_PROFILE_DIR?.trim() || defaultProfileDir(platform, env, home));
  validateProfilePath(profileDir, cwd, platform);
  return {
    backend,
    browser,
    executablePath,
    profileDir,
    headless: booleanValue(env.WEBMAIL_HEADLESS, false, 'WEBMAIL_HEADLESS'),
    cdpEndpoint,
    timeoutMs: timeoutValue(env.WEBMAIL_BROWSER_TIMEOUT_MS),
    outlookUrl: assertOutlookUrl(env.WEBMAIL_URL?.trim() || 'https://partner.outlook.cn/mail/'),
  };
}

type AccessFn = (path: string, mode?: number) => Promise<void>;

function platformCandidates(platform: NodeJS.Platform, env: NodeJS.ProcessEnv): BrowserExecutable[] {
  if (platform === 'win32') {
    const roots = [env['PROGRAMFILES(X86)'], env.PROGRAMFILES, env.LOCALAPPDATA].filter((value): value is string => Boolean(value));
    return roots.flatMap(root => [
      { name: 'edge' as const, path: join(root, 'Microsoft', 'Edge', 'Application', 'msedge.exe') },
      { name: 'chrome' as const, path: join(root, 'Google', 'Chrome', 'Application', 'chrome.exe') },
    ]);
  }
  if (platform === 'darwin') {
    return [
      { name: 'edge', path: '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge' },
      { name: 'chrome', path: '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome' },
      { name: 'chromium', path: '/Applications/Chromium.app/Contents/MacOS/Chromium' },
    ];
  }
  return [];
}

function pathCommandNames(browser: BrowserChoice, platform: NodeJS.Platform): BrowserExecutable[] {
  const suffix = platform === 'win32' ? '.exe' : '';
  const all: BrowserExecutable[] = [
    { name: 'edge', path: `microsoft-edge${suffix}` },
    { name: 'edge', path: `microsoft-edge-stable${suffix}` },
    { name: 'chrome', path: `google-chrome${suffix}` },
    { name: 'chrome', path: `google-chrome-stable${suffix}` },
    { name: 'chromium', path: `chromium${suffix}` },
    { name: 'chromium', path: `chromium-browser${suffix}` },
  ];
  return browser === 'auto' ? all : all.filter(candidate => candidate.name === browser);
}

async function accessible(path: string, platform: NodeJS.Platform, accessFn: AccessFn): Promise<boolean> {
  try {
    await accessFn(path, platform === 'win32' ? constants.F_OK : constants.X_OK);
    return true;
  } catch {
    return false;
  }
}

export async function resolveBrowserExecutable(
  config: PlaywrightConfig,
  options: { env?: NodeJS.ProcessEnv; platform?: NodeJS.Platform; accessFn?: AccessFn } = {},
): Promise<BrowserExecutable> {
  const env = options.env ?? process.env;
  const platform = options.platform ?? process.platform;
  const accessFn = options.accessFn ?? access;
  if (config.executablePath) {
    if (!isAbsolute(config.executablePath) || !await accessible(config.executablePath, platform, accessFn)) {
      throw new AppError('BROWSER_NOT_FOUND', `WEBMAIL_EXECUTABLE_PATH 不存在或不可执行：${config.executablePath}`);
    }
    const filename = config.executablePath.toLowerCase();
    const inferred = /msedge/.test(filename) ? 'edge' : /chrome/.test(filename) ? 'chrome' : 'chromium';
    return { name: config.browser === 'auto' ? inferred : config.browser, path: config.executablePath };
  }

  const installed = platformCandidates(platform, env)
    .filter(candidate => config.browser === 'auto' || candidate.name === config.browser);
  for (const candidate of installed) {
    if (await accessible(candidate.path, platform, accessFn)) return candidate;
  }

  const pathDirs = (env.PATH || '').split(delimiter).filter(Boolean);
  for (const command of pathCommandNames(config.browser, platform)) {
    for (const directory of pathDirs) {
      const candidate = join(directory, command.path);
      if (await accessible(candidate, platform, accessFn)) return { ...command, path: candidate };
    }
  }
  throw new AppError('BROWSER_NOT_FOUND', '未找到 Edge、Chrome 或 Chromium；请安装浏览器或设置 WEBMAIL_EXECUTABLE_PATH。');
}

export async function prepareProfileDirectory(config: PlaywrightConfig): Promise<void> {
  if (config.cdpEndpoint) return;
  await mkdir(config.profileDir, { recursive: true, mode: 0o700 });
  const actual = await realpath(config.profileDir);
  if (inside(actual, process.cwd())) {
    throw new AppError('INVALID_ARGUMENT', 'WEBMAIL_PROFILE_DIR 解析后位于代码仓库中。');
  }
  if (process.platform !== 'win32') await chmod(actual, 0o700);
}

export function redactCdpEndpoint(endpoint: string): string {
  try {
    const url = new URL(endpoint);
    return `${url.protocol}//${url.host}${url.pathname}`;
  } catch {
    return '<configured endpoint>';
  }
}
