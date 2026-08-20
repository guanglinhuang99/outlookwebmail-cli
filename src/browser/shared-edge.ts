import { readFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { AppError } from '../util/errors.js';

export interface SharedEdgeDiscoveryOptions {
  env?: NodeJS.ProcessEnv;
  platform?: NodeJS.Platform;
  home?: string;
  readFileFn?: typeof readFile;
}

export function defaultEdgeUserDataDir(options: SharedEdgeDiscoveryOptions = {}): string {
  const env = options.env ?? process.env;
  const platform = options.platform ?? process.platform;
  const home = options.home ?? homedir();
  if (platform === 'win32') {
    return join(env.LOCALAPPDATA || join(home, 'AppData', 'Local'), 'Microsoft', 'Edge', 'User Data');
  }
  if (platform === 'darwin') return join(home, 'Library', 'Application Support', 'Microsoft Edge');
  return join(home, '.config', 'microsoft-edge');
}

export async function discoverSharedEdgeEndpoint(options: SharedEdgeDiscoveryOptions = {}): Promise<string> {
  const userDataDir = defaultEdgeUserDataDir(options);
  const activePortPath = join(userDataDir, 'DevToolsActivePort');
  let content: string;
  try {
    content = await (options.readFileFn ?? readFile)(activePortPath, 'utf8');
  } catch (error) {
    const code = (error as NodeJS.ErrnoException).code;
    const detail = code === 'ENOENT' ? '未找到 DevToolsActivePort' : `无法读取 DevToolsActivePort（${code ?? 'unknown'}）`;
    throw new AppError(
      'SHARED_EDGE_NOT_AVAILABLE',
      `${detail}：${activePortPath}。请在日常 Edge 的 edge://inspect 中启用“Allow remote debugging for this browser instance”。`,
      { cause: error },
    );
  }
  const lines = content.split(/\r?\n/).map(line => line.trim());
  const port = Number(lines[0]);
  if (!Number.isInteger(port) || port < 1 || port > 65_535) {
    throw new AppError('SHARED_EDGE_NOT_AVAILABLE', `DevToolsActivePort 中的端口无效：${activePortPath}。请重新启用 Edge 远程调试。`);
  }
  const webSocketPath = lines[1] ?? '';
  if (/^\/devtools\/browser\/[A-Za-z0-9._-]+$/.test(webSocketPath)) {
    return `ws://127.0.0.1:${port}${webSocketPath}`;
  }
  return `http://127.0.0.1:${port}`;
}
