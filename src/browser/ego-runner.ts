import { spawn } from 'node:child_process';
import { AppError } from '../util/errors.js';

const RESULT_MARKER = '__webmail_result__';
const MAX_OUTPUT_BYTES = 20 * 1024 * 1024;

interface MarkedResult<T> {
  [RESULT_MARKER]: true;
  result: T;
}

export interface EgoRunResult<T> {
  stdout: string;
  stderr: string;
  value: T;
}

export interface BrowserScriptRunner {
  run<T>(body: string, timeoutMs?: number): Promise<EgoRunResult<T>>;
  close?(): Promise<void>;
}

export interface EgoRunnerOptions {
  command?: string;
  args?: string[];
  maxOutputBytes?: number;
}

export function parseMarkedResult<T>(stdout: string): T {
  const lines = stdout.split(/\r?\n/);

  for (let index = lines.length - 1; index >= 0; index -= 1) {
    const line = lines[index]?.trim();
    if (!line) continue;

    try {
      const candidate = JSON.parse(line) as Partial<MarkedResult<T>>;
      if (candidate[RESULT_MARKER] === true && Object.hasOwn(candidate, 'result')) {
        return candidate.result as T;
      }
    } catch {
      // ego-browser may emit ordinary log lines; only marked JSON is a result.
    }
  }

  throw new AppError(
    'EGO_BROWSER_ERROR',
    `ego-browser 未返回可识别的结果（stdout ${Buffer.byteLength(stdout, 'utf8')} bytes）。`,
  );
}

function parseProcessResult<T>(stdout: string, stderr: string): T {
  if (/agentDelegatedToUser|user[- ]owned|user is controlling|用户.*控制/i.test(`${stdout}\n${stderr}`)) {
    throw new AppError('OUTLOOK_NOT_READY', 'Ego Lite 任务空间当前由用户控制。请把控制权交还给 Agent 后重试。');
  }
  try {
    return parseMarkedResult<T>(stdout);
  } catch {
    try {
      return parseMarkedResult<T>(stderr);
    } catch {
      throw new AppError(
        'EGO_BROWSER_ERROR',
        `ego-browser 未返回可识别的结果（stdout ${Buffer.byteLength(stdout, 'utf8')} bytes，stderr ${Buffer.byteLength(stderr, 'utf8')} bytes）。`,
      );
    }
  }
}

function stderrSuffix(stderr: string): string {
  const value = stderr.trim();
  if (!value) return '';
  return `：${value.slice(-1000)}`;
}

export class EgoRunner implements BrowserScriptRunner {
  private readonly command: string;
  private readonly args: string[];
  private readonly maxOutputBytes: number;

  constructor(options: EgoRunnerOptions = {}) {
    this.command = options.command ?? 'ego-browser';
    this.args = options.args ?? ['nodejs'];
    this.maxOutputBytes = options.maxOutputBytes ?? MAX_OUTPUT_BYTES;
  }

  async run<T>(body: string, timeoutMs = 30_000): Promise<EgoRunResult<T>> {
    return await new Promise<EgoRunResult<T>>((resolve, reject) => {
      const child = spawn(this.command, this.args, {
        stdio: ['pipe', 'pipe', 'pipe'],
        shell: false,
      });

      let stdout = '';
      let stderr = '';
      let outputBytes = 0;
      let timedOut = false;
      let settled = false;

      const finishReject = (error: AppError): void => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        reject(error);
      };

      const collect = (target: 'stdout' | 'stderr', chunk: Buffer): void => {
        outputBytes += chunk.byteLength;
        if (outputBytes > this.maxOutputBytes) {
          child.kill('SIGTERM');
          finishReject(new AppError('EGO_BROWSER_ERROR', 'ego-browser 输出超过安全上限。'));
          return;
        }

        if (target === 'stdout') stdout += chunk.toString('utf8');
        else stderr += chunk.toString('utf8');
      };

      child.stdout.on('data', (chunk: Buffer) => collect('stdout', chunk));
      child.stderr.on('data', (chunk: Buffer) => collect('stderr', chunk));

      child.on('error', (error: NodeJS.ErrnoException) => {
        const code = error.code === 'ENOENT' ? 'EGO_BROWSER_NOT_FOUND' : 'EGO_BROWSER_ERROR';
        const message = error.code === 'ENOENT'
          ? '未找到 ego-browser，请先安装并完成 Ego Lite 初始化。'
          : `无法启动 ego-browser：${error.message}`;
        finishReject(new AppError(code, message, { cause: error }));
      });

      child.on('close', (exitCode) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);

        if (timedOut) {
          reject(new AppError('TIMEOUT', `ego-browser 操作超过 ${timeoutMs}ms。`));
          return;
        }

        if (exitCode !== 0) {
          const detail = `${stdout}\n${stderr}`;
          if (/agentDelegatedToUser|user[- ]owned|user is controlling|用户.*控制/i.test(detail)) {
            reject(new AppError('OUTLOOK_NOT_READY', 'Ego Lite 任务空间当前由用户控制。请在完成登录或检查后把控制权交还给 Agent，再重新执行命令。'));
          } else {
            reject(new AppError('EGO_BROWSER_ERROR', `ego-browser 退出码为 ${exitCode}${stderrSuffix(stderr)}`));
          }
          return;
        }

        try {
          resolve({ stdout, stderr, value: parseProcessResult<T>(stdout, stderr) });
        } catch (error) {
          reject(error);
        }
      });

      const timer = setTimeout(() => {
        timedOut = true;
        child.kill('SIGTERM');
      }, timeoutMs);

      child.stdin.on('error', (error: NodeJS.ErrnoException) => {
        if (error.code !== 'EPIPE') {
          finishReject(new AppError('EGO_BROWSER_ERROR', `无法向 ego-browser 写入脚本：${error.message}`));
        }
      });
      child.stdin.end(body);
    });
  }
}
