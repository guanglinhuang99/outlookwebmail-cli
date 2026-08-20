import { randomUUID } from 'node:crypto';
import { homedir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import type { DatedMailListOptions, DatedMailListResult } from '../outlook/service.js';
import type { MailSummary } from '../types/mail.js';
import { AppError } from '../util/errors.js';
import { replaceFileAtomically } from '../export/obsidian.js';

interface WatchState {
  version: 1;
  updatedAt: string;
  scopes: Record<string, string[]>;
}

export interface MailWatchSource {
  listByDate(options: DatedMailListOptions): Promise<DatedMailListResult>;
}

export interface NewMessageEvent {
  type: 'message.new';
  observedAt: string;
  directory: string;
  date: string;
  message: MailSummary;
}

export interface MailWatchOptions {
  directory?: string | null;
  intervalSeconds?: number;
  emitExisting?: boolean;
  statePath?: string;
  iterations?: number;
  signal?: AbortSignal;
  now?: () => Date;
  sleep?: (milliseconds: number, signal?: AbortSignal) => Promise<void>;
  onEvent: (event: NewMessageEvent) => void | Promise<void>;
}

export interface MailWatchResult {
  polls: number;
  emitted: number;
  statePath: string;
  stopped: boolean;
}

export function defaultWatchStatePath(): string {
  return join(homedir(), '.webmail-cli', 'watch-state.json');
}

async function readState(path: string): Promise<WatchState> {
  try {
    const parsed = JSON.parse(await readFile(path, 'utf8')) as Partial<WatchState>;
    if (parsed.version !== 1 || !parsed.scopes || typeof parsed.scopes !== 'object') throw new Error('invalid state');
    return parsed as WatchState;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
      return { version: 1, updatedAt: new Date(0).toISOString(), scopes: {} };
    }
    throw new AppError('OPERATION_FAILED', `无法读取 watch 状态文件：${path}`, { cause: error });
  }
}

async function writeState(path: string, state: WatchState): Promise<void> {
  await mkdir(dirname(path), { recursive: true, mode: 0o700 });
  const temporary = join(dirname(path), `.${basename(path)}.${randomUUID()}.tmp`);
  await writeFile(temporary, `${JSON.stringify(state, null, 2)}\n`, { encoding: 'utf8', flag: 'wx', mode: 0o600 });
  await replaceFileAtomically(temporary, path);
}

async function collectToday(source: MailWatchSource, directory?: string | null): Promise<DatedMailListResult> {
  const messages: MailSummary[] = [];
  let cursor: string | null = null;
  let first: DatedMailListResult | null = null;
  for (let page = 0; page < 100; page += 1) {
    const result = await source.listByDate({ directory, limit: 100, cursor });
    first ??= result;
    messages.push(...result.messages);
    cursor = result.nextCursor;
    if (!cursor) return { ...first, messages, nextCursor: null, hasMore: false };
  }
  throw new AppError('OPERATION_FAILED', 'watch 单次轮询超过 100 页，已停止。');
}

function defaultSleep(milliseconds: number, signal?: AbortSignal): Promise<void> {
  if (signal?.aborted) return Promise.resolve();
  return new Promise(resolvePromise => {
    const timer = setTimeout(resolvePromise, milliseconds);
    timer.unref?.();
    signal?.addEventListener('abort', () => {
      clearTimeout(timer);
      resolvePromise();
    }, { once: true });
  });
}

export async function watchMail(source: MailWatchSource, options: MailWatchOptions): Promise<MailWatchResult> {
  const intervalSeconds = options.intervalSeconds ?? 30;
  if (!Number.isInteger(intervalSeconds) || intervalSeconds < 5 || intervalSeconds > 3_600) {
    throw new AppError('INVALID_ARGUMENT', '--interval 必须是 5 到 3600 之间的整数秒。');
  }
  const iterations = options.iterations ?? 0;
  if (!Number.isInteger(iterations) || iterations < 0) {
    throw new AppError('INVALID_ARGUMENT', '--iterations 必须是大于或等于 0 的整数。');
  }
  const statePath = resolve(options.statePath ?? defaultWatchStatePath());
  const state = await readState(statePath);
  const now = options.now ?? (() => new Date());
  const sleep = options.sleep ?? defaultSleep;
  let polls = 0;
  let emitted = 0;

  while (!options.signal?.aborted && (iterations === 0 || polls < iterations)) {
    const result = await collectToday(source, options.directory);
    polls += 1;
    const date = result.date ?? result.fromDate;
    const directory = result.directory?.path ?? (options.directory?.trim() || '收件箱');
    const scope = `${directory}\u001f${date}`;
    const previous = new Set(state.scopes[scope] ?? []);
    const isBaseline = !(scope in state.scopes) && !options.emitExisting;
    const newMessages = isBaseline ? [] : result.messages.filter(message => !previous.has(message.stableId)).reverse();
    for (const message of newMessages) {
      await options.onEvent({ type: 'message.new', observedAt: now().toISOString(), directory, date, message });
      emitted += 1;
    }
    state.scopes[scope] = result.messages.map(message => message.stableId).slice(0, 10_000);
    const scopeKeys = Object.keys(state.scopes).sort().reverse();
    for (const oldScope of scopeKeys.slice(30)) delete state.scopes[oldScope];
    state.updatedAt = now().toISOString();
    await writeState(statePath, state);
    if (options.signal?.aborted || (iterations !== 0 && polls >= iterations)) break;
    await sleep(intervalSeconds * 1_000, options.signal);
  }
  return { polls, emitted, statePath, stopped: Boolean(options.signal?.aborted) };
}
