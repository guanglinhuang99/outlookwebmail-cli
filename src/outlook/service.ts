import { createHash } from 'node:crypto';
import { constants, createReadStream } from 'node:fs';
import { copyFile, mkdir, mkdtemp, rm, stat, unlink } from 'node:fs/promises';
import { basename, extname, join, resolve } from 'node:path';
import type { BrowserBackend } from '../browser/backend.js';
import type {
  AttachmentDownloadResult,
  AttachmentSummary,
  ComposeActionResult,
  ComposeOptions,
  ComposeResult,
  ConversationResult,
  DownloadAllResult,
  DraftMessage,
  DraftUpdateOptions,
  FolderSummary,
  ForwardOptions,
  ForwardResult,
  MailExportFormat,
  MailExportResult,
  MailMessage,
  MailSummary,
  MessageActionResult,
  MessageLocator,
  MessageStateActionResult,
  ObsidianExportResult,
  ObsidianSyncResult,
  RawMessageRow,
  ReplyActionResult,
  ReplyResult,
} from '../types/mail.js';
import type { DoctorCheck, DoctorResult, InspectResult, MessageInspectResult, StatusResult } from '../types/inspect.js';
import { AppError } from '../util/errors.js';
import { messageFingerprint, normalizeText, stableMessageId } from '../util/text.js';
import { EgoInboxParser, type InboxParser } from './inbox-parser.js';
import { EgoMessageParser, type MessageParser } from './message-parser.js';
import { FolderParser } from './folder-parser.js';
import { SessionStore } from '../session/session-store.js';
import { detectOutlookState, isAllowedOutlookUrl } from './state.js';
import { MutationStore } from '../safety/mutation-store.js';
import {
  attachmentLink,
  chooseExportPaths,
  renderObsidianMarkdown,
  writeMarkdownAtomically,
} from '../export/obsidian.js';
import { syncObsidian } from '../sync/obsidian-sync.js';

const SEARCH_SELECTOR = 'input[role="combobox"][aria-label^="搜索"], input[role="combobox"][aria-label^="Search"]';
const EXIT_SEARCH_SELECTOR = 'button[aria-label="退出搜索"], button[aria-label="Exit search"]';
const SEARCH_STATE_SCRIPT = String.raw`
(() => {
  const inputs = Array.from(document.querySelectorAll('input[role="combobox"]'))
    .filter(el => /搜索|search/i.test((el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('placeholder') || '')));
  return { count: inputs.length, value: inputs.length === 1 ? inputs[0].value : null };
})()
`;
const SEARCH_RESULT_STATE_SCRIPT = String.raw`
(() => {
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const inputs = Array.from(document.querySelectorAll('input[role="combobox"]'))
    .filter(el => /搜索|search/i.test((el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('placeholder') || '')));
  const lists = Array.from(document.querySelectorAll('[role="listbox"]'))
    .filter(el => /邮件列表|mail list/i.test(el.getAttribute('aria-label') || ''));
  const first = lists.length === 1 ? lists[0].querySelector('[role="option"]') : null;
  const busy = Array.from(document.querySelectorAll('[aria-busy="true"], [role="progressbar"]'))
    .some(el => { const rect=el.getBoundingClientRect(); return rect.width>0&&rect.height>0; });
  const empty = /未找到结果|无结果|no results|nothing found/i.test(document.body.innerText || '');
  return { count: inputs.length, value: inputs.length === 1 ? inputs[0].value : null, listCount: lists.length,
    firstKey: first ? clean(first.id || first.getAttribute('aria-label') || first.textContent).slice(0, 500) : null, busy, empty };
})()
`;

interface SearchState {
  count: number;
  value: string | null;
}

interface SearchResultState extends SearchState {
  listCount: number;
  firstKey: string | null;
  busy: boolean;
  empty: boolean;
}

type RowBuckets = Map<string, RawMessageRow[]>;

function mergeVirtualPage(target: RowBuckets, rows: RawMessageRow[]): void {
  const pageOccurrences = new Map<string, number>();
  for (const row of rows) {
    const fingerprint = messageFingerprint(row);
    const occurrence = pageOccurrences.get(fingerprint) ?? 0;
    pageOccurrences.set(fingerprint, occurrence + 1);
    const bucket = target.get(fingerprint) ?? [];
    if (!bucket[occurrence]) bucket[occurrence] = row;
    target.set(fingerprint, bucket);
  }
}

function flattenRowBuckets(buckets: RowBuckets): RawMessageRow[] {
  return Array.from(buckets.values()).flat();
}

export interface InboxOptions {
  limit: number;
  unreadOnly?: boolean;
  directory?: string | null;
}

export interface MailListResult {
  directory: FolderSummary | null;
  messages: MailSummary[];
}

export interface DatedMailListResult extends MailListResult {
  date: string | null;
  fromDate: string;
  toDate: string;
  nextCursor: string | null;
  hasMore: boolean;
  complete: boolean;
  searchQuery: string | null;
  scope: 'selected-folder';
  filters: {
    sender: string | null;
    subject: string | null;
    unread: boolean;
    hasAttachments: boolean;
  };
}

export interface DatedMailListOptions {
  date?: string | null;
  fromDate?: string | null;
  toDate?: string | null;
  directory?: string | null;
  limit?: number;
  cursor?: string | null;
  sender?: string | null;
  subject?: string | null;
  unread?: boolean;
  hasAttachments?: boolean;
}

function currentShanghaiDate(): string {
  return new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
  }).format(new Date());
}

function normalizeMailDate(value: string | null | undefined, option = '--date', defaultToday = true): string | null {
  const normalized = value?.normalize('NFKC').trim() || (defaultToday ? currentShanghaiDate() : null);
  if (!normalized) return null;
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(normalized);
  if (!match) throw new AppError('INVALID_ARGUMENT', `${option} 必须使用 YYYY-MM-DD 格式。`);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new AppError('INVALID_ARGUMENT', `${option} 不是有效的日历日期。`);
  }
  return normalized;
}

interface DateWindow {
  date: string | null;
  fromDate: string;
  toDate: string;
}

function normalizeDateWindow(options: DatedMailListOptions): DateWindow {
  const hasDate = Boolean(options.date?.normalize('NFKC').trim());
  const hasRange = Boolean(options.fromDate?.normalize('NFKC').trim() || options.toDate?.normalize('NFKC').trim());
  if (hasDate && hasRange) {
    throw new AppError('INVALID_ARGUMENT', '--date 不能与 --from-date/--to-date 同时使用。');
  }
  if (hasRange) {
    const fromDate = normalizeMailDate(options.fromDate, '--from-date', false);
    const toDate = normalizeMailDate(options.toDate, '--to-date', false);
    if (!fromDate || !toDate) throw new AppError('INVALID_ARGUMENT', '日期范围必须同时提供 --from-date 和 --to-date。');
    if (fromDate > toDate) throw new AppError('INVALID_ARGUMENT', '--from-date 不能晚于 --to-date。');
    return { date: null, fromDate, toDate };
  }
  const date = normalizeMailDate(options.date, '--date', true)!;
  return { date, fromDate: date, toDate: date };
}

function validateMailId(id: string): void {
  if (!/^\d+$/.test(id) && !/^m_[A-Za-z0-9_-]{20}$/.test(id)) {
    throw new AppError('INVALID_ARGUMENT', '邮件 ID 必须是列表返回的数字短 ID 或 stableId。');
  }
}

function encodeCursor(scope: string, offset: number): string {
  return Buffer.from(JSON.stringify({ v: 1, scope, offset }), 'utf8').toString('base64url');
}

function decodeCursor(value: string | null | undefined, scope: string): number {
  if (!value?.trim()) return 0;
  try {
    const parsed = JSON.parse(Buffer.from(value.trim(), 'base64url').toString('utf8')) as unknown;
    if (!parsed || typeof parsed !== 'object') throw new Error('invalid cursor');
    const cursor = parsed as { v?: unknown; scope?: unknown; offset?: unknown };
    if (cursor.v !== 1 || cursor.scope !== scope || !Number.isInteger(cursor.offset) || Number(cursor.offset) < 0) {
      throw new Error('invalid cursor');
    }
    return Number(cursor.offset);
  } catch (error) {
    throw new AppError('INVALID_ARGUMENT', '--cursor 无效或不属于当前目录和筛选条件。', { cause: error });
  }
}

function normalizeRecipients(values: string[] | undefined): string[] {
  const recipients = new Map<string, string>();
  for (const value of values ?? []) {
    const normalized = value.normalize('NFKC').trim();
    const key = normalized.toLocaleLowerCase('en-US');
    if (normalized && !recipients.has(key)) recipients.set(key, normalized);
  }
  return Array.from(recipients.values());
}

async function sha256File(path: string): Promise<string> {
  return await new Promise<string>((resolveHash, rejectHash) => {
    const hash = createHash('sha256');
    const stream = createReadStream(path);
    stream.on('error', rejectHash);
    stream.on('data', chunk => hash.update(chunk));
    stream.on('end', () => resolveHash(hash.digest('hex')));
  });
}

async function copyWithoutOverwrite(source: string, directory: string, preferredFilename: string): Promise<{
  filename: string; path: string;
}> {
  const extension = extname(preferredFilename);
  const stem = basename(preferredFilename, extension) || 'attachment';
  for (let suffix = 0; suffix < 10_000; suffix += 1) {
    const filename = suffix === 0 ? `${stem}${extension}` : `${stem} (${suffix + 1})${extension}`;
    const target = join(directory, filename);
    try {
      await copyFile(source, target, constants.COPYFILE_EXCL);
      await unlink(source);
      return { filename, path: target };
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
    }
  }
  throw new AppError('OPERATION_FAILED', `文件 ${preferredFilename} 的同名版本过多，无法生成安全文件名。`);
}

function requirePageUrl(page: { url?: string; dialog?: unknown }): string {
  if (page.dialog) {
    throw new AppError('OUTLOOK_NOT_READY', 'Outlook 页面存在未处理的浏览器对话框。');
  }

  const url = page.url ?? null;
  if (!url) {
    throw new AppError('OUTLOOK_NOT_READY', '无法取得 Outlook 页面地址。');
  }

  return url;
}

export class OutlookService {
  private readonly inboxParser: InboxParser;
  private readonly messageParser: MessageParser;
  private readonly folderParser: FolderParser;

  constructor(
    private readonly backend: BrowserBackend,
    private readonly sessionStore = new SessionStore(),
    inboxParser?: InboxParser,
    messageParser?: MessageParser,
    private readonly mutationStore = new MutationStore(),
  ) {
    this.inboxParser = inboxParser ?? new EgoInboxParser(backend);
    this.messageParser = messageParser ?? new EgoMessageParser(backend);
    this.folderParser = new FolderParser(backend);
  }

  private backendName(): 'ego-lite' | 'playwright' {
    return this.backend.name ?? 'ego-lite';
  }

  async close(): Promise<void> {
    await this.backend.close?.();
  }

  private locatorForRow(row: RawMessageRow, stableId?: string): MessageLocator {
    return {
      stableHint: row.stableHint,
      stableId,
      subject: row.subject,
      senderName: row.senderName,
      senderAddress: row.senderAddress,
      receivedAt: row.receivedAt,
      receivedAtText: row.receivedAtText,
      preview: row.preview,
      hasAttachments: row.hasAttachments,
      unread: row.unread,
    };
  }

  private summariesForRows(rows: RawMessageRow[], firstId = 1): MailSummary[] {
    const collisions = new Map<string, number>();
    return rows.map((row, index) => {
      const baseStableId = stableMessageId(row);
      const occurrence = collisions.get(baseStableId) ?? 0;
      collisions.set(baseStableId, occurrence + 1);
      const stableId = occurrence === 0
        ? baseStableId
        : stableMessageId({ ...row, stableHint: `collision:${baseStableId}:${occurrence}` });
      return this.summaryForRow(row, String(firstId + index), stableId);
    });
  }

  private summaryForRow(row: RawMessageRow, id: string, stableId = stableMessageId(row)): MailSummary {
    return {
      id,
      stableId,
      sender: { name: row.senderName, address: row.senderAddress },
      subject: row.subject,
      receivedAt: row.receivedAt,
      receivedAtText: row.receivedAtText,
      preview: row.preview,
      unread: row.unread,
      hasAttachments: row.hasAttachments,
    };
  }

  private async writeListSession(
    source: string,
    messages: MailSummary[],
    rows: RawMessageRow[],
    metadata: { query?: string | null; directory?: string | null; complete?: boolean } = {},
  ): Promise<void> {
    const previous = await this.sessionStore.read();
    const stableEntries = new Map(Object.entries(previous?.stableMessages ?? {}));
    for (const [index, row] of rows.entries()) {
      const stableId = messages[index]!.stableId;
      stableEntries.delete(stableId);
      stableEntries.set(stableId, this.locatorForRow(row, stableId));
    }
    while (stableEntries.size > 1_000) stableEntries.delete(stableEntries.keys().next().value!);
    await this.sessionStore.write({
      version: 1,
      updatedAt: new Date().toISOString(),
      source,
      ...metadata,
      messages: Object.fromEntries(messages.map((message, index) => [message.id, this.locatorForRow(rows[index]!, message.stableId)])),
      stableMessages: Object.fromEntries(stableEntries),
    });
  }

  private async ensureReady(): Promise<void> {
    const observation = await this.backend.status();
    await this.requireAuthenticatedPage(observation.page, observation.snapshot);
    await this.requireMailReady();
  }

  private async requireMailReady(): Promise<void> {
    const probe = await this.backend.waitUntilMailReady(30_000);
    if (probe.ready) return;
    if (!isAllowedOutlookUrl(probe.url)) {
      await this.requireAuthenticatedPage({ url: probe.url ?? undefined }, '');
    }
    const diagnosticUrl = (() => {
      try {
        const url = new URL(probe.url ?? '');
        return `${url.protocol}//${url.host}${url.pathname}`;
      } catch {
        return '<unknown>';
      }
    })();
    throw new AppError(
      'OUTLOOK_NOT_READY',
      `Outlook 邮件页面尚未就绪。URL=${diagnosticUrl}；标题=${probe.title ?? '<unknown>'}；` +
      `搜索框=${probe.searchInputs}，Inbox=${probe.inboxFolders}，邮件列表=${probe.mailLists}，登录 iframe=${probe.loginFrames}，busy=${probe.busy}。` +
      '请等待页面加载完成；若仍停留在登录页，请在浏览器中完成登录后重试。',
    );
  }

  private async requireAuthenticatedPage(
    page: { url?: string; dialog?: unknown },
    snapshot: string,
  ): Promise<string> {
    const url = requirePageUrl(page);
    if (detectOutlookState(url, snapshot) !== 'AUTH_REQUIRED') return url;

    const backendName = this.backendName();

    try {
      const handoff = await this.backend.handoffForLogin();
      if (handoff.handedOff) {
        throw new AppError(
          'AUTH_REQUIRED',
          backendName === 'playwright'
            ? '已在 Playwright 浏览器中打开 Outlook。请完成登录，然后重新执行命令。'
            : '已在 Ego Lite 中打开 Outlook 并将页面交给你。请完成登录，再将控制权交还给 Agent，然后重新执行命令。',
        );
      }
      throw new AppError(
        'AUTH_REQUIRED',
        backendName === 'playwright'
          ? '已打开 Outlook，但浏览器窗口不可交互；请设置 WEBMAIL_HEADLESS=false 后完成登录。'
          : '已在 Ego Lite 中打开 Outlook，但未能把页面控制权交给你；请打开 Ego Lite 完成登录后重试。',
      );
    } catch (error) {
      if (error instanceof AppError && error.code === 'AUTH_REQUIRED') throw error;
      throw new AppError(
        'AUTH_REQUIRED',
        backendName === 'playwright'
          ? 'Outlook 尚未登录；Playwright 无法打开登录页面，请检查浏览器配置后重试。'
          : 'Outlook 尚未登录；已尝试在 Ego Lite 中打开登录页面，但控制权交接失败。请打开 Ego Lite 登录后重试。',
        { cause: error },
      );
    }
  }

  private async searchState(): Promise<SearchState> {
    const value = await this.backend.eval<SearchState>(SEARCH_STATE_SCRIPT);
    if (value.count !== 1) {
      throw new AppError('UI_CHANGED', `预期找到 1 个 Outlook 搜索框，实际找到 ${value.count} 个。`);
    }
    return value;
  }

  private async clearSearch(): Promise<void> {
    const state = await this.searchState();
    if (!state.value) return;
    await this.backend.clickAndWait(EXIT_SEARCH_SELECTOR, 0);
    for (let attempt = 0; attempt < 20; attempt += 1) {
      const clearedState = await this.searchState();
      if (!clearedState.value) return;
      await this.backend.wait(250);
    }
    throw new AppError('OUTLOOK_NOT_READY', '已点击 Outlook 的退出搜索按钮，但搜索状态在 5 秒内仍未清除。');
  }

  private async submitSearch(query: string): Promise<void> {
    await this.searchState();
    await this.backend.fillAndPress(SEARCH_SELECTOR, query, 'Enter', 0);
    let previousFingerprint: string | null = null;
    let stablePolls = 0;
    let last: SearchResultState | null = null;
    for (let attempt = 0; attempt < 40; attempt += 1) {
      last = await this.backend.eval<SearchResultState>(SEARCH_RESULT_STATE_SCRIPT);
      if (last.count !== 1) throw new AppError('UI_CHANGED', `预期找到 1 个 Outlook 搜索框，实际找到 ${last.count} 个。`);
      const fingerprint = `${last.listCount}:${last.firstKey ?? '<empty>'}:${last.empty}`;
      if (last.value === query && !last.busy && (last.listCount === 1 || last.empty)) {
        stablePolls = fingerprint === previousFingerprint ? stablePolls + 1 : 1;
        if (stablePolls >= 2) return;
      } else {
        stablePolls = 0;
      }
      previousFingerprint = fingerprint;
      await this.backend.wait(250);
    }
    throw new AppError(
      'OUTLOOK_NOT_READY',
      `Outlook 搜索结果在 10 秒内未稳定。query=${JSON.stringify(query)}；value=${JSON.stringify(last?.value ?? null)}；` +
      `mailLists=${last?.listCount ?? 0}；busy=${last?.busy ?? false}；empty=${last?.empty ?? false}。`,
    );
  }

  private async collectRows(target: number, resetScroll: boolean): Promise<RawMessageRow[]> {
    const rowsByFingerprint: RowBuckets = new Map();
    let rows = resetScroll ? await this.inboxParser.resetAndExtract() : await this.inboxParser.extract();
    for (let attempt = 0; attempt <= 5; attempt += 1) {
      mergeVirtualPage(rowsByFingerprint, rows);

      if (flattenRowBuckets(rowsByFingerprint).length >= target || attempt === 5) break;
      rows = await this.inboxParser.scrollAndExtract();
    }

    return flattenRowBuckets(rowsByFingerprint).slice(0, 100);
  }

  private async list(
    source: string,
    options: InboxOptions,
    directory: FolderSummary | null,
    metadata: { query?: string | null; complete?: boolean } = {},
  ): Promise<MailListResult> {
    if (!Number.isInteger(options.limit) || options.limit < 1 || options.limit > 100) {
      throw new AppError('INVALID_ARGUMENT', '--limit 必须是 1 到 100 之间的整数。');
    }

    const collectionTarget = options.unreadOnly ? 100 : options.limit;
    const rawRows = await this.collectRows(collectionTarget, true);
    const selectedRows = (options.unreadOnly ? rawRows.filter(row => row.unread === true) : rawRows)
      .slice(0, options.limit);
    const messages = this.summariesForRows(selectedRows);
    await this.writeListSession(source, messages, selectedRows, {
      directory: directory?.path ?? null,
      complete: false,
      ...metadata,
    });

    return { directory, messages };
  }

  private async selectDirectory(directory?: string | null): Promise<FolderSummary> {
    const normalizedDirectory = directory?.normalize('NFKC').trim() || null;
    const result = await this.backend.selectInboxFolder(normalizedDirectory);
    if (result.count > 1) {
      throw new AppError('AMBIGUOUS_FOLDER', `目录 ${normalizedDirectory} 对应 ${result.count} 个候选；请使用 folders 返回的完整 path。`);
    }
    if (result.count === 0 || !result.folder) {
      throw new AppError('FOLDER_NOT_FOUND', normalizedDirectory
        ? `Inbox 下找不到目录：${normalizedDirectory}`
        : '找不到 Inbox 目录。');
    }
    if (!result.selected) {
      throw new AppError('OPERATION_FAILED', `Outlook 未能选中目录：${result.folder.path}`);
    }
    return result.folder;
  }

  async status(): Promise<StatusResult> {
    const observation = await this.backend.status();
    const url = await this.requireAuthenticatedPage(observation.page, observation.snapshot);
    await this.requireMailReady();
    const detectedState = detectOutlookState(url, observation.snapshot);
    return {
      backend: this.backendName(),
      browser: observation.browserName ?? null,
      url,
      title: observation.title,
      state: detectedState === 'UNKNOWN' ? 'INBOX' : detectedState,
      mailReady: true,
      browserSession: observation.browserSession ?? null,
    };
  }

  async doctor(): Promise<DoctorResult> {
    const checks: DoctorCheck[] = [];
    const nodeMajor = Number(process.versions.node.split('.')[0]);
    checks.push({
      name: 'node',
      status: nodeMajor >= 24 ? 'pass' : 'fail',
      message: nodeMajor >= 24 ? `Node.js ${process.versions.node} 满足 >=24。` : `Node.js ${process.versions.node} 低于要求的 24。`,
    });

    let observation;
    try {
      observation = await this.backend.status();
      const backendName = this.backendName();
      const backendLabel = backendName === 'playwright' ? 'Playwright' : 'Ego Lite';
      checks.push({
        name: backendName,
        status: observation.connected ? 'pass' : 'fail',
        message: observation.connected ? `${backendLabel} 已连接。` : `${backendLabel} 未连接。`,
      });
      if (!observation.connected) return { ok: false, checks };
    } catch (error) {
      const backendName = this.backendName();
      const backendLabel = backendName === 'playwright' ? 'Playwright' : 'Ego Lite';
      checks.push({ name: backendName, status: 'fail', message: `${backendLabel} 检查失败：${error instanceof Error ? error.message : String(error)}` });
      return { ok: false, checks };
    }

    let url: string;
    try {
      url = requirePageUrl(observation.page);
    } catch (error) {
      checks.push({ name: 'authentication', status: 'fail', message: error instanceof Error ? error.message : String(error) });
      return { ok: false, checks };
    }
    const state = detectOutlookState(url, observation.snapshot);
    if (state === 'AUTH_REQUIRED') {
      const location = this.backendName() === 'playwright' ? 'Playwright 浏览器' : 'Ego Lite';
      checks.push({ name: 'authentication', status: 'fail', message: `Outlook 尚未登录；请在${location}中登录后重试。` });
      return { ok: false, checks };
    }
    let ready = false;
    let domMessage = 'DOM 尚未检查。';
    try {
      const probe = await this.backend.waitUntilMailReady(30_000);
      ready = probe.ready;
      domMessage = `DOM：搜索框 ${probe.searchInputs}，Inbox ${probe.inboxFolders}，邮件列表 ${probe.mailLists}，登录 iframe ${probe.loginFrames}，busy=${probe.busy}。`;
    } catch (error) {
      domMessage = `DOM 等待失败：${error instanceof Error ? error.message : String(error)}`;
    }
    checks.push({ name: 'authentication', status: ready ? 'pass' : 'warn', message: ready ? `Outlook 已登录，当前状态：${state}。` : 'Outlook 地址可访问，但邮件页面尚未就绪。' });
    checks.push({ name: 'dom', status: ready ? 'pass' : 'fail', message: domMessage });

    return { ok: checks.every(check => check.status !== 'fail'), checks };
  }

  async inspect(): Promise<InspectResult> {
    const observation = await this.backend.inspect();
    const url = await this.requireAuthenticatedPage(observation.page, observation.snapshot);
    return {
      backend: this.backendName(),
      capturedAt: new Date().toISOString(),
      state: detectOutlookState(url, observation.snapshot),
      ...observation,
    };
  }

  async inspectMessage(): Promise<MessageInspectResult> {
    const observation = await this.backend.inspectMessage();
    const url = await this.requireAuthenticatedPage(observation.page, observation.snapshot);
    return {
      backend: this.backendName(),
      capturedAt: new Date().toISOString(),
      state: detectOutlookState(url, observation.snapshot),
      ...observation,
    };
  }

  async inbox(options: InboxOptions): Promise<MailListResult> {
    await this.ensureReady();
    await this.clearSearch();
    const directory = await this.selectDirectory(options.directory);
    return await this.list(`folder:${directory.path}`, options, directory);
  }

  async search(query: string, options: InboxOptions): Promise<MailListResult> {
    const normalizedQuery = query.normalize('NFKC').trim();
    if (!normalizedQuery) throw new AppError('INVALID_ARGUMENT', '搜索内容不能为空。');
    await this.ensureReady();
    await this.clearSearch();
    await this.submitSearch(normalizedQuery);
    return await this.list(`search:${normalizedQuery}`, options, null, { query: normalizedQuery, complete: false });
  }

  async listByDate(options: DatedMailListOptions = {}): Promise<DatedMailListResult> {
    const window = normalizeDateWindow(options);
    const limit = options.limit ?? 20;
    if (!Number.isInteger(limit) || limit < 1 || limit > 100) {
      throw new AppError('INVALID_ARGUMENT', '--limit 必须是 1 到 100 之间的整数。');
    }
    const sender = options.sender?.normalize('NFKC').trim() || null;
    const subject = options.subject?.normalize('NFKC').trim() || null;
    await this.ensureReady();
    await this.clearSearch();
    const directory = await this.selectDirectory(options.directory);
    const searchQuery = window.date ? `received:${window.date}` : null;
    if (searchQuery) await this.submitSearch(searchQuery);
    const scope = createHash('sha256').update(JSON.stringify({
      directory: directory.path,
      ...window,
      sender: normalizeText(sender),
      subject: normalizeText(subject),
      unread: Boolean(options.unread),
      hasAttachments: Boolean(options.hasAttachments),
    })).digest('base64url').slice(0, 20);
    const offset = decodeCursor(options.cursor, scope);
    const target = offset + limit + 1;
    const matchingRows: RowBuckets = new Map();
    let rows = await this.inboxParser.resetAndExtract();
    let previousPage = '';
    let scanComplete = false;
    let sawRows = false;
    let sawParsedDate = false;
    for (let attempt = 0; attempt < 100; attempt += 1) {
      const filteredRows: RawMessageRow[] = [];
      for (const row of rows) {
        sawRows = true;
        const receivedDate = row.receivedAt?.slice(0, 10) ?? null;
        if (receivedDate) sawParsedDate = true;
        if (!receivedDate || receivedDate < window.fromDate || receivedDate > window.toDate) continue;
        if (sender && !normalizeText(row.senderAddress || row.senderName).includes(normalizeText(sender))) continue;
        if (subject && !normalizeText(row.subject).includes(normalizeText(subject))) continue;
        if (options.unread && row.unread !== true) continue;
        if (options.hasAttachments && row.hasAttachments !== true) continue;
        filteredRows.push(row);
      }
      mergeVirtualPage(matchingRows, filteredRows);
      if (flattenRowBuckets(matchingRows).length >= target) {
        break;
      }
      if (rows.some(row => row.receivedAt && row.receivedAt.slice(0, 10) < window.fromDate)) {
        scanComplete = true;
        break;
      }
      const page = rows.map(messageFingerprint).join('\u001e');
      if (!rows.length || page === previousPage) {
        scanComplete = true;
        break;
      }
      previousPage = page;
      rows = await this.inboxParser.scrollAndExtract();
    }
    if (searchQuery && sawRows && !sawParsedDate) {
      throw new AppError('UI_CHANGED', 'Outlook 日期搜索返回了邮件，但无法解析任何邮件日期；已拒绝把结果误报为空。');
    }
    if (!scanComplete && flattenRowBuckets(matchingRows).length === 0) {
      throw new AppError('OUTLOOK_NOT_READY', `${window.fromDate} 至 ${window.toDate} 的搜索尚未完成，当前未取得可验证结果。`);
    }

    const allRows = flattenRowBuckets(matchingRows);
    const selectedRows = allRows.slice(offset, offset + limit);
    const hasMore = allRows.length > offset + limit || !scanComplete;
    const messages = this.summariesForRows(selectedRows, offset + 1);
    const source = window.date
      ? `date:${directory.path}:${window.date}`
      : `range:${directory.path}:${window.fromDate}:${window.toDate}`;
    await this.writeListSession(source, messages, selectedRows, {
      query: searchQuery,
      directory: directory.path,
      complete: scanComplete,
    });
    return {
      date: window.date,
      fromDate: window.fromDate,
      toDate: window.toDate,
      directory,
      messages,
      hasMore,
      complete: scanComplete,
      searchQuery,
      scope: 'selected-folder',
      nextCursor: hasMore ? encodeCursor(scope, offset + limit) : null,
      filters: { sender, subject, unread: Boolean(options.unread), hasAttachments: Boolean(options.hasAttachments) },
    };
  }

  async today(directoryName?: string | null): Promise<DatedMailListResult> {
    return await this.listByDate({ directory: directoryName });
  }

  async folders(): Promise<{ root: '收件箱'; folders: FolderSummary[] }> {
    await this.ensureReady();
    return { root: '收件箱', folders: await this.folderParser.extract() };
  }

  private async sessionLocator(id: string): Promise<MessageLocator> {
    const session = await this.sessionStore.read();
    if (!session) {
      throw new AppError('INVALID_ARGUMENT', '当前没有邮件列表 Session；请先运行 list、today、inbox 或 search。');
    }
    const message = /^\d+$/.test(id) ? session.messages[id] : session.stableMessages?.[id];
    if (!message) {
      throw new AppError('INVALID_ARGUMENT', `Session 中不存在邮件 ID ${id}；请重新运行 list、today、inbox 或 search。`);
    }
    return message;
  }

  async read(id: string): Promise<MailMessage> {
    validateMailId(id);
    await this.ensureReady();
    const locator = await this.sessionLocator(id);
    const result = await this.messageParser.openAndExtract(locator);
    if (result.matchCount > 1) {
      throw new AppError('AMBIGUOUS_MESSAGE', `邮件 ID ${id} 对应 ${result.matchCount} 个候选，已拒绝自动打开。`);
    }
    if (result.matchCount === 0) {
      throw new AppError('MESSAGE_NOT_FOUND', `无法重新定位邮件 ID ${id}。`);
    }
    if (!result.message) {
      throw new AppError('UI_CHANGED', `已定位邮件 ID ${id}，但阅读窗格未能按已确认结构解析。`);
    }
    if (!result.message.subject || !result.message.bodyText) {
      throw new AppError('UI_CHANGED', '邮件已打开，但主题或正文无法按已确认结构解析。');
    }

    return {
      id,
      stableId: locator.stableId ?? stableMessageId(locator),
      subject: result.message.subject,
      from: { name: result.message.fromName, address: result.message.fromAddress },
      to: result.message.to,
      cc: result.message.cc,
      receivedAt: result.message.receivedAt,
      receivedAtText: result.message.receivedAtText,
      bodyText: result.message.bodyText,
      bodyTruncated: Boolean(result.message.bodyTruncated),
      bodyBytes: result.message.bodyBytes ?? Buffer.byteLength(result.message.bodyText, 'utf8'),
      attachments: result.message.attachments.map((attachment, index) => ({
        id: String(index + 1),
        filename: attachment.filename,
        sizeText: attachment.sizeText,
      })),
      unreadRestored: await this.restoreUnreadAfterRead(id, locator),
    };
  }

  private async restoreUnreadAfterRead(id: string, locator: MessageLocator): Promise<boolean> {
    if (locator.unread !== true) return false;
    const result = await this.backend.setReadState(locator, true);
    this.assertStatePerformed(id, result);
    return true;
  }

  async attachments(id: string): Promise<{ attachments: AttachmentSummary[] }> {
    const message = await this.read(id);
    return { attachments: message.attachments };
  }

  private assertPerformed(id: string, result: MessageActionResult): void {
    if (result.matchCount > 1 || result.status === 'message_ambiguous') {
      throw new AppError('AMBIGUOUS_MESSAGE', `邮件 ID ${id} 对应多个候选，已拒绝操作。`);
    }
    if (result.matchCount === 0 || result.status === 'message_not_found') {
      throw new AppError('MESSAGE_NOT_FOUND', `无法重新定位邮件 ID ${id}。`);
    }
    if (result.status === 'folder_not_found') {
      throw new AppError('FOLDER_NOT_FOUND', '找不到完全匹配的目标目录。');
    }
    if (result.status === 'folder_ambiguous') {
      throw new AppError('AMBIGUOUS_FOLDER', `目标目录对应 ${result.folderMatches ?? '多个'} 个候选，已拒绝移动。`);
    }
    if (result.status === 'attachment_not_found') {
      throw new AppError('ATTACHMENT_NOT_FOUND', '邮件中不存在指定附件 ID。');
    }
    if (!result.performed || result.status !== 'performed') {
      throw new AppError('OPERATION_FAILED', `Outlook 操作失败：${result.status}。`);
    }
    if (!result.verified) {
      throw new AppError('OPERATION_UNKNOWN', 'Outlook 已接受操作，但无法验证最终页面状态；禁止自动重试。');
    }
  }

  private async recordMutationError(requestId: string, appError: AppError, action: string, mailId: string): Promise<void> {
    if (['OPERATION_UNKNOWN', 'PLAYWRIGHT_TIMEOUT', 'TIMEOUT'].includes(appError.code)) {
      await this.mutationStore.uncertain(requestId, appError.code, { action, mailId });
      return;
    }
    await this.mutationStore.fail(requestId, appError.code, { action, mailId });
  }

  async delete(
    id: string,
    confirmed: boolean,
    requestId?: string | null,
  ): Promise<{ id: string; deleted: true; verified: true; sessionInvalidated: true; requestId: string; deduplicated: boolean }> {
    if (!confirmed) {
      throw new AppError('CONFIRMATION_REQUIRED', '删除邮件必须显式提供 --yes；邮件将移入“已删除邮件”。');
    }
    validateMailId(id);
    const normalizedRequestId = this.mutationStore.validateRequestId(requestId);
    const payloadHash = this.mutationStore.payloadHash('delete', { id });
    const prior = await this.mutationStore.prior<Awaited<ReturnType<OutlookService['delete']>>>(
      normalizedRequestId, 'delete', payloadHash, id,
    );
    if (prior) return { ...prior, deduplicated: true };
    const locator = await this.sessionLocator(id);
    const auditId = locator.stableId ?? stableMessageId(locator);
    await this.mutationStore.begin(normalizedRequestId, 'delete', payloadHash, auditId);
    try {
      await this.ensureReady();
      const result = await this.backend.deleteMessage(locator);
      this.assertPerformed(id, result);
    } catch (error) {
      const appError = error instanceof AppError ? error : new AppError('OPERATION_FAILED', '删除邮件失败。', { cause: error });
      await this.recordMutationError(normalizedRequestId, appError, 'delete', auditId);
      throw error;
    }
    const response = { id, deleted: true as const, verified: true as const, sessionInvalidated: true as const, requestId: normalizedRequestId, deduplicated: false };
    await this.mutationStore.succeed(normalizedRequestId, response, { action: 'delete', mailId: auditId });
    await this.sessionStore.clear();
    return response;
  }

  async move(
    id: string,
    folder: string,
    confirmed: boolean,
    requestId?: string | null,
  ): Promise<{ id: string; folder: string; moved: true; verified: true; sessionInvalidated: true; requestId: string; deduplicated: boolean }> {
    if (!confirmed) throw new AppError('CONFIRMATION_REQUIRED', '移动邮件必须显式提供 --yes。');
    validateMailId(id);
    const normalizedFolder = folder.normalize('NFKC').trim();
    if (!normalizedFolder) throw new AppError('INVALID_ARGUMENT', '目标目录不能为空。');
    const normalizedRequestId = this.mutationStore.validateRequestId(requestId);
    const payloadHash = this.mutationStore.payloadHash('move', { id, folder: normalizedFolder });
    const prior = await this.mutationStore.prior<Awaited<ReturnType<OutlookService['move']>>>(
      normalizedRequestId, 'move', payloadHash, id,
    );
    if (prior) return { ...prior, deduplicated: true };
    const locator = await this.sessionLocator(id);
    const auditId = locator.stableId ?? stableMessageId(locator);
    await this.mutationStore.begin(normalizedRequestId, 'move', payloadHash, auditId);
    try {
      await this.ensureReady();
      const result = await this.backend.moveMessage(locator, normalizedFolder);
      this.assertPerformed(id, result);
    } catch (error) {
      const appError = error instanceof AppError ? error : new AppError('OPERATION_FAILED', '移动邮件失败。', { cause: error });
      await this.recordMutationError(normalizedRequestId, appError, 'move', auditId);
      throw error;
    }
    const response = { id, folder: normalizedFolder, moved: true as const, verified: true as const, sessionInvalidated: true as const, requestId: normalizedRequestId, deduplicated: false };
    await this.mutationStore.succeed(normalizedRequestId, response, { action: 'move', mailId: auditId });
    await this.sessionStore.clear();
    return response;
  }

  private assertReplyPerformed(id: string, result: ReplyActionResult): void {
    if (result.matchCount > 1 || result.status === 'message_ambiguous') {
      throw new AppError('AMBIGUOUS_MESSAGE', `邮件 ID ${id} 对应多个候选，已拒绝回复。`);
    }
    if (result.matchCount === 0 || result.status === 'message_not_found') {
      throw new AppError('MESSAGE_NOT_FOUND', `无法重新定位邮件 ID ${id}。`);
    }
    if (result.status === 'reply_control_ambiguous') {
      throw new AppError('UI_CHANGED', 'Outlook 页面中出现多个匹配的回复按钮，已拒绝继续。');
    }
    if (!result.verified) {
      const handoff = result.handedOff ? '；当前回复窗口已交给用户手工检查' : '';
      throw new AppError(result.performed && !result.draft ? 'OPERATION_UNKNOWN' : 'OPERATION_FAILED', `Outlook 回复操作失败：${result.status}${handoff}。`);
    }
  }

  private assertComposePerformed(result: ComposeActionResult, operation: string): void {
    if ((result.matchCount && result.matchCount > 1) || result.status === 'message_ambiguous') {
      throw new AppError('AMBIGUOUS_MESSAGE', `${operation}对应多个邮件候选，已拒绝继续。`);
    }
    if (result.status === 'message_not_found') throw new AppError('MESSAGE_NOT_FOUND', `${operation}的原邮件或草稿不存在。`);
    if (!result.verified || !result.performed) {
      throw new AppError(result.performed && !result.draft ? 'OPERATION_UNKNOWN' : 'OPERATION_FAILED', `${operation}失败：${result.status}。`);
    }
  }

  private async normalizeAttachmentPaths(paths: string[] | undefined): Promise<string[]> {
    const normalized = Array.from(new Set((paths ?? []).map(path => resolve(path))));
    for (const path of normalized) {
      let info;
      try {
        info = await stat(path);
      } catch (error) {
        throw new AppError('INVALID_ARGUMENT', `附件文件不存在：${path}`, { cause: error });
      }
      if (!info.isFile()) throw new AppError('INVALID_ARGUMENT', `附件路径不是普通文件：${path}`);
    }
    return normalized;
  }

  async compose(options: ComposeOptions, requestId?: string | null): Promise<ComposeResult> {
    const normalized: ComposeOptions = {
      to: normalizeRecipients(options.to), cc: normalizeRecipients(options.cc), bcc: normalizeRecipients(options.bcc),
      subject: options.subject.normalize('NFKC').trim(),
      content: options.content,
      attachments: await this.normalizeAttachmentPaths(options.attachments),
      draft: options.draft,
    };
    if (!normalized.subject) throw new AppError('INVALID_ARGUMENT', '--subject 不能为空。');
    if (!normalized.content.trim()) throw new AppError('INVALID_ARGUMENT', '--content 不能为空。');
    if (!normalized.draft && normalized.to.length === 0) throw new AppError('INVALID_ARGUMENT', '自动发送新邮件时至少需要一个 --to。');
    const normalizedRequestId = normalized.draft ? null : this.mutationStore.validateRequestId(requestId);
    const payloadHash = normalizedRequestId ? this.mutationStore.payloadHash('compose-send', normalized) : null;
    const auditId = `out_${createHash('sha256').update(JSON.stringify({ to: normalized.to, subject: normalized.subject })).digest('base64url').slice(0, 20)}`;
    if (normalizedRequestId && payloadHash) {
      const prior = await this.mutationStore.prior<ComposeResult>(normalizedRequestId, 'compose-send', payloadHash, auditId);
      if (prior) return { ...prior, deduplicated: true };
      await this.mutationStore.begin(normalizedRequestId, 'compose-send', payloadHash, auditId);
    }
    let result: ComposeActionResult;
    try {
      await this.ensureReady();
      result = await this.backend.composeMessage(normalized);
      this.assertComposePerformed(result, '新建邮件');
    } catch (error) {
      if (normalizedRequestId) {
        const appError = error instanceof AppError ? error : new AppError('OPERATION_FAILED', '新建邮件失败。', { cause: error });
        await this.recordMutationError(normalizedRequestId, appError, 'compose-send', auditId);
      }
      throw error;
    }
    const response: ComposeResult = normalized.draft ? {
      draft: true, sent: false, requiresManualSend: true, handedOff: result.handedOff,
      verified: true, attachmentCount: result.attachmentCount,
    } : {
      draft: false, sent: true, requiresManualSend: false, handedOff: false,
      verified: true, attachmentCount: result.attachmentCount,
      requestId: normalizedRequestId!, deduplicated: false,
    };
    if (normalizedRequestId) await this.mutationStore.succeed(normalizedRequestId, response, { action: 'compose-send', mailId: auditId });
    return response;
  }

  async forward(
    id: string,
    options: ForwardOptions,
    requestId?: string | null,
  ): Promise<ForwardResult> {
    validateMailId(id);
    const normalized: ForwardOptions = {
      to: normalizeRecipients(options.to), cc: normalizeRecipients(options.cc), bcc: normalizeRecipients(options.bcc),
      content: options.content,
      attachments: await this.normalizeAttachmentPaths(options.attachments),
      draft: options.draft,
    };
    if (!normalized.content.trim()) throw new AppError('INVALID_ARGUMENT', '--content 不能为空。');
    if (!normalized.draft && normalized.to.length === 0) throw new AppError('INVALID_ARGUMENT', '自动转发时至少需要一个 --to。');
    const normalizedRequestId = normalized.draft ? null : this.mutationStore.validateRequestId(requestId);
    const payloadHash = normalizedRequestId ? this.mutationStore.payloadHash('forward-send', { id, ...normalized }) : null;
    if (normalizedRequestId && payloadHash) {
      const prior = await this.mutationStore.prior<ForwardResult>(normalizedRequestId, 'forward-send', payloadHash, id);
      if (prior) return { ...prior, deduplicated: true };
    }
    const locator = await this.sessionLocator(id);
    const auditId = locator.stableId ?? stableMessageId(locator);
    if (normalizedRequestId && payloadHash) await this.mutationStore.begin(normalizedRequestId, 'forward-send', payloadHash, auditId);
    let result: ComposeActionResult;
    try {
      await this.ensureReady();
      result = await this.backend.forwardMessage(locator, normalized);
      this.assertComposePerformed(result, '转发邮件');
    } catch (error) {
      if (normalizedRequestId) {
        const appError = error instanceof AppError ? error : new AppError('OPERATION_FAILED', '转发邮件失败。', { cause: error });
        await this.recordMutationError(normalizedRequestId, appError, 'forward-send', auditId);
      }
      throw error;
    }
    const response: ForwardResult = normalized.draft ? {
      id, draft: true, sent: false, requiresManualSend: true, handedOff: result.handedOff,
      verified: true, attachmentCount: result.attachmentCount, originalAttachmentsPreserved: true,
    } : {
      id, draft: false, sent: true, requiresManualSend: false, handedOff: false,
      verified: true, attachmentCount: result.attachmentCount, originalAttachmentsPreserved: true,
      requestId: normalizedRequestId!, deduplicated: false,
    };
    if (normalizedRequestId) await this.mutationStore.succeed(normalizedRequestId, response, { action: 'forward-send', mailId: auditId });
    return response;
  }

  async reply(
    id: string,
    content: string,
    draft = true,
    replyAll = false,
    requestId?: string | null,
  ): Promise<ReplyResult> {
    validateMailId(id);
    if (!content.trim()) throw new AppError('INVALID_ARGUMENT', '--content 不能为空。');
    const normalizedRequestId = draft ? null : this.mutationStore.validateRequestId(requestId);
    const payloadHash = normalizedRequestId
      ? this.mutationStore.payloadHash('reply-send', { id, content, replyAll })
      : null;
    if (normalizedRequestId && payloadHash) {
      const prior = await this.mutationStore.prior<ReplyResult>(normalizedRequestId, 'reply-send', payloadHash, id);
      if (prior) return { ...prior, deduplicated: true };
    }
    const locator = await this.sessionLocator(id);
    const auditId = locator.stableId ?? stableMessageId(locator);
    if (normalizedRequestId && payloadHash) {
      await this.mutationStore.begin(normalizedRequestId, 'reply-send', payloadHash, auditId);
    }
    let result: ReplyActionResult;
    try {
      await this.ensureReady();
      result = await this.backend.replyMessage(locator, content, draft, replyAll);
      this.assertReplyPerformed(id, result);
    } catch (error) {
      if (normalizedRequestId) {
        const appError = error instanceof AppError ? error : new AppError('OPERATION_FAILED', '发送回复失败。', { cause: error });
        await this.recordMutationError(normalizedRequestId, appError, 'reply-send', auditId);
      }
      throw error;
    }

    if (draft) {
      if (result.status !== 'draft_ready' || !result.performed || !result.handedOff) {
        throw new AppError('OPERATION_FAILED', '回复草稿已处理，但未能验证草稿窗口已交给用户。');
      }
      return {
        id,
        draft: true,
        replyAll,
        sent: false,
        requiresManualSend: true,
        handedOff: true,
        verified: true,
      };
    }

    if (result.status !== 'sent' || !result.performed) {
      throw new AppError('OPERATION_FAILED', 'Outlook 未能验证回复已经发送。');
    }
    const response: ReplyResult = {
      id,
      draft: false,
      replyAll,
      sent: true,
      requiresManualSend: false,
      handedOff: false,
      verified: true,
      requestId: normalizedRequestId!,
      deduplicated: false,
    };
    await this.mutationStore.succeed(normalizedRequestId!, response, { action: 'reply-send', mailId: auditId });
    return response;
  }

  async drafts(options: { limit: number }): Promise<MailListResult> {
    await this.ensureReady();
    await this.clearSearch();
    const directory = await this.backend.selectSystemFolder('草稿');
    if (directory.count !== 1 || !directory.selected || !directory.folder) {
      throw new AppError('FOLDER_NOT_FOUND', '无法选择 Outlook 草稿目录。');
    }
    return await this.list('system:草稿', { limit: options.limit }, directory.folder);
  }

  async readDraft(id: string): Promise<DraftMessage> {
    validateMailId(id);
    const locator = await this.sessionLocator(id);
    await this.ensureReady();
    const opened = await this.backend.openDraft(locator, true);
    if (opened.matchCount > 1) throw new AppError('AMBIGUOUS_MESSAGE', `草稿 ID ${id} 对应多个候选。`);
    if (opened.matchCount !== 1 || !opened.draft) throw new AppError('MESSAGE_NOT_FOUND', `无法定位草稿 ID ${id}。`);
    return {
      id, stableId: locator.stableId ?? stableMessageId(locator),
      to: opened.draft.to, cc: opened.draft.cc, bcc: opened.draft.bcc,
      subject: opened.draft.subject, bodyText: opened.draft.bodyText,
      attachments: opened.draft.attachments.map((attachment, index) => ({ id: String(index + 1), ...attachment })),
    };
  }

  async updateDraft(id: string, options: DraftUpdateOptions): Promise<{
    id: string; updated: true; verified: true; sessionInvalidated: true; attachmentCount: number;
  }> {
    validateMailId(id);
    if (options.to === undefined && options.cc === undefined && options.bcc === undefined
      && options.subject === undefined && options.content === undefined && options.attachments.length === 0) {
      throw new AppError('INVALID_ARGUMENT', 'draft-update 至少需要一个要修改的字段或附件。');
    }
    const normalized: DraftUpdateOptions = {
      to: options.to === undefined ? undefined : normalizeRecipients(options.to),
      cc: options.cc === undefined ? undefined : normalizeRecipients(options.cc),
      bcc: options.bcc === undefined ? undefined : normalizeRecipients(options.bcc),
      subject: options.subject?.normalize('NFKC').trim(),
      content: options.content,
      attachments: await this.normalizeAttachmentPaths(options.attachments),
    };
    const locator = await this.sessionLocator(id);
    await this.ensureReady();
    const result = await this.backend.updateDraft(locator, normalized);
    this.assertComposePerformed(result, '修改草稿');
    await this.sessionStore.clear();
    return { id, updated: true, verified: true, sessionInvalidated: true, attachmentCount: result.attachmentCount };
  }

  async sendDraft(id: string, requestId?: string | null): Promise<{
    id: string; sent: true; verified: true; sessionInvalidated: true; requestId: string; deduplicated: boolean;
  }> {
    validateMailId(id);
    const normalizedRequestId = this.mutationStore.validateRequestId(requestId);
    const payloadHash = this.mutationStore.payloadHash('draft-send', { id });
    const prior = await this.mutationStore.prior<Awaited<ReturnType<OutlookService['sendDraft']>>>(normalizedRequestId, 'draft-send', payloadHash, id);
    if (prior) return { ...prior, deduplicated: true };
    const locator = await this.sessionLocator(id);
    const auditId = locator.stableId ?? stableMessageId(locator);
    await this.mutationStore.begin(normalizedRequestId, 'draft-send', payloadHash, auditId);
    try {
      await this.ensureReady();
      const result = await this.backend.sendDraft(locator);
      this.assertComposePerformed(result, '发送草稿');
    } catch (error) {
      const appError = error instanceof AppError ? error : new AppError('OPERATION_FAILED', '发送草稿失败。', { cause: error });
      await this.recordMutationError(normalizedRequestId, appError, 'draft-send', auditId);
      throw error;
    }
    const response = { id, sent: true as const, verified: true as const, sessionInvalidated: true as const, requestId: normalizedRequestId, deduplicated: false };
    await this.mutationStore.succeed(normalizedRequestId, response, { action: 'draft-send', mailId: auditId });
    await this.sessionStore.clear();
    return response;
  }

  async discardDraft(id: string, confirmed: boolean, requestId?: string | null): Promise<{
    id: string; discarded: true; verified: true; sessionInvalidated: true; requestId: string; deduplicated: boolean;
  }> {
    if (!confirmed) throw new AppError('CONFIRMATION_REQUIRED', '放弃草稿必须显式提供 --yes。');
    validateMailId(id);
    const normalizedRequestId = this.mutationStore.validateRequestId(requestId);
    const payloadHash = this.mutationStore.payloadHash('draft-discard', { id });
    const prior = await this.mutationStore.prior<Awaited<ReturnType<OutlookService['discardDraft']>>>(normalizedRequestId, 'draft-discard', payloadHash, id);
    if (prior) return { ...prior, deduplicated: true };
    const locator = await this.sessionLocator(id);
    const auditId = locator.stableId ?? stableMessageId(locator);
    await this.mutationStore.begin(normalizedRequestId, 'draft-discard', payloadHash, auditId);
    try {
      await this.ensureReady();
      const result = await this.backend.discardDraft(locator);
      this.assertPerformed(id, result);
    } catch (error) {
      const appError = error instanceof AppError ? error : new AppError('OPERATION_FAILED', '放弃草稿失败。', { cause: error });
      await this.recordMutationError(normalizedRequestId, appError, 'draft-discard', auditId);
      throw error;
    }
    const response = { id, discarded: true as const, verified: true as const, sessionInvalidated: true as const, requestId: normalizedRequestId, deduplicated: false };
    await this.mutationStore.succeed(normalizedRequestId, response, { action: 'draft-discard', mailId: auditId });
    await this.sessionStore.clear();
    return response;
  }

  private assertStatePerformed(id: string, result: MessageStateActionResult): void {
    this.assertPerformed(id, result);
  }

  async markRead(id: string, unread: boolean): Promise<{ id: string; unread: boolean; changed: boolean; verified: true }> {
    validateMailId(id);
    await this.ensureReady();
    const result = await this.backend.setReadState(await this.sessionLocator(id), unread);
    this.assertStatePerformed(id, result);
    return { id, unread, changed: Boolean(result.changed), verified: true };
  }

  async flag(id: string, flagged: boolean): Promise<{ id: string; flagged: boolean; changed: boolean; verified: true }> {
    validateMailId(id);
    await this.ensureReady();
    const result = await this.backend.setFlagState(await this.sessionLocator(id), flagged);
    this.assertStatePerformed(id, result);
    return { id, flagged, changed: Boolean(result.changed), verified: true };
  }

  async categorize(id: string, category: string, applied: boolean): Promise<{
    id: string; category: string; applied: boolean; changed: boolean; verified: true;
  }> {
    validateMailId(id);
    const normalizedCategory = category.normalize('NFKC').trim();
    if (!normalizedCategory) throw new AppError('INVALID_ARGUMENT', '--category 不能为空。');
    await this.ensureReady();
    const result = await this.backend.setCategoryState(await this.sessionLocator(id), normalizedCategory, applied);
    this.assertStatePerformed(id, result);
    return { id, category: normalizedCategory, applied, changed: Boolean(result.changed), verified: true };
  }

  async archive(id: string, confirmed: boolean, requestId?: string | null): Promise<{
    id: string; archived: true; verified: true; sessionInvalidated: true; requestId: string; deduplicated: boolean;
  }> {
    const moved = await this.move(id, '存档', confirmed, requestId);
    return {
      id: moved.id, archived: true, verified: true, sessionInvalidated: true,
      requestId: moved.requestId, deduplicated: moved.deduplicated,
    };
  }

  async conversation(id: string): Promise<ConversationResult> {
    validateMailId(id);
    const locator = await this.sessionLocator(id);
    await this.ensureReady();
    const result = await this.backend.getConversation(locator);
    if (result.matchCount > 1) throw new AppError('AMBIGUOUS_MESSAGE', `邮件 ID ${id} 对应多个候选。`);
    if (result.matchCount !== 1 || result.messages.length === 0) throw new AppError('MESSAGE_NOT_FOUND', `无法读取邮件 ID ${id} 的会话。`);
    return {
      id, stableId: locator.stableId ?? stableMessageId(locator), subject: locator.subject, complete: result.complete,
      messages: result.messages.map((message, index) => ({
        index: index + 1,
        subject: message.subject,
        from: { name: message.fromName, address: message.fromAddress },
        to: message.to, cc: message.cc,
        receivedAt: message.receivedAt, receivedAtText: message.receivedAtText,
        bodyText: message.bodyText,
        bodyTruncated: Boolean(message.bodyTruncated),
        bodyBytes: message.bodyBytes ?? Buffer.byteLength(message.bodyText, 'utf8'),
        attachments: message.attachments.map((attachment, attachmentIndex) => ({ id: String(attachmentIndex + 1), ...attachment })),
      })),
    };
  }

  async downloadAttachment(
    id: string,
    attachmentId: string,
    outputDirectory: string,
  ): Promise<AttachmentDownloadResult> {
    validateMailId(id);
    if (!/^[1-9]\d*$/.test(attachmentId)) throw new AppError('INVALID_ARGUMENT', '附件 ID 必须是 attachments/read 返回的正整数。');
    const directory = resolve(outputDirectory);
    await mkdir(directory, { recursive: true });
    await this.ensureReady();
    const result = await this.backend.downloadAttachment(
      await this.sessionLocator(id),
      Number(attachmentId) - 1,
      directory,
    );
    this.assertPerformed(id, result);
    return result;
  }

  async downloadAll(id: string, outputDirectory: string): Promise<DownloadAllResult> {
    validateMailId(id);
    const directory = resolve(outputDirectory);
    await mkdir(directory, { recursive: true });
    const stagingDirectory = await mkdtemp(join(directory, '.webmail-download-'));
    try {
      const message = await this.read(id);
      const attachments = [];
      for (const attachment of message.attachments) {
        const downloaded = await this.downloadAttachment(id, attachment.id, stagingDirectory);
        if (!downloaded.path || !downloaded.filename || downloaded.bytes === undefined) {
          throw new AppError('OPERATION_FAILED', `附件 ${attachment.id} 下载结果缺少文件信息。`);
        }
        const saved = await copyWithoutOverwrite(downloaded.path, directory, downloaded.filename);
        attachments.push({
          id: attachment.id,
          filename: saved.filename,
          path: saved.path,
          bytes: downloaded.bytes,
          sha256: await sha256File(saved.path),
        });
      }
      return { id, outputDirectory: directory, attachments };
    } finally {
      await rm(stagingDirectory, { recursive: true, force: true });
    }
  }

  async exportObsidian(id: string, outputDirectory: string): Promise<ObsidianExportResult> {
    if (!outputDirectory.normalize('NFKC').trim()) {
      throw new AppError('INVALID_ARGUMENT', '导出目录不能为空。');
    }

    const message = await this.read(id);
    const paths = await chooseExportPaths(message, outputDirectory);
    if (paths.attachmentDirectory) await mkdir(paths.attachmentDirectory, { recursive: true });

    const attachments = [];
    for (const attachment of message.attachments) {
      const downloaded = await this.downloadAttachment(id, attachment.id, paths.attachmentDirectory!);
      if (!downloaded.path || !downloaded.filename || downloaded.bytes === undefined) {
        throw new AppError('OPERATION_FAILED', `附件 ${attachment.id} 已下载，但返回结果缺少文件信息。`);
      }
      attachments.push({
        id: attachment.id,
        filename: downloaded.filename,
        path: downloaded.path,
        bytes: downloaded.bytes,
        link: attachmentLink(paths.markdownPath, downloaded.path),
      });
    }

    const markdown = renderObsidianMarkdown(message, attachments);
    const bytes = await writeMarkdownAtomically(paths.markdownPath, markdown);
    return {
      id,
      markdownPath: paths.markdownPath,
      attachmentDirectory: paths.attachmentDirectory,
      attachments,
      bytes,
    };
  }

  async exportMessage(id: string, outputDirectory: string, format: MailExportFormat = 'md'): Promise<MailExportResult> {
    if (format === 'md') return { ...await this.exportObsidian(id, outputDirectory), format: 'md' };
    if (format !== 'eml') throw new AppError('INVALID_ARGUMENT', 'format 必须是 md 或 eml。');
    if (!outputDirectory.normalize('NFKC').trim()) {
      throw new AppError('INVALID_ARGUMENT', '导出目录不能为空。');
    }

    validateMailId(id);
    const directory = resolve(outputDirectory);
    await mkdir(directory, { recursive: true });
    const stagingDirectory = await mkdtemp(join(directory, '.webmail-eml-'));
    try {
      await this.ensureReady();
      const downloaded = await this.backend.downloadMessageAsEml(await this.sessionLocator(id), stagingDirectory);
      this.assertPerformed(id, downloaded);
      if (!downloaded.path || !downloaded.filename || downloaded.bytes === undefined) {
        throw new AppError('OPERATION_FAILED', 'Outlook EML 下载结果缺少文件信息。');
      }
      const preferredFilename = downloaded.filename.toLowerCase().endsWith('.eml')
        ? downloaded.filename
        : `${downloaded.filename}.eml`;
      const saved = await copyWithoutOverwrite(downloaded.path, directory, preferredFilename);
      return {
        id,
        format: 'eml',
        filename: saved.filename,
        emlPath: saved.path,
        attachmentCount: downloaded.attachmentCount ?? 0,
        bytes: downloaded.bytes,
      };
    } finally {
      await rm(stagingDirectory, { recursive: true, force: true });
    }
  }

  async exportBatch(options: DatedMailListOptions, outputDirectory: string): Promise<{
    outputDirectory: string;
    fromDate: string;
    toDate: string;
    directory: FolderSummary | null;
    exports: ObsidianExportResult[];
  }> {
    if (!outputDirectory.normalize('NFKC').trim()) throw new AppError('INVALID_ARGUMENT', '导出目录不能为空。');
    const directory = resolve(outputDirectory);
    await mkdir(directory, { recursive: true });
    const exports: ObsidianExportResult[] = [];
    let cursor: string | null = null;
    let first: DatedMailListResult | null = null;
    for (let page = 0; page < 100; page += 1) {
      const result = await this.listByDate({ ...options, limit: options.limit ?? 20, cursor });
      first ??= result;
      for (const message of result.messages) exports.push(await this.exportObsidian(message.stableId, directory));
      cursor = result.nextCursor;
      if (!cursor) {
        return {
          outputDirectory: directory,
          fromDate: first.fromDate,
          toDate: first.toDate,
          directory: first.directory,
          exports,
        };
      }
    }
    throw new AppError('OPERATION_FAILED', '批量导出超过 100 页，已停止以避免无限循环。');
  }

  async syncObsidian(options: DatedMailListOptions, outputDirectory: string): Promise<ObsidianSyncResult> {
    return await syncObsidian(this, options, outputDirectory);
  }
}
