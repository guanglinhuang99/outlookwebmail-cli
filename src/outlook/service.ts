import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';
import type { BrowserBackend } from '../browser/backend.js';
import type {
  AttachmentDownloadResult,
  AttachmentSummary,
  FolderSummary,
  MailMessage,
  MailSummary,
  MessageActionResult,
  MessageLocator,
  RawMessageRow,
} from '../types/mail.js';
import type { InspectResult, MessageInspectResult, StatusResult } from '../types/inspect.js';
import { AppError } from '../util/errors.js';
import { messageFingerprint } from '../util/text.js';
import { EgoInboxParser, type InboxParser } from './inbox-parser.js';
import { EgoMessageParser, type MessageParser } from './message-parser.js';
import { FolderParser } from './folder-parser.js';
import { SessionStore } from '../session/session-store.js';
import { detectOutlookState } from './state.js';

const SEARCH_SELECTOR = 'input[role="combobox"][aria-label^="搜索"], input[role="combobox"][aria-label^="Search"]';
const EXIT_SEARCH_SELECTOR = 'button[aria-label="退出搜索"], button[aria-label="Exit search"]';
const SEARCH_STATE_SCRIPT = String.raw`
(() => {
  const inputs = Array.from(document.querySelectorAll('input[role="combobox"]'))
    .filter(el => /搜索|search/i.test((el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('placeholder') || '')));
  return { count: inputs.length, value: inputs.length === 1 ? inputs[0].value : null };
})()
`;

interface SearchState {
  count: number;
  value: string | null;
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

function requireReadyPage(page: { url?: string; dialog?: unknown }, snapshot: string): string {
  if (page.dialog) {
    throw new AppError('OUTLOOK_NOT_READY', 'Outlook 页面存在未处理的浏览器对话框。');
  }

  const url = page.url ?? null;
  const state = detectOutlookState(url, snapshot);
  if (state === 'AUTH_REQUIRED') {
    throw new AppError('AUTH_REQUIRED', '请在 Ego Lite 中重新登录 Outlook，然后再次执行命令。');
  }

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
  ) {
    this.inboxParser = inboxParser ?? new EgoInboxParser(backend);
    this.messageParser = messageParser ?? new EgoMessageParser(backend);
    this.folderParser = new FolderParser(backend);
  }

  private async ensureReady(): Promise<void> {
    const observation = await this.backend.status();
    requireReadyPage(observation.page, observation.snapshot);
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
    await this.backend.clickAndWait(EXIT_SEARCH_SELECTOR, 1_000);
    const clearedState = await this.searchState();
    if (clearedState.value) {
      throw new AppError('UI_CHANGED', '已点击 Outlook 的退出搜索按钮，但搜索状态仍未清除。');
    }
  }

  private async collectRows(target: number, resetScroll: boolean): Promise<RawMessageRow[]> {
    const rowsByFingerprint = new Map<string, RawMessageRow>();
    let rows = resetScroll ? await this.inboxParser.resetAndExtract() : await this.inboxParser.extract();
    for (let attempt = 0; attempt <= 5; attempt += 1) {
      for (const row of rows) {
        const fingerprint = messageFingerprint(row);
        if (!rowsByFingerprint.has(fingerprint)) rowsByFingerprint.set(fingerprint, row);
      }

      if (rowsByFingerprint.size >= target || attempt === 5) break;
      rows = await this.inboxParser.scrollAndExtract();
    }

    return Array.from(rowsByFingerprint.values()).slice(0, 100);
  }

  private async list(source: string, options: InboxOptions, directory: FolderSummary | null): Promise<MailListResult> {
    if (!Number.isInteger(options.limit) || options.limit < 1 || options.limit > 100) {
      throw new AppError('INVALID_ARGUMENT', '--limit 必须是 1 到 100 之间的整数。');
    }

    const collectionTarget = options.unreadOnly ? 100 : options.limit;
    const rawRows = await this.collectRows(collectionTarget, true);
    const selectedRows = (options.unreadOnly ? rawRows.filter(row => row.unread === true) : rawRows)
      .slice(0, options.limit);
    const messages: MailSummary[] = selectedRows.map((row, index) => ({
      id: String(index + 1),
      sender: { name: row.senderName, address: row.senderAddress },
      subject: row.subject,
      receivedAt: row.receivedAt,
      receivedAtText: row.receivedAtText,
      preview: row.preview,
      unread: row.unread,
      hasAttachments: row.hasAttachments,
    }));

    await this.sessionStore.write({
      version: 1,
      updatedAt: new Date().toISOString(),
      source,
      messages: Object.fromEntries(messages.map(message => [message.id, {
        subject: message.subject,
        senderName: message.sender.name,
        senderAddress: message.sender.address,
        receivedAt: message.receivedAt,
        receivedAtText: message.receivedAtText,
        preview: message.preview,
        hasAttachments: message.hasAttachments,
      }])),
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
    const url = requireReadyPage(observation.page, observation.snapshot);
    return {
      backend: 'ego-lite',
      url,
      title: observation.title,
      state: detectOutlookState(url, observation.snapshot),
    };
  }

  async inspect(): Promise<InspectResult> {
    const observation = await this.backend.inspect();
    const url = requireReadyPage(observation.page, observation.snapshot);
    return {
      backend: 'ego-lite',
      capturedAt: new Date().toISOString(),
      state: detectOutlookState(url, observation.snapshot),
      ...observation,
    };
  }

  async inspectMessage(): Promise<MessageInspectResult> {
    const observation = await this.backend.inspectMessage();
    const url = requireReadyPage(observation.page, observation.snapshot);
    return {
      backend: 'ego-lite',
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
    await this.searchState();
    await this.backend.fillAndPress(SEARCH_SELECTOR, normalizedQuery, 'Enter', 1_500);
    return await this.list(`search:${normalizedQuery}`, options, null);
  }

  async today(directoryName?: string | null): Promise<MailListResult> {
    await this.ensureReady();
    await this.clearSearch();
    const directory = await this.selectDirectory(directoryName);
    const today = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
    }).format(new Date());
    const rowsByFingerprint = new Map<string, RawMessageRow>();
    let rows = await this.inboxParser.resetAndExtract();
    let previousPage = '';
    let complete = false;
    for (let attempt = 0; attempt < 100; attempt += 1) {
      for (const row of rows) {
        if (row.receivedAt?.slice(0, 10) === today) rowsByFingerprint.set(messageFingerprint(row), row);
      }
      if (rows.some(row => row.receivedAt && row.receivedAt.slice(0, 10) < today)) {
        complete = true;
        break;
      }
      const page = rows.map(messageFingerprint).join('\u001e');
      if (!rows.length || page === previousPage) {
        complete = true;
        break;
      }
      previousPage = page;
      rows = await this.inboxParser.scrollAndExtract();
    }
    if (!complete) {
      throw new AppError('OPERATION_FAILED', '今日邮件超过 100 个虚拟滚动页，已拒绝返回可能不完整的结果。');
    }

    const messages: MailSummary[] = Array.from(rowsByFingerprint.values()).map((row, index) => ({
      id: String(index + 1),
      sender: { name: row.senderName, address: row.senderAddress },
      subject: row.subject,
      receivedAt: row.receivedAt,
      receivedAtText: row.receivedAtText,
      preview: row.preview,
      unread: row.unread,
      hasAttachments: row.hasAttachments,
    }));
    await this.sessionStore.write({
      version: 1,
      updatedAt: new Date().toISOString(),
      source: `today:${directory.path}:${today}`,
      messages: Object.fromEntries(messages.map(message => [message.id, {
        subject: message.subject,
        senderName: message.sender.name,
        senderAddress: message.sender.address,
        receivedAt: message.receivedAt,
        receivedAtText: message.receivedAtText,
        preview: message.preview,
        hasAttachments: message.hasAttachments,
      }])),
    });
    return { directory, messages };
  }

  async folders(): Promise<{ root: '收件箱'; folders: FolderSummary[] }> {
    await this.ensureReady();
    return { root: '收件箱', folders: await this.folderParser.extract() };
  }

  private async sessionLocator(id: string): Promise<MessageLocator> {
    const session = await this.sessionStore.read();
    if (!session) {
      throw new AppError('INVALID_ARGUMENT', '当前没有邮件列表 Session；请先运行 inbox 或 search。');
    }
    const message = session.messages[id];
    if (!message) {
      throw new AppError('INVALID_ARGUMENT', `Session 中不存在邮件 ID ${id}；请重新运行 inbox 或 search。`);
    }
    return message;
  }

  async read(id: string): Promise<MailMessage> {
    if (!/^\d+$/.test(id)) throw new AppError('INVALID_ARGUMENT', '邮件 ID 必须是 inbox/search 返回的数字短 ID。');
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
      subject: result.message.subject,
      from: { name: result.message.fromName, address: result.message.fromAddress },
      to: result.message.to,
      cc: result.message.cc,
      receivedAt: result.message.receivedAt,
      receivedAtText: result.message.receivedAtText,
      bodyText: result.message.bodyText,
      attachments: result.message.attachments.map((attachment, index) => ({
        id: String(index + 1),
        filename: attachment.filename,
        sizeText: attachment.sizeText,
      })),
    };
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
      throw new AppError('OPERATION_FAILED', 'Outlook 已接受操作，但无法验证最终页面状态。');
    }
  }

  async delete(id: string, confirmed: boolean): Promise<{ id: string; deleted: true; verified: true; sessionInvalidated: true }> {
    if (!confirmed) {
      throw new AppError('CONFIRMATION_REQUIRED', '删除邮件必须显式提供 --yes；邮件将移入“已删除邮件”。');
    }
    if (!/^\d+$/.test(id)) throw new AppError('INVALID_ARGUMENT', '邮件 ID 必须是 inbox/search/today 返回的数字短 ID。');
    await this.ensureReady();
    const result = await this.backend.deleteMessage(await this.sessionLocator(id));
    this.assertPerformed(id, result);
    await this.sessionStore.clear();
    return { id, deleted: true, verified: true, sessionInvalidated: true };
  }

  async move(
    id: string,
    folder: string,
    confirmed: boolean,
  ): Promise<{ id: string; folder: string; moved: true; verified: true; sessionInvalidated: true }> {
    if (!confirmed) throw new AppError('CONFIRMATION_REQUIRED', '移动邮件必须显式提供 --yes。');
    if (!/^\d+$/.test(id)) throw new AppError('INVALID_ARGUMENT', '邮件 ID 必须是 inbox/search/today 返回的数字短 ID。');
    const normalizedFolder = folder.normalize('NFKC').trim();
    if (!normalizedFolder) throw new AppError('INVALID_ARGUMENT', '目标目录不能为空。');
    await this.ensureReady();
    const result = await this.backend.moveMessage(await this.sessionLocator(id), normalizedFolder);
    this.assertPerformed(id, result);
    await this.sessionStore.clear();
    return { id, folder: normalizedFolder, moved: true, verified: true, sessionInvalidated: true };
  }

  async downloadAttachment(
    id: string,
    attachmentId: string,
    outputDirectory: string,
  ): Promise<AttachmentDownloadResult> {
    if (!/^\d+$/.test(id)) throw new AppError('INVALID_ARGUMENT', '邮件 ID 必须是 inbox/search/today 返回的数字短 ID。');
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
}
