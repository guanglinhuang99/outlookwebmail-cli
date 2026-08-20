import { createHash, randomUUID } from 'node:crypto';
import { constants } from 'node:fs';
import {
  access,
  copyFile,
  mkdir,
  mkdtemp,
  open,
  readFile,
  rename,
  rm,
  stat,
  unlink,
  writeFile,
} from 'node:fs/promises';
import { basename, extname, join, relative, resolve, sep } from 'node:path';
import type { DatedMailListOptions, DatedMailListResult } from '../outlook/service.js';
import type {
  AttachmentDownloadResult,
  DownloadedAttachmentWithHash,
  MailMessage,
  ObsidianSyncItem,
  ObsidianSyncResult,
} from '../types/mail.js';
import { attachmentLink, exportStem, renderObsidianMarkdown, replaceFileAtomically, writeMarkdownAtomically } from '../export/obsidian.js';
import { AppError } from '../util/errors.js';

const MANIFEST_NAME = '.webmail-cli-index.json';
const ATTACHMENT_INDEX_NAME = '_attachments-index.md';
const MANAGED_END = '<!-- webmail-cli:managed-end -->';
const IMAGE_EXTENSIONS = new Set(['.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp']);

interface ManifestAttachment {
  id: string;
  filename: string;
  path: string;
  bytes: number;
  sha256: string;
}

interface ManifestEntry {
  markdownPath: string;
  contentHash: string;
  subject: string;
  receivedAt: string | null;
  updatedAt: string;
  attachments: ManifestAttachment[];
}

interface SyncManifest {
  version: 1;
  updatedAt: string;
  messages: Record<string, ManifestEntry>;
}

export interface ObsidianSyncSource {
  listByDate(options: DatedMailListOptions): Promise<DatedMailListResult>;
  read(id: string): Promise<MailMessage>;
  downloadAttachment(id: string, attachmentId: string, outputDirectory: string): Promise<AttachmentDownloadResult>;
}

async function exists(path: string): Promise<boolean> {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

function safeFilename(value: string): string {
  const normalized = basename(value).normalize('NFKC')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
    .trim();
  return (normalized || 'attachment').slice(0, 120) || 'attachment';
}

function messageHash(message: MailMessage): string {
  return createHash('sha256').update(JSON.stringify({
    stableId: message.stableId,
    subject: message.subject,
    from: message.from,
    to: message.to,
    cc: message.cc,
    receivedAt: message.receivedAt,
    receivedAtText: message.receivedAtText,
    bodyText: message.bodyText,
    attachments: message.attachments,
  })).digest('hex');
}

async function sha256File(path: string): Promise<string> {
  const buffer = await readFile(path);
  return createHash('sha256').update(buffer).digest('hex');
}

function relativeManifestPath(root: string, path: string): string {
  return relative(root, path).split(sep).join('/');
}

function resolveManifestPath(root: string, path: string): string {
  const target = resolve(root, ...path.split('/'));
  if (target !== root && !target.startsWith(`${root}${sep}`)) {
    throw new AppError('OPERATION_FAILED', 'Obsidian 索引包含越出同步目录的路径。');
  }
  return target;
}

async function readManifest(path: string): Promise<SyncManifest> {
  try {
    const parsed = JSON.parse(await readFile(path, 'utf8')) as Partial<SyncManifest>;
    if (parsed.version !== 1 || !parsed.messages || typeof parsed.messages !== 'object') throw new Error('invalid manifest');
    return parsed as SyncManifest;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'ENOENT') {
      return { version: 1, updatedAt: new Date(0).toISOString(), messages: {} };
    }
    throw new AppError('OPERATION_FAILED', `无法读取 Obsidian 同步索引：${path}`, { cause: error });
  }
}

async function writeJsonAtomically(path: string, value: unknown): Promise<void> {
  const temporary = join(resolve(path, '..'), `.${basename(path)}.${randomUUID()}.tmp`);
  await writeFile(temporary, `${JSON.stringify(value, null, 2)}\n`, { encoding: 'utf8', flag: 'wx', mode: 0o600 });
  await replaceFileAtomically(temporary, path);
}

async function lockManifest(path: string): Promise<() => Promise<void>> {
  const lockPath = `${path}.lock`;
  try {
    const handle = await open(lockPath, 'wx', 0o600);
    await handle.writeFile(`${process.pid}\n${new Date().toISOString()}\n`);
    await handle.close();
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'EEXIST') {
      throw new AppError('OPERATION_FAILED', `已有 Obsidian 同步任务正在运行：${lockPath}`);
    }
    throw error;
  }
  return async () => { await unlink(lockPath).catch(() => undefined); };
}

async function replaceManagedMarkdown(path: string, generated: string): Promise<number> {
  let suffix = '';
  try {
    const existing = await readFile(path, 'utf8');
    const marker = existing.indexOf(MANAGED_END);
    if (marker >= 0) suffix = existing.slice(marker + MANAGED_END.length).replace(/^\n?/, '\n');
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
  }
  const markdown = `${generated.trimEnd()}\n\n${MANAGED_END}${suffix || '\n'}`;
  return await writeMarkdownAtomically(path, markdown);
}

function renderManagedMarkdown(
  message: MailMessage,
  markdownPath: string,
  attachments: DownloadedAttachmentWithHash[],
): string {
  const rendered = attachments.map(attachment => ({
    ...attachment,
    link: attachmentLink(markdownPath, attachment.path),
  }));
  let markdown = renderObsidianMarkdown(message, rendered);
  for (const attachment of rendered) {
    if (!IMAGE_EXTENSIONS.has(extname(attachment.filename).toLowerCase())) continue;
    const ordinary = `- [${attachment.filename.replace(/([\\[\]])/g, '\\$1')}](${attachment.link})`;
    markdown = markdown.replace(ordinary, `${ordinary}\n  ![${attachment.filename.replace(/([\\[\]])/g, '\\$1')}](${attachment.link})`);
  }
  return markdown;
}

async function validExistingEntry(root: string, entry: ManifestEntry, contentHash: string): Promise<boolean> {
  if (entry.contentHash !== contentHash) return false;
  const markdownPath = resolveManifestPath(root, entry.markdownPath);
  if (!await exists(markdownPath)) return false;
  for (const attachment of entry.attachments) {
    const path = resolveManifestPath(root, attachment.path);
    if (!await exists(path)) return false;
    const file = await stat(path);
    if (file.size !== attachment.bytes || await sha256File(path) !== attachment.sha256) return false;
  }
  return true;
}

async function installAttachment(source: string, target: string): Promise<void> {
  await mkdir(resolve(target, '..'), { recursive: true });
  const temporary = join(resolve(target, '..'), `.${basename(target)}.${randomUUID()}.tmp`);
  await copyFile(source, temporary, constants.COPYFILE_EXCL);
  await unlink(target).catch(error => {
    if ((error as NodeJS.ErrnoException).code !== 'ENOENT') throw error;
  });
  await rename(temporary, target);
}

function markdownPathFor(message: MailMessage, root: string): string {
  const stableSuffix = message.stableId.replace(/[^A-Za-z0-9_-]/g, '').slice(-12) || createHash('sha256').update(message.stableId).digest('hex').slice(0, 12);
  return join(root, `${exportStem(message)}-${stableSuffix}.md`);
}

async function syncMessage(
  source: ObsidianSyncSource,
  root: string,
  messageId: string,
  manifest: SyncManifest,
): Promise<ObsidianSyncItem> {
  const message = await source.read(messageId);
  const prior = manifest.messages[message.stableId];
  const hash = messageHash(message);
  const markdownPath = prior ? resolveManifestPath(root, prior.markdownPath) : markdownPathFor(message, root);
  if (prior && await validExistingEntry(root, prior, hash)) {
    return {
      id: messageId,
      stableId: message.stableId,
      status: 'unchanged',
      markdownPath,
      attachmentDirectory: prior.attachments.length ? resolve(root, 'attachments', message.stableId) : null,
      attachments: prior.attachments.map(attachment => ({ ...attachment, path: resolveManifestPath(root, attachment.path) })),
    };
  }

  const attachmentDirectory = message.attachments.length ? join(root, 'attachments', message.stableId) : null;
  const staging = await mkdtemp(join(root, '.webmail-sync-'));
  const attachments: DownloadedAttachmentWithHash[] = [];
  try {
    for (const attachment of message.attachments) {
      const downloaded = await source.downloadAttachment(messageId, attachment.id, staging);
      if (!downloaded.path || downloaded.bytes === undefined) {
        throw new AppError('OPERATION_FAILED', `附件 ${attachment.id} 下载结果缺少文件信息。`);
      }
      const target = join(attachmentDirectory!, `${attachment.id}-${safeFilename(downloaded.filename ?? attachment.filename)}`);
      await installAttachment(downloaded.path, target);
      attachments.push({
        id: attachment.id,
        filename: downloaded.filename ?? attachment.filename,
        path: target,
        bytes: (await stat(target)).size,
        sha256: await sha256File(target),
      });
    }
  } finally {
    await rm(staging, { recursive: true, force: true });
  }

  const currentPaths = new Set(attachments.map(item => resolve(item.path)));
  for (const old of prior?.attachments ?? []) {
    const oldPath = resolveManifestPath(root, old.path);
    if (!currentPaths.has(oldPath)) await unlink(oldPath).catch(() => undefined);
  }
  await replaceManagedMarkdown(markdownPath, renderManagedMarkdown(message, markdownPath, attachments));
  manifest.messages[message.stableId] = {
    markdownPath: relativeManifestPath(root, markdownPath),
    contentHash: hash,
    subject: message.subject,
    receivedAt: message.receivedAt,
    updatedAt: new Date().toISOString(),
    attachments: attachments.map(attachment => ({
      ...attachment,
      path: relativeManifestPath(root, attachment.path),
    })),
  };
  return {
    id: messageId,
    stableId: message.stableId,
    status: prior ? 'updated' : 'created',
    markdownPath,
    attachmentDirectory,
    attachments,
  };
}

function renderAttachmentIndex(root: string, manifest: SyncManifest): string {
  const lines = ['# 邮件附件索引', '', '| 邮件 | 附件 | SHA-256 | 字节 |', '| --- | --- | --- | ---: |'];
  for (const entry of Object.values(manifest.messages).sort((a, b) => (b.receivedAt ?? '').localeCompare(a.receivedAt ?? ''))) {
    const mailLink = encodeURI(entry.markdownPath);
    for (const attachment of entry.attachments) {
      const attachmentLinkValue = encodeURI(attachment.path);
      lines.push(`| [${entry.subject.replace(/\|/g, '\\|')}](${mailLink}) | [${attachment.filename.replace(/\|/g, '\\|')}](${attachmentLinkValue}) | \`${attachment.sha256}\` | ${attachment.bytes} |`);
    }
  }
  if (lines.length === 4) lines.push('| - | 无附件 | - | 0 |');
  return `${lines.join('\n')}\n`;
}

export async function syncObsidian(
  source: ObsidianSyncSource,
  options: DatedMailListOptions,
  outputDirectory: string,
): Promise<ObsidianSyncResult> {
  if (!outputDirectory.normalize('NFKC').trim()) throw new AppError('INVALID_ARGUMENT', '同步目录不能为空。');
  const root = resolve(outputDirectory);
  await mkdir(root, { recursive: true });
  const manifestPath = join(root, MANIFEST_NAME);
  const release = await lockManifest(manifestPath);
  try {
    const manifest = await readManifest(manifestPath);
    const items: ObsidianSyncItem[] = [];
    let cursor: string | null = null;
    let first: DatedMailListResult | null = null;
    for (let page = 0; page < 100; page += 1) {
      const result = await source.listByDate({ ...options, limit: options.limit ?? 20, cursor });
      first ??= result;
      for (const message of result.messages) items.push(await syncMessage(source, root, message.stableId, manifest));
      cursor = result.nextCursor;
      if (!cursor) break;
      if (page === 99) throw new AppError('OPERATION_FAILED', 'Obsidian 同步超过 100 页，已停止。');
    }
    if (!first) throw new AppError('OPERATION_FAILED', 'Obsidian 同步未获得邮件列表。');
    manifest.updatedAt = new Date().toISOString();
    await writeJsonAtomically(manifestPath, manifest);
    const attachmentIndexPath = join(root, ATTACHMENT_INDEX_NAME);
    await writeMarkdownAtomically(attachmentIndexPath, renderAttachmentIndex(root, manifest));
    return {
      outputDirectory: root,
      fromDate: first.fromDate,
      toDate: first.toDate,
      directory: first.directory,
      manifestPath,
      attachmentIndexPath,
      created: items.filter(item => item.status === 'created').length,
      updated: items.filter(item => item.status === 'updated').length,
      unchanged: items.filter(item => item.status === 'unchanged').length,
      items,
    };
  } finally {
    await release();
  }
}
