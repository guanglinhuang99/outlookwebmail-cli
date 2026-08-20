import { createHash, randomUUID } from 'node:crypto';
import { access, mkdir, rename, writeFile } from 'node:fs/promises';
import { basename, join, relative, resolve, sep } from 'node:path';
import type { ExportedAttachment, MailAddress, MailMessage } from '../types/mail.js';
import { AppError } from '../util/errors.js';

export interface ObsidianExportPaths {
  outputDirectory: string;
  stem: string;
  markdownPath: string;
  attachmentDirectory: string | null;
}

function addressText(address: MailAddress): string {
  if (address.name && address.address) return `${address.name} <${address.address}>`;
  return address.name ?? address.address ?? '';
}

function frontmatterValue(value: string): string {
  return JSON.stringify(value);
}

function timestampPrefix(receivedAt: string | null): string {
  if (!receivedAt) return 'unknown-date';
  const match = receivedAt.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})/);
  return match ? `${match[1]}-${match[2]}-${match[3]}_${match[4]}${match[5]}` : 'unknown-date';
}

function safeSubject(subject: string): string {
  const normalized = subject
    .normalize('NFKC')
    .replace(/[\u0000-\u001f\u007f]/g, ' ')
    .replace(/[\\/:*?"<>|#[\]]/g, '-')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '')
    .trim();
  return (normalized || 'untitled').slice(0, 80).replace(/[. ]+$/g, '') || 'untitled';
}

export function exportStem(message: MailMessage): string {
  const hash = createHash('sha256')
    .update([message.subject, addressText(message.from), message.receivedAt ?? ''].join('\u001f'))
    .digest('hex')
    .slice(0, 8);
  return `${timestampPrefix(message.receivedAt)}-${safeSubject(message.subject)}-${hash}`;
}

async function exists(path: string): Promise<boolean> {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

export async function chooseExportPaths(
  message: MailMessage,
  outputDirectory: string,
): Promise<ObsidianExportPaths> {
  const directory = resolve(outputDirectory);
  await mkdir(directory, { recursive: true });
  const baseStem = exportStem(message);

  for (let index = 1; index <= 10_000; index += 1) {
    const stem = index === 1 ? baseStem : `${baseStem}-${index}`;
    const markdownPath = join(directory, `${stem}.md`);
    const attachmentDirectory = message.attachments.length ? join(directory, 'attachments', stem) : null;
    if (!await exists(markdownPath) && (!attachmentDirectory || !await exists(attachmentDirectory))) {
      return { outputDirectory: directory, stem, markdownPath, attachmentDirectory };
    }
  }

  throw new AppError('OPERATION_FAILED', '导出目录中的同名邮件版本过多，无法分配新文件名。');
}

function markdownLinkLabel(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/\[/g, '\\[').replace(/\]/g, '\\]');
}

function encodeRelativePath(path: string): string {
  return path.split(sep).map(segment => encodeURIComponent(segment)).join('/');
}

export function attachmentLink(markdownPath: string, attachmentPath: string): string {
  const relativePath = relative(resolve(markdownPath, '..'), resolve(attachmentPath));
  if (!relativePath || relativePath === '..' || relativePath.startsWith(`..${sep}`)) {
    throw new AppError('OPERATION_FAILED', '附件不在 Markdown 导出目录内，无法生成安全的相对链接。');
  }
  return encodeRelativePath(relativePath);
}

export function renderObsidianMarkdown(message: MailMessage, attachments: ExportedAttachment[]): string {
  const from = addressText(message.from);
  const to = message.to.map(addressText).filter(Boolean);
  const cc = message.cc.map(addressText).filter(Boolean);
  const receivedAt = message.receivedAt ?? message.receivedAtText ?? '';
  const attachmentLines = attachments.length
    ? attachments.map(attachment => `- [${markdownLinkLabel(attachment.filename)}](${attachment.link})`).join('\n')
    : '- 无';

  return [
    '---',
    'type: email',
    `title: ${frontmatterValue(message.subject)}`,
    `received_at: ${frontmatterValue(receivedAt)}`,
    `from: ${frontmatterValue(from)}`,
    `to: ${JSON.stringify(to)}`,
    `cc: ${JSON.stringify(cc)}`,
    'tags: [email]',
    '---',
    '',
    `# ${message.subject}`,
    '',
    '> [!info] 邮件信息',
    `> 发件人：${from || '未知'}`,
    `> 收件人：${to.join('；') || '未解析'}`,
    `> 抄送：${cc.join('；') || '无'}`,
    `> 时间：${message.receivedAtText ?? message.receivedAt ?? '未知'}`,
    '',
    '## 附件',
    '',
    attachmentLines,
    '',
    '## 正文',
    '',
    message.bodyText,
    '',
  ].join('\n');
}

export async function writeMarkdownAtomically(path: string, markdown: string): Promise<number> {
  const temporaryPath = join(resolve(path, '..'), `.${basename(path)}.${randomUUID()}.tmp`);
  await writeFile(temporaryPath, markdown, { encoding: 'utf8', flag: 'wx' });
  await rename(temporaryPath, path);
  return Buffer.byteLength(markdown, 'utf8');
}
