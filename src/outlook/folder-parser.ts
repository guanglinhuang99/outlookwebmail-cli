import { z } from 'zod';
import type { BrowserBackend } from '../browser/backend.js';
import type { FolderSummary } from '../types/mail.js';
import { AppError } from '../util/errors.js';

const folderSchema = z.object({
  name: z.string().min(1),
  path: z.string().min(1),
  level: z.number().int().positive(),
  expanded: z.boolean().nullable(),
});

const extractionSchema = z.object({
  accountCount: z.number().int().nonnegative(),
  inboxCount: z.number().int().nonnegative(),
  complete: z.boolean(),
  folders: z.array(folderSchema),
});

export const INBOX_FOLDER_SCAN_SCRIPT = String.raw`
(() => {
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const rows = Array.from(document.querySelectorAll('[role="treeitem"]')).map((item, index) => ({
    item,
    index,
    name: clean(Array.from(item.children).find(el => el.tagName === 'SPAN' && clean(el.textContent))?.textContent),
    level: Number(item.getAttribute('aria-level')) || 1,
    expanded: item.hasAttribute('aria-expanded') ? item.getAttribute('aria-expanded') === 'true' : null
  })).filter(row => row.name);
  const accountRows = rows.filter(row => row.level === 1 && /@/.test(row.name));
  if (accountRows.length !== 1) {
    return { accountCount: accountRows.length, inboxCount: 0, folders: [], collapsed: null, hasCollapsed: false };
  }
  const accountIndex = rows.indexOf(accountRows[0]);
  const accountEnd = rows.findIndex((row, index) => index > accountIndex && row.level === 1);
  const mailboxRows = rows.slice(accountIndex, accountEnd < 0 ? rows.length : accountEnd);
  const inboxRows = mailboxRows.filter(row => row.level === 2 && /^(收件箱|inbox)$/i.test(row.name));
  if (inboxRows.length !== 1) {
    return { accountCount: 1, inboxCount: inboxRows.length, folders: [], collapsed: null, hasCollapsed: false };
  }
  const inboxIndex = mailboxRows.indexOf(inboxRows[0]);
  const inboxLevel = inboxRows[0].level;
  const inboxEnd = mailboxRows.findIndex((row, index) => index > inboxIndex && row.level <= inboxLevel);
  const subtree = mailboxRows.slice(inboxIndex + 1, inboxEnd < 0 ? mailboxRows.length : inboxEnd);
  const stack = [inboxRows[0].name];
  const folders = subtree.map(row => {
    stack.length = Math.max(1, row.level - inboxLevel);
    const path = [...stack, row.name].join('/');
    stack[row.level - inboxLevel] = row.name;
    return { name: row.name, path, level: row.level, expanded: row.expanded, item: row.item };
  });
  const collapsed = folders.find(folder => folder.expanded === false);
  let collapsedTarget = null;
  if (collapsed) {
    const button = collapsed.item.querySelector('button');
    if (button) {
      button.scrollIntoView({ block: 'center', inline: 'nearest' });
      const rect = button.getBoundingClientRect();
      collapsedTarget = {
        path: collapsed.path,
        rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
      };
    }
  }
  return {
    accountCount: 1,
    inboxCount: 1,
    folders: folders.map(({ item, ...folder }) => folder),
    collapsed: collapsedTarget,
    hasCollapsed: Boolean(collapsed)
  };
})()
`;

export function buildInboxFolderTargetScript(directory: string | null): string {
  return String.raw`
(() => {
  const requested = ${JSON.stringify(directory)};
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const normalized = value => clean(value).toLowerCase();
  const rows = Array.from(document.querySelectorAll('[role="treeitem"]')).map(item => ({
    item,
    name: clean(Array.from(item.children).find(el => el.tagName === 'SPAN' && clean(el.textContent))?.textContent),
    level: Number(item.getAttribute('aria-level')) || 1,
    expanded: item.hasAttribute('aria-expanded') ? item.getAttribute('aria-expanded') === 'true' : null
  })).filter(row => row.name);
  const accountIndex = rows.findIndex(row => row.level === 1 && /@/.test(row.name));
  if (accountIndex < 0) return { count: 0, folder: null, selected: false, rect: null };
  const accountEnd = rows.findIndex((row, index) => index > accountIndex && row.level === 1);
  const mailboxRows = rows.slice(accountIndex, accountEnd < 0 ? rows.length : accountEnd);
  const inboxIndex = mailboxRows.findIndex(row => row.level === 2 && /^(收件箱|inbox)$/i.test(row.name));
  if (inboxIndex < 0) return { count: 0, folder: null, selected: false, rect: null };
  const inbox = mailboxRows[inboxIndex];
  const inboxEnd = mailboxRows.findIndex((row, index) => index > inboxIndex && row.level <= inbox.level);
  const subtree = mailboxRows.slice(inboxIndex + 1, inboxEnd < 0 ? mailboxRows.length : inboxEnd);
  const stack = [inbox.name];
  const folders = subtree.map(row => {
    stack.length = Math.max(1, row.level - inbox.level);
    const path = [...stack, row.name].join('/');
    stack[row.level - inbox.level] = row.name;
    return { ...row, path };
  });
  let matches;
  if (!requested) {
    matches = [{ ...inbox, path: inbox.name }];
  } else {
    const pathMatches = folders.filter(folder => normalized(folder.path) === normalized(requested));
    matches = pathMatches.length > 0
      ? pathMatches
      : folders.filter(folder => normalized(folder.name) === normalized(requested));
  }
  if (matches.length !== 1) return { count: matches.length, folder: null, selected: false, rect: null };
  const match = matches[0];
  match.item.scrollIntoView({ block: 'center', inline: 'nearest' });
  const rect = match.item.getBoundingClientRect();
  const selected = Array.from(match.item.children).some(el => /已选择|selected/i.test(clean(el.textContent)));
  return {
    count: 1,
    folder: { name: match.name, path: match.path, level: match.level, expanded: match.expanded },
    selected,
    rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
  };
})()
`;
}

export class FolderParser {
  constructor(private readonly backend: BrowserBackend) {}

  async extract(): Promise<FolderSummary[]> {
    const result = extractionSchema.parse(await this.backend.listInboxFolders());
    if (result.accountCount !== 1 || result.inboxCount !== 1) {
      throw new AppError(
        'UI_CHANGED',
        `预期找到 1 个主邮箱和 1 个 Inbox，实际为 ${result.accountCount} 和 ${result.inboxCount}。`,
      );
    }
    if (!result.complete) {
      throw new AppError('OPERATION_FAILED', 'Inbox 子目录未能完整展开。');
    }
    return result.folders;
  }
}
