import { mkdtemp, readFile, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import {
  chooseExportPaths,
  exportStem,
  renderObsidianMarkdown,
  writeMarkdownAtomically,
} from '../src/export/obsidian.js';
import type { MailMessage } from '../src/types/mail.js';

const temporaryDirectories: string[] = [];

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map(path => rm(path, { recursive: true, force: true })));
});

const message: MailMessage = {
  id: '1',
  subject: '风险/报告: 终稿',
  from: { name: '张三', address: 'zhangsan@example.com' },
  to: [{ name: '李四', address: 'lisi@example.com' }],
  cc: [],
  receivedAt: '2026-08-20T09:30:00+08:00',
  receivedAtText: '2026/8/20 9:30',
  bodyText: '第一行\n\n第二行',
  attachments: [{ id: '1', filename: '风险 报告[终].xlsx', sizeText: '12 KB' }],
};

describe('Obsidian export formatting', () => {
  it('creates a filesystem-safe and stable export stem', () => {
    const stem = exportStem(message);

    expect(stem).toMatch(/^2026-08-20_0930-风险-报告- 终稿-[a-f0-9]{8}$/);
    expect(exportStem(message)).toBe(stem);
  });

  it('renders Obsidian frontmatter, body and a relative attachment link', () => {
    const markdown = renderObsidianMarkdown(message, [{
      id: '1',
      filename: '风险 报告[终].xlsx',
      path: '/vault/mail/attachments/item/风险 报告[终].xlsx',
      bytes: 12,
      link: 'attachments/item/%E9%A3%8E%E9%99%A9%20%E6%8A%A5%E5%91%8A%5B%E7%BB%88%5D.xlsx',
    }]);

    expect(markdown).toContain('title: "风险/报告: 终稿"');
    expect(markdown).toContain('from: "张三 <zhangsan@example.com>"');
    expect(markdown).toContain('to: ["李四 <lisi@example.com>"]');
    expect(markdown).toContain('- [风险 报告\\[终\\].xlsx](attachments/item/');
    expect(markdown).toContain('## 正文\n\n第一行\n\n第二行');
  });

  it('does not overwrite an existing Markdown export', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'webmail-obsidian-test-'));
    temporaryDirectories.push(directory);
    const first = await chooseExportPaths({ ...message, attachments: [] }, directory);
    await writeMarkdownAtomically(first.markdownPath, 'first');
    const second = await chooseExportPaths({ ...message, attachments: [] }, directory);
    await writeMarkdownAtomically(second.markdownPath, 'second');

    expect(second.stem).toBe(`${first.stem}-2`);
    await expect(readFile(first.markdownPath, 'utf8')).resolves.toBe('first');
    await expect(readFile(second.markdownPath, 'utf8')).resolves.toBe('second');
    await expect(stat(join(directory, 'attachments'))).rejects.toMatchObject({ code: 'ENOENT' });
  });
});
