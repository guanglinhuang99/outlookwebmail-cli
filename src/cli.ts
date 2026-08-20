#!/usr/bin/env node

import { Command, CommanderError, InvalidArgumentError } from 'commander';
import { createBrowserBackend } from './browser/browser-factory.js';
import { OutlookService } from './outlook/service.js';
import { AppError, toAppError } from './util/errors.js';
import { errorResult, successResult, writeJson, writePretty } from './util/output.js';
import { watchMail } from './watch/mail-watcher.js';

function requestedJson(argv: string[]): boolean {
  return argv.includes('--json');
}

function parseLimit(value: string): number {
  const limit = Number(value);
  if (!Number.isInteger(limit) || limit < 1 || limit > 100) {
    throw new InvalidArgumentError('必须是 1 到 100 之间的整数');
  }
  return limit;
}

function parseBoolean(value: string): boolean {
  const normalized = value.normalize('NFKC').trim().toLowerCase();
  if (normalized === 'true') return true;
  if (normalized === 'false') return false;
  throw new InvalidArgumentError('必须是 true 或 false');
}

function parseWatchInterval(value: string): number {
  const interval = Number(value);
  if (!Number.isInteger(interval) || interval < 5 || interval > 3_600) {
    throw new InvalidArgumentError('必须是 5 到 3600 之间的整数秒');
  }
  return interval;
}

function parseIterations(value: string): number {
  const iterations = Number(value);
  if (!Number.isInteger(iterations) || iterations < 0) {
    throw new InvalidArgumentError('必须是大于或等于 0 的整数');
  }
  return iterations;
}

function parseRecipients(value?: string): string[] {
  return (value ?? '').split(/[,;；\n]/).map(item => item.normalize('NFKC').trim()).filter(Boolean);
}

function collectValue(value: string, previous: string[]): string[] {
  return [...previous, value];
}

async function execute<T>(json: boolean, operation: () => Promise<T>): Promise<void> {
  try {
    const data = await operation();
    if (json) writeJson(successResult(data));
    else writePretty(data);
  } catch (error) {
    const appError = toAppError(error);
    process.exitCode = appError.exitCode;
    if (json) writeJson(errorResult(appError));
    else process.stderr.write(`[${appError.code}] ${appError.message}\n`);
  }
}

export function createProgram(service = new OutlookService(createBrowserBackend())): Command {
  const program = new Command();
  program
    .name('webmail')
    .description('Outlook Web CLI through Playwright (Ego Lite fallback)')
    .version('0.3.0')
    .showHelpAfterError()
    .exitOverride();

  program
    .command('status')
    .description('检查浏览器后端与 Outlook 登录状态')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.status());
    });

  program
    .command('doctor')
    .description('检查 Node、浏览器后端、登录状态和 Outlook DOM 兼容性')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { json?: boolean }) => {
      await execute(Boolean(options.json), async () => {
        const result = await service.doctor();
        if (!result.ok) process.exitCode = 6;
        return result;
      });
    });

  program
    .command('inspect')
    .description('采集 Outlook 页面结构，供 DOM parser 校准')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.inspect());
    });

  program
    .command('inbox')
    .description('读取 Outlook 收件箱邮件列表')
    .option('-n, --limit <number>', '最多返回的邮件数', parseLimit, 20)
    .option('--unread', '只返回未读邮件')
    .option('--dir <directory>', 'Inbox 子目录名称或 folders 返回的完整 path')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { limit: number; unread?: boolean; dir?: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.inbox({
        limit: options.limit,
        unreadOnly: Boolean(options.unread),
        directory: options.dir,
      }));
    });

  program
    .command('inspect-message')
    .description('采集当前打开邮件的结构，供 MessageParser 校准')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.inspectMessage());
    });

  program
    .command('search')
    .description('使用 Outlook 原生搜索')
    .argument('<query>', '原样交给 Outlook 的搜索内容')
    .option('-n, --limit <number>', '最多返回的邮件数', parseLimit, 20)
    .option('--json', '输出单一 JSON envelope')
    .action(async (query: string, options: { limit: number; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.search(query, { limit: options.limit }));
    });

  program
    .command('list')
    .description('按日期、目录和筛选条件分页列出邮件')
    .option('--date <date>', '日期（YYYY-MM-DD）；省略或为空时使用今天')
    .option('--from-date <date>', '日期范围起点（YYYY-MM-DD，须与 --to-date 同时使用）')
    .option('--to-date <date>', '日期范围终点（YYYY-MM-DD，须与 --from-date 同时使用）')
    .option('--dir <directory>', 'Inbox 子目录名称或 folders 返回的完整 path；省略或为空时使用 Inbox')
    .option('-n, --limit <number>', '每页最多返回的邮件数', parseLimit, 20)
    .option('--cursor <cursor>', '上一页返回的 nextCursor')
    .option('--sender <text>', '按发件人姓名或地址包含匹配')
    .option('--subject <text>', '按主题包含匹配')
    .option('--unread', '只返回未读邮件')
    .option('--has-attachments', '只返回有附件的邮件')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: {
      date?: string; fromDate?: string; toDate?: string; dir?: string; limit: number; cursor?: string;
      sender?: string; subject?: string; unread?: boolean; hasAttachments?: boolean; json?: boolean;
    }) => {
      await execute(Boolean(options.json), () => service.listByDate({
        date: options.date,
        fromDate: options.fromDate,
        toDate: options.toDate,
        directory: options.dir,
        limit: options.limit,
        cursor: options.cursor,
        sender: options.sender,
        subject: options.subject,
        unread: Boolean(options.unread),
        hasAttachments: Boolean(options.hasAttachments),
      }));
    });

  program
    .command('today')
    .description('列出 Inbox 或指定子目录中今天收到的全部邮件')
    .option('--dir <directory>', 'Inbox 子目录名称或 folders 返回的完整 path')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { dir?: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.today(options.dir));
    });

  program
    .command('folders')
    .description('递归列出 Inbox 下的全部子目录')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.folders());
    });

  program
    .command('read')
    .description('按当前 Session 短 ID 读取唯一邮件')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.read(id));
    });

  program
    .command('attachments')
    .description('列出邮件附件元数据，不下载附件')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.attachments(id));
    });

  program
    .command('compose')
    .description('新建邮件；默认保存草稿并交给用户，显式关闭 draft 时自动发送')
    .option('--to <addresses>', '收件人，多个地址用逗号或分号分隔')
    .option('--cc <addresses>', '抄送，多个地址用逗号或分号分隔')
    .option('--bcc <addresses>', '密件抄送，多个地址用逗号或分号分隔')
    .requiredOption('--subject <text>', '邮件主题')
    .requiredOption('--content <text>', '邮件正文')
    .option('--attach <path>', '附加本地文件，可重复提供', collectValue, [])
    .option('--draft <boolean>', 'true 保存草稿；false 自动发送', parseBoolean, true)
    .option('--request-id <id>', '自动发送时必填的幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: {
      to?: string; cc?: string; bcc?: string; subject: string; content: string;
      attach: string[]; draft: boolean; requestId?: string; json?: boolean;
    }) => {
      await execute(Boolean(options.json), () => service.compose({
        to: parseRecipients(options.to), cc: parseRecipients(options.cc), bcc: parseRecipients(options.bcc),
        subject: options.subject, content: options.content, attachments: options.attach, draft: options.draft,
      }, options.requestId));
    });

  program
    .command('forward')
    .description('转发邮件并保留原附件；默认保存草稿')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .requiredOption('--to <addresses>', '收件人，多个地址用逗号或分号分隔')
    .option('--cc <addresses>', '抄送')
    .option('--bcc <addresses>', '密件抄送')
    .requiredOption('--content <text>', '转发附言')
    .option('--attach <path>', '额外附加本地文件，可重复提供', collectValue, [])
    .option('--draft <boolean>', 'true 保存草稿；false 自动发送', parseBoolean, true)
    .option('--request-id <id>', '自动发送时必填的幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: {
      to: string; cc?: string; bcc?: string; content: string; attach: string[];
      draft: boolean; requestId?: string; json?: boolean;
    }) => {
      await execute(Boolean(options.json), () => service.forward(id, {
        to: parseRecipients(options.to), cc: parseRecipients(options.cc), bcc: parseRecipients(options.bcc),
        content: options.content, attachments: options.attach, draft: options.draft,
      }, options.requestId));
    });

  program
    .command('drafts')
    .description('列出草稿目录中的邮件')
    .option('-n, --limit <number>', '最多返回的草稿数', parseLimit, 20)
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { limit: number; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.drafts({ limit: options.limit }));
    });

  program
    .command('draft-read')
    .description('读取一封草稿的收件人、主题、正文和附件')
    .argument('<id>', 'drafts 返回的数字短 ID 或 stableId')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.readDraft(id));
    });

  program
    .command('draft-update')
    .description('修改并保存草稿；只修改显式提供的字段')
    .argument('<id>', 'drafts 返回的数字短 ID 或 stableId')
    .option('--to <addresses>', '替换收件人；传空字符串可清空')
    .option('--cc <addresses>', '替换抄送')
    .option('--bcc <addresses>', '替换密件抄送')
    .option('--subject <text>', '替换主题')
    .option('--content <text>', '替换正文')
    .option('--attach <path>', '追加本地附件，可重复提供', collectValue, [])
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: {
      to?: string; cc?: string; bcc?: string; subject?: string; content?: string; attach: string[]; json?: boolean;
    }) => {
      await execute(Boolean(options.json), () => service.updateDraft(id, {
        to: options.to === undefined ? undefined : parseRecipients(options.to),
        cc: options.cc === undefined ? undefined : parseRecipients(options.cc),
        bcc: options.bcc === undefined ? undefined : parseRecipients(options.bcc),
        subject: options.subject, content: options.content, attachments: options.attach,
      }));
    });

  program
    .command('draft-send')
    .description('发送现有草稿')
    .argument('<id>', 'drafts 返回的数字短 ID 或 stableId')
    .requiredOption('--request-id <id>', '幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { requestId: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.sendDraft(id, options.requestId));
    });

  program
    .command('draft-discard')
    .description('永久放弃一封草稿')
    .argument('<id>', 'drafts 返回的数字短 ID 或 stableId')
    .option('--yes', '确认放弃草稿')
    .requiredOption('--request-id <id>', '幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { yes?: boolean; requestId: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.discardDraft(id, Boolean(options.yes), options.requestId));
    });

  program
    .command('mark-read')
    .description('将邮件标记为已读')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.markRead(id, false));
    });

  program
    .command('mark-unread')
    .description('将邮件标记为未读')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.markRead(id, true));
    });

  program
    .command('flag')
    .description('设置或清除邮件旗标')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--state <boolean>', 'true 设置旗标；false 清除旗标', parseBoolean, true)
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { state: boolean; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.flag(id, options.state));
    });

  program
    .command('category')
    .description('添加或移除现有 Outlook 分类')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .requiredOption('--category <name>', '分类的精确名称')
    .option('--applied <boolean>', 'true 添加；false 移除', parseBoolean, true)
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { category: string; applied: boolean; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.categorize(id, options.category, options.applied));
    });

  program
    .command('archive')
    .description('将邮件移动到“存档”目录')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--yes', '确认归档')
    .requiredOption('--request-id <id>', '幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { yes?: boolean; requestId: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.archive(id, Boolean(options.yes), options.requestId));
    });

  program
    .command('conversation')
    .description('读取当前邮件会话中已加载的完整往来记录')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.conversation(id));
    });

  program
    .command('reply')
    .description('回复一封邮件；默认生成草稿，显式关闭 draft 时自动发送')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .requiredOption('--content <text>', '回复正文')
    .option('--draft <boolean>', 'true 时交给用户手工发送；false 时自动发送', parseBoolean, true)
    .option('--replyall <boolean>', 'true 时全部答复；false 时只答复发件人', parseBoolean, false)
    .option('--request-id <id>', '自动发送时必填的幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { content: string; draft: boolean; replyall: boolean; requestId?: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.reply(
        id,
        options.content,
        options.draft,
        options.replyall,
        options.requestId,
      ));
    });

  program
    .command('download')
    .description('下载指定邮件附件到本地目录')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .argument('<attachment-id>', 'read/attachments 返回的附件数字短 ID')
    .option('-o, --output <directory>', '下载目录', 'downloads')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, attachmentId: string, options: { output: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.downloadAttachment(id, attachmentId, options.output));
    });

  program
    .command('download-all')
    .description('下载一封邮件的全部附件并计算 SHA-256')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('-o, --output <directory>', '下载目录', 'downloads')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { output: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.downloadAll(id, options.output));
    });

  program
    .command('export')
    .description('将一封邮件导出为 Obsidian Markdown，并下载全部附件')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .requiredOption('-o, --output <directory>', 'Markdown 导出目录')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { output: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.exportObsidian(id, options.output));
    });

  program
    .command('export-batch')
    .description('按日期、目录和筛选条件批量导出 Obsidian Markdown')
    .requiredOption('-o, --output <directory>', 'Markdown 导出目录')
    .option('--date <date>', '指定日期；默认今天')
    .option('--from-date <date>', '日期范围起点')
    .option('--to-date <date>', '日期范围终点')
    .option('--dir <directory>', 'Inbox 子目录')
    .option('-n, --limit <number>', '每页邮件数', parseLimit, 20)
    .option('--sender <text>', '发件人包含匹配')
    .option('--subject <text>', '主题包含匹配')
    .option('--unread', '只导出未读邮件')
    .option('--has-attachments', '只导出带附件邮件')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: {
      output: string; date?: string; fromDate?: string; toDate?: string; dir?: string; limit: number;
      sender?: string; subject?: string; unread?: boolean; hasAttachments?: boolean; json?: boolean;
    }) => {
      await execute(Boolean(options.json), () => service.exportBatch({
        date: options.date, fromDate: options.fromDate, toDate: options.toDate, directory: options.dir,
        limit: options.limit, sender: options.sender, subject: options.subject,
        unread: Boolean(options.unread), hasAttachments: Boolean(options.hasAttachments),
      }, options.output));
    });

  program
    .command('sync-obsidian')
    .description('按日期和目录增量同步邮件到 Obsidian，重复运行会跳过未变化邮件')
    .requiredOption('-o, --output <directory>', 'Obsidian Vault 中的目标目录')
    .option('--date <date>', '指定日期；默认今天')
    .option('--from-date <date>', '日期范围起点')
    .option('--to-date <date>', '日期范围终点')
    .option('--dir <directory>', 'Inbox 子目录')
    .option('-n, --limit <number>', '每页邮件数', parseLimit, 20)
    .option('--sender <text>', '发件人包含匹配')
    .option('--subject <text>', '主题包含匹配')
    .option('--unread', '只同步未读邮件')
    .option('--has-attachments', '只同步带附件邮件')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: {
      output: string; date?: string; fromDate?: string; toDate?: string; dir?: string; limit: number;
      sender?: string; subject?: string; unread?: boolean; hasAttachments?: boolean; json?: boolean;
    }) => {
      await execute(Boolean(options.json), () => service.syncObsidian({
        date: options.date, fromDate: options.fromDate, toDate: options.toDate, directory: options.dir,
        limit: options.limit, sender: options.sender, subject: options.subject,
        unread: Boolean(options.unread), hasAttachments: Boolean(options.hasAttachments),
      }, options.output));
    });

  program
    .command('watch')
    .description('轮询今天的新邮件并逐行输出 JSONL；首次运行默认只建立基线')
    .option('--dir <directory>', 'Inbox 子目录；默认 Inbox')
    .option('--interval <seconds>', '轮询间隔秒数', parseWatchInterval, 30)
    .option('--emit-existing', '首次运行也输出当前已有邮件')
    .option('--state <path>', '自定义持久化状态文件')
    .option('--iterations <number>', '轮询次数；0 表示持续运行', parseIterations, 0)
    .action(async (options: {
      dir?: string; interval: number; emitExisting?: boolean; state?: string; iterations: number;
    }) => {
      const controller = new AbortController();
      const stop = () => controller.abort();
      process.once('SIGINT', stop);
      process.once('SIGTERM', stop);
      try {
        await watchMail(service, {
          directory: options.dir,
          intervalSeconds: options.interval,
          emitExisting: Boolean(options.emitExisting),
          statePath: options.state,
          iterations: options.iterations,
          signal: controller.signal,
          onEvent: event => { process.stdout.write(`${JSON.stringify(event)}\n`); },
        });
      } finally {
        process.removeListener('SIGINT', stop);
        process.removeListener('SIGTERM', stop);
      }
    });

  program
    .command('move')
    .description('将邮件移动到完全匹配的 Outlook 目录')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .argument('<folder>', 'folders 返回的精确目录名')
    .option('--yes', '确认执行移动')
    .requiredOption('--request-id <id>', '幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, folder: string, options: { yes?: boolean; requestId: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.move(id, folder, Boolean(options.yes), options.requestId));
    });

  program
    .command('delete')
    .description('将邮件移入“已删除邮件”')
    .argument('<id>', '列表返回的数字短 ID 或 stableId')
    .option('--yes', '确认执行删除')
    .requiredOption('--request-id <id>', '幂等请求 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { yes?: boolean; requestId: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.delete(id, Boolean(options.yes), options.requestId));
    });

  return program;
}

export async function main(argv = process.argv): Promise<void> {
  const service = new OutlookService(createBrowserBackend());
  try {
    await createProgram(service).parseAsync(argv);
  } catch (error) {
    if (error instanceof CommanderError) {
      if (error.code === 'commander.helpDisplayed' || error.code === 'commander.version') return;
      const appError = new AppError('INVALID_ARGUMENT', error.message);
      process.exitCode = appError.exitCode;
      if (requestedJson(argv)) writeJson(errorResult(appError));
      return;
    }
    throw error;
  } finally {
    await service.close().catch(() => undefined);
  }
}

await main();
