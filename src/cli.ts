#!/usr/bin/env node

import { Command, CommanderError, InvalidArgumentError } from 'commander';
import { EgoLiteBackend } from './browser/ego-lite.js';
import { OutlookService } from './outlook/service.js';
import { AppError, toAppError } from './util/errors.js';
import { errorResult, successResult, writeJson, writePretty } from './util/output.js';

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

export function createProgram(service = new OutlookService(new EgoLiteBackend())): Command {
  const program = new Command();
  program
    .name('webmail')
    .description('Outlook Web CLI through Ego Lite')
    .version('0.2.0')
    .showHelpAfterError()
    .exitOverride();

  program
    .command('status')
    .description('检查 Ego Lite 与 Outlook 登录状态')
    .option('--json', '输出单一 JSON envelope')
    .action(async (options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.status());
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
    .argument('<id>', 'inbox/search 返回的数字短 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.read(id));
    });

  program
    .command('attachments')
    .description('列出邮件附件元数据，不下载附件')
    .argument('<id>', 'inbox/search 返回的数字短 ID')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { json?: boolean }) => {
      await execute(Boolean(options.json), () => service.attachments(id));
    });

  program
    .command('download')
    .description('下载指定邮件附件到本地目录')
    .argument('<id>', 'inbox/search/today 返回的邮件数字短 ID')
    .argument('<attachment-id>', 'read/attachments 返回的附件数字短 ID')
    .option('-o, --output <directory>', '下载目录', 'downloads')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, attachmentId: string, options: { output: string; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.downloadAttachment(id, attachmentId, options.output));
    });

  program
    .command('move')
    .description('将邮件移动到完全匹配的 Outlook 目录')
    .argument('<id>', 'inbox/search/today 返回的邮件数字短 ID')
    .argument('<folder>', 'folders 返回的精确目录名')
    .option('--yes', '确认执行移动')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, folder: string, options: { yes?: boolean; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.move(id, folder, Boolean(options.yes)));
    });

  program
    .command('delete')
    .description('将邮件移入“已删除邮件”')
    .argument('<id>', 'inbox/search/today 返回的邮件数字短 ID')
    .option('--yes', '确认执行删除')
    .option('--json', '输出单一 JSON envelope')
    .action(async (id: string, options: { yes?: boolean; json?: boolean }) => {
      await execute(Boolean(options.json), () => service.delete(id, Boolean(options.yes)));
    });

  return program;
}

export async function main(argv = process.argv): Promise<void> {
  try {
    await createProgram().parseAsync(argv);
  } catch (error) {
    if (error instanceof CommanderError) {
      if (error.code === 'commander.helpDisplayed' || error.code === 'commander.version') return;
      const appError = new AppError('INVALID_ARGUMENT', error.message);
      process.exitCode = appError.exitCode;
      if (requestedJson(argv)) writeJson(errorResult(appError));
      return;
    }
    throw error;
  }
}

await main();
