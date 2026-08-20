import * as z from 'zod/v4';
import type { OutlookService } from '../outlook/service.js';
import { toAppError } from '../util/errors.js';

type ToolArguments = Record<string, unknown>;

export interface WebmailToolDefinition {
  name: string;
  title: string;
  description: string;
  inputSchema: z.ZodObject<z.ZodRawShape>;
  readOnly: boolean;
  destructive?: boolean;
  invoke(args: ToolArguments): Promise<unknown>;
}

const id = z.string().min(1).describe('list_messages/search_messages 返回的短 ID 或 stableId');
const requestId = z.string().min(8).describe('调用方生成的唯一幂等请求 ID');
const recipients = z.array(z.string().min(1)).default([]);
const dateListSchema = z.object({
  date: z.string().optional(),
  fromDate: z.string().optional(),
  toDate: z.string().optional(),
  directory: z.string().optional(),
  limit: z.number().int().min(1).max(100).default(20),
  cursor: z.string().optional(),
  sender: z.string().optional(),
  subject: z.string().optional(),
  unread: z.boolean().default(false),
  hasAttachments: z.boolean().default(false),
});

function tool<T extends z.ZodRawShape>(
  service: OutlookService,
  definition: Omit<WebmailToolDefinition, 'inputSchema' | 'invoke'> & {
    inputSchema: z.ZodObject<T>;
    run: (args: z.output<z.ZodObject<T>>, service: OutlookService) => Promise<unknown>;
  },
): WebmailToolDefinition {
  return {
    ...definition,
    inputSchema: definition.inputSchema,
    invoke: async args => await definition.run(definition.inputSchema.parse(args), service),
  };
}

export function createWebmailTools(service: OutlookService): WebmailToolDefinition[] {
  return [
    tool(service, {
      name: 'status', title: 'Webmail status', description: '检查浏览器后端和 Outlook 登录状态。',
      inputSchema: z.object({}), readOnly: true, run: async (_args, current) => await current.status(),
    }),
    tool(service, {
      name: 'list_folders', title: 'List mail folders', description: '递归列出 Inbox 下的目录。',
      inputSchema: z.object({}), readOnly: true, run: async (_args, current) => await current.folders(),
    }),
    tool(service, {
      name: 'list_messages', title: 'List messages', description: '按日期、目录、过滤条件和游标分页列出邮件；日期默认今天，目录默认 Inbox。',
      inputSchema: dateListSchema, readOnly: true, run: async (args, current) => await current.listByDate(args),
    }),
    tool(service, {
      name: 'search_messages', title: 'Search messages', description: '使用 Outlook 原生搜索邮件。',
      inputSchema: z.object({ query: z.string().min(1), limit: z.number().int().min(1).max(100).default(20) }),
      readOnly: true, run: async (args, current) => await current.search(args.query, { limit: args.limit }),
    }),
    tool(service, {
      name: 'get_message', title: 'Get message', description: '读取邮件正文、收发件人和附件元数据。',
      inputSchema: z.object({ id }), readOnly: true, run: async (args, current) => await current.read(args.id),
    }),
    tool(service, {
      name: 'get_conversation', title: 'Get conversation', description: '读取邮件会话中已加载的全部往来。',
      inputSchema: z.object({ id }), readOnly: true, run: async (args, current) => await current.conversation(args.id),
    }),
    tool(service, {
      name: 'download_attachments', title: 'Download attachments', description: '下载一封邮件的全部附件并计算 SHA-256。',
      inputSchema: z.object({ id, outputDirectory: z.string().min(1) }), readOnly: false,
      run: async (args, current) => await current.downloadAll(args.id, args.outputDirectory),
    }),
    tool(service, {
      name: 'sync_obsidian', title: 'Sync to Obsidian', description: '按日期和目录增量同步邮件、附件和附件索引到 Obsidian。',
      inputSchema: dateListSchema.omit({ cursor: true }).extend({ outputDirectory: z.string().min(1) }), readOnly: false,
      run: async (args, current) => {
        const { outputDirectory, ...options } = args;
        return await current.syncObsidian(options, outputDirectory);
      },
    }),
    tool(service, {
      name: 'create_message', title: 'Create message', description: '新建邮件；draft 默认 true，只有 draft=false 且提供 requestId 时才自动发送。',
      inputSchema: z.object({
        to: recipients, cc: recipients, bcc: recipients, subject: z.string().min(1), content: z.string().min(1),
        attachments: z.array(z.string().min(1)).default([]), draft: z.boolean().default(true), requestId: requestId.optional(),
      }),
      readOnly: false,
      run: async (args, current) => await current.compose({
        to: args.to, cc: args.cc, bcc: args.bcc, subject: args.subject, content: args.content,
        attachments: args.attachments, draft: args.draft,
      }, args.requestId),
    }),
    tool(service, {
      name: 'reply_message', title: 'Reply to message', description: '回复邮件；draft 默认 true，replyAll 默认 false。自动发送必须提供 requestId。',
      inputSchema: z.object({
        id, content: z.string().min(1), draft: z.boolean().default(true), replyAll: z.boolean().default(false),
        requestId: requestId.optional(),
      }),
      readOnly: false,
      run: async (args, current) => await current.reply(args.id, args.content, args.draft, args.replyAll, args.requestId),
    }),
    tool(service, {
      name: 'forward_message', title: 'Forward message', description: '转发邮件并保留原附件；默认保存草稿。',
      inputSchema: z.object({
        id, to: recipients, cc: recipients, bcc: recipients, content: z.string().min(1),
        attachments: z.array(z.string().min(1)).default([]), draft: z.boolean().default(true), requestId: requestId.optional(),
      }),
      readOnly: false,
      run: async (args, current) => await current.forward(args.id, {
        to: args.to, cc: args.cc, bcc: args.bcc, content: args.content,
        attachments: args.attachments, draft: args.draft,
      }, args.requestId),
    }),
    tool(service, {
      name: 'set_read_state', title: 'Set read state', description: '把邮件标记为已读或未读。',
      inputSchema: z.object({ id, unread: z.boolean() }), readOnly: false,
      run: async (args, current) => await current.markRead(args.id, args.unread),
    }),
    tool(service, {
      name: 'set_flag', title: 'Set message flag', description: '设置或清除邮件旗标。',
      inputSchema: z.object({ id, flagged: z.boolean() }), readOnly: false,
      run: async (args, current) => await current.flag(args.id, args.flagged),
    }),
    tool(service, {
      name: 'set_category', title: 'Set message category', description: '添加或移除一个现有 Outlook 分类。',
      inputSchema: z.object({ id, category: z.string().min(1), applied: z.boolean() }), readOnly: false,
      run: async (args, current) => await current.categorize(args.id, args.category, args.applied),
    }),
    tool(service, {
      name: 'move_message', title: 'Move message', description: '移动邮件。必须 confirmed=true 并提供 requestId。',
      inputSchema: z.object({ id, folder: z.string().min(1), confirmed: z.boolean(), requestId }),
      readOnly: false, destructive: true,
      run: async (args, current) => await current.move(args.id, args.folder, args.confirmed, args.requestId),
    }),
    tool(service, {
      name: 'archive_message', title: 'Archive message', description: '将邮件移到存档。必须 confirmed=true 并提供 requestId。',
      inputSchema: z.object({ id, confirmed: z.boolean(), requestId }), readOnly: false, destructive: true,
      run: async (args, current) => await current.archive(args.id, args.confirmed, args.requestId),
    }),
    tool(service, {
      name: 'delete_message', title: 'Delete message', description: '将邮件移入已删除邮件。必须 confirmed=true 并提供 requestId。',
      inputSchema: z.object({ id, confirmed: z.boolean(), requestId }), readOnly: false, destructive: true,
      run: async (args, current) => await current.delete(args.id, args.confirmed, args.requestId),
    }),
    tool(service, {
      name: 'list_drafts', title: 'List drafts', description: '列出草稿邮件。',
      inputSchema: z.object({ limit: z.number().int().min(1).max(100).default(20) }), readOnly: true,
      run: async (args, current) => await current.drafts({ limit: args.limit }),
    }),
    tool(service, {
      name: 'get_draft', title: 'Get draft', description: '读取草稿内容和附件元数据。',
      inputSchema: z.object({ id }), readOnly: true, run: async (args, current) => await current.readDraft(args.id),
    }),
    tool(service, {
      name: 'update_draft', title: 'Update draft', description: '更新草稿字段或追加附件。',
      inputSchema: z.object({
        id, to: z.array(z.string()).optional(), cc: z.array(z.string()).optional(), bcc: z.array(z.string()).optional(),
        subject: z.string().optional(), content: z.string().optional(), attachments: z.array(z.string()).default([]),
      }),
      readOnly: false,
      run: async (args, current) => await current.updateDraft(args.id, {
        to: args.to, cc: args.cc, bcc: args.bcc, subject: args.subject, content: args.content, attachments: args.attachments,
      }),
    }),
    tool(service, {
      name: 'send_draft', title: 'Send draft', description: '发送草稿；必须提供 requestId。',
      inputSchema: z.object({ id, requestId }), readOnly: false,
      run: async (args, current) => await current.sendDraft(args.id, args.requestId),
    }),
    tool(service, {
      name: 'discard_draft', title: 'Discard draft', description: '放弃草稿；必须 confirmed=true 并提供 requestId。',
      inputSchema: z.object({ id, confirmed: z.boolean(), requestId }), readOnly: false, destructive: true,
      run: async (args, current) => await current.discardDraft(args.id, args.confirmed, args.requestId),
    }),
  ];
}

function jsonObject(value: unknown): Record<string, unknown> {
  const normalized = JSON.parse(JSON.stringify(value)) as unknown;
  return normalized && typeof normalized === 'object' && !Array.isArray(normalized)
    ? normalized as Record<string, unknown>
    : { result: normalized };
}

export async function invokeWebmailTool(definition: WebmailToolDefinition, args: ToolArguments): Promise<{
  content: Array<{ type: 'text'; text: string }>;
  structuredContent: Record<string, unknown>;
  isError?: true;
}> {
  try {
    const data = jsonObject(await definition.invoke(args));
    return { content: [{ type: 'text', text: JSON.stringify(data) }], structuredContent: data };
  } catch (error) {
    if (error instanceof z.ZodError) {
      const data = {
        error: {
          code: 'INVALID_ARGUMENT',
          message: error.issues.map(issue => `${issue.path.join('.') || 'arguments'}: ${issue.message}`).join('; '),
        },
      };
      return { content: [{ type: 'text', text: JSON.stringify(data) }], structuredContent: data, isError: true };
    }
    const appError = toAppError(error);
    const data = { error: { code: appError.code, message: appError.message } };
    return { content: [{ type: 'text', text: JSON.stringify(data) }], structuredContent: data, isError: true };
  }
}
