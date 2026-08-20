import { McpServer } from '@modelcontextprotocol/server';
import type { OutlookService } from '../outlook/service.js';
import { createWebmailTools, invokeWebmailTool } from './tools.js';

export function createWebmailMcpServer(service: OutlookService): McpServer {
  const server = new McpServer(
    { name: 'outlook-webmail', version: '0.3.0' },
    {
      instructions: '通过已登录的 Outlook Web 浏览器读写邮件。先调用 list_messages/search_messages，再用返回的 ID 操作邮件。破坏性操作必须显式确认并提供幂等 requestId。',
    },
  );
  for (const definition of createWebmailTools(service)) {
    server.registerTool(
      definition.name,
      {
        title: definition.title,
        description: definition.description,
        inputSchema: definition.inputSchema,
        annotations: {
          readOnlyHint: definition.readOnly,
          destructiveHint: Boolean(definition.destructive),
          idempotentHint: Boolean(definition.destructive),
        },
      },
      async args => await invokeWebmailTool(definition, args as Record<string, unknown>),
    );
  }
  return server;
}
