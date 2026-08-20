import { describe, expect, it, vi } from 'vitest';
import type { OutlookService } from '../src/outlook/service.js';
import { createSerialExecutor } from '../src/mcp/server.js';
import { createWebmailTools, invokeWebmailTool } from '../src/mcp/tools.js';

describe('MCP webmail tools', () => {
  it('提供标准化工具，并保持回复默认草稿和单独回复', async () => {
    const reply = vi.fn(async () => ({ draft: true, sent: false }));
    const service = { reply } as unknown as OutlookService;
    const tools = createWebmailTools(service);
    expect(tools.map(tool => tool.name)).toEqual(expect.arrayContaining([
      'list_messages', 'get_message', 'reply_message', 'move_message', 'delete_message', 'sync_obsidian',
    ]));
    const definition = tools.find(tool => tool.name === 'reply_message')!;
    const result = await invokeWebmailTool(definition, { id: '1', content: '收到' });
    expect(reply).toHaveBeenCalledWith('1', '收到', true, false, undefined);
    expect(result.isError).toBeUndefined();
  });

  it('破坏性工具要求 confirmed 和 requestId，错误不暴露堆栈', async () => {
    const service = { delete: vi.fn() } as unknown as OutlookService;
    const definition = createWebmailTools(service).find(tool => tool.name === 'delete_message')!;
    const result = await invokeWebmailTool(definition, { id: '1' });
    expect(result.isError).toBe(true);
    expect(result.structuredContent).toMatchObject({ error: { code: 'INVALID_ARGUMENT' } });
    expect(result.content[0]!.text).not.toContain('stack');
    expect(service.delete).not.toHaveBeenCalled();
  });

  it('serializes stateful browser operations even when MCP calls arrive concurrently', async () => {
    const serial = createSerialExecutor();
    const order: string[] = [];
    let releaseFirst!: () => void;
    const gate = new Promise<void>(resolve => { releaseFirst = resolve; });
    const first = serial(async () => { order.push('first:start'); await gate; order.push('first:end'); });
    const second = serial(async () => { order.push('second:start'); order.push('second:end'); });
    await Promise.resolve();
    expect(order).toEqual(['first:start']);
    releaseFirst();
    await Promise.all([first, second]);
    expect(order).toEqual(['first:start', 'first:end', 'second:start', 'second:end']);
  });
});
