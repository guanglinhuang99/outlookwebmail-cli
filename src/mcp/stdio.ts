#!/usr/bin/env node

import { serveStdio } from '@modelcontextprotocol/server/stdio';
import { createBrowserBackend } from '../browser/browser-factory.js';
import { OutlookService } from '../outlook/service.js';
import { createWebmailMcpServer } from './server.js';

const services = new Set<OutlookService>();
const handle = serveStdio(() => {
  const service = new OutlookService(createBrowserBackend());
  services.add(service);
  return createWebmailMcpServer(service);
}, {
  onerror: error => process.stderr.write(`[MCP_ERROR] ${error.message}\n`),
});

async function shutdown(): Promise<void> {
  await handle.close().catch(() => undefined);
  await Promise.all(Array.from(services, service => service.close().catch(() => undefined)));
}

process.once('SIGINT', () => { void shutdown().finally(() => process.exit(0)); });
process.once('SIGTERM', () => { void shutdown().finally(() => process.exit(0)); });
