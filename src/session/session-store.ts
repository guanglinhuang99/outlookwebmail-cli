import { randomUUID } from 'node:crypto';
import { mkdir, readFile, rename, unlink, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { z } from 'zod';
import { AppError } from '../util/errors.js';

const sessionMessageSchema = z.object({
  subject: z.string(),
  senderName: z.string().nullable(),
  senderAddress: z.string().nullable(),
  receivedAt: z.string().nullable().default(null),
  receivedAtText: z.string().nullable(),
  preview: z.string().nullable(),
  hasAttachments: z.boolean().nullable().default(null),
});

const sessionSchema = z.object({
  version: z.literal(1),
  updatedAt: z.string(),
  source: z.string(),
  messages: z.record(z.string(), sessionMessageSchema),
  stableMessages: z.record(z.string(), sessionMessageSchema).optional(),
});

export type MailSession = z.infer<typeof sessionSchema>;

export class SessionStore {
  readonly path: string;

  constructor(path = join(homedir(), '.webmail-cli', 'session.json')) {
    this.path = path;
  }

  async read(): Promise<MailSession | null> {
    let raw: string;
    try {
      raw = await readFile(this.path, 'utf8');
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return null;
      throw error;
    }

    try {
      return sessionSchema.parse(JSON.parse(raw));
    } catch (error) {
      throw new AppError('SESSION_INVALID', `Session 文件无效：${this.path}`, { cause: error });
    }
  }

  async write(session: MailSession): Promise<void> {
    const value = sessionSchema.parse(session);
    const directory = dirname(this.path);
    const tempPath = join(directory, `.session-${process.pid}-${randomUUID()}.tmp`);
    await mkdir(directory, { recursive: true, mode: 0o700 });

    try {
      await writeFile(tempPath, `${JSON.stringify(value, null, 2)}\n`, { encoding: 'utf8', mode: 0o600 });
      await rename(tempPath, this.path);
    } catch (error) {
      await unlink(tempPath).catch(() => undefined);
      throw error;
    }
  }

  async clear(): Promise<void> {
    await unlink(this.path).catch((error: NodeJS.ErrnoException) => {
      if (error.code !== 'ENOENT') throw error;
    });
  }
}
