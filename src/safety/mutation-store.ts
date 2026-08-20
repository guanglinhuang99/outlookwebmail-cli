import { createHash, randomUUID } from 'node:crypto';
import { appendFile, chmod, mkdir, readFile, rename, stat, unlink, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { z } from 'zod';
import { AppError } from '../util/errors.js';

const recordSchema = z.object({
  requestId: z.string(),
  action: z.string(),
  payloadHash: z.string(),
  mailId: z.string().optional(),
  status: z.enum(['pending', 'succeeded', 'failed', 'unknown']),
  createdAt: z.string(),
  completedAt: z.string().optional(),
  result: z.unknown().optional(),
  errorCode: z.string().optional(),
});

const fileSchema = z.object({
  version: z.literal(1),
  records: z.record(z.string(), recordSchema),
});

type MutationFile = z.infer<typeof fileSchema>;

export interface MutationAuditEvent {
  requestId: string;
  action: string;
  mailId: string;
  status: 'pending' | 'succeeded' | 'failed' | 'unknown' | 'deduplicated';
  errorCode?: string;
}

export class MutationStore {
  readonly path: string;
  readonly auditPath: string;

  constructor(
    path = join(homedir(), '.webmail-cli', 'mutations.json'),
    auditPath = join(homedir(), '.webmail-cli', 'audit.jsonl'),
  ) {
    this.path = path;
    this.auditPath = auditPath;
  }

  validateRequestId(requestId?: string | null): string {
    const normalized = requestId?.normalize('NFKC').trim() || '';
    if (!/^[A-Za-z0-9._:-]{8,128}$/.test(normalized)) {
      throw new AppError('INVALID_ARGUMENT', '--request-id 必须由 8 到 128 个字母、数字、点、下划线、冒号或连字符组成。');
    }
    return normalized;
  }

  payloadHash(action: string, payload: unknown): string {
    return createHash('sha256').update(`${action}\u001f${JSON.stringify(payload)}`).digest('base64url');
  }

  private async withFileLock<T>(operation: () => Promise<T>): Promise<T> {
    const lockPath = `${this.path}.lock`;
    await mkdir(dirname(lockPath), { recursive: true, mode: 0o700 });
    for (let attempt = 0; attempt < 80; attempt += 1) {
      try {
        await writeFile(lockPath, `${JSON.stringify({ pid: process.pid, createdAt: new Date().toISOString() })}\n`, {
          encoding: 'utf8', mode: 0o600, flag: 'wx',
        });
        try {
          return await operation();
        } finally {
          await unlink(lockPath).catch(() => undefined);
        }
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
        const age = Date.now() - (await stat(lockPath).catch(() => ({ mtimeMs: Date.now() }))).mtimeMs;
        if (age > 30_000) {
          await unlink(lockPath).catch(() => undefined);
          continue;
        }
        await new Promise(resolve => setTimeout(resolve, 25));
      }
    }
    throw new AppError('OPERATION_FAILED', '幂等状态正被另一进程更新，请稍后重试。');
  }

  private async read(): Promise<MutationFile> {
    try {
      return fileSchema.parse(JSON.parse(await readFile(this.path, 'utf8')));
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'ENOENT') return { version: 1, records: {} };
      throw new AppError('SESSION_INVALID', `幂等状态文件无效：${this.path}`, { cause: error });
    }
  }

  private async write(value: MutationFile): Promise<void> {
    const parsed = fileSchema.parse(value);
    const directory = dirname(this.path);
    const temporary = join(directory, `.mutations-${process.pid}-${randomUUID()}.tmp`);
    await mkdir(directory, { recursive: true, mode: 0o700 });
    try {
      await writeFile(temporary, `${JSON.stringify(parsed, null, 2)}\n`, { encoding: 'utf8', mode: 0o600 });
      await rename(temporary, this.path);
      await chmod(this.path, 0o600);
    } catch (error) {
      await unlink(temporary).catch(() => undefined);
      throw error;
    }
  }

  private async audit(event: MutationAuditEvent): Promise<void> {
    await mkdir(dirname(this.auditPath), { recursive: true, mode: 0o700 });
    await appendFile(this.auditPath, `${JSON.stringify({ timestamp: new Date().toISOString(), ...event })}\n`, {
      encoding: 'utf8', mode: 0o600,
    });
    await chmod(this.auditPath, 0o600);
  }

  async prior<T>(requestId: string, action: string, payloadHash: string, mailId: string): Promise<T | null> {
    const record = (await this.read()).records[requestId];
    if (!record) return null;
    if (record.action !== action || record.payloadHash !== payloadHash) {
      throw new AppError('INVALID_ARGUMENT', `--request-id ${requestId} 已用于不同的操作或参数。`);
    }
    if (record.status === 'succeeded') {
      await this.audit({ requestId, action, mailId: record.mailId ?? mailId, status: 'deduplicated' });
      return record.result as T;
    }
    if (record.status === 'pending' || record.status === 'unknown') {
      throw new AppError(
        'OPERATION_UNKNOWN',
        `请求 ${requestId} 状态未决；请保留该 request-id 并人工核查邮箱，禁止换用新 request-id 自动重试。`,
      );
    }
    throw new AppError('OPERATION_FAILED', `请求 ${requestId} 曾明确失败；请核查邮箱后再决定是否创建新请求。`);
  }

  async begin(requestId: string, action: string, payloadHash: string, mailId: string): Promise<void> {
    const claimPath = `${this.path}.${requestId}.claim`;
    await mkdir(dirname(claimPath), { recursive: true, mode: 0o700 });
    try {
      await writeFile(claimPath, `${JSON.stringify({ requestId, action, payloadHash })}\n`, {
        encoding: 'utf8', mode: 0o600, flag: 'wx',
      });
    } catch (error) {
      if ((error as NodeJS.ErrnoException).code === 'EEXIST') {
        throw new AppError('OPERATION_FAILED', `请求 ${requestId} 已被另一进程接收；请先核查执行结果。`);
      }
      throw error;
    }
    await this.withFileLock(async () => {
      const file = await this.read();
      if (file.records[requestId]) {
        throw new AppError('OPERATION_FAILED', `请求 ${requestId} 已存在，无法重复开始。`);
      }
      file.records[requestId] = {
        requestId, action, payloadHash, mailId, status: 'pending', createdAt: new Date().toISOString(),
      };
      await this.write(file);
    });
    await this.audit({ requestId, action, mailId, status: 'pending' });
  }

  async succeed(requestId: string, result: unknown, event: Omit<MutationAuditEvent, 'requestId' | 'status'>): Promise<void> {
    await this.withFileLock(async () => {
      const file = await this.read();
      const record = file.records[requestId];
      if (!record) throw new AppError('SESSION_INVALID', `找不到请求状态：${requestId}`);
      record.status = 'succeeded';
      record.completedAt = new Date().toISOString();
      record.result = result;
      await this.write(file);
    });
    await this.audit({ requestId, ...event, status: 'succeeded' });
  }

  async fail(requestId: string, errorCode: string, event: Omit<MutationAuditEvent, 'requestId' | 'status' | 'errorCode'>): Promise<void> {
    const updated = await this.withFileLock(async () => {
      const file = await this.read();
      const record = file.records[requestId];
      if (!record) return false;
      record.status = 'failed';
      record.completedAt = new Date().toISOString();
      record.errorCode = errorCode;
      await this.write(file);
      return true;
    });
    if (!updated) return;
    await this.audit({ requestId, ...event, status: 'failed', errorCode });
  }

  async uncertain(requestId: string, errorCode: string, event: Omit<MutationAuditEvent, 'requestId' | 'status' | 'errorCode'>): Promise<void> {
    const updated = await this.withFileLock(async () => {
      const file = await this.read();
      const record = file.records[requestId];
      if (!record) return false;
      record.status = 'unknown';
      record.completedAt = new Date().toISOString();
      record.errorCode = errorCode;
      await this.write(file);
      return true;
    });
    if (!updated) return;
    await this.audit({ requestId, ...event, status: 'unknown', errorCode });
  }
}
