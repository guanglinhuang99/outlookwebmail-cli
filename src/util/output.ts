import type { CliResult } from '../types/result.js';
import { toAppError } from './errors.js';

export function successResult<T>(data: T): CliResult<T> {
  return { ok: true, data };
}

export function errorResult(error: unknown): CliResult<never> {
  const appError = toAppError(error);
  return {
    ok: false,
    error: {
      code: appError.code,
      message: appError.message,
    },
  };
}

export function writeJson(value: unknown): void {
  process.stdout.write(`${JSON.stringify(value)}\n`);
}

export function writePretty(value: unknown): void {
  process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
}
