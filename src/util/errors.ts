import type { ErrorCode } from '../types/result.js';

const EXIT_CODES: Record<ErrorCode, number> = {
  INVALID_ARGUMENT: 2,
  AUTH_REQUIRED: 3,
  OUTLOOK_NOT_READY: 4,
  UI_CHANGED: 4,
  MESSAGE_NOT_FOUND: 5,
  AMBIGUOUS_MESSAGE: 5,
  ATTACHMENT_NOT_FOUND: 5,
  FOLDER_NOT_FOUND: 5,
  AMBIGUOUS_FOLDER: 5,
  CONFIRMATION_REQUIRED: 2,
  OPERATION_FAILED: 6,
  OPERATION_UNKNOWN: 6,
  SESSION_INVALID: 5,
  EGO_BROWSER_NOT_FOUND: 6,
  EGO_BROWSER_ERROR: 6,
  BROWSER_NOT_FOUND: 6,
  SHARED_EDGE_NOT_AVAILABLE: 6,
  PROFILE_ACCESS_DENIED: 6,
  PROFILE_LOCKED: 6,
  PLAYWRIGHT_TIMEOUT: 6,
  PLAYWRIGHT_ERROR: 6,
  TIMEOUT: 6,
};

export class AppError extends Error {
  readonly code: ErrorCode;
  readonly exitCode: number;

  constructor(code: ErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'AppError';
    this.code = code;
    this.exitCode = EXIT_CODES[code];
  }
}

export function toAppError(error: unknown): AppError {
  if (error instanceof AppError) {
    return error;
  }

  if (error instanceof Error) {
    return new AppError('EGO_BROWSER_ERROR', error.message, { cause: error });
  }

  return new AppError('EGO_BROWSER_ERROR', '发生未知错误。');
}
