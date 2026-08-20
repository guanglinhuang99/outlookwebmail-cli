export type ErrorCode =
  | 'EGO_BROWSER_NOT_FOUND'
  | 'EGO_BROWSER_ERROR'
  | 'BROWSER_NOT_FOUND'
  | 'SHARED_EDGE_NOT_AVAILABLE'
  | 'PROFILE_ACCESS_DENIED'
  | 'PROFILE_LOCKED'
  | 'PLAYWRIGHT_TIMEOUT'
  | 'PLAYWRIGHT_ERROR'
  | 'AUTH_REQUIRED'
  | 'OUTLOOK_NOT_READY'
  | 'UI_CHANGED'
  | 'MESSAGE_NOT_FOUND'
  | 'AMBIGUOUS_MESSAGE'
  | 'ATTACHMENT_NOT_FOUND'
  | 'FOLDER_NOT_FOUND'
  | 'AMBIGUOUS_FOLDER'
  | 'CONFIRMATION_REQUIRED'
  | 'OPERATION_FAILED'
  | 'OPERATION_UNKNOWN'
  | 'INVALID_ARGUMENT'
  | 'SESSION_INVALID'
  | 'TIMEOUT';

export type CliResult<T> =
  | { ok: true; data: T }
  | {
      ok: false;
      error: {
        code: ErrorCode;
        message: string;
      };
    };
