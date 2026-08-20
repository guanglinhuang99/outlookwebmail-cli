export interface MailAddress {
  name: string | null;
  address: string | null;
}

export interface MailSummary {
  id: string;
  sender: MailAddress;
  subject: string;
  receivedAt: string | null;
  receivedAtText: string | null;
  preview: string | null;
  unread: boolean | null;
  hasAttachments: boolean | null;
}

export interface AttachmentSummary {
  id: string;
  filename: string;
  sizeText: string | null;
}

export interface FolderSummary {
  name: string;
  path: string;
  level: number;
  expanded: boolean | null;
}

export interface InboxFolderListResult {
  accountCount: number;
  inboxCount: number;
  complete: boolean;
  folders: FolderSummary[];
}

export interface FolderSelectionResult {
  count: number;
  selected: boolean;
  folder: FolderSummary | null;
}

export interface MailMessage {
  id: string;
  subject: string;
  from: MailAddress;
  to: MailAddress[];
  cc: MailAddress[];
  receivedAt: string | null;
  receivedAtText: string | null;
  bodyText: string;
  attachments: AttachmentSummary[];
}

export interface RawMessageRow {
  stableHint: string | null;
  senderName: string | null;
  senderAddress: string | null;
  subject: string;
  receivedAt: string | null;
  receivedAtText: string | null;
  preview: string | null;
  unread: boolean | null;
  hasAttachments: boolean | null;
}

export interface RawMessage {
  subject: string;
  fromName: string | null;
  fromAddress: string | null;
  to: MailAddress[];
  cc: MailAddress[];
  receivedAt: string | null;
  receivedAtText: string | null;
  bodyText: string;
  attachments: Array<{
    filename: string;
    sizeText: string | null;
  }>;
}

export interface MessageLocator {
  subject: string;
  senderName: string | null;
  senderAddress: string | null;
  receivedAt: string | null;
  receivedAtText: string | null;
  preview: string | null;
  hasAttachments: boolean | null;
}

export interface MessageOpenResult {
  matchCount: number;
  message: RawMessage | null;
}

export type MessageActionStatus =
  | 'performed'
  | 'message_not_found'
  | 'message_ambiguous'
  | 'reading_pane_not_ready'
  | 'control_not_found'
  | 'folder_not_found'
  | 'folder_ambiguous'
  | 'attachment_not_found'
  | 'download_failed';

export interface MessageActionResult {
  matchCount: number;
  status: MessageActionStatus;
  performed: boolean;
  verified: boolean;
  folderMatches?: number;
}

export interface AttachmentDownloadResult extends MessageActionResult {
  attachmentCount?: number;
  attachmentId?: string;
  filename?: string;
  path?: string;
  bytes?: number;
}

export type ReplyActionStatus =
  | 'draft_ready'
  | 'sent'
  | 'message_not_found'
  | 'message_ambiguous'
  | 'reading_pane_not_ready'
  | 'reply_control_not_found'
  | 'reply_control_ambiguous'
  | 'editor_not_ready'
  | 'content_not_verified'
  | 'send_control_not_found'
  | 'send_not_verified';

export interface ReplyActionResult {
  matchCount: number;
  status: ReplyActionStatus;
  performed: boolean;
  verified: boolean;
  draft: boolean;
  replyAll: boolean;
  handedOff?: boolean;
}

export interface ReplyResult {
  id: string;
  draft: boolean;
  replyAll: boolean;
  sent: boolean;
  requiresManualSend: boolean;
  handedOff: boolean;
  verified: true;
}

export interface ExportedAttachment {
  id: string;
  filename: string;
  path: string;
  bytes: number;
  link: string;
}

export interface ObsidianExportResult {
  id: string;
  markdownPath: string;
  attachmentDirectory: string | null;
  attachments: ExportedAttachment[];
  bytes: number;
}
