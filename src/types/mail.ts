export interface MailAddress {
  name: string | null;
  address: string | null;
}

export interface MailSummary {
  id: string;
  stableId: string;
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
  stableId: string;
  subject: string;
  from: MailAddress;
  to: MailAddress[];
  cc: MailAddress[];
  receivedAt: string | null;
  receivedAtText: string | null;
  bodyText: string;
  bodyTruncated: boolean;
  bodyBytes: number;
  attachments: AttachmentSummary[];
  unreadRestored: boolean;
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
  bodyTruncated?: boolean;
  bodyBytes?: number;
  attachments: Array<{
    filename: string;
    sizeText: string | null;
  }>;
}

export interface MessageLocator {
  stableHint?: string | null;
  stableId?: string;
  subject: string;
  senderName: string | null;
  senderAddress: string | null;
  receivedAt: string | null;
  receivedAtText: string | null;
  preview: string | null;
  hasAttachments: boolean | null;
  unread?: boolean | null;
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
  requestId?: string;
  deduplicated?: boolean;
}

export interface ComposeOptions {
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  content: string;
  attachments: string[];
  draft: boolean;
}

export type ComposeActionStatus =
  | 'draft_ready'
  | 'sent'
  | 'compose_control_not_found'
  | 'compose_control_ambiguous'
  | 'editor_not_ready'
  | 'field_not_ready'
  | 'content_not_verified'
  | 'recipient_not_verified'
  | 'attachment_not_verified'
  | 'send_control_not_found'
  | 'send_not_verified'
  | 'message_not_found'
  | 'message_ambiguous'
  | 'reading_pane_not_ready';

export interface ComposeActionResult {
  status: ComposeActionStatus;
  performed: boolean;
  verified: boolean;
  draft: boolean;
  handedOff: boolean;
  attachmentCount: number;
  matchCount?: number;
}

export interface ComposeResult {
  draft: boolean;
  sent: boolean;
  requiresManualSend: boolean;
  handedOff: boolean;
  verified: true;
  attachmentCount: number;
  requestId?: string;
  deduplicated?: boolean;
}

export interface ForwardOptions {
  to: string[];
  cc: string[];
  bcc: string[];
  content: string;
  attachments: string[];
  draft: boolean;
}

export interface ForwardResult extends ComposeResult {
  id: string;
  originalAttachmentsPreserved: true;
}

export interface DraftMessage {
  id: string;
  stableId: string;
  to: MailAddress[];
  cc: MailAddress[];
  bcc: MailAddress[];
  subject: string;
  bodyText: string;
  attachments: AttachmentSummary[];
}

export interface RawDraft {
  to: MailAddress[];
  cc: MailAddress[];
  bcc: MailAddress[];
  subject: string;
  bodyText: string;
  attachments: Array<{ filename: string; sizeText: string | null }>;
}

export interface DraftOpenResult {
  matchCount: number;
  draft: RawDraft | null;
  closed?: boolean;
}

export interface DraftUpdateOptions {
  to?: string[];
  cc?: string[];
  bcc?: string[];
  subject?: string;
  content?: string;
  attachments: string[];
}

export type MessageStateKind = 'read' | 'flagged' | 'category';

export interface MessageStateActionResult extends MessageActionResult {
  changed?: boolean;
  state?: boolean;
  category?: string;
}

export interface DownloadedAttachmentWithHash {
  id: string;
  filename: string;
  path: string;
  bytes: number;
  sha256: string;
}

export interface DownloadAllResult {
  id: string;
  outputDirectory: string;
  attachments: DownloadedAttachmentWithHash[];
}

export interface ConversationMessage extends Omit<MailMessage, 'id' | 'stableId' | 'unreadRestored'> {
  index: number;
}

export interface ConversationResult {
  id: string;
  stableId: string;
  subject: string;
  messages: ConversationMessage[];
  complete: boolean;
}

export interface ConversationOpenResult {
  matchCount: number;
  messages: RawMessage[];
  complete: boolean;
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

export type ObsidianSyncStatus = 'created' | 'updated' | 'unchanged';

export interface ObsidianSyncItem {
  id: string;
  stableId: string;
  status: ObsidianSyncStatus;
  markdownPath: string;
  attachmentDirectory: string | null;
  attachments: DownloadedAttachmentWithHash[];
}

export interface ObsidianSyncResult {
  outputDirectory: string;
  fromDate: string;
  toDate: string;
  directory: FolderSummary | null;
  manifestPath: string;
  attachmentIndexPath: string;
  created: number;
  updated: number;
  unchanged: number;
  items: ObsidianSyncItem[];
}
