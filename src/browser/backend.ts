import type { BrowserStatus, InspectResult, MessageInspectResult } from '../types/inspect.js';
import type {
  AttachmentDownloadResult,
  FolderSelectionResult,
  InboxFolderListResult,
  MessageActionResult,
  MessageLocator,
  MessageOpenResult,
  ReplyActionResult,
} from '../types/mail.js';

export interface BrowserBackend {
  status(): Promise<BrowserStatus>;
  snapshot(): Promise<string>;
  eval<T>(script: string): Promise<T>;
  click(refOrLocator: string): Promise<void>;
  clickAndWait(refOrLocator: string, waitMs: number): Promise<void>;
  fill(refOrLocator: string, text: string): Promise<void>;
  fillAndPress(refOrLocator: string, text: string, key: string, waitMs: number): Promise<void>;
  press(key: string): Promise<void>;
  scrollBy(x: number, y: number): Promise<void>;
  wheel(refOrLocator: string, y: number, steps?: number): Promise<void>;
  wheelAndEval<T>(refOrLocator: string, y: number, steps: number, waitMs: number, script: string): Promise<T>;
  wait(ms: number): Promise<void>;
  inspect(): Promise<Omit<InspectResult, 'backend' | 'capturedAt' | 'state'>>;
  inspectMessage(): Promise<Omit<MessageInspectResult, 'backend' | 'capturedAt' | 'state'>>;
  openAndExtractMessage(locator: MessageLocator): Promise<MessageOpenResult>;
  listInboxFolders(): Promise<InboxFolderListResult>;
  selectInboxFolder(directory: string | null): Promise<FolderSelectionResult>;
  deleteMessage(locator: MessageLocator): Promise<MessageActionResult>;
  moveMessage(locator: MessageLocator, folder: string): Promise<MessageActionResult>;
  downloadAttachment(locator: MessageLocator, attachmentIndex: number, outputDirectory: string): Promise<AttachmentDownloadResult>;
  replyMessage(locator: MessageLocator, content: string, draft: boolean, replyAll: boolean): Promise<ReplyActionResult>;
}
