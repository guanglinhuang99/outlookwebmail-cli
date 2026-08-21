import type { BrowserStatus, InspectResult, LoginHandoffResult, MailReadyProbe, MessageInspectResult } from '../types/inspect.js';
import type {
  AttachmentDownloadResult,
  ComposeActionResult,
  ComposeOptions,
  ConversationOpenResult,
  DraftOpenResult,
  DraftUpdateOptions,
  FolderSelectionResult,
  ForwardOptions,
  InboxFolderListResult,
  MessageActionResult,
  MessageDownloadResult,
  MessageStateActionResult,
  MessageLocator,
  MessageOpenResult,
  ReplyActionResult,
} from '../types/mail.js';

export interface BrowserBackend {
  readonly name?: BrowserBackendName;
  status(): Promise<BrowserStatus>;
  waitUntilMailReady(timeoutMs?: number): Promise<MailReadyProbe>;
  handoffForLogin(): Promise<LoginHandoffResult>;
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
  downloadMessageAsEml(locator: MessageLocator, outputDirectory: string): Promise<MessageDownloadResult>;
  replyMessage(locator: MessageLocator, content: string, draft: boolean, replyAll: boolean): Promise<ReplyActionResult>;
  composeMessage(options: ComposeOptions): Promise<ComposeActionResult>;
  forwardMessage(locator: MessageLocator, options: ForwardOptions): Promise<ComposeActionResult>;
  selectSystemFolder(folder: '草稿'): Promise<FolderSelectionResult>;
  openDraft(locator: MessageLocator, closeAfterRead?: boolean): Promise<DraftOpenResult>;
  updateDraft(locator: MessageLocator, options: DraftUpdateOptions): Promise<ComposeActionResult>;
  sendDraft(locator: MessageLocator): Promise<ComposeActionResult>;
  discardDraft(locator: MessageLocator): Promise<MessageActionResult>;
  setReadState(locator: MessageLocator, unread: boolean): Promise<MessageStateActionResult>;
  setFlagState(locator: MessageLocator, flagged: boolean): Promise<MessageStateActionResult>;
  setCategoryState(locator: MessageLocator, category: string, applied: boolean): Promise<MessageStateActionResult>;
  getConversation(locator: MessageLocator): Promise<ConversationOpenResult>;
  close?(): Promise<void>;
}

export type BrowserBackendName = 'playwright' | 'ego-lite';
