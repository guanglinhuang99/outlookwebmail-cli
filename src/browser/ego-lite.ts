import type { BrowserBackend, BrowserBackendName } from './backend.js';
import { EgoRunner, type BrowserScriptRunner } from './ego-runner.js';
import type { BrowserStatus, InspectResult, LoginHandoffResult, MessageInspectResult, PageInfo } from '../types/inspect.js';
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
  MessageStateActionResult,
  MessageLocator,
  MessageOpenResult,
  ReplyActionResult,
} from '../types/mail.js';
import { basename } from 'node:path';
import {
  ATTACHMENT_CANDIDATES,
  IFRAME_CANDIDATES,
  LIST_CANDIDATES,
  MESSAGE_ROW_CANDIDATES,
  PAGE_DOM_INVENTORY,
  SCROLL_CANDIDATES,
  MESSAGE_BODY_CANDIDATES,
  MESSAGE_HEADER_CANDIDATES,
} from '../outlook/dom-probes.js';
import { buildMessageExtractScript, buildMessageRowMatchScript } from '../outlook/message-parser.js';
import { buildInboxFolderTargetScript, INBOX_FOLDER_SCAN_SCRIPT } from '../outlook/folder-parser.js';

const TASK_SPACE = 'webmail-cli';
const OUTLOOK_URL = 'https://partner.outlook.cn/mail/';

function markedResult(expression: string): string {
  return `cliLog(JSON.stringify({ __webmail_result__: true, result: ${expression} }));`;
}

function bootstrap(): string {
  return `
const task = await useOrCreateTaskSpace(${JSON.stringify(TASK_SPACE)});
await openOrReuseTab(${JSON.stringify(OUTLOOK_URL)}, { wait: true, timeout: 30 });
`;
}

function resumeTaskSpace(): string {
  return `
const task = await useOrCreateTaskSpace(${JSON.stringify(TASK_SPACE)});
const tab = await ensureRealTab();
if (!tab) throw new Error('Outlook tab is unavailable');
`;
}

interface RawInspectResult {
  taskSpaceId: string | number | null;
  page: PageInfo;
  snapshot: string;
  domInventory: unknown[];
  listCandidates: unknown[];
  messageRowCandidates: unknown[];
  scrollCandidates: unknown[];
  iframes: unknown[];
}

interface RawMessageInspectResult {
  page: PageInfo;
  snapshot: string;
  bodyCandidates: unknown[];
  headerCandidates: unknown[];
  attachmentCandidates: unknown[];
  iframes: unknown[];
}

interface RawInboxFolderScan extends Omit<InboxFolderListResult, 'complete'> {
  collapsed: { path: string; rect: { x: number; y: number } } | null;
  hasCollapsed: boolean;
}

function actionFailure(opened: MessageOpenResult): MessageActionResult | null {
  if (opened.matchCount === 1 && opened.message) return null;
  return {
    matchCount: opened.matchCount,
    status: opened.matchCount > 1
      ? 'message_ambiguous'
      : opened.matchCount === 0
        ? 'message_not_found'
        : 'reading_pane_not_ready',
    performed: false,
    verified: false,
  };
}

export class EgoLiteBackend implements BrowserBackend {
  readonly name: BrowserBackendName = 'ego-lite';

  constructor(protected readonly runner: BrowserScriptRunner = new EgoRunner()) {}

  async close(): Promise<void> {
    await this.runner.close?.();
  }

  private async runComposeAction(
    options: ComposeOptions | ForwardOptions,
    control: '新建' | '转发',
    matchCount = 1,
    preservedAttachmentFilenames: string[] = [],
  ): Promise<ComposeActionResult> {
    const fieldTarget = (label: string): string => String.raw`(() => {
  const label = ${JSON.stringify(label)};
  const inViewport = el => {
    const rect = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight
      && style.display !== 'none' && style.visibility !== 'hidden';
  };
  const fields = Array.from(document.querySelectorAll('[contenteditable="true"][aria-label]')).filter(inViewport)
    .filter(el => (el.getAttribute('aria-label') || '').normalize('NFKC').trim() === label);
  if (fields.length !== 1) return { count: fields.length, rect: null };
  const rect = fields[0].getBoundingClientRect();
  return { count: 1, rect: { x: rect.x + Math.min(20, rect.width / 2), y: rect.y + Math.min(20, rect.height / 2) } };
})()`;
    const recipientFocus = (label: string): string => String.raw`(() => {
  const label = ${JSON.stringify(label)};
  const fields = Array.from(document.querySelectorAll('[contenteditable="true"][aria-label]'))
    .filter(el => { const rect=el.getBoundingClientRect(); return rect.width>0&&rect.height>0&&rect.bottom>0&&rect.top<innerHeight; })
    .filter(el => (el.getAttribute('aria-label') || '').normalize('NFKC').trim() === label);
  if (fields.length !== 1) return { count: fields.length, focused: false };
  const field = fields[0];
  field.focus();
  const selection = window.getSelection();
  const range = document.createRange();
  range.selectNodeContents(field);
  range.collapse(false);
  selection.removeAllRanges();
  selection.addRange(range);
  return { count: 1, focused: document.activeElement === field };
})()`;
    const recipientSteps = (label: string, values: string[]): string => values.map((value, index) => `
if (!fieldFailure) {
  const field_${label.length}_${index} = await js(${JSON.stringify(recipientFocus(label))});
  if (field_${label.length}_${index}.count !== 1 || !field_${label.length}_${index}.focused) fieldFailure = true;
  else {
    await cdp('Input.insertText', { text: ${JSON.stringify(value)} });
    await pressKey('Enter');
    await wait(0.25);
  }
}`).join('');
    const bccReveal = options.bcc.length ? `
const bccButton = await js(${JSON.stringify(String.raw`(() => {
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const buttons = Array.from(document.querySelectorAll('button,[role="button"]')).filter(inViewport)
    .filter(el => /^(密件抄送|密送|bcc)$/i.test(clean(el.textContent) || clean(el.getAttribute('aria-label'))));
  if (buttons.length !== 1) return { count: buttons.length, rect: null };
  const rect = buttons[0].getBoundingClientRect(); return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`)});
if (bccButton.count === 1 && bccButton.rect) { await click(bccButton.rect); await wait(0.2); }
else fieldFailure = true;
` : '';
    const attachmentSteps = options.attachments.map(path => `
if (!fieldFailure) {
  await uploadFile('input[type="file"][data-testid="local-computer-filein"]:not([accept])', ${JSON.stringify(path)});
  await wait(0.5);
}
`).join('');
    const expectedFilenames = [
      ...preservedAttachmentFilenames,
      ...options.attachments.map(path => basename(path)),
    ];
    const editorTargetScript = fieldTarget('邮件正文');
    const verificationScript = String.raw`(() => {
  const expected = ${JSON.stringify({ ...options, attachmentFilenames: expectedFilenames })};
  const clean = value => (value || '').normalize('NFKC').replace(/[\u200b\ufeff]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const one = selector => Array.from(document.querySelectorAll(selector)).filter(inViewport);
  const subjectFields = one('input[aria-label="主题"], input[aria-label="Subject"]');
  const bodyFields = one('[contenteditable="true"][role="textbox"][aria-label="邮件正文"], [contenteditable="true"][role="textbox"][aria-label="Message body"]');
  let composerRoot = bodyFields[0] || null;
  for (let depth = 0; composerRoot && depth < 12; depth += 1, composerRoot = composerRoot.parentElement) {
    if (composerRoot.querySelector('input[aria-label="主题"],input[aria-label="Subject"]')
      && Array.from(composerRoot.querySelectorAll('button[aria-label]')).some(button => /^(发送|send)$/i.test((button.getAttribute('aria-label') || '').trim()))) break;
  }
  const recipientOk = (label, values) => {
    if (!values.length) return true;
    const fields = one('[contenteditable="true"][aria-label="' + label + '"]');
    if (fields.length !== 1) return false;
    const container = fields[0].parentElement && fields[0].parentElement.parentElement || fields[0];
    const text = clean(container.innerText || container.textContent);
    return values.every(value => text.includes(clean(value)) || clean(fields[0].getAttribute('aria-label')).includes(clean(value)));
  };
  const pageText = clean(composerRoot && composerRoot.innerText);
  return {
    subject: subjectFields.length === 1 && clean(subjectFields[0].value) === clean(expected.subject || subjectFields[0].value),
    content: bodyFields.length === 1 && clean(bodyFields[0].innerText || bodyFields[0].textContent).includes(clean(expected.content)),
    recipients: recipientOk('收件人', expected.to) && recipientOk('抄送', expected.cc) && recipientOk('密件抄送', expected.bcc),
    attachments: expected.attachmentFilenames.every(filename => pageText.includes(clean(filename))),
    attachmentCount: expected.attachmentFilenames.length
  };
})()`;
    const start = control === '新建' ? bootstrap() : resumeTaskSpace();
    const controlMatcher = control === '新建' ? '^(新建|new mail)$' : '^(转发|forward)$';
    const failureStatus = control === '新建' ? 'compose_control_not_found' : 'compose_control_not_found';
    const script = `${start}
const controlTarget = await js(${JSON.stringify(`(() => {
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const matcher = /${controlMatcher}/i;
  const buttons = Array.from(document.querySelectorAll('button[aria-label]')).filter(inViewport)
    .filter(el => matcher.test((el.getAttribute('aria-label') || '').normalize('NFKC').trim()));
  if (buttons.length !== 1) return { count: buttons.length, rect: null };
  const rect = buttons[0].getBoundingClientRect(); return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`)});
if (controlTarget.count !== 1 || !controlTarget.rect) {
  ${markedResult(`{ status: controlTarget.count > 1 ? 'compose_control_ambiguous' : '${failureStatus}', performed: false, verified: false, draft: ${JSON.stringify(options.draft)}, handedOff: false, attachmentCount: 0, matchCount: ${JSON.stringify(matchCount)} }`)}
} else {
  await click(controlTarget.rect);
  let editor = null;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    await wait(0.25);
    editor = await js(${JSON.stringify(editorTargetScript)});
    if (editor.count === 1 && editor.rect) break;
  }
  if (!editor || editor.count !== 1 || !editor.rect) {
    const handoff = await handOffTaskSpace(task.id);
    ${markedResult(`{ status: 'editor_not_ready', performed: false, verified: false, draft: ${JSON.stringify(options.draft)}, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), attachmentCount: 0, matchCount: ${JSON.stringify(matchCount)} }`)}
  } else {
    let fieldFailure = false;
    ${bccReveal}
    ${recipientSteps('收件人', options.to)}
    ${recipientSteps('抄送', options.cc)}
    ${recipientSteps('密件抄送', options.bcc)}
    if (!fieldFailure && ${JSON.stringify('subject' in options)}) {
      await fillInput('input[aria-label="主题"], input[aria-label="Subject"]', ${JSON.stringify('subject' in options ? options.subject : '')});
    }
    if (!fieldFailure) {
      editor = await js(${JSON.stringify(editorTargetScript)});
      if (!editor || !editor.rect) fieldFailure = true;
      else {
        await click(editor.rect);
        await cdp('Input.insertText', { text: ${JSON.stringify(options.content)} });
      }
    }
    ${attachmentSteps}
    await wait(0.5);
    if (fieldFailure) {
      const handoff = await handOffTaskSpace(task.id);
      ${markedResult(`{ status: 'field_not_ready', performed: true, verified: false, draft: ${JSON.stringify(options.draft)}, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), attachmentCount: 0, matchCount: ${JSON.stringify(matchCount)} }`)}
    } else {
      const verified = await js(${JSON.stringify(verificationScript)});
      const failure = !verified.subject || !verified.content
        ? 'content_not_verified'
        : !verified.recipients
          ? 'recipient_not_verified'
          : !verified.attachments
            ? 'attachment_not_verified'
            : null;
      if (failure) {
        const handoff = await handOffTaskSpace(task.id);
        ${markedResult(`{ status: failure, performed: true, verified: false, draft: ${JSON.stringify(options.draft)}, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), attachmentCount: verified.attachmentCount, matchCount: ${JSON.stringify(matchCount)} }`)}
      } else if (${JSON.stringify(options.draft)}) {
        const saveButtons = await js(${JSON.stringify(String.raw`(() => {
          const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
          const buttons = Array.from(document.querySelectorAll('button[aria-label]')).filter(inViewport).filter(el => /^(保存草稿|save draft)$/i.test((el.getAttribute('aria-label') || '').trim()));
          if (buttons.length !== 1) return { count: buttons.length, rect: null }; const rect = buttons[0].getBoundingClientRect(); return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
        })()`)});
        if (saveButtons.rect) { await click(saveButtons.rect); await wait(0.5); }
        const handoff = await handOffTaskSpace(task.id);
        ${markedResult(`{ status: 'draft_ready', performed: true, verified: true, draft: true, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), attachmentCount: verified.attachmentCount, matchCount: ${JSON.stringify(matchCount)} }`)}
      } else {
        const sendTarget = await js(${JSON.stringify(String.raw`(() => {
          const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
          const buttons = Array.from(document.querySelectorAll('button[aria-label]')).filter(inViewport).filter(el => /^(发送|send)$/i.test((el.getAttribute('aria-label') || '').trim()));
          if (buttons.length !== 1) return { count: buttons.length, rect: null }; const rect = buttons[0].getBoundingClientRect(); return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
        })()`)});
        if (!sendTarget.rect) {
          const handoff = await handOffTaskSpace(task.id);
          ${markedResult(`{ status: 'send_control_not_found', performed: true, verified: false, draft: false, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), attachmentCount: verified.attachmentCount, matchCount: ${JSON.stringify(matchCount)} }`)}
        } else {
          await click(sendTarget.rect);
          let closed = false;
          for (let attempt = 0; attempt < 30; attempt += 1) {
            await wait(0.3);
            closed = await js(${JSON.stringify(String.raw`(() => {
              const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
              return !Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]')).filter(inViewport).some(el => /^(邮件正文|message body)$/i.test((el.getAttribute('aria-label') || '').trim()));
            })()`)});
            if (closed) break;
          }
          ${markedResult(`{ status: closed ? 'sent' : 'send_not_verified', performed: true, verified: closed, draft: false, handedOff: false, attachmentCount: verified.attachmentCount, matchCount: ${JSON.stringify(matchCount)} }`)}
        }
      }
    }
  }
}
`;
    return (await this.runner.run<ComposeActionResult>(script, 120_000)).value;
  }

  async status(): Promise<BrowserStatus> {
    const script = `${bootstrap()}
const page = await pageInfo();
const snapshot = page && page.dialog ? '' : await snapshotText();
${markedResult(`{
  connected: true,
  taskSpaceId: task && task.id != null ? task.id : null,
  url: page && page.url ? page.url : null,
  title: page && page.title ? page.title : null,
  page,
  snapshot
}`)}
`;
    return (await this.runner.run<BrowserStatus>(script, 40_000)).value;
  }

  async handoffForLogin(): Promise<LoginHandoffResult> {
    const script = `
const task = await useOrCreateTaskSpace(${JSON.stringify(TASK_SPACE)});
let tab = await ensureRealTab();
if (!tab) tab = await openOrReuseTab(${JSON.stringify(OUTLOOK_URL)}, { wait: true, timeout: 30 });
const page = await pageInfo();
const handoff = await handOffTaskSpace(task.id);
${markedResult(`{
  taskSpaceId: task && task.id != null ? task.id : null,
  url: page && page.url ? page.url : null,
  handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned')
}`)}
`;
    return (await this.runner.run<LoginHandoffResult>(script, 40_000)).value;
  }

  async snapshot(): Promise<string> {
    const script = `${bootstrap()}
const value = await snapshotText();
${markedResult('value')}
`;
    return (await this.runner.run<string>(script, 40_000)).value;
  }

  async eval<T>(script: string): Promise<T> {
    const body = `${bootstrap()}
const value = await js(${JSON.stringify(script)});
${markedResult('value')}
`;
    return (await this.runner.run<T>(body, 40_000)).value;
  }

  async click(refOrLocator: string): Promise<void> {
    const script = `${bootstrap()}
await click(${JSON.stringify(refOrLocator)});
${markedResult('null')}
`;
    await this.runner.run<null>(script, 45_000);
  }

  async clickAndWait(refOrLocator: string, waitMs: number): Promise<void> {
    const script = `${bootstrap()}
await click(${JSON.stringify(refOrLocator)});
await wait(${JSON.stringify(waitMs / 1000)});
${markedResult('null')}
`;
    await this.runner.run<null>(script, Math.max(45_000, waitMs + 35_000));
  }

  async fill(refOrLocator: string, text: string): Promise<void> {
    const script = `${bootstrap()}
await fillInput(${JSON.stringify(refOrLocator)}, ${JSON.stringify(text)});
${markedResult('null')}
`;
    await this.runner.run<null>(script, 45_000);
  }

  async fillAndPress(refOrLocator: string, text: string, key: string, waitMs: number): Promise<void> {
    const script = `${bootstrap()}
await fillInput(${JSON.stringify(refOrLocator)}, ${JSON.stringify(text)});
await pressKey(${JSON.stringify(key)});
await wait(${JSON.stringify(waitMs / 1000)});
${markedResult('null')}
`;
    await this.runner.run<null>(script, Math.max(45_000, waitMs + 35_000));
  }

  async press(key: string): Promise<void> {
    const script = `${bootstrap()}
await pressKey(${JSON.stringify(key)});
${markedResult('null')}
`;
    await this.runner.run<null>(script, 45_000);
  }

  async scrollBy(_x: number, y: number): Promise<void> {
    const script = `${bootstrap()}
await scrollBy(${JSON.stringify(y)});
${markedResult('null')}
`;
    await this.runner.run<null>(script, 45_000);
  }

  async wheel(refOrLocator: string, y: number, steps = 1): Promise<void> {
    const script = `${bootstrap()}
await hover(${JSON.stringify(refOrLocator)});
for (let step = 0; step < ${JSON.stringify(steps)}; step += 1) {
  await scroll({ dy: ${JSON.stringify(y)} });
  if (step + 1 < ${JSON.stringify(steps)}) await wait(0.1);
}
${markedResult('null')}
`;
    await this.runner.run<null>(script, 45_000);
  }

  async wheelAndEval<T>(
    refOrLocator: string,
    y: number,
    steps: number,
    waitMs: number,
    script: string,
  ): Promise<T> {
    const targetScript = `(() => {
  const element = document.querySelector(${JSON.stringify(refOrLocator)});
  if (!element) return null;
  const rect = element.getBoundingClientRect();
  return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
})()`;
    const body = `${bootstrap()}
let wheelTarget = null;
for (let attempt = 0; attempt < 20; attempt += 1) {
  wheelTarget = await js(${JSON.stringify(targetScript)});
  if (wheelTarget) break;
  await wait(0.3);
}
if (!wheelTarget) throw new Error('Mail list wheel target not found');
await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: wheelTarget.x, y: wheelTarget.y });
for (let step = 0; step < ${JSON.stringify(steps)}; step += 1) {
  await cdp('Input.dispatchMouseEvent', {
    type: 'mouseWheel',
    x: wheelTarget.x,
    y: wheelTarget.y,
    deltaX: 0,
    deltaY: ${JSON.stringify(y)}
  });
  if (step + 1 < ${JSON.stringify(steps)}) await wait(0.1);
}
await wait(${JSON.stringify(waitMs / 1000)});
const value = await js(${JSON.stringify(script)});
${markedResult('value')}
`;
    return (await this.runner.run<T>(body, Math.max(45_000, waitMs + 35_000))).value;
  }

  async wait(ms: number): Promise<void> {
    const seconds = ms / 1000;
    const script = `${bootstrap()}
await wait(${JSON.stringify(seconds)});
${markedResult('null')}
`;
    await this.runner.run<null>(script, Math.max(40_000, ms + 35_000));
  }

  async inspect(): Promise<Omit<InspectResult, 'backend' | 'capturedAt' | 'state'>> {
    const script = `${bootstrap()}
const page = await pageInfo();
const snapshot = page && page.dialog ? '' : await snapshotText();
const domInventory = page && page.dialog ? [] : await js(${JSON.stringify(PAGE_DOM_INVENTORY)});
const listCandidates = page && page.dialog ? [] : await js(${JSON.stringify(LIST_CANDIDATES)});
const messageRowCandidates = page && page.dialog ? [] : await js(${JSON.stringify(MESSAGE_ROW_CANDIDATES)});
const scrollCandidates = page && page.dialog ? [] : await js(${JSON.stringify(SCROLL_CANDIDATES)});
const iframes = page && page.dialog ? [] : await js(${JSON.stringify(IFRAME_CANDIDATES)});
${markedResult(`{
  taskSpaceId: task && task.id != null ? task.id : null,
  page,
  snapshot,
  domInventory,
  listCandidates,
  messageRowCandidates,
  scrollCandidates,
  iframes
}`)}
`;
    const value = (await this.runner.run<RawInspectResult>(script, 60_000)).value;
    return {
      page: value.page,
      snapshot: value.snapshot,
      domInventory: value.domInventory,
      listCandidates: value.listCandidates,
      messageRowCandidates: value.messageRowCandidates,
      scrollCandidates: value.scrollCandidates,
      iframes: value.iframes,
    };
  }

  async inspectMessage(): Promise<Omit<MessageInspectResult, 'backend' | 'capturedAt' | 'state'>> {
    const script = `${bootstrap()}
const page = await pageInfo();
const snapshot = page && page.dialog ? '' : await snapshotText();
const bodyCandidates = page && page.dialog ? [] : await js(${JSON.stringify(MESSAGE_BODY_CANDIDATES)});
const headerCandidates = page && page.dialog ? [] : await js(${JSON.stringify(MESSAGE_HEADER_CANDIDATES)});
const attachmentCandidates = page && page.dialog ? [] : await js(${JSON.stringify(ATTACHMENT_CANDIDATES)});
const iframes = page && page.dialog ? [] : await js(${JSON.stringify(IFRAME_CANDIDATES)});
${markedResult('{ page, snapshot, bodyCandidates, headerCandidates, attachmentCandidates, iframes }')}
`;
    return (await this.runner.run<RawMessageInspectResult>(script, 60_000)).value;
  }

  async listInboxFolders(): Promise<InboxFolderListResult> {
    const script = `${bootstrap()}
let scan = null;
let complete = false;
let previousCollapsedPath = null;
for (let attempt = 0; attempt < 100; attempt += 1) {
  scan = await js(${JSON.stringify(INBOX_FOLDER_SCAN_SCRIPT)});
  if (scan.accountCount !== 1 || scan.inboxCount !== 1) break;
  if (!scan.collapsed) {
    complete = !scan.hasCollapsed;
    break;
  }
  if (scan.collapsed.path === previousCollapsedPath) break;
  previousCollapsedPath = scan.collapsed.path;
  await click(scan.collapsed.rect);
  await wait(0.8);
}
${markedResult(`{
  accountCount: scan ? scan.accountCount : 0,
  inboxCount: scan ? scan.inboxCount : 0,
  complete,
  folders: scan ? scan.folders : []
}`)}
`;
    return (await this.runner.run<InboxFolderListResult>(script, 110_000)).value;
  }

  async selectInboxFolder(directory: string | null): Promise<FolderSelectionResult> {
    const folders = await this.listInboxFolders();
    if (!folders.complete || folders.accountCount !== 1 || folders.inboxCount !== 1) {
      return { count: 0, selected: false, folder: null };
    }
    const targetScript = buildInboxFolderTargetScript(directory);
    const script = `${bootstrap()}
let target = await js(${JSON.stringify(targetScript)});
if (target.count !== 1 || !target.rect || !target.folder) {
  ${markedResult("{ count: target.count, selected: false, folder: null }")}
} else if (target.selected) {
  let listReady = false;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    const listCount = await js(${JSON.stringify(String.raw`(() => Array.from(document.querySelectorAll('[role="listbox"]')).filter(el => /邮件列表|mail list/i.test(el.getAttribute('aria-label') || '')).length)()`)});
    if (listCount === 1) { listReady = true; break; }
    await wait(0.3);
  }
  ${markedResult("{ count: 1, selected: listReady, folder: target.folder }")}
} else {
  await click(target.rect);
  for (let attempt = 0; attempt < 15; attempt += 1) {
    await wait(0.3);
    target = await js(${JSON.stringify(targetScript)});
    if (target.count === 1 && target.selected) break;
  }
  let listReady = false;
  if (target.count === 1 && target.selected) {
    for (let attempt = 0; attempt < 20; attempt += 1) {
      const listCount = await js(${JSON.stringify(String.raw`(() => Array.from(document.querySelectorAll('[role="listbox"]')).filter(el => /邮件列表|mail list/i.test(el.getAttribute('aria-label') || '')).length)()`)});
      if (listCount === 1) { listReady = true; break; }
      await wait(0.3);
    }
  }
  ${markedResult("{ count: target.count, selected: target.count === 1 && target.selected && listReady, folder: target.folder }")}
}
`;
    return (await this.runner.run<FolderSelectionResult>(script, 60_000)).value;
  }

  async openAndExtractMessage(locator: MessageLocator): Promise<MessageOpenResult> {
    const matchScript = buildMessageRowMatchScript(locator);
    const extractScript = buildMessageExtractScript(locator);
    const searchSelector = 'input[role="combobox"][aria-label^="搜索"], input[role="combobox"][aria-label^="Search"]';
    const exitSearchSelector = 'button[aria-label="退出搜索"], button[aria-label="Exit search"]';
    const script = `${bootstrap()}
let searchReady = false;
for (let attempt = 0; attempt < 20; attempt += 1) {
  const searchCount = await js(${JSON.stringify(`(() => document.querySelectorAll(${JSON.stringify(searchSelector)}).length)()`)});
  if (searchCount === 1) { searchReady = true; break; }
  await wait(0.3);
}
if (!searchReady) throw new Error('Outlook search input did not become ready');
const activeSearch = await js(${JSON.stringify(`(() => {
  const input = document.querySelector(${JSON.stringify(searchSelector)});
  return Boolean(input && input.value);
})()`)});
if (activeSearch) {
  await click(${JSON.stringify(exitSearchSelector)});
  await wait(0.8);
}
await fillInput(${JSON.stringify(searchSelector)}, ${JSON.stringify(locator.subject)});
await pressKey('Enter');
let initialScan = null;
for (let attempt = 0; attempt < 20; attempt += 1) {
  const scan = await js(${JSON.stringify(matchScript)});
  if (scan.listCount === 1 && scan.matches.length > 0) {
    initialScan = scan;
    break;
  }
  await wait(0.5);
}

const seenMatches = new Map();
const previewMatches = new Map();
const matchPages = new Map();
let previousFirstKey = null;
for (let pageIndex = 0; pageIndex < 6; pageIndex += 1) {
  const scan = pageIndex === 0 && initialScan ? initialScan : await js(${JSON.stringify(matchScript)});
  if (scan.listCount !== 1) {
    ${markedResult("{ matchCount: 0, message: null }")}
    break;
  }
  if (pageIndex > 0 && scan.firstKey === previousFirstKey) break;
  previousFirstKey = scan.firstKey;
  for (const match of scan.matches) {
    if (!seenMatches.has(match.logicalKey)) {
      seenMatches.set(match.logicalKey, match);
      matchPages.set(match.logicalKey, pageIndex);
    }
    if (match.previewMatches) previewMatches.set(match.logicalKey, match);
  }
  if (!scan.wheelTarget || pageIndex === 5) break;
  await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: scan.wheelTarget.x, y: scan.wheelTarget.y });
  await cdp('Input.dispatchMouseEvent', {
    type: 'mouseWheel', x: scan.wheelTarget.x, y: scan.wheelTarget.y, deltaX: 0, deltaY: 900
  });
  await wait(0.5);
}

const selectedMatches = seenMatches.size === 1 ? seenMatches : (previewMatches.size === 1 ? previewMatches : null);
if (!selectedMatches) {
  ${markedResult("{ matchCount: seenMatches.size, message: null }")}
} else {
  const uniqueKey = Array.from(selectedMatches.keys())[0];
  const selectedIdentityKey = selectedMatches.get(uniqueKey).identityKey;
  const matchPage = matchPages.get(uniqueKey);
  let previousTopKey = null;
  for (let attempt = 0; attempt < 6; attempt += 1) {
    const scan = await js(${JSON.stringify(matchScript)});
    if (!scan.wheelTarget || scan.firstKey === previousTopKey) break;
    previousTopKey = scan.firstKey;
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: scan.wheelTarget.x, y: scan.wheelTarget.y });
    await cdp('Input.dispatchMouseEvent', {
      type: 'mouseWheel', x: scan.wheelTarget.x, y: scan.wheelTarget.y, deltaX: 0, deltaY: -900
    });
    await wait(0.4);
  }
  for (let pageIndex = 0; pageIndex < matchPage; pageIndex += 1) {
    const scan = await js(${JSON.stringify(matchScript)});
    if (!scan.wheelTarget) break;
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: scan.wheelTarget.x, y: scan.wheelTarget.y });
    await cdp('Input.dispatchMouseEvent', {
      type: 'mouseWheel', x: scan.wheelTarget.x, y: scan.wheelTarget.y, deltaX: 0, deltaY: 900
    });
    await wait(0.5);
  }
  let selected = null;
  for (let attempt = 0; attempt < 10; attempt += 1) {
    const finalScan = await js(${JSON.stringify(matchScript)});
    selected = finalScan.matches.find(match => match.identityKey === selectedIdentityKey)
      || (finalScan.matches.length === 1 ? finalScan.matches[0] : null);
    if (selected) break;
    await wait(0.3);
  }
  if (!selected) {
    ${markedResult("{ matchCount: 0, message: null }")}
  } else {
    const x = selected.rect.x + selected.rect.width / 2;
    const y = selected.rect.y + selected.rect.height / 2;
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y });
    await cdp('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: 1 });
    await cdp('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: 1 });
    let message = null;
    for (let attempt = 0; attempt < 15; attempt += 1) {
      await wait(0.3);
      message = await js(${JSON.stringify(extractScript)});
      if (message && message.subject && message.bodyText
        && (!${JSON.stringify(locator.hasAttachments)} || message.attachments.length > 0)) break;
    }
    ${markedResult("{ matchCount: 1, message }")}
  }
}
`;
    return (await this.runner.run<MessageOpenResult>(script, 75_000)).value;
  }

  async deleteMessage(locator: MessageLocator): Promise<MessageActionResult> {
    const opened = await this.openAndExtractMessage(locator);
    const failure = actionFailure(opened);
    if (failure) return failure;

    const script = `${resumeTaskSpace()}
const control = await js(${JSON.stringify(String.raw`(() => {
  const visible = el => {
    const rect = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
  };
  const buttons = Array.from(document.querySelectorAll('button[aria-label]')).filter(visible)
    .filter(el => /^(删除|delete)$/i.test((el.getAttribute('aria-label') || '').trim()));
  if (buttons.length !== 1) return { count: buttons.length, rect: null };
  const rect = buttons[0].getBoundingClientRect();
  return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`)});
if (!control.rect) {
  ${markedResult("{ matchCount: 1, status: 'control_not_found', performed: false, verified: false }")}
} else {
  await click(control.rect);
  await wait(1);
  const verified = await js(${JSON.stringify(`(() => {
    const expected = ${JSON.stringify(locator.subject)}.normalize('NFKC').replace(/\\s+/g, ' ').trim();
    const clean = value => (value || '').normalize('NFKC').replace(/\\s+/g, ' ').trim();
    const pane = document.querySelector('[role="main"][aria-label="阅读窗格"], [role="main"][aria-label="Reading pane"]');
    if (!pane) return true;
    return !Array.from(pane.querySelectorAll('span[role="heading"]')).some(el => clean(el.textContent) === expected);
  })()`)});
  ${markedResult("{ matchCount: 1, status: 'performed', performed: true, verified }")}
}
`;
    return (await this.runner.run<MessageActionResult>(script, 50_000)).value;
  }

  async moveMessage(locator: MessageLocator, folder: string): Promise<MessageActionResult> {
    const opened = await this.openAndExtractMessage(locator);
    const failure = actionFailure(opened);
    if (failure) return failure;

    const script = `${resumeTaskSpace()}
const moveControl = await js(${JSON.stringify(String.raw`(() => {
  const visible = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0; };
  const buttons = Array.from(document.querySelectorAll('button[aria-label]')).filter(visible)
    .filter(el => /^(移动|move)$/i.test((el.getAttribute('aria-label') || '').trim()));
  if (buttons.length !== 1) return { count: buttons.length, rect: null };
  const rect = buttons[0].getBoundingClientRect();
  return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`)});
if (!moveControl.rect) {
  ${markedResult("{ matchCount: 1, status: 'control_not_found', performed: false, verified: false }")}
} else {
  await click(moveControl.rect);
  await wait(0.5);
  const folderInput = '[role="menu"][aria-label="移动"] input[placeholder="搜索文件夹"], [role="menu"][aria-label="Move"] input[placeholder*="Search" i]';
  await fillInput(folderInput, ${JSON.stringify(folder)});
  await wait(0.8);
  const target = await js(${JSON.stringify(`(() => {
    const expected = ${JSON.stringify(folder)}.normalize('NFKC').trim().toLowerCase();
    const visible = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0; };
    const menus = Array.from(document.querySelectorAll('[role="menu"]')).filter(visible)
      .filter(el => /移动|move/i.test(el.getAttribute('aria-label') || ''));
    const items = menus.flatMap(menu => Array.from(menu.querySelectorAll('[role="menuitem"]'))).filter(visible)
      .filter(el => (el.getAttribute('aria-label') || '').normalize('NFKC').trim().toLowerCase() === expected);
    if (items.length !== 1) return { count: items.length, rect: null };
    const rect = items[0].getBoundingClientRect();
    return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
  })()`)});
  if (!target.rect) {
    await pressKey('Escape');
    ${markedResult("{ matchCount: 1, status: target.count > 1 ? 'folder_ambiguous' : 'folder_not_found', performed: false, verified: false, folderMatches: target.count }")}
  } else {
    await click(target.rect);
    await wait(1);
    const verified = await js(${JSON.stringify(String.raw`(() => {
      const menus = Array.from(document.querySelectorAll('[role="menu"]'));
      return !menus.some(el => /移动|move/i.test(el.getAttribute('aria-label') || '') && el.getBoundingClientRect().width > 0);
    })()`)});
    ${markedResult("{ matchCount: 1, status: 'performed', performed: true, verified }")}
  }
}
`;
    return (await this.runner.run<MessageActionResult>(script, 55_000)).value;
  }

  async replyMessage(
    locator: MessageLocator,
    content: string,
    draft: boolean,
    replyAll: boolean,
  ): Promise<ReplyActionResult> {
    const opened = await this.openAndExtractMessage(locator);
    if (opened.matchCount !== 1 || !opened.message) {
      return {
        matchCount: opened.matchCount,
        status: opened.matchCount > 1
          ? 'message_ambiguous'
          : opened.matchCount === 0
            ? 'message_not_found'
            : 'reading_pane_not_ready',
        performed: false,
        verified: false,
        draft,
        replyAll,
      };
    }

    const replyTargetScript = String.raw`(() => {
  const replyAll = ${JSON.stringify(replyAll)};
  const matcher = replyAll ? /^(全部答复|reply all)$/i : /^(答复|reply)$/i;
  const pane = document.querySelector('[role="main"][aria-label="阅读窗格"], [role="main"][aria-label="Reading pane"]');
  const buttons = pane ? Array.from(pane.querySelectorAll('button[aria-label]'))
    .filter(el => matcher.test((el.getAttribute('aria-label') || '').normalize('NFKC').trim())) : [];
  if (buttons.length !== 1) return { count: buttons.length, rect: null };
  buttons[0].scrollIntoView({ block: 'center', inline: 'center' });
  const rect = buttons[0].getBoundingClientRect();
  const inViewport = rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight;
  return { count: 1, rect: inViewport ? { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } : null };
})()`;
    const editorScript = String.raw`(() => {
  const inViewport = el => {
    const rect = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight
      && style.display !== 'none' && style.visibility !== 'hidden';
  };
  const editors = Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]'))
    .filter(inViewport)
    .filter(el => /^(邮件正文|message body)$/i.test((el.getAttribute('aria-label') || '').normalize('NFKC').trim()));
  if (editors.length !== 1) return { count: editors.length, rect: null };
  const rect = editors[0].getBoundingClientRect();
  return { count: 1, rect: { x: rect.x + Math.min(20, rect.width / 2), y: rect.y + Math.min(20, rect.height / 2) } };
})()`;
    const contentVerifiedScript = String.raw`(() => {
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const editors = Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]'))
    .filter(inViewport)
    .filter(el => /^(邮件正文|message body)$/i.test((el.getAttribute('aria-label') || '').normalize('NFKC').trim()));
  if (editors.length !== 1) return false;
  const clean = value => (value || '').normalize('NFKC').replace(/\r\n/g, '\n').replace(/\u00a0/g, ' ').trim();
  return clean(editors[0].innerText || editors[0].textContent).includes(clean(${JSON.stringify(content)}));
})()`;
    const sendTargetScript = String.raw`(() => {
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const buttons = Array.from(document.querySelectorAll('button[aria-label]')).filter(inViewport)
    .filter(el => /^(发送|send)$/i.test((el.getAttribute('aria-label') || '').normalize('NFKC').trim()));
  if (buttons.length !== 1) return { count: buttons.length, rect: null };
  const rect = buttons[0].getBoundingClientRect();
  return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`;
    const editorClosedScript = String.raw`(() => {
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  return !Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]'))
    .filter(inViewport)
    .some(el => /^(邮件正文|message body)$/i.test((el.getAttribute('aria-label') || '').normalize('NFKC').trim()));
})()`;
    const commonResult = `draft: ${JSON.stringify(draft)}, replyAll: ${JSON.stringify(replyAll)}`;
    const script = `${resumeTaskSpace()}
const replyTarget = await js(${JSON.stringify(replyTargetScript)});
if (replyTarget.count !== 1) {
  ${markedResult(`{ matchCount: 1, status: replyTarget.count > 1 ? 'reply_control_ambiguous' : 'reply_control_not_found', performed: false, verified: false, ${commonResult} }`)}
} else if (!replyTarget.rect) {
  ${markedResult(`{ matchCount: 1, status: 'reply_control_not_found', performed: false, verified: false, ${commonResult} }`)}
} else {
  await click(replyTarget.rect);
  let editor = null;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    await wait(0.3);
    editor = await js(${JSON.stringify(editorScript)});
    if (editor.count === 1 && editor.rect) break;
  }
  if (!editor || editor.count !== 1 || !editor.rect) {
    const handoff = await handOffTaskSpace(task.id);
    ${markedResult(`{ matchCount: 1, status: 'editor_not_ready', performed: false, verified: false, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), ${commonResult} }`)}
  } else {
    await click(editor.rect);
    await cdp('Input.insertText', { text: ${JSON.stringify(content)} });
    await wait(0.5);
    const contentVerified = await js(${JSON.stringify(contentVerifiedScript)});
    if (!contentVerified) {
      const handoff = await handOffTaskSpace(task.id);
      ${markedResult(`{ matchCount: 1, status: 'content_not_verified', performed: true, verified: false, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), ${commonResult} }`)}
    } else if (${JSON.stringify(draft)}) {
      const handoff = await handOffTaskSpace(task.id);
      ${markedResult(`{ matchCount: 1, status: 'draft_ready', performed: true, verified: true, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), ${commonResult} }`)}
    } else {
      const sendTarget = await js(${JSON.stringify(sendTargetScript)});
      if (sendTarget.count !== 1 || !sendTarget.rect) {
        const handoff = await handOffTaskSpace(task.id);
        ${markedResult(`{ matchCount: 1, status: 'send_control_not_found', performed: true, verified: false, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), ${commonResult} }`)}
      } else {
        await click(sendTarget.rect);
        let sent = false;
        for (let attempt = 0; attempt < 30; attempt += 1) {
          await wait(0.3);
          sent = await js(${JSON.stringify(editorClosedScript)});
          if (sent) break;
        }
        if (sent) {
          ${markedResult(`{ matchCount: 1, status: 'sent', performed: true, verified: true, handedOff: false, ${commonResult} }`)}
        } else {
          const handoff = await handOffTaskSpace(task.id);
          ${markedResult(`{ matchCount: 1, status: 'send_not_verified', performed: true, verified: false, handedOff: Boolean(handoff.done || handoff.skipped === 'user-owned'), ${commonResult} }`)}
        }
      }
    }
  }
}
`;
    return (await this.runner.run<ReplyActionResult>(script, 90_000)).value;
  }

  async composeMessage(options: ComposeOptions): Promise<ComposeActionResult> {
    return await this.runComposeAction(options, '新建');
  }

  async forwardMessage(locator: MessageLocator, options: ForwardOptions): Promise<ComposeActionResult> {
    const opened = await this.openAndExtractMessage(locator);
    if (opened.matchCount !== 1 || !opened.message) {
      return {
        status: opened.matchCount > 1
          ? 'message_ambiguous'
          : opened.matchCount === 0
            ? 'message_not_found'
            : 'reading_pane_not_ready',
        performed: false,
        verified: false,
        draft: options.draft,
        handedOff: false,
        attachmentCount: 0,
        matchCount: opened.matchCount,
      };
    }
    return await this.runComposeAction(
      options,
      '转发',
      1,
      opened.message.attachments.map(attachment => attachment.filename),
    );
  }

  async selectSystemFolder(folder: '草稿'): Promise<FolderSelectionResult> {
    const targetScript = String.raw`(() => {
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const matcher = /^(草稿|drafts?)/i;
  const items = Array.from(document.querySelectorAll('[role="treeitem"]')).filter(el => matcher.test(clean(el.getAttribute('title'))));
  if (!items.length) return { count: 0, rect: null, folder: null, selected: false };
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  let visible = items.filter(inViewport);
  if (!visible.length) {
    items[items.length - 1].scrollIntoView({ block: 'center', inline: 'nearest' });
    visible = items.filter(inViewport);
  }
  if (visible.length !== 1) return { count: visible.length, rect: null, folder: null, selected: false };
  const item = visible[0];
  const rect = item.getBoundingClientRect();
  const selected = item.getAttribute('aria-selected') === 'true' || /已选择|selected/i.test(clean(item.textContent));
  return {
    count: 1,
    rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 },
    selected,
    folder: { name: ${JSON.stringify(folder)}, path: ${JSON.stringify(folder)}, level: Number(item.getAttribute('aria-level')) || 2, expanded: null }
  };
})()`;
    const script = `${bootstrap()}
let target = await js(${JSON.stringify(targetScript)});
if (target.count !== 1 || !target.rect || !target.folder) {
  ${markedResult('{ count: target.count, selected: false, folder: null }')}
} else {
  if (!target.selected) { await click(target.rect); await wait(0.8); }
  target = await js(${JSON.stringify(targetScript)});
  const listCount = await js(${JSON.stringify(String.raw`(() => Array.from(document.querySelectorAll('[role="listbox"]')).filter(el => /邮件列表|mail list/i.test(el.getAttribute('aria-label') || '')).length)()`)});
  ${markedResult('{ count: target.count, selected: target.count === 1 && listCount === 1, folder: target.folder }')}
}
`;
    return (await this.runner.run<FolderSelectionResult>(script, 55_000)).value;
  }

  async openDraft(locator: MessageLocator, closeAfterRead = false): Promise<DraftOpenResult> {
    const matchScript = buildMessageRowMatchScript(locator);
    const searchSelector = 'input[role="combobox"][aria-label^="搜索"], input[role="combobox"][aria-label^="Search"]';
    const exitSearchSelector = 'button[aria-label="退出搜索"], button[aria-label="Exit search"]';
    const extractScript = String.raw`(() => {
  const clean = value => (value || '').normalize('NFKC').replace(/[\u200b\ufeff]/g, '').replace(/\s+/g, ' ').trim();
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const fields = label => Array.from(document.querySelectorAll('[contenteditable="true"][aria-label]')).filter(inViewport)
    .filter(el => (el.getAttribute('aria-label') || '').normalize('NFKC').trim() === label);
  const addresses = label => {
    const matches = fields(label);
    if (matches.length !== 1) return [];
    const root = matches[0].parentElement && matches[0].parentElement.parentElement || matches[0];
    const buttons = Array.from(root.querySelectorAll('[role="button"][aria-label],button[aria-label]')).filter(inViewport);
    const values = [root.innerText || root.textContent || '', ...buttons.map(button =>
      (button.getAttribute('aria-label') || '') + ' ' + (button.getAttribute('title') || '')
    )].map(clean).filter(Boolean);
    const emails = Array.from(new Set(values.flatMap(value =>
      Array.from(value.matchAll(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig), match => match[0])
    )));
    if (emails.length) return emails.map(address => ({ name: null, address }));
    return Array.from(new Set(buttons.map(button => clean(button.getAttribute('aria-label'))).filter(Boolean)))
      .map(name => ({ name, address: null }));
  };
  const subjects = Array.from(document.querySelectorAll('input[aria-label="主题"],input[aria-label="Subject"]')).filter(inViewport);
  const bodies = Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label="邮件正文"],[contenteditable="true"][role="textbox"][aria-label="Message body"]')).filter(inViewport);
  if (subjects.length !== 1 || bodies.length !== 1) return null;
  let composerRoot = bodies[0];
  for (let depth = 0; composerRoot && depth < 12; depth += 1, composerRoot = composerRoot.parentElement) {
    if (composerRoot.querySelector('input[aria-label="主题"],input[aria-label="Subject"]')
      && Array.from(composerRoot.querySelectorAll('button[aria-label]')).some(button => /^(发送|send)$/i.test((button.getAttribute('aria-label') || '').trim()))) break;
  }
  const attachmentRoots = Array.from((composerRoot || document).querySelectorAll('[role="listbox"]')).filter(inViewport).filter(el => /文件附件|attachments?/i.test(el.getAttribute('aria-label') || ''));
  const attachments = attachmentRoots.flatMap(root => Array.from(root.querySelectorAll('[role="option"]')).map(option => {
    const filenameElement = Array.from(option.querySelectorAll('[title]')).find(el => /\.[a-z0-9]{2,8}$/i.test(clean(el.getAttribute('title'))));
    const filename = clean(filenameElement && filenameElement.getAttribute('title'));
    const sizeMatch = clean(option.innerText).match(/\b\d+(?:\.\d+)?\s*(?:bytes?|[kmgt]b|字节)\b/i);
    return { filename, sizeText: sizeMatch ? sizeMatch[0] : null };
  })).filter(item => item.filename);
  return {
    to: addresses('收件人'), cc: addresses('抄送'), bcc: addresses('密件抄送'),
    subject: clean(subjects[0].value),
    bodyText: (bodies[0].innerText || '').replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ').trim(),
    attachments
    };
})()`;
    const script = `${bootstrap()}
const currentDraft = await js(${JSON.stringify(extractScript)});
const currentDraftMatches = currentDraft && currentDraft.subject.normalize('NFKC').trim()
  === ${JSON.stringify(locator.subject)}.normalize('NFKC').trim();
if (currentDraftMatches) {
  ${markedResult('{ matchCount: 1, draft: currentDraft, closed: false }')}
} else {
let searchReady = false;
for (let attempt = 0; attempt < 20; attempt += 1) {
  const searchCount = await js(${JSON.stringify(`(() => document.querySelectorAll(${JSON.stringify(searchSelector)}).length)()`)});
  if (searchCount === 1) { searchReady = true; break; }
  await wait(0.3);
}
if (!searchReady) throw new Error('Outlook search input did not become ready');
const activeSearch = await js(${JSON.stringify(`(() => { const input=document.querySelector(${JSON.stringify(searchSelector)}); return Boolean(input && input.value); })()`)});
if (activeSearch) { await click(${JSON.stringify(exitSearchSelector)}); await wait(0.8); }
await fillInput(${JSON.stringify(searchSelector)}, ${JSON.stringify(locator.subject)});
await pressKey('Enter');
await wait(1.5);
let previousTopKey = null;
for (let attempt = 0; attempt < 12; attempt += 1) {
  const scan = await js(${JSON.stringify(matchScript)});
  if (!scan.wheelTarget || scan.firstKey === previousTopKey) break;
  previousTopKey = scan.firstKey;
  await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: scan.wheelTarget.x, y: scan.wheelTarget.y });
  await cdp('Input.dispatchMouseEvent', { type: 'mouseWheel', x: scan.wheelTarget.x, y: scan.wheelTarget.y, deltaX: 0, deltaY: -900 });
  await wait(0.35);
}
const seenMatches = new Map();
const previewMatches = new Map();
const matchPages = new Map();
let previousFirstKey = null;
for (let pageIndex = 0; pageIndex < 20; pageIndex += 1) {
  const scan = await js(${JSON.stringify(matchScript)});
  if (scan.listCount !== 1 || (pageIndex > 0 && scan.firstKey === previousFirstKey)) break;
  previousFirstKey = scan.firstKey;
  for (const match of scan.matches) {
    if (!seenMatches.has(match.logicalKey)) {
      seenMatches.set(match.logicalKey, match);
      matchPages.set(match.logicalKey, pageIndex);
    }
    if (match.previewMatches) previewMatches.set(match.logicalKey, match);
  }
  if (!scan.wheelTarget || pageIndex === 19) break;
  await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: scan.wheelTarget.x, y: scan.wheelTarget.y });
  await cdp('Input.dispatchMouseEvent', { type: 'mouseWheel', x: scan.wheelTarget.x, y: scan.wheelTarget.y, deltaX: 0, deltaY: 900 });
  await wait(0.4);
}
const selectedMatches = seenMatches.size === 1 ? seenMatches : (previewMatches.size === 1 ? previewMatches : null);
if (!selectedMatches) {
  ${markedResult('{ matchCount: seenMatches.size, draft: null }')}
} else {
  const uniqueKey = Array.from(selectedMatches.keys())[0];
  const selectedIdentityKey = selectedMatches.get(uniqueKey).identityKey;
  const matchPage = matchPages.get(uniqueKey);
  previousTopKey = null;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    const scan = await js(${JSON.stringify(matchScript)});
    if (!scan.wheelTarget || scan.firstKey === previousTopKey) break;
    previousTopKey = scan.firstKey;
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: scan.wheelTarget.x, y: scan.wheelTarget.y });
    await cdp('Input.dispatchMouseEvent', { type: 'mouseWheel', x: scan.wheelTarget.x, y: scan.wheelTarget.y, deltaX: 0, deltaY: -900 });
    await wait(0.35);
  }
  for (let pageIndex = 0; pageIndex < matchPage; pageIndex += 1) {
    const scan = await js(${JSON.stringify(matchScript)});
    if (!scan.wheelTarget) break;
    await cdp('Input.dispatchMouseEvent', { type: 'mouseMoved', x: scan.wheelTarget.x, y: scan.wheelTarget.y });
    await cdp('Input.dispatchMouseEvent', { type: 'mouseWheel', x: scan.wheelTarget.x, y: scan.wheelTarget.y, deltaX: 0, deltaY: 900 });
    await wait(0.4);
  }
  const finalScan = await js(${JSON.stringify(matchScript)});
  const target = finalScan.matches.find(match => match.identityKey === selectedIdentityKey)
    || (finalScan.matches.length === 1 ? finalScan.matches[0] : null);
  if (!target) {
    ${markedResult('{ matchCount: 0, draft: null }')}
  } else {
  await click({ x: target.rect.x + target.rect.width / 2, y: target.rect.y + target.rect.height / 2 });
  let draft = null;
  for (let attempt = 0; attempt < 20; attempt += 1) {
    await wait(0.3);
    draft = await js(${JSON.stringify(extractScript)});
    if (draft) break;
  }
  let closed = false;
  if (draft && ${JSON.stringify(closeAfterRead)}) {
    const close = await js(${JSON.stringify(String.raw`(() => { const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight};const bs=Array.from(document.querySelectorAll('button[aria-label]')).filter(visible).filter(el=>/^(关闭|close)$/i.test((el.getAttribute('aria-label')||'').trim()));if(bs.length!==1)return null;const r=bs[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2} })()`)});
    if (close) { await click(close); await wait(0.6); }
    closed = await js(${JSON.stringify(String.raw`(() => !Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]')).some(el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight&&/^(邮件正文|message body)$/i.test((el.getAttribute('aria-label')||'').trim())}))()`)});
  }
  ${markedResult('{ matchCount: 1, draft, closed }')}
  }
}
}
`;
    return (await this.runner.run<DraftOpenResult>(script, 90_000)).value;
  }

  async updateDraft(locator: MessageLocator, options: DraftUpdateOptions): Promise<ComposeActionResult> {
    const opened = await this.openDraft(locator);
    if (opened.matchCount !== 1 || !opened.draft) {
      return {
        status: opened.matchCount > 1 ? 'message_ambiguous' : 'message_not_found',
        performed: false, verified: false, draft: true, handedOff: false,
        attachmentCount: 0, matchCount: opened.matchCount,
      };
    }
    const replaceRecipients = (label: string, values?: string[]): string => values === undefined ? '' : `
const recipient_${label.length} = await js(${JSON.stringify(`(() => { const fields=Array.from(document.querySelectorAll('[contenteditable="true"][aria-label="${label}"]')).filter(el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight});if(fields.length!==1)return false;const field=fields[0];field.focus();const selection=window.getSelection();const range=document.createRange();range.selectNodeContents(field);selection.removeAllRanges();selection.addRange(range);return document.activeElement===field })()`)});
if (!recipient_${label.length}) fieldFailure = true;
else {
  await pressKey('Backspace');
  ${values.map(value => `await cdp('Input.insertText', { text: ${JSON.stringify(value)} }); await pressKey('Enter'); await wait(0.2);`).join('\n')}
}
`;
    const attachments = options.attachments.map(path => `
await uploadFile('input[type="file"][data-testid="local-computer-filein"]:not([accept])', ${JSON.stringify(path)});
await wait(0.5);
`).join('');
    const verificationScript = String.raw`(() => {
  const expected = ${JSON.stringify({ ...options, attachmentFilenames: options.attachments.map(path => basename(path)) })};
  const clean = value => (value || '').normalize('NFKC').replace(/[\u200b\ufeff]/g, '').replace(/\s+/g, ' ').trim().toLowerCase();
  const visible = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const recipientMatches = (label, values) => {
    if (values === undefined) return true;
    const fields = Array.from(document.querySelectorAll('[contenteditable="true"][aria-label="' + label + '"]')).filter(visible);
    if (fields.length !== 1) return false;
    const root = fields[0].parentElement && fields[0].parentElement.parentElement || fields[0];
    const text = clean((root.innerText || root.textContent || '') + ' ' + Array.from(root.querySelectorAll('button')).map(button =>
      (button.getAttribute('aria-label') || '') + ' ' + (button.getAttribute('title') || '')
    ).join(' '));
    return values.length ? values.every(value => text.includes(clean(value))) : !/@/.test(text);
  };
  const subjects = Array.from(document.querySelectorAll('input[aria-label="主题"],input[aria-label="Subject"]')).filter(visible);
  const bodies = Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label="邮件正文"],[contenteditable="true"][role="textbox"][aria-label="Message body"]')).filter(visible);
  let composerRoot = bodies[0] || null;
  for (let depth = 0; composerRoot && depth < 12; depth += 1, composerRoot = composerRoot.parentElement) {
    if (composerRoot.querySelector('input[aria-label="主题"],input[aria-label="Subject"]')
      && Array.from(composerRoot.querySelectorAll('button[aria-label]')).some(button => /^(发送|send)$/i.test((button.getAttribute('aria-label') || '').trim()))) break;
  }
  const pageText = clean(composerRoot && composerRoot.innerText);
  return {
    subject: expected.subject === undefined || subjects.length === 1 && clean(subjects[0].value) === clean(expected.subject),
    content: expected.content === undefined || bodies.length === 1 && clean(bodies[0].innerText || bodies[0].textContent).includes(clean(expected.content)),
    recipients: recipientMatches('收件人', expected.to) && recipientMatches('抄送', expected.cc) && recipientMatches('密件抄送', expected.bcc),
    attachments: expected.attachmentFilenames.every(filename => pageText.includes(clean(filename)))
  };
})()`;
    const script = `${resumeTaskSpace()}
let fieldFailure = false;
${options.bcc !== undefined ? `
const bccExists = await js(${JSON.stringify(String.raw`(() => Array.from(document.querySelectorAll('[contenteditable="true"][aria-label]')).some(el => /^(密件抄送|bcc)$/i.test((el.getAttribute('aria-label') || '').trim()) && el.getBoundingClientRect().width > 0))()`)});
if (!bccExists) {
  const bccButton = await js(${JSON.stringify(String.raw`(() => { const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0};const bs=Array.from(document.querySelectorAll('button,[role="button"]')).filter(visible).filter(el=>/^(密件抄送|密送|bcc)$/i.test((el.textContent||el.getAttribute('aria-label')||'').normalize('NFKC').trim()));if(bs.length!==1)return null;const r=bs[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2} })()`)});
  if (bccButton) { await click(bccButton); await wait(0.2); } else fieldFailure = true;
}
` : ''}
${replaceRecipients('收件人', options.to)}
${replaceRecipients('抄送', options.cc)}
${replaceRecipients('密件抄送', options.bcc)}
${options.subject !== undefined ? `await fillInput('input[aria-label="主题"],input[aria-label="Subject"]', ${JSON.stringify(options.subject)});` : ''}
${options.content !== undefined ? `
const bodyFocused = await js(${JSON.stringify(String.raw`(() => {
  const fields=Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]'))
    .filter(el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight})
    .filter(el=>/^(邮件正文|message body)$/i.test((el.getAttribute('aria-label')||'').trim()));
  if(fields.length!==1)return false;
  const field=fields[0];
  field.focus();
  const selection=window.getSelection();
  const range=document.createRange();
  range.selectNodeContents(field);
  selection.removeAllRanges();
  selection.addRange(range);
  return document.activeElement===field;
})()`)});
if (!bodyFocused) fieldFailure = true;
else {
  await pressKey('Backspace');
  await cdp('Input.insertText', { text: ${JSON.stringify(options.content)} });
}
` : ''}
${attachments}
await wait(1.5);
const save = await js(${JSON.stringify(String.raw`(() => { const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight};const bs=Array.from(document.querySelectorAll('button[aria-label]')).filter(visible).filter(el=>/^(保存草稿|save draft)$/i.test((el.getAttribute('aria-label')||'').trim()));if(bs.length!==1)return null;const r=bs[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2} })()`)});
if (!fieldFailure && save) { await click(save); await wait(2); }
const verification = !fieldFailure && save ? await js(${JSON.stringify(verificationScript)}) : null;
const verified = Boolean(verification && verification.subject && verification.content && verification.recipients && verification.attachments);
let closed = false;
if (verified) {
  const close = await js(${JSON.stringify(String.raw`(() => { const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight};const bs=Array.from(document.querySelectorAll('button[aria-label]')).filter(visible).filter(el=>/^(关闭|close)$/i.test((el.getAttribute('aria-label')||'').trim()));if(bs.length!==1)return null;const r=bs[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2} })()`)});
  if (close) { await click(close); await wait(0.6); }
  closed = await js(${JSON.stringify(String.raw`(() => !Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]')).some(el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight&&/^(邮件正文|message body)$/i.test((el.getAttribute('aria-label')||'').trim())}))()`)});
}
${markedResult(`{
  status: fieldFailure || !save ? 'field_not_ready' : verified ? 'draft_ready' : 'content_not_verified',
  performed: true,
  verified,
  draft: true,
  handedOff: false,
  attachmentCount: ${JSON.stringify(options.attachments.length)},
  matchCount: 1
}`)}
`;
    const result = (await this.runner.run<ComposeActionResult>(script, 100_000)).value;
    if (!result.verified) return result;

    await this.runner.run<null>(`${resumeTaskSpace()}
await cdp('Page.reload', { ignoreCache: true });
await wait(3);
${markedResult('null')}
`, 60_000);

    const persisted = await this.openDraft({
      ...locator,
      subject: options.subject ?? locator.subject,
      receivedAt: null,
      receivedAtText: null,
      preview: options.content ?? locator.preview,
      hasAttachments: options.attachments.length ? true : locator.hasAttachments,
    }, true);
    const normalizedAddresses = (addresses: Array<{ name: string | null; address: string | null }>): string[] =>
      addresses.map(address => (address.address || address.name || '').normalize('NFKC').trim().toLocaleLowerCase('en-US')).filter(Boolean);
    const recipientsMatch = (
      expected: string[] | undefined,
      actual: Array<{ name: string | null; address: string | null }>,
    ): boolean => expected === undefined || (() => {
      const normalizedExpected = expected.map(value => value.normalize('NFKC').trim().toLocaleLowerCase('en-US'));
      const normalizedActual = normalizedAddresses(actual);
      return normalizedExpected.length === normalizedActual.length
        && normalizedExpected.every(value => normalizedActual.includes(value));
    })();
    const draft = persisted.draft;
    const persistedVerified = persisted.matchCount === 1 && Boolean(draft)
      && (options.subject === undefined || draft!.subject.normalize('NFKC').trim() === options.subject.normalize('NFKC').trim())
      && (options.content === undefined || draft!.bodyText.normalize('NFKC').includes(options.content.normalize('NFKC')))
      && recipientsMatch(options.to, draft!.to)
      && recipientsMatch(options.cc, draft!.cc)
      && recipientsMatch(options.bcc, draft!.bcc)
      && options.attachments.every(path => draft!.attachments.some(attachment => attachment.filename === basename(path)));
    return {
      ...result,
      status: persistedVerified ? 'draft_ready' : 'content_not_verified',
      verified: persistedVerified,
    };
  }

  async sendDraft(locator: MessageLocator): Promise<ComposeActionResult> {
    const opened = await this.openDraft(locator);
    if (opened.matchCount !== 1 || !opened.draft) {
      return { status: opened.matchCount > 1 ? 'message_ambiguous' : 'message_not_found', performed: false, verified: false, draft: false, handedOff: false, attachmentCount: 0, matchCount: opened.matchCount };
    }
    const script = `${resumeTaskSpace()}
const send = await js(${JSON.stringify(String.raw`(() => { const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight};const bs=Array.from(document.querySelectorAll('button[aria-label]')).filter(visible).filter(el=>/^(发送|send)$/i.test((el.getAttribute('aria-label')||'').trim()));if(bs.length!==1)return null;const r=bs[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2} })()`)});
if (!send) {
  ${markedResult("{ status: 'send_control_not_found', performed: false, verified: false, draft: false, handedOff: false, attachmentCount: 0, matchCount: 1 }")}
} else {
  await click(send);
  let closed = false;
  for (let attempt = 0; attempt < 30; attempt += 1) {
    await wait(0.3);
    closed = await js(${JSON.stringify(String.raw`(() => !Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]')).some(el => { const r=el.getBoundingClientRect(); return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight&&/^(邮件正文|message body)$/i.test((el.getAttribute('aria-label')||'').trim()); }))()`)});
    if (closed) break;
  }
  ${markedResult("{ status: closed ? 'sent' : 'send_not_verified', performed: true, verified: closed, draft: false, handedOff: false, attachmentCount: 0, matchCount: 1 }")}
}
`;
    return (await this.runner.run<ComposeActionResult>(script, 75_000)).value;
  }

  async discardDraft(locator: MessageLocator): Promise<MessageActionResult> {
    const opened = await this.openDraft(locator);
    if (opened.matchCount !== 1 || !opened.draft) {
      return { matchCount: opened.matchCount, status: opened.matchCount > 1 ? 'message_ambiguous' : 'message_not_found', performed: false, verified: false };
    }
    const script = `${resumeTaskSpace()}
const discard = await js(${JSON.stringify(String.raw`(() => { const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight};const bs=Array.from(document.querySelectorAll('button[aria-label]')).filter(visible).filter(el=>/^(放弃|discard)$/i.test((el.getAttribute('aria-label')||'').trim()));if(bs.length!==1)return null;const r=bs[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2} })()`)});
if (!discard) {
  ${markedResult("{ matchCount: 1, status: 'control_not_found', performed: false, verified: false }")}
} else {
  await click(discard); await wait(0.3);
  const confirm = await js(${JSON.stringify(String.raw`(() => { const clean=v=>(v||'').normalize('NFKC').trim();const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight};const dialogs=Array.from(document.querySelectorAll('[role="dialog"],[role="alertdialog"]')).filter(visible);const bs=dialogs.flatMap(d=>Array.from(d.querySelectorAll('button'))).filter(visible).filter(b=>/^(确定|ok|discard)$/i.test(clean(b.textContent)));if(dialogs.length===1&&bs.length===1){const r=bs[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2}}return null })()`)});
  if (confirm) { await click(confirm); await wait(0.5); }
  const verified = await js(${JSON.stringify(String.raw`(() => !Array.from(document.querySelectorAll('[contenteditable="true"][role="textbox"][aria-label]')).some(el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight&&/^(邮件正文|message body)$/i.test((el.getAttribute('aria-label')||'').trim())}))()`)});
  ${markedResult("{ matchCount: 1, status: verified ? 'performed' : 'control_not_found', performed: true, verified }")}
}
`;
    return (await this.runner.run<MessageActionResult>(script, 65_000)).value;
  }

  private async setBinaryMessageState(
    locator: MessageLocator,
    kind: 'read' | 'flagged',
    desired: boolean,
  ): Promise<MessageStateActionResult> {
    const opened = await this.openAndExtractMessage(locator);
    const failure = actionFailure(opened);
    if (failure) return failure;
    const desiredAction = kind === 'read'
      ? (desired ? /^(标记为已读|mark as read)$/i : /^(标记为未读|mark as unread)$/i)
      : (desired ? /^(标记此邮件|flag this message|flag)$/i : /^(清除标志|取消标记|clear flag|unflag)$/i);
    const oppositeAction = kind === 'read'
      ? (desired ? /^(标记为未读|mark as unread)$/i : /^(标记为已读|mark as read)$/i)
      : (desired ? /^(清除标志|取消标记|clear flag|unflag)$/i : /^(标记此邮件|flag this message|flag)$/i);
    const stateScript = String.raw`(() => {
  const desiredAction = ${desiredAction};
  const oppositeAction = ${oppositeAction};
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const selected = Array.from(document.querySelectorAll('[role="listbox"] [role="option"]')).filter(inViewport)
    .filter(el => el.getAttribute('aria-selected') === 'true' || /已选择|selected/i.test(clean(el.getAttribute('aria-label'))));
  if (selected.length !== 1) return { count: selected.length, already: false, rect: null };
  const buttons = Array.from(selected[0].querySelectorAll('button')).filter(inViewport);
  const desiredButtons = buttons.filter(el => desiredAction.test(clean(el.getAttribute('title')) || clean(el.getAttribute('aria-label'))));
  const oppositeButtons = buttons.filter(el => oppositeAction.test(clean(el.getAttribute('title')) || clean(el.getAttribute('aria-label'))));
  if (oppositeButtons.length === 1) return { count: 1, already: true, rect: null };
  if (desiredButtons.length !== 1) return { count: desiredButtons.length, already: false, rect: null };
  const rect = desiredButtons[0].getBoundingClientRect();
  return { count: 1, already: false, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`;
    const script = `${resumeTaskSpace()}
let state = await js(${JSON.stringify(stateScript)});
if (state.count !== 1) {
  ${markedResult(`{ matchCount: 1, status: 'control_not_found', performed: false, verified: false, changed: false, state: ${JSON.stringify(desired)} }`)}
} else if (state.already) {
  ${markedResult(`{ matchCount: 1, status: 'performed', performed: true, verified: true, changed: false, state: ${JSON.stringify(desired)} }`)}
} else {
  await click(state.rect); await wait(0.6);
  state = await js(${JSON.stringify(stateScript)});
  const verified = state.count === 1 && state.already;
  ${markedResult(`{ matchCount: 1, status: verified ? 'performed' : 'control_not_found', performed: true, verified, changed: true, state: ${JSON.stringify(desired)} }`)}
}
`;
    return (await this.runner.run<MessageStateActionResult>(script, 55_000)).value;
  }

  async setReadState(locator: MessageLocator, unread: boolean): Promise<MessageStateActionResult> {
    return await this.setBinaryMessageState(locator, 'read', !unread);
  }

  async setFlagState(locator: MessageLocator, flagged: boolean): Promise<MessageStateActionResult> {
    return await this.setBinaryMessageState(locator, 'flagged', flagged);
  }

  async setCategoryState(locator: MessageLocator, category: string, applied: boolean): Promise<MessageStateActionResult> {
    const opened = await this.openAndExtractMessage(locator);
    const failure = actionFailure(opened);
    if (failure) return failure;
    const categoryScript = String.raw`(() => {
  const expected = ${JSON.stringify(category)}.normalize('NFKC').trim().toLowerCase();
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const items = Array.from(document.querySelectorAll('[role="menuitemcheckbox"]')).filter(inViewport)
    .filter(el => (el.getAttribute('aria-label') || '').normalize('NFKC').trim().toLowerCase() === expected);
  if (items.length !== 1) return { count: items.length, checked: null, rect: null };
  const rect = items[0].getBoundingClientRect();
  return { count: 1, checked: items[0].getAttribute('aria-checked') === 'true', rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`;
    const controlScript = String.raw`(() => {
  const inViewport = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0 && rect.bottom > 0 && rect.top < innerHeight; };
  const buttons = Array.from(document.querySelectorAll('button[aria-label]')).filter(inViewport)
    .filter(el => /^(分类|categorize)$/i.test((el.getAttribute('aria-label') || '').normalize('NFKC').trim()));
  if (buttons.length !== 1) return null;
  const rect = buttons[0].getBoundingClientRect(); return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
})()`;
    const script = `${resumeTaskSpace()}
const control = await js(${JSON.stringify(controlScript)});
if (!control) {
  ${markedResult(`{ matchCount: 1, status: 'control_not_found', performed: false, verified: false, changed: false, state: ${JSON.stringify(applied)}, category: ${JSON.stringify(category)} }`)}
} else {
  await click(control); await wait(0.4);
  let item = await js(${JSON.stringify(categoryScript)});
  if (item.count !== 1 || item.checked === null) {
    await pressKey('Escape');
    ${markedResult(`{ matchCount: 1, status: 'control_not_found', performed: false, verified: false, changed: false, state: ${JSON.stringify(applied)}, category: ${JSON.stringify(category)} }`)}
  } else if (item.checked === ${JSON.stringify(applied)}) {
    await pressKey('Escape');
    ${markedResult(`{ matchCount: 1, status: 'performed', performed: true, verified: true, changed: false, state: ${JSON.stringify(applied)}, category: ${JSON.stringify(category)} }`)}
  } else {
    await click(item.rect); await wait(0.3);
    const apply = await js(${JSON.stringify(String.raw`(() => { const visible=el=>{const r=el.getBoundingClientRect();return r.width>0&&r.height>0&&r.bottom>0&&r.top<innerHeight};const items=Array.from(document.querySelectorAll('[role="menuitem"]')).filter(visible).filter(el=>/^(应用|apply)$/i.test((el.getAttribute('aria-label')||el.textContent||'').normalize('NFKC').trim()));if(items.length!==1)return null;const r=items[0].getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2} })()`)});
    if (apply) { await click(apply); await wait(0.5); } else { await pressKey('Escape'); }
    await click(control); await wait(0.3);
    item = await js(${JSON.stringify(categoryScript)});
    const verified = item.count === 1 && item.checked === ${JSON.stringify(applied)};
    await pressKey('Escape');
    ${markedResult(`{ matchCount: 1, status: verified ? 'performed' : 'control_not_found', performed: true, verified, changed: true, state: ${JSON.stringify(applied)}, category: ${JSON.stringify(category)} }`)}
  }
}
`;
    return (await this.runner.run<MessageStateActionResult>(script, 65_000)).value;
  }

  async getConversation(locator: MessageLocator): Promise<ConversationOpenResult> {
    const opened = await this.openAndExtractMessage(locator);
    if (opened.matchCount !== 1 || !opened.message) return { matchCount: opened.matchCount, messages: [], complete: false };
    const script = `${resumeTaskSpace()}
const value = await js(${JSON.stringify(String.raw`(() => {
  const fallbackSubject = ${JSON.stringify(locator.subject)};
  const fallbackSenderAddress = ${JSON.stringify(locator.senderAddress)};
  const pane = document.querySelector('[role="main"][aria-label="阅读窗格"], [role="main"][aria-label="Reading pane"]');
  if (!pane) return { messages: [], complete: false };
  const visible = el => { const rect = el.getBoundingClientRect(); const style=getComputedStyle(el); return rect.width>0&&rect.height>0&&style.display!=='none'&&style.visibility!=='hidden'; };
  const clean = value => (value || '').normalize('NFKC').replace(/[\u200b\ufeff]/g, '').replace(/\s+/g, ' ').trim();
  const documents = Array.from(pane.querySelectorAll('[role="document"]')).filter(visible);
  const messages = documents.map(documentElement => {
    let container = documentElement.parentElement;
    for (let depth=0; container && container!==pane && depth<8; depth+=1, container=container.parentElement) {
      if (container.querySelector('[role="button"][aria-label^="发件人:"], [role="button"][aria-label^="From:"]')) break;
    }
    container = container && container!==pane ? container : pane;
    const headings = Array.from(container.querySelectorAll('[role="heading"]')).filter(visible);
    const fromElement = container.querySelector('[role="button"][aria-label^="发件人:"], [role="button"][aria-label^="From:"]');
    const toHeading = headings.find(el => /^(收件人|to)[:：]/i.test(clean(el.textContent)));
    const ccHeading = headings.find(el => /^(抄送|cc)[:：]/i.test(clean(el.textContent)));
    const dateHeading = headings.find(el => /(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/.test(clean(el.textContent)));
    const recipients = heading => heading ? Array.from(heading.querySelectorAll('[role="button"][aria-label]')).map(el => ({ name: clean(el.getAttribute('aria-label')) || null, address: null })) : [];
    const dateText = clean(dateHeading && dateHeading.textContent) || null;
    const dateMatch = (dateText || '').match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/);
    const receivedAt = dateMatch ? new Date(Number(dateMatch[1]),Number(dateMatch[2])-1,Number(dateMatch[3]),Number(dateMatch[4]),Number(dateMatch[5])).toISOString() : null;
    const bodyText = (documentElement.innerText || '').replace(/\r\n?/g,'\n').replace(/\u00a0/g,' ').replace(/[ \t]+\n/g,'\n').replace(/\n{3,}/g,'\n\n').trim().slice(0,100*1024);
    const attachmentRoots = Array.from(container.querySelectorAll('[role="listbox"]')).filter(el=>/文件附件|attachments?/i.test(el.getAttribute('aria-label')||''));
    const attachments = attachmentRoots.flatMap(root=>Array.from(root.querySelectorAll('[role="option"]')).map(option=>{const filenameElement=Array.from(option.querySelectorAll('[title]')).find(el=>/\.[a-z0-9]{2,8}$/i.test(clean(el.getAttribute('title'))));const filename=clean(filenameElement&&filenameElement.getAttribute('title'));const sizeMatch=clean(option.innerText).match(/\b\d+(?:\.\d+)?\s*(?:bytes?|[kmgt]b|字节)\b/i);return{filename,sizeText:sizeMatch?sizeMatch[0]:null}})).filter(item=>item.filename);
    return { subject: fallbackSubject, fromName: clean(fromElement&&fromElement.textContent)||null, fromAddress: fallbackSenderAddress, to: recipients(toHeading), cc: recipients(ccHeading), receivedAt, receivedAtText: dateText, bodyText, attachments };
  }).filter(message=>message.bodyText);
  const expandable = Array.from(pane.querySelectorAll('[aria-expanded="false"]')).filter(el=>/展开|expand|show/i.test((el.getAttribute('aria-label')||el.getAttribute('title')||'').normalize('NFKC')));
  return { messages, complete: expandable.length===0 };
})()`)});
${markedResult('{ matchCount: 1, messages: value.messages, complete: value.complete }')}
`;
    return (await this.runner.run<ConversationOpenResult>(script, 55_000)).value;
  }

  async downloadAttachment(
    locator: MessageLocator,
    attachmentIndex: number,
    outputDirectory: string,
  ): Promise<AttachmentDownloadResult> {
    const opened = await this.openAndExtractMessage(locator);
    const failure = actionFailure(opened);
    if (failure) return failure;

    const attachmentTargetScript = `(() => {
  const index = ${JSON.stringify(attachmentIndex)};
  const visible = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0; };
  const clean = value => (value || '').normalize('NFKC').replace(/\\s+/g, ' ').trim();
  const pane = document.querySelector('[role="main"][aria-label="阅读窗格"], [role="main"][aria-label="Reading pane"]');
  const lists = pane ? Array.from(pane.querySelectorAll('[role="listbox"]')).filter(el => /文件附件|attachments?/i.test(el.getAttribute('aria-label') || '')) : [];
  const options = lists.flatMap(list => Array.from(list.querySelectorAll('[role="option"]'))).filter(visible);
  const option = options[index];
  if (!option) return { count: options.length, target: null };
  const filenameElement = Array.from(option.querySelectorAll('[title]')).find(el => /\\.[a-z0-9]{2,8}$/i.test(clean(el.getAttribute('title'))));
  const button = Array.from(option.querySelectorAll('button')).find(el => /更多操作|more actions/i.test((el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')));
  if (!button) return { count: options.length, target: null };
  const rect = button.getBoundingClientRect();
  return { count: options.length, target: { filename: clean(filenameElement && filenameElement.getAttribute('title')), rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } } };
})()`;
    const downloadControlScript = String.raw`(() => {
  const visible = el => { const rect = el.getBoundingClientRect(); return rect.width > 0 && rect.height > 0; };
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const items = Array.from(document.querySelectorAll('[role="menu"] [role="menuitem"]')).filter(visible)
    .filter(el => /下载|download/i.test(clean(el.textContent)));
  if (items.length !== 1) return { count: items.length, rect: null };
  const rect = items[0].getBoundingClientRect();
  return { count: 1, rect: { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 } };
})()`;
    const script = `${resumeTaskSpace()}
const fs = await import('node:fs/promises');
const pathModule = await import('node:path');
const outputDirectory = ${JSON.stringify(outputDirectory)};
await cdp('Page.setDownloadBehavior', { behavior: 'allow', downloadPath: outputDirectory });
const before = new Map();
for (const name of await fs.readdir(outputDirectory)) {
  const info = await fs.stat(pathModule.join(outputDirectory, name));
  before.set(name, info.mtimeMs);
}
let target = null;
let attachmentCount = 0;
for (let attempt = 0; attempt < 20; attempt += 1) {
  const value = await js(${JSON.stringify(attachmentTargetScript)});
  attachmentCount = value.count;
  target = value.target;
  if (target) break;
  await wait(0.3);
}
if (!target) {
  ${markedResult("{ matchCount: 1, status: 'attachment_not_found', performed: false, verified: false, attachmentCount }")}
} else {
  await click(target.rect);
  await wait(0.4);
  const download = await js(${JSON.stringify(downloadControlScript)});
  if (!download.rect) {
    await pressKey('Escape');
    ${markedResult("{ matchCount: 1, status: 'control_not_found', performed: false, verified: false, attachmentCount }")}
  } else {
    await click(download.rect);
    let downloaded = null;
    for (let attempt = 0; attempt < 60; attempt += 1) {
      await wait(0.5);
      for (const name of await fs.readdir(outputDirectory)) {
        if (name.endsWith('.crdownload')) continue;
        const fullPath = pathModule.join(outputDirectory, name);
        const info = await fs.stat(fullPath);
        if (!before.has(name) || before.get(name) !== info.mtimeMs) {
          downloaded = { filename: name, path: fullPath, bytes: info.size };
          break;
        }
      }
      if (downloaded) break;
    }
    if (!downloaded) {
      ${markedResult("{ matchCount: 1, status: 'download_failed', performed: true, verified: false, attachmentCount }")}
    } else {
      ${markedResult(`{
        matchCount: 1,
        status: 'performed',
        performed: true,
        verified: true,
        attachmentCount,
        attachmentId: String(${JSON.stringify(attachmentIndex)} + 1),
        ...downloaded
      }`)}
    }
  }
}
`;
    return (await this.runner.run<AttachmentDownloadResult>(script, 110_000)).value;
  }
}
