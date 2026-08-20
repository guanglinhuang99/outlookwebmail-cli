import type { BrowserBackend } from './backend.js';
import { EgoRunner } from './ego-runner.js';
import type { BrowserStatus, InspectResult, MessageInspectResult, PageInfo } from '../types/inspect.js';
import type {
  AttachmentDownloadResult,
  FolderSelectionResult,
  InboxFolderListResult,
  MessageActionResult,
  MessageLocator,
  MessageOpenResult,
  ReplyActionResult,
} from '../types/mail.js';
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
  constructor(private readonly runner = new EgoRunner()) {}

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
