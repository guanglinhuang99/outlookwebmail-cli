import { z } from 'zod';
import type { BrowserBackend } from '../browser/backend.js';
import type { MessageLocator, RawMessage } from '../types/mail.js';

const addressSchema = z.object({ name: z.string().nullable(), address: z.string().nullable() });
const rawMessageSchema = z.object({
  subject: z.string(),
  fromName: z.string().nullable(),
  fromAddress: z.string().nullable(),
  to: z.array(addressSchema),
  cc: z.array(addressSchema),
  receivedAt: z.string().nullable(),
  receivedAtText: z.string().nullable(),
  bodyText: z.string(),
  bodyTruncated: z.boolean().optional(),
  bodyBytes: z.number().int().nonnegative().optional(),
  attachments: z.array(z.object({ filename: z.string(), sizeText: z.string().nullable() })),
});

const openResultSchema = z.object({
  matchCount: z.number().int().nonnegative(),
  message: rawMessageSchema.nullable(),
});

export function buildMessageRowMatchScript(locator: MessageLocator): string {
  return String.raw`
(() => {
  const target = ${JSON.stringify(locator)};
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim().toLowerCase();
  const timeToken = value => {
    const match = clean(value).match(/(\d{1,2}:\d{2})(?!.*\d{1,2}:\d{2})/);
    return match ? match[1] : clean(value);
  };
  const minuteKey = value => {
    const match = clean(value).match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})[t\s]+(\d{1,2}):(\d{2})/i);
    if (!match) return null;
    const pad = part => String(Number(part)).padStart(2, '0');
    return match[1] + '-' + pad(match[2]) + '-' + pad(match[3]) + ' ' + pad(match[4]) + ':' + match[5];
  };
  const roots = Array.from(document.querySelectorAll('[role="listbox"]'))
    .filter(el => /邮件列表|mail list/i.test(el.getAttribute('aria-label') || ''));
  if (roots.length !== 1) return { listCount: roots.length, matches: [], firstKey: null, wheelTarget: null };
  const root = roots[0];
  const rootRect = root.getBoundingClientRect();
  const rows = Array.from(root.querySelectorAll('[role="option"]'));
  const parsed = rows.map(row => {
    const titledSpans = Array.from(row.querySelectorAll('span[title]'));
    const sender = titledSpans.find(el => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test((el.getAttribute('title') || '').trim()));
    const time = titledSpans.find(el => /(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/.test(el.getAttribute('title') || ''));
    const avatar = row.querySelector('[role="img"][aria-label]');
    const subjectBlock = time ? time.parentElement : null;
    const subject = subjectBlock ? Array.from(subjectBlock.children).find(el => el !== time && clean(el.textContent)) : null;
    const preview = subjectBlock ? subjectBlock.nextElementSibling : null;
    const rect = row.getBoundingClientRect();
    const parsedRow = {
      key: row.id || row.getAttribute('aria-label') || [sender && sender.getAttribute('title'), subject && subject.textContent, time && time.textContent].join('|'),
      stableHint: [row.getAttribute('data-item-id'), row.getAttribute('data-message-id'), row.getAttribute('data-convid'), row.getAttribute('data-conversation-id')]
        .map(clean).find(Boolean) || null,
      senderName: clean(sender ? sender.textContent : avatar && avatar.getAttribute('aria-label')),
      senderAddress: clean(sender && sender.getAttribute('title')),
      subject: clean(subject && subject.textContent),
      receivedAt: clean(time && time.getAttribute('title')),
      receivedAtText: clean(time && time.textContent),
      preview: clean(preview && preview.textContent),
      rect: { x: rect.x, y: rect.y, width: rect.width, height: rect.height },
      inViewport: rect.bottom > 0 && rect.top < window.innerHeight
    };
    const identityKey = [
      parsedRow.senderAddress || parsedRow.senderName,
      parsedRow.subject,
      minuteKey(parsedRow.receivedAt) || timeToken(parsedRow.receivedAtText)
    ].join('|');
    return {
      ...parsedRow,
      identityKey,
      logicalKey: [identityKey, parsedRow.preview].join('|')
    };
  });
  const matches = parsed.filter(row =>
    row.inViewport &&
    (!target.stableHint || row.stableHint === clean(target.stableHint)) &&
    row.subject === clean(target.subject) &&
    (!target.senderAddress && !target.senderName
      ? true
      : target.senderAddress
        ? row.senderAddress === clean(target.senderAddress)
        : row.senderName === clean(target.senderName)) &&
    (target.receivedAt
      ? minuteKey(row.receivedAt) === minuteKey(target.receivedAt)
      : !target.receivedAtText || timeToken(row.receivedAtText) === timeToken(target.receivedAtText))
  ).map(row => ({
    ...row,
    previewMatches: !target.preview || row.preview === clean(target.preview)
  }));
  return {
    listCount: roots.length,
    matches,
    firstKey: parsed[0] ? parsed[0].key : null,
    wheelTarget: { x: rootRect.x + rootRect.width / 2, y: rootRect.y + rootRect.height / 2 }
  };
})()
`;
}

export function buildMessageExtractScript(locator: MessageLocator): string {
  return String.raw`
(() => {
  const fallbackSenderAddress = ${JSON.stringify(locator.senderAddress)};
  const pane = document.querySelector('[role="main"][aria-label="阅读窗格"], [role="main"][aria-label="Reading pane"]');
  if (!pane) return null;
  const visible = el => {
    const rect = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return rect.width > 0 && rect.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
  };
  const clean = value => (value || '').normalize('NFKC').replace(/[\u200b\ufeff]/g, '').replace(/\s+/g, ' ').trim();
  const headings = Array.from(pane.querySelectorAll('[role="heading"]')).filter(visible);
  const subjectElement = headings.find(el => el.tagName === 'SPAN' && clean(el.getAttribute('title')) === clean(el.textContent));
  const fromElement = pane.querySelector('[role="button"][aria-label^="发件人:"], [role="button"][aria-label^="From:"]');
  const toHeading = headings.find(el => /^(收件人|to)[:：]/i.test(clean(el.textContent)));
  const ccHeading = headings.find(el => /^(抄送|cc)[:：]/i.test(clean(el.textContent)));
  const dateHeading = headings.find(el => /(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/.test(clean(el.textContent)));
  const documentElement = Array.from(pane.querySelectorAll('[role="document"]')).find(visible);

  const recipients = heading => heading
    ? Array.from(heading.querySelectorAll('[role="button"][aria-label]')).map(el => ({
        name: clean(el.getAttribute('aria-label')) || null,
        address: null
      }))
    : [];

  function toIso(value) {
    const match = value.match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/);
    if (!match) return null;
    const [, year, month, day, hour, minute] = match;
    const local = new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute));
    const offsetMinutes = -local.getTimezoneOffset();
    const sign = offsetMinutes >= 0 ? '+' : '-';
    const offset = Math.abs(offsetMinutes);
    const pad = part => String(part).padStart(2, '0');
    return year + '-' + pad(month) + '-' + pad(day) + 'T' + pad(hour) + ':' + minute + ':00' + sign
      + pad(Math.floor(offset / 60)) + ':' + pad(offset % 60);
  }

  function truncateUtf8(value, maximumBytes) {
    const encoder = new TextEncoder();
    const encoded = encoder.encode(value);
    if (encoded.length <= maximumBytes) return { text: value, truncated: false, bytes: encoded.length };
    let output = '';
    let bytes = 0;
    for (const character of value) {
      const size = encoder.encode(character).length;
      if (bytes + size > maximumBytes) break;
      output += character;
      bytes += size;
    }
    return { text: output, truncated: true, bytes: encoded.length };
  }

  const body = documentElement
    ? (documentElement.innerText || '').replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim()
    : '';
  const bodyResult = truncateUtf8(body, 100 * 1024);
  const attachmentRoots = Array.from(pane.querySelectorAll('[role="listbox"]'))
    .filter(el => /文件附件|attachments?/i.test(el.getAttribute('aria-label') || ''));
  const attachments = attachmentRoots.flatMap(root => Array.from(root.querySelectorAll('[role="option"]')).map(option => {
    const titleCandidates = Array.from(option.querySelectorAll('[title]')).map(el => clean(el.getAttribute('title')))
      .filter(title => title && !/^(更多操作|more actions|下载|download|预览|preview|打开|open)$/i.test(title)
        && !/^\d+(?:\.\d+)?\s*(?:bytes?|[kmgt]b|字节)$/i.test(title));
    const textCandidates = (option.innerText || '').split(/\r?\n/).map(clean)
      .filter(text => text && !/^(更多操作|more actions|下载|download|预览|preview|打开|open)$/i.test(text)
        && !/^\d+(?:\.\d+)?\s*(?:bytes?|[kmgt]b|字节)$/i.test(text));
    const filename = titleCandidates[0] || textCandidates[0] || '';
    const sizeMatch = clean(option.innerText).match(/\b\d+(?:\.\d+)?\s*(?:bytes?|[kmgt]b|字节)\b/i);
    return { filename, sizeText: sizeMatch ? sizeMatch[0] : null };
  })).filter(item => item.filename);

  const receivedAtText = clean(dateHeading && dateHeading.textContent) || null;
  return {
    subject: clean(subjectElement && subjectElement.textContent),
    fromName: clean(fromElement && fromElement.textContent) || null,
    fromAddress: fallbackSenderAddress,
    to: recipients(toHeading),
    cc: recipients(ccHeading),
    receivedAt: toIso(receivedAtText || ''),
    receivedAtText,
    bodyText: bodyResult.text,
    bodyTruncated: bodyResult.truncated,
    bodyBytes: bodyResult.bytes,
    attachments
  };
})()
`;
}

export interface MessageParser {
  openAndExtract(locator: MessageLocator): Promise<{ matchCount: number; message: RawMessage | null }>;
}

export class EgoMessageParser implements MessageParser {
  constructor(private readonly backend: BrowserBackend) {}

  async openAndExtract(locator: MessageLocator): Promise<{ matchCount: number; message: RawMessage | null }> {
    return openResultSchema.parse(await this.backend.openAndExtractMessage(locator));
  }
}
