import { z } from 'zod';
import type { BrowserBackend } from '../browser/backend.js';
import type { RawMessageRow } from '../types/mail.js';
import { AppError } from '../util/errors.js';

const rawMessageRowSchema = z.object({
  stableHint: z.string().nullable(),
  senderName: z.string().nullable(),
  senderAddress: z.string().nullable(),
  subject: z.string(),
  receivedAt: z.string().nullable(),
  receivedAtText: z.string().nullable(),
  preview: z.string().nullable(),
  unread: z.boolean().nullable(),
  hasAttachments: z.boolean().nullable(),
});

const extractionSchema = z.object({
  listCount: z.number().int().nonnegative(),
  rowCount: z.number().int().nonnegative(),
  rows: z.array(rawMessageRowSchema),
});

const MAIL_LIST_SELECTOR = '[role="listbox"][aria-label^="邮件列表"], [role="listbox"][aria-label^="Mail list"]';

export const INBOX_EXTRACT_SCRIPT = String.raw`
(() => {
  const clean = value => (value || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  const listRoots = Array.from(document.querySelectorAll('[role="listbox"]'))
    .filter(el => /邮件列表|mail list/i.test(el.getAttribute('aria-label') || ''));

  if (listRoots.length !== 1) {
    return { listCount: listRoots.length, rowCount: 0, rows: [] };
  }

  const root = listRoots[0];
  const rowElements = Array.from(root.querySelectorAll('[role="option"]'));

  function toIso(title) {
    const match = title.match(/(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/);
    if (!match) return null;
    const [, year, month, day, hour, minute] = match;
    const local = new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute));
    const offsetMinutes = -local.getTimezoneOffset();
    const sign = offsetMinutes >= 0 ? '+' : '-';
    const offset = Math.abs(offsetMinutes);
    const pad = value => String(value).padStart(2, '0');
    return year + '-' + pad(month) + '-' + pad(day) + 'T' + pad(hour) + ':' + minute + ':00' + sign
      + pad(Math.floor(offset / 60)) + ':' + pad(offset % 60);
  }

  const rows = rowElements.map(row => {
    const ariaLabel = clean(row.getAttribute('aria-label'));
    const titledSpans = Array.from(row.querySelectorAll('span[title]'));
    const senderElement = titledSpans.find(el => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean(el.getAttribute('title'))));
    const timeElement = titledSpans.find(el => /(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})\s+(\d{1,2}):(\d{2})/.test(el.getAttribute('title') || ''));
    const avatarElement = row.querySelector('[role="img"][aria-label]');
    const subjectBlock = timeElement ? timeElement.parentElement : null;
    const subjectElement = subjectBlock
      ? Array.from(subjectBlock.children).find(el => el !== timeElement && clean(el.textContent).length > 0)
      : null;
    const previewBlock = subjectBlock ? subjectBlock.nextElementSibling : null;
    const markerTitles = Array.from(row.querySelectorAll('button[title]')).map(el => clean(el.getAttribute('title')));
    const stableHint = [
      row.getAttribute('data-item-id'),
      row.getAttribute('data-message-id'),
      row.getAttribute('data-convid'),
      row.getAttribute('data-conversation-id')
    ].map(clean).find(Boolean) || null;

    let unread = null;
    if (markerTitles.some(title => /标记为已读|mark as read/i.test(title))) unread = true;
    else if (markerTitles.some(title => /标记为未读|mark as unread/i.test(title))) unread = false;
    else if (/^未读\b|^unread\b/i.test(ariaLabel)) unread = true;

    const timeTitle = clean(timeElement && timeElement.getAttribute('title'));
    return {
      stableHint,
      senderName: clean(senderElement ? senderElement.textContent : avatarElement && avatarElement.getAttribute('aria-label')) || null,
      senderAddress: clean(senderElement && senderElement.getAttribute('title')) || null,
      subject: clean(subjectElement && subjectElement.textContent),
      receivedAt: toIso(timeTitle),
      receivedAtText: clean(timeElement && timeElement.textContent) || null,
      preview: clean(previewBlock && previewBlock.textContent) || null,
      unread,
      hasAttachments: /带附件|has attachments?/i.test(ariaLabel)
    };
  }).filter(row => row.subject.length > 0);

  return { listCount: listRoots.length, rowCount: rowElements.length, rows };
})()
`;

export interface InboxParser {
  extract(): Promise<RawMessageRow[]>;
  resetAndExtract(): Promise<RawMessageRow[]>;
  scrollAndExtract(): Promise<RawMessageRow[]>;
}

export class EgoInboxParser implements InboxParser {
  constructor(private readonly backend: BrowserBackend) {}

  async extract(): Promise<RawMessageRow[]> {
    return this.parseExtraction(await this.backend.eval<unknown>(INBOX_EXTRACT_SCRIPT));
  }

  private parseExtraction(value: unknown): RawMessageRow[] {
    const result = extractionSchema.parse(value);
    if (result.listCount !== 1) {
      throw new AppError('UI_CHANGED', `预期找到 1 个邮件列表，实际找到 ${result.listCount} 个。`);
    }
    if (result.rowCount > 0 && result.rows.length === 0) {
      throw new AppError('UI_CHANGED', `邮件列表包含 ${result.rowCount} 行，但没有任何一行能按已确认结构解析。`);
    }
    return result.rows;
  }

  async resetAndExtract(): Promise<RawMessageRow[]> {
    let previousFirst: string | null = null;
    let rows: RawMessageRow[] = [];
    for (let attempt = 0; attempt < 6; attempt += 1) {
      const value = await this.backend.wheelAndEval<unknown>(MAIL_LIST_SELECTOR, -900, 1, 400, INBOX_EXTRACT_SCRIPT);
      rows = this.parseExtraction(value);
      const first = rows[0];
      const firstFingerprint = first
        ? [first.senderAddress, first.subject, first.receivedAtText, first.preview].join('\u001f')
        : '';
      if (firstFingerprint === previousFirst) break;
      previousFirst = firstFingerprint;
    }
    return rows;
  }

  async scrollAndExtract(): Promise<RawMessageRow[]> {
    const value = await this.backend.wheelAndEval<unknown>(MAIL_LIST_SELECTOR, 900, 1, 500, INBOX_EXTRACT_SCRIPT);
    return this.parseExtraction(value);
  }
}
