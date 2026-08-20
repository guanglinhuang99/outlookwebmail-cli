import { createHash } from 'node:crypto';

export function normalizeText(value: string | null | undefined): string {
  return (value ?? '').normalize('NFKC').replace(/\s+/g, ' ').trim().toLowerCase();
}

export function messageFingerprint(parts: {
  senderName?: string | null;
  senderAddress?: string | null;
  subject: string;
  receivedAtText?: string | null;
  preview?: string | null;
}): string {
  return [
    normalizeText(parts.senderAddress || parts.senderName),
    normalizeText(parts.subject),
    normalizeText(parts.receivedAtText),
    normalizeText(parts.preview).slice(0, 80),
  ].join('\u001f');
}

export function stableMessageId(parts: {
  senderName?: string | null;
  senderAddress?: string | null;
  subject: string;
  receivedAt?: string | null;
  receivedAtText?: string | null;
  preview?: string | null;
}): string {
  const timestamp = normalizeText(parts.receivedAt || parts.receivedAtText);
  const fallback = timestamp ? '' : normalizeText(parts.preview).slice(0, 120);
  const identity = [
    normalizeText(parts.senderAddress || parts.senderName),
    normalizeText(parts.subject),
    timestamp,
    fallback,
  ].join('\u001f');
  return `m_${createHash('sha256').update(identity).digest('base64url').slice(0, 20)}`;
}
