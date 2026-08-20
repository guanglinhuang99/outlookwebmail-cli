import { createHash } from 'node:crypto';

export function normalizeText(value: string | null | undefined): string {
  return (value ?? '').normalize('NFKC').replace(/\s+/g, ' ').trim().toLowerCase();
}

function normalizeStableHint(value: string | null | undefined): string {
  return (value ?? '').normalize('NFKC').trim();
}

export function messageFingerprint(parts: {
  stableHint?: string | null;
  senderName?: string | null;
  senderAddress?: string | null;
  subject: string;
  receivedAtText?: string | null;
  preview?: string | null;
}): string {
  const stableHint = normalizeStableHint(parts.stableHint);
  if (stableHint) return ['stable-hint', stableHint].join('\u001f');
  return [
    normalizeText(parts.senderAddress || parts.senderName),
    normalizeText(parts.subject),
    normalizeText(parts.receivedAtText),
    normalizeText(parts.preview).slice(0, 80),
  ].join('\u001f');
}

export function stableMessageId(parts: {
  stableHint?: string | null;
  senderName?: string | null;
  senderAddress?: string | null;
  subject: string;
  receivedAt?: string | null;
  receivedAtText?: string | null;
  preview?: string | null;
}): string {
  const stableHint = normalizeStableHint(parts.stableHint);
  const timestamp = normalizeText(parts.receivedAt || parts.receivedAtText);
  const identity = stableHint
    ? ['stable-hint', stableHint].join('\u001f')
    : [
        normalizeText(parts.senderAddress || parts.senderName),
        normalizeText(parts.subject),
        timestamp,
        normalizeText(parts.preview).slice(0, 120),
      ].join('\u001f');
  return `m_${createHash('sha256').update(identity).digest('base64url').slice(0, 20)}`;
}
