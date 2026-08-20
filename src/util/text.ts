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
