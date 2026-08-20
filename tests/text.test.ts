import { describe, expect, it } from 'vitest';
import { messageFingerprint, normalizeText, stableMessageId } from '../src/util/text.js';

describe('mail text helpers', () => {
  it('normalizes whitespace, width and case', () => {
    expect(normalizeText('  ＲＩＳＫ\n Report  ')).toBe('risk report');
  });

  it('creates the same fingerprint from equivalent text', () => {
    const first = messageFingerprint({ senderName: '张三', subject: 'Risk  Report', receivedAtText: '09:00', preview: 'Hello' });
    const second = messageFingerprint({ senderName: ' 张三 ', subject: 'risk report', receivedAtText: '09:00', preview: ' hello ' });
    expect(first).toBe(second);
  });

  it('creates deterministic opaque stable message IDs', () => {
    const first = stableMessageId({ senderAddress: 'USER@example.com', subject: ' Risk Report ', receivedAt: '2026-08-20T09:00:00+08:00' });
    const second = stableMessageId({ senderAddress: 'user@example.com', subject: 'risk report', receivedAt: '2026-08-20T09:00:00+08:00' });
    expect(first).toBe(second);
    expect(first).toMatch(/^m_[A-Za-z0-9_-]{20}$/);
    expect(first).not.toContain('user');
  });
});
