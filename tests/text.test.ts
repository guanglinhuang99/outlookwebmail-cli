import { describe, expect, it } from 'vitest';
import { messageFingerprint, normalizeText } from '../src/util/text.js';

describe('mail text helpers', () => {
  it('normalizes whitespace, width and case', () => {
    expect(normalizeText('  ＲＩＳＫ\n Report  ')).toBe('risk report');
  });

  it('creates the same fingerprint from equivalent text', () => {
    const first = messageFingerprint({ senderName: '张三', subject: 'Risk  Report', receivedAtText: '09:00', preview: 'Hello' });
    const second = messageFingerprint({ senderName: ' 张三 ', subject: 'risk report', receivedAtText: '09:00', preview: ' hello ' });
    expect(first).toBe(second);
  });
});
