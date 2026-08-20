import { describe, expect, it } from 'vitest';
import { detectOutlookState, isAllowedOutlookUrl } from '../src/outlook/state.js';

describe('isAllowedOutlookUrl', () => {
  it('accepts only the exact HTTPS Outlook host and mail path', () => {
    expect(isAllowedOutlookUrl('https://partner.outlook.cn/mail/inbox')).toBe(true);
    expect(isAllowedOutlookUrl('http://partner.outlook.cn/mail/')).toBe(false);
    expect(isAllowedOutlookUrl('https://partner.outlook.cn.evil.example/mail/')).toBe(false);
    expect(isAllowedOutlookUrl('https://partner.outlook.cn/calendar/')).toBe(false);
  });
});

describe('detectOutlookState', () => {
  const url = 'https://partner.outlook.cn/mail/';

  it('detects login, inbox and search states', () => {
    expect(detectOutlookState('https://login.example/', 'Sign in')).toBe('AUTH_REQUIRED');
    expect(detectOutlookState(url, '导航 收件箱 草稿')).toBe('INBOX');
    expect(detectOutlookState(url, '搜索结果 收件箱')).toBe('SEARCH_RESULTS');
  });

  it('returns UNKNOWN when no calibrated signal exists', () => {
    expect(detectOutlookState(url, 'Outlook')).toBe('UNKNOWN');
  });
});
