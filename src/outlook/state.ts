export type OutlookState =
  | 'AUTH_REQUIRED'
  | 'INBOX'
  | 'SEARCH_RESULTS'
  | 'MESSAGE_OPEN'
  | 'COMPOSE_OPEN'
  | 'UNKNOWN';

const OUTLOOK_HOST = 'partner.outlook.cn';

export function isAllowedOutlookUrl(value: string | null | undefined): boolean {
  if (!value) return false;

  try {
    const url = new URL(value);
    return url.protocol === 'https:' && url.hostname === OUTLOOK_HOST && url.pathname.startsWith('/mail');
  } catch {
    return false;
  }
}

export function detectOutlookState(url: string | null, snapshot: string): OutlookState {
  const normalized = snapshot.normalize('NFKC');

  if (!isAllowedOutlookUrl(url)) return 'AUTH_REQUIRED';
  if (/密码|登录|登入|sign in|password/i.test(normalized)) return 'AUTH_REQUIRED';
  if (/新邮件|新建邮件|发送|discard|放弃草稿/i.test(normalized) && /message body|邮件正文|正文/i.test(normalized)) {
    return 'COMPOSE_OPEN';
  }
  if (/搜索结果|search results/i.test(normalized)) return 'SEARCH_RESULTS';
  if (/回复|转发|reply|forward/i.test(normalized) && /收件人|发件人|to:|from:/i.test(normalized)) {
    return 'MESSAGE_OPEN';
  }
  if (/收件箱|inbox/i.test(normalized)) return 'INBOX';
  return 'UNKNOWN';
}
