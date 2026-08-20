import { describe, expect, it, vi } from 'vitest';
import type { BrowserBackend } from '../src/browser/backend.js';
import { createBrowserBackend, createFallbackBackend } from '../src/browser/browser-factory.js';
import { AppError } from '../src/util/errors.js';

function backend(name: 'playwright' | 'ego-lite', status: () => Promise<unknown>): BrowserBackend {
  return {
    name,
    status,
  } as unknown as BrowserBackend;
}

describe('browser backend factory', () => {
  it('uses Playwright by default and preserves explicit Ego Lite selection', () => {
    const common = { cwd: '/work/repo', home: '/users/test' };
    expect(createBrowserBackend({ ...common, env: {} }).name).toBe('playwright');
    expect(createBrowserBackend({ ...common, env: { WEBMAIL_BACKEND: 'playwright' } }).name).toBe('playwright');
    expect(createBrowserBackend({ ...common, env: { WEBMAIL_BACKEND: 'ego-lite' } }).name).toBe('ego-lite');
  });

  it('falls back once when Playwright cannot start', async () => {
    const primaryStatus = vi.fn(async () => { throw new AppError('BROWSER_NOT_FOUND', 'missing'); });
    const fallbackStatus = vi.fn(async () => ({ connected: true }));
    const selected = createFallbackBackend(
      backend('playwright', primaryStatus),
      backend('ego-lite', fallbackStatus),
    );

    await expect(selected.status()).resolves.toEqual({ connected: true });
    await expect(selected.status()).resolves.toEqual({ connected: true });
    expect(selected.name).toBe('ego-lite');
    expect(primaryStatus).toHaveBeenCalledOnce();
    expect(fallbackStatus).toHaveBeenCalledTimes(2);
  });

  it('does not hide non-startup Playwright errors', async () => {
    const fallbackStatus = vi.fn(async () => ({ connected: true }));
    const selected = createFallbackBackend(
      backend('playwright', async () => { throw new AppError('AUTH_REQUIRED', 'login'); }),
      backend('ego-lite', fallbackStatus),
    );

    await expect(selected.status()).rejects.toMatchObject({ code: 'AUTH_REQUIRED' });
    expect(fallbackStatus).not.toHaveBeenCalled();
  });
});
