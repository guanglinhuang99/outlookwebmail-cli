import type { BrowserBackend } from './backend.js';
import { EgoLiteBackend } from './ego-lite.js';
import { PlaywrightBackend } from './playwright.js';
import { loadPlaywrightConfig, type ConfigOptions } from './playwright-config.js';

function canFallback(error: unknown): boolean {
  if (!error || typeof error !== 'object' || !('code' in error)) return false;
  return ['BROWSER_NOT_FOUND', 'PROFILE_LOCKED', 'PLAYWRIGHT_ERROR', 'PLAYWRIGHT_TIMEOUT'].includes(String(error.code));
}

export function createFallbackBackend(primary: BrowserBackend, fallback: BrowserBackend): BrowserBackend {
  let selected: BrowserBackend | null = null;
  return new Proxy(primary, {
    get(_target, property) {
      if (property === 'name') return selected?.name ?? primary.name;
      if (property === 'close') {
        return async () => {
          await primary.close?.();
          if (fallback !== primary) await fallback.close?.();
        };
      }
      const primaryValue = Reflect.get(primary, property);
      if (typeof primaryValue !== 'function') return primaryValue;
      return async (...args: unknown[]) => {
        if (selected) {
          const selectedValue = Reflect.get(selected, property);
          return await selectedValue.apply(selected, args);
        }
        try {
          const result = await primaryValue.apply(primary, args);
          selected = primary;
          return result;
        } catch (primaryError) {
          if (!canFallback(primaryError)) throw primaryError;
          try {
            const fallbackValue = Reflect.get(fallback, property);
            const result = await fallbackValue.apply(fallback, args);
            selected = fallback;
            return result;
          } catch {
            throw primaryError;
          }
        }
      };
    },
  });
}

export function createBrowserBackend(options: ConfigOptions = {}): BrowserBackend {
  const config = loadPlaywrightConfig(options);
  if (config.backend === 'ego-lite') return new EgoLiteBackend();
  const playwright = new PlaywrightBackend(config);
  if (config.backend === 'playwright') return playwright;
  return createFallbackBackend(playwright, new EgoLiteBackend());
}
