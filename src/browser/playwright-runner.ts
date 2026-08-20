import type { Browser, BrowserContext, CDPSession, Page } from 'playwright-core';
import { chromium } from 'playwright-core';
import { AppError } from '../util/errors.js';
import { parseMarkedResult, type BrowserScriptRunner, type EgoRunResult } from './ego-runner.js';
import {
  loadPlaywrightConfig,
  prepareProfileDirectory,
  redactCdpEndpoint,
  resolveBrowserExecutable,
  type BrowserExecutable,
  type PlaywrightConfig,
} from './playwright-config.js';

type ScriptFunction = (...args: unknown[]) => Promise<void>;
type AsyncFunctionConstructor = new (...args: string[]) => ScriptFunction;
type ClickTarget = string | { x: number; y: number };

const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor as AsyncFunctionConstructor;

export class PlaywrightRunner implements BrowserScriptRunner {
  private browser: Browser | null = null;
  private context: BrowserContext | null = null;
  private page: Page | null = null;
  private cdpSession: CDPSession | null = null;
  private launchPromise: Promise<void> | null = null;
  private lastDialog: { type: string; message: string } | null = null;
  private browserExecutable: BrowserExecutable | null = null;
  private readonly observedPages = new WeakSet<Page>();

  constructor(readonly config: PlaywrightConfig = loadPlaywrightConfig()) {}

  get executable(): BrowserExecutable | null {
    return this.browserExecutable;
  }

  private mapError(error: unknown): AppError {
    if (error instanceof AppError) return error;
    const message = error instanceof Error ? error.message : String(error);
    if (/executable .* (doesn.t exist|not found)|enoent/i.test(message)) {
      return new AppError('BROWSER_NOT_FOUND', '未找到可用的 Edge、Chrome 或 Chromium。', { cause: error as Error });
    }
    if (/user data directory is already in use|profile.*(locked|in use)|processsingleton/i.test(message)) {
      return new AppError('PROFILE_LOCKED', 'Playwright 专用 Profile 正被其他进程使用；请关闭其他 webmail-cli 实例后重试。', { cause: error as Error });
    }
    if (/timeout/i.test(message)) {
      return new AppError('PLAYWRIGHT_TIMEOUT', `Playwright 操作超时：${message.slice(0, 500)}`, { cause: error as Error });
    }
    if (/target page, context or browser has been closed|targetclosed/i.test(message)) {
      return new AppError('PLAYWRIGHT_ERROR', 'Playwright 浏览器页面或上下文已关闭。', { cause: error as Error });
    }
    return new AppError('PLAYWRIGHT_ERROR', `Playwright 操作失败：${message.slice(0, 1000)}`, { cause: error as Error });
  }

  private async launch(): Promise<void> {
    try {
      if (this.config.cdpEndpoint) {
        this.browser = await chromium.connectOverCDP(this.config.cdpEndpoint, { timeout: this.config.timeoutMs });
        this.context = this.browser.contexts()[0] ?? null;
        if (!this.context) throw new AppError('PLAYWRIGHT_ERROR', `CDP 未返回可用浏览器上下文：${redactCdpEndpoint(this.config.cdpEndpoint)}`);
      } else {
        this.browserExecutable = await resolveBrowserExecutable(this.config);
        await prepareProfileDirectory(this.config);
        this.context = await chromium.launchPersistentContext(this.config.profileDir, {
          executablePath: this.browserExecutable.path,
          headless: this.config.headless,
          acceptDownloads: true,
          timeout: this.config.timeoutMs,
        });
      }
      this.context.setDefaultTimeout(this.config.timeoutMs);
      this.context.setDefaultNavigationTimeout(this.config.timeoutMs);
      this.page = this.context.pages().find(page => this.isOutlookUrl(page.url()))
        ?? this.context.pages()[0]
        ?? await this.context.newPage();
      this.observePage(this.page);
      this.context.on('page', page => this.observePage(page));
    } catch (error) {
      await this.close().catch(() => undefined);
      throw this.mapError(error);
    }
  }

  private async ensureLaunched(): Promise<void> {
    if (this.context && this.page && !this.page.isClosed()) return;
    this.launchPromise ??= this.launch().finally(() => {
      this.launchPromise = null;
    });
    await this.launchPromise;
  }

  private observePage(page: Page): void {
    if (this.observedPages.has(page)) return;
    this.observedPages.add(page);
    page.on('dialog', async dialog => {
      this.lastDialog = { type: dialog.type(), message: dialog.message().slice(0, 300) };
      await dialog.dismiss().catch(() => undefined);
    });
  }

  private isOutlookUrl(rawUrl: string): boolean {
    try {
      const url = new URL(rawUrl);
      return url.hostname.toLowerCase() === new URL(this.config.outlookUrl).hostname.toLowerCase()
        && url.pathname.startsWith('/mail');
    } catch {
      return false;
    }
  }

  private currentPage(): Page {
    if (!this.page || this.page.isClosed()) throw new AppError('PLAYWRIGHT_ERROR', 'Playwright 页面不可用。');
    return this.page;
  }

  private async openOrReuseTab(rawUrl: string): Promise<Page> {
    const url = new URL(rawUrl);
    if (url.protocol !== 'https:' || url.hostname.toLowerCase() !== new URL(this.config.outlookUrl).hostname.toLowerCase()) {
      throw new AppError('INVALID_ARGUMENT', '浏览器后端拒绝打开非 Outlook 目标地址。');
    }
    const context = this.context;
    if (!context) throw new AppError('PLAYWRIGHT_ERROR', 'Playwright 浏览器上下文不可用。');
    const existing = context.pages().find(page => this.isOutlookUrl(page.url()) && !page.isClosed());
    const nextPage = existing ?? (this.page && !this.page.isClosed() ? this.page : await context.newPage());
    if (this.page !== nextPage) this.cdpSession = null;
    this.page = nextPage;
    this.observePage(this.page);
    if (!this.isOutlookUrl(this.page.url())) {
      await this.page.goto(url.href, { waitUntil: 'domcontentloaded', timeout: this.config.timeoutMs });
    }
    await this.page.bringToFront();
    return this.page;
  }

  private async evaluate<T>(expression: string): Promise<T> {
    return await this.currentPage().evaluate(expression) as T;
  }

  private async click(target: ClickTarget): Promise<void> {
    const page = this.currentPage();
    if (typeof target !== 'string') {
      await page.mouse.click(target.x, target.y);
      return;
    }
    const locator = page.locator(target);
    const count = await locator.count();
    if (count !== 1) throw new AppError('UI_CHANGED', `点击目标应唯一匹配，实际匹配 ${count} 个：${target}`);
    await locator.click();
  }

  private async fillInput(selector: string, value: string): Promise<void> {
    const locator = this.currentPage().locator(selector);
    const count = await locator.count();
    if (count !== 1) throw new AppError('UI_CHANGED', `输入目标应唯一匹配，实际匹配 ${count} 个：${selector}`);
    await locator.fill(value);
  }

  private async cdp(method: string, params: Record<string, unknown> = {}): Promise<unknown> {
    const page = this.currentPage();
    if (method === 'Input.insertText') {
      await page.keyboard.insertText(String(params.text ?? ''));
      return {};
    }
    if (method === 'Page.reload') {
      await page.reload({ waitUntil: 'domcontentloaded', timeout: this.config.timeoutMs });
      return {};
    }
    if (!this.cdpSession) {
      const context = this.context;
      if (!context) throw new AppError('PLAYWRIGHT_ERROR', 'Playwright 浏览器上下文不可用。');
      this.cdpSession = await context.newCDPSession(page);
    }
    return await this.cdpSession.send(method as never, params as never);
  }

  private helpers(stdout: string[]): Record<string, unknown> {
    return {
      useOrCreateTaskSpace: async (_name: string) => ({ id: 'playwright' }),
      openOrReuseTab: async (url: string) => await this.openOrReuseTab(url),
      ensureRealTab: async () => this.page && !this.page.isClosed() ? this.page : null,
      pageInfo: async () => {
        const page = this.currentPage();
        const metrics = await page.evaluate(() => ({
          w: innerWidth,
          h: innerHeight,
          sx: scrollX,
          sy: scrollY,
          pw: document.documentElement.scrollWidth,
          ph: document.documentElement.scrollHeight,
        }));
        const dialog = this.lastDialog;
        this.lastDialog = null;
        return { url: page.url(), title: await page.title(), ...metrics, dialog };
      },
      snapshotText: async () => await this.currentPage().locator('body').innerText(),
      js: async (expression: string) => await this.evaluate(expression),
      click: async (target: ClickTarget) => await this.click(target),
      wait: async (seconds: number) => await this.currentPage().waitForTimeout(Math.max(0, seconds * 1_000)),
      fillInput: async (selector: string, value: string) => await this.fillInput(selector, value),
      pressKey: async (key: string) => await this.currentPage().keyboard.press(key),
      scrollBy: async (y: number) => await this.currentPage().mouse.wheel(0, y),
      hover: async (target: ClickTarget) => {
        if (typeof target === 'string') await this.currentPage().locator(target).hover();
        else await this.currentPage().mouse.move(target.x, target.y);
      },
      scroll: async ({ dx = 0, dy = 0 }: { dx?: number; dy?: number }) => await this.currentPage().mouse.wheel(dx, dy),
      cdp: async (method: string, params?: Record<string, unknown>) => await this.cdp(method, params),
      handOffTaskSpace: async (_id: string | number) => ({ done: true }),
      uploadFile: async (selector: string, path: string) => await this.currentPage().locator(selector).setInputFiles(path),
      cliLog: (value: unknown) => stdout.push(String(value)),
    };
  }

  async run<T>(body: string, timeoutMs = this.config.timeoutMs): Promise<EgoRunResult<T>> {
    await this.ensureLaunched();
    const stdout: string[] = [];
    const helpers = this.helpers(stdout);
    const names = Object.keys(helpers);
    let timer: NodeJS.Timeout | undefined;
    try {
      const execute = new AsyncFunction(...names, body);
      const timeout = new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => {
          void this.close();
          reject(new AppError('PLAYWRIGHT_TIMEOUT', `Playwright 操作超过 ${timeoutMs}ms。`));
        }, timeoutMs);
      });
      await Promise.race([execute(...Object.values(helpers)), timeout]);
      const output = stdout.join('\n');
      return { stdout: output, stderr: '', value: parseMarkedResult<T>(output) };
    } catch (error) {
      throw this.mapError(error);
    } finally {
      clearTimeout(timer);
    }
  }

  async close(): Promise<void> {
    this.cdpSession = null;
    const context = this.context;
    const browser = this.browser;
    this.page = null;
    this.context = null;
    this.browser = null;
    if (context && !this.config.cdpEndpoint) await context.close().catch(() => undefined);
    if (browser) await browser.close().catch(() => undefined);
  }
}
