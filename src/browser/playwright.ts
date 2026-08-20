import type { BrowserBackendName } from './backend.js';
import { EgoLiteBackend } from './ego-lite.js';
import { PlaywrightRunner } from './playwright-runner.js';
import type { PlaywrightConfig } from './playwright-config.js';

export class PlaywrightBackend extends EgoLiteBackend {
  override readonly name: BrowserBackendName = 'playwright';
  private readonly playwrightRunner: PlaywrightRunner;

  constructor(config?: PlaywrightConfig) {
    const runner = new PlaywrightRunner(config);
    super(runner);
    this.playwrightRunner = runner;
  }

  override async status() {
    const status = await super.status();
    return {
      ...status,
      browserName: this.playwrightRunner.executable?.name
        ?? (this.playwrightRunner.config.shareEdge ? 'edge' : this.playwrightRunner.config.cdpEndpoint ? 'cdp' : null),
      browserSession: this.playwrightRunner.session,
    };
  }
}
