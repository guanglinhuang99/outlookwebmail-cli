import { describe, expect, it, vi } from 'vitest';
import { defaultEdgeUserDataDir, discoverSharedEdgeEndpoint } from '../src/browser/shared-edge.js';

describe('shared Edge discovery', () => {
  it('uses the default Edge user data directory on macOS and Windows', () => {
    expect(defaultEdgeUserDataDir({ platform: 'darwin', home: '/Users/test', env: {} }))
      .toBe('/Users/test/Library/Application Support/Microsoft Edge');
    expect(defaultEdgeUserDataDir({ platform: 'win32', home: 'C:\\Users\\test', env: { LOCALAPPDATA: 'C:\\Local' } }))
      .toContain('Microsoft');
  });

  it('reads the remote debugging port and only returns a loopback endpoint', async () => {
    const readFileFn = vi.fn().mockResolvedValue('57652\n/devtools/browser/example\n');
    await expect(discoverSharedEdgeEndpoint({
      platform: 'darwin', home: '/Users/test', env: {}, readFileFn,
    })).resolves.toBe('ws://127.0.0.1:57652/devtools/browser/example');
    expect(readFileFn).toHaveBeenCalledWith(
      '/Users/test/Library/Application Support/Microsoft Edge/DevToolsActivePort', 'utf8',
    );
  });

  it('reports how to enable remote debugging when the port file is absent', async () => {
    const missing = Object.assign(new Error('missing'), { code: 'ENOENT' });
    await expect(discoverSharedEdgeEndpoint({
      platform: 'linux', home: '/home/test', env: {}, readFileFn: vi.fn().mockRejectedValue(missing),
    })).rejects.toMatchObject({
      code: 'SHARED_EDGE_NOT_AVAILABLE',
      message: expect.stringContaining('edge://inspect'),
    });
  });

  it('rejects malformed ports', async () => {
    await expect(discoverSharedEdgeEndpoint({
      platform: 'darwin', home: '/Users/test', env: {}, readFileFn: vi.fn().mockResolvedValue('invalid\n'),
    })).rejects.toMatchObject({ code: 'SHARED_EDGE_NOT_AVAILABLE' });
  });
});
