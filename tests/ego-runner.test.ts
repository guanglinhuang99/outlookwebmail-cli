import { describe, expect, it } from 'vitest';
import { EgoRunner, parseMarkedResult } from '../src/browser/ego-runner.js';
import { AppError } from '../src/util/errors.js';

describe('parseMarkedResult', () => {
  it('returns the last marked JSON result', () => {
    const stdout = [
      'ordinary log',
      '{"__webmail_result__":true,"result":{"value":1}}',
      'another log',
      '{"__webmail_result__":true,"result":{"value":2}}',
    ].join('\n');

    expect(parseMarkedResult<{ value: number }>(stdout)).toEqual({ value: 2 });
  });

  it('rejects output without a marked result', () => {
    expect(() => parseMarkedResult('not json\n{"result":1}')).toThrow(AppError);
  });
});

describe('EgoRunner', () => {
  const runner = new EgoRunner({ command: process.execPath, args: [] });

  it('runs a stdin Node script and parses its result', async () => {
    const result = await runner.run<{ ok: boolean }>(
      'console.log(JSON.stringify({__webmail_result__:true,result:{ok:true}}));',
      2_000,
    );
    expect(result.value).toEqual({ ok: true });
  });

  it('accepts cliLog-style marked results from stderr', async () => {
    const result = await runner.run<number>(
      'console.error(JSON.stringify({__webmail_result__:true,result:42}));',
      2_000,
    );
    expect(result.value).toBe(42);
  });

  it('maps a non-zero exit to EGO_BROWSER_ERROR', async () => {
    await expect(runner.run('console.error("boom"); process.exitCode = 7;', 2_000)).rejects.toMatchObject({
      code: 'EGO_BROWSER_ERROR',
    });
  });

  it('terminates a timed-out child', async () => {
    await expect(runner.run('setInterval(() => {}, 1000);', 50)).rejects.toMatchObject({
      code: 'TIMEOUT',
    });
  });
});
