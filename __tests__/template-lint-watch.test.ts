import { runCli } from '../src/extensions/cli';
import { EventEmitter } from 'events';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { deck, shape } from './helpers/lint-deck';

/**
 * `pptx-templates lint --watch`. The watch runs until the signal aborts;
 * timings are shortened so the suite stays fast.
 */
const startWatch = (...args: string[]) => {
  const controller = new AbortController();
  const out = { stdout: '', stderr: '' };
  const done = runCli(['lint', '--watch', ...args], {
    stdout: (t) => (out.stdout += t),
    stderr: (t) => (out.stderr += t),
    signal: controller.signal,
    watchTiming: { debounceMs: 50, retries: 3, retryDelayMs: 100 },
  });
  /** Waits until stdout has `count` runs (each ends with the footer). */
  const runs = async (count: number) => {
    const deadline = Date.now() + 8000;
    while ((out.stdout.match(/Watching \d+ file/g) ?? []).length < count) {
      if (Date.now() > deadline) {
        throw new Error(`Timed out waiting for run ${count}:\n${out.stdout}`);
      }
      await new Promise((r) => setTimeout(r, 20));
    }
  };
  /** Output of the last run only. */
  const last = () => {
    const parts = out.stdout.split(/\n\[\d\d:\d\d:\d\d\] Watching[^\n]*\n/);
    return parts[parts.length - 2];
  };
  const stop = async () => {
    controller.abort();
    return done;
  };
  return { out, runs, last, stop, done };
};

/** How PowerPoint saves: write a temporary file, rename it over the deck. */
const saveLikePowerPoint = (file: string, content: Buffer) => {
  const tmp = path.join(path.dirname(file), `PPT${Date.now()}.tmp`);
  fs.writeFileSync(tmp, content);
  fs.renameSync(tmp, file);
};

describe('lint --watch', () => {
  let dir: string;
  let broken: Buffer;
  let clean: Buffer;
  beforeAll(async () => {
    broken = await deck([shape('T', [['{{IF user.vip}}']])]);
    clean = await deck([shape('T', [['{{user.name}}']])]);
  });
  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-lint-watch-'));
  });
  afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));

  test('lints again after a save that replaces the file', async () => {
    const file = path.join(dir, 'deck.pptx');
    fs.writeFileSync(file, broken);
    const watch = startWatch(file);
    await watch.runs(1);
    expect(watch.last()).toContain('(unclosed-block)');
    expect(watch.out.stdout).toContain(
      'Watching 1 file(s) for changes. Press Ctrl+C to stop.',
    );

    saveLikePowerPoint(file, clean);
    await watch.runs(2);
    expect(watch.last()).toBe('No problems found in 1 file(s)\n');

    saveLikePowerPoint(file, broken);
    await watch.runs(3);
    expect(watch.last()).toContain('(unclosed-block)');

    // Errors don't end the watch; stopping it resolves with 0
    expect(await watch.stop()).toBe(0);
    expect(watch.out.stderr).toBe('');
  });

  test('watches several decks in different folders', async () => {
    fs.mkdirSync(path.join(dir, 'de'));
    fs.mkdirSync(path.join(dir, 'en'));
    const de = path.join(dir, 'de', 'deck.pptx');
    const en = path.join(dir, 'en', 'deck.pptx');
    fs.writeFileSync(de, clean);
    fs.writeFileSync(en, clean);
    const watch = startWatch(de, en);
    await watch.runs(1);
    expect(watch.last()).toBe('No problems found in 2 file(s)\n');
    expect(watch.out.stdout).toContain('Watching 2 file(s)');

    // Saving one deck lints both: the summary stays complete
    saveLikePowerPoint(en, broken);
    await watch.runs(2);
    expect(watch.last()).toContain(en);
    expect(watch.last()).not.toContain(de);
    expect(watch.last()).toContain('1 problem(s): 1 error(s), 0 warning(s)');
    await watch.stop();
  });

  test('ignores other files in the directory', async () => {
    const file = path.join(dir, 'deck.pptx');
    fs.writeFileSync(file, clean);
    const watch = startWatch(file);
    await watch.runs(1);
    // PowerPoint's owner file while the deck is open
    fs.writeFileSync(path.join(dir, '~$deck.pptx'), 'owner');
    fs.writeFileSync(path.join(dir, 'other.pptx'), broken);
    await new Promise((r) => setTimeout(r, 300));
    await watch.stop();
    expect(watch.out.stdout.match(/Watching \d+ file/g)).toHaveLength(1);
  });

  test('reloads the --data file when it changes', async () => {
    const file = path.join(dir, 'deck.pptx');
    const data = path.join(dir, 'data.json');
    fs.writeFileSync(file, clean);
    fs.writeFileSync(data, JSON.stringify({ user: { nmae: 'Ada' } }));
    const watch = startWatch(file, '--data', data);
    await watch.runs(1);
    expect(watch.last()).toContain('(unknown-property)');
    expect(watch.out.stdout).toContain('Watching 2 file(s)');

    fs.writeFileSync(data, JSON.stringify({ user: { name: 'Ada' } }));
    await watch.runs(2);
    expect(watch.last()).toBe('No problems found in 1 file(s)\n');
    await watch.stop();
  });

  test('retries a file caught mid-save instead of reporting it', async () => {
    const file = path.join(dir, 'deck.pptx');
    fs.writeFileSync(file, clean);
    const watch = startWatch(file);
    await watch.runs(1);

    // Half-written ZIP, complete 150 ms later: after the debounce, before
    // the retries give up
    fs.writeFileSync(file, clean.subarray(0, 100));
    setTimeout(() => fs.writeFileSync(file, broken), 150);
    await watch.runs(2);
    expect(watch.last()).toContain('(unclosed-block)');
    expect(watch.out.stderr).toBe('');
    await watch.stop();
  });

  test('reports an unreadable file and keeps watching', async () => {
    const file = path.join(dir, 'deck.pptx');
    fs.writeFileSync(file, clean);
    const watch = startWatch(file);
    await watch.runs(1);

    fs.writeFileSync(file, 'not a zip');
    await watch.runs(2);
    expect(watch.out.stderr).toContain(`Cannot read ${file}:`);

    fs.writeFileSync(file, clean);
    await watch.runs(3);
    expect(watch.last()).toBe('No problems found in 1 file(s)\n');
    await watch.stop();
  });

  test('clears the screen before each run when asked', async () => {
    const file = path.join(dir, 'deck.pptx');
    fs.writeFileSync(file, clean);
    const controller = new AbortController();
    let stdout = '';
    const done = runCli(['lint', '-w', file], {
      stdout: (t) => (stdout += t),
      stderr: () => undefined,
      clearScreen: true,
      signal: controller.signal,
      watchTiming: { debounceMs: 50 },
    });
    while (!stdout.includes('Watching')) {
      await new Promise((r) => setTimeout(r, 20));
    }
    controller.abort();
    await done;
    expect(stdout.startsWith('\u001b[2J\u001b[3J\u001b[H')).toBe(true);
  });

  test('control characters in a broken --data file do not reach the terminal', async () => {
    const file = path.join(dir, 'deck.pptx');
    const data = path.join(dir, 'data.json');
    fs.writeFileSync(file, clean);
    // JSON.parse quotes the content in its error message
    fs.writeFileSync(data, '\u001b[2J not json');
    const watch = startWatch(file, '--data', data);
    await watch.runs(1);
    await watch.stop();
    expect(watch.out.stderr).toContain(`Cannot read data ${data}:`);
    expect(watch.out.stderr).not.toContain('\u001b');
  });

  test('ends with exit code 2 when the watcher fails', async () => {
    const file = path.join(dir, 'deck.pptx');
    fs.writeFileSync(file, clean);
    const fake = Object.assign(new EventEmitter(), { close: jest.fn() });
    const spy = jest
      .spyOn(fs, 'watch')
      .mockReturnValue(fake as unknown as fs.FSWatcher);
    try {
      const watch = startWatch(file);
      await watch.runs(1);
      fake.emit('error', new Error('EPERM: operation not permitted, watch'));
      // Resolves without the signal
      expect(await watch.done).toBe(2);
      expect(watch.out.stderr).toContain('Watching failed: EPERM');
      expect(fake.close).toHaveBeenCalled();
    } finally {
      spy.mockRestore();
    }
  });

  test('rejects --watch with --json, and a missing directory', async () => {
    let stderr = '';
    const io = { stdout: () => undefined, stderr: (t: string) => (stderr += t) };
    expect(await runCli(['lint', '--watch', '--json', 'deck.pptx'], io)).toBe(2);
    expect(stderr).toContain('--watch cannot be combined with --json');

    stderr = '';
    const missing = path.join(dir, 'no-such-dir', 'deck.pptx');
    expect(await runCli(['lint', '--watch', missing], io)).toBe(2);
    expect(stderr).toContain('Cannot watch:');
  });
});
