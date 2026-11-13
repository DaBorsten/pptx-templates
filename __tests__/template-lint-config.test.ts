import { runCli } from '../src/extensions/cli';
import { readConfig, resolveConfig } from '../src/extensions/template-lint/config';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { deck, shape } from './helpers/lint-deck';

/**
 * A typical layout: per customer folder a cover and a report deck,
 * each in German and English; some customers have no report.
 * Each deck kind and language has its own example data.
 */
const makeProject = async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-config-'));
  const write = (file: string, content: string | Buffer) => {
    fs.mkdirSync(path.dirname(path.join(dir, file)), { recursive: true });
    fs.writeFileSync(path.join(dir, file), content);
  };
  const first = {
    de: await deck([shape('Titel', [['{{titel}}']])]),
    en: await deck([shape('Title', [['{{title}}']])]),
  };
  const projects = {
    de: await deck([shape('Liste', [['{{FOR p IN projekte}}{{p.name}}{{END-FOR}}']])]),
    en: await deck([shape('List', [['{{FOR p IN projects}}{{p.name}}{{END-FOR}}']])]),
  };
  for (const folder of ['customer-a', 'customer-b']) {
    write(`${folder}/cover_de.pptx`, first.de);
    write(`${folder}/cover_en.pptx`, first.en);
  }
  write('customer-a/report_de.pptx', projects.de);
  write('customer-a/report_en.pptx', projects.en);
  // PowerPoint's owner file while a deck is open: never a template
  write('customer-a/~$cover_de.pptx', 'owner');

  write('data/cover_de.json', JSON.stringify({ titel: 'Hallo', intern: 1 }));
  write('data/cover_en.json', JSON.stringify({ title: 'Hello' }));
  write('data/report_de.json', JSON.stringify({ projekte: [{ name: 'A', jahr: 1 }] }));
  write('data/report_en.json', JSON.stringify({ projects: [{ name: 'A' }] }));

  const config = {
    templates: [
      { files: '{customer}/cover_{lang}.pptx', data: 'data/cover_{lang}.json' },
      { files: '{customer}/report_{lang}.pptx', data: 'data/report_{lang}.json' },
    ],
  };
  const configFile = path.join(dir, 'pptx-templates.json');
  fs.writeFileSync(configFile, JSON.stringify(config));
  return { dir, write, configFile, first, projects };
};

const run = async (...args: string[]) => {
  let stdout = '';
  let stderr = '';
  const code = await runCli(args, {
    stdout: (t) => (stdout += t),
    stderr: (t) => (stderr += t),
  });
  return { code, stdout, stderr };
};

const writeConfig = (file: string, config: unknown) =>
  fs.writeFileSync(file, JSON.stringify(config));

describe('resolveConfig', () => {
  test('pairs each deck with the data of its kind and language', async () => {
    const { dir, configFile } = await makeProject();
    const { jobs, unmatched } = resolveConfig(readConfig(configFile), configFile);
    const rel = (p: string) => path.relative(dir, p).split(path.sep).join('/');
    expect(jobs.map((j) => [rel(j.file), rel(j.dataFile)])).toEqual([
      ['customer-a/cover_de.pptx', 'data/cover_de.json'],
      ['customer-a/cover_en.pptx', 'data/cover_en.json'],
      ['customer-a/report_de.pptx', 'data/report_de.json'],
      ['customer-a/report_en.pptx', 'data/report_en.json'],
      // customer-b has no report: not an error
      ['customer-b/cover_de.pptx', 'data/cover_de.json'],
      ['customer-b/cover_en.pptx', 'data/cover_en.json'],
    ]);
    expect(unmatched).toEqual([]);
  });

  test('a placeholder used twice must match the same text', async () => {
    const { dir, write } = await makeProject();
    write('de/deck_de.pptx', 'x');
    write('de/deck_en.pptx', 'x');
    const file = path.join(dir, 'c.json');
    writeConfig(file, { templates: [{ files: '{lang}/deck_{lang}.pptx' }] });
    const { jobs } = resolveConfig(readConfig(file), file);
    expect(jobs.map((j) => path.basename(j.file))).toEqual(['deck_de.pptx']);
  });

  test('characters with a regex meaning in names are literal text', async () => {
    const { dir, write } = await makeProject();
    // Only characters every file system allows (no ? or \ on Windows)
    const names = ['deck (1)+[a]^$.pptx', 'v1.2-final.pptx'];
    names.forEach((name) => write(`lit/${name}`, 'x'));
    // Look-alikes that a regex without escaping would match
    write('lit/deck 1+a^.pptx', 'x');
    write('lit/v1x2-finalxpptx', 'x');
    const file = path.join(dir, 'c.json');
    writeConfig(file, { templates: [{ files: names.map((n) => `lit/${n}`) }] });
    const { jobs, unmatched } = resolveConfig(readConfig(file), file);
    expect(jobs.map((j) => path.basename(j.file)).sort()).toEqual([...names].sort());
    expect(unmatched).toEqual([]);
  });

  test('** matches any number of folders; hidden folders are skipped', async () => {
    const { dir, write } = await makeProject();
    write('a/b/c/deep.pptx', 'x');
    write('top.pptx', 'x');
    write('.git/x.pptx', 'x');
    const file = path.join(dir, 'c.json');
    writeConfig(file, { templates: [{ files: ['**/deep.pptx', '**/top.pptx', '**/x.pptx'] }] });
    const { jobs, unmatched } = resolveConfig(readConfig(file), file);
    expect(jobs.map((j) => path.basename(j.file))).toEqual(['deep.pptx', 'top.pptx']);
    expect(unmatched).toEqual(['**/x.pptx']);
  });

  test('rejects mistakes with the file and entry named', async () => {
    const { dir } = await makeProject();
    const file = path.join(dir, 'c.json');
    const error = (config: unknown) => {
      writeConfig(file, config);
      try {
        resolveConfig(readConfig(file), file);
      } catch (e) {
        return (e as Error).message;
      }
      return 'no error';
    };
    expect(error({ templates: [] })).toBe(`${file}: "templates" must be a non-empty list`);
    expect(error({ templates: [{ file: 'x.pptx' }] })).toContain(
      'templates[0]: unknown key "file" (allowed: files, data, schema)',
    );
    expect(error({ templates: [{ files: '*.pptx', data: 'data/{lang}.json' }] })).toContain(
      'templates[0]: "data" uses {lang}, but "files" does not define it',
    );
    expect(
      error({ templates: [{ files: ['{lang}/a.pptx', 'b.pptx'], data: '{lang}.json' }] }),
    ).toContain('does not define it in every pattern');
    expect(error({ templates: [{ files: 'a/**' }] })).toContain('must end with a file name pattern');
    expect(error({ templates: [{ files: '{1x}/a.pptx' }] })).toContain(
      '"{1x}" is not a valid placeholder name',
    );
    expect(
      error({ templates: [{ files: 'customer-a/cover_de.pptx' }, { files: '*/cover_*.pptx' }] }),
    ).toContain('matches templates[0] and templates[1]');
  });
});

describe('lint --config', () => {
  test('lints each deck against its own data', async () => {
    const { dir, configFile, write } = await makeProject();
    // The German cover uses the English name: only it is wrong
    write('customer-b/cover_de.pptx', await deck([shape('Titel', [['{{title}}']])]));
    const { code, stdout, stderr } = await run('lint', '--config', configFile);
    expect(stderr).toBe('');
    expect(code).toBe(0); // warnings only
    expect(stdout).toContain(path.join(dir, 'customer-b', 'cover_de.pptx'));
    expect(stdout).toContain('"title" is not in the data (did you mean "titel"?)');
    expect(stdout).toContain('1 problem(s): 0 error(s), 1 warning(s)');
  });

  test('uses pptx-templates.json from the current folder', async () => {
    const { dir } = await makeProject();
    const cwd = process.cwd();
    process.chdir(dir);
    try {
      const { code, stdout } = await run('lint');
      expect(code).toBe(0);
      expect(stdout).toBe('No problems found in 6 file(s)\n');
    } finally {
      process.chdir(cwd);
    }
  });

  test('notes patterns that match nothing; nothing at all is an error', async () => {
    const { dir, configFile } = await makeProject();
    writeConfig(configFile, {
      templates: [{ files: '*/cover_*.pptx' }, { files: '*/agenda_*.pptx' }],
    });
    const some = await run('lint', '-c', configFile);
    expect(some.code).toBe(0);
    expect(some.stderr).toBe('No file matches "*/agenda_*.pptx".\n');

    writeConfig(configFile, { templates: [{ files: '*/agenda_*.pptx' }] });
    const none = await run('lint', '-c', configFile);
    expect(none.code).toBe(2);
    expect(none.stderr).toContain(`${configFile}: no file matches its patterns`);
    fs.rmSync(dir, { recursive: true, force: true });
  });

  test('a missing data file is reported', async () => {
    const { dir, configFile } = await makeProject();
    fs.rmSync(path.join(dir, 'data', 'report_en.json'));
    const { code, stderr } = await run('lint', '-c', configFile);
    expect(code).toBe(2);
    expect(stderr).toContain(`Cannot read data ${path.join(dir, 'data', 'report_en.json')}`);
  });

  test('--config cannot be combined with files or --data', async () => {
    const { stderr, code } = await run('lint', '-c', 'x.json', 'deck.pptx');
    expect(code).toBe(2);
    expect(stderr).toContain('--config names the files and their data');
  });
});

describe('usage --config', () => {
  test('lists unused fields per data source', async () => {
    const { configFile } = await makeProject();
    const { code, stdout } = await run('usage', '-c', configFile);
    expect(code).toBe(0);
    // Both German covers read titel; neither reads intern
    expect(stdout).toMatch(/Unused in all 2 file\(s\) using \S*cover_de\.json \(1\)\n {2}intern\n/);
    // Only one deck per report data: no "unused in all" for it
    expect(stdout).not.toContain('report_de.json (');
    expect(stdout).toContain('Every field is used in at least one of the 2 file(s) using');
  });
});

describe('lint --config --watch', () => {
  test('picks up a new folder and a changed data file', async () => {
    const { dir, write, first } = await makeProject();
    const configFile = path.join(dir, 'pptx-templates.json');
    const controller = new AbortController();
    let stdout = '';
    const done = runCli(['lint', '-c', configFile, '--watch'], {
      stdout: (t) => (stdout += t),
      stderr: () => undefined,
      signal: controller.signal,
      watchTiming: { debounceMs: 50 },
    });
    const runs = async (count: number) => {
      const deadline = Date.now() + 8000;
      while ((stdout.match(/Watching \d+ file/g) ?? []).length < count) {
        if (Date.now() > deadline) throw new Error(`Timed out:\n${stdout}`);
        await new Promise((r) => setTimeout(r, 20));
      }
    };
    await runs(1);
    expect(stdout).toContain('No problems found in 6 file(s)');

    // A new folder, then a deck in it
    fs.mkdirSync(path.join(dir, 'customer-c'));
    await runs(2);
    write('customer-c/cover_en.pptx', first.en);
    await runs(3);
    expect(stdout).toContain('No problems found in 7 file(s)');

    // The English data loses "title": the English covers are wrong now
    write('data/cover_en.json', JSON.stringify({ name: 'x' }));
    await runs(4);
    expect(stdout).toContain('3 problem(s): 0 error(s), 3 warning(s)');

    controller.abort();
    expect(await done).toBe(0);
  });
});
