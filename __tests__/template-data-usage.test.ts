import { templateDataUsage } from '../src/index';
import { runCli } from '../src/extensions/cli';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { deck, shape, table } from './helpers/lint-deck';

const data = {
  kunde: { name: 'Ada', firma: 'ACME', ustId: 'DE1', adresse: { ort: 'Bonn', plz: '53111' } },
  produkte: [
    { name: 'A', preis: 1, sku: 'x' },
    { name: 'B', preis: 2, rabatt: 0.1 },
  ],
  datum: '2026-11-13',
  intern: { notiz: 'n', id: 1 },
};

/** Usage of a one-slide deck whose shapes have one paragraph per tag. */
const usageOf = async (
  tags: string[],
  options: Parameters<typeof templateDataUsage>[1] = { data },
) => templateDataUsage(await deck([shape('Text', tags.map((t) => [t]))]), options);

const paths = (entries: { path: string }[]) => entries.map((e) => e.path);

test('lists used, unused and missing paths', async () => {
  const usage = await usageOf(['{{kunde.name}}', '{{kunde.nmae}}', '{{datum}}']);
  expect(paths(usage.used)).toEqual(['kunde.name', 'datum']);
  expect(paths(usage.missing)).toEqual(['kunde.nmae']);
  // A subtree nothing reads is one entry
  expect(usage.unused).toEqual([
    'kunde.firma',
    'kunde.ustId',
    'kunde.adresse',
    'produkte',
    'intern',
  ]);
});

test('traces loop variables back to the list', async () => {
  const usage = await usageOf([
    '{{FOR p IN produkte}}',
    '{{p.name}}: {{$p.preis}} ({{$idx}})',
    '{{END-FOR}}',
  ]);
  expect(paths(usage.used)).toEqual(['produkte', 'produkte[].name', 'produkte[].preis']);
  // Fields of any entry count: sku is only in the first, rabatt in the second
  expect(usage.unused).toEqual([
    'kunde',
    'produkte[].sku',
    'produkte[].rabatt',
    'datum',
    'intern',
  ]);
});

test('nested loops and EXEC names', async () => {
  const nested = {
    gruppen: [{ titel: 'G', posten: [{ text: 't', betrag: 1 }] }],
    kunde: { adresse: { ort: 'Bonn', plz: '1' } },
  };
  const usage = await usageOf(
    [
      '{{EXEC a = kunde.adresse}}',
      '{{a.ort}}',
      '{{FOR g IN gruppen}}{{FOR x IN g.posten}}{{x.betrag}}{{END-FOR}}{{END-FOR}}',
    ],
    { data: nested },
  );
  expect(paths(usage.used)).toEqual([
    'kunde.adresse',
    'kunde.adresse.ort',
    'gruppen',
    'gruppen[].posten',
    'gruppen[].posten[].betrag',
  ]);
  // In the order of the data
  expect(usage.unused).toEqual([
    'gruppen[].titel',
    'gruppen[].posten[].text',
    'kunde.adresse.plz',
  ]);
});

test('a value tested by IF is read, its fields are not', async () => {
  const usage = await usageOf(['{{IF kunde}}x{{END-IF}}', '{{IF intern.id > 0}}y{{END-IF}}']);
  expect(usage.unused).toContain('kunde.name');
  expect(usage.unused).toContain('intern.notiz');
  expect(usage.unused).not.toContain('intern.id');
});

test('a value output or passed on counts as a whole', async () => {
  const fmt = (v: unknown) => String(v);
  const usage = await usageOf(['{{fmt(kunde)}}', '{{produkte.map(p => p.name)}}'], {
    data,
    additionalJsContext: { fmt },
  });
  // Helpers from additionalJsContext are not data
  expect(paths(usage.used)).toEqual(['kunde', 'produkte']);
  expect(usage.unused).toEqual(['datum', 'intern']);
});

test('length reads the list, not its entries', async () => {
  const usage = await usageOf(['{{produkte.length}} Produkte']);
  expect(paths(usage.used)).toEqual(['produkte']);
  // No field of any entry is read: one entry for all of them
  expect(usage.unused).toEqual(['kunde', 'produkte[]', 'datum', 'intern']);
});

test('a computed key counts the whole object', async () => {
  const usage = await usageOf(['{{kunde[datum]}}']);
  expect(paths(usage.used)).toEqual(['kunde', 'datum']);
  expect(usage.unused).not.toContain('kunde.name');
});

test('reports where each path is read', async () => {
  const pptx = await deck([
    shape('Titel', [['{{kunde.name}}']]),
    table('Tabelle', [['{{kunde.name}}', '{{datum}}']]),
  ]);
  const usage = await templateDataUsage(pptx, { data });
  expect(usage.used[0]).toEqual({
    path: 'kunde.name',
    locations: [
      { part: 'ppt/slides/slide1.xml', slide: 1, shape: 'Titel' },
      { part: 'ppt/slides/slide2.xml', slide: 2, shape: 'Tabelle' },
    ],
  });
});

test('works with a JSON Schema', async () => {
  const schema = {
    type: 'object',
    properties: {
      kunde: { $ref: '#/$defs/kunde' },
      positionen: { type: 'array', items: { type: 'object', properties: { text: {}, menge: {} } } },
      meta: { type: 'object' },
    },
    $defs: { kunde: { type: 'object', properties: { name: { type: 'string' }, mail: { type: 'string' } } } },
  };
  const usage = await usageOf(
    ['{{kunde.name}}', '{{FOR p IN positionen}}{{p.menge}}{{END-FOR}}', '{{meta.version}}', '{{kunde.tel}}'],
    { schema },
  );
  // meta has no listed properties: any key is fine
  expect(paths(usage.used)).toEqual(['kunde.name', 'positionen', 'positionen[].menge', 'meta.version']);
  expect(paths(usage.missing)).toEqual(['kunde.tel']);
  expect(usage.unused).toEqual(['kunde.mail', 'positionen[].text']);
});

test('without data or schema, only lists what is read', async () => {
  const usage = await usageOf(['{{FOR p IN liste}}{{p.a}}{{END-FOR}}', '{{x["b c"]}}'], {});
  expect(paths(usage.used)).toEqual(['liste', 'liste[].a', 'x["b c"]']);
  expect(usage).toMatchObject({ unused: [], missing: [] });
});

test('tags in layouts are not counted', async () => {
  const pptx = await deck([shape('T', [['{{datum}}']])], {
    'ppt/slideLayouts/slideLayout1.xml': `<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree>${shape('L', [['{{kunde.name}}']])}</p:spTree></p:cSld></p:sldLayout>`,
  });
  const usage = await templateDataUsage(pptx, { data });
  expect(paths(usage.used)).toEqual(['datum']);
});

describe('more tag kinds and expressions', () => {
  test('ELSE-IF tests like IF; IMAGE uses its value as a whole', async () => {
    const usage = await usageOf(['{{IF kunde.firma}}a{{ELSE-IF intern}}b{{END-IF}}', '{{IMAGE kunde.adresse}}'], {
      data,
    });
    expect(paths(usage.used)).toEqual(['kunde.firma', 'intern', 'kunde.adresse']);
    // intern is only tested; the image gets the whole address
    expect(usage.unused).toEqual(['kunde.name', 'kunde.ustId', 'produkte', 'datum', 'intern.notiz', 'intern.id']);
  });

  test('EXEC inside a loop traces the loop variable', async () => {
    const usage = await usageOf([
      '{{FOR p IN produkte}}',
      '{{EXEC x = p.preis}}{{EXEC summe = p.rabatt * 2}}',
      '{{x}} {{summe}}',
      '{{END-FOR}}',
    ]);
    expect(paths(usage.used)).toEqual(['produkte', 'produkte[].preis', 'produkte[].rabatt']);
    expect(usage.unused).toEqual(['kunde', 'produkte[].name', 'produkte[].sku', 'datum', 'intern']);
  });

  test('EXEC names stay bound after the loop, and alias each other', async () => {
    const usage = await usageOf(['{{EXEC k = kunde}}{{EXEC a = k.adresse}}', '{{a.plz}}']);
    expect(paths(usage.used)).toEqual(['kunde', 'kunde.adresse', 'kunde.adresse.plz']);
    expect(usage.unused).toContain('kunde.adresse.ort');
    expect(usage.unused).toContain('kunde.name');
  });

  test('||, ?? and the branches of ?: pass the value on; comparisons do not', async () => {
    const passed = await usageOf([
      "{{kunde.adresse || kunde.firma}}",
      "{{intern ?? 'none'}}",
      "{{datum ? produkte : ''}}",
    ]);
    // Everything below adresse, intern and produkte counts as used
    expect(passed.unused).toEqual(['kunde.name', 'kunde.ustId']);

    const compared = await usageOf([
      "{{kunde.adresse == null ? 'x' : 'y'}}",
      '{{IF !intern}}x{{END-IF}}',
      "{{produkte != undefined && 'ja'}}",
    ]);
    expect(compared.unused).toEqual([
      'kunde.name',
      'kunde.firma',
      'kunde.ustId',
      'kunde.adresse.ort',
      'kunde.adresse.plz',
      'produkte[]',
      'datum',
      'intern.notiz',
      'intern.id',
    ]);
  });

  test('a method on a string reads the string, not more', async () => {
    const usage = await usageOf(['{{kunde.name.toUpperCase()}}', '{{datum.length}}']);
    expect(paths(usage.used)).toEqual(['kunde.name', 'datum']);
    expect(usage.missing).toEqual([]);
  });
});

describe('more schemas and data', () => {
  test('allOf, anyOf and oneOf branches all contribute fields', async () => {
    const schema = {
      type: 'object',
      properties: {
        kunde: {
          allOf: [
            { type: 'object', properties: { name: {} } },
            { type: 'object', properties: { mail: {} } },
          ],
        },
        kontakt: {
          anyOf: [
            { type: 'object', properties: { tel: {} } },
            { type: 'object', properties: { fax: {} } },
          ],
        },
        wert: { oneOf: [{ type: 'string' }, { type: 'number' }] },
      },
    };
    const usage = await usageOf(['{{kunde.mail}}', '{{kontakt.fax}}', '{{kunde.tel}}'], { schema });
    expect(paths(usage.used)).toEqual(['kunde.mail', 'kontakt.fax']);
    expect(paths(usage.missing)).toEqual(['kunde.tel']);
    expect(usage.unused).toEqual(['kunde.name', 'kontakt.tel', 'wert']);
  });

  test('additionalProperties: true next to listed fields', async () => {
    const schema = {
      type: 'object',
      properties: {
        meta: { type: 'object', properties: { version: {}, autor: {} }, additionalProperties: true },
      },
    };
    const usage = await usageOf(['{{meta.version}}', '{{meta.sonstiges.x}}'], { schema });
    // Unlisted keys are allowed: used, not missing
    expect(paths(usage.used)).toEqual(['meta.version', 'meta.sonstiges.x']);
    expect(usage.missing).toEqual([]);
    expect(usage.unused).toEqual(['meta.autor']);
  });

  test('recursive schemas stop instead of looping', async () => {
    const schema = {
      type: 'object',
      properties: { baum: { $ref: '#/$defs/knoten' } },
      $defs: {
        knoten: {
          type: 'object',
          properties: { name: {}, kinder: { type: 'array', items: { $ref: '#/$defs/knoten' } } },
        },
      },
    };
    const usage = await usageOf(['{{baum.kinder[0].kinder[0].name}}'], { schema });
    expect(paths(usage.used)).toEqual(['baum.kinder[].kinder[].name']);
    expect(usage.unused).toContain('baum.name');
  });

  test('data and schema together: fields of both count', async () => {
    const usage = await usageOf(['{{kunde.name}}'], {
      data: { kunde: { name: 'Ada' } },
      schema: { type: 'object', properties: { kunde: { type: 'object', properties: { mail: {} } } } },
    });
    expect(usage.unused).toEqual(['kunde.mail']);
    expect(usage.missing).toEqual([]);
  });

  test('keys that are not names are quoted', async () => {
    const usage = await usageOf(['{{kunde["b c"]}}', "{{kunde['ü-1']}}"], {
      data: { kunde: { 'b c': 1, 'x.y': 2, 'ü-1': 3, 'äb': 4 } },
    });
    expect(paths(usage.used)).toEqual(['kunde["b c"]', 'kunde["ü-1"]']);
    // Letters outside ASCII are names: no quotes
    expect(usage.unused).toEqual(['kunde["x.y"]', 'kunde.äb']);
  });

  test('a list index reads the entries like a loop', async () => {
    const usage = await usageOf(['{{produkte[0].name}}', "{{produkte[1]['preis']}}"]);
    expect(paths(usage.used)).toEqual(['produkte[].name', 'produkte[].preis']);
    expect(usage.unused).toContain('produkte[].sku');
  });
});

describe('CLI', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-usage-'));
  const run = async (...args: string[]) => {
    let stdout = '';
    let stderr = '';
    const code = await runCli(args, {
      stdout: (t) => (stdout += t),
      stderr: (t) => (stderr += t),
    });
    return { code, stdout, stderr };
  };
  const dataFile = path.join(dir, 'data.json');
  fs.writeFileSync(dataFile, JSON.stringify(data));
  let first: string;
  let second: string;
  beforeAll(async () => {
    first = path.join(dir, 'first.pptx');
    second = path.join(dir, 'second.pptx');
    fs.writeFileSync(
      first,
      await deck([shape('Titel', [['{{kunde.name}}'], ['{{kunde.nmae}}']])]),
    );
    fs.writeFileSync(second, await deck([shape('Fuß', [['{{datum}}'], ['{{kunde.name}}']])]));
  });

  test('prints used, unused and missing per file', async () => {
    const { code, stdout } = await run('usage', first, '--data', dataFile);
    expect(code).toBe(0);
    expect(stdout).toBe(
      `${first}\n` +
        '  used (1)\n' +
        '    kunde.name  slide 1 › "Titel"\n' +
        '  unused (6)\n' +
        '    kunde.firma\n' +
        '    kunde.ustId\n' +
        '    kunde.adresse\n' +
        '    produkte\n' +
        '    datum\n' +
        '    intern\n' +
        '  missing in data (1)\n' +
        '    kunde.nmae  slide 1 › "Titel"\n',
    );
  });

  test('with several files, adds what none of them reads', async () => {
    const { stdout } = await run('usage', first, second, '--data', dataFile);
    expect(stdout).toContain(`${second}\n  used (2)\n`);
    expect(stdout.endsWith(
      '\nUnused in all 2 file(s) (5)\n' +
        '  kunde.firma\n' +
        '  kunde.ustId\n' +
        '  kunde.adresse\n' +
        '  produkte\n' +
        '  intern\n',
    )).toBe(true);
  });

  test('--json', async () => {
    const { code, stdout } = await run('usage', '--json', first, '--data', dataFile);
    expect(code).toBe(0);
    const report = JSON.parse(stdout);
    expect(report.files[first].missing).toEqual([
      { path: 'kunde.nmae', locations: [{ part: 'ppt/slides/slide1.xml', slide: 1, shape: 'Titel' }] },
    ]);
    expect(report.unusedInAll).toEqual([
      expect.objectContaining({ data: dataFile, files: [first] }),
    ]);
    expect(report.unusedInAll[0].unused).toContain('intern');
  });

  test('without data, says that unused fields need it', async () => {
    const { stdout } = await run('usage', first);
    expect(stdout).toContain('kunde.nmae');
    expect(stdout).toContain('Pass --data or --schema to see unused fields.');
  });

  test('--watch runs again when the data changes', async () => {
    const watchData = path.join(dir, 'watch-data.json');
    fs.writeFileSync(watchData, JSON.stringify({ kunde: { name: 'Ada' } }));
    const controller = new AbortController();
    let stdout = '';
    const done = runCli(['usage', first, '--data', watchData, '--watch'], {
      stdout: (t) => (stdout += t),
      stderr: () => undefined,
      signal: controller.signal,
      watchTiming: { debounceMs: 50 },
    });
    const runs = async (count: number) => {
      const deadline = Date.now() + 8000;
      while ((stdout.match(/Watching 2 file/g) ?? []).length < count) {
        if (Date.now() > deadline) throw new Error(`Timed out:\n${stdout}`);
        await new Promise((r) => setTimeout(r, 20));
      }
    };
    await runs(1);
    expect(stdout).toContain('unused: none');
    fs.writeFileSync(watchData, JSON.stringify({ kunde: { name: 'Ada', tel: '1' } }));
    await runs(2);
    expect(stdout).toContain('unused (1)\n    kunde.tel\n');
    controller.abort();
    expect(await done).toBe(0);
  });

  test('invalid calls', async () => {
    expect((await run('usage')).code).toBe(2);
    expect((await run('usage', '--max-warnings', '1', first)).stderr).toContain(
      'Unknown option --max-warnings',
    );
    expect((await run('usage', path.join(dir, 'missing.pptx'))).code).toBe(2);
    expect((await run('usage', '--help')).stdout).toContain('Usage: pptx-templates usage');
  });
});
