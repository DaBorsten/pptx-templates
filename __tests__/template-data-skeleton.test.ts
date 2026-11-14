import Automizer, {
  lintTemplate,
  mergeDataSkeleton,
  modify,
  templateDataSkeleton,
} from '../src/index';
import { expectXml } from './helpers/expect-xml';
import { runCli } from '../src/extensions/cli';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { deck, shape, table } from './helpers/lint-deck';

/** A one-slide deck whose shape has one paragraph per tag. */
const deckOf = (tags: string[]) => deck([shape('Text', tags.map((t) => [t]))]);

const skeletonOf = async (tags: string[]) => templateDataSkeleton(await deckOf(tags));

test('builds objects, lists and placeholder values from the tags', async () => {
  const { data, functions, errorCount } = await skeletonOf([
    '{{kunde.name}}, {{kunde.adresse.ort}}',
    '{{IF kunde.vip}}VIP{{END-IF}}',
    '{{FOR p IN produkte}}',
    '{{p.name}}: {{euro(p.preis)}} {{p.menge.toFixed(0)}}',
    '{{END-FOR}}',
    '{{datum.toUpperCase()}}',
  ]);
  expect(data).toEqual({
    kunde: { name: 'kunde.name', adresse: { ort: 'kunde.adresse.ort' }, vip: true },
    produkte: [{ name: 'produkte[].name', preis: 'produkte[].preis', menge: 1 }],
    datum: 'datum',
  });
  expect(functions).toEqual(['euro']);
  expect(errorCount).toBe(0);
});

test('a FOR whose entries are output as a whole is a list of values', async () => {
  const { data } = await skeletonOf([
    '{{FOR t IN tags}}{{t}}{{END-FOR}}',
    '{{FOR g IN gruppen}}{{FOR x IN g.posten}}{{x.betrag}}{{END-FOR}}{{END-FOR}}',
    '{{namen.join(", ")}}',
  ]);
  expect(data).toEqual({
    tags: ['tags[]'],
    gruppen: [{ posten: [{ betrag: 'gruppen[].posten[].betrag' }] }],
    namen: ['namen[]'],
  });
});

test('EXEC names, list indices and tables are traced to the data', async () => {
  const { data } = await templateDataSkeleton(
    await deck([
      shape('Text', [['{{EXEC a = kunde.adresse}}'], ['{{a.plz}}'], ['{{posten[0].text}}']]),
      table('Tabelle', [['{{FOR r IN zeilen}}{{r.wert}}'], ['{{END-FOR}}']]),
    ]),
  );
  expect(data).toEqual({
    kunde: { adresse: { plz: 'kunde.adresse.plz' } },
    posten: [{ text: 'posten[].text' }],
    zeilen: [{ wert: 'zeilen[].wert' }],
  });
});

test('combines several decks into one data object', async () => {
  const { data } = await templateDataSkeleton([
    await deckOf(['{{kunde.name}}', '{{datum}}']),
    await deckOf(['{{kunde.firma}}', '{{FOR p IN produkte}}{{p.name}}{{END-FOR}}']),
  ]);
  expect(data).toEqual({
    kunde: { name: 'kunde.name', firma: 'kunde.firma' },
    datum: 'datum',
    produkte: [{ name: 'produkte[].name' }],
  });
});

test('the skeleton lints without warnings', async () => {
  const pptx = await deckOf([
    '{{kunde.name}} {{IF kunde.vip}}!{{END-IF}}',
    '{{FOR p IN produkte}}{{p.name}} {{p.preis.toFixed(2)}}{{END-FOR}}',
    '{{namen.join(", ")}}',
  ]);
  const { data } = await templateDataSkeleton(pptx);
  const { issues } = await lintTemplate(pptx, { data });
  expect(issues).toEqual([]);
});

test('tag errors are counted and skipped', async () => {
  const { data, errorCount } = await skeletonOf(['{{kunde.name}}', "{{IF x == 'a}}"]);
  expect(data).toEqual({ kunde: { name: 'kunde.name' } });
  expect(errorCount).toBeGreaterThan(0);
});

test('fields named like string or list methods stay fields, in any order', async () => {
  const expected = {
    kunde: { link: 'kunde.link', name: 'kunde.name' },
    config: { values: 'config.values', keys: 'config.keys' },
    produkte: [{ link: 'produkte[].link', small: 'produkte[].small' }],
  };
  const tags = [
    '{{kunde.link}} {{kunde.name}}',
    '{{config.values}} {{config.keys}}',
    '{{FOR p IN produkte}}{{p.link}} {{p.small}}{{END-FOR}}',
  ];
  const pptx = await deckOf(tags);
  const { data } = await templateDataSkeleton(pptx);
  expect(data).toEqual(expected);
  expect((await lintTemplate(pptx, { data })).issues).toEqual([]);

  const reversed = await skeletonOf(['{{kunde.name}} {{kunde.link}}', ...tags.slice(1)]);
  expect(reversed.data).toEqual(expected);
});

test('length and called methods describe the value, not a field', async () => {
  const { data, functions } = await skeletonOf([
    '{{IF produkte.length > 0}}x{{END-IF}}',
    '{{FOR p IN produkte}}{{p.name.trim()}}{{END-FOR}}',
    '{{titel.length}}',
  ]);
  expect(data).toEqual({ produkte: [{ name: 'produkte[].name' }], titel: 'titel' });
  expect(functions).toEqual([]);
});

test('operands of a calculation are numbers, operands of + are output', async () => {
  const { data } = await skeletonOf([
    '{{summe * 2}} {{-rabatt}} {{netto - skonto}}',
    '{{vorname + " " + nachname}}',
    '{{IF preis > 100}}x{{END-IF}}',
  ]);
  expect(data).toEqual({
    summe: 1,
    rabatt: 1,
    netto: 1,
    skonto: 1,
    vorname: 'vorname',
    nachname: 'nachname',
    preis: true,
  });
});

test('a method called on an object is a helper, the object stays one', async () => {
  const { data, functions } = await skeletonOf(['{{kunde.format()}}', '{{kunde.name}}']);
  expect(data).toEqual({ kunde: { name: 'kunde.name' } });
  expect(functions).toEqual(['kunde.format']);
  expect((await skeletonOf(['{{kunde.format()}}'])).data).toEqual({ kunde: {} });
});

test('forbidden keys never become data', async () => {
  const { data, errorCount } = await skeletonOf(['{{kunde.__proto__.x}} {{kunde.constructor}}']);
  expect(data).toEqual({});
  expect(errorCount).toBe(2);
});

describe('mergeDataSkeleton', () => {
  test('adds fields named like Object.prototype members', () => {
    const { data, added } = mergeDataSkeleton(
      { kunde: {} },
      { toString: 'toString', kunde: { valueOf: 'kunde.valueOf' } },
    );
    expect(JSON.parse(JSON.stringify(data))).toEqual({
      toString: 'toString',
      kunde: { valueOf: 'kunde.valueOf' },
    });
    expect(added).toEqual(['toString', 'kunde.valueOf']);
    expect(Object.getPrototypeOf(data)).toBe(Object.prototype);
  });

  test('keeps existing values and adds missing fields', () => {
    const { data, added } = mergeDataSkeleton(
      {
        kunde: { name: 'Ada' },
        produkte: [{ name: 'A' }, { name: 'B', preis: 2 }, 'kein Objekt'],
        extra: 1,
      },
      {
        kunde: { name: 'kunde.name', ort: 'kunde.ort' },
        produkte: [{ name: 'produkte[].name', preis: 'produkte[].preis' }],
        datum: 'datum',
      },
    );
    expect(data).toEqual({
      kunde: { name: 'Ada', ort: 'kunde.ort' },
      produkte: [
        { name: 'A', preis: 'produkte[].preis' },
        { name: 'B', preis: 2 },
        'kein Objekt',
      ],
      extra: 1,
      datum: 'datum',
    });
    expect(added).toEqual(['kunde.ort', 'produkte[].preis', 'datum']);
  });

  test('a value of another kind stays as it is', () => {
    const { data, added } = mergeDataSkeleton(
      { kunde: 'Ada', liste: [] },
      { kunde: { name: 'kunde.name' }, liste: [{ a: 'liste[].a' }] },
    );
    expect(data).toEqual({ kunde: 'Ada', liste: [] });
    expect(added).toEqual([]);
  });
});

describe('CLI', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-init-'));
  const run = async (...args: string[]) => {
    let stdout = '';
    let stderr = '';
    const code = await runCli(args, {
      stdout: (t) => (stdout += t),
      stderr: (t) => (stderr += t),
    });
    return { code, stdout, stderr };
  };
  let first: string;
  let second: string;
  beforeAll(async () => {
    first = path.join(dir, 'first.pptx');
    second = path.join(dir, 'second.pptx');
    fs.writeFileSync(first, await deckOf(['{{kunde.name}}', '{{euro(summe)}}']));
    fs.writeFileSync(second, await deckOf(['{{FOR p IN produkte}}{{p.name}}{{END-FOR}}']));
  });

  test('prints the combined data of several decks', async () => {
    const { code, stdout, stderr } = await run('init', first, second);
    expect(code).toBe(0);
    expect(JSON.parse(stdout)).toEqual({
      kunde: { name: 'kunde.name' },
      summe: 'summe',
      produkte: [{ name: 'produkte[].name' }],
    });
    expect(stderr).toContain('Called as functions, so not in the data: euro');
  });

  test('creates the file, then only adds missing fields', async () => {
    const out = path.join(dir, 'data', 'deck.json');
    let result = await run('init', first, '-o', out);
    expect(result.code).toBe(0);
    expect(result.stdout).toContain('created');
    expect(JSON.parse(fs.readFileSync(out, 'utf8'))).toEqual({
      kunde: { name: 'kunde.name' },
      summe: 'summe',
    });

    // The user fills in a value, then a second deck uses the same data
    fs.writeFileSync(out, JSON.stringify({ kunde: { name: 'Ada' }, summe: 12 }));
    result = await run('init', first, second, '-o', out);
    expect(result.stdout).toContain('completed');
    expect(result.stdout).toContain('+ produkte');
    expect(JSON.parse(fs.readFileSync(out, 'utf8'))).toEqual({
      kunde: { name: 'Ada' },
      summe: 12,
      produkte: [{ name: 'produkte[].name' }],
    });

    result = await run('init', first, second, '-o', out);
    expect(result.stdout).toContain('up to date');
    // Written through a temp file that is renamed: none is left behind
    expect(fs.readdirSync(path.dirname(out))).toEqual(['deck.json']);
  });

  test('never overwrites a data file it cannot read', async () => {
    const out = path.join(dir, 'broken.json');
    fs.writeFileSync(out, '{ nope');
    const { code, stderr } = await run('init', first, '-o', out);
    expect(code).toBe(2);
    expect(stderr).toContain('it was not changed');
    expect(fs.readFileSync(out, 'utf8')).toBe('{ nope');
  });

  test('writes every data file of a config, each from its decks', async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-init-config-'));
    fs.mkdirSync(path.join(root, 'kunde-a'));
    fs.mkdirSync(path.join(root, 'kunde-b'));
    fs.writeFileSync(path.join(root, 'kunde-a', 'cover_de.pptx'), await deckOf(['{{titel}}']));
    fs.writeFileSync(path.join(root, 'kunde-b', 'cover_de.pptx'), await deckOf(['{{untertitel}}']));
    fs.writeFileSync(path.join(root, 'kunde-a', 'report_de.pptx'), await deckOf(['{{zahlen}}']));
    const config = path.join(root, 'pptx-templates.json');
    fs.writeFileSync(
      config,
      JSON.stringify({
        templates: [
          { files: '{customer}/cover_{lang}.pptx', data: 'data/cover_{lang}.json' },
          { files: '{customer}/report_{lang}.pptx' },
        ],
      }),
    );
    const { code, stdout, stderr } = await run('init', '--config', config);
    expect(code).toBe(0);
    expect(stdout).toContain('from 2 file(s)');
    expect(stderr).toContain('has no "data"');
    expect(JSON.parse(fs.readFileSync(path.join(root, 'data', 'cover_de.json'), 'utf8'))).toEqual({
      titel: 'titel',
      untertitel: 'untertitel',
    });
  });

  test('rejects files together with --config', async () => {
    const { code, stderr } = await run('init', first, '--config', 'x.json');
    expect(code).toBe(2);
    expect(stderr).toContain('do not pass files or --out');
  });
});

/**
 * TemplateLogic.pptx was saved by PowerPoint itself (not built by a test
 * helper): three slides with IF, a table row with FOR, EXEC, a helper
 * function and a list method.
 */
describe('a template saved by PowerPoint', () => {
  const template = path.join(__dirname, 'pptx-templates', 'TemplateLogic.pptx');
  const out = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-init-real-')), 'data.json');
  const euro = (n: unknown) => `${n} EUR`;

  const run = async (...args: string[]) => {
    let stdout = '';
    let stderr = '';
    const code = await runCli(args, {
      stdout: (t) => (stdout += t),
      stderr: (t) => (stderr += t),
    });
    return { code, stdout, stderr };
  };

  test('init writes the data the tags read', async () => {
    const { code, stdout, stderr } = await run('init', template, '-o', out);
    expect(code).toBe(0);
    expect(stdout).toContain('created');
    expect(stderr).toContain('Called as functions, so not in the data: euro');

    const data = JSON.parse(fs.readFileSync(out, 'utf8'));
    expect(data).toEqual({
      kunde: {
        name: 'kunde.name',
        vip: true,
        adresse: {
          strasse: 'kunde.adresse.strasse',
          plz: 'kunde.adresse.plz',
          ort: 'kunde.adresse.ort',
        },
        ansprechpartner: {
          vorname: 'kunde.ansprechpartner.vorname',
          nachname: 'kunde.ansprechpartner.nachname',
        },
      },
      angebot: { gueltigBis: 'angebot.gueltigBis' },
      produkte: [
        { name: 'produkte[].name', menge: 'produkte[].menge', preis: 'produkte[].preis' },
      ],
      summe: 'summe',
      rabatt: 'rabatt',
      themen: ['themen[]'],
      anhaenge: [{ titel: 'anhaenge[].titel' }],
    });
    // In the order the tags read them, slide by slide
    expect(Object.keys(data)).toEqual([
      'kunde',
      'angebot',
      'produkte',
      'summe',
      'rabatt',
      'themen',
      'anhaenge',
    ]);
  });

  test('the written data lints and renders the template', async () => {
    const data = JSON.parse(fs.readFileSync(out, 'utf8'));
    const lint = await lintTemplate(template, { data, additionalJsContext: { euro } });
    expect(lint.issues).toEqual([]);

    const file = 'template-logic-init.test.pptx';
    const automizer = new Automizer({
      templateDir: `${__dirname}/pptx-templates`,
      outputDir: `${__dirname}/pptx-output`,
    })
      .loadRoot('RootTemplate.pptx')
      .load('TemplateLogic.pptx');
    for (const n of [1, 2, 3]) {
      automizer.addSlide('TemplateLogic.pptx', n, (slide) =>
        slide.modify(modify.renderTemplate(data, { additionalJsContext: { euro } })),
      );
    }
    await automizer.write(file);

    const texts = async (n: number) =>
      [...(await expectXml(file, `ppt/slides/slide${n}.xml`)).raw().matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
        .map((p) => [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => t[1]).join(''))
        .filter(Boolean);

    expect(await texts(2)).toEqual([
      'Angebot für kunde.name',
      'Danke für Ihre Treue!',
      'kunde.adresse.strasse, kunde.adresse.plz kunde.adresse.ort',
      'Gültig bis angebot.gueltigBis',
    ]);
    // One table row per list entry; "rabatt" > 0 is false, so that line goes
    expect(await texts(3)).toEqual([
      'Produkt',
      'Menge',
      'Preis',
      'produkte[].name',
      'produkte[].menge',
      'produkte[].preis EUR',
      'Summe: summe EUR',
    ]);
    expect(await texts(4)).toEqual([
      'Ihr Kontakt: kunde.ansprechpartner.vorname kunde.ansprechpartner.nachname',
      'Themen: themen[]',
      '1. anhaenge[].titel ',
    ]);
  });
});
