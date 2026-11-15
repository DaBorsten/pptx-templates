import fs from 'fs';
import os from 'os';
import path from 'path';
import JSZip from 'jszip';
import { renderDeck, TemplateRenderError } from '../src/index';
import { runCli } from '../src/extensions/cli';
import { checkPptxInvariants } from './helpers/pptx-invariants';

const template = `${__dirname}/pptx-templates/TemplateLogic.pptx`;
const outputDir = `${__dirname}/pptx-output`;

const data = {
  kunde: {
    name: 'Ada',
    vip: true,
    adresse: { strasse: 'Hauptstr. 1', plz: '53111', ort: 'Bonn' },
    ansprechpartner: { vorname: 'Grace', nachname: 'Hopper' },
  },
  angebot: { gueltigBis: '31.12.2026' },
  produkte: [
    { name: 'Äpfel', menge: 3, preis: 1.5 },
    { name: 'Birnen', menge: 2, preis: 2 },
  ],
  summe: 8.5,
  rabatt: 0,
  themen: ['Obst', 'Gemüse'],
  anhaenge: [{ titel: 'AGB' }],
};
const euro = (n: number) => `${n.toFixed(2)} €`;

/** The text of each slide in presentation order, runs joined. */
const slideTexts = async (pptx: Buffer): Promise<string[]> => {
  const zip = await JSZip.loadAsync(pptx);
  const rels = await zip.file('ppt/_rels/presentation.xml.rels').async('text');
  const presentation = await zip.file('ppt/presentation.xml').async('text');
  const targets = new Map(
    [...rels.matchAll(/Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [
      m[1],
      m[2],
    ]),
  );
  const ids = [...presentation.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g)].map(
    (m) => m[1],
  );
  return Promise.all(
    ids.map(async (id) => {
      const xml = await zip.file(`ppt/${targets.get(id)}`).async('text');
      return [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
    }),
  );
};

/** Writes the result like a test output and checks the package invariants. */
const keep = async (name: string, pptx: Buffer) => {
  fs.writeFileSync(`${outputDir}/${name}`, pptx);
  const { errors } = await checkPptxInvariants(pptx);
  expect(errors).toEqual([]);
};

test('renderDeck renders the tags of every slide', async () => {
  const pptx = await renderDeck(template, data, {
    additionalJsContext: { euro },
  });
  await keep('template-render.test.pptx', pptx);

  const texts = await slideTexts(pptx);
  expect(texts).toHaveLength(3);
  expect(texts[0]).toContain('Angebot für Ada');
  expect(texts[0]).toContain('Danke für Ihre Treue!');
  expect(texts[0]).toContain('Hauptstr. 1, 53111 Bonn');
  expect(texts[1]).toContain('Äpfel');
  expect(texts[1]).toContain('1.50 €');
  expect(texts[1]).toContain('Summe: 8.50 €');
  expect(texts[1]).not.toContain('Rabatt');
  expect(texts[2]).toContain('Ihr Kontakt: Grace Hopper');
  expect(texts[2]).toContain('Themen: Obst, Gemüse');
  texts.forEach((text) => expect(text).not.toContain('{{'));

  // The template's own slides are gone, not only unlisted
  const zip = await JSZip.loadAsync(pptx);
  const slideParts = Object.keys(zip.files).filter((f) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(f),
  );
  expect(slideParts).toHaveLength(3);
});

test('renderDeck keeps only the given slides, in their order', async () => {
  const pptx = await renderDeck(
    fs.readFileSync(template),
    { ...data, rabatt: 10 },
    { additionalJsContext: { euro }, slides: [2, 1] },
  );
  await keep('template-render-slides.test.pptx', pptx);

  const texts = await slideTexts(pptx);
  expect(texts).toHaveLength(2);
  expect(texts[0]).toContain('Rabatt: 10 %');
  expect(texts[1]).toContain('Angebot für Ada');
});

test('renderDeck names the slide of a failing tag', async () => {
  // euro is missing
  const error = await renderDeck(template, data).catch((e) => e);
  expect(error).toBeInstanceOf(TemplateRenderError);
  expect(error.message).toMatch(/^slide 2: \{\{euro\(p\.preis\)\}\}: /);
});

test('renderDeck rejects slide numbers the deck does not have', async () => {
  await expect(renderDeck(template, data, { slides: [4] })).rejects.toThrow(
    'The deck has no slide 4 (slides: 1, 2, 3)',
  );
});

describe('pptx-templates render', () => {
  let dir: string;
  const run = async (...args: string[]) => {
    let stdout = '';
    let stderr = '';
    const code = await runCli(['render', ...args], {
      stdout: (t) => (stdout += t),
      stderr: (t) => (stderr += t),
    });
    return { code, stdout, stderr };
  };

  beforeEach(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-render-'));
    fs.writeFileSync(path.join(dir, 'data.json'), JSON.stringify(data));
    fs.writeFileSync(
      path.join(dir, 'helpers.js'),
      'module.exports = { euro: (n) => n.toFixed(2) + " €" };',
    );
  });
  afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));

  test('writes the rendered deck', async () => {
    const out = path.join(dir, 'out', 'Angebot.pptx');
    const result = await run(
      template,
      '--data',
      path.join(dir, 'data.json'),
      '--helpers',
      path.join(dir, 'helpers.js'),
      '-o',
      out,
    );
    expect(result).toMatchObject({ code: 0, stderr: '' });
    expect(result.stdout).toContain('rendered');
    const texts = await slideTexts(fs.readFileSync(out));
    expect(texts[1]).toContain('Summe: 8.50 €');
  });

  test('exit code 1 for a tag that fails', async () => {
    const out = path.join(dir, 'out.pptx');
    const result = await run(
      template,
      '--data',
      path.join(dir, 'data.json'),
      '-o',
      out,
    );
    expect(result.code).toBe(1);
    expect(result.stderr).toContain('slide 2: {{euro(p.preis)}}');
    expect(fs.existsSync(out)).toBe(false);
  });

  test('exit code 2 for an invalid call', async () => {
    expect((await run(template)).code).toBe(2);
    expect((await run(template)).stderr).toContain('Missing --data, --out');
    expect((await run(template, '--slides', '1,x')).code).toBe(2);
    const badData = path.join(dir, 'bad.json');
    fs.writeFileSync(badData, '[1]');
    const result = await run(template, '--data', badData, '-o', 'x.pptx');
    expect(result.code).toBe(2);
    expect(result.stderr).toContain('expected a JSON object');
  });
});
