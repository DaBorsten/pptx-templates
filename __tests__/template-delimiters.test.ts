import Automizer, {
  lintTemplate,
  modify,
  templateDataSkeleton,
  templateDataUsage,
  XmlElement,
} from '../src/index';
import { runCli } from '../src/extensions/cli';
import { expectXml } from './helpers/expect-xml';
import { deck, shape } from './helpers/lint-deck';
import fs from 'fs';
import os from 'os';
import path from 'path';

/**
 * Other tag delimiters, including different ones for start and end, through
 * every tool: renderTemplate, lint, usage, init and the CLI options.
 */
const PAIRS: [string, string][] = [
  ['[[', ']]'],
  ['<%', '%>'],
  ['${', '}'],
  ['«', '»'],
  ['+++', '+++'],
];

const data = {
  user: { name: 'Ada', vip: true },
  items: [{ name: 'Apples' }, { name: 'Pears' }],
};

/** Template paragraphs with the given delimiters; strings are runs. */
const template = (o: string, c: string): string[][] => [
  // A tag torn across runs, as PowerPoint stores it
  ['Hello ', `${o}us`, 'er.na', `me${c}!`],
  [`${o}IF user.vip${c}VIP${o}ELSE${c}normal${o}END-IF${c}`],
  [`${o}FOR i IN items${c}(${o}i.name${c})${o}END-FOR${c}`],
];

const paragraphTexts = (xml: string): string[] =>
  [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((p) =>
    [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => t[1]).join(''),
  );

const writeTemplate =
  (paragraphs: string[][]) =>
  (element: XmlElement): void =>
    modify.setMultiText(
      paragraphs.map((runs) => ({
        paragraph: {},
        textRuns: runs.map((text) => ({ text })),
      })),
    )(element);

describe.each(PAIRS)('delimiters %s … %s', (o, c) => {
  const id = `${PAIRS.findIndex(([a]) => a === o)}`;

  test('renderTemplate renders them', async () => {
    const file = `template-delimiters-${id}.test.pptx`;
    await new Automizer({
      templateDir: `${__dirname}/pptx-templates`,
      outputDir: `${__dirname}/pptx-output`,
    })
      .loadRoot('RootTemplate.pptx')
      .load('TextReplace.pptx')
      .addSlide('TextReplace.pptx', 1, (slide) => {
        slide.modifyElement('setText', writeTemplate(template(o, c)));
        slide.modify(modify.renderTemplate(data, { openingTag: o, closingTag: c }));
      })
      .write(file);
    const slide = await expectXml(file, 'ppt/slides/slide2.xml');
    const texts = paragraphTexts(slide.raw());
    expect(texts).toEqual(expect.arrayContaining(['Hello Ada!', 'VIP', '(Apples)(Pears)']));
  });

  test('lint, usage and init read them', async () => {
    const pptx = await deck([shape('Text', template(o, c))]);
    const options = { openingTag: o, closingTag: c };

    expect((await lintTemplate(pptx, { ...options, data })).issues).toEqual([]);
    // An unclosed block is still found
    const broken = await deck([shape('Text', [[`${o}IF user.vip${c}x`]])]);
    expect((await lintTemplate(broken, options)).issues.map((i) => i.rule)).toEqual([
      'unclosed-block',
    ]);

    const usage = await templateDataUsage(pptx, { ...options, data });
    expect(usage.used.map((u) => u.path)).toEqual([
      'user.name',
      'user.vip',
      'items',
      'items[].name',
    ]);
    expect(usage.missing).toEqual([]);

    expect((await templateDataSkeleton(pptx, options)).data).toEqual({
      user: { name: 'user.name', vip: true },
      items: [{ name: 'items[].name' }],
    });
  });

  test('the CLI options and the config file pass them on', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-delimiters-'));
    const file = path.join(dir, 'deck.pptx');
    fs.writeFileSync(file, await deck([shape('Text', template(o, c))]));
    const dataFile = path.join(dir, 'data.json');
    fs.writeFileSync(dataFile, JSON.stringify(data));
    const run = async (...args: string[]) => {
      let stdout = '';
      let stderr = '';
      const code = await runCli(args, {
        stdout: (t) => (stdout += t),
        stderr: (t) => (stderr += t),
      });
      return { code, stdout, stderr };
    };
    const tags = ['--opening-tag', o, '--closing-tag', c];

    let result = await run('lint', file, '--data', dataFile, ...tags);
    expect(result.stdout).toContain('No problems found');
    expect(result.code).toBe(0);

    result = await run('usage', file, '--json', ...tags);
    expect(Object.values(JSON.parse(result.stdout).files)[0]).toMatchObject({
      used: expect.arrayContaining([expect.objectContaining({ path: 'items[].name' })]),
    });

    result = await run('init', file, ...tags);
    expect(JSON.parse(result.stdout)).toEqual({
      user: { name: 'user.name', vip: true },
      items: [{ name: 'items[].name' }],
    });

    // Same delimiters from pptx-templates.json
    const config = path.join(dir, 'pptx-templates.json');
    fs.writeFileSync(
      config,
      JSON.stringify({
        openingTag: o,
        closingTag: c,
        templates: [{ files: 'deck.pptx', data: 'data.json' }],
      }),
    );
    result = await run('lint', '--config', config);
    expect(result.stdout).toContain('No problems found');
  });
});
