import JSZip from 'jszip';
import Automizer, { countFlowPages, renderDeck } from '../src/index';

const fonts = [`${__dirname}/fonts`];
const SLIDE = 'ppt/slides/slide2.xml';
const lorem =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. ';

/** A column x (inches) of the template: one card, three text boxes. */
const COLUMNS = [8.86, 4.68, 0.5];

/**
 * A template slide with three cards side by side, like a CV page with three
 * projects: customer, content (with a heading) and period, each its own
 * text box of a different height, plus a grey background. Every column is
 * a PowerPoint group. The cards are created right to left.
 */
const writeCardTemplate = async (
  file: string,
  loop = '{{FOR p IN flow(projects, 1)}}',
  heading = 'Content:',
): Promise<Buffer> => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
    verbosity: 0,
  });
  const lines = (texts: string[]) =>
    texts.map((text) => ({ text, options: { breakLine: true } }));
  const style = {
    w: 3.95,
    fontSize: 11,
    fontFace: 'Liberation Sans',
    valign: 'top' as const,
  };
  await automizer
    .loadRoot('EmptySlide.pptx')
    .load('EmptySlide.pptx', 'empty')
    .addSlide('empty', 1, (slide) => {
      slide.generate((s) => {
        s.addText('Page {{flowPage}} of {{flowPages}}', {
          ...style,
          x: 0.5,
          y: 0.2,
          w: 12,
          h: 0.4,
        });
        COLUMNS.forEach((x) => {
          s.addShape('rect', {
            x,
            y: 1,
            w: 3.95,
            h: 6,
            fill: { color: 'E6E6E6' },
            objectName: 'Background',
          });
          s.addText(lines([loop, 'Customer: {{p.customer}}', '{{END-FOR}}']), {
            ...style,
            x,
            y: 1,
            h: 0.5,
          });
          s.addText(lines([heading, loop, '{{p.content}}', '{{END-FOR}}']), {
            ...style,
            x,
            y: 1.6,
            h: 2,
          });
          s.addText(lines([loop, 'Period: {{p.period}}', '{{END-FOR}}']), {
            ...style,
            x,
            y: 6.4,
            h: 0.5,
          });
        });
      });
    })
    .write(file);

  // Group the shapes of each column, as PowerPoint stores Ctrl+G
  const zip = await JSZip.loadAsync(
    await import('fs').then((fs) =>
      fs.promises.readFile(`${__dirname}/pptx-output/${file}`),
    ),
  );
  let xml = await zip.file(SLIDE).async('string');
  const shapes = [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)].map((m) => m[0]);
  COLUMNS.forEach((column, i) => {
    const x = Math.round(column * 914400);
    const members = shapes.filter((sp) => sp.includes(`<a:off x="${x}"`));
    members.forEach((sp) => (xml = xml.replace(sp, '')));
    const xfrm =
      `<a:off x="${x}" y="914400"/><a:ext cx="3611880" cy="5486400"/>` +
      `<a:chOff x="${x}" y="914400"/><a:chExt cx="3611880" cy="5486400"/>`;
    const group =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${900 + i}" name="Card ${i + 1}"/>` +
      '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      `<p:grpSpPr><a:xfrm>${xfrm}</a:xfrm></p:grpSpPr>` +
      members.join('') +
      '</p:grpSp>';
    xml = xml.replace('</p:spTree>', `${group}</p:spTree>`);
  });
  zip.file(SLIDE, xml);
  return zip.generateAsync({ type: 'nodebuffer' });
};

const project = (n: number, content = lorem) => ({
  customer: `Customer ${n}`,
  content: content.trim(),
  period: `${n}/2020`,
});

/** The text of an XML fragment: the content of all a:t, joined by `sep`. */
const textOf = (xml: string, sep = ''): string =>
  [...xml.matchAll(/<a:t(?:\s[^>]*)?>([^<]*)<\/a:t>/g)]
    .map((m) => m[1])
    .join(sep);

/** Per slide: per card, left to right, its text (runs joined). */
const deckCards = async (pptx: Buffer) => {
  const zip = await JSZip.loadAsync(pptx);
  const rels = await zip
    .file('ppt/_rels/presentation.xml.rels')
    .async('string');
  const presentation = await zip.file('ppt/presentation.xml').async('string');
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
      const xml = await zip.file(`ppt/${targets.get(id)}`).async('string');
      const cards = [...xml.matchAll(/<p:grpSp>[\s\S]*?<\/p:grpSp>/g)]
        .map((m) => m[0])
        .map((group) => ({
          x: Number(/<a:off x="(\d+)"/.exec(group)[1]),
          text: textOf(group, ' ').replace(/\s+/g, ' ').trim(),
          backgrounds: (group.match(/val="E6E6E6"/g) ?? []).length,
        }))
        .sort((a, b) => a.x - b.x);
      return { text: textOf(xml), cards };
    }),
  );
};

test('flow() fills grouped text boxes as one card per entry', async () => {
  const template = await writeCardTemplate('text-flow-cards-template.pptx');
  const projects = [1, 2, 3, 4, 5, 6, 7].map((n) => project(n));

  expect(await countFlowPages(template, 2, { projects }, { fonts })).toBe(3);

  const pptx = await renderDeck(template, { projects }, { fonts });
  const slides = await deckCards(pptx);
  expect(slides).toHaveLength(3);
  expect(slides[0].text).toContain('Page 1 of 3');

  // left to right, though created right to left; every box of a card
  // shows the same project
  const customers = slides.map((s) =>
    s.cards.map((c) => /Customer: (Customer \d)/.exec(c.text)?.[1]),
  );
  expect(customers).toEqual([
    ['Customer 1', 'Customer 2', 'Customer 3'],
    ['Customer 4', 'Customer 5', 'Customer 6'],
    ['Customer 7'],
  ]);
  slides
    .flatMap((s) => s.cards)
    .forEach((card) => {
      const n = /Customer (\d)/.exec(card.text)[1];
      expect(card.text).toContain(`Period: ${n}/2020`);
      expect(card.text).toContain('Content: Lorem ipsum');
      expect(card.backgrounds).toBe(1);
    });
  // the empty cards of the last slide are removed, with their background
  expect(slides[2].cards).toHaveLength(1);
  expect(slides[2].text.match(/Content:/g)).toHaveLength(1);
});

// Like a single box: the entries after it are not placed either
test('an entry that is too large for a box of the card fits no card', async () => {
  const template = await writeCardTemplate(
    'text-flow-cards-template-overflow.pptx',
  );
  const projects = [project(1), project(2, lorem.repeat(12)), project(3)];
  await expect(
    renderDeck(template, { projects }, { fonts, flowOverflow: 'error' }),
  ).rejects.toThrow(/2 of 3 entries fit into none/);
});

test('{ measure: false } counts the entries only: a long text overflows its box', async () => {
  const template = await writeCardTemplate(
    'text-flow-cards-template-unmeasured.pptx',
    '{{FOR p IN flow(projects, 1, { measure: false })}}',
  );
  const projects = [1, 2, 3, 4].map((n) =>
    project(n, n === 2 ? lorem.repeat(12) : lorem),
  );
  const slides = await deckCards(
    await renderDeck(template, { projects }, { flowOverflow: 'error' }),
  );
  expect(
    slides.map((s) =>
      s.cards.map((c) => /Customer: (Customer \d)/.exec(c.text)?.[1]),
    ),
  ).toEqual([['Customer 1', 'Customer 2', 'Customer 3'], ['Customer 4']]);
  expect(slides[0].text).toContain('Page 1 of 2');
});

/** The cards of every slide, each slide rendered by itself (no cache). */
const pagesOneByOne = async (
  template: Buffer,
  data: { projects: unknown[] },
  pages: number,
) => {
  const slides = [];
  for (let flowPage = 0; flowPage < pages; flowPage++) {
    const deck = await renderDeck(template, data, { fonts, flowPage });
    slides.push((await deckCards(deck))[0].cards.map((c) => c.text));
  }
  return slides;
};

describe('renderDeck measures a list once for all copies of the slide', () => {
  // entries of one, two and three times the text: a card takes one or two
  const sizes = [1, 3, 1, 1, 2, 2, 1, 3, 3, 1, 1];
  // the slides rendered one by one measure the list once per slide
  const TIMEOUT = 60000;

  test(
    'entries of different sizes: same slides as rendered one by one',
    async () => {
      const template = await writeCardTemplate(
        'text-flow-cards-template-sizes.pptx',
        '{{FOR p IN flow(projects)}}',
      );
      let measured = 0;
      const projects = sizes.map((size, i) => ({
        ...project(i + 1),
        get content() {
          measured++;
          return lorem.repeat(size).trim();
        },
      }));

      const slides = await deckCards(
        await renderDeck(template, { projects }, { fonts }),
      );
      const calls = measured;
      const cards = slides.map((s) => s.cards.map((c) => c.text));
      // not every slide takes the same number of entries
      const perSlide = cards.map(
        (slide) => slide.join(' ').match(/Customer: /g).length,
      );
      expect(new Set(perSlide).size).toBeGreaterThan(1);
      expect(perSlide.reduce((a, b) => a + b)).toBe(sizes.length);

      expect(cards).toEqual(
        await pagesOneByOne(template, { projects }, slides.length),
      );
      // each entry once per content box (three cards) to measure it, and
      // once on its slide: not once more per copy of the slide
      expect(calls).toBeLessThanOrEqual(4 * sizes.length);
    },
    TIMEOUT,
  );

  test(
    'text around the loop that differs per copy is measured per copy',
    async () => {
      const template = await writeCardTemplate(
        'text-flow-cards-template-heading.pptx',
        '{{FOR p IN flow(projects)}}',
        `Content:{{IF flowPage > 1}} ${lorem.repeat(2)}{{END-IF}}`,
      );
      const projects = sizes.map((size, i) =>
        project(i + 1, lorem.repeat(Math.min(size, 2))),
      );
      const slides = await deckCards(
        await renderDeck(template, { projects }, { fonts }),
      );
      expect(slides.length).toBeGreaterThan(2);
      expect(slides.map((s) => s.cards.map((c) => c.text))).toEqual(
        await pagesOneByOne(template, { projects }, slides.length),
      );
    },
    TIMEOUT,
  );
});
