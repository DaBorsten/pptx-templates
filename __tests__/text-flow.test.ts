import Automizer, {
  lintTemplate,
  modify,
  RenderTemplateOptions,
  templateDataSkeleton,
} from '../src/index';
import { FontRegistry } from '../src/extensions/text-flow/fonts';
import { flowArguments } from '../src/extensions/text-flow';
import { expectXml } from './helpers/expect-xml';

// removeExistingSlides keeps the template's slide1.xml: the new slide is #2
const SLIDE = 'ppt/slides/slide2.xml';
const fonts = [`${__dirname}/fonts`];
const lorem =
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. ';
const lengths = [
  60, 140, 30, 200, 90, 45, 170, 110, 20, 130, 75, 190, 55, 100, 160, 40, 120,
  80, 150, 35, 180, 65, 95, 140, 50, 110, 85, 170,
];
const entries = (count: number) =>
  lengths.slice(0, count).map((n, i) => ({
    title: `${i + 1}. Entry`,
    text: lorem.repeat(2).slice(0, n).trim(),
  }));

type Box = { name: string; x: number; w?: number; before?: string[] };

/** The loop as PowerPoint stores it: one paragraph per line. */
const loop = (list = 'items') => [
  `{{FOR item IN flow(${list})}}`,
  '{{item.title}}',
  '{{item.text}}',
  '{{END-FOR}}',
];

/**
 * One slide with text boxes of 11 pt Liberation Sans, 6.2" high, that all
 * contain the same flow() loop, and optional extra text boxes.
 */
const render = async (
  file: string,
  boxes: Box[],
  data: Record<string, unknown>,
  options: RenderTemplateOptions = {},
  /** Other text boxes; `after`: later in the z-order than the flow boxes. */
  extra: { name: string; y: number; text: string[]; after?: boolean }[] = [],
  body = loop,
) => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
    verbosity: 0,
  });
  return automizer
    .loadRoot('EmptySlide.pptx')
    .load('EmptySlide.pptx', 'empty')
    .addSlide('empty', 1, (slide) => {
      slide.generate((s) => {
        const text = (lines: string[]) =>
          lines.map((t) => ({
            text: t,
            options: { breakLine: true, bold: t === '{{item.title}}' },
          }));
        const addExtra = (after: boolean) =>
          extra
            .filter((e) => !!e.after === after)
            .forEach((e) =>
              s.addText(text(e.text), {
                x: 0.5,
                y: e.y,
                w: 12,
                h: 0.3,
                fontSize: 8,
                fontFace: 'Liberation Sans',
                objectName: e.name,
              }),
            );
        addExtra(false);
        boxes.forEach((b) =>
          s.addText(text([...(b.before ?? []), ...body()]), {
            x: b.x,
            y: 1,
            w: b.w ?? 3.95,
            h: 6.2,
            fontSize: 11,
            fontFace: 'Liberation Sans',
            valign: 'top',
            paraSpaceAfter: 4,
            objectName: b.name,
          }),
        );
        addExtra(true);
      });
      slide.modify(modify.renderTemplate(data, { fonts, ...options }));
    })
    .write(file);
};

/**
 * Entry numbers per flow box, boxes from left to right. generate() names
 * shapes itself, so the boxes are told apart by position (all at y = 1").
 */
const columnsOf = async (file: string): Promise<string[][]> => {
  const xml = (await expectXml(file, SLIDE)).raw();
  return [...xml.matchAll(/<p:sp>[\s\S]*?<\/p:sp>/g)]
    .map((m) => m[0])
    .filter((sp) => sp.includes('y="914400"'))
    .map((sp) => ({
      x: Number(/<a:off x="(\d+)"/.exec(sp)[1]),
      entries: [...sp.matchAll(/<a:t>(\d+)\. Entry<\/a:t>/g)].map((t) => t[1]),
    }))
    .sort((a, b) => a.x - b.x)
    .map((c) => c.entries);
};

/** The text of the slide, runs joined. */
const slideText = async (file: string): Promise<string> =>
  (await expectXml(file, SLIDE)).raw().replace(/<[^>]+>/g, '');

const range = (from: number, to: number) =>
  Array.from({ length: to - from + 1 }, (_, i) => String(from + i));

// Created right to left: flow() follows the position, not the z-order
const columns: Box[] = [
  { name: 'Right', x: 8.86 },
  { name: 'Middle', x: 4.68 },
  { name: 'Left', x: 0.5 },
];

test('flow() fills text boxes column by column, never splitting an entry', async () => {
  await render('text-flow-columns.pptx', columns, { items: entries(24) });

  // Verified against LibreOffice: each column is full, the next entry
  // would not fit any more (see docs/text-flow.md)
  expect(await columnsOf('text-flow-columns.pptx')).toEqual([
    range(1, 8),
    range(9, 16),
    range(17, 24),
  ]);
});

test('flow() leaves later boxes empty when the entries fit before', async () => {
  await render('text-flow-few.pptx', columns, { items: entries(5) });

  expect(await columnsOf('text-flow-few.pptx')).toEqual([range(1, 5), [], []]);
});

test('flow() measures the text around the loop and the width of each box', async () => {
  await render(
    'text-flow-layout.pptx',
    [
      // a narrow first column with a heading above the loop
      { name: 'Left', x: 0.5, w: 2.5, before: ['Heading', 'Second line'] },
      { name: 'Middle', x: 3.2, w: 4.8 },
      { name: 'Right', x: 8.2, w: 4.6 },
    ],
    { items: entries(24) },
  );

  const [left, middle, right] = await columnsOf('text-flow-layout.pptx');
  expect(left.length).toBeLessThan(8);
  expect(middle.length).toBeGreaterThan(8);
  // in order, nothing lost, nothing twice
  expect([...left, ...middle, ...right]).toEqual(range(1, 24));
});

test('flowRest() returns the entries that do not fit; flowOverflow: error throws', async () => {
  const warn = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  await render('text-flow-rest.pptx', columns, { items: entries(28) }, {}, [
    {
      name: 'Rest',
      y: 0.2,
      text: ['Rest: {{flowRest(items).map(i => i.title).join(", ")}}'],
    },
  ]);
  warn.mockRestore();

  expect((await columnsOf('text-flow-rest.pptx'))[2]).toEqual(range(17, 24));
  expect(await slideText('text-flow-rest.pptx')).toContain(
    'Rest: 25. Entry, 26. Entry, 27. Entry, 28. Entry',
  );

  await expect(
    render(
      'text-flow-error.pptx',
      columns,
      { items: entries(28) },
      {
        flowOverflow: 'error',
      },
    ),
  ).rejects.toThrow(/4 of 28 entries do not fit/);
});

test('flow() measures without repeating EXEC side effects', async () => {
  await render(
    'text-flow-exec.pptx',
    columns,
    { items: entries(24) },
    {},
    [
      { name: 'Init', y: 0.1, text: ['{{EXEC count = 0}}'] },
      { name: 'Total', y: 7.3, text: ['Total: {{count}}'], after: true },
    ],
    () => [
      '{{FOR item IN flow(items)}}',
      '{{item.title}}{{! count = count + 1}}',
      '{{item.text}}',
      '{{END-FOR}}',
    ],
  );
  expect(await slideText('text-flow-exec.pptx')).toContain('Total: 24');
});

test('flow() needs the loop around whole paragraphs', async () => {
  await expect(
    render(
      'text-flow-inline.pptx',
      columns,
      { items: entries(3) },
      {},
      [],
      () => ['{{FOR item IN flow(items)}}{{item.title}}, {{END-FOR}}'],
    ),
  ).rejects.toThrow(/paragraphs of their own/);
});

test('flow() with an empty list leaves the boxes empty', async () => {
  await render('text-flow-empty.pptx', columns, { items: [] });
  expect(await columnsOf('text-flow-empty.pptx')).toEqual([[], [], []]);
});

test('fonts: metric aliases and a warning for missing fonts', () => {
  const registry = new FontRegistry(fonts, false);
  // Arial is measured with its metric twin Liberation Sans
  expect(registry.face('Arial').info.families).toContain('liberation sans');
  expect(registry.face('Arial', true).info.bold).toBe(true);

  const warn = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  expect(registry.face('No Such Font').info.families).toContain(
    'liberation sans',
  );
  expect(warn).toHaveBeenCalledWith(
    expect.anything(),
    expect.stringContaining('"No Such Font" not found'),
  );
  warn.mockRestore();
});

test('flowArguments finds the list of each flow() call', () => {
  expect(flowArguments('FOR item IN flow(items)')).toEqual([
    { source: 'items', key: 'items' },
  ]);
  expect(flowArguments('FOR x IN flow( data.list.filter(i => i.on) )')).toEqual(
    [
      {
        source: 'data.list.filter(i => i.on)',
        key: 'data.list.filter(i=>i.on)',
      },
    ],
  );
  expect(flowArguments('FOR x IN overflow(items)')).toEqual([]);
  expect(flowArguments('flowRest(items).length', 'flowRest')).toEqual([
    { source: 'items', key: 'items' },
  ]);
});

test('the linter knows flow() and traces its list', async () => {
  // the template itself: the same boxes without renderTemplate
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
    verbosity: 0,
  });
  await automizer
    .loadRoot('EmptySlide.pptx')
    .load('EmptySlide.pptx', 'empty')
    .addSlide('empty', 1, (slide) => {
      slide.generate((s) => {
        s.addText(
          loop().map((t) => ({ text: t, options: { breakLine: true } })),
          {
            x: 1,
            y: 1,
            w: 4,
            h: 4,
            objectName: 'Box',
          },
        );
      });
    })
    .write('text-flow-template.pptx');
  const template = `${__dirname}/pptx-output/text-flow-template.pptx`;

  const result = await lintTemplate(template, {
    data: { items: [{ title: 'A' }] },
  });
  const messages = result.issues.map((i) => i.message);
  expect(messages.some((m) => m.includes('"flow"'))).toBe(false);
  expect(messages).toContain('"text" does not exist on "item" (object)');

  const skeleton = await templateDataSkeleton(template);
  expect(skeleton.data).toEqual({
    items: [{ title: 'items[].title', text: 'items[].text' }],
  });
  expect(skeleton.functions).toEqual([]);
});
