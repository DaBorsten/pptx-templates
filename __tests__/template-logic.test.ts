import Automizer, {
  listTemplateCommands,
  modify,
  RenderTemplateOptions,
  ShapeModificationCallback,
  TableData,
  TemplateData,
  XmlDocument,
  XmlElement,
} from '../src/index';
import { expectXml } from './helpers/expect-xml';
import { imageDimensions } from '../src/helper/image-dimensions';
import fs from 'fs';
import JSZip from 'jszip';

/**
 * Writes template text into a shape the way PowerPoint stores it: every inner
 * array is one paragraph, every string one run - so tags may be torn apart.
 */
const writeTemplate =
  (paragraphs: string[][]) =>
  (element: XmlElement): void =>
    modify.setMultiText(
      paragraphs.map((runs) => ({
        paragraph: {},
        textRuns: runs.map((text) => ({ text })),
      })),
    )(element);

/** Sets the text of each table cell; `cells[r][c]`, undefined = keep. */
const writeCells =
  (cells: (string | undefined)[][]) =>
  (element: XmlElement): void => {
    const rows = element.getElementsByTagName('a:tr');
    cells.forEach((row, r) =>
      row.forEach((text, c) => {
        if (text === undefined) return;
        const cell = rows[r].getElementsByTagName('a:tc')[c];
        const runs = cell.getElementsByTagName('a:r');
        runs[0].getElementsByTagName('a:t')[0].textContent = text;
        for (let i = runs.length - 1; i > 0; i--) {
          runs[i].parentNode.removeChild(runs[i]);
        }
      }),
    );
  };

const paragraphTexts = (xml: string): string[] =>
  [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map((p) =>
    [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => t[1]).join(''),
  );

/** The XML of one <p:sp>, found by its name. */
const shapeXml = (xml: string, name: string): string => {
  const start = xml.indexOf(`name="${name}"`);
  return xml.slice(start, xml.indexOf('</p:sp>', start));
};

const data = {
  user: { name: 'Ada', vip: true, orders: 3 },
  items: [
    { name: 'Apples', price: 1.5 },
    { name: 'Pears', price: 2 },
  ],
  empty: [] as string[],
  upper: (s: string) => s.toUpperCase(),
  note: 'line one\nline two',
};

test('render template logic in text shapes and tables', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  await automizer
    .loadRoot('RootTemplate.pptx')
    .load('TextReplace.pptx')
    .load('SlideWithTables.pptx', 'tables')
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', [
        writeTemplate([
          ['Hello {{us', 'er.na', 'me}}!'],
          ['{{IF user.vip}}'],
          ['Dear VIP, you have {{user.orders}} orders.'],
          ['{{ELSE}}'],
          ['Hi stranger.'],
          ['{{END-IF}}'],
          ['{{FOR item IN items}}'],
          ['{{$idx + 1}}. {{upper($item.name)}}: {{item.price.toFixed(2)}}'],
          ['{{END-FOR}}'],
          [
            '{{FOR x IN empty}}never{{END-FOR}}Status: {{IF user.orders > 5}}gold{{ELSE-IF user.orders > 2}}silver{{ELSE}}none{{END-IF}}',
          ],
          ["{{IF user.name == 'Ada'}}quoted ok{{END-IF}}"],
          ['{{= note}}'],
        ]),
        modify.renderTemplate(data),
      ]);
    })
    .addSlide('tables', 1, (slide) => {
      // Blocks left open in a cell span table rows
      slide.modifyElement('TableWithHeader', [
        writeCells([
          [],
          [
            '{{FOR item IN items}}{{item.name}}',
            '{{item.price}}',
            '{{IF item.price > 1.75}}high{{ELSE}}low{{END-IF}}',
            '{{END-FOR}}',
          ],
          ['{{IF user.vip}}VIP row', '{{user.name}}', '', '{{END-IF}}'],
          ['{{IF !user.vip}}hidden', '', '', '{{END-IF}}'],
        ]),
        modify.renderTemplate(data),
      ]);
      // Rows holding nothing but control tags are dropped
      slide.modifyElement('TableDefault', [
        writeCells([
          ['{{FOR item IN items}}', '', '', ''],
          ['{{item.name}}', '', '', ''],
          ['{{END-FOR}}', '', '', ''],
        ]),
        modify.renderTemplate(data),
      ]);
    })
    .write('render-template.test.pptx');

  const slide = await expectXml(
    'render-template.test.pptx',
    'ppt/slides/slide2.xml',
  );
  // The slide's other shape keeps its {{tags}}: only 'setText' was rendered
  const shape = shapeXml(slide.raw(), 'setText');
  expect(shape).not.toMatch(/\{\{|\}\}/);
  expect(paragraphTexts(shape)).toEqual([
    'Hello Ada!',
    'Dear VIP, you have 3 orders.',
    '1. APPLES: 1.50',
    '2. PEARS: 2.00',
    'Status: silver',
    'quoted ok',
    'line oneline two',
  ]);
  // Multiline values become a soft line break inside the paragraph
  slide.toContainElementTimes('a:br', 1);

  const table = await expectXml(
    'render-template.test.pptx',
    'ppt/slides/slide3.xml',
  );
  expect(table.raw()).not.toMatch(/\{\{|\}\}/);
  const frames = table.elements('p:graphicFrame');
  const frame = (name: string) =>
    frames.find(
      (f) => f.getElementsByTagName('p:cNvPr')[0].getAttribute('name') === name,
    );
  const cellTexts = (name: string) =>
    Array.from(frame(name).getElementsByTagName('a:tr')).map((row) =>
      Array.from(row.getElementsByTagName('a:tc')).map((cell) =>
        Array.from(cell.getElementsByTagName('a:t'))
          .map((t) => t.textContent)
          .join(''),
      ),
    );
  const withHeader = cellTexts('TableWithHeader');
  expect(withHeader.slice(1)).toEqual([
    ['Apples', '1.5', 'low', ''],
    ['Pears', '2', 'high', ''],
    ['VIP row', 'Ada', '', ''],
  ]);
  expect(withHeader[0][0]).toBe('Header 1');
  expect(cellTexts('TableDefault').map((row) => row[0])).toEqual([
    'Apples',
    'Pears',
  ]);
  // The graphic frame shrinks to the remaining rows
  const defaultFrame = frame('TableDefault');
  const rowHeights = Array.from(defaultFrame.getElementsByTagName('a:tr'))
    .map((row) => Number(row.getAttribute('h')))
    .reduce((a, b) => a + b, 0);
  const xfrm = defaultFrame.getElementsByTagName('p:xfrm')[0];
  expect(xfrm.getElementsByTagName('a:ext')[0].getAttribute('cy')).toBe(
    String(rowHeights),
  );
});

test('render template logic on a whole slide; skips falsy blocks', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  await automizer
    .loadRoot('RootTemplate.pptx')
    .load('TextReplace.pptx')
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'setText',
        writeTemplate([['{{IF user.vip}}'], ['only for vips'], ['{{END-IF}}']]),
      );
      slide.modifyElement(
        'replaceText',
        writeTemplate([['Name: {{user.name}}']]),
      );
      slide.modify(
        modify.renderTemplate({ user: { name: 'Bob', vip: false } }),
      );
    })
    .write('render-template-slide.test.pptx');

  const slide = await expectXml(
    'render-template-slide.test.pptx',
    'ppt/slides/slide2.xml',
  );
  expect(slide.raw()).not.toMatch(/only for vips|\{\{/);
  expect(paragraphTexts(slide.raw())).toContain('Name: Bob');
});

test('template errors reject write()', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const write = (paragraphs: string[][], strict = false) =>
    automizer
      .loadRoot('RootTemplate.pptx')
      .load('TextReplace.pptx')
      .addSlide('TextReplace.pptx', 1, (slide) => {
        slide.modifyElement('setText', [
          writeTemplate(paragraphs),
          modify.renderTemplate({}, { strict }),
        ]);
      })
      .write('render-template-error.test.pptx');

  await expect(write([['{{IF a}}open']])).rejects.toThrow(
    /Unclosed template block \{\{IF a\}\}/,
  );
  await expect(write([['{{missing}}']], true)).rejects.toThrow(
    /"missing" is not defined/,
  );
  await expect(write([['{{a.constructor}}']])).rejects.toThrow(/not allowed/);
});

/** Renders `paragraphs` into the 'setText' shape and returns its paragraphs. */
const renderShape = async (
  file: string,
  paragraphs: string[][],
  templateData: TemplateData | Promise<TemplateData>,
  options?: RenderTemplateOptions,
) => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });
  await automizer
    .loadRoot('RootTemplate.pptx')
    .load('TextReplace.pptx')
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', [
        writeTemplate(paragraphs),
        modify.renderTemplate(templateData, options),
      ]);
    })
    .write(file);
  const slide = await expectXml(file, 'ppt/slides/slide2.xml');
  return { slide, texts: paragraphTexts(shapeXml(slide.raw(), 'setText')) };
};

test('typographic quotes need fixSmartQuotes, as in docx-templates', async () => {
  const paragraphs = [
    ['{{IF user.name == ’Ada’}}yes{{END-IF}} {{“a” + ‚b‘ + „c“}}'],
    ['{{ALIAS greet INS ‘Hi ’ + user.name}}{{*greet}}'],
  ];
  await expect(
    renderShape('template-logic-smart-quotes-off.test.pptx', paragraphs, data),
  ).rejects.toThrow(/Unexpected character "’"/);

  const { texts } = await renderShape(
    'template-logic-smart-quotes.test.pptx',
    paragraphs,
    data,
    { fixSmartQuotes: true },
  );
  expect(texts).toEqual(['yes abc', 'Hi Ada']);
});

test('setTable markers like {{each:row}} stay as text', async () => {
  const { texts } = await renderShape(
    'template-logic-table-marker.test.pptx',
    [['{{each:row}}'], ['{{ each:subSub2 }} {{user.name}}']],
    data,
  );
  expect(texts).toEqual(['{{each:row}}', '{{ each:subSub2 }} Ada']);
});

describe('nested tables: setTable markers next to template tags', () => {
  const tableData: TableData = {
    body: [
      { values: ['top left', 123, 345, 'subsub3-1', 'subsub3-2', 'Last'] },
      { values: [undefined, 't1', 't2', 't3', 't3', ''] },
      { values: ['label 0', 1, 2, 3, 3, 'l0'] },
      { values: ['label 1', 123, 345, 4563, 4671, 'l1'] },
      { values: ['label 2', 123, 345, 4562, 4672] },
      { values: ['label 3', 123, 345, 4561, 4673, 'l3'] },
      { values: [undefined, 'Foo', 'ter', 4564, 4674, ''] },
    ],
  };
  const setTable = modify.setTable(tableData, {
    expand: [
      { mode: 'row', tag: '{{each:row}}', count: 3 },
      { mode: 'column', tag: '{{each:subSub3}}', count: 1 },
    ],
  });
  const render = modify.renderTemplate(data);

  /** NestedTable3 on slide 4 of NestedTables.pptx; returns its <a:tbl>. */
  const tableXml = async (
    file: string,
    callbacks: ShapeModificationCallback[],
  ) => {
    await new Automizer({
      templateDir: `${__dirname}/pptx-templates`,
      outputDir: `${__dirname}/pptx-output`,
    })
      .loadRoot('RootTemplate.pptx')
      .load('NestedTables.pptx', 'tables')
      .addSlide('tables', 4, (slide) => {
        slide.modifyElement('NestedTable3', callbacks);
      })
      .write(file);
    const xml = (await expectXml(file, 'ppt/slides/slide2.xml')).raw();
    return xml.slice(xml.indexOf('<a:tbl>'), xml.indexOf('</a:tbl>'));
  };

  /** Text of each cell, row by row. */
  const cellTexts = (tbl: string): string[][] =>
    tbl
      .split(/<a:tr[ >]/)
      .slice(1)
      .map((row) =>
        row
          .split(/<a:tc[ >]/)
          .slice(1)
          .map((cell) =>
            [...cell.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => t[1]).join(''),
          ),
      );

  test('renderTemplate fills tags in cells and leaves the markers', async () => {
    // Template rows: 0 {{each:sub}} … Top, Last; 1 {{each:subSub*}};
    // 2 {{each:row}} - - - -; 3 Footer …
    const tbl = await tableXml('template-logic-nested-table-mixed.test.pptx', [
      writeCells([
        [undefined, undefined, undefined, '{{user.name}}', '{{IF user.vip}}VIP{{ELSE}}-{{END-IF}}'],
        [],
        [undefined, '{{upper(user.name)}}', undefined, undefined, '{{user.orders}} {{each:keep}}'],
        ['{{FOR i IN items}}{{i.name}} {{END-FOR}}'],
      ]),
      render,
    ]);
    expect(cellTexts(tbl)).toEqual([
      ['A1-A2', '{{each:sub}}', '', 'Ada', 'VIP'],
      ['', '{{each:subSub}}', '{{each:subSub2}}', '{{each:subSub3}}', ''],
      ['{{each:row}}', 'ADA', '-', '-', '3 {{each:keep}}'],
      ['Apples Pears ', 'Footer', 'Footer2', 'Footer3', '-'],
    ]);
  });

  test('renderTemplate before or after setTable gives the same table', async () => {
    // Cells setTable keeps: the 6th of the shorter "label 2" row (copied
    // from the {{each:row}} row) and the first footer cell (undefined)
    const cells = writeCells([
      [],
      [],
      [undefined, undefined, undefined, undefined, '{{user.name}}'],
      ['{{IF user.vip}}VIP{{END-IF}}'],
    ]);

    const unrendered = await tableXml('template-logic-nested-table.test.pptx', [
      cells,
      setTable,
    ]);
    expect(unrendered.match(/<a:tr /g)).toHaveLength(7);
    expect(unrendered).toContain('<a:t>subsub3-2</a:t>');
    expect(unrendered).not.toContain('{{each:');

    const renderFirst = await tableXml(
      'template-logic-nested-table-render-first.test.pptx',
      [cells, render, setTable],
    );
    const renderAfter = await tableXml(
      'template-logic-nested-table-render-after.test.pptx',
      [cells, setTable, render],
    );
    // renderTemplate renumbers a16:rowId (10000 + row) in tables it renders,
    // setTable from 0: both are fine as long as they are unique per table
    const rowIds = (tbl: string) =>
      [...tbl.matchAll(/a16:rowId[^>]*val="(\d+)"/g)].map((m) => m[1]);
    [renderFirst, renderAfter].forEach((tbl) => {
      expect(rowIds(tbl)).toHaveLength(7);
      expect(new Set(rowIds(tbl)).size).toBe(7);
    });
    const withoutRowIds = (tbl: string) =>
      tbl.replace(/(a16:rowId[^>]*val=")\d+"/g, '$1"');
    expect(withoutRowIds(renderAfter)).toBe(withoutRowIds(renderFirst));

    const texts = cellTexts(renderFirst);
    expect(texts).toHaveLength(7);
    expect(texts[4]).toEqual(['label 2', '123', '345', '4562', '4672', 'Ada']);
    expect(texts[6][0]).toBe('VIP');
    expect(renderFirst).not.toContain('{{');
    // Apart from the rendered cells, the table is what setTable alone makes
    expect(
      renderFirst.replace('<a:t>Ada</a:t>', '').replace('<a:t>VIP</a:t>', ''),
    ).toBe(
      unrendered
        .replace('<a:t>{{user.name}}</a:t>', '')
        .replace('<a:t>{{IF user.vip}}VIP{{END-IF}}</a:t>', ''),
    );
  });
});

test('EXEC, ALIAS, arrow functions, object/array literals', async () => {
  const { texts } = await renderShape(
    'template-logic-exec.test.pptx',
    [
      ["{{EXEC total = 0; label = 'Sum'}}"],
      ['{{FOR item IN items}}{{! total = total + item.price}}{{END-FOR}}'],
      ['{{label}}: {{total}}'],
      ['{{ALIAS who INS user.name.toUpperCase()}}'],
      ['Hi {{*who}} / {{*who}}'],
      ['{{items.filter(i => i.price > 1.75).map(i => i.name)}}'],
      ['{{[3, 1, 2].sort((a, b) => a - b)}}'],
      ['{{({ a: 1, b: user.name }).b}} {{user?.missing?.deep ?? "fallback"}}'],
      ['{{FOR n IN [1, 2]}}{{EXEC sq = n * n}}{{sq}};{{END-FOR}}'],
    ],
    data,
  );
  expect(texts).toEqual([
    'Sum: 3.5',
    'Hi ADA / ADA',
    'Pears',
    '1, 2, 3',
    'Ada fallback',
    '1;4;',
  ]);
});

test('HTML inserts formatted paragraphs, LINK inserts hyperlinks', async () => {
  const { slide, texts } = await renderShape(
    'template-logic-html-link.test.pptx',
    [
      ['Intro'],
      ['{{HTML body}}'],
      ['See {{LINK ({ url: site, label: "our site" })}} and {{LINK site}}.'],
      ['Back to {{LINK ({ url: 1, label: "start" })}}'],
    ],
    {
      body: '<p>First <b>bold</b></p><ul><li>One</li><li>Two</li></ul>',
      site: 'https://example.com/',
    },
  );
  expect(texts).toEqual([
    'Intro',
    'First bold',
    'One',
    'Two',
    'See our site and https://example.com/.',
    'Back to start',
  ]);
  const shape = shapeXml(slide.raw(), 'setText');
  expect(shape).toMatch(/<a:rPr[^>]*\bb="1"/);
  // Both external links share one relationship
  const ids = [
    ...shape.matchAll(/<a:hlinkClick r:id="([^"]+)"( action="[^"]+")?/g),
  ];
  expect(ids).toHaveLength(3);
  expect(ids[0][1]).toBe(ids[1][1]);
  expect(ids[2][2]).toContain('hlinksldjump');

  const rels = await expectXml(
    'template-logic-html-link.test.pptx',
    'ppt/slides/_rels/slide2.xml.rels',
  );
  const rel = (id: string) =>
    rels.elements('Relationship').find((r) => r.getAttribute('Id') === id);
  expect(rel(ids[0][1]).getAttribute('Target')).toBe('https://example.com/');
  expect(rel(ids[0][1]).getAttribute('TargetMode')).toBe('External');
  expect(rel(ids[2][1]).getAttribute('Target')).toBe('../slides/slide1.xml');
});

test('async data, additionalJsContext, delimiters, line breaks', async () => {
  const { texts, slide } = await renderShape(
    'template-logic-options.test.pptx',
    [
      ['+++loadName()+++ +++fmt(price)+++'],
      ['+++FOR x IN loadList()+++[+++x+++]+++END-FOR+++'],
      ['+++text+++'],
    ],
    Promise.resolve({
      loadName: async () => 'Async Ada',
      loadList: () => Promise.resolve(['a', 'b']),
      price: 2,
      text: 'no\nbreak',
    }),
    {
      cmdDelimiter: '+++',
      processLineBreaks: false,
      additionalJsContext: { fmt: (n: number) => `${n.toFixed(2)} €` },
    },
  );
  expect(texts).toEqual(['Async Ada 2.00 €', '[a][b]', 'no\nbreak']);
  expect(shapeXml(slide.raw(), 'setText')).not.toContain('<a:br');
});

test('errorHandler, failFast: false and rejectNullish', async () => {
  const handled = await renderShape(
    'template-logic-handler.test.pptx',
    [['A {{missing.call()}} B {{IF broken(}}x{{END-IF}}']],
    {},
    { errorHandler: (_error, command) => `<${command}>` },
  );
  expect(handled.texts).toEqual(['A &lt;{{missing.call()}}&gt; B x']);

  await expect(
    renderShape(
      'template-logic-failfast.test.pptx',
      [['{{a.b()}} {{c.d()}}']],
      {},
      { failFast: false },
    ),
  ).rejects.toThrow(
    /2 error\(s\):\n- \{\{a\.b\(\)\}\}: Not a function[\s\S]*\{\{c\.d\(\)\}\}/,
  );

  await expect(
    renderShape(
      'template-logic-nullish.test.pptx',
      [['{{user.missing}}']],
      data,
      { rejectNullish: true },
    ),
  ).rejects.toThrow(/\{\{user\.missing\}\}: "user\.missing" is undefined/);
});

test('slide.modify() resolves the slide relationships for LINK', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });
  await automizer
    .loadRoot('RootTemplate.pptx')
    .load('TextReplace.pptx')
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement('setText', writeTemplate([['{{LINK url}}']]));
      slide.modify(modify.renderTemplate({ url: 'https://example.org/' }));
    })
    .write('template-logic-slide-link.test.pptx');

  const slide = await expectXml(
    'template-logic-slide-link.test.pptx',
    'ppt/slides/slide2.xml',
  );
  const relId = /<a:hlinkClick r:id="([^"]+)"/.exec(slide.raw())[1];
  const rels = await expectXml(
    'template-logic-slide-link.test.pptx',
    'ppt/slides/_rels/slide2.xml.rels',
  );
  expect(
    rels
      .elements('Relationship')
      .find((r) => r.getAttribute('Id') === relId)
      .getAttribute('Target'),
  ).toBe('https://example.org/');
});

test('listTemplateCommands lists the commands of a template', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });
  await automizer
    .loadRoot('RootTemplate.pptx')
    .load('TextReplace.pptx')
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'setText',
        writeTemplate([
          ['{{IF us', 'er.vip}}'],
          ['{{ALIAS n = user.name}}{{*n}}'],
          ['{{END-IF}}{{FOR $p IN people}}{{END-FOR}}{{each:row}}'],
        ]),
      );
    })
    .write('template-logic-list.test.pptx');

  const commands = await listTemplateCommands(
    `${__dirname}/pptx-output/template-logic-list.test.pptx`,
  );
  expect(
    commands
      .filter((c) => c.shape === 'setText')
      .map(({ slide, type, code }) => [slide, type, code]),
  ).toEqual([
    [2, 'IF', 'user.vip'],
    [2, 'ALIAS', '= user.name'],
    [2, 'INS', 'user.name'],
    [2, 'END-IF', ''],
    [2, 'FOR', 'people'],
    [2, 'END-FOR', ''],
  ]);
});

describe('IMAGE', () => {
  const png = fs.readFileSync(`${__dirname}/media/test.png`);
  const feather = fs.readFileSync(`${__dirname}/media/feather.png`);

  const build = (
    file: string,
    setTextTemplate: string[][],
    replaceTextTemplate: string[][],
    templateData: TemplateData,
    viaModifyElement = false,
  ) => {
    const automizer = new Automizer({
      templateDir: `${__dirname}/pptx-templates`,
      outputDir: `${__dirname}/pptx-output`,
    });
    return automizer
      .loadRoot('RootTemplate.pptx')
      .load('TextReplace.pptx')
      .addSlide('TextReplace.pptx', 1, (slide) => {
        slide.modifyElement('replaceText', writeTemplate(replaceTextTemplate));
        if (viaModifyElement) {
          slide.modifyElement('setText', [
            writeTemplate(setTextTemplate),
            modify.renderTemplate(templateData),
          ]);
        } else {
          slide.modifyElement('setText', writeTemplate(setTextTemplate));
          slide.modify(modify.renderTemplate(templateData));
        }
      })
      .write(file);
  };

  const pictures = (doc: XmlDocument) =>
    Array.from(doc.getElementsByTagName('p:pic')).filter((pic) =>
      /^Picture \d+$/.test(
        pic.getElementsByTagName('p:cNvPr')[0].getAttribute('name'),
      ),
    );
  const box = (pic: XmlElement) => {
    const off = pic.getElementsByTagName('a:off')[0];
    const ext = pic.getElementsByTagName('a:ext')[0];
    return {
      x: Number(off.getAttribute('x')),
      y: Number(off.getAttribute('y')),
      cx: Number(ext.getAttribute('cx')),
      cy: Number(ext.getAttribute('cy')),
    };
  };

  test('places pictures where the shape is, writes media once', async () => {
    const file = 'template-logic-image.test.pptx';
    await build(
      file,
      [['{{IMAGE ({ data: logo, extension: ".png", alt: "Logo" })}}']],
      [
        [
          'Gallery {{FOR p IN pics}}{{IMAGE ({ data: p, extension: "png", width: 2 })}}{{END-FOR}}',
        ],
      ],
      { logo: png, pics: [`data:image/png;base64,${feather.toString('base64')}`, png] },
    );

    const slide = await expectXml(file, 'ppt/slides/slide2.xml');
    const doc = slide.doc();
    const names = Array.from(doc.getElementsByTagName('p:cNvPr')).map((n) =>
      n.getAttribute('name'),
    );
    // The shape that held nothing but the tag is replaced ...
    expect(names).not.toContain('setText');
    // ... a shape with text left keeps it, its tag removed
    expect(names).toContain('replaceText');
    expect(paragraphTexts(shapeXml(slide.raw(), 'replaceText'))).toEqual([
      'Gallery ',
    ]);

    const all = pictures(doc);
    expect(all).toHaveLength(3);
    const descr = (p: XmlElement) =>
      p.getElementsByTagName('p:cNvPr')[0].getAttribute('descr');
    const logo = all.find((p) => descr(p) === 'Logo');
    const gallery = all.filter((p) => p !== logo); // feather, png
    // Fit into the frame keeping the aspect ratio of test.png
    const { width, height } = imageDimensions(png);
    const fit = box(logo);
    expect(Math.abs(fit.cx / fit.cy - width / height)).toBeLessThan(0.01);
    // width: 2 (cm) → 720000 EMU, side by side
    expect(box(gallery[0]).cx).toBe(720000);
    expect(box(gallery[1]).cx).toBe(720000);
    expect(box(gallery[1]).x).toBeGreaterThan(box(gallery[0]).x);
    const pics = [logo, ...gallery];
    // cNvPr ids stay unique
    const ids = Array.from(doc.getElementsByTagName('p:cNvPr')).map((n) =>
      n.getAttribute('id'),
    );
    expect(new Set(ids).size).toBe(ids.length);

    // logo and pics[1] are the same file: stored once
    const embeds = pics.map((p) =>
      p.getElementsByTagName('a:blip')[0].getAttribute('r:embed'),
    );
    const rels = await expectXml(file, 'ppt/slides/_rels/slide2.xml.rels');
    const target = (id: string) =>
      rels
        .elements('Relationship')
        .find((r) => r.getAttribute('Id') === id)
        .getAttribute('Target');
    expect(target(embeds[0])).toBe(target(embeds[2]));
    expect(target(embeds[0])).not.toBe(target(embeds[1]));

    const zip = await JSZip.loadAsync(
      fs.readFileSync(`${__dirname}/pptx-output/${file}`),
    );
    const media = await zip.file(`ppt/${target(embeds[0]).slice(3)}`).async('nodebuffer');
    expect(media.equals(png)).toBe(true);
    const types = await expectXml(file, '[Content_Types].xml');
    expect(
      types
        .elements('Default')
        .some((d) => d.getAttribute('Extension') === 'png'),
    ).toBe(true);
  });

  test('IMAGE errors', async () => {
    await expect(
      build(
        'template-logic-image-error.test.pptx',
        [['{{IMAGE ({ data: logo, extension: ".png" })}}']],
        [['x']],
        { logo: png },
        true,
      ),
    ).rejects.toThrow(/IMAGE needs slide\.modify/);

    await expect(
      build(
        'template-logic-image-error.test.pptx',
        [['{{IMAGE ({ data: logo, extension: ".tiff" })}}']],
        [['x']],
        { logo: png },
      ),
    ).rejects.toThrow(/unsupported extension "tiff"/);
  });
});
