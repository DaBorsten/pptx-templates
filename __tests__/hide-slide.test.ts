import fs from 'fs';
import JSZip from 'jszip';
import Automizer, {
  isSlideHidden,
  modify,
  lintTemplate,
  listTemplateCommands,
  renderDeck,
  TemplateRenderError,
} from '../src/index';
import { checkPptxInvariants } from './helpers/pptx-invariants';

// Three slides ("Slide 1" … "Slide 3"), each with an empty notes body
const template = `${__dirname}/pptx-templates/threeSlidePres.pptx`;
const outputDir = `${__dirname}/pptx-output`;

const xmlText = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** The template with `notes` written into the notes of slide 1, 2, 3. */
const withNotes = async (
  notes: Record<number, string | string[]>,
  slides: Record<number, string> = {},
): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(fs.readFileSync(template));
  for (const [n, text] of Object.entries(notes)) {
    const part = `ppt/notesSlides/notesSlide${n}.xml`;
    const xml = await zip.file(part).async('string');
    const lines = Array.isArray(text) ? text : [text];
    const paragraphs = lines
      .map((line) => `<a:p><a:r><a:t>${xmlText(line)}</a:t></a:r></a:p>`)
      .join('');
    zip.file(
      part,
      xml.replace(
        /<a:p>(?:(?!<a:p>)[\s\S])*?<a:t><\/a:t>[\s\S]*?<\/a:p>/,
        paragraphs,
      ),
    );
  }
  for (const [n, paragraphs] of Object.entries(slides)) {
    const part = `ppt/slides/slide${n}.xml`;
    const xml = await zip.file(part).async('string');
    // The first text body: its paragraphs replaced
    zip.file(
      part,
      xml.replace(
        /(<p:txBody>[\s\S]*?<a:lstStyle\/>)[\s\S]*?(<\/p:txBody>)/,
        `$1${paragraphs}$2`,
      ),
    );
  }
  return zip.generateAsync({ type: 'nodebuffer' });
};

/** Slide and notes text of each output slide, in presentation order. */
const deck = async (pptx: Buffer) => {
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
  const textOf = (xml: string) =>
    [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
  return Promise.all(
    ids.map(async (id) => {
      const path = `ppt/${targets.get(id)}`;
      const slide = await zip.file(path).async('text');
      const slideRels = await zip
        .file(path.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels'))
        .async('text');
      const notesTarget =
        /Target="\.\.\/notesSlides\/(notesSlide\d+\.xml)"/.exec(slideRels)?.[1];
      const notes = notesTarget
        ? await zip.file(`ppt/notesSlides/${notesTarget}`).async('text')
        : '';
      return { text: textOf(slide), notes: textOf(notes) };
    }),
  );
};

const keep = async (name: string, pptx: Buffer) => {
  fs.writeFileSync(`${outputDir}/${name}`, pptx);
  const { errors } = await checkPptxInvariants(pptx);
  expect(errors).toEqual([]);
};

test('renderDeck leaves out a slide whose HIDE-SLIDE condition is true', async () => {
  const pptx = await withNotes({
    2: '{{HIDE-SLIDE projects.length <= 2}}',
    3: '{{HIDE-SLIDE !projects.length}}',
  });

  const few = await renderDeck(pptx, { projects: [1] });
  await keep('hide-slide.test.pptx', few);
  expect((await deck(few)).map((s) => s.text)).toEqual(['Slide 1', 'Slide 3']);

  const many = await renderDeck(pptx, { projects: [1, 2, 3] });
  expect((await deck(many)).map((s) => s.text)).toEqual([
    'Slide 1',
    'Slide 2',
    'Slide 3',
  ]);

  const none = await renderDeck(pptx, { projects: [] });
  expect((await deck(none)).map((s) => s.text)).toEqual(['Slide 1']);
});

test('HIDE-SLIDE without a condition always hides; several tags hide when one is true', async () => {
  const pptx = await withNotes({
    1: '{{HIDE-SLIDE}}',
    2: ['{{HIDE-SLIDE a}}', '{{hide-slide b}}'],
  });
  const texts = async (data: Record<string, unknown>) =>
    (await deck(await renderDeck(pptx, data))).map((s) => s.text);

  expect(await texts({ a: false, b: false })).toEqual(['Slide 2', 'Slide 3']);
  expect(await texts({ a: false, b: true })).toEqual(['Slide 3']);
  expect(await texts({ a: true, b: false })).toEqual(['Slide 3']);
});

test('the HIDE-SLIDE tag is removed from the notes of the output', async () => {
  const pptx = await withNotes({
    2: ['Talk about the projects.', '{{HIDE-SLIDE hide}}'],
    3: 'Before {{HIDE-SLIDE hide}}after',
  });
  const slides = await deck(await renderDeck(pptx, { hide: false }));
  expect(slides[1].notes).toBe('Talk about the projects.');
  expect(slides[2].notes).toBe('Before after');
  slides.forEach((s) => expect(s.notes).not.toContain('HIDE'));
});

test('a tag split across runs in the notes is found and removed', async () => {
  const zip = await JSZip.loadAsync(fs.readFileSync(template));
  const part = 'ppt/notesSlides/notesSlide2.xml';
  const xml = await zip.file(part).async('string');
  zip.file(
    part,
    xml.replace(
      '<a:r><a:t></a:t></a:r>',
      '<a:r><a:t>{{HIDE-</a:t></a:r><a:r><a:rPr lang="de-DE"/><a:t>SLIDE hide}}</a:t></a:r>',
    ),
  );
  const pptx = await zip.generateAsync({ type: 'nodebuffer' });

  expect((await deck(await renderDeck(pptx, { hide: true }))).length).toBe(2);
  const kept = await deck(await renderDeck(pptx, { hide: false }));
  expect(kept[1].notes).not.toContain('SLIDE');
});

test('a tag split inside the keyword is found and removed', async () => {
  // spell check splits runs anywhere, also within HIDE
  const zip = await JSZip.loadAsync(fs.readFileSync(template));
  const part = 'ppt/notesSlides/notesSlide2.xml';
  const xml = await zip.file(part).async('string');
  zip.file(
    part,
    xml.replace(
      '<a:r><a:t></a:t></a:r>',
      '<a:r><a:t>Notes {{HI</a:t></a:r><a:r><a:rPr lang="de-DE"/><a:t>DE-SLIDE off}}</a:t></a:r>',
    ),
  );
  const pptx = await zip.generateAsync({ type: 'nodebuffer' });
  // the keyword is nowhere else in the part
  expect(/hide/i.test(xml)).toBe(false);

  expect((await deck(await renderDeck(pptx, { off: true }))).length).toBe(2);
  const kept = await deck(await renderDeck(pptx, { off: false }));
  expect(kept[1].notes).toBe('Notes ');
});

test('the condition sees additionalJsContext, other delimiters and fixSmartQuotes', async () => {
  const pptx = await withNotes({ 2: '<<HIDE-SLIDE isEmpty(list)>>' });
  const slides = await deck(
    await renderDeck(
      pptx,
      { list: [] },
      {
        openingTag: '<<',
        closingTag: '>>',
        additionalJsContext: { isEmpty: (l: unknown[]) => !l.length },
      },
    ),
  );
  expect(slides.map((s) => s.text)).toEqual(['Slide 1', 'Slide 3']);

  const quoted = await withNotes({ 2: '{{HIDE-SLIDE lang == ’de’}}' });
  const de = await deck(
    await renderDeck(quoted, { lang: 'de' }, { fixSmartQuotes: true }),
  );
  expect(de).toHaveLength(2);
});

test('HIDE-SLIDE can hide single copies of a flow() slide', async () => {
  const loop =
    '<a:p><a:r><a:t>{{FOR i IN flow(items, 1, { measure: false })}}</a:t></a:r></a:p>' +
    '<a:p><a:r><a:t>Item {{i}}</a:t></a:r></a:p>' +
    '<a:p><a:r><a:t>{{END-FOR}}</a:t></a:r></a:p>';
  const pptx = await withNotes(
    { 1: '{{HIDE-SLIDE flowPage == 2 || flowPages > 5}}' },
    { 1: loop },
  );

  const slides = await deck(await renderDeck(pptx, { items: [1, 2, 3] }));
  expect(slides.map((s) => s.text)).toEqual([
    'Item 1',
    'Item 3',
    'Slide 2',
    'Slide 3',
  ]);

  const long = await deck(
    await renderDeck(pptx, { items: [1, 2, 3, 4, 5, 6] }),
  );
  expect(long.map((s) => s.text)).toEqual(['Slide 2', 'Slide 3']);
});

test('renderDeck with every slide hidden writes a deck without slides', async () => {
  const pptx = await withNotes({
    1: '{{HIDE-SLIDE}}',
    2: '{{HIDE-SLIDE}}',
    3: '{{HIDE-SLIDE}}',
  });
  const result = await renderDeck(pptx, {});
  await keep('hide-slide-empty.test.pptx', result);
  expect(await deck(result)).toEqual([]);
});

test('a failing condition rejects with the slide number', async () => {
  const pptx = await withNotes({ 2: '{{HIDE-SLIDE check()}}' });
  const data = {
    check: () => {
      throw new Error('no data');
    },
  };
  const error = await renderDeck(pptx, data).catch((e) => e);
  expect(error).toBeInstanceOf(TemplateRenderError);
  expect(error.message).toBe('slide 2: {{HIDE-SLIDE check()}}: no data');

  // errorHandler decides instead
  const slides = await deck(
    await renderDeck(pptx, data, { errorHandler: () => true }),
  );
  expect(slides.map((s) => s.text)).toEqual(['Slide 1', 'Slide 3']);
});

test('isSlideHidden answers for addSlide users', async () => {
  const pptx = await withNotes({
    2: '{{HIDE-SLIDE projects.length <= 2}}',
    3: '{{HIDE-SLIDE flowPage > 1}}',
  });
  expect(await isSlideHidden(pptx, 1, { projects: [] })).toBe(false);
  expect(await isSlideHidden(pptx, 2, { projects: [] })).toBe(true);
  expect(await isSlideHidden(pptx, 2, { projects: [1, 2, 3] })).toBe(false);
  expect(await isSlideHidden(pptx, 3, {})).toBe(false);
  expect(await isSlideHidden(pptx, 3, {}, { flowPage: 1, flowPages: 2 })).toBe(
    true,
  );
});

test('HIDE-SLIDE on the slide itself is removed with a warning, not applied', async () => {
  const pptx = await withNotes(
    {},
    { 2: '<a:p><a:r><a:t>{{HIDE-SLIDE true}}Slide 2</a:t></a:r></a:p>' },
  );
  const slides = await deck(await renderDeck(pptx, {}));
  expect(slides.map((s) => s.text)).toEqual(['Slide 1', 'Slide 2', 'Slide 3']);

  // renderDeck is quiet (verbosity 0); renderTemplate on its own warns
  const warn = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
  const zip = await new Automizer({ removeExistingSlides: true })
    .loadRoot(pptx)
    .load(pptx, 'deck')
    .addSlide('deck', 2, (slide) => slide.modify(modify.renderTemplate({})))
    .getJSZip();
  const xml = await zip.file('ppt/slides/slide4.xml').async('text');
  expect(xml).toContain('<a:t>Slide 2</a:t>');
  expect(warn).toHaveBeenCalledWith(
    '[pptx-templates]',
    expect.stringContaining('HIDE-SLIDE works only in the speaker notes'),
  );
  warn.mockRestore();

  const commands = await listTemplateCommands(pptx);
  expect(commands).toContainEqual(
    expect.objectContaining({ slide: 2, type: 'HIDE-SLIDE', code: 'true' }),
  );
});

test('the linter checks HIDE-SLIDE in the notes and warns about it elsewhere', async () => {
  const pptx = await withNotes(
    {
      2: ['{{HIDE-SLIDE projekts.length <= 2}}', '{{kunde}}'],
      3: '{{HIDE-SLIDE foo(}}',
    },
    { 1: '<a:p><a:r><a:t>{{HIDE-SLIDE true}}Slide 1</a:t></a:r></a:p>' },
  );
  const { issues } = await lintTemplate(pptx, { data: { projects: [] } });
  const summary = issues.map((i) => [i.slide, i.rule, i.tag]);

  expect(summary).toContainEqual([
    1,
    'unsupported-location',
    '{{HIDE-SLIDE true}}',
  ]);
  expect(summary).toContainEqual([
    2,
    'unknown-variable',
    '{{HIDE-SLIDE projekts.length <= 2}}',
  ]);
  // other tags in the notes are still not rendered
  expect(summary).toContainEqual([2, 'unsupported-location', '{{kunde}}']);
  expect(summary).toContainEqual([3, 'syntax', '{{HIDE-SLIDE foo(}}']);
  expect(
    issues.filter((i) => i.tag?.startsWith('{{HIDE-SLIDE projekts')),
  ).toHaveLength(1);
});

test('flowContinue() goes on with the entries earlier slides do not show', async () => {
  const loop = (list: string, perBox: number) =>
    `<a:p><a:r><a:t>{{FOR i IN flow(${list}, ${perBox}, { measure: false })}}</a:t></a:r></a:p>` +
    '<a:p><a:r><a:t>Item {{i}}</a:t></a:r></a:p>' +
    '<a:p><a:r><a:t>{{END-FOR}}</a:t></a:r></a:p>';
  const pptx = await withNotes(
    {
      // only the first copy of slide 1
      1: '{{HIDE-SLIDE flowPage > 1}}',
      2: '{{HIDE-SLIDE compact}}',
      3: '{{HIDE-SLIDE !flowContinue(items).length}}',
    },
    {
      1: loop('items', 2),
      2: loop('again ? items : flowContinue(items)', 1),
      3: loop('flowContinue(items)', 1),
    },
  );
  const text = async (data: object) =>
    (await deck(await renderDeck(pptx, data))).map((s) => s.text);

  const items = [1, 2, 3, 4];
  expect(await text({ items, again: false, compact: false })).toEqual([
    'Item 1Item 2',
    'Item 3',
    'Item 4',
  ]);
  expect(await text({ items, again: true, compact: false })).toEqual([
    'Item 1Item 2',
    'Item 1',
    'Item 2',
    'Item 3',
    'Item 4',
  ]);
  // slide 2 hidden: slide 3 goes on after slide 1
  expect(await text({ items, again: false, compact: true })).toEqual([
    'Item 1Item 2',
    'Item 3',
    'Item 4',
  ]);
  // everything on slide 1: slide 3 has nothing left
  expect(await text({ items: [1], again: false, compact: true })).toEqual([
    'Item 1',
  ]);
});

const flowLoop = (list: string, perBox: number) =>
  `<a:p><a:r><a:t>{{FOR i IN flow(${list}, ${perBox}, { measure: false })}}</a:t></a:r></a:p>` +
  '<a:p><a:r><a:t>Item {{i}}</a:t></a:r></a:p>' +
  '<a:p><a:r><a:t>{{END-FOR}}</a:t></a:r></a:p>';

test('flowContinue() skips strings and numbers only of the same list', async () => {
  const pptx = await withNotes(
    { 3: '{{HIDE-SLIDE true}}' },
    {
      1: flowLoop('tags', 5),
      2: flowLoop('flowContinue(skills)', 5),
    },
  );
  const out = await renderDeck(pptx, {
    tags: ['Java'],
    skills: ['Java', 'Go'],
  });
  expect((await deck(out)).map((s) => s.text)).toEqual([
    'Item Java',
    'Item JavaItem Go',
  ]);
});

test('flowContinue() of a list derived with flowContinue() counts for the list', async () => {
  const pptx = await withNotes(
    { 1: '{{HIDE-SLIDE flowPage > 1}}' },
    {
      1: flowLoop('items', 1),
      2: flowLoop('flowContinue(items)', 1),
      3: flowLoop('flowContinue(items)', 1),
    },
  );
  const out = await renderDeck(pptx, { items: ['a', 'b', 'c'] });
  // slide 2 shows b and c; slide 3 has nothing left: one empty copy
  expect((await deck(out)).map((s) => s.text)).toEqual([
    'Item a',
    'Item b',
    'Item c',
    '',
  ]);
});

test('flowContinue() with an explicit flowPage option', async () => {
  const pptx = await withNotes(
    {},
    {
      1: flowLoop('items', 2),
      2: flowLoop('flowContinue(items)', 1),
    },
  );
  const out = await renderDeck(
    pptx,
    { items: [1, 2, 3, 4] },
    { flowPage: 0, slides: [1, 2] },
  );
  expect((await deck(out)).map((s) => s.text)).toEqual([
    'Item 1Item 2',
    'Item 3',
  ]);
});
