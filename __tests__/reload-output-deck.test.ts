import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { renderDeck } from '../src/index';

/*
 * A deck written by pptx-templates, loaded again as root: its slide parts
 * are not numbered 1..n (removeExistingSlides + cleanup leave slide4.xml,
 * slide5.xml, …). New slides must not reuse those numbers.
 */
const templates = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;

const slideParts = (zip: JSZip) =>
  Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));

const duplicateOverrides = async (zip: JSZip) => {
  const xml = await zip.file('[Content_Types].xml').async('string');
  const names = [...xml.matchAll(/PartName="([^"]+)"/g)].map((m) => m[1]);
  return names.filter((name, i) => names.indexOf(name) !== i);
};

/** Slide texts in presentation order. */
const texts = async (zip: JSZip) => {
  const presentation = await zip.file('ppt/presentation.xml').async('string');
  const rels = await zip
    .file('ppt/_rels/presentation.xml.rels')
    .async('string');
  const result: string[] = [];
  for (const [, id] of presentation.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g)) {
    const target = new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels)[1];
    const xml = await zip.file(`ppt/${target}`).async('string');
    result.push(
      [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('|'),
    );
  }
  return result;
};

const rendered = () =>
  renderDeck(fs.readFileSync(`${templates}/threeSlidePres.pptx`), {});

test('the output of renderDeck starts above slide 1', async () => {
  // the precondition of the tests below
  const zip = await JSZip.loadAsync(await rendered());
  expect(slideParts(zip).sort()).toEqual([
    'ppt/slides/slide4.xml',
    'ppt/slides/slide5.xml',
    'ppt/slides/slide6.xml',
  ]);
});

test('slides added to a reloaded deck do not overwrite its slides', async () => {
  const deck = await rendered();
  const pres = new Automizer({ verbosity: 0 })
    .loadRoot(deck)
    .load(`${templates}/threeSlidePres.pptx`, 'tpl');
  pres.addSlide('tpl', 1);
  pres.addSlide('tpl', 2);

  const zip = await pres.getJSZip();
  fs.writeFileSync(
    `${outputDir}/reload-output-deck-append.test.pptx`,
    await zip.generateAsync({ type: 'nodebuffer' }),
  );

  expect(await texts(zip)).toEqual([
    'Slide 1',
    'Slide 2',
    'Slide 3',
    'Slide 1',
    'Slide 2',
  ]);
  expect(slideParts(zip)).toHaveLength(5);
  expect(await duplicateOverrides(zip)).toEqual([]);
});

test('a reloaded deck can be copied into itself (two passes)', async () => {
  const deck = await rendered();
  const pres = new Automizer({
    removeExistingSlides: true,
    cleanup: true,
    verbosity: 0,
  })
    .loadRoot(deck)
    .load(deck, 'deck');
  for (const n of await pres.getTemplate('deck').getAllSlideNumbers()) {
    pres.addSlide('deck', n);
  }

  const zip = await pres.getJSZip();
  fs.writeFileSync(
    `${outputDir}/reload-output-deck-two-pass.test.pptx`,
    await zip.generateAsync({ type: 'nodebuffer' }),
  );

  expect(await texts(zip)).toEqual(['Slide 1', 'Slide 2', 'Slide 3']);
  expect(slideParts(zip)).toHaveLength(3);
  expect(await duplicateOverrides(zip)).toEqual([]);
  // every slide keeps its own notes part
  const notes = Object.keys(zip.files).filter((f) =>
    /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(f),
  );
  expect(notes).toHaveLength(3);
});
