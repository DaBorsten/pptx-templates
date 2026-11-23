import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { countFlowPages, renderDeck } from '../src/index';
import { FontRegistry } from '../src/extensions/text-flow/fonts';
import { generatePptx } from '../src/extensions/template-render/zip-writer';
import { removeUnreachableParts } from '../src/extensions/template-render/prune';
import { readText } from '../src/extensions/template-render/zip-text';
import { checkPptxInvariants } from './helpers/pptx-invariants';

/*
 * What renderDeck does instead of Automizer's `cleanup` and JSZip's
 * generateAsync, and that it counts flow() pages without building a deck.
 */
const templates = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;
const fonts = [`${__dirname}/fonts`];
const REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

const parts = (zip: JSZip) =>
  Object.keys(zip.files)
    .filter((name) => !zip.files[name].dir)
    .sort();

/** Compression method of each part, from the local file headers. */
const methods = (pptx: Buffer): Map<string, number> => {
  const result = new Map<string, number>();
  let at = 0;
  while (pptx.readUInt32LE(at) === 0x04034b50) {
    const method = pptx.readUInt16LE(at + 8);
    const size = pptx.readUInt32LE(at + 18);
    const nameLength = pptx.readUInt16LE(at + 26);
    const extra = pptx.readUInt16LE(at + 28);
    const name = pptx.toString('utf8', at + 30, at + 30 + nameLength);
    result.set(name, method);
    at += 30 + nameLength + extra + size;
  }
  return result;
};

test('renderDeck keeps only the parts its slides reach', async () => {
  // slide 1 shows image1.jpg; slide 2 also image2.png, image3.png, image4.svg
  const template = fs.readFileSync(`${templates}/SlideWithImages.pptx`);
  const pptx = await renderDeck(template, {}, { slides: [1] });
  fs.writeFileSync(`${outputDir}/template-render-reachable.pptx`, pptx);

  const { errors, knownIssues } = await checkPptxInvariants(pptx);
  expect(errors).toEqual([]);
  expect(knownIssues).toEqual([]);

  const zip = await JSZip.loadAsync(pptx);
  const names = parts(zip);
  expect(names.filter((n) => n.startsWith('ppt/media/'))).toEqual([
    'ppt/media/image1.jpg',
  ]);
  // the template's own slides are gone, the rendered one is slide3
  expect(names.filter((n) => n.startsWith('ppt/slides/'))).toEqual([
    'ppt/slides/_rels/slide3.xml.rels',
    'ppt/slides/slide3.xml',
  ]);
  // masters, layouts and theme stay
  expect(names).toContain('ppt/slideLayouts/slideLayout1.xml');
  expect(names).toContain('ppt/theme/theme1.xml');

  // no content type for a part that is not there
  const types = await zip.file('[Content_Types].xml').async('string');
  const overrides = [...types.matchAll(/PartName="\/([^"]+)"/g)].map(
    (m) => m[1],
  );
  expect(overrides.filter((part) => !names.includes(part))).toEqual([]);
  expect(overrides).toContain('ppt/slides/slide3.xml');
});

test('a picture whose relationship the part does not use is removed', async () => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templates}/SlideWithImages.pptx`),
  );
  zip.file(
    'ppt/media/unused.png',
    zip.file('ppt/media/image3.png').async('nodebuffer'),
  );
  const rels = 'ppt/slides/_rels/slide1.xml.rels';
  zip.file(
    rels,
    (await zip.file(rels).async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rId77" Type="${REL}image" Target="../media/unused.png"/></Relationships>`,
    ),
  );
  // also reachable through another relationship type: stays
  zip.file('ppt/media/linked.png', 'PNG');
  zip.file(
    rels,
    (await zip.file(rels).async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rId78" Type="${REL}oleObject" Target="../media/linked.png"/></Relationships>`,
    ),
  );

  await removeUnreachableParts(zip);
  const names = parts(zip);
  expect(names).not.toContain('ppt/media/unused.png');
  expect(names).toContain('ppt/media/linked.png');
  // used by <a:blip r:embed="rId2"> of slide 1
  expect(names).toContain('ppt/media/image1.jpg');
  // the relationship stays, as with Automizer's cleanup
  expect(await zip.file(rels).async('string')).toContain('rId77');
});

test('generatePptx writes a zip JSZip reads back, deflated parts reused', async () => {
  const source = await JSZip.loadAsync(
    fs.readFileSync(`${templates}/SlideWithImages.pptx`),
  );
  source.file('ppt/slides/slide1.xml', '<?xml version="1.0"?><p:sld>Ä</p:sld>');
  source.file('docProps/custom.bin', new Uint8Array([0, 1, 2, 250]));
  const pptx = await generatePptx(source);

  const zip = await JSZip.loadAsync(pptx, { checkCRC32: true });
  expect(parts(zip)).toEqual(parts(source));
  expect(Object.keys(zip.files)[0]).toBe('[Content_Types].xml');
  for (const name of parts(source)) {
    expect(await zip.file(name).async('nodebuffer')).toEqual(
      await source.file(name).async('nodebuffer'),
    );
  }
  expect(await zip.file('ppt/slides/slide1.xml').async('string')).toContain(
    'Ä',
  );

  const method = methods(pptx);
  // changed and XML parts deflated, stored pictures stay stored
  expect(method.get('ppt/slides/slide1.xml')).toBe(8);
  expect(method.get('docProps/custom.bin')).toBe(8);
  expect(method.get('ppt/slides/slide2.xml')).toBe(8);
  expect(method.get('ppt/media/image1.jpg')).toBe(0);
});

test('readText reads parts as JSZip does', async () => {
  const source = new JSZip();
  source.file('a.xml', '﻿<a>Ä</a>');
  source.file('b.xml', 'stored', { compression: 'STORE' });
  const zip = await JSZip.loadAsync(
    await source.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }),
  );
  zip.file('c.xml', 'changed');
  for (const name of ['a.xml', 'b.xml', 'c.xml']) {
    expect(await readText(zip, name)).toBe(
      await zip.file(name).async('string'),
    );
  }
  expect(await readText(zip, 'missing.xml')).toBeUndefined();
});

/** One slide with three flow() columns, as in text-flow.test.ts. */
const flowTemplate = async (file: string): Promise<string> => {
  await new Automizer({
    templateDir: templates,
    outputDir,
    removeExistingSlides: true,
    verbosity: 0,
  })
    .loadRoot('EmptySlide.pptx')
    .load('EmptySlide.pptx', 'empty')
    .addSlide('empty', 1, (slide) => {
      slide.generate((s) => {
        s.addText('Page {{flowPage}} of {{flowPages}} {{LINK site}}', {
          x: 0.5,
          y: 0.2,
          w: 12,
          h: 0.4,
          fontSize: 12,
          fontFace: 'Liberation Sans',
        });
        [0.5, 4.68, 8.86].forEach((x) =>
          s.addText(
            ['{{FOR item IN flow(items)}}', '{{item}}', '{{END-FOR}}'].map(
              (text) => ({ text, options: { breakLine: true } }),
            ),
            {
              x,
              y: 1,
              w: 3.95,
              h: 6.2,
              fontSize: 11,
              fontFace: 'Liberation Sans',
              valign: 'top',
            },
          ),
        );
      });
    })
    .write(file);
  return `${outputDir}/${file}`;
};

test('renderDeck counts flow() pages without building a deck per slide', async () => {
  const template = await flowTemplate('template-render-count-template.pptx');
  const data = {
    items: Array.from({ length: 150 }, (_, i) => `Entry ${i + 1}`),
    site: { url: 'https://example.com', label: 'Site' },
  };
  expect(await countFlowPages(template, 2, data, { fonts })).toBe(2);

  const getJSZip = jest.spyOn(Automizer.prototype, 'getJSZip');
  try {
    const pptx = await renderDeck(template, data, { fonts });
    // one Automizer run: the deck itself
    expect(getJSZip).toHaveBeenCalledTimes(1);
    const zip = await JSZip.loadAsync(pptx);
    const presentation = await zip.file('ppt/presentation.xml').async('string');
    expect(presentation.match(/<p:sldId /g)).toHaveLength(2);
  } finally {
    getJSZip.mockRestore();
  }
});

test('the font folders are listed once per process', () => {
  const readdir = jest.spyOn(fs, 'readdirSync');
  try {
    new FontRegistry(fonts, false).face('Liberation Sans');
    const first = readdir.mock.calls.length;
    new FontRegistry(fonts, false).face('Liberation Sans');
    expect(readdir.mock.calls.length).toBe(first);
  } finally {
    readdir.mockRestore();
  }
});
