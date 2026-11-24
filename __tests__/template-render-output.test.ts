import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { countFlowPages, renderDeck } from '../src/index';
import { FontRegistry } from '../src/extensions/text-flow/fonts';
import { generatePptx } from '../src/extensions/template-render/zip-writer';
import { removeUnreachableParts } from '../src/extensions/template-render/prune';
import { readText } from '../src/extensions/template-render/zip-text';
import { checkPptxInvariants } from './helpers/pptx-invariants';
import ArchiveJszip from '../src/helper/archive/archive-jszip';
import { MediaDeduplicator } from '../src/helper/media-deduplicator';
import * as images from '../src/extensions/template-logic/images';

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

test('the system font folders are listed once per process', () => {
  // a typeface no folder has: the system fonts are searched
  const lookUp = () => {
    try {
      new FontRegistry([], true).face('No Such Typeface 7c1e');
    } catch {
      // a machine without system fonts: nothing to measure with
    }
  };
  lookUp();
  const readdir = jest.spyOn(fs, 'readdirSync');
  try {
    lookUp();
    expect(readdir).not.toHaveBeenCalled();
  } finally {
    readdir.mockRestore();
  }
});

test('a font added to a fonts folder is found by the next render', () => {
  const folder = `${outputDir}/font-folder-added`;
  fs.rmSync(folder, { recursive: true, force: true });
  fs.mkdirSync(folder, { recursive: true });
  // empty folder, no system fonts: nothing to measure with
  expect(() =>
    new FontRegistry([folder], false).face('Liberation Sans'),
  ).toThrow(/no font to measure/);

  fs.copyFileSync(
    `${fonts[0]}/LiberationSans-Regular.ttf`,
    `${folder}/LiberationSans-Regular.ttf`,
  );
  const face = new FontRegistry([folder], false).face('Liberation Sans');
  expect(face.info.file.replace(/\\/g, '/')).toBe(
    `${folder}/LiberationSans-Regular.ttf`.replace(/\\/g, '/'),
  );
});

describe('MediaDeduplicator hashes only pictures of the same size', () => {
  const archiveWith = async (files: Record<string, Buffer>) => {
    const zip = new JSZip();
    Object.entries(files).forEach(([name, content]) =>
      zip.file(`ppt/media/${name}`, content),
    );
    const archive = new ArchiveJszip(
      await zip.generateAsync({ type: 'nodebuffer' }),
      { mode: 'jszip' },
    );
    archive.archive = await JSZip.loadAsync(
      await zip.generateAsync({ type: 'nodebuffer' }),
    );
    return archive;
  };

  test('finds identical files, reads no file of another size', async () => {
    const archive = await archiveWith({
      'a.png': Buffer.from('AAAA'),
      'b.png': Buffer.from('BBBB'),
      'c.png': Buffer.from('CCCCCCCC'),
      'd.png': Buffer.from('AAAA'),
    });
    const read = jest.spyOn(archive, 'read');
    const media = new MediaDeduplicator(archive);

    expect(await media.find(Buffer.from('BBBB'))).toBe('b.png');
    // the first of two identical files, as before
    expect(await media.find(Buffer.from('AAAA'))).toBe('a.png');
    expect(await media.find(Buffer.from('XXXX'))).toBeUndefined();
    expect(read.mock.calls.map((c) => c[0]).sort()).toEqual([
      'ppt/media/a.png',
      'ppt/media/b.png',
      'ppt/media/d.png',
    ]);

    // an added file is found; an existing identical one comes first
    media.add(Buffer.from('NEW!'), 'new.png');
    media.add(Buffer.from('CCCCCCCC'), 'copy.png');
    expect(await media.find(Buffer.from('NEW!'))).toBe('new.png');
    expect(await media.find(Buffer.from('CCCCCCCC'))).toBe('c.png');
  });

  test('without sizes it hashes every file up front', async () => {
    const archive = await archiveWith({ 'a.png': Buffer.from('AAAA') });
    // a part without a known size: e.g. another archive type
    (
      archive.archive.files['ppt/media/a.png'] as unknown as { _data: unknown }
    )._data = undefined;
    const media = new MediaDeduplicator(archive);
    jest
      .spyOn(archive, 'read')
      .mockResolvedValue(Buffer.from('AAAA') as unknown as string);
    expect(await media.find(Buffer.from('AAAA'))).toBe('a.png');
  });
});

test('flow() trial renders and the count load no IMAGE pictures', async () => {
  const template = await flowTemplate('template-render-image-template.pptx');
  // the header box of the flow slide shows a picture
  const zip = await JSZip.loadAsync(fs.readFileSync(template));
  const slide = 'ppt/slides/slide2.xml';
  zip.file(
    slide,
    (await zip.file(slide).async('string')).replace(
      '{{LINK site}}',
      '{{IMAGE photo}}',
    ),
  );
  const data = {
    items: Array.from({ length: 150 }, (_, i) => `Entry ${i + 1}`),
    photo: fs.readFileSync(`${__dirname}/media/test.png`),
  };
  const loadImage = jest.spyOn(images, 'loadImage');
  try {
    const pptx = await renderDeck(
      await zip.generateAsync({ type: 'nodebuffer' }),
      data,
      { fonts },
    );
    // once per slide of the output: 2 pages
    expect(loadImage).toHaveBeenCalledTimes(2);
    const out = await JSZip.loadAsync(pptx);
    const media = parts(out).filter((n) => n.startsWith('ppt/media/'));
    expect(media).toHaveLength(1);
    // a picture the render adds is stored, as PowerPoint does
    expect(methods(pptx).get(media[0])).toBe(0);
    const { errors } = await checkPptxInvariants(pptx);
    expect(errors).toEqual([]);
  } finally {
    loadImage.mockRestore();
  }
});
