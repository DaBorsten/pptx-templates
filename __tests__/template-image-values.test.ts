import fs from 'fs';
import os from 'os';
import path from 'path';
import JSZip from 'jszip';
import { renderDeck, templateDataSkeleton } from '../src/index';
import { runCli } from '../src/extensions/cli';
import { deck, shape } from './helpers/lint-deck';

/**
 * Pictures as plain values: `{{IMAGE logo}}` with a Buffer, a data URL or
 * a file path relative to `imageRoot`, without `{ data, extension }`.
 */

const png = fs.readFileSync(`${__dirname}/media/test.png`);
const feather = fs.readFileSync(`${__dirname}/media/feather.png`);

const PARAGRAPH =
  '<a:p><a:r><a:rPr lang="de-DE" sz="1200"/><a:t>This is an almost empty slide.</a:t></a:r></a:p>';

/** EmptySlide.pptx with `tag` as the text of its text box. */
const deckWith = async (tag: string): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-templates/EmptySlide.pptx`),
  );
  const slide = 'ppt/slides/slide1.xml';
  const xml = await zip.file(slide).async('text');
  const text = tag.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  expect(xml).toContain(PARAGRAPH);
  zip.file(slide, xml.replace(PARAGRAPH, `<a:p><a:r><a:t>${text}</a:t></a:r></a:p>`));
  return zip.generateAsync({ type: 'nodebuffer' });
};

/** The media files of the rendered deck. */
const media = async (pptx: Buffer): Promise<Buffer[]> => {
  const zip = await JSZip.loadAsync(pptx);
  const files = Object.keys(zip.files).filter((f) =>
    /^ppt\/media\/.+\.(png|jpe?g|gif)$/.test(f),
  );
  return Promise.all(files.map((f) => zip.file(f).async('nodebuffer')));
};

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-image-'));
  fs.mkdirSync(path.join(dir, 'images'));
  fs.writeFileSync(path.join(dir, 'images', 'vip.png'), png);
  fs.writeFileSync(path.join(dir, 'images', 'standard.png'), feather);
});
afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));

test('a Buffer or data URL is the picture, its type from the first bytes', async () => {
  const pptx = await deckWith('{{IMAGE logo}}');
  for (const logo of [png, `data:image/png;base64,${png.toString('base64')}`]) {
    const [file] = await media(await renderDeck(pptx, { logo }));
    expect(file.equals(png)).toBe(true);
  }
  // Also inside an object, without extension
  const [file] = await media(
    await renderDeck(await deckWith('{{IMAGE ({ data: logo, width: 2 })}}'), {
      logo: png,
    }),
  );
  expect(file.equals(png)).toBe(true);
});

test('a file path is read from imageRoot', async () => {
  const pptx = await deckWith(
    '{{IF vip}}{{IMAGE logoVip}}{{ELSE}}{{IMAGE logoStandard}}{{END-IF}}',
  );
  const values = {
    logoVip: 'images/vip.png',
    logoStandard: 'images/standard.png',
  };
  for (const [vip, expected] of [
    [true, png],
    [false, feather],
  ] as const) {
    const [file] = await media(
      await renderDeck(pptx, { ...values, vip }, { imageRoot: dir }),
    );
    expect(file.equals(expected)).toBe(true);
  }
});

test('file paths need imageRoot and stay inside it', async () => {
  const pptx = await deckWith('{{IMAGE logo}}');
  await expect(renderDeck(pptx, { logo: 'images/vip.png' })).rejects.toThrow(
    /"images\/vip.png" is a file path, but no imageRoot is set/,
  );
  await expect(
    renderDeck(pptx, { logo: '../vip.png' }, { imageRoot: dir }),
  ).rejects.toThrow(/outside of the imageRoot folder/);
  await expect(
    renderDeck(pptx, { logo: 'images/none.png' }, { imageRoot: dir }),
  ).rejects.toThrow(/file "images\/none.png" not found/);
});

test('the data skeleton names image paths and leaves them null', async () => {
  const pptx = await deck([
    shape('Bild', [
      ['{{IMAGE logo}}'],
      ['{{IMAGE ({ data: kunde.vip ? kunde.foto : platzhalter, width: 3 })}}'],
      ['{{FOR p IN produkte}}{{IMAGE p.bild}}{{p.name}}{{END-FOR}}'],
    ]),
  ]);
  const { data, images } = await templateDataSkeleton(pptx);
  expect(images).toEqual([
    'logo',
    'kunde.foto',
    'platzhalter',
    'produkte[].bild',
  ]);
  expect(data).toEqual({
    logo: null,
    kunde: { vip: true, foto: null },
    platzhalter: null,
    produkte: [{ bild: null, name: 'produkte[].name' }],
  });
});

test('pptx-templates render reads paths next to the data file', async () => {
  const deckFile = path.join(dir, 'deck.pptx');
  fs.writeFileSync(deckFile, await deckWith('{{IMAGE logo}}'));
  fs.writeFileSync(
    path.join(dir, 'data.json'),
    JSON.stringify({ logo: 'images/vip.png' }),
  );
  const run = (...args: string[]) =>
    runCli(['render', deckFile, '-o', path.join(dir, 'out.pptx'), ...args], {
      stdout: () => undefined,
      stderr: () => undefined,
    });

  expect(await run('--data', path.join(dir, 'data.json'))).toBe(0);
  let [file] = await media(fs.readFileSync(path.join(dir, 'out.pptx')));
  expect(file.equals(png)).toBe(true);

  // --image-root: paths from another folder
  fs.writeFileSync(
    path.join(dir, 'images', 'data.json'),
    JSON.stringify({ logo: 'standard.png' }),
  );
  expect(await run('--data', path.join(dir, 'images', 'data.json'))).toBe(0);
  [file] = await media(fs.readFileSync(path.join(dir, 'out.pptx')));
  expect(file.equals(feather)).toBe(true);
  expect(
    await run(
      '--data',
      path.join(dir, 'images', 'data.json'),
      '--image-root',
      os.tmpdir(),
    ),
  ).toBe(1);
});
