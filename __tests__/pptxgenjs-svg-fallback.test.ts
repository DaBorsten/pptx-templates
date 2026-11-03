import fs from 'fs';
import path from 'path';
import JSZip from 'jszip';
import Automizer from '../src/automizer';
import { expectXml } from './helpers/expect-xml';

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47]);
const outputFile = 'pptxgenjs-svg-fallback.test.pptx';
const slidePart = 'ppt/slides/slide2.xml';

const resolveTarget = (target: string) =>
  path.posix.normalize(path.posix.join('ppt/slides', target));

test('an SVG generated with pptxgenjs keeps a separate PNG fallback', async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });

  const pres = automizer
    .loadRoot(`RootTemplate.pptx`)
    .load(`EmptySlide.pptx`, 'empty');

  pres.addSlide('empty', 1, (slide) => {
    slide.generate((pptxGenJSSlide) => {
      pptxGenJSSlide.addImage({
        path: `${__dirname}/media/test.png`,
        x: 1,
        y: 1,
      });
      pptxGenJSSlide.addImage({
        path: `${__dirname}/media/test.svg`,
        x: 3,
        y: 1,
      });
    });
  });

  await pres.write(outputFile);

  const slide = await expectXml(outputFile, slidePart);
  slide.toContainElementTimes('asvg:svgBlip', 1);
  const svgBlip = slide.elements('asvg:svgBlip')[0];
  const fallbackBlip = slide
    .elements('a:blip')
    .find((blip) => blip.getElementsByTagName('asvg:svgBlip').length > 0);

  const rels = await expectXml(outputFile, 'ppt/slides/_rels/slide2.xml.rels');
  const targetOf = (rId: string) =>
    rels
      .elements('Relationship')
      .find((rel) => rel.getAttribute('Id') === rId)
      ?.getAttribute('Target');

  const svgTarget = targetOf(svgBlip.getAttribute('r:embed'));
  const fallbackTarget = targetOf(fallbackBlip.getAttribute('r:embed'));
  expect(svgTarget).toMatch(/\.svg$/);
  expect(fallbackTarget).toMatch(/\.png$/);
  expect(fallbackTarget).not.toBe(svgTarget);

  const archive = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-output/${outputFile}`),
  );
  const read = (target: string) =>
    archive.file(resolveTarget(target)).async('nodebuffer');

  // The SVG part holds the SVG, the fallback a real PNG (not the SVG bytes)
  expect((await read(svgTarget)).toString('utf8')).toMatch(/^<svg/);
  expect((await read(fallbackTarget)).subarray(0, 4)).toEqual(PNG_SIGNATURE);

  // The regular PNG next to it is copied unchanged
  const original = fs.readFileSync(`${__dirname}/media/test.png`);
  const media = archive.file(/^ppt\/media\//);
  const contents = await Promise.all(media.map((f) => f.async('nodebuffer')));
  expect(contents.some((content) => content.equals(original))).toBe(true);
  expect(media).toHaveLength(3);

  const contentTypes = await expectXml(outputFile, '[Content_Types].xml');
  contentTypes.toHaveAttribute('Default', 'Extension', 'svg');
});
