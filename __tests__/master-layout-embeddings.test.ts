import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { renderDeck } from '../src/index';
import { checkPptxInvariants } from './helpers/pptx-invariants';

/*
 * cleanup (removeExistingSlides + cleanup, as in renderDeck) keeps the
 * masters and layouts of the root template, so it must keep the embeddings
 * they point at. It removed every embedding that was not copied with a
 * slide, and the master/layout relationships pointed at missing parts.
 */
const templates = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;
const REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

/**
 * SlideLayoutWithOle.pptx: slideLayout10 already has an OLE object
 * (oleObject1.bin). This adds a second one to slideMaster1
 * (oleObject2.bin), as e.g. think-cell puts them on masters.
 */
const masterAndLayoutOleTemplate = async (): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templates}/SlideLayoutWithOle.pptx`),
  );
  zip.file('ppt/embeddings/oleObject2.bin', 'OLE-OF-MASTER-1');
  const rels = 'ppt/slideMasters/_rels/slideMaster1.xml.rels';
  zip.file(
    rels,
    (await zip.file(rels).async('string')).replace(
      '</Relationships>',
      `<Relationship Id="rId99" Type="${REL}oleObject" Target="../embeddings/oleObject2.bin"/></Relationships>`,
    ),
  );
  return zip.generateAsync({ type: 'nodebuffer' });
};

const embeddingsAndTargets = async (pptx: Buffer) => {
  const zip = await JSZip.loadAsync(pptx);
  const files = Object.keys(zip.files)
    .filter((f) => /^ppt\/embeddings\/[^/]+$/.test(f))
    .sort();
  const targets: string[] = [];
  for (const f of Object.keys(zip.files)) {
    if (!/^ppt\/slide(Master|Layout)s\/_rels\/.+\.rels$/.test(f)) continue;
    const xml = await zip.file(f).async('string');
    for (const m of xml.matchAll(/Target="\.\.\/embeddings\/([^"]+)"/g)) {
      targets.push(`ppt/embeddings/${m[1]}`);
    }
  }
  return { zip, files, targets: [...new Set(targets)].sort() };
};

const check = async (name: string, pptx: Buffer) => {
  fs.writeFileSync(`${outputDir}/${name}`, pptx);
  const { errors, knownIssues } = await checkPptxInvariants(pptx);
  expect(errors).toEqual([]);
  expect(knownIssues.filter((i) => i.includes('embeddings/'))).toEqual([]);

  const { zip, files, targets } = await embeddingsAndTargets(pptx);
  expect(targets).toEqual([
    'ppt/embeddings/oleObject1.bin',
    'ppt/embeddings/oleObject2.bin',
  ]);
  expect(files).toEqual(targets);
  expect(
    await zip.file('ppt/embeddings/oleObject2.bin').async('string'),
  ).toBe('OLE-OF-MASTER-1');
};

test('renderDeck keeps the embeddings of masters and layouts', async () => {
  const pptx = await renderDeck(await masterAndLayoutOleTemplate(), {});
  await check('master-layout-embeddings-render.test.pptx', pptx);
});

test('addSlide + removeExistingSlides + cleanup keeps the embeddings of masters and layouts', async () => {
  const template = await masterAndLayoutOleTemplate();
  const pres = new Automizer({
    templateDir: templates,
    outputDir,
    removeExistingSlides: true,
    cleanup: true,
  });
  pres.loadRoot(template).load(template, 'src');
  pres.addSlide('src', 1);
  const zip = await pres.getJSZip();
  const pptx = await zip.generateAsync({ type: 'nodebuffer' });
  await check('master-layout-embeddings-addslide.test.pptx', pptx);
});
