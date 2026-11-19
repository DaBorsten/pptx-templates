import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { renderDeck } from '../src/index';
import { checkPptxInvariants } from './helpers/pptx-invariants';

/*
 * cleanup (removeExistingSlides + cleanup, as in renderDeck) keeps the
 * masters and layouts of the root template, so it must keep the parts they
 * point at. It removed every embedding and chart that was not copied with
 * a slide, and the master/layout relationships pointed at missing parts.
 */
const templates = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;
const REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

const addRelationship = async (zip: JSZip, rels: string, rel: string) =>
  zip.file(
    rels,
    (await zip.file(rels).async('string')).replace(
      '</Relationships>',
      `${rel}</Relationships>`,
    ),
  );

/**
 * SlideLayoutWithOle.pptx: slideLayout10 already has an OLE object
 * (oleObject1.bin). This adds a second one to slideMaster1
 * (oleObject2.bin), as e.g. think-cell puts them on masters, and a chart
 * with its style, colors and workbook to slideLayout1 (chart99.xml, taken
 * from ChartBarsStacked.pptx).
 */
const masterAndLayoutPartsTemplate = async (): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templates}/SlideLayoutWithOle.pptx`),
  );
  zip.file('ppt/embeddings/oleObject2.bin', 'OLE-OF-MASTER-1');
  await addRelationship(
    zip,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `<Relationship Id="rId99" Type="${REL}oleObject" Target="../embeddings/oleObject2.bin"/>`,
  );

  const charts = await JSZip.loadAsync(
    fs.readFileSync(`${templates}/ChartBarsStacked.pptx`),
  );
  const copy = async (from: string, to: string) =>
    zip.file(to, await charts.file(from).async('nodebuffer'));
  await copy('ppt/charts/chart1.xml', 'ppt/charts/chart99.xml');
  await copy('ppt/charts/style1.xml', 'ppt/charts/style99.xml');
  await copy('ppt/charts/colors1.xml', 'ppt/charts/colors99.xml');
  await copy(
    'ppt/embeddings/Microsoft_Excel_Worksheet.xlsx',
    'ppt/embeddings/Microsoft_Excel_Worksheet99.xlsx',
  );
  zip.file(
    'ppt/charts/_rels/chart99.xml.rels',
    (await charts.file('ppt/charts/_rels/chart1.xml.rels').async('string'))
      .replace(
        'Microsoft_Excel_Worksheet.xlsx',
        'Microsoft_Excel_Worksheet99.xlsx',
      )
      .replace('colors1.xml', 'colors99.xml')
      .replace('style1.xml', 'style99.xml'),
  );
  await addRelationship(
    zip,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `<Relationship Id="rId99" Type="${REL}chart" Target="../charts/chart99.xml"/>`,
  );
  const types = '[Content_Types].xml';
  const override = (part: string, type: string) =>
    `<Override PartName="/ppt/charts/${part}" ContentType="application/${type}"/>`;
  let contentTypes = await zip.file(types).async('string');
  if (!contentTypes.includes('Extension="xlsx"')) {
    contentTypes = contentTypes.replace(
      '<Default ',
      '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Default ',
    );
  }
  zip.file(
    types,
    contentTypes.replace(
      '</Types>',
      override(
        'chart99.xml',
        'vnd.openxmlformats-officedocument.drawingml.chart+xml',
      ) +
        override('style99.xml', 'vnd.ms-office.chartstyle+xml') +
        override('colors99.xml', 'vnd.ms-office.chartcolorstyle+xml') +
        '</Types>',
    ),
  );
  return zip.generateAsync({ type: 'nodebuffer' });
};

const filesIn = (zip: JSZip, dir: string) =>
  Object.keys(zip.files)
    .filter((f) => new RegExp(`^${dir}/[^/]+$`).test(f))
    .sort();

const check = async (name: string, pptx: Buffer) => {
  fs.writeFileSync(`${outputDir}/${name}`, pptx);
  const { errors, knownIssues } = await checkPptxInvariants(pptx);
  expect(errors).toEqual([]);
  expect(knownIssues.filter((i) => /embeddings\/|charts\//.test(i))).toEqual(
    [],
  );

  const zip = await JSZip.loadAsync(pptx);
  expect(filesIn(zip, 'ppt/embeddings')).toEqual([
    'ppt/embeddings/Microsoft_Excel_Worksheet99.xlsx',
    'ppt/embeddings/oleObject1.bin',
    'ppt/embeddings/oleObject2.bin',
  ]);
  expect(filesIn(zip, 'ppt/charts')).toEqual([
    'ppt/charts/chart99.xml',
    'ppt/charts/colors99.xml',
    'ppt/charts/style99.xml',
  ]);
  expect(await zip.file('ppt/embeddings/oleObject2.bin').async('string')).toBe(
    'OLE-OF-MASTER-1',
  );
};

test('renderDeck keeps the embeddings and charts of masters and layouts', async () => {
  const pptx = await renderDeck(await masterAndLayoutPartsTemplate(), {});
  await check('master-layout-parts-render.test.pptx', pptx);
});

test('addSlide + removeExistingSlides + cleanup keeps the embeddings and charts of masters and layouts', async () => {
  const template = await masterAndLayoutPartsTemplate();
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
  await check('master-layout-parts-addslide.test.pptx', pptx);
});
