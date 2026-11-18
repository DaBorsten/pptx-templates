import fs from 'fs';
import JSZip from 'jszip';
import Automizer, { renderDeck } from '../src/index';
import { checkPptxInvariants } from './helpers/pptx-invariants';

/*
 * Embedded objects and speaker notes when slides are copied: several copies
 * of one slide (flow()), slides left out (HIDE-SLIDE), and embeddings of
 * different slides in one deck.
 */
const templates = `${__dirname}/pptx-templates`;
const outputDir = `${__dirname}/pptx-output`;
const REL =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

const oleFrame = (id: number, rId: string, imageRId: string) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Embedded ${id}"/>` +
  '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
  '<p:xfrm><a:off x="100000" y="100000"/><a:ext cx="1000000" cy="1000000"/></p:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">' +
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
  `<mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v"><p:oleObj name="Object" r:id="${rId}" imgW="100" imgH="100" progId="Excel.Sheet.12"><p:embed/></p:oleObj></mc:Choice>` +
  `<mc:Fallback><p:oleObj name="Object" r:id="${rId}" imgW="100" imgH="100" progId="Excel.Sheet.12"><p:embed/>` +
  `<p:pic><p:nvPicPr><p:cNvPr id="${id + 100}" name="Preview ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="${imageRId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  '<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
  '</p:oleObj></mc:Fallback></mc:AlternateContent></a:graphicData></a:graphic></p:graphicFrame>';

const replaceNotes = (xml: string, lines: string[]) =>
  xml.replace(
    /<a:p>(?:(?!<a:p>)[\s\S])*?<a:t><\/a:t>[\s\S]*?<\/a:p>/,
    lines.map((l) => `<a:p><a:r><a:t>${l}</a:t></a:r></a:p>`).join(''),
  );

/**
 * threeSlidePres.pptx with an embedded object on every slide, as
 * PowerPoint writes them: slide 1 an OLE object (oleObject1.bin), slide 2
 * an embedded Excel worksheet (a "package" relationship), slide 3 another
 * OLE object (oleObject2.bin). The notes point back to their slides.
 */
const embeddingsTemplate = async (
  options: {
    slide1?: string;
    notes?: Record<number, string[]>;
  } = {},
): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${templates}/threeSlidePres.pptx`),
  );
  const png = await (
    await JSZip.loadAsync(
      fs.readFileSync(`${templates}/SlideLayoutWithOle.pptx`),
    )
  )
    .file('ppt/media/image4.png')
    .async('nodebuffer');
  zip.file('ppt/media/preview1.png', png);
  zip.file('ppt/embeddings/oleObject1.bin', 'OLE-OF-SLIDE-1');
  zip.file('ppt/embeddings/Microsoft_Excel_Worksheet.xlsx', 'XLSX-OF-SLIDE-2');
  zip.file('ppt/embeddings/oleObject2.bin', 'OLE-OF-SLIDE-3');
  const embeds: Record<number, [string, string]> = {
    1: ['oleObject', '../embeddings/oleObject1.bin'],
    2: ['package', '../embeddings/Microsoft_Excel_Worksheet.xlsx'],
    3: ['oleObject', '../embeddings/oleObject2.bin'],
  };
  for (const [n, [type, target]] of Object.entries(embeds)) {
    const rels = `ppt/slides/_rels/slide${n}.xml.rels`;
    zip.file(
      rels,
      (await zip.file(rels).async('string')).replace(
        '</Relationships>',
        `<Relationship Id="rId7" Type="${REL}${type}" Target="${target}"/>` +
          `<Relationship Id="rId8" Type="${REL}image" Target="../media/preview1.png"/></Relationships>`,
      ),
    );
    const slide = `ppt/slides/slide${n}.xml`;
    zip.file(
      slide,
      (await zip.file(slide).async('string')).replace(
        '</p:spTree>',
        `${oleFrame(90 + Number(n), 'rId7', 'rId8')}</p:spTree>`,
      ),
    );
    const notesRels = `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`;
    zip.file(
      notesRels,
      (await zip.file(notesRels).async('string')).replace(
        '</Relationships>',
        `<Relationship Id="rId2" Type="${REL}slide" Target="../slides/slide${n}.xml"/></Relationships>`,
      ),
    );
  }
  for (const [n, lines] of Object.entries(options.notes ?? {})) {
    const part = `ppt/notesSlides/notesSlide${n}.xml`;
    zip.file(part, replaceNotes(await zip.file(part).async('string'), lines));
  }
  if (options.slide1) {
    const part = 'ppt/slides/slide1.xml';
    zip.file(
      part,
      (await zip.file(part).async('string')).replace(
        /(<p:txBody>[\s\S]*?<a:lstStyle\/>)[\s\S]*?(<\/p:txBody>)/,
        `$1${options.slide1}$2`,
      ),
    );
  }
  let contentTypes = await zip.file('[Content_Types].xml').async('string');
  for (const [ext, type] of [
    ['bin', 'application/vnd.openxmlformats-officedocument.oleObject'],
    [
      'xlsx',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ],
    ['png', 'image/png'],
  ]) {
    if (!contentTypes.includes(`Extension="${ext}"`)) {
      contentTypes = contentTypes.replace(
        '<Default ',
        `<Default Extension="${ext}" ContentType="${type}"/><Default `,
      );
    }
  }
  zip.file('[Content_Types].xml', contentTypes);
  return zip.generateAsync({ type: 'nodebuffer' });
};

type OutputSlide = {
  part: string;
  text: string;
  notes?: { part: string; text: string; backTo: string[] };
  embeddings: { type: string; part: string; content: string }[];
};

const relsOf = (part: string) =>
  part.replace(/^(.*)\/([^/]+)$/, '$1/_rels/$2.rels');

const relationships = (xml: string) =>
  [...xml.matchAll(/<Relationship ([^>]*)\/>/g)].map((m) => ({
    id: /Id="([^"]+)"/.exec(m[1])[1],
    type: /Type="[^"]*\/([^"/]+)"/.exec(m[1])[1],
    target: /Target="([^"]+)"/.exec(m[1])[1],
  }));

const resolve = (from: string, target: string) => {
  const parts = from.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop();
    else parts.push(segment);
  }
  return parts.join('/');
};

const textOf = (xml: string) =>
  [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('|');

/** The slides of a deck in presentation order, with notes and embeddings. */
const inspect = async (pptx: Buffer) => {
  const zip = await JSZip.loadAsync(pptx);
  const read = (part: string) => zip.file(part)?.async('string');
  const presentation = await read('ppt/presentation.xml');
  const presentationRels = relationships(
    await read('ppt/_rels/presentation.xml.rels'),
  );
  const slides: OutputSlide[] = [];
  for (const id of [
    ...presentation.matchAll(/<p:sldId [^>]*r:id="([^"]+)"/g),
  ].map((m) => m[1])) {
    const part = resolve(
      'ppt/presentation.xml',
      presentationRels.find((r) => r.id === id).target,
    );
    const rels = relationships(await read(relsOf(part)));
    const slide: OutputSlide = {
      part,
      text: textOf(await read(part)),
      embeddings: [],
    };
    for (const rel of rels) {
      const target = resolve(part, rel.target);
      if (rel.type === 'notesSlide') {
        slide.notes = {
          part: target,
          text: textOf(await read(target)),
          backTo: relationships(await read(relsOf(target)))
            .filter((r) => r.type === 'slide')
            .map((r) => resolve(target, r.target)),
        };
      }
      if (/embeddings\//.test(target)) {
        slide.embeddings.push({
          type: rel.type,
          part: target,
          content: (await read(target)) ?? 'MISSING',
        });
      }
    }
    slides.push(slide);
  }
  const contentTypes = await read('[Content_Types].xml');
  const partNames = [...contentTypes.matchAll(/PartName="([^"]+)"/g)].map(
    (m) => m[1],
  );
  return {
    zip,
    slides,
    duplicateContentTypes: partNames.filter(
      (p, i) => partNames.indexOf(p) !== i,
    ),
    notesParts: Object.keys(zip.files).filter((f) =>
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(f),
    ),
  };
};

const keep = async (name: string, pptx: Buffer) => {
  fs.writeFileSync(`${outputDir}/${name}`, pptx);
  const { errors } = await checkPptxInvariants(pptx);
  expect(errors).toEqual([]);
};

test('renderDeck: every copy of a flow() slide gets its own embedding and notes', async () => {
  const loop =
    '<a:p><a:r><a:t>{{FOR i IN flow(items, 1, { measure: false })}}</a:t></a:r></a:p>' +
    '<a:p><a:r><a:t>Item {{i}}</a:t></a:r></a:p>' +
    '<a:p><a:r><a:t>{{END-FOR}}</a:t></a:r></a:p>';
  const template = await embeddingsTemplate({
    slide1: loop,
    notes: {
      1: ['Talk about {{i}}', '{{HIDE-SLIDE flowPage == 2}}'],
      2: ['Internal: do not ship', '{{HIDE-SLIDE hideExcel}}'],
    },
  });

  const pptx = await renderDeck(template, {
    items: [1, 2, 3],
    hideExcel: false,
  });
  await keep('embeddings-notes-flow.test.pptx', pptx);
  const { slides, duplicateContentTypes, notesParts } = await inspect(pptx);

  expect(slides.map((s) => s.text)).toEqual([
    'Item |1',
    'Item |3',
    'Slide 2',
    'Slide 3',
  ]);
  // Each slide has its own notes part that points back to it
  slides.forEach((slide) => {
    expect(slide.notes.backTo).toEqual([slide.part]);
  });
  expect(new Set(slides.map((s) => s.notes.part)).size).toBe(4);
  expect(slides[0].notes.text).toBe('Talk about {{i}}');
  expect(slides[1].notes.text).toBe('Talk about {{i}}');
  expect(slides[2].notes.text).toBe('Internal: do not ship');
  // Notes parts only for the slides of the output
  expect(notesParts.sort()).toEqual(slides.map((s) => s.notes.part).sort());

  // Every embedding exists, with the content of its own source slide
  expect(slides.map((s) => s.embeddings.map((e) => e.content))).toEqual([
    ['OLE-OF-SLIDE-1'],
    ['OLE-OF-SLIDE-1'],
    ['XLSX-OF-SLIDE-2'],
    ['OLE-OF-SLIDE-3'],
  ]);
  // The embedded worksheet stays a package, with its extension
  expect(slides[2].embeddings[0].type).toBe('package');
  expect(slides[2].embeddings[0].part).toMatch(/\.xlsx$/);
  // The copies do not share a part
  const parts = slides.flatMap((s) => s.embeddings.map((e) => e.part));
  expect(new Set(parts).size).toBe(parts.length);
  expect(duplicateContentTypes).toEqual([]);
});

test('renderDeck: the notes of a hidden slide are not in the output', async () => {
  const template = await embeddingsTemplate({
    notes: { 2: ['Internal: do not ship', '{{HIDE-SLIDE hideExcel}}'] },
  });

  const pptx = await renderDeck(template, { hideExcel: true });
  await keep('embeddings-notes-hidden.test.pptx', pptx);
  const { zip, slides, notesParts } = await inspect(pptx);

  expect(slides.map((s) => s.text)).toEqual(['Slide 1', 'Slide 3']);
  expect(notesParts.sort()).toEqual(slides.map((s) => s.notes.part).sort());
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir) continue;
    const content = await zip.file(name).async('string');
    expect(content).not.toContain('Internal: do not ship');
  }
  // Nor its embedding
  expect(Object.keys(zip.files).filter((f) => /\.xlsx$/.test(f))).toEqual([]);
});

test('addSlide: embeddings of different slides do not overwrite each other', async () => {
  const template = await embeddingsTemplate();
  const pres = new Automizer({
    removeExistingSlides: true,
    verbosity: 0,
  })
    .loadRoot(template)
    .load(template, 'embeddings');
  // Slide 1 and 3 both get the relationship id of the next free rId
  pres.addSlide('embeddings', 1);
  pres.addSlide('embeddings', 3);
  pres.addSlide('embeddings', 2);

  const pptx = await (
    await pres.getJSZip()
  ).generateAsync({ type: 'nodebuffer' });
  await keep('embeddings-notes-addslide.test.pptx', pptx);
  const { slides, duplicateContentTypes } = await inspect(pptx);

  expect(slides.map((s) => s.embeddings.map((e) => e.content))).toEqual([
    ['OLE-OF-SLIDE-1'],
    ['OLE-OF-SLIDE-3'],
    ['XLSX-OF-SLIDE-2'],
  ]);
  expect(duplicateContentTypes).toEqual([]);
});

test('addSlide: a copied chartEx chart does not overwrite the charts of the deck', async () => {
  const pres = new Automizer({
    templateDir: templates,
    verbosity: 0,
  })
    .loadRoot('ChartWaterfall.pptx')
    .load('ChartWaterfall.pptx', 'waterfall');
  pres.addSlide('waterfall', 1);

  const pptx = await (
    await pres.getJSZip()
  ).generateAsync({ type: 'nodebuffer' });
  await keep('embeddings-notes-chartex.test.pptx', pptx);
  const { zip, slides, duplicateContentTypes } = await inspect(pptx);

  // The charts a slide shows (a copied slide keeps the unused relationship
  // of the source slide's chart, a known upstream leftover)
  const chartsOf = async (part: string) => {
    const xml = await zip.file(part).async('string');
    return relationships(await zip.file(relsOf(part)).async('string'))
      .filter((r) => r.type === 'chartEx' && xml.includes(`"${r.id}"`))
      .map((r) => resolve(part, r.target));
  };
  expect(await chartsOf(slides[0].part)).toEqual(['ppt/charts/chartEx1.xml']);
  expect(await chartsOf(slides[1].part)).toEqual(['ppt/charts/chartEx2.xml']);
  // The original chart keeps its workbook
  expect(
    relationships(
      await zip.file('ppt/charts/_rels/chartEx1.xml.rels').async('string'),
    ).find((r) => r.type === 'package').target,
  ).toBe('../embeddings/Microsoft_Excel_Worksheet.xlsx');
  expect(duplicateContentTypes).toEqual([]);
});
