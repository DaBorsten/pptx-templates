import path from 'path';
import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { listTemplateCommands, renderDeck } from '../src/index';
import { deck, lint } from './helpers/lint-deck';
import { opaquePng, pngColorType, Rgb } from './helpers/png';

/**
 * `{{IMAGE …}}` in the alt text of a picture placeholder: the picture
 * replaces the backup picture there, cropped (a:srcRect) to cover the
 * placeholder, which keeps its position, shape and link to the layout.
 */

const png = fs.readFileSync(`${__dirname}/media/test.png`); // 200 × 129
const feather = fs.readFileSync(`${__dirname}/media/feather.png`); // 190 × 191

// Picture placeholder of slideLayout9 in EmptySlide.pptx
const FRAME = { cx: 6172200, cy: 4873625 };
const LAYOUT = 'ppt/slideLayouts/slideLayout9.xml';
const SLIDE = 'ppt/slides/slide1.xml';
const RELS = 'ppt/slides/_rels/slide1.xml.rels';
const REL_IMAGE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** Zip path of a relationship target of a slide (`../media/x.png`). */
const slidePart = (target: string) => path.posix.join('ppt/slides', target);

const cNvPr = (descr: string) =>
  `<p:cNvPr id="10" name="Bildplatzhalter 1" descr="${descr
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')}"/>`;

/** A picture in the placeholder (the backup), as PowerPoint saves it. */
const backupPicture = (descr: string) =>
  `<p:pic><p:nvPicPr>${cNvPr(descr)}` +
  '<p:cNvPicPr><a:picLocks noGrp="1" noChangeAspect="1"/></p:cNvPicPr>' +
  '<p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rIdBackup"/><a:srcRect t="5000" b="5000"/>' +
  '<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr/></p:pic>';

/** The placeholder without a picture. */
const emptyPlaceholder = (descr: string) =>
  `<p:sp><p:nvSpPr>${cNvPr(descr)}` +
  '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
  '<p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>';

/** A plain circle with text. */
const circle = (descr: string) =>
  `<p:sp><p:nvSpPr>${cNvPr(descr)}<p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>' +
  '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Name</a:t></a:r></a:p></p:txBody></p:sp>';

/**
 * EmptySlide.pptx with `shape` on slide 1, which uses slideLayout9; its
 * picture placeholder (`type="pic"`, PowerPoint's "Picture (fill)") is a
 * circle there, of `size`. The backup picture is feather.png.
 */
const deckWith = async (
  shape: string,
  size: { cx: number; cy: number } = FRAME,
): Promise<Buffer> => {
  const zip = await JSZip.loadAsync(
    fs.readFileSync(`${__dirname}/pptx-templates/EmptySlide.pptx`),
  );
  const layout = await zip.file(LAYOUT).async('text');
  const placeholderFrame = `<a:xfrm><a:off x="5183188" y="987425"/><a:ext cx="${FRAME.cx}" cy="${FRAME.cy}"/></a:xfrm>`;
  expect(layout).toContain(placeholderFrame);
  zip.file(
    LAYOUT,
    layout.replace(
      placeholderFrame,
      placeholderFrame.replace(
        /<a:ext [^>]*\/>/,
        `<a:ext cx="${size.cx}" cy="${size.cy}"/>`,
      ) + '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>',
    ),
  );

  const slide = await zip.file(SLIDE).async('text');
  zip.file(SLIDE, slide.replace('</p:spTree>', `${shape}</p:spTree>`));
  zip.file(
    RELS,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout9.xml"/>' +
      `<Relationship Id="rIdBackup" Type="${REL_IMAGE}" Target="../media/backup.png"/>` +
      '</Relationships>',
  );
  zip.file('ppt/media/backup.png', feather);
  const types = await zip.file('[Content_Types].xml').async('text');
  if (!types.includes('Extension="png"')) {
    zip.file(
      '[Content_Types].xml',
      types.replace(
        '<Default ',
        '<Default Extension="png" ContentType="image/png"/><Default ',
      ),
    );
  }
  return zip.generateAsync({ type: 'nodebuffer' });
};

type Rendered = {
  slide: Document;
  rels: Element[];
  media: (target: string) => Promise<Buffer>;
};

const render = async (
  shape: string,
  data: Record<string, unknown>,
  options?: Parameters<typeof renderDeck>[2],
): Promise<Rendered> => {
  const zip = await JSZip.loadAsync(
    await renderDeck(await deckWith(shape), data, options),
  );
  const slidePath = Object.keys(zip.files).find((f) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(f),
  );
  const relsPath = slidePath.replace('slides/', 'slides/_rels/') + '.rels';
  const parse = async (file: string) =>
    new DOMParser().parseFromString(
      await zip.file(file).async('text'),
      'application/xml',
    ) as unknown as Document;
  const rels = Array.from(
    (await parse(relsPath)).getElementsByTagName('Relationship'),
  );
  return {
    slide: await parse(slidePath),
    rels,
    media: (target) => zip.file(slidePart(target)).async('nodebuffer'),
  };
};

const first = (doc: Document | Element, tag: string) =>
  doc.getElementsByTagName(tag)[0];
const directChildren = (el: Element) =>
  Array.from(el.childNodes)
    .filter((n) => n.nodeType === 1)
    .map((n) => (n as Element).tagName);

/** srcRect attributes that crop `ratio` evenly to the frame's ratio. */
const cover = (ratio: number) => {
  const frame = FRAME.cx / FRAME.cy;
  return ratio > frame
    ? { l: Math.round(((1 - frame / ratio) / 2) * 100000) }
    : { t: Math.round(((1 - ratio / frame) / 2) * 100000) };
};

test('replaces the backup picture, cropped to cover the inherited frame', async () => {
  const { slide, rels, media } = await render(backupPicture('{{IMAGE foto}}'), {
    foto: png,
  });

  const pic = first(slide, 'p:pic');
  // Still the picture in the placeholder: PowerPoint can replace it by hand
  expect(first(pic, 'p:ph').getAttribute('idx')).toBe('1');
  expect(first(pic, 'p:cNvPr').hasAttribute('descr')).toBe(false);
  // Position and circle stay inherited from the layout
  expect(directChildren(first(pic, 'p:spPr'))).toEqual([]);

  // Wider than the frame: cropped left and right, evenly
  const srcRect = first(pic, 'a:srcRect');
  const { l } = cover(200 / 129);
  expect(l).toBeGreaterThan(0);
  expect(srcRect.getAttribute('l')).toBe(String(l));
  expect(srcRect.getAttribute('r')).toBe(String(l));
  expect(srcRect.hasAttribute('t')).toBe(false);
  expect(directChildren(first(pic, 'p:blipFill'))).toEqual([
    'a:blip',
    'a:srcRect',
    'a:stretch',
  ]);

  const rel = rels.find(
    (r) =>
      r.getAttribute('Id') === first(pic, 'a:blip').getAttribute('r:embed'),
  );
  expect((await media(rel.getAttribute('Target'))).equals(png)).toBe(true);
  // The relationship of the backup picture is gone
  expect(rels.filter((r) => r.getAttribute('Type') === REL_IMAGE)).toHaveLength(
    1,
  );
});

test('a picture taller than the frame is cropped at top and bottom', async () => {
  const { slide } = await render(backupPicture('{{IMAGE foto}}'), {
    foto: feather,
  });
  const srcRect = first(slide, 'a:srcRect');
  const { t } = cover(190 / 191);
  expect(srcRect.getAttribute('t')).toBe(String(t));
  expect(srcRect.getAttribute('b')).toBe(String(t));
  expect(srcRect.hasAttribute('l')).toBe(false);
});

test('without a value the backup picture stays', async () => {
  for (const foto of [null, undefined, '']) {
    const { slide, rels } = await render(backupPicture('{{IMAGE foto}}'), {
      foto,
    });
    const pic = first(slide, 'p:pic');
    expect(first(pic, 'a:blip').getAttribute('r:embed')).toBe('rIdBackup');
    expect(first(pic, 'a:srcRect').getAttribute('t')).toBe('5000');
    expect(first(pic, 'p:cNvPr').hasAttribute('descr')).toBe(false);
    expect(rels.some((r) => r.getAttribute('Id') === 'rIdBackup')).toBe(true);
  }
});

test('alt sets the alt text of the picture', async () => {
  const { slide } = await render(
    backupPicture("{{IMAGE ({ data: foto, alt: 'Ada Lovelace' })}}"),
    { foto: png },
  );
  expect(first(first(slide, 'p:pic'), 'p:cNvPr').getAttribute('descr')).toBe(
    'Ada Lovelace',
  );
});

test('an empty placeholder becomes a picture in that placeholder', async () => {
  const { slide } = await render(emptyPlaceholder('{{IMAGE foto}}'), {
    foto: png,
  });
  expect(slide.getElementsByTagName('p:sp')).toHaveLength(1); // the text box
  const pic = first(slide, 'p:pic');
  expect(directChildren(pic)).toEqual(['p:nvPicPr', 'p:blipFill', 'p:spPr']);
  expect(first(pic, 'p:ph').getAttribute('type')).toBe('pic');
  expect(first(pic, 'a:picLocks').getAttribute('noGrp')).toBe('1');
  expect(first(pic, 'p:cNvPr').getAttribute('name')).toBe('Bildplatzhalter 1');
  expect(first(pic, 'a:srcRect').getAttribute('l')).toBe(
    String(cover(200 / 129).l),
  );
});

test('a shape that is no placeholder gets the picture as its fill', async () => {
  const { slide } = await render(circle('{{IMAGE foto}}'), { foto: feather });
  const sp = Array.from(slide.getElementsByTagName('p:sp')).find(
    (s) => first(s, 'p:cNvPr').getAttribute('name') === 'Bildplatzhalter 1',
  );
  const spPr = first(sp, 'p:spPr');
  // the fill replaces the red one, in schema order
  expect(directChildren(spPr)).toEqual(['a:xfrm', 'a:prstGeom', 'a:blipFill']);
  // 2:1 frame, square picture: the fill rectangle reaches past the top
  // and bottom (a:fillRect, which LibreOffice reads; a:srcRect stays empty)
  const srcRect = first(spPr, 'a:srcRect');
  expect(srcRect.attributes).toHaveLength(0);
  const fillRect = first(spPr, 'a:fillRect');
  const side = -((2 / (190 / 191) - 1) / 2) * 100000;
  expect(Number(fillRect.getAttribute('t'))).toBeCloseTo(side, -1);
  expect(Number(fillRect.getAttribute('b'))).toBeCloseTo(side, -1);
  expect(fillRect.hasAttribute('l')).toBe(false);
  expect(first(sp, 'a:t').textContent).toBe('Name');
});

test('an alt text that is not only tags with an IMAGE stays as it is', async () => {
  for (const descr of [
    'Preis {{ in Euro',
    'Foto: {{IMAGE foto}}',
    '{{name}}',
    '{{IF a}}{{END-IF}}',
  ]) {
    const { slide } = await render(backupPicture(descr), {
      foto: png,
      name: 'x',
      a: true,
    });
    const pic = first(slide, 'p:pic');
    expect(first(pic, 'p:cNvPr').getAttribute('descr')).toBe(descr);
    expect(first(pic, 'a:blip').getAttribute('r:embed')).toBe('rIdBackup');
  }
});

test('IMAGE in the alt text uses the configured delimiters', async () => {
  const options = { cmdDelimiter: ['[[', ']]'] as [string, string] };
  const filled = await render(
    backupPicture('[[IMAGE foto]]'),
    { foto: png },
    options,
  );
  const pic = first(filled.slide, 'p:pic');
  expect(first(pic, 'a:blip').getAttribute('r:embed')).not.toBe('rIdBackup');
  expect(first(pic, 'p:cNvPr').hasAttribute('descr')).toBe(false);
  // {{…}} is plain text with other delimiters
  const kept = await render(
    backupPicture('{{IMAGE foto}}'),
    { foto: png },
    options,
  );
  expect(
    first(first(kept.slide, 'p:pic'), 'p:cNvPr').getAttribute('descr'),
  ).toBe('{{IMAGE foto}}');
});

test('rejects other tags or size options in the alt text', async () => {
  await expect(
    render(backupPicture('{{IMAGE foto}}{{name}}'), { foto: png, name: 'x' }),
  ).rejects.toThrow(/only IMAGE, IF, ELSE-IF, ELSE, FOR and EXEC/);
  await expect(
    render(backupPicture('{{IMAGE ({ data: foto, width: 3 })}}'), {
      foto: png,
    }),
  ).rejects.toThrow(/"width" is not supported/);
});

test('an alt text that does not parse leaves no tag in the output', async () => {
  for (const descr of ['{{IMAGE foto}}{{name}}', '{{IF a}}{{IMAGE foto}}']) {
    const errors: string[] = [];
    const zip = await JSZip.loadAsync(
      await renderDeck(
        await deckWith(backupPicture(descr)),
        { foto: png, name: 'x', a: true },
        {
          errorHandler: (error) => {
            errors.push(error.message);
            return undefined;
          },
        },
      ),
    );
    expect(errors).toHaveLength(1);
    const slidePath = Object.keys(zip.files).find((f) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(f),
    );
    const slide = await zip.file(slidePath).async('text');
    expect(slide).not.toContain('descr=');
    // the backup picture stays
    expect(slide).toContain('r:embed="rIdBackup"');
  }
});

/** The relationship id of the picture in the first p:pic. */
const embedded = (r: Rendered) =>
  first(first(r.slide, 'p:pic'), 'a:blip').getAttribute('r:embed');
const pictureOf = (r: Rendered) =>
  r.media(
    r.rels
      .find((rel) => rel.getAttribute('Id') === embedded(r))
      .getAttribute('Target'),
  );

test('IF, ELSE-IF and ELSE pick the picture', async () => {
  const alt =
    '{{IF mode == "a"}}{{IMAGE a}}{{ELSE-IF mode == "b"}}{{IMAGE b}}' +
    '{{ELSE}}{{IMAGE null}}{{END-IF}}';
  const a = await render(backupPicture(alt), { mode: 'a', a: png, b: feather });
  expect((await pictureOf(a)).equals(png)).toBe(true);
  const b = await render(backupPicture(alt), { mode: 'b', a: png, b: feather });
  expect((await pictureOf(b)).equals(feather)).toBe(true);
  // ELSE yields no picture: the backup stays, the alt text goes
  const none = await render(backupPicture(alt), { mode: 'c', a: png });
  expect(embedded(none)).toBe('rIdBackup');
  expect(first(none.slide, 'p:cNvPr').hasAttribute('descr')).toBe(false);
});

test('an IF that is false keeps the backup picture', async () => {
  const r = await render(backupPicture('{{IF show}}{{IMAGE foto}}{{END-IF}}'), {
    show: false,
    foto: png,
  });
  expect(embedded(r)).toBe('rIdBackup');
});

test('the ternary operator picks the picture', async () => {
  const alt = '{{IMAGE female ? f : m}}';
  const f = await render(backupPicture(alt), {
    female: true,
    f: png,
    m: feather,
  });
  expect((await pictureOf(f)).equals(png)).toBe(true);
  const m = await render(backupPicture(alt), {
    female: false,
    f: png,
    m: feather,
  });
  expect((await pictureOf(m)).equals(feather)).toBe(true);
});

test('FOR and EXEC work in the alt text', async () => {
  const r = await render(
    backupPicture(
      '{{EXEC p = people[1]}}{{FOR x IN [p]}}{{IMAGE x.foto}}{{END-FOR}}',
    ),
    { people: [{ foto: feather }, { foto: png }] },
  );
  expect((await pictureOf(r)).equals(png)).toBe(true);
});

test('more than one picture from one alt text is an error', async () => {
  await expect(
    render(backupPicture('{{FOR f IN fotos}}{{IMAGE f}}{{END-FOR}}'), {
      fotos: [png, feather],
    }),
  ).rejects.toThrow(/yields 2 pictures/);
});

test('a placeholder with text keeps its text and gets the picture as its fill', async () => {
  const body =
    '<p:sp><p:nvSpPr><p:cNvPr id="10" name="Inhalt" descr="{{IMAGE foto}}"/>' +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
    '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Bleibt</a:t></a:r></a:p></p:txBody></p:sp>';
  const { slide } = await render(body, { foto: png });
  expect(slide.getElementsByTagName('p:pic')).toHaveLength(0);
  const sp = Array.from(slide.getElementsByTagName('p:sp')).find(
    (s) => first(s, 'p:cNvPr').getAttribute('name') === 'Inhalt',
  );
  expect(first(sp, 'a:t').textContent).toBe('Bleibt');
  expect(first(sp, 'p:ph').getAttribute('idx')).toBe('1');
  expect(directChildren(first(sp, 'p:spPr'))).toEqual(['a:blipFill']);
});

test('a title placeholder is filled, not turned into a picture', async () => {
  const title =
    '<p:sp><p:nvSpPr><p:cNvPr id="10" name="Titel" descr="{{IMAGE foto}}"/>' +
    '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
    '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>';
  const { slide } = await render(title, { foto: png });
  expect(slide.getElementsByTagName('p:pic')).toHaveLength(0);
  const sp = Array.from(slide.getElementsByTagName('p:sp')).find(
    (s) => first(s, 'p:cNvPr').getAttribute('name') === 'Titel',
  );
  expect(first(sp, 'p:ph').getAttribute('type')).toBe('title');
  expect(directChildren(first(sp, 'p:spPr'))).toEqual(['a:xfrm', 'a:blipFill']);
});

const tableWithAlt = (descr: string) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10" name="Tabelle" descr="${descr}"/>` +
  '<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></p:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
  '<a:tbl><a:tblGrid><a:gridCol w="1000"/></a:tblGrid><a:tr h="1000"><a:tc>' +
  '<a:txBody><a:bodyPr/><a:p><a:endParaRPr lang="de-DE"/></a:p></a:txBody></a:tc></a:tr></a:tbl>' +
  '</a:graphicData></a:graphic></p:graphicFrame>';

test('an unsupported element is an error and adds no media file', async () => {
  const frame = tableWithAlt('{{IMAGE foto}}');
  await expect(render(frame, { foto: png })).rejects.toThrow(
    /not on <p:graphicFrame>/,
  );
  // errorHandler: the render goes on, without a picture in the deck
  const errors: string[] = [];
  const zip = await JSZip.loadAsync(
    await renderDeck(
      await deckWith(frame),
      { foto: png },
      {
        errorHandler: (error) => {
          errors.push(error.message);
          return undefined;
        },
      },
    ),
  );
  expect(errors).toEqual([expect.stringMatching(/not on <p:graphicFrame>/)]);
  const media = Object.keys(zip.files).filter(
    (f) => f.startsWith('ppt/media/') && !zip.files[f].dir,
  );
  const stored = await Promise.all(
    media.map((f) => zip.file(f).async('nodebuffer')),
  );
  expect(stored.some((file) => file.equals(png))).toBe(false);
});

/** A card: a group with a flow() text box and a picture. */
const card = (n: number, x: number) => {
  const loop = '{{FOR p IN flow(people, 1, { measure: false })}}';
  const id = 20 + n * 3;
  return (
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${id}" name="Karte ${n}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="2000000" cy="3000000"/>` +
    `<a:chOff x="${x}" y="0"/><a:chExt cx="2000000" cy="3000000"/></a:xfrm></p:grpSpPr>` +
    `<p:sp><p:nvSpPr><p:cNvPr id="${id + 1}" name="Name ${n}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${x}" y="2000000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
    `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${loop}</a:t></a:r></a:p>` +
    '<a:p><a:r><a:t>{{p.name}}</a:t></a:r></a:p>' +
    '<a:p><a:r><a:t>{{END-FOR}}</a:t></a:r></a:p></p:txBody></p:sp>' +
    `<p:pic><p:nvPicPr><p:cNvPr id="${id + 2}" name="Foto ${n}" descr="${loop}{{IMAGE p.foto}}{{END-FOR}}"/>` +
    '<p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdBackup"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    `<p:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="2000000" cy="2000000"/></a:xfrm>` +
    '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:grpSp>'
  );
};

test('in a flow() card, flow() in the alt text gives the entry of that card', async () => {
  const third = opaquePng(4, 2, () => [0, 130, 200]);
  const people = [
    { name: 'Ada', foto: png },
    { name: 'Grace', foto: feather },
    { name: 'Hedy', foto: third },
  ];
  const pptx = await renderDeck(await deckWith(card(1, 0) + card(2, 3000000)), {
    people,
  });
  const zip = await JSZip.loadAsync(pptx);
  const parse = async (file: string) =>
    new DOMParser().parseFromString(
      await zip.file(file).async('text'),
      'application/xml',
    ) as unknown as Document;
  const presentation = await parse('ppt/presentation.xml');
  const presRels = Array.from(
    (await parse('ppt/_rels/presentation.xml.rels')).getElementsByTagName(
      'Relationship',
    ),
  );
  const slides = Array.from(presentation.getElementsByTagName('p:sldId')).map(
    (sldId) =>
      'ppt/' +
      presRels
        .find((r) => r.getAttribute('Id') === sldId.getAttribute('r:id'))
        .getAttribute('Target'),
  );
  expect(slides).toHaveLength(2);

  /** Name, picture and alt text of each card on a slide. */
  const cards = async (file: string) => {
    const doc = await parse(file);
    const rels = Array.from(
      (
        await parse(file.replace('slides/', 'slides/_rels/') + '.rels')
      ).getElementsByTagName('Relationship'),
    );
    return Promise.all(
      Array.from(doc.getElementsByTagName('p:grpSp')).map(async (group) => {
        const pic = first(group, 'p:pic');
        const id = first(pic, 'a:blip').getAttribute('r:embed');
        const target = rels
          .find((r) => r.getAttribute('Id') === id)
          .getAttribute('Target');
        return {
          name: Array.from(group.getElementsByTagName('a:t'))
            .map((t) => t.textContent)
            .join(''),
          foto: await zip.file(slidePart(target)).async('nodebuffer'),
          descr: first(pic, 'p:cNvPr').hasAttribute('descr'),
        };
      }),
    );
  };

  const page1 = await cards(slides[0]);
  expect(page1.map((c) => c.name)).toEqual(['Ada', 'Grace']);
  expect(page1[0].foto.equals(png)).toBe(true);
  expect(page1[1].foto.equals(feather)).toBe(true);
  expect(page1.some((c) => c.descr)).toBe(false);
  // the second card of the last slide stays empty and is removed
  const page2 = await cards(slides[1]);
  expect(page2.map((c) => c.name)).toEqual(['Hedy']);
  expect(page2[0].foto.equals(third)).toBe(true);
});

test('flow() in the alt text of a shape outside a card is an error', async () => {
  await expect(
    render(
      card(1, 0) +
        backupPicture(
          '{{FOR p IN flow(people, 1, { measure: false })}}{{IMAGE p.foto}}{{END-FOR}}',
        ),
      { people: [{ name: 'Ada', foto: png }] },
    ),
  ).rejects.toThrow(/works only on a shape in a card/);
});

/** A text box with `body` as its text. */
const textBox = (body: string) =>
  '<p:sp><p:nvSpPr><p:cNvPr id="20" name="Box"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
  `<p:txBody><a:bodyPr/><a:p><a:r><a:t>${body}</a:t></a:r></a:p></p:txBody></p:sp>`;

test('alt text from the data is content, not a template', async () => {
  for (const caption of ['{{IMAGE other}}', 'Preis {{ in Euro']) {
    const { slide, rels, media } = await render(
      textBox('{{IMAGE ({ data: foto, alt: caption })}}'),
      { foto: png, other: feather, caption },
    );
    const pic = first(slide, 'p:pic');
    expect(first(pic, 'p:cNvPr').getAttribute('descr')).toBe(caption);
    const embed = first(pic, 'a:blip').getAttribute('r:embed');
    const target = rels
      .find((r) => r.getAttribute('Id') === embed)
      .getAttribute('Target');
    expect((await media(target)).equals(png)).toBe(true);
  }
});

/** lint-deck parts declare no r: namespace */
const lintPicture = (descr: string) =>
  backupPicture(descr).replace(' r:embed="rIdBackup"', '');

test('listTemplateCommands lists the tag in the alt text', async () => {
  const commands = await listTemplateCommands(
    await deck([lintPicture('{{IMAGE foto}}')]),
  );
  expect(commands).toEqual([
    {
      slide: 1,
      shape: 'Bildplatzhalter 1',
      raw: '{{IMAGE foto}}',
      type: 'IMAGE',
      code: 'foto',
    },
  ]);
});

test('lint accepts IMAGE in the alt text, also in IF and FOR, nothing else', async () => {
  expect(await lint([lintPicture('{{IMAGE foto}}')])).toEqual([]);
  expect(await lint([emptyPlaceholder('{{IMAGE foto}}')])).toEqual([]);
  expect(
    await lint([
      lintPicture('{{IF a}}{{IMAGE a}}{{ELSE}}{{IMAGE b ? b : c}}{{END-IF}}'),
    ]),
  ).toEqual([]);
  expect(
    await lint([lintPicture('{{FOR x IN list}}{{IMAGE x.foto}}{{END-FOR}}')]),
  ).toEqual([]);

  const [unclosed] = await lint([lintPicture('{{IF a}}{{IMAGE a}}')]);
  expect(unclosed).toMatchObject({ severity: 'error', tag: '{{IF a}}' });
  const [table] = await lint([
    tableWithAlt('{{IMAGE foto}}').replace(/ uri="[^"]*"/, ''),
  ]);
  expect(table.message).toMatch(/not on <p:graphicFrame>/);

  const [other] = await lint([lintPicture('{{IMAGE foto}}{{name}}')]);
  expect(other).toMatchObject({
    severity: 'error',
    rule: 'unsupported-location',
    shape: 'Bildplatzhalter 1',
    tag: '{{name}}',
  });

  // Not a template: stays as it is, a warning for the tag in it
  for (const descr of ['Foto {{IMAGE foto}}', '{{name}}']) {
    const issues = await lint([lintPicture(descr)]);
    expect(issues).toEqual([
      expect.objectContaining({
        severity: 'warning',
        rule: 'unsupported-location',
        shape: 'Bildplatzhalter 1',
        message: expect.stringMatching(/stays as it is/),
      }),
    ]);
  }
  expect(await lint([lintPicture('Preis {{ in Euro')])).toEqual([]);
});

test('lint rejects size and rotation for IMAGE in the alt text', async () => {
  const issues = await lint([
    lintPicture(
      '{{IMAGE a ? ({ data: a, width: 3, rotation: 90 }) : ({ data: b, width: 2, alt: "B" })}}',
    ),
  ]);
  expect(issues.map((issue) => issue.message)).toEqual([
    expect.stringMatching(/"width" is not supported there/),
    expect.stringMatching(/"rotation" is not supported there/),
  ]);
  expect(issues[0]).toMatchObject({
    severity: 'error',
    rule: 'unsupported-location',
  });
  // the other options are fine there, and in a text box size is too
  expect(
    await lint([
      lintPicture('{{IMAGE ({ data: a, alt: "A", shape: "ellipse" })}}'),
    ]),
  ).toEqual([]);
});

test('circle placeholder (Picture fill) with an opaque landscape picture', async () => {
  // 2:1, every pixel opaque: colored stripes and a round dot in the
  // middle, which stays round only if the picture is not distorted
  const stripes: Rgb[] = [
    [230, 25, 75],
    [245, 130, 48],
    [255, 225, 25],
    [60, 180, 75],
    [70, 240, 240],
    [0, 130, 200],
    [145, 30, 180],
    [240, 50, 230],
  ];
  const [width, height] = [800, 400];
  const photo = opaquePng(width, height, (x, y) =>
    (x - width / 2) ** 2 + (y - height / 2) ** 2 < (height * 0.3) ** 2
      ? [20, 20, 20]
      : stripes[Math.floor((x / width) * stripes.length)],
  );
  expect(pngColorType(photo)).toBe(2);

  const square = { cx: 4800000, cy: 4800000 };
  const pptx = await renderDeck(
    await deckWith(backupPicture('{{IMAGE foto}}'), square),
    { foto: photo },
  );
  // Open in PowerPoint: a circle filled with the middle of the picture,
  // the dark dot round and centered, no empty edge
  fs.mkdirSync(`${__dirname}/pptx-output`, { recursive: true });
  fs.writeFileSync(
    `${__dirname}/pptx-output/template-image-alt-text-circle.pptx`,
    pptx,
  );

  const zip = await JSZip.loadAsync(pptx);
  const slidePath = Object.keys(zip.files).find((f) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(f),
  );
  const slide = new DOMParser().parseFromString(
    await zip.file(slidePath).async('text'),
    'application/xml',
  ) as unknown as Document;
  const pic = first(slide, 'p:pic');
  expect(first(pic, 'p:ph').getAttribute('type')).toBe('pic');
  // circle and size come from the layout
  expect(directChildren(first(pic, 'p:spPr'))).toEqual([]);
  expect(await zip.file(LAYOUT).async('text')).toContain(
    `<a:ext cx="4800000" cy="4800000"/></a:xfrm><a:prstGeom prst="ellipse">`,
  );
  // 2:1 into 1:1: a quarter off on the left and on the right
  const srcRect = first(pic, 'a:srcRect');
  expect(srcRect.getAttribute('l')).toBe('25000');
  expect(srcRect.getAttribute('r')).toBe('25000');
  expect(srcRect.hasAttribute('t')).toBe(false);

  const media = Object.keys(zip.files).filter(
    (f) => f.startsWith('ppt/media/') && !zip.files[f].dir,
  );
  const stored = await Promise.all(
    media.map((f) => zip.file(f).async('nodebuffer')),
  );
  expect(stored.some((file) => file.equals(photo))).toBe(true);
  // the backup picture is gone
  expect(stored.some((file) => file.equals(feather))).toBe(false);
});

test('lint ignores alt text without an opening delimiter, like the renderer', async () => {
  for (const descr of ['Text }}', 'Pfeil -> }} Ende', 'Foto']) {
    expect(await lint([lintPicture(descr)])).toEqual([]);
    const { slide } = await render(backupPicture(descr), {});
    const pic = first(slide, 'p:pic');
    expect(first(pic, 'p:cNvPr').getAttribute('descr')).toBe(descr);
    expect(first(pic, 'a:blip').getAttribute('r:embed')).toBe('rIdBackup');
  }
});

test('the crop covers the shape as shown, in a group scaled unevenly', async () => {
  // A square picture in a group stretched to twice its width: on the
  // slide the frame is 2:1
  const group =
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="30" name="Gruppe"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/>' +
    '<a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="2000000"/></a:xfrm></p:grpSpPr>' +
    '<p:pic><p:nvPicPr><p:cNvPr id="31" name="Foto" descr="{{IMAGE foto}}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    '<p:blipFill><a:blip r:embed="rIdBackup"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="2000000"/></a:xfrm>' +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:grpSp>';
  const { slide } = await render(group, {
    foto: opaquePng(4, 4, () => [0, 130, 200]),
  });
  const srcRect = first(first(slide, 'p:pic'), 'a:srcRect');
  // 1:1 into 2:1: a quarter off at the top and at the bottom
  expect(srcRect.getAttribute('t')).toBe('25000');
  expect(srcRect.getAttribute('b')).toBe('25000');
  expect(srcRect.hasAttribute('l')).toBe(false);
});

/*
 * Example for opening in PowerPoint: the template with the tags in the alt
 * texts (….input.pptx) and what renderDeck makes of it (….output.pptx).
 */

/** Stripes with a dark dot in the middle: the dot stays round only undistorted. */
const dotPicture = (width: number, height: number, stripes: Rgb[]) =>
  opaquePng(width, height, (x, y) =>
    (x - width / 2) ** 2 + (y - height / 2) ** 2 <
    (Math.min(width, height) * 0.3) ** 2
      ? [20, 20, 20]
      : stripes[Math.floor((x / width) * stripes.length)],
  );

const RAINBOW: Rgb[] = [
  [230, 25, 75],
  [245, 130, 48],
  [255, 225, 25],
  [60, 180, 75],
  [70, 240, 240],
  [0, 130, 200],
  [145, 30, 180],
  [240, 50, 230],
];

const label = (id: number, x: number, y: number, text: string, cx = 4800000) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Beschriftung ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="600000"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
  `<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="de-DE" sz="1100"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;

const picture = (
  id: number,
  name: string,
  descr: string,
  frame: { x: number; y: number; cx: number; cy: number },
  geometry = 'rect',
) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}" descr="${descr}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  '<p:blipFill><a:blip r:embed="rIdBackup"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  `<p:spPr><a:xfrm><a:off x="${frame.x}" y="${frame.y}"/><a:ext cx="${frame.cx}" cy="${frame.cy}"/></a:xfrm>` +
  `<a:prstGeom prst="${geometry}"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;

test('example deck: template and result for PowerPoint', async () => {
  const shapes = [
    // 1. The empty circle picture placeholder of the layout, as PowerPoint
    //    saves it: "Bild durch Klicken hinzufügen", tag in the alt text
    '<p:sp><p:nvSpPr><p:cNvPr id="10" name="Bildplatzhalter 1" descr="{{IMAGE quer}}"/>' +
      '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' +
      '<p:nvPr><p:ph type="pic" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="de-DE"/></a:p></p:txBody></p:sp>',
    label(
      41,
      5183188,
      5850000,
      '1: leerer Bildplatzhalter (Kreis), Querformat, Punkt rund',
    ),
    // 6. A second picture placeholder (portrait, right) with a backup
    //    picture in it: the picture replaces the backup
    '<p:pic><p:nvPicPr><p:cNvPr id="51" name="Bildplatzhalter mit Ersatzbild" descr="{{IMAGE quer}}"/>' +
      '<p:cNvPicPr><a:picLocks noGrp="1" noChangeAspect="1"/></p:cNvPicPr>' +
      '<p:nvPr><p:ph type="pic" idx="13"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rIdBackup"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr/></p:pic>',
    label(
      52,
      10250000,
      3450000,
      '6: Bildplatzhalter mit Ersatzbild (Feder), Querformat',
      1800000,
    ),
    // 2. A rounded rectangle with a red fill and text: the picture becomes its fill
    '<p:sp><p:nvSpPr><p:cNvPr id="42" name="Form mit Text" descr="{{IMAGE hoch}}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="400000" y="400000"/><a:ext cx="2800000" cy="1600000"/></a:xfrm>' +
      '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>' +
      '<p:txBody><a:bodyPr anchor="b"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="de-DE" sz="2000" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Text bleibt</a:t></a:r></a:p></p:txBody></p:sp>',
    label(43, 400000, 2020000, '2: Form, Hochformat als Füllung, Text bleibt'),
    // 3. A picture in a group stretched to twice its width
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="44" name="Gestreckte Gruppe"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="400000" y="2500000"/><a:ext cx="2800000" cy="1400000"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="1400000" cy="1400000"/></a:xfrm></p:grpSpPr>' +
      picture(45, 'Bild in Gruppe', '{{IMAGE quadrat}}', {
        x: 0,
        y: 0,
        cx: 1400000,
        cy: 1400000,
      }) +
      '</p:grpSp>',
    label(46, 400000, 3920000, '3: gestreckte Gruppe, quadratisch, Punkt rund'),
    // 4. IF that is false: the backup picture (feather) stays
    picture(47, 'Bild mit IF', '{{IF zeigen}}{{IMAGE quer}}{{END-IF}}', {
      x: 400000,
      y: 4400000,
      cx: 1300000,
      cy: 1300000,
    }),
    label(48, 400000, 5750000, '4: IF falsch, Feder bleibt'),
    // 5. Ternary operator, in an ellipse
    picture(
      49,
      'Bild mit Ternär',
      '{{IMAGE blau ? blauBild : gruenBild}}',
      { x: 2000000, y: 4400000, cx: 1300000, cy: 1300000 },
      'ellipse',
    ),
    label(50, 2000000, 6050000, '5: Ternär, blauer Kreis'),
  ].join('');

  const data = {
    quer: dotPicture(800, 400, RAINBOW),
    hoch: dotPicture(300, 600, RAINBOW),
    quadrat: dotPicture(400, 400, RAINBOW),
    zeigen: false,
    blau: true,
    blauBild: dotPicture(200, 200, [[0, 90, 200]]),
    gruenBild: dotPicture(200, 200, [[40, 160, 60]]),
  };

  // A second picture placeholder on the layout: a copy of the first,
  // portrait, in the right column
  const base = await JSZip.loadAsync(
    await deckWith(shapes, { cx: 4800000, cy: 4800000 }),
  );
  const layout = await base.file(LAYOUT).async('text');
  const original = layout.match(
    /<p:sp><p:nvSpPr><p:cNvPr id="3" name="Bildplatzhalter 2">.*?<\/p:sp>/,
  )[0];
  const second = original
    .replace(
      /<p:cNvPr id="3" name="Bildplatzhalter 2">.*?<\/p:cNvPr>/,
      '<p:cNvPr id="13" name="Bildplatzhalter 13"/>',
    )
    .replace('idx="1"', 'idx="13"')
    .replace(
      /<a:xfrm>.*?<\/a:xfrm><a:prstGeom prst="ellipse"><a:avLst\/><\/a:prstGeom>/,
      '<a:xfrm><a:off x="10250000" y="987425"/><a:ext cx="1700000" cy="2400000"/></a:xfrm>',
    );
  expect(second).toContain('<p:ph type="pic" idx="13"/>');
  expect(second).toContain('cx="1700000"');
  base.file(LAYOUT, layout.replace(original, original + second));
  const input = await base.generateAsync({ type: 'nodebuffer' });
  const output = await renderDeck(input, data);
  fs.mkdirSync(`${__dirname}/pptx-output`, { recursive: true });
  fs.writeFileSync(
    `${__dirname}/pptx-output/template-image-alt-text-example.input.pptx`,
    input,
  );
  fs.writeFileSync(
    `${__dirname}/pptx-output/template-image-alt-text-example.output.pptx`,
    output,
  );

  const zip = await JSZip.loadAsync(output);
  const slidePath = Object.keys(zip.files).find((f) =>
    /^ppt\/slides\/slide\d+\.xml$/.test(f),
  );
  const parse = async (file: string) =>
    new DOMParser().parseFromString(
      await zip.file(file).async('text'),
      'application/xml',
    ) as unknown as Document;
  const slide = await parse(slidePath);
  const rels = Array.from(
    (
      await parse(slidePath.replace('slides/', 'slides/_rels/') + '.rels')
    ).getElementsByTagName('Relationship'),
  );
  const named = (name: string) =>
    Array.from(slide.getElementsByTagName('p:cNvPr')).find(
      (c) => c.getAttribute('name') === name,
    ).parentNode.parentNode as Element;
  const pictureIn = (shape: Element) => {
    const id = first(shape, 'a:blip').getAttribute('r:embed');
    const target = rels
      .find((r) => r.getAttribute('Id') === id)
      .getAttribute('Target');
    return zip.file(slidePart(target)).async('nodebuffer');
  };
  const crop = (shape: Element) => {
    const rect = first(shape, 'a:srcRect');
    return ['l', 't', 'r', 'b'].map((a) => rect.getAttribute(a) || '0');
  };

  // No tag left in any alt text
  expect(await zip.file(slidePath).async('text')).not.toContain('{{');

  // The empty placeholder is now a picture in that placeholder
  const placeholder = named('Bildplatzhalter 1');
  expect(placeholder.tagName).toBe('p:pic');
  expect(first(placeholder, 'p:ph').getAttribute('idx')).toBe('1');
  expect((await pictureIn(placeholder)).equals(data.quer)).toBe(true);
  expect(crop(placeholder)).toEqual(['25000', '0', '25000', '0']);

  // Backup replaced; 2:1 into 1700000 × 2400000: left and right off
  const withBackup = named('Bildplatzhalter mit Ersatzbild');
  expect(first(withBackup, 'p:ph').getAttribute('idx')).toBe('13');
  expect((await pictureIn(withBackup)).equals(data.quer)).toBe(true);
  const side = String(Math.round(((1 - 1700000 / 2400000 / 2) / 2) * 100000));
  expect(crop(withBackup)).toEqual([side, '0', side, '0']);

  const filled = named('Form mit Text');
  expect(filled.tagName).toBe('p:sp');
  expect((await pictureIn(filled)).equals(data.hoch)).toBe(true);
  // 1:2 into 1.75:1: the fill rectangle reaches (1.75 / 0.5 − 1) / 2 of
  // the height past the top and the bottom
  expect(crop(filled)).toEqual(['0', '0', '0', '0']);
  const fillRect = first(filled, 'a:fillRect');
  expect(
    ['l', 't', 'r', 'b'].map((a) => fillRect.getAttribute(a) || '0'),
  ).toEqual(['0', '-125000', '0', '-125000']);
  expect(first(filled, 'a:t').textContent).toBe('Text bleibt');

  const grouped = named('Bild in Gruppe');
  expect((await pictureIn(grouped)).equals(data.quadrat)).toBe(true);
  expect(crop(grouped)).toEqual(['0', '25000', '0', '25000']);

  expect(first(named('Bild mit IF'), 'a:blip').getAttribute('r:embed')).toBe(
    'rIdBackup',
  );
  expect(
    (await pictureIn(named('Bild mit Ternär'))).equals(data.blauBild),
  ).toBe(true);
});
