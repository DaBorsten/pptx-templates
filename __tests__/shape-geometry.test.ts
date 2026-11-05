import Automizer, {
  CornerRadius,
  modify,
  TemplateData,
  XmlDocument,
  XmlElement,
} from '../src/index';
import { expectXml } from './helpers/expect-xml';
import fs from 'fs';

// removeExistingSlides keeps slide1.xml of the root as a part, so the
// added slides start at slide2.xml.
const newAutomizer = () =>
  new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
    removeExistingSlides: true,
  });

/** The <p:sp> or <p:pic> with the given name. */
const findShape = (doc: XmlDocument, name: string): XmlElement => {
  const shape = Array.from(doc.getElementsByTagName('p:cNvPr')).find(
    (n) => n.getAttribute('name') === name,
  );
  if (!shape) throw new Error(`no shape ${name}`);
  return shape.parentNode.parentNode as XmlElement;
};

const directChildren = (element: XmlElement, tag?: string) =>
  Array.from(element.childNodes).filter(
    (n): n is XmlElement =>
      n.nodeType === 1 && (!tag || (n as XmlElement).nodeName === tag),
  );

/** prst, adjust values and the child order of p:spPr. */
const geometry = (shape: XmlElement) => {
  const spPr = directChildren(shape, 'p:spPr')[0];
  const prstGeom = directChildren(spPr, 'a:prstGeom');
  expect(prstGeom).toHaveLength(1);
  expect(directChildren(spPr, 'a:custGeom')).toHaveLength(0);
  const adjustments = Object.fromEntries(
    Array.from(prstGeom[0].getElementsByTagName('a:gd')).map((gd) => [
      gd.getAttribute('name'),
      gd.getAttribute('fmla'),
    ]),
  );
  return {
    prst: prstGeom[0].getAttribute('prst'),
    adjustments,
    children: directChildren(spPr).map((n) => n.nodeName),
  };
};

const shorterSide = (shape: XmlElement) => {
  const xfrm = shape.getElementsByTagName('a:xfrm')[0];
  const ext = xfrm.getElementsByTagName('a:ext')[0];
  return Math.min(
    Number(ext.getAttribute('cx')),
    Number(ext.getAttribute('cy')),
  );
};

describe('modify.setGeometry / modify.setCornerRadius', () => {
  test('crops pictures and shapes to presets, rounds corners', async () => {
    const file = 'shape-geometry.test.pptx';
    await newAutomizer()
      .loadRoot('RootTemplate.pptx')
      .load('SlideWithImages.pptx', 'images')
      .load('SlideWithShapes.pptx', 'shapes')
      .addSlide('images', 2, (slide) => {
        slide.modifyElement('imagePNG', modify.setGeometry('ellipse'));
        slide.modifyElement('imageJPG', modify.setCornerRadius('12px'));
        slide.modifyElement('imagePNGduotone', [
          modify.setCornerRadius('20%', 'round2SameRect'),
          modify.setOutline({
            weight: 25400,
            color: { type: 'srgbClr', value: '1F4E79' },
          }),
        ]);
      })
      .addSlide('shapes', 2, (slide) => {
        slide.modifyElement('Cloud', modify.setGeometry('heart'));
        slide.modifyElement(
          'Star',
          modify.setGeometry('star5', { adj: 30000 }),
        );
        slide.modifyElement('Arrow', modify.setCornerRadius('80%'));
        slide.modifyElement('Drum', modify.setCornerRadius(0.5));
      })
      .write(file);

    const images = (await expectXml(file, 'ppt/slides/slide2.xml')).doc();
    // Preset replaced, nothing else in p:spPr moved
    const png = geometry(findShape(images, 'imagePNG'));
    expect(png).toMatchObject({ prst: 'ellipse', adjustments: {} });
    expect(png.children.indexOf('a:prstGeom')).toBe(
      png.children.indexOf('a:xfrm') + 1,
    );

    const jpgShape = findShape(images, 'imageJPG');
    const jpg = geometry(jpgShape);
    expect(jpg.prst).toBe('roundRect');
    expect(jpg.adjustments).toEqual({
      adj: `val ${Math.round(((12 * 9525) / shorterSide(jpgShape)) * 100000)}`,
    });

    const duotone = geometry(findShape(images, 'imagePNGduotone'));
    // PowerPoint repairs an <a:avLst> that names only some guides:
    // adj2 is completed with its default
    expect(duotone.prst).toBe('round2SameRect');
    expect(duotone.adjustments).toEqual({ adj1: 'val 20000', adj2: 'val 0' });
    // Outline comes after the geometry (CT_ShapeProperties order)
    expect(duotone.children.indexOf('a:ln')).toBeGreaterThan(
      duotone.children.indexOf('a:prstGeom'),
    );

    const shapes = (await expectXml(file, 'ppt/slides/slide3.xml')).doc();
    expect(geometry(findShape(shapes, 'Cloud'))).toMatchObject({
      prst: 'heart',
      adjustments: {},
    });
    // hf and vf completed with the defaults of presetShapeDefinitions.xml
    expect(geometry(findShape(shapes, 'Star')).adjustments).toEqual({
      adj: 'val 30000',
      hf: 'val 105146',
      vf: 'val 110557',
    });
    // Capped at 50% of the shorter side
    expect(geometry(findShape(shapes, 'Arrow'))).toMatchObject({
      prst: 'roundRect',
      adjustments: { adj: 'val 50000' },
    });
    const drumShape = findShape(shapes, 'Drum');
    expect(geometry(drumShape).adjustments).toEqual({
      adj: `val ${Math.round(
        Math.min((360000 * 0.5) / shorterSide(drumShape), 0.5) * 100000,
      )}`,
    });
  });

  test('a rounded preset keeps its kind', async () => {
    const file = 'shape-geometry-keep.test.pptx';
    await newAutomizer()
      .loadRoot('RootTemplate.pptx')
      .load('SlideWithImages.pptx', 'images')
      .addSlide('images', 2, (slide) => {
        slide.modifyElement('imagePNG', [
          modify.setGeometry('round2DiagRect'),
          modify.setCornerRadius('5%'),
        ]);
      })
      .write(file);

    const doc = (await expectXml(file, 'ppt/slides/slide2.xml')).doc();
    const kept = geometry(findShape(doc, 'imagePNG'));
    expect(kept.prst).toBe('round2DiagRect');
    expect(kept.adjustments).toEqual({ adj1: 'val 5000', adj2: 'val 0' });
  });

  test.each<[string, () => (element: XmlElement) => void, RegExp]>([
    [
      'unknown preset',
      () => modify.setGeometry('blob' as never),
      /invalid shape: blob/,
    ],
    [
      'unknown adjustment',
      () => modify.setGeometry('star5', { adj9: 1 }),
      /shape star5 has no adjustment adj9 \(use adj, hf, vf\)/,
    ],
    [
      'adjustment on a shape without any',
      () => modify.setGeometry('ellipse', { adj: 1 }),
      /shape ellipse has no adjustment adj \(it has none\)/,
    ],
    [
      'inherited property name as adjustment',
      () => modify.setGeometry('ellipse', { constructor: 1 }),
      /shape ellipse has no adjustment constructor/,
    ],
    [
      'bad radius',
      () => modify.setCornerRadius('1em' as CornerRadius),
      /setCornerRadius: invalid corner radius: 1em/,
    ],
  ])('rejects %s', async (_, callback, message) => {
    await expect(
      newAutomizer()
        .loadRoot('RootTemplate.pptx')
        .load('SlideWithImages.pptx', 'images')
        .addSlide('images', 2, (slide) => {
          slide.modifyElement('imagePNG', callback());
        })
        .write('shape-geometry-error.test.pptx'),
    ).rejects.toThrow(message);
  });
});

describe('IMAGE shape, cornerRadius, border', () => {
  const png = fs.readFileSync(`${__dirname}/media/test.png`);

  const render = (file: string, tag: string, data: TemplateData) =>
    newAutomizer()
      .loadRoot('RootTemplate.pptx')
      .load('TextReplace.pptx')
      .addSlide('TextReplace.pptx', 1, (slide) => {
        slide.modifyElement(
          'setText',
          modify.setMultiText([{ paragraph: {}, textRuns: [{ text: tag }] }]),
        );
        slide.modify(modify.renderTemplate(data));
      })
      .write(file);

  const picture = async (file: string) => {
    const doc = (await expectXml(file, 'ppt/slides/slide2.xml')).doc();
    return Array.from(doc.getElementsByTagName('p:pic')).find((pic) =>
      /^Picture \d+$/.test(
        pic.getElementsByTagName('p:cNvPr')[0].getAttribute('name'),
      ),
    );
  };

  test('rounded corners and border, like docx-templates', async () => {
    const file = 'shape-geometry-image.test.pptx';
    await render(
      file,
      `{{IMAGE ({ data: logo, extension: '.png', width: 4, height: 2, cornerRadius: 0.5, border: { color: '#1f4e79', width: 2 } })}}`,
      { logo: png },
    );
    const pic = await picture(file);
    // 0.5 cm of a 2 cm short side = 25%
    expect(geometry(pic)).toMatchObject({
      prst: 'roundRect',
      adjustments: { adj: 'val 25000' },
      children: ['a:xfrm', 'a:prstGeom', 'a:ln'],
    });
    const ln = pic.getElementsByTagName('a:ln')[0];
    expect(ln.getAttribute('w')).toBe('25400');
    expect(ln.getElementsByTagName('a:srgbClr')[0].getAttribute('val')).toBe(
      '1F4E79',
    );
  });

  test('any preset with raw adjustments', async () => {
    const file = 'shape-geometry-image-shape.test.pptx';
    await render(
      file,
      `{{IMAGE ({ data: logo, extension: '.png', shape: 'star5', shapeAdjustments: { adj: 30000 } })}}`,
      { logo: png },
    );
    const pic = await picture(file);
    expect(geometry(pic)).toMatchObject({
      prst: 'star5',
    });
    expect(geometry(pic).adjustments).toEqual({
      adj: 'val 30000',
      hf: 'val 105146',
      vf: 'val 110557',
    });
    expect(pic.getElementsByTagName('a:ln')).toHaveLength(0);
  });

  test.each([
    [`shape: 'blob'`, /IMAGE: invalid shape: blob/],
    [
      `shape: 'ellipse', cornerRadius: '10%'`,
      /cornerRadius is not supported for shape ellipse/,
    ],
    [`cornerRadius: 'x'`, /IMAGE: invalid corner radius: x/],
    [`border: { color: 'red', width: 1 }`, /IMAGE: invalid border color red/],
  ])('rejects %s', async (options, message) => {
    await expect(
      render(
        'shape-geometry-image-error.test.pptx',
        `{{IMAGE ({ data: logo, extension: '.png', ${options} })}}`,
        { logo: png },
      ),
    ).rejects.toThrow(message);
  });
});
