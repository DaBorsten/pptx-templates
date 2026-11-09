import IArchive from '../../interfaces/iarchive';
import { RootPresTemplate } from '../../interfaces/root-pres-template';
import { XmlElement } from '../../types/xml-types';
import { XmlHelper } from '../../helper/xml-helper';
import { XmlRelationshipHelper } from '../../helper/xml-relationship-helper';
import { PptPaths } from '../../helper/ppt-paths';
import { imageDimensions } from '../../helper/image-dimensions';
import { ContentTypeExtension } from '../../enums/content-type-map';
import ModifyShapeHelper from '../../helper/modify-shape-helper';
import { applyGeometry, resolveGeometry } from '../shape-geometry';
import { TemplateImage } from './types';

/**
 * {{IMAGE …}} support: writes the image into ppt/media (deduplicated), adds
 * the content type and the slide relationship, and places a <p:pic> where
 * the shape holding the tag is.
 */

const REL_TYPE_IMAGE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

const MIME_EXTENSIONS: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
};
const SUPPORTED = ['png', 'jpg', 'jpeg', 'gif'];

const EMU_PER_CM = 360000;
const EMU_PER_PT = 12700;

export type ImageTarget = {
  archive: IArchive;
  template: RootPresTemplate;
  relsPath: string;
  relations: XmlElement;
  /** Highest shape id handed out so far in this render. */
  lastShapeId?: number;
};

type Frame = { x: number; y: number; cx: number; cy: number };

type LoadedImage = {
  spec: TemplateImage;
  content: Buffer;
  extension: string;
};

/** Validates an IMAGE value and decodes its data. */
export const loadImage = (value: unknown): LoadedImage => {
  if (!value || typeof value !== 'object') {
    throw new Error(
      'IMAGE needs an object, e.g. {{IMAGE ({ data, extension: ".png" })}}',
    );
  }
  const spec = value as TemplateImage;
  let { data } = spec;
  let extension = spec.extension?.replace(/^\./, '').toLowerCase();

  if (typeof data === 'string') {
    const dataUrl = /^data:([^;,]+);base64,([\s\S]*)$/.exec(data);
    if (dataUrl) {
      extension ??= MIME_EXTENSIONS[dataUrl[1].toLowerCase()];
      data = dataUrl[2];
    }
    data = Buffer.from(data as string, 'base64');
  }

  let content: Buffer;
  if (Buffer.isBuffer(data)) content = data;
  else if (data instanceof ArrayBuffer) content = Buffer.from(data);
  else if (data instanceof Uint8Array) content = Buffer.from(data);
  else throw new Error('IMAGE: "data" must be a Buffer or a base64 string');

  if (!extension || !SUPPORTED.includes(extension)) {
    throw new Error(
      `IMAGE: unsupported extension "${extension ?? ''}" (use ${SUPPORTED.join(', ')})`,
    );
  }
  if (spec.border) {
    const { color, width } = spec.border;
    if (typeof color !== 'string' || !/^#?[0-9a-f]{6}$/i.test(color)) {
      throw new Error(
        `IMAGE: invalid border color ${String(color)} (a hex RGB color, e.g. '#ff0000')`,
      );
    }
    if (!Number.isFinite(width) || width < 0) {
      throw new Error(`IMAGE: invalid border width ${String(width)} (in pt)`);
    }
  }
  // Fails early on a bad shape or radius; the real size comes with the slot.
  geometryFor(spec, { cx: 1, cy: 1 });
  return { spec, content, extension };
};

const geometryFor = (spec: TemplateImage, size: { cx: number; cy: number }) => {
  try {
    return resolveGeometry(spec, size);
  } catch (error) {
    throw new Error(`IMAGE: ${(error as Error).message}`, { cause: error });
  }
};

/** shape/cornerRadius/shapeAdjustments and border of an IMAGE picture. */
const styleShape = (picture: XmlElement, image: LoadedImage, box: Frame) => {
  const { spec } = image;
  const spPr = XmlHelper.getFirstDirectChild(picture, ['p:spPr']);
  applyGeometry(spPr, geometryFor(spec, box));
  if (spec.border) {
    ModifyShapeHelper.setOutline({
      weight: Math.round(spec.border.width * EMU_PER_PT),
      color: {
        type: 'srgbClr',
        value: spec.border.color.replace(/^#/, '').toUpperCase(),
      },
    })(picture);
  }
};

const childElement = (parent: XmlElement, tagName: string) => {
  for (let node = parent?.firstChild; node; node = node.nextSibling) {
    if (node.nodeType === 1 && (node as XmlElement).tagName === tagName) {
      return node as XmlElement;
    }
  }
  return undefined;
};

const shapeFrame = (shape: XmlElement): Frame | undefined => {
  const xfrm = childElement(childElement(shape, 'p:spPr'), 'a:xfrm');
  const off = childElement(xfrm, 'a:off');
  const ext = childElement(xfrm, 'a:ext');
  if (!off || !ext) return undefined;
  return {
    x: Number(off.getAttribute('x')),
    y: Number(off.getAttribute('y')),
    cx: Number(ext.getAttribute('cx')),
    cy: Number(ext.getAttribute('cy')),
  };
};

/** Size in EMU: explicit cm, the other side from the aspect ratio, else fit. */
const imageSize = (image: LoadedImage, slot: Frame) => {
  let ratio = 1;
  try {
    const { width, height } = imageDimensions(image.content);
    if (width > 0 && height > 0) ratio = width / height;
  } catch {
    // Unknown dimensions: square
  }
  const { width, height } = image.spec;
  if (width && height) {
    return { cx: width * EMU_PER_CM, cy: height * EMU_PER_CM };
  }
  if (width)
    return { cx: width * EMU_PER_CM, cy: (width * EMU_PER_CM) / ratio };
  if (height) {
    return { cx: height * EMU_PER_CM * ratio, cy: height * EMU_PER_CM };
  }
  const scale = Math.min(slot.cx / ratio, slot.cy);
  return { cx: scale * ratio, cy: scale };
};

const mediaFile = async (
  target: ImageTarget,
  image: LoadedImage,
): Promise<string> => {
  const { archive, template } = target;
  const existing = await template.mediaDeduplicator.find(image.content);
  if (existing) return existing;

  let name: string;
  do {
    name = `image${template.incrementCounter('images')}.${image.extension}`;
  } while (archive.fileExists(PptPaths.media(name)));

  await archive.write(PptPaths.media(name), image.content);
  archive.contentTracker?.trackFile(PptPaths.media(name));
  template.mediaDeduplicator.add(image.content, name);
  await XmlHelper.appendImageExtensionToContentType(
    archive,
    image.extension as ContentTypeExtension,
  );
  return name;
};

const addRelation = (target: ImageTarget, file: string): string => {
  const root = target.relations;
  const id = `rId${XmlRelationshipHelper.getMaxId(
    root.getElementsByTagName('Relationship'),
    'Id',
    true,
  )}`;
  const attributes = {
    Id: id,
    Type: REL_TYPE_IMAGE,
    Target: `../media/${file}`,
  };
  const rel = root.ownerDocument.createElement('Relationship');
  Object.entries(attributes).forEach(([k, v]) => rel.setAttribute(k, v));
  root.appendChild(rel);
  target.archive.contentTracker?.trackRelation(target.relsPath, attributes);
  return id;
};

/**
 * Scans the slide once, then counts up: xmldom's live NodeList does not
 * reliably pick up the pictures inserted in between.
 */
const nextShapeId = (target: ImageTarget, shape: XmlElement): number => {
  target.lastShapeId ??= XmlHelper.collectionToArray(
    shape.ownerDocument.getElementsByTagName('p:cNvPr'),
  ).reduce((max, el) => Math.max(max, Number(el.getAttribute('id')) || 0), 0);
  return ++target.lastShapeId;
};

const createPicture = (
  shape: XmlElement,
  relId: string,
  image: LoadedImage,
  box: Frame,
  id: number,
): XmlElement => {
  const doc = shape.ownerDocument;
  const el = (
    tag: string,
    attrs: Record<string, string | number> = {},
    ...children: XmlElement[]
  ) => {
    const node = doc.createElement(tag) as unknown as XmlElement;
    Object.entries(attrs).forEach(([k, v]) => node.setAttribute(k, String(v)));
    children.forEach((child) => node.appendChild(child));
    return node;
  };
  const xfrmAttrs: Record<string, number> = {};
  if (image.spec.rotation) {
    xfrmAttrs.rot = Math.round(image.spec.rotation * 60000);
  }

  return el(
    'p:pic',
    {},
    el(
      'p:nvPicPr',
      {},
      el('p:cNvPr', {
        id,
        name: `Picture ${id}`,
        ...(image.spec.alt ? { descr: image.spec.alt } : {}),
      }),
      el('p:cNvPicPr', {}, el('a:picLocks', { noChangeAspect: 1 })),
      el('p:nvPr'),
    ),
    el(
      'p:blipFill',
      {},
      el('a:blip', { 'r:embed': relId }),
      el('a:stretch', {}, el('a:fillRect')),
    ),
    el(
      'p:spPr',
      {},
      el(
        'a:xfrm',
        xfrmAttrs,
        el('a:off', { x: Math.round(box.x), y: Math.round(box.y) }),
        el('a:ext', { cx: Math.round(box.cx), cy: Math.round(box.cy) }),
      ),
    ),
  );
};

/**
 * Places the images rendered from one shape's text: side by side, each in
 * an equal share of the shape's frame and centered there. The pictures are
 * inserted right after the shape; the shape is removed if no text is left.
 */
export const placeImages = async (
  target: ImageTarget,
  shape: XmlElement,
  images: LoadedImage[],
  removeShape: boolean,
): Promise<void> => {
  const frame = shapeFrame(shape);
  if (!frame) {
    throw new Error(
      'IMAGE needs a shape with its own position and size ' +
        '(placeholders that inherit them from the layout are not supported)',
    );
  }
  const slotWidth = frame.cx / images.length;
  let anchor = shape.nextSibling;

  for (let i = 0; i < images.length; i++) {
    const slot = { ...frame, x: frame.x + i * slotWidth, cx: slotWidth };
    const size = imageSize(images[i], slot);
    const box = {
      x: slot.x + (slot.cx - size.cx) / 2,
      y: slot.y + (slot.cy - size.cy) / 2,
      ...size,
    };
    const relId = addRelation(target, await mediaFile(target, images[i]));
    const picture = createPicture(
      shape,
      relId,
      images[i],
      box,
      nextShapeId(target, shape),
    );
    styleShape(picture, images[i], box);
    shape.parentNode.insertBefore(picture, anchor);
    anchor = picture.nextSibling;
  }

  if (removeShape) shape.parentNode.removeChild(shape);
};
