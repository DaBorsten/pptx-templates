import fs from 'fs';
import path from 'path';
import IArchive from '../../interfaces/iarchive';
import { RootPresTemplate } from '../../interfaces/root-pres-template';
import { XmlElement } from '../../types/xml-types';
import { XmlHelper } from '../../helper/xml-helper';
import { XmlRelationshipHelper } from '../../helper/xml-relationship-helper';
import { PptPaths } from '../../helper/ppt-paths';
import { imageDimensions } from '../../helper/image-dimensions';
import { ContentTypeExtension } from '../../enums/content-type-map';
import ModifyShapeHelper from '../../helper/modify-shape-helper';
import {
  applyGeometry,
  resolveGeometry,
  SPPR_CHILD_ORDER,
} from '../shape-geometry';
import { inheritedPlaceholders, SlideStyles } from '../text-flow/styles';
import { FILL_FIXED_OPTIONS, TemplateImage } from './types';

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
/** A string ending like this is a file path: base64 has no dot. */
const FILE_PATH = /\.(png|jpe?g|gif)$/i;

const EMU_PER_CM = 360000;
const EMU_PER_PT = 12700;

export type ImageTarget = {
  archive: IArchive;
  template: RootPresTemplate;
  relsPath: string;
  relations: XmlElement;
  /** Highest shape id handed out so far in this render. */
  lastShapeId?: number;
  /** Layout and master of the slide, for placeholders without own size. */
  styles?: () => Promise<Pick<SlideStyles, 'layout' | 'master'> | undefined>;
};

type Frame = { x: number; y: number; cx: number; cy: number };

type LoadedImage = {
  spec: TemplateImage;
  content: Buffer;
  extension: string;
};

/** The type of a picture from its first bytes, for data without extension. */
const sniffExtension = (content: Buffer): string | undefined => {
  if (content.subarray(0, 8).equals(Buffer.from('89504e470d0a1a0a', 'hex'))) {
    return 'png';
  }
  if (content[0] === 0xff && content[1] === 0xd8 && content[2] === 0xff) {
    return 'jpeg';
  }
  if (/^GIF8[79]a/.test(content.subarray(0, 6).toString('latin1'))) {
    return 'gif';
  }
  return undefined;
};

const isBinary = (value: unknown): value is Buffer | ArrayBuffer | Uint8Array =>
  Buffer.isBuffer(value) ||
  value instanceof ArrayBuffer ||
  value instanceof Uint8Array;

/** `images/logo.png` from `imageRoot`; the path cannot leave that folder. */
const readImageFile = (file: string, imageRoot: string | undefined) => {
  if (!imageRoot) {
    throw new Error(
      `IMAGE: "${file}" is a file path, but no imageRoot is set to read it from`,
    );
  }
  const root = path.resolve(imageRoot);
  const resolved = path.resolve(root, file);
  const back = path.relative(root, resolved);
  if (!back || back.startsWith('..') || path.isAbsolute(back)) {
    throw new Error(`IMAGE: "${file}" is outside of the imageRoot folder`);
  }
  try {
    return fs.readFileSync(resolved);
  } catch {
    throw new Error(`IMAGE: file "${file}" not found in ${root}`);
  }
};

/**
 * Validates an IMAGE value and loads its data. The value is an object like
 * `{ data, width, … }` or only the picture: a Buffer, a data URL, base64
 * or a file path relative to `imageRoot` (`images/logo.png`).
 */
export const loadImage = (value: unknown, imageRoot?: string): LoadedImage => {
  const spec = (
    typeof value === 'string' || isBinary(value) ? { data: value } : value
  ) as TemplateImage;
  if (!spec || typeof spec !== 'object') {
    throw new Error(
      'IMAGE needs a picture or an object, e.g. {{IMAGE logo}} or {{IMAGE ({ data: logo, width: 4 })}}',
    );
  }
  const { data } = spec;
  let extension = spec.extension?.replace(/^\./, '').toLowerCase();

  let content: Buffer;
  if (typeof data === 'string') {
    const dataUrl = /^data:([^;,]+);base64,([\s\S]*)$/.exec(data);
    if (dataUrl) {
      extension ??= MIME_EXTENSIONS[dataUrl[1].toLowerCase()];
      content = Buffer.from(dataUrl[2], 'base64');
    } else if (FILE_PATH.test(data.trim())) {
      content = readImageFile(data.trim(), imageRoot);
      extension ??= path.extname(data.trim()).slice(1).toLowerCase();
    } else {
      content = Buffer.from(data, 'base64');
    }
  } else if (Buffer.isBuffer(data)) content = data;
  else if (data instanceof ArrayBuffer) content = Buffer.from(data);
  else if (data instanceof Uint8Array) content = Buffer.from(data);
  else {
    throw new Error(
      'IMAGE: "data" must be a Buffer, a data URL, base64 or a file path',
    );
  }
  extension ??= sniffExtension(content);

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
  applyBorder(picture, spec);
};

const applyBorder = (shape: XmlElement, spec: TemplateImage) => {
  if (!spec.border) return;
  ModifyShapeHelper.setOutline({
    weight: Math.round(spec.border.width * EMU_PER_PT),
    color: {
      type: 'srgbClr',
      value: spec.border.color.replace(/^#/, '').toUpperCase(),
    },
  })(shape);
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

/** Width / height of the picture; 1 if its dimensions are unknown. */
const aspectRatio = (image: LoadedImage): number => {
  try {
    const { width, height } = imageDimensions(image.content);
    if (width > 0 && height > 0) return width / height;
  } catch {
    // Unknown dimensions: square
  }
  return 1;
};

/** Size in EMU: explicit cm, the other side from the aspect ratio, else fit. */
const imageSize = (image: LoadedImage, slot: Frame) => {
  const ratio = aspectRatio(image);
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

/*
 * {{IMAGE …}} in the alt text of a picture or shape: the picture takes the
 * place of the one there (or fills the shape), cropped to cover it.
 */

const FILL_TAGS = [
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
];

const createElement = (
  doc: XmlElement['ownerDocument'],
  tag: string,
  attrs: Record<string, string | number> = {},
  ...children: XmlElement[]
): XmlElement => {
  const node = doc.createElement(tag) as unknown as XmlElement;
  Object.entries(attrs).forEach(([k, v]) => node.setAttribute(k, String(v)));
  children.forEach((child) => node.appendChild(child));
  return node;
};

/** Placeholder types PowerPoint can put a picture into. */
const PICTURE_PLACEHOLDERS = ['pic', 'obj', 'clipArt'];

/**
 * An empty placeholder that can hold a picture: it then becomes a picture
 * in that placeholder. A placeholder with text, or one for a title or
 * body, is filled like any shape, so its text stays.
 */
const takesPicture = (shape: XmlElement): boolean => {
  const ph = childElement(
    childElement(childElement(shape, 'p:nvSpPr'), 'p:nvPr'),
    'p:ph',
  );
  if (!ph) return false;
  if (!PICTURE_PLACEHOLDERS.includes(ph.getAttribute('type') || 'obj')) {
    return false;
  }
  const body = childElement(shape, 'p:txBody');
  return !body || (body.textContent ?? '').trim() === '';
};

/**
 * How the groups around `shape` stretch it: ext / chExt of each group,
 * multiplied. 1 × 1 outside a group.
 */
const groupScale = (shape: XmlElement): { sx: number; sy: number } => {
  let sx = 1;
  let sy = 1;
  for (
    let group = shape.parentNode as XmlElement;
    group?.tagName === 'p:grpSp';
    group = group.parentNode as XmlElement
  ) {
    const xfrm = childElement(childElement(group, 'p:grpSpPr'), 'a:xfrm');
    const ext = childElement(xfrm, 'a:ext');
    const chExt = childElement(xfrm, 'a:chExt');
    const ratio = (attr: string) => {
      const child = Number(chExt?.getAttribute(attr));
      const own = Number(ext?.getAttribute(attr));
      return child > 0 && own > 0 ? own / child : 1;
    };
    sx *= ratio('cx');
    sy *= ratio('cy');
  }
  return { sx, sy };
};

/**
 * Size of a shape as shown on the slide (scaled by the groups around it);
 * a placeholder without its own takes the layout's.
 */
const coverSize = async (
  target: ImageTarget,
  shape: XmlElement,
): Promise<{ cx: number; cy: number } | undefined> => {
  let size: { cx: number; cy: number } | undefined = shapeFrame(shape);
  if (!size) {
    const styles = await target.styles?.();
    if (!styles) return undefined;
    size = inheritedPlaceholders(shape, styles).map(shapeFrame).find(Boolean);
    if (!size) return undefined;
  }
  const { sx, sy } = groupScale(shape);
  return { cx: size.cx * sx, cy: size.cy * sy };
};

type Insets = Record<string, number>;

/**
 * <a:srcRect> insets that crop the picture to the aspect ratio of the
 * frame, evenly on both sides, so that it covers the frame without
 * distortion. Values in 1/1000 %.
 */
const coverCrop = (ratio: number, size: { cx: number; cy: number }): Insets => {
  const frame = size.cx / size.cy;
  if (ratio > frame) {
    const side = Math.round(((1 - frame / ratio) / 2) * 100000);
    return side > 0 ? { l: side, r: side } : {};
  }
  if (ratio < frame) {
    const side = Math.round(((1 - ratio / frame) / 2) * 100000);
    return side > 0 ? { t: side, b: side } : {};
  }
  return {};
};

/**
 * The same cover as <a:fillRect> insets: the fill rectangle reaches past
 * the shape (negative insets), the picture is stretched into it. For the
 * picture fill of a shape (a:blipFill), where LibreOffice ignores
 * a:srcRect; PowerPoint reads both.
 */
const coverStretch = (
  ratio: number,
  size: { cx: number; cy: number },
): Insets => {
  const frame = size.cx / size.cy;
  if (ratio > frame) {
    const side = -Math.round(((ratio / frame - 1) / 2) * 100000);
    return side < 0 ? { l: side, r: side } : {};
  }
  if (ratio < frame) {
    const side = -Math.round(((frame / ratio - 1) / 2) * 100000);
    return side < 0 ? { t: side, b: side } : {};
  }
  return {};
};

/** blip + srcRect + stretch, in a p:blipFill or a:blipFill. */
const fillBlip = (
  blipFill: XmlElement,
  relId: string,
  crop: Insets,
  stretch: Insets = {},
): void => {
  const doc = blipFill.ownerDocument;
  // A fresh blip: the old one may carry an SVG or effects of the backup
  for (const tag of ['a:blip', 'a:srcRect', 'a:tile', 'a:stretch']) {
    for (
      let el = childElement(blipFill, tag);
      el;
      el = childElement(blipFill, tag)
    ) {
      blipFill.removeChild(el);
    }
  }
  const first = blipFill.firstChild;
  blipFill.insertBefore(
    createElement(doc, 'a:blip', { 'r:embed': relId }),
    first,
  );
  blipFill.insertBefore(createElement(doc, 'a:srcRect', crop), first);
  blipFill.insertBefore(
    createElement(
      doc,
      'a:stretch',
      {},
      createElement(doc, 'a:fillRect', stretch),
    ),
    first,
  );
};

/** Relationship ids referenced in `el` and below (r:embed, r:link, r:id, …). */
const relationIds = (el: XmlElement | undefined): string[] => {
  if (!el) return [];
  const ids: string[] = [];
  const visit = (node: XmlElement) => {
    for (let i = 0; i < node.attributes.length; i++) {
      const attr = node.attributes[i];
      if (attr.name.startsWith('r:')) ids.push(attr.value);
    }
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1) visit(n as XmlElement);
    }
  };
  visit(el);
  return ids;
};

/** Drops the image relationships of `ids` that nothing references any more. */
const dropUnusedRelations = (
  target: ImageTarget,
  doc: XmlElement['ownerDocument'],
  ids: string[],
): void => {
  if (!ids.length) return;
  const used = new Set(
    relationIds(doc.documentElement as unknown as XmlElement),
  );
  XmlHelper.collectionToArray(
    target.relations.getElementsByTagName('Relationship'),
  )
    .filter(
      (rel) =>
        ids.includes(rel.getAttribute('Id')) &&
        !used.has(rel.getAttribute('Id')) &&
        rel.getAttribute('Type') === REL_TYPE_IMAGE,
    )
    .forEach((rel) => rel.parentNode.removeChild(rel));
};

/** The empty placeholder `shape` as a picture, as PowerPoint inserts one. */
const pictureFromPlaceholder = (shape: XmlElement): XmlElement => {
  const doc = shape.ownerDocument;
  const nvSpPr = childElement(shape, 'p:nvSpPr');
  const spPr = childElement(shape, 'p:spPr');
  return createElement(
    doc,
    'p:pic',
    {},
    createElement(
      doc,
      'p:nvPicPr',
      {},
      cloneNodeDeep(childElement(nvSpPr, 'p:cNvPr')),
      createElement(
        doc,
        'p:cNvPicPr',
        {},
        createElement(doc, 'a:picLocks', { noGrp: 1, noChangeAspect: 1 }),
      ),
      cloneNodeDeep(childElement(nvSpPr, 'p:nvPr')),
    ),
    createElement(doc, 'p:blipFill'),
    spPr ? cloneNodeDeep(spPr) : createElement(doc, 'p:spPr'),
  );
};

const cloneNodeDeep = (el: XmlElement): XmlElement =>
  el.cloneNode(true) as XmlElement;

/**
 * Puts the picture into a picture or shape that has the tag in its alt
 * text, cropped so that it covers the shape without distortion:
 * - a picture (also one in a placeholder, e.g. a backup picture) gets the
 *   new picture instead of its own; position, shape and placeholder stay;
 * - an empty placeholder becomes a picture in that placeholder;
 * - any other shape gets the picture as its fill.
 *
 * Returns the element that now holds the alt text.
 */
export const fillWithImage = async (
  target: ImageTarget,
  shape: XmlElement,
  image: LoadedImage,
): Promise<XmlElement> => {
  const { spec } = image;
  if (shape?.tagName !== 'p:pic' && shape?.tagName !== 'p:sp') {
    throw new Error(
      `IMAGE in alt text works on pictures, placeholders and shapes, not on <${shape?.tagName}>`,
    );
  }
  for (const key of FILL_FIXED_OPTIONS) {
    if (spec[key] !== undefined) {
      throw new Error(
        `IMAGE in alt text fills the shape; "${key}" is not supported there ` +
          '(size and rotation come from the shape)',
      );
    }
  }
  const size = await coverSize(target, shape);
  if (!size?.cx || !size.cy) {
    throw new Error(
      'IMAGE: the shape has no size, and no placeholder on the layout to inherit one from',
    );
  }

  const doc = shape.ownerDocument;
  const relId = addRelation(target, await mediaFile(target, image));
  const ratio = aspectRatio(image);
  const crop = coverCrop(ratio, size);
  let holder = shape;
  let replaced: string[] = [];

  if (shape.tagName === 'p:pic') {
    const blipFill = childElement(shape, 'p:blipFill');
    replaced = relationIds(blipFill);
    fillBlip(blipFill, relId, crop);
  } else if (takesPicture(shape)) {
    holder = pictureFromPlaceholder(shape);
    fillBlip(childElement(holder, 'p:blipFill'), relId, crop);
    shape.parentNode.replaceChild(holder, shape);
  } else {
    const spPr = childElement(shape, 'p:spPr');
    FILL_TAGS.forEach((tag) => {
      const old = childElement(spPr, tag);
      if (!old) return;
      replaced.push(...relationIds(old));
      spPr.removeChild(old);
    });
    const blipFill = createElement(doc, 'a:blipFill', { rotWithShape: 1 });
    fillBlip(blipFill, relId, {}, coverStretch(ratio, size));
    XmlHelper.insertInSchemaOrder(spPr, blipFill, SPPR_CHILD_ORDER);
  }

  // The shape keeps its own geometry (e.g. the circle of the layout)
  // unless the tag asks for another one
  if (spec.shape || spec.cornerRadius != null || spec.shapeAdjustments) {
    applyGeometry(
      XmlHelper.getFirstDirectChild(holder, ['p:spPr']),
      geometryFor(spec, size),
    );
  }
  applyBorder(holder, spec);
  dropUnusedRelations(target, doc, replaced);
  return holder;
};
