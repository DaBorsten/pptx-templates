import { XmlElement } from '../../types/xml-types';
import { XmlHelper } from '../../helper/xml-helper';
import { PresetAdjustmentDefaults, ShapePresets } from './presets';
import {
  CornerRadius,
  RoundedShapePreset,
  ShapeAdjustments,
  ShapePreset,
} from './types';

/**
 * Shape geometry: crop a picture or shape to any preset geometry ("Crop to
 * Shape" in PowerPoint) and set corner radii in cm, px, pt or percent.
 */

export { ShapePresets } from './presets';
export * from './types';

/** Name of the adjust value that holds the corner radius, per shape. */
const CORNER_RADIUS_ADJUSTMENTS: Record<RoundedShapePreset, string> = {
  roundRect: 'adj',
  round1Rect: 'adj',
  round2SameRect: 'adj1',
  round2DiagRect: 'adj1',
  snipRoundRect: 'adj1',
};

const EMU_PER_UNIT: Record<string, number> = {
  cm: 360000,
  mm: 36000,
  pt: 12700,
  px: 9525, // CSS pixel, 96 dpi
};

/** CT_ShapeProperties sequence (children of p:spPr). */
const SPPR_CHILD_ORDER = [
  'a:xfrm',
  'a:custGeom',
  'a:prstGeom',
  'a:noFill',
  'a:solidFill',
  'a:gradFill',
  'a:blipFill',
  'a:pattFill',
  'a:grpFill',
  'a:ln',
  'a:effectLst',
  'a:effectDag',
  'a:scene3d',
  'a:sp3d',
  'a:extLst',
];

export type Geometry = { prst: ShapePreset; adjustments: ShapeAdjustments };

export const isShapePreset = (name: unknown): name is ShapePreset =>
  typeof name === 'string' &&
  (ShapePresets as readonly string[]).includes(name);

export const isRoundedShapePreset = (
  name: unknown,
): name is RoundedShapePreset =>
  typeof name === 'string' && Object.hasOwn(CORNER_RADIUS_ADJUSTMENTS, name);

const assertPreset = (shape: unknown): ShapePreset => {
  if (!isShapePreset(shape)) {
    throw new Error(
      `invalid shape: ${String(shape)} (must be a DrawingML preset geometry, e.g. 'roundRect' or 'ellipse')`,
    );
  }
  return shape;
};

/**
 * Validates adjust values and completes them: PowerPoint asks to repair a
 * file whose <a:avLst> names only some of the guides of a preset, so the
 * missing ones get their defaults (in definition order). Empty stays empty,
 * which means "all defaults".
 */
const completeAdjustments = (
  prst: ShapePreset,
  adjustments: ShapeAdjustments = {},
): ShapeAdjustments => {
  const defaults = PresetAdjustmentDefaults[prst] ?? {};
  for (const [name, value] of Object.entries(adjustments)) {
    if (!Object.hasOwn(defaults, name)) {
      const known = Object.keys(defaults);
      throw new Error(
        `shape ${prst} has no adjustment ${name} ` +
          (known.length ? `(use ${known.join(', ')})` : '(it has none)'),
      );
    }
    if (!Number.isFinite(value)) {
      throw new Error(`invalid shape adjustment ${name}: ${value}`);
    }
  }
  if (!Object.keys(adjustments).length) return {};
  return { ...defaults, ...adjustments };
};

/**
 * Corner radius as fraction of the shorter side. `shorterSide` (EMU) is only
 * needed for absolute units.
 */
export const cornerRadiusFraction = (
  radius: CornerRadius,
  shorterSide?: number,
): number => {
  let fraction = NaN;
  let absolute: number | undefined;

  if (typeof radius === 'number') {
    absolute = radius * EMU_PER_UNIT.cm;
  } else if (typeof radius === 'string') {
    const match = /^\s*(\d+(?:\.\d+)?)\s*(%|cm|mm|pt|px)?\s*$/i.exec(radius);
    if (match) {
      const value = Number(match[1]);
      const unit = (match[2] ?? 'cm').toLowerCase();
      if (unit === '%') fraction = value / 100;
      else absolute = value * EMU_PER_UNIT[unit];
    }
  }

  if (absolute !== undefined) {
    if (!(shorterSide > 0)) {
      throw new Error(
        `corner radius ${radius} needs the size of the shape; ` +
          `use a percentage (e.g. '10%') for shapes without their own size`,
      );
    }
    fraction = absolute / shorterSide;
  }

  if (!Number.isFinite(fraction) || fraction < 0) {
    throw new Error(
      `invalid corner radius: ${String(radius)} ` +
        `(a number in cm, or a string like '10%', '12px', '4pt', '0.5cm')`,
    );
  }
  return Math.min(fraction, 0.5);
};

/**
 * Resolves the docx-templates style options into a preset geometry and its
 * adjust values. `shapeAdjustments` win over `cornerRadius`.
 */
export const resolveGeometry = (
  options: {
    shape?: string;
    cornerRadius?: CornerRadius;
    shapeAdjustments?: ShapeAdjustments;
  },
  size?: { cx: number; cy: number },
): Geometry => {
  if (options.cornerRadius == null) {
    const prst = options.shape == null ? 'rect' : assertPreset(options.shape);
    return {
      prst,
      adjustments: completeAdjustments(prst, options.shapeAdjustments),
    };
  }

  const prst = assertPreset(options.shape ?? 'roundRect');
  if (!isRoundedShapePreset(prst)) {
    throw new Error(
      `cornerRadius is not supported for shape ${prst} (use one of ` +
        `${Object.keys(CORNER_RADIUS_ADJUSTMENTS).join(', ')} or shapeAdjustments)`,
    );
  }
  const fraction = cornerRadiusFraction(
    options.cornerRadius,
    size ? Math.min(size.cx, size.cy) : undefined,
  );
  return {
    prst,
    adjustments: completeAdjustments(prst, {
      [CORNER_RADIUS_ADJUSTMENTS[prst]]: Math.round(fraction * 100000),
      ...options.shapeAdjustments,
    }),
  };
};

/** The p:spPr of a shape, picture or connector (direct child only). */
const shapeProperties = (element: XmlElement, helper: string): XmlElement => {
  const spPr = XmlHelper.getFirstDirectChild(element, ['p:spPr']);
  if (!spPr) {
    throw new Error(
      `${helper}: <${element.nodeName}> has no shape properties ` +
        `(only shapes and pictures have a geometry; modify group members individually)`,
    );
  }
  return spPr;
};

const shapeSize = (spPr: XmlElement) => {
  const xfrm = XmlHelper.getFirstDirectChild(spPr, ['a:xfrm']);
  const ext = xfrm && XmlHelper.getFirstDirectChild(xfrm, ['a:ext']);
  if (!ext) return undefined;
  return {
    cx: Number(ext.getAttribute('cx')),
    cy: Number(ext.getAttribute('cy')),
  };
};

const currentPreset = (spPr: XmlElement): string | undefined =>
  XmlHelper.getFirstDirectChild(spPr, ['a:prstGeom'])?.getAttribute('prst');

/** Writes <a:prstGeom> into p:spPr, replacing a preset or custom geometry. */
export const applyGeometry = (spPr: XmlElement, geometry: Geometry): void => {
  const doc = spPr.ownerDocument;
  const prstGeom = doc.createElement('a:prstGeom');
  prstGeom.setAttribute('prst', geometry.prst);
  const avLst = doc.createElement('a:avLst');
  for (const [name, value] of Object.entries(geometry.adjustments)) {
    const gd = doc.createElement('a:gd');
    gd.setAttribute('name', name);
    gd.setAttribute('fmla', `val ${Math.round(value)}`);
    avLst.appendChild(gd);
  }
  prstGeom.appendChild(avLst);

  const existing = XmlHelper.getFirstDirectChild(spPr, [
    'a:prstGeom',
    'a:custGeom',
  ]);
  if (existing) {
    spPr.replaceChild(prstGeom, existing);
  } else {
    XmlHelper.insertInSchemaOrder(
      spPr,
      prstGeom as unknown as XmlElement,
      SPPR_CHILD_ORDER,
    );
  }
};

/**
 * Crops a picture or shape to a preset geometry, like "Crop to Shape" in
 * PowerPoint: `setGeometry('ellipse')`, `setGeometry('heart')`. Replaces the
 * current preset or custom geometry.
 *
 * @param shape - Any DrawingML preset geometry (see `ShapePresets`)
 * @param adjustments - Optional raw adjust values, e.g. `{ adj: 25000 }`
 */
export const setGeometry =
  (shape: ShapePreset, adjustments?: ShapeAdjustments) =>
  (element: XmlElement): void => {
    const spPr = shapeProperties(element, 'setGeometry');
    const prst = assertPreset(shape);
    applyGeometry(spPr, {
      prst,
      adjustments: completeAdjustments(prst, adjustments),
    });
  };

/**
 * Rounds the corners of a picture or shape. The radius is measured on the
 * shorter side and capped at half of it:
 * `setCornerRadius('10%')`, `setCornerRadius('12px')`, `setCornerRadius(0.5)`
 * (cm). A shape that already is one of the rounded presets keeps it;
 * anything else becomes a `roundRect`, unless `shape` says otherwise.
 *
 * @param radius - Number in cm, or a string in %, px, pt, cm or mm
 * @param shape - Rounded preset to use, e.g. 'round2SameRect' (top corners)
 */
export const setCornerRadius =
  (radius: CornerRadius, shape?: RoundedShapePreset) =>
  (element: XmlElement): void => {
    const spPr = shapeProperties(element, 'setCornerRadius');
    const current = currentPreset(spPr);
    const preset =
      shape ?? (isRoundedShapePreset(current) ? current : 'roundRect');
    let geometry: Geometry;
    try {
      geometry = resolveGeometry(
        { shape: preset, cornerRadius: radius },
        shapeSize(spPr),
      );
    } catch (error) {
      throw new Error(`setCornerRadius: ${(error as Error).message}`, {
        cause: error,
      });
    }
    applyGeometry(spPr, geometry);
  };
