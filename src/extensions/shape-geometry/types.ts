import { ShapePresets } from './presets';

/** A DrawingML preset geometry, e.g. 'roundRect', 'ellipse', 'heart'. */
export type ShapePreset = (typeof ShapePresets)[number];

/** Preset geometries whose corner radius `setCornerRadius` can set. */
export type RoundedShapePreset =
  | 'roundRect'
  | 'round1Rect'
  | 'round2SameRect'
  | 'round2DiagRect'
  | 'snipRoundRect';

/**
 * Corner radius, measured on the shorter side of the shape:
 * - a number: centimeters (as in docx-templates)
 * - a string with unit: `'10%'` (of the shorter side), `'12px'` (96 dpi),
 *   `'4pt'`, `'0.5cm'`, `'5mm'`
 *
 * The radius is capped at half the shorter side (`'50%'`).
 */
export type CornerRadius = number | string;

/**
 * Raw adjust values of a preset geometry (the yellow handles in
 * PowerPoint), keyed by guide name, e.g. `{ adj: 25000 }`. Names and value
 * ranges depend on the shape; most are in 1/100000 of the shorter side.
 */
export type ShapeAdjustments = { [name: string]: number };

/** Border of an `{{IMAGE}}` picture (docx-templates compatible). */
export type ImageBorder = {
  /** Hex RGB color, e.g. '#ff0000' or 'ff0000'. */
  color: string;
  /** Width in points. */
  width: number;
};
