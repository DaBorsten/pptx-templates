/**
 * Fork extensions (not part of upstream singerla/pptx-automizer).
 *
 * Everything the fork adds lives below src/extensions/ and is wired into the
 * public API from this one file, so upstream merges stay conflict-free.
 * See FORK.md.
 */
import { renderTemplate } from './template-logic';
import { setCornerRadius, setGeometry } from './shape-geometry';

/** Spread into the public `modify` namespace in src/index.ts. */
export const extensionModifiers = {
  renderTemplate,
  setGeometry,
  setCornerRadius,
};

export * from './template-logic';
export { ShapePresets } from './shape-geometry';
// `export type`: Bun transpiles file by file and can't tell types from
// values in a mixed re-export; it would look for these at runtime.
export type {
  ShapePreset,
  RoundedShapePreset,
  CornerRadius,
  ShapeAdjustments,
  ImageBorder,
} from './shape-geometry';
