/**
 * Fork extensions (not part of upstream singerla/pptx-automizer).
 *
 * Everything the fork adds lives below src/extensions/ and is wired into the
 * public API from this one file, so upstream merges stay conflict-free.
 * See FORK.md.
 */
import { renderTemplate } from './template-logic';

/** Spread into the public `modify` namespace in src/index.ts. */
export const extensionModifiers = {
  renderTemplate,
};

export * from './template-logic';
