/**
 * Escapes a string so that it matches itself literally in a RegExp, both
 * as a pattern and inside a character class (`[^…]`). Replaces the
 * `regexp.escape` polyfill, which pulled in 91 packages; `RegExp.escape`
 * itself is only built into Node 24 and later.
 *
 * The result is meant for patterns without the `u`/`v` flag, where an
 * escaped `-` or `/` is allowed.
 */
export const escapeRegExp = (value: string): string =>
  value.replace(/[\\^$.*+?()[\]{}|\-/]/g, '\\$&');
