import { DOMParser } from '@xmldom/xmldom';
import { XmlDocument } from '../types/xml-types';
import { cloneDocument } from './xml-clone';

/**
 * Parses XML parts, and copies the result for a part it has parsed before
 * instead of parsing it again: copying (cloneDocument) takes about a
 * quarter of the time. A slide added several times, or the many identical
 * `.rels` parts of a deck, are read again and again.
 *
 * A text is remembered when it is parsed for the first time; the second
 * parse keeps an untouched copy, which the third and later ones copy. A
 * part read once costs nothing extra. Texts above MAX_TEXT, and any once
 * LIMIT characters of text are remembered, are always parsed.
 */
export class XmlParseCache {
  static readonly MAX_TEXT = 2_000_000;
  static readonly LIMIT = 32_000_000;

  private readonly seen = new Map<string, XmlDocument | null>();
  private chars = 0;

  parse(xmlString: string): XmlDocument {
    const known = this.seen.get(xmlString);
    if (known) {
      const copy = cloneDocument(known);
      if (copy) return copy;
    }
    const parsed = new DOMParser().parseFromString(
      xmlString,
      'application/xml',
    ) as unknown as XmlDocument;
    if (known === null) {
      const pristine = cloneDocument(parsed);
      if (pristine) this.seen.set(xmlString, pristine);
    } else if (
      known === undefined &&
      xmlString.length <= XmlParseCache.MAX_TEXT &&
      this.chars + xmlString.length <= XmlParseCache.LIMIT
    ) {
      this.seen.set(xmlString, null);
      this.chars += xmlString.length;
    }
    return parsed;
  }
}
