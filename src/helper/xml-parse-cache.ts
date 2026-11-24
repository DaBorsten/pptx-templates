import { createHash } from 'crypto';
import { DOMParser } from '@xmldom/xmldom';
import { XmlDocument } from '../types/xml-types';
import { cloneDocument } from './xml-clone';

/**
 * Parses XML parts, and copies the result for a part it has parsed before
 * instead of parsing it again: copying (cloneDocument) takes about a
 * quarter of the time. A slide added several times, or the many identical
 * `.rels` parts of a deck, are read again and again.
 *
 * Memory: a text parsed for the first time leaves only its checksum (a
 * text seen once keeps nothing else). The second parse of the same text
 * keeps the text and an untouched copy of its document, which the third
 * and later ones copy. A parsed document takes about 36 bytes per
 * character of XML, so the copies kept are limited to LIMIT characters of
 * text (about 150 MB at most, a few MB for usual templates), and texts
 * above MAX_TEXT are always parsed. A checksum that matches by chance only
 * keeps a copy that is never used: copies are looked up by the full text.
 */
export class XmlParseCache {
  /** Longer texts are always parsed. */
  static readonly MAX_TEXT = 1_000_000;
  /** Characters of the texts whose parsed copy is kept. */
  static readonly LIMIT = 4_000_000;
  /** Checksums of texts seen once (about 100 bytes each). */
  static readonly MAX_SEEN = 100_000;

  private readonly seenOnce = new Set<string>();
  private readonly copies = new Map<string, XmlDocument>();
  private chars = 0;

  /** Characters of text whose parsed copy is kept (for tests). */
  get keptChars(): number {
    return this.chars;
  }

  parse(xmlString: string): XmlDocument {
    const known = this.copies.get(xmlString);
    if (known) {
      const copy = cloneDocument(known);
      if (copy) return copy;
    }
    const parsed = new DOMParser().parseFromString(
      xmlString,
      'application/xml',
    ) as unknown as XmlDocument;
    if (!known && xmlString.length <= XmlParseCache.MAX_TEXT) {
      this.remember(xmlString, parsed);
    }
    return parsed;
  }

  private remember(xmlString: string, parsed: XmlDocument): void {
    const checksum = createHash('sha1').update(xmlString).digest('base64');
    if (!this.seenOnce.has(checksum)) {
      if (this.seenOnce.size < XmlParseCache.MAX_SEEN) {
        this.seenOnce.add(checksum);
      }
      return;
    }
    if (this.chars + xmlString.length > XmlParseCache.LIMIT) return;
    // copied before the caller changes `parsed`
    const pristine = cloneDocument(parsed);
    if (!pristine) return;
    this.seenOnce.delete(checksum);
    this.copies.set(xmlString, pristine);
    this.chars += xmlString.length;
  }
}
