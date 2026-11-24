import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { XmlParseCache } from '../src/helper/xml-parse-cache';

/*
 * The archive's parse cache: copies of parts parsed before, the same as
 * parsing them again, within a memory budget.
 */
const rels = (id: number) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Id="rId${id}" Type="t" Target="../slideLayouts/slideLayout1.xml"/>` +
  '</Relationships>';

const serialize = (doc: unknown) =>
  new XMLSerializer().serializeToString(doc as never);

test('a text parsed again is the same as parsing it, and its own document', () => {
  const cache = new XmlParseCache();
  const text = rels(1);
  const expected = serialize(
    new DOMParser().parseFromString(text, 'application/xml'),
  );
  const docs = [1, 2, 3, 4].map(() => cache.parse(text));
  docs.forEach((doc) => expect(serialize(doc)).toBe(expected));

  // changing one document changes none of the others, nor later parses
  docs[2].documentElement.setAttribute('changed', '1');
  docs[1].documentElement.removeChild(docs[1].documentElement.firstChild);
  expect(serialize(docs[3])).toBe(expected);
  expect(serialize(cache.parse(text))).toBe(expected);
});

test('a text seen once keeps nothing but its checksum', () => {
  const cache = new XmlParseCache();
  for (let i = 0; i < 1000; i++) cache.parse(rels(i));
  expect(cache.keptChars).toBe(0);
  cache.parse(rels(7));
  expect(cache.keptChars).toBe(rels(7).length);
});

test('copies are kept up to LIMIT characters, longer texts are parsed', () => {
  const cache = new XmlParseCache();
  const long = `<a>${'x'.repeat(XmlParseCache.MAX_TEXT)}</a>`;
  cache.parse(long);
  cache.parse(long);
  expect(cache.keptChars).toBe(0);

  const part = (i: number) =>
    `<a i="${i}">${'y'.repeat(XmlParseCache.MAX_TEXT - 20)}</a>`;
  for (let i = 0; i < 6; i++) {
    cache.parse(part(i));
    cache.parse(part(i));
  }
  expect(cache.keptChars).toBeLessThanOrEqual(XmlParseCache.LIMIT);
  expect(cache.keptChars).toBeGreaterThan(XmlParseCache.LIMIT / 2);
  // over the budget: still parsed correctly
  expect(cache.parse(part(5)).documentElement.getAttribute('i')).toBe('5');
});
