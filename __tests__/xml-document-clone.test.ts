import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { cloneDocument } from '../src/helper/xml-clone';

/**
 * cloneDocument stands in for parsing the same XML again (the parse cache
 * of Archive.parseXml); the copy must be what a new parse gives.
 */

// xmldom internals (_nsMap, own fields) are compared field by field
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyNode = Record<string, any>;

/**
 * The scalar properties as they read: own ones and the ones of the
 * prototype (a parsed element has its own `prefix: null`, a copy reads
 * the same null from the prototype, as xmldom's cloneNode makes it).
 */
const PROTOTYPE_SCALARS = [
  'nodeName',
  'nodeType',
  'tagName',
  'localName',
  'prefix',
  'namespaceURI',
  'name',
  'value',
  'nodeValue',
  'data',
  'target',
  'specified',
];
const scalars = (node: AnyNode): string =>
  JSON.stringify(
    [...new Set([...Object.keys(node), ...PROTOTYPE_SCALARS])]
      .filter((k) => k !== '_inc')
      .filter((k) => typeof node[k] !== 'object' || node[k] === null)
      .sort()
      .map((k) => [k, node[k]]),
  );

const ownMap = (map: AnyNode | undefined) =>
  map ? JSON.stringify(Object.keys(map).map((k) => [k, map[k]])) : '-';

const difference = (
  a: AnyNode,
  b: AnyNode,
  docA: AnyNode,
  docB: AnyNode,
  at: string,
): string | undefined => {
  if (scalars(a) !== scalars(b)) return `${at}: ${scalars(a)} / ${scalars(b)}`;
  if (a.ownerDocument !== docA || b.ownerDocument !== docB) {
    return `${at}: ownerDocument`;
  }
  if (ownMap(a._nsMap) !== ownMap(b._nsMap)) return `${at}: _nsMap`;
  if (a.attributes) {
    if (a.attributes.length !== b.attributes.length) return `${at}: attrs`;
    for (let i = 0; i < a.attributes.length; i++) {
      const x = a.attributes.item(i);
      const y = b.attributes.item(i);
      if (y.ownerElement !== b) return `${at}: ownerElement`;
      const d = difference(x, y, docA, docB, `${at}@${x.name}`);
      if (d) return d;
    }
    // lookups by name and namespace, as code uses them
    for (let i = 0; i < a.attributes.length; i++) {
      const x = a.attributes.item(i);
      if (b.getAttribute(x.name) !== a.getAttribute(x.name)) {
        return `${at}: getAttribute(${x.name})`;
      }
      if (
        x.namespaceURI &&
        b.getAttributeNS(x.namespaceURI, x.localName) !==
          a.getAttributeNS(x.namespaceURI, x.localName)
      ) {
        return `${at}: getAttributeNS(${x.name})`;
      }
    }
  }
  const ca = Array.from((a.childNodes ?? []) as AnyNode[]);
  const cb = Array.from((b.childNodes ?? []) as AnyNode[]);
  if (ca.length !== cb.length) return `${at}: child count`;
  for (let i = 0; i < ca.length; i++) {
    if (cb[i].parentNode !== b) return `${at}/${i}: parentNode`;
    const d = difference(ca[i], cb[i], docA, docB, `${at}/${i}`);
    if (d) return d;
  }
  return undefined;
};

const expectSameAsParse = (xml: string, at: string) => {
  const parse = () =>
    new DOMParser().parseFromString(
      xml,
      'application/xml',
    ) as unknown as AnyNode;
  const parsed = parse();
  const copy = cloneDocument(parse()) as AnyNode;
  expect(copy).toBeDefined();
  expect(difference(parsed, copy, parsed, copy, at)).toBeUndefined();
  for (const key of ['contentType', 'type', 'doctype']) {
    expect([at, key, copy[key]]).toEqual([at, key, parsed[key]]);
  }
  expect(copy.ownerDocument === copy).toBe(parsed.ownerDocument === parsed);
  expect(copy.implementation).toBeTruthy();
  expect(copy.documentElement?.tagName).toBe(parsed.documentElement?.tagName);
  expect(copy.documentElement).toBe(copy.lastChild);
  const s = new XMLSerializer();
  expect(s.serializeToString(copy as never)).toBe(
    s.serializeToString(parsed as never),
  );
  // namespaces resolve the same from every element
  const elsA = Array.from(parsed.getElementsByTagName('*')) as AnyNode[];
  const elsB = Array.from(copy.getElementsByTagName('*')) as AnyNode[];
  expect(elsB.length).toBe(elsA.length);
  for (let i = 0; i < elsA.length; i += Math.ceil(elsA.length / 30)) {
    const prefix = elsA[i].prefix ?? null;
    expect(elsB[i].lookupNamespaceURI(prefix)).toBe(
      elsA[i].lookupNamespaceURI(prefix),
    );
    expect(elsB[i].lookupPrefix(elsA[i].namespaceURI)).toBe(
      elsA[i].lookupPrefix(elsA[i].namespaceURI),
    );
  }
};

test('cloneDocument gives what parsing again gives, for every template part', async () => {
  const dir = `${__dirname}/pptx-templates`;
  let compared = 0;
  for (const file of fs.readdirSync(dir).filter((f) => f.endsWith('.pptx'))) {
    const zip = await JSZip.loadAsync(fs.readFileSync(`${dir}/${file}`));
    for (const part of Object.keys(zip.files)) {
      if (!/\.(xml|rels)$/.test(part)) continue;
      expectSameAsParse(
        await zip.file(part).async('string'),
        `${file}:${part}`,
      );
      compared++;
    }
  }
  expect(compared).toBeGreaterThan(1000);
}, 120_000);

test('cloneDocument: declaration, whitespace, default namespaces', () => {
  expectSameAsParse(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Types xmlns="urn:t"><Default Extension="xml" ContentType="a/b"/>' +
      '<inner xmlns="" xmlns:q="urn:q" q:a="1">text &amp; more</inner></Types>',
    'default namespace',
  );
  expectSameAsParse('<root/>', 'no declaration');
  // other children: parse instead
  const withComment = new DOMParser().parseFromString(
    '<!-- c --><root/>',
    'application/xml',
  );
  expect(cloneDocument(withComment)).toBeUndefined();
  const html = new DOMParser().parseFromString('<p/>', 'text/html');
  expect(cloneDocument(html)).toBeUndefined();
});
