import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { cloneXml } from '../src/extensions/template-logic/clone';

/**
 * cloneXml replaces xmldom's cloneNode in the template renderer for speed;
 * it must produce the same nodes. Compared node by node: own scalar
 * properties (what cloneNode copies), the properties read through the
 * prototype, namespace map, attributes, children and their links.
 */

const PROPS = [
  'nodeName',
  'tagName',
  'localName',
  'prefix',
  'namespaceURI',
  'nodeType',
  'data',
  'value',
  'nodeValue',
  'name',
  'specified',
];

// xmldom internals (_nsMap, own fields) are compared field by field
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyNode = Record<string, any>;

const scalars = (node: AnyNode): string => {
  const result: [string, unknown][] = Object.keys(node)
    .filter((k) => typeof node[k] !== 'object' || node[k] === null)
    .map((k) => [k, node[k]]);
  PROPS.forEach((k) => result.push([`@${k}`, node[k]]));
  return JSON.stringify(result.sort(([a], [b]) => (a < b ? -1 : 1)));
};

const difference = (a: AnyNode, b: AnyNode, at: string): string | undefined => {
  if (scalars(a) !== scalars(b)) return `${at}: ${scalars(a)} / ${scalars(b)}`;
  if (a.ownerDocument !== b.ownerDocument) return `${at}: ownerDocument`;
  if (JSON.stringify(a._nsMap) !== JSON.stringify(b._nsMap)) {
    return `${at}: _nsMap`;
  }
  const attrsA = a.attributes;
  const attrsB = b.attributes;
  if (!!attrsA !== !!attrsB) return `${at}: attributes`;
  if (attrsA) {
    if (attrsA.length !== attrsB.length) return `${at}: attribute count`;
    for (let i = 0; i < attrsA.length; i++) {
      if (attrsA.item(i).ownerElement !== a) return `${at}: ownerElement`;
      if (attrsB.item(i).ownerElement !== b) return `${at}: ownerElement`;
      const d = difference(attrsA.item(i), attrsB.item(i), `${at}@${i}`);
      if (d) return d;
    }
  }
  const childrenA = Array.from(a.childNodes ?? []) as AnyNode[];
  const childrenB = Array.from(b.childNodes ?? []) as AnyNode[];
  if (childrenA.length !== childrenB.length) return `${at}: child count`;
  for (let i = 0; i < childrenA.length; i++) {
    if (childrenA[i].parentNode !== a || childrenB[i].parentNode !== b) {
      return `${at}: parentNode`;
    }
    const d = difference(childrenA[i], childrenB[i], `${at}/${i}`);
    if (d) return d;
  }
  return undefined;
};

const expectSameClone = (node: AnyNode, deep: boolean, at: string) => {
  const native = node.cloneNode(deep);
  const fast = cloneXml(node, deep);
  expect(difference(native, fast, at)).toBeUndefined();
  expect(fast.parentNode).toBeNull();
  const serializer = new XMLSerializer();
  expect(serializer.serializeToString(fast)).toBe(
    serializer.serializeToString(native),
  );
};

test('cloneXml clones the parts of templates like cloneNode', async () => {
  const dir = `${__dirname}/pptx-templates`;
  // text, tables, charts, groups, placeholders
  const files = [
    'TextReplace.pptx',
    'RootTemplate.pptx',
    'ChartBarsStacked.pptx',
    'SlideMasters.pptx',
  ].filter((f) => fs.existsSync(`${dir}/${f}`));
  expect(files.length).toBeGreaterThan(1);

  let compared = 0;
  for (const file of files) {
    const zip = await JSZip.loadAsync(fs.readFileSync(`${dir}/${file}`));
    const parts = Object.keys(zip.files).filter((n) => /\.(xml|rels)$/.test(n));
    for (const part of parts) {
      const doc = new DOMParser().parseFromString(
        await zip.file(part).async('string'),
        'application/xml',
      );
      const elements = Array.from(
        doc.getElementsByTagName('*'),
      ) as unknown as AnyNode[];
      // every element of small parts, an even sample of large ones
      const step = Math.ceil(elements.length / 60);
      const sample = elements.filter((_, i) => i % step === 0);
      for (const element of [doc.documentElement, ...sample]) {
        expectSameClone(element, true, `${file}:${part}:${element.tagName}`);
        expectSameClone(element, false, `${file}:${part}:${element.tagName}`);
        compared++;
      }
    }
  }
  expect(compared).toBeGreaterThan(500);
  // ~3 s alone; the full suite runs it in parallel with everything else
}, 30_000);

test('cloneXml falls back to cloneNode for other nodes', () => {
  const doc = new DOMParser().parseFromString(
    '<a:p xmlns:a="urn:a" xmlns:r="urn:r" r:id="1">' +
      '<!--c--><![CDATA[x < y]]><?pi data?><a:t xml:space="preserve"> t &amp; </a:t></a:p>',
    'application/xml',
  ) as unknown as AnyNode;
  const p = doc.documentElement;
  expectSameClone(p, true, 'p');
  Array.from(p.childNodes as AnyNode[]).forEach((child, i) =>
    expectSameClone(child, true, `child ${i}`),
  );
  expectSameClone(doc.createTextNode('a < b'), false, 'text');
});
