import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import {
  elementsByTagName,
  firstElementByTagName,
} from '../src/helper/xml-find';
import { XmlDocument, XmlElement } from '../src/types/xml-types';

/** elementsByTagName stands in for getElementsByTagName: same nodes, order. */

const TAGS = ['*', 'a:t', 'a:p', 'p:ph', 'p:cNvPr', 'Relationship', 'nope'];

const expectSame = (root: XmlDocument | XmlElement, at: string) => {
  for (const tag of TAGS) {
    const expected = Array.from(root.getElementsByTagName(tag));
    const actual = elementsByTagName(root, tag);
    expect([at, tag, actual.length]).toEqual([at, tag, expected.length]);
    expect(actual.every((el, i) => el === expected[i])).toBe(true);
    expect(firstElementByTagName(root, tag)).toBe(expected[0]);
  }
};

test('elementsByTagName finds what getElementsByTagName finds', async () => {
  const dir = `${__dirname}/pptx-templates`;
  let compared = 0;
  for (const file of fs.readdirSync(dir).filter((f) => f.endsWith('.pptx'))) {
    const zip = await JSZip.loadAsync(fs.readFileSync(`${dir}/${file}`));
    for (const part of Object.keys(zip.files)) {
      if (!/(slide\d+|slideLayout\d+)\.xml(\.rels)?$/.test(part)) continue;
      const doc = new DOMParser().parseFromString(
        await zip.file(part).async('string'),
        'application/xml',
      ) as unknown as XmlDocument;
      expectSame(doc, `${file}:${part}`);
      const elements = Array.from(doc.getElementsByTagName('*'));
      elements
        .filter((_, i) => i % 25 === 0)
        .forEach((el) => expectSame(el, `${file}:${part}:${el.tagName}`));
      compared++;
    }
  }
  expect(compared).toBeGreaterThan(100);
}, 60_000);

test('elementsByTagName: leaves, detached elements, HTML', () => {
  const doc = new DOMParser().parseFromString(
    '<r><a><b/><b><b/></b></a><b/>text</r>',
    'application/xml',
  ) as unknown as XmlDocument;
  expectSame(doc, 'doc');
  const a = doc.getElementsByTagName('a')[0];
  a.parentNode.removeChild(a);
  expectSame(a, 'detached');
  expectSame(doc.getElementsByTagName('b')[0], 'leaf');
  const html = new DOMParser().parseFromString(
    '<html><body><P>x</P><p>y</p></body></html>',
    'text/html',
  ) as unknown as XmlDocument;
  for (const tag of ['p', 'P', '*']) {
    expect(elementsByTagName(html, tag)).toEqual(
      Array.from(html.getElementsByTagName(tag)),
    );
  }
});
