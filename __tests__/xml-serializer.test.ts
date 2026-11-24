import fs from 'fs';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { serializeXml } from '../src/helper/xml-serializer';

/**
 * serializeXml replaces xmldom's XMLSerializer when parts are written; it
 * must produce the same string for every node.
 */

const xmldom = (node: unknown) =>
  new XMLSerializer().serializeToString(node as never);

const expectSame = (node: unknown, at: string) => {
  const expected = xmldom(node);
  const actual = serializeXml(node);
  if (actual !== expected) {
    let i = 0;
    while (actual[i] === expected[i]) i++;
    throw new Error(
      `${at}: differs at ${i}\n  xmldom: ${expected.slice(i - 80, i + 80)}\n  fast:   ${actual.slice(i - 80, i + 80)}`,
    );
  }
};

const parse = (xml: string, type = 'application/xml') =>
  new DOMParser().parseFromString(xml, type);

test('serializeXml writes every part of the templates like XMLSerializer', async () => {
  const dir = `${__dirname}/pptx-templates`;
  let compared = 0;
  for (const file of fs.readdirSync(dir).filter((f) => f.endsWith('.pptx'))) {
    const zip = await JSZip.loadAsync(fs.readFileSync(`${dir}/${file}`));
    const parts = Object.keys(zip.files).filter((n) => /\.(xml|rels)$/.test(n));
    for (const part of parts) {
      const doc = parse(await zip.file(part).async('string'));
      expectSame(doc, `${file}:${part}`);
      // elements by themselves: their namespaces are declared higher up
      const elements = Array.from(doc.getElementsByTagName('*'));
      const step = Math.ceil(elements.length / 20);
      elements
        .filter((_, i) => i % step === 0)
        .forEach((el) => expectSame(el, `${file}:${part}:${el.tagName}`));
      compared++;
    }
  }
  expect(compared).toBeGreaterThan(1000);
}, 60_000);

test('serializeXml: namespaces, escaping and other node types', () => {
  const doc = parse(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<!DOCTYPE r SYSTEM "r.dtd">' +
      '<r xmlns="urn:default" xmlns:a="urn:a" a:x="1 &amp; 2 &lt; &quot;3&quot; &#9;&#10;&#13;">' +
      '<a:t xml:space="preserve"> x &lt; y &amp; z &gt; w ]]&gt; </a:t>' +
      '<inner xmlns="urn:other"><deep/></inner>' +
      '<reset xmlns=""><plain/></reset>' +
      '<!-- comment --><![CDATA[ a ]]> b ]]><?pi some data?>' +
      '<a:empty/><b:x xmlns:b="urn:b" b:y="2"/>' +
      '</r>',
  );
  expectSame(doc, 'document');
  Array.from(doc.getElementsByTagName('*')).forEach((el) =>
    expectSame(el, el.tagName),
  );
  Array.from(doc.childNodes).forEach((n, i) => expectSame(n, `child ${i}`));

  // nodes made by code: without namespace, with an undeclared one, moved
  const root = doc.documentElement;
  const noNs = doc.createElement('a:rPr');
  noNs.setAttribute('lang', 'de-DE');
  root.appendChild(noNs);
  const undeclared = doc.createElementNS('urn:new', 'n:thing');
  undeclared.setAttributeNS('urn:attr', 'q:attr', 'v');
  undeclared.appendChild(doc.createTextNode('text\r\nwith "quotes" & <tags>'));
  root.appendChild(undeclared);
  const defaultNs = doc.createElementNS('urn:default', 'child');
  root.appendChild(defaultNs);
  const otherDefault = doc.createElementNS('urn:a', 'unprefixed');
  root.appendChild(otherDefault);
  const fragment = doc.createDocumentFragment();
  fragment.appendChild(doc.createElement('f'));
  fragment.appendChild(doc.createTextNode('&'));
  expectSame(doc, 'document with created nodes');
  [noNs, undeclared, defaultNs, otherDefault, fragment].forEach((n, i) =>
    expectSame(n, `created ${i}`),
  );
  expectSame(root.getAttributeNode('a:x'), 'attribute');
  expectSame(doc.createTextNode('a < b'), 'detached text');

  // an element whose namespace is in no declaration at all
  const bare = parse('<x:root xmlns:x="urn:x"><x:a/></x:root>');
  const detached = bare.documentElement.firstChild;
  bare.documentElement.removeChild(detached);
  expectSame(detached, 'detached element');
});

test('serializeXml hands HTML to xmldom', () => {
  const html = parse(
    '<html><body><br/><p>a &amp; b</p></body></html>',
    'text/html',
  );
  expectSame(html, 'html document');
  const xhtml = parse(
    '<html xmlns="http://www.w3.org/1999/xhtml"><body><br/><p/></body></html>',
  );
  expectSame(xhtml, 'xhtml document');
});
