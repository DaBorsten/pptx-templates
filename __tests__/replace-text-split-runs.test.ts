import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { modify } from '../src/index';
import { XmlElement } from '../src/types/xml-types';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

const run = (text: string, attrs = '') =>
  `<a:r><a:rPr lang="de-DE" ${attrs}/><a:t>${text}</a:t></a:r>`;

function shape(paragraphXml: string): XmlElement {
  const doc = new DOMParser().parseFromString(
    `<p:sp ${NS}><p:txBody><a:p>${paragraphXml}</a:p></p:txBody></p:sp>`,
    'text/xml',
  );
  return doc.documentElement as unknown as XmlElement;
}

function replace(el: XmlElement, replaceBy: string) {
  modify.replaceText({ replace: 'name', by: { text: replaceBy } })(el);
  return new XMLSerializer().serializeToString(el as any);
}

const texts = (xml: string) =>
  [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]);

test('tag split into several runs by the spell checker is replaced', () => {
  const el = shape(
    run('Hallo {{') +
      run('name', 'err="1"') +
      run('}}, willkommen'),
  );
  const xml = replace(el, 'Anna');

  expect(texts(xml).join('')).toBe('Hallo Anna, willkommen');
  expect(xml).not.toMatch(/\{\{|\}\}/);
});

test('opening and closing tag characters in separate runs', () => {
  const el = shape(
    run('{') + run('{') + run('name', 'err="1"') + run('}') + run('}'),
  );
  expect(texts(replace(el, 'Anna')).join('')).toBe('Anna');
});

test('runs without a split tag keep their own formatting', () => {
  const el = shape(
    run('fett ', 'b="1"') + run('{{name}}') + run(' normal', 'i="1"'),
  );
  const xml = replace(el, 'Anna');

  expect(texts(xml)).toEqual(['fett ', 'Anna', ' normal']);
  expect(xml).toContain('b="1"');
  expect(xml).toContain('i="1"');
});

test('text around a split tag keeps the formatting of its own run', () => {
  const el = shape(
    run('fett {{', 'b="1"') + run('name', 'err="1"') + run('}} normal', 'i="1"'),
  );
  const xml = replace(el, 'Anna');
  const runs = [...xml.matchAll(/<a:r>(.*?)<\/a:r>/g)].map((m) => m[1]);

  expect(texts(xml)).toEqual(['fett ', 'Anna', ' normal']);
  expect(runs[0]).toContain('b="1"');
  expect(runs[1]).toContain('b="1"');
  expect(runs[2]).toContain('i="1"');
  expect(runs[2]).not.toContain('b="1"');
});
