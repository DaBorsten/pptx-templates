import fs from 'fs';
import JSZip from 'jszip';
import ArchiveJszip from '../src/helper/archive/archive-jszip';
import { readFast } from '../src/helper/archive/jszip-fast-read';

/**
 * ArchiveJszip.read() reads loaded and written parts without JSZip's
 * pipeline (readFast); it must return what `file.async(type)` returns.
 */

const same = (a: string | Buffer, b: string | Buffer) =>
  typeof a === 'string' ? a === b : Buffer.isBuffer(b) && a.equals(b);

test('read() returns what JSZip returns, for every part of the templates', async () => {
  const dir = `${__dirname}/pptx-templates`;
  let fast = 0;
  for (const name of fs.readdirSync(dir).filter((f) => f.endsWith('.pptx'))) {
    const content = fs.readFileSync(`${dir}/${name}`);
    const archive = new ArchiveJszip(content, { mode: 'jszip' });
    const zip = await JSZip.loadAsync(content);
    for (const part of Object.keys(zip.files)) {
      if (zip.files[part].dir) continue;
      for (const type of ['string', 'nodebuffer'] as const) {
        const expected = await zip.files[part].async(type);
        const actual = await archive.read(part, type);
        expect(same(expected, actual)).toBe(true);
        if (readFast(zip.files[part], type) !== undefined) fast++;
      }
    }
  }
  // the fast way is taken, not only the fallback
  expect(fast).toBeGreaterThan(1000);
}, 60_000);

test('read() of written parts, odd text and replaced parts', async () => {
  const zip = new JSZip();
  const texts = [
    '<a>plain</a>',
    '﻿<a>with BOM</a>',
    '<a>Umlaute äöü € 😀 and more</a>',
    // a surrogate pair across JSZip's 16 KB chunks
    'x'.repeat(16 * 1024 - 1) + '😀' + 'y'.repeat(100),
    '\uD800 lone surrogate',
    '',
  ];
  const buffers = [
    Buffer.from('<a>bytes äöü</a>'),
    Buffer.from([0x3c, 0x61, 0xff, 0xfe, 0x3e]), // not UTF-8
    Buffer.concat([Buffer.alloc(16 * 1024 - 1, 0x61), Buffer.from('😀')]),
    Buffer.alloc(0),
  ];
  const content = await zip.generateAsync({ type: 'nodebuffer' });
  const archive = new ArchiveJszip(content, { mode: 'jszip' });
  const reference = new JSZip();
  const data = [...texts, ...buffers];
  for (let i = 0; i < data.length; i++) {
    await archive.write(`part${i}.xml`, data[i]);
    reference.file(`part${i}.xml`, data[i]);
  }
  for (let i = 0; i < data.length; i++) {
    for (const type of ['string', 'nodebuffer'] as const) {
      const expected = await reference.file(`part${i}.xml`).async(type);
      const actual = await archive.read(`part${i}.xml`, type);
      expect([i, type, same(expected, actual)]).toEqual([i, type, true]);
    }
  }

  // replaced without write(): the data given to write() is outdated
  archive.archive.file('part0.xml', '<b>replaced</b>');
  expect(await archive.read('part0.xml', 'string')).toBe('<b>replaced</b>');

  // the returned buffer is a copy
  const buffer = (await archive.read('part6.xml', 'nodebuffer')) as Buffer;
  buffer[0] = 0;
  expect(((await archive.read('part6.xml', 'nodebuffer')) as Buffer)[0]).toBe(
    0x3c,
  );
});
