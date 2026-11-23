import zlib from 'zlib';
import JSZip from 'jszip';

/**
 * Packs a JSZip into a .pptx with Node's zlib: the job of
 * `zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })`,
 * about four times faster. JSZip deflates in JavaScript (pako), with a
 * fresh compressor per part and a macrotask per chunk; a deck has dozens of
 * small XML parts, so most of that time is overhead.
 *
 * Parts taken over unchanged from the template keep their compressed bytes,
 * as with JSZip; stored pictures and other binary parts also stay stored,
 * and pictures the render adds (JPEG, PNG, GIF) are stored as PowerPoint
 * does (JSZip deflated them, for about 1 % less, at most of the time it
 * took to pack a deck with photos). No ZIP64: a deck needs less
 * than 4 GB and 65535 parts; above that, JSZip packs it.
 */
export const generatePptx = async (zip: JSZip): Promise<Buffer> => {
  const names = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  // PowerPoint expects nothing about the order, but this is the usual one
  const first = names.indexOf('[Content_Types].xml');
  if (first > 0) names.unshift(...names.splice(first, 1));

  const entries: Entry[] = [];
  for (const name of names) entries.push(await entryOf(zip.files[name]));

  const size = entries.reduce(
    (sum, e) => sum + 30 + e.name.length + e.data.length,
    0,
  );
  if (entries.length > 0xffff || size > 0xfffffffe) {
    return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }

  const chunks: Buffer[] = [];
  const central: Buffer[] = [];
  let offset = 0;
  for (const entry of entries) {
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    writeCommon(local, 4, entry);
    chunks.push(local, entry.name, entry.data);

    const header = Buffer.alloc(46);
    header.writeUInt32LE(0x02014b50, 0);
    header.writeUInt16LE(20, 4); // made by: 2.0, MS-DOS
    writeCommon(header, 6, entry);
    // comment length, disk number, internal and external attributes: 0
    header.writeUInt32LE(offset, 42);
    central.push(header, entry.name);

    offset += local.length + entry.name.length + entry.data.length;
  }
  const directory = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, directory, end]);
};

type Entry = {
  name: Buffer;
  /** The bytes as stored. */
  data: Buffer;
  method: 0 | 8;
  crc: number;
  size: number;
  date: Date;
};

/** The fields local and central headers share, from `at` on. */
const writeCommon = (buffer: Buffer, at: number, entry: Entry): void => {
  buffer.writeUInt16LE(20, at); // version needed: 2.0
  buffer.writeUInt16LE(0x0800, at + 2); // names are UTF-8
  buffer.writeUInt16LE(entry.method, at + 4);
  const { date } = entry;
  buffer.writeUInt16LE(
    (date.getUTCHours() << 11) |
      (date.getUTCMinutes() << 5) |
      (date.getUTCSeconds() >> 1),
    at + 6,
  );
  buffer.writeUInt16LE(
    (Math.max(date.getUTCFullYear() - 1980, 0) << 9) |
      ((date.getUTCMonth() + 1) << 5) |
      date.getUTCDate(),
    at + 8,
  );
  buffer.writeUInt32LE(entry.crc >>> 0, at + 10);
  buffer.writeUInt32LE(entry.data.length, at + 14);
  buffer.writeUInt32LE(entry.size, at + 18);
  buffer.writeUInt16LE(entry.name.length, at + 22);
  buffer.writeUInt16LE(0, at + 24); // extra field length
};

/** JSZip's view of a part loaded from a file and not changed since. */
type Loaded = {
  compressedSize: number;
  uncompressedSize: number;
  crc32: number;
  compression: { magic: string };
  compressedContent: Uint8Array;
};

const isLoaded = (data: unknown): data is Loaded =>
  !!data &&
  typeof data === 'object' &&
  'compressedContent' in data &&
  'compression' in data &&
  (data as Loaded).compressedContent instanceof Uint8Array &&
  (data as Loaded).compressedContent.length === (data as Loaded).compressedSize;

const DEFLATE = '\x08\x00';
const STORE = '\x00\x00';

/** Parts that deflate well. */
const XML = /\.(xml|rels|vml)$/i;

/** Formats that are compressed already: deflate hardly shrinks them. */
const COMPRESSED = /\.(jpe?g|png|gif|webp|mp4|m4v|mov|mp3|m4a|zip)$/i;

const bufferOf = (data: Uint8Array): Buffer =>
  Buffer.from(data.buffer, data.byteOffset, data.byteLength);

/** Whether a string part holds bytes (one per character), not text. */
const binary = (file: JSZip.JSZipObject): boolean =>
  !!(file as unknown as { _dataBinary?: boolean })._dataBinary;

const entryOf = async (file: JSZip.JSZipObject): Promise<Entry> => {
  const name = Buffer.from(file.name, 'utf8');
  const date = file.date ?? new Date();
  // internal, but what JSZip itself reuses for unchanged parts
  const data = (file as unknown as { _data: unknown })._data;
  // Deflated parts as they are. Stored media too: PowerPoint stores
  // pictures and videos, which deflate hardly shrinks
  if (
    isLoaded(data) &&
    (data.compression.magic === DEFLATE ||
      (data.compression.magic === STORE && !XML.test(file.name)))
  ) {
    return {
      name,
      data: bufferOf(data.compressedContent),
      method: data.compression.magic === DEFLATE ? 8 : 0,
      crc: data.crc32,
      size: data.uncompressedSize,
      date,
    };
  }
  // other stored parts are deflated, as JSZip does with 'DEFLATE'
  const content =
    isLoaded(data) && data.compression.magic === STORE
      ? bufferOf(data.compressedContent)
      : typeof data === 'string' && !binary(file)
        ? Buffer.from(data, 'utf8')
        : await file.async('nodebuffer');
  // pictures and videos added by the render (IMAGE): already compressed
  const stored = COMPRESSED.test(file.name);
  return {
    name,
    data: stored ? content : zlib.deflateRawSync(content),
    method: stored ? 0 : 8,
    crc: crc32(content),
    size: content.length,
    date,
  };
};

let table: Int32Array | undefined;

/** zlib.crc32 where Node has it (20.15, 22.2), else a table. */
const crc32 = (data: Buffer): number => {
  const native = (zlib as { crc32?: (data: Buffer) => number }).crc32;
  if (native) return native(data);
  table ??= Int32Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c;
  });
  let crc = -1;
  for (let i = 0; i < data.length; i++) {
    crc = table[(crc ^ data[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ -1) >>> 0;
};
