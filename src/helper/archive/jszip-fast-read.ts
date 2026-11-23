import zlib from 'zlib';
import JSZip from 'jszip';

/**
 * `file.async('string' | 'nodebuffer')` of JSZip, with the same result,
 * without JSZip's pipeline: JSZip inflates in JavaScript (pako) and
 * passes every chunk of 16 KB through a worker chain with a macrotask per
 * chunk, which made reading and packing parts a large share of a run.
 *
 * Two cases are read directly, everything else goes to JSZip:
 * - a part loaded from the file and unchanged since: its compressed bytes,
 *   inflated with Node's zlib (stored or deflated, sizes checked);
 * - a part last written through ArchiveJszip.write(): the data given
 *   (rememberWritten). A part replaced with `zip.file()` since is a new
 *   object, which has none.
 *
 * In Node, JSZip converts with `Buffer#toString('utf-8')` and
 * `Buffer.from(string, 'utf-8')`, but chunk by chunk. That only matters
 * for bytes that are not valid UTF-8 and for strings with surrogates; both
 * go to JSZip (a decoded U+FFFD, or any surrogate when encoding).
 */

type Loaded = {
  compressedSize: number;
  uncompressedSize: number;
  compression: { magic: string };
  compressedContent: Uint8Array;
};

const STORE = '\x00\x00';
const DEFLATE = '\x08\x00';

/**
 * The bytes of a loaded, unchanged part, or undefined. `copy: false` may
 * return a view of the loaded zip, for a caller that only reads it.
 */
const loadedBytes = (
  file: JSZip.JSZipObject,
  copy: boolean,
): Buffer | undefined => {
  // internal, but what JSZip itself reads the part from
  const data = (file as unknown as { _data: unknown })._data as Loaded;
  const content = data?.compressedContent;
  if (
    !(content instanceof Uint8Array) ||
    content.length !== data.compressedSize ||
    typeof data.compression?.magic !== 'string'
  ) {
    return undefined;
  }
  const bytes = Buffer.from(
    content.buffer,
    content.byteOffset,
    content.byteLength,
  );
  if (data.compression.magic === STORE) {
    // a copy, as JSZip returns: the caller may keep or change it
    if (bytes.length !== data.uncompressedSize) return undefined;
    return copy ? Buffer.from(bytes) : bytes;
  }
  if (data.compression.magic === DEFLATE) {
    const inflated = zlib.inflateRawSync(bytes);
    return inflated.length === data.uncompressedSize ? inflated : undefined;
  }
  return undefined;
};

/**
 * The data given for a zip object, keyed by the object: nothing shared
 * between archives or Automizer instances, and gone with the object.
 */
const writtenData = new WeakMap<JSZip.JSZipObject, string | Buffer>();

/** Called by ArchiveJszip.write() with the object `zip.file()` made. */
export const rememberWritten = (
  file: JSZip.JSZipObject | undefined,
  data: string | Buffer,
): void => {
  if (file && (typeof data === 'string' || Buffer.isBuffer(data))) {
    writtenData.set(file, data);
  }
};

const SURROGATE = /[\uD800-\uDFFF]/;

const toText = (bytes: Buffer): string | undefined => {
  const text = bytes.toString('utf-8');
  return text.includes('�') ? undefined : text;
};

/**
 * The contents of `file` as JSZip's `async(type)` returns them, or
 * undefined: read it with JSZip. `copy: false` may return the written
 * buffer itself, for a caller that only reads it.
 */
export const readFast = (
  file: JSZip.JSZipObject,
  type: 'string' | 'nodebuffer',
  copy = true,
): string | Buffer | undefined => {
  const written = writtenData.get(file);
  if (written !== undefined) {
    if (typeof written === 'string') {
      if (type === 'string') return written;
      return SURROGATE.test(written)
        ? undefined
        : Buffer.from(written, 'utf-8');
    }
    if (type === 'string') return toText(written);
    return copy ? Buffer.from(written) : written;
  }
  // decoded to a new string anyway: no copy needed for 'string'
  const bytes = loadedBytes(file, type === 'nodebuffer' && copy);
  if (!bytes) return undefined;
  return type === 'nodebuffer' ? bytes : toText(bytes);
};
