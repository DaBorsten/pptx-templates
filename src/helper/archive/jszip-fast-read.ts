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
 *
 * Inflating never produces more than the size the zip header gives: a
 * part that would (a corrupt zip, or a zip bomb with a false size) is
 * rejected as JSZip rejects it, but before it fills the memory. Parts above
 * SYNC_INFLATE_MAX are inflated on the thread pool (readFastAsync), so a
 * large part does not block the event loop.
 */

type Loaded = {
  compressedSize: number;
  uncompressedSize: number;
  compression: { magic: string };
  compressedContent: Uint8Array;
};

const STORE = '\x00\x00';
const DEFLATE = '\x08\x00';

/** Uncompressed size up to which a part is inflated synchronously. */
export const SYNC_INFLATE_MAX = 1024 * 1024;

/** A loaded, unchanged part as the zip holds it. */
type Raw = {
  name: string;
  /** A view of the loaded zip: never handed out without a copy. */
  bytes: Buffer;
  stored: boolean;
  /** Uncompressed size from the zip header. */
  size: number;
};

const loadedRaw = (file: JSZip.JSZipObject): Raw | undefined => {
  // internal, but what JSZip itself reads the part from
  const data = (file as unknown as { _data: unknown })._data as Loaded;
  const content = data?.compressedContent;
  const magic = data?.compression?.magic;
  const size = data?.uncompressedSize;
  if (
    !(content instanceof Uint8Array) ||
    content.length !== data.compressedSize ||
    (magic !== STORE && magic !== DEFLATE) ||
    !Number.isSafeInteger(size) ||
    size < 0
  ) {
    return undefined;
  }
  const bytes = Buffer.from(
    content.buffer,
    content.byteOffset,
    content.byteLength,
  );
  if (magic === STORE && bytes.length !== size) return undefined;
  return { name: file.name, bytes, stored: magic === STORE, size };
};

/** The error JSZip would end with, raised before inflating any further. */
const tooLarge = (raw: Raw, cause: unknown): Error =>
  new Error(
    `${raw.name}: inflates to more than the ${raw.size} bytes the zip ` +
      'header gives (corrupt file or zip bomb)',
    { cause },
  );

/**
 * What to do with an inflate error: past the size of the header → reject;
 * anything else (bad data, a size zlib cannot allocate) → undefined, JSZip
 * reads the part and reports it as before.
 */
const inflateFailed = (raw: Raw, error: unknown): undefined => {
  if ((error as { code?: string })?.code === 'ERR_BUFFER_TOO_LARGE') {
    throw tooLarge(raw, error);
  }
  return undefined;
};

// maxOutputLength must be at least 1; an empty part inflates to 0 bytes
const inflateOptions = (raw: Raw): zlib.ZlibOptions => ({
  maxOutputLength: Math.max(raw.size, 1),
});

const inflateSync = (raw: Raw): Buffer | undefined => {
  let inflated: Buffer;
  try {
    inflated = zlib.inflateRawSync(raw.bytes, inflateOptions(raw));
  } catch (error) {
    return inflateFailed(raw, error);
  }
  return inflated.length === raw.size ? inflated : undefined;
};

const inflateAsync = (raw: Raw): Promise<Buffer | undefined> =>
  new Promise((resolve, reject) => {
    const done = (error: Error | null, inflated?: Buffer) => {
      try {
        if (error) resolve(inflateFailed(raw, error));
        else resolve(inflated.length === raw.size ? inflated : undefined);
      } catch (rejected) {
        reject(rejected);
      }
    };
    try {
      zlib.inflateRaw(raw.bytes, inflateOptions(raw), done);
    } catch (error) {
      // options zlib refuses (a size it cannot allocate)
      done(error as Error);
    }
  });

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

/** A written part as `type`, or undefined: JSZip converts it. */
const fromWritten = (
  written: string | Buffer,
  type: 'string' | 'nodebuffer',
  copy: boolean,
): string | Buffer | undefined => {
  if (typeof written === 'string') {
    if (type === 'string') return written;
    return SURROGATE.test(written) ? undefined : Buffer.from(written, 'utf-8');
  }
  if (type === 'string') return toText(written);
  return copy ? Buffer.from(written) : written;
};

/** Inflated or stored bytes as `type`. */
const fromBytes = (
  raw: Raw,
  bytes: Buffer | undefined,
  type: 'string' | 'nodebuffer',
  copy: boolean,
): string | Buffer | undefined => {
  if (!bytes) return undefined;
  if (type === 'string') return toText(bytes); // a new string anyway
  // stored bytes are a view of the loaded zip; JSZip returns a copy the
  // caller may keep or change
  return raw.stored && copy ? Buffer.from(bytes) : bytes;
};

/**
 * The contents of `file` as JSZip's `async(type)` returns them, or
 * undefined: read it with JSZip (or with readFastAsync, for a deflated
 * part above SYNC_INFLATE_MAX). `copy: false` may return the written
 * buffer or a view of the zip itself, for a caller that only reads it.
 * Throws for a part that inflates to more than its header says.
 */
export const readFast = (
  file: JSZip.JSZipObject,
  type: 'string' | 'nodebuffer',
  copy = true,
): string | Buffer | undefined => {
  const written = writtenData.get(file);
  if (written !== undefined) return fromWritten(written, type, copy);
  const raw = loadedRaw(file);
  if (!raw) return undefined;
  if (raw.stored) return fromBytes(raw, raw.bytes, type, copy);
  if (raw.size > SYNC_INFLATE_MAX) return undefined;
  return fromBytes(raw, inflateSync(raw), type, copy);
};

/**
 * readFast, with large deflated parts inflated on the thread pool instead
 * of leaving them to JSZip. Rejects for a part that inflates to more than
 * its header says.
 */
export const readFastAsync = async (
  file: JSZip.JSZipObject,
  type: 'string' | 'nodebuffer',
  copy = true,
): Promise<string | Buffer | undefined> => {
  const written = writtenData.get(file);
  if (written !== undefined) return fromWritten(written, type, copy);
  const raw = loadedRaw(file);
  if (!raw || raw.stored || raw.size <= SYNC_INFLATE_MAX) {
    return readFast(file, type, copy);
  }
  return fromBytes(raw, await inflateAsync(raw), type, copy);
};
