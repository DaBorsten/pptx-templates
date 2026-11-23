import zlib from 'zlib';
import JSZip from 'jszip';

/** JSZip's view of a part loaded from a file and not changed since. */
type Loaded = {
  compressedSize: number;
  uncompressedSize: number;
  compression: { magic: string };
  compressedContent: Uint8Array;
};

/**
 * The text of a part of `zip`, or undefined if there is none: what
 * `zip.file(name)?.async('string')` returns. A part loaded from the file
 * and unchanged since is inflated with Node's zlib right away; JSZip
 * inflates in JavaScript and waits a macrotask per part, which adds up
 * for the many small parts renderDeck reads.
 */
export const readText = async (
  zip: JSZip,
  name: string,
): Promise<string | undefined> => {
  const file = zip.file(name);
  if (!file) return undefined;
  // internal, but what JSZip itself reads the part from
  const data = (file as unknown as { _data: unknown })._data as Loaded;
  const content = data?.compressedContent;
  if (
    content instanceof Uint8Array &&
    content.length === data.compressedSize &&
    typeof data.compression?.magic === 'string'
  ) {
    const bytes = Buffer.from(
      content.buffer,
      content.byteOffset,
      content.byteLength,
    );
    if (data.compression.magic === '\x00\x00') {
      return utf8(bytes);
    }
    if (data.compression.magic === '\x08\x00') {
      const inflated = zlib.inflateRawSync(bytes);
      if (inflated.length === data.uncompressedSize) return utf8(inflated);
    }
  }
  if (typeof data === 'string') {
    // set as text (zip.file(name, string)); JSZip returns it as it is
    if (!(file as unknown as { _dataBinary?: boolean })._dataBinary) {
      return data;
    }
  }
  return file.async('string');
};

/** UTF-8, a byte order mark kept as JSZip keeps it. */
const utf8 = (bytes: Buffer): string => bytes.toString('utf8');
