import fs from 'fs';

/**
 * Minimal reader for TrueType/OpenType fonts (.ttf, .otf) and collections
 * (.ttc): just what text measurement needs - names, style, units per em and
 * glyph advance widths. No shaping, no kerning (PowerPoint does not kern by
 * default either).
 */

/** Name and style of one face, read without loading the whole file. */
export type FontFaceInfo = {
  file: string;
  /** Index inside a .ttc collection, else 0. */
  index: number;
  /** Lower-case family names (legacy and typographic). */
  families: string[];
  bold: boolean;
  italic: boolean;
};

type TableDirectory = Map<string, { offset: number; length: number }>;

/** Random access to a font file: partial reads while indexing. */
type Source = {
  read(offset: number, length: number): Buffer;
};

const fileSource = (fd: number): Source => ({
  read(offset, length) {
    const buffer = Buffer.alloc(length);
    const bytes = fs.readSync(fd, buffer, 0, length, offset);
    return bytes === length ? buffer : buffer.subarray(0, bytes);
  },
});

const bufferSource = (data: Buffer): Source => ({
  read: (offset, length) => data.subarray(offset, offset + length),
});

const tag = (b: Buffer, offset: number) =>
  b.toString('latin1', offset, offset + 4);

/** Offsets of the faces in a file: one, or several for a collection. */
const faceOffsets = (source: Source): number[] => {
  const header = source.read(0, 12);
  if (header.length < 12) throw new Error('not a font file');
  if (tag(header, 0) !== 'ttcf') {
    const version = header.readUInt32BE(0);
    if (version !== 0x00010000 && !['OTTO', 'true'].includes(tag(header, 0))) {
      throw new Error('not a font file');
    }
    return [0];
  }
  const count = header.readUInt32BE(8);
  const offsets = source.read(12, count * 4);
  return Array.from({ length: count }, (_, i) => offsets.readUInt32BE(i * 4));
};

const readDirectory = (source: Source, faceOffset: number): TableDirectory => {
  const header = source.read(faceOffset, 12);
  const count = header.readUInt16BE(4);
  const records = source.read(faceOffset + 12, count * 16);
  const tables: TableDirectory = new Map();
  for (let i = 0; i < count; i++) {
    tables.set(tag(records, i * 16), {
      offset: records.readUInt32BE(i * 16 + 8),
      length: records.readUInt32BE(i * 16 + 12),
    });
  }
  return tables;
};

const readTable = (
  source: Source,
  tables: TableDirectory,
  name: string,
): Buffer | undefined => {
  const entry = tables.get(name);
  return entry ? source.read(entry.offset, entry.length) : undefined;
};

/** Family names (IDs 1 and 16) from the `name` table, lower case. */
const familyNames = (name: Buffer | undefined): string[] => {
  if (!name || name.length < 6) return [];
  const count = name.readUInt16BE(2);
  const strings = name.readUInt16BE(4);
  const result = new Set<string>();
  for (let i = 0; i < count; i++) {
    const record = 6 + i * 12;
    if (record + 12 > name.length) break;
    const platform = name.readUInt16BE(record);
    const id = name.readUInt16BE(record + 6);
    if (id !== 1 && id !== 16) continue;
    const length = name.readUInt16BE(record + 8);
    const start = strings + name.readUInt16BE(record + 10);
    const raw = name.subarray(start, start + length);
    let text: string;
    if (platform === 0 || platform === 3) {
      // UTF-16BE
      const swapped = Buffer.from(raw);
      swapped.swap16();
      text = swapped.toString('utf16le');
    } else {
      text = raw.toString('latin1');
    }
    if (text.trim()) result.add(text.trim().toLowerCase());
  }
  return [...result];
};

const styleOf = (
  os2: Buffer | undefined,
  head: Buffer | undefined,
): { bold: boolean; italic: boolean } => {
  if (os2 && os2.length >= 64) {
    const selection = os2.readUInt16BE(62);
    const weight = os2.readUInt16BE(4);
    return {
      bold: (selection & 0x20) !== 0 || weight >= 600,
      italic: (selection & 0x01) !== 0,
    };
  }
  const macStyle = head && head.length >= 46 ? head.readUInt16BE(44) : 0;
  return { bold: (macStyle & 1) !== 0, italic: (macStyle & 2) !== 0 };
};

const describe = (source: Source, file: string): FontFaceInfo[] =>
  faceOffsets(source).map((offset, index) => {
    const tables = readDirectory(source, offset);
    return {
      file,
      index,
      families: familyNames(readTable(source, tables, 'name')),
      ...styleOf(
        readTable(source, tables, 'OS/2'),
        readTable(source, tables, 'head'),
      ),
    };
  });

/** Names and styles of the faces in a font file (partial reads only). */
export const readFontInfo = (file: string): FontFaceInfo[] => {
  const fd = fs.openSync(file, 'r');
  try {
    return describe(fileSource(fd), file);
  } finally {
    fs.closeSync(fd);
  }
};

/** A loaded face: measures text in points. */
export class FontFace {
  readonly unitsPerEm: number;
  private readonly advances: Uint16Array;
  private readonly cmap: (codePoint: number) => number;
  private readonly cache = new Map<number, number>();

  constructor(
    data: Buffer,
    readonly info: FontFaceInfo,
  ) {
    const source = bufferSource(data);
    const tables = readDirectory(source, faceOffsets(source)[info.index] ?? 0);
    const require = (name: string) => {
      const table = readTable(source, tables, name);
      if (!table) throw new Error(`font has no ${name} table`);
      return table;
    };
    const head = require('head');
    const hhea = require('hhea');
    const hmtx = require('hmtx');
    const maxp = require('maxp');
    this.unitsPerEm = head.readUInt16BE(18) || 1000;
    const metrics = hhea.readUInt16BE(34);
    const glyphs = maxp.readUInt16BE(4);
    this.advances = new Uint16Array(Math.max(glyphs, metrics));
    for (let g = 0; g < this.advances.length; g++) {
      const record = Math.min(g, metrics - 1) * 4;
      this.advances[g] =
        record + 2 <= hmtx.length ? hmtx.readUInt16BE(record) : 0;
    }
    this.cmap = parseCmap(require('cmap'));
  }

  /** Advance width of one character in em (1 = font size). */
  advance(codePoint: number): number {
    let width = this.cache.get(codePoint);
    if (width === undefined) {
      const glyph = this.cmap(codePoint);
      width = (this.advances[glyph] ?? 0) / this.unitsPerEm;
      this.cache.set(codePoint, width);
    }
    return width;
  }

  /** Whether the face has a glyph for the character. */
  has(codePoint: number): boolean {
    return this.cmap(codePoint) !== 0;
  }

  static load(info: FontFaceInfo, data?: Buffer): FontFace {
    return new FontFace(data ?? fs.readFileSync(info.file), info);
  }

  /** Faces of a font file given as contents (option `fonts`). */
  static describe(data: Buffer, name: string): FontFaceInfo[] {
    return describe(bufferSource(data), name);
  }
}

/** Character → glyph lookup from a cmap table (formats 4 and 12). */
const parseCmap = (cmap: Buffer): ((codePoint: number) => number) => {
  const count = cmap.readUInt16BE(2);
  let best: { offset: number; rank: number } | undefined;
  for (let i = 0; i < count; i++) {
    const platform = cmap.readUInt16BE(4 + i * 8);
    const encoding = cmap.readUInt16BE(6 + i * 8);
    const offset = cmap.readUInt32BE(8 + i * 8);
    const format = cmap.readUInt16BE(offset);
    let rank = 0;
    if (format === 12 && (platform === 3 || platform === 0)) rank = 3;
    else if (format === 4 && platform === 3 && encoding === 1) rank = 2;
    else if (format === 4 && platform === 0) rank = 1;
    if (rank && (!best || rank > best.rank)) best = { offset, rank };
  }
  if (!best) return () => 0;

  const at = best.offset;
  if (cmap.readUInt16BE(at) === 12) {
    const groups = cmap.readUInt32BE(at + 12);
    return (cp) => {
      let lo = 0;
      let hi = groups - 1;
      while (lo <= hi) {
        const mid = (lo + hi) >> 1;
        const g = at + 16 + mid * 12;
        const start = cmap.readUInt32BE(g);
        const end = cmap.readUInt32BE(g + 4);
        if (cp < start) hi = mid - 1;
        else if (cp > end) lo = mid + 1;
        else return cmap.readUInt32BE(g + 8) + (cp - start);
      }
      return 0;
    };
  }

  // format 4
  const segments = cmap.readUInt16BE(at + 6) / 2;
  const ends = at + 14;
  const starts = ends + segments * 2 + 2;
  const deltas = starts + segments * 2;
  const rangeOffsets = deltas + segments * 2;
  return (cp) => {
    if (cp > 0xffff) return 0;
    for (let s = 0; s < segments; s++) {
      if (cp > cmap.readUInt16BE(ends + s * 2)) continue;
      const start = cmap.readUInt16BE(starts + s * 2);
      if (cp < start) return 0;
      const delta = cmap.readInt16BE(deltas + s * 2);
      const rangeOffset = cmap.readUInt16BE(rangeOffsets + s * 2);
      if (rangeOffset === 0) return (cp + delta) & 0xffff;
      const glyphAt = rangeOffsets + s * 2 + rangeOffset + (cp - start) * 2;
      if (glyphAt + 2 > cmap.length) return 0;
      const glyph = cmap.readUInt16BE(glyphAt);
      return glyph === 0 ? 0 : (glyph + delta) & 0xffff;
    }
    return 0;
  };
};
