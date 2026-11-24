import fs from 'fs';
import os from 'os';
import path from 'path';
import { log } from '../../helper/logger';
import { FontFace, FontFaceInfo, readFontInfo } from './font-file';

const FONT_FILE = /\.(ttf|otf|ttc)$/i;

/**
 * Fonts with the same metrics under another name. A server without the
 * Microsoft fonts usually has these (LibreOffice ships them), and text set in
 * them breaks exactly like in the original.
 */
const METRIC_ALIASES: Record<string, string[]> = {
  arial: ['liberation sans', 'arimo'],
  helvetica: ['liberation sans', 'arimo', 'arial'],
  'times new roman': ['liberation serif', 'tinos'],
  times: ['liberation serif', 'tinos', 'times new roman'],
  'courier new': ['liberation mono', 'cousine'],
  courier: ['liberation mono', 'cousine', 'courier new'],
  calibri: ['carlito'],
  cambria: ['caladea'],
  'liberation sans': ['arial', 'arimo'],
  'liberation serif': ['times new roman', 'tinos'],
  'liberation mono': ['courier new', 'cousine'],
  carlito: ['calibri'],
  caladea: ['cambria'],
};

/** Used, with a warning, when a font is not installed at all. */
const FALLBACK = [
  'liberation sans',
  'arial',
  'arimo',
  'helvetica',
  'dejavu sans',
  'noto sans',
];

export const systemFontFolders = (): string[] => {
  const home = os.homedir();
  switch (process.platform) {
    case 'win32':
      return [
        path.join(process.env.WINDIR ?? 'C:\\Windows', 'Fonts'),
        path.join(
          process.env.LOCALAPPDATA ?? '',
          'Microsoft',
          'Windows',
          'Fonts',
        ),
      ];
    case 'darwin':
      return [
        '/System/Library/Fonts',
        '/Library/Fonts',
        path.join(home, 'Library', 'Fonts'),
      ];
    default:
      return [
        '/usr/share/fonts',
        '/usr/local/share/fonts',
        path.join(home, '.fonts'),
        path.join(home, '.local', 'share', 'fonts'),
      ];
  }
};

const listFontFiles = (entry: string, depth = 0): string[] => {
  let stat: fs.Stats;
  try {
    stat = fs.statSync(entry);
  } catch {
    return [];
  }
  if (stat.isFile()) return FONT_FILE.test(entry) ? [entry] : [];
  if (!stat.isDirectory() || depth > 6) return [];
  let names: string[];
  try {
    names = fs.readdirSync(entry).sort();
  } catch {
    return [];
  }
  return names.flatMap((name) =>
    listFontFiles(path.join(entry, name), depth + 1),
  );
};

/*
 * Font files never change while a process runs, so their names and parsed
 * metrics are cached for the whole process. This is a read-only cache of
 * file contents, not state of an Automizer instance.
 */
const infoCache = new Map<string, FontFaceInfo[]>();
const faceCache = new Map<string, FontFace>();
/**
 * The font files of the system font folders. Listing them takes a stat()
 * per file (hundreds on Windows) and every render with flow() may need
 * them, so they are listed once per process. Fonts installed while the
 * process runs are found after a restart, as by most applications.
 *
 * The `fonts` folders a caller passes are listed by each FontRegistry
 * (each render) instead: a caller may add font files to them at any time.
 */
let systemFontFiles: string[] | undefined;

const listSystemFontFiles = (): string[] => {
  systemFontFiles ??= systemFontFolders().flatMap((entry) =>
    listFontFiles(entry),
  );
  return systemFontFiles;
};

const infoOf = (file: string): FontFaceInfo[] => {
  let info = infoCache.get(file);
  if (!info) {
    try {
      info = readFontInfo(file);
    } catch {
      info = []; // not a readable font: skip
    }
    infoCache.set(file, info);
  }
  return info;
};

const faceOf = (info: FontFaceInfo): FontFace => {
  const key = `${info.file}#${info.index}`;
  let face = faceCache.get(key);
  if (!face) {
    face = FontFace.load(info);
    faceCache.set(key, face);
  }
  return face;
};

/** Finds the font file for a typeface name, bold and italic. */
export class FontRegistry {
  private ownIndex?: FontFaceInfo[];
  private index?: FontFaceInfo[];
  private readonly resolved = new Map<string, FontFace>();
  private readonly warned = new Set<string>();

  /**
   * @param folders Font files or folders, searched in this order. The
   *   system font folders are appended unless `system` is false.
   */
  constructor(
    private readonly folders: string[] = [],
    private readonly system = true,
  ) {}

  /** The faces of `folders` only, listed once per registry. */
  private ownFaces(): FontFaceInfo[] {
    this.ownIndex ??= this.folders
      .flatMap((entry) => listFontFiles(entry))
      .flatMap(infoOf);
    return this.ownIndex;
  }

  /** The faces of `folders`, then those of the system font folders. */
  private faces(): FontFaceInfo[] {
    this.index ??= this.system
      ? [...this.ownFaces(), ...listSystemFontFiles().flatMap(infoOf)]
      : this.ownFaces();
    return this.index;
  }

  private find(
    family: string,
    bold: boolean,
    italic: boolean,
  ): FontFaceInfo | undefined {
    const score = (f: FontFaceInfo) =>
      (f.bold === bold ? 2 : 0) + (f.italic === italic ? 1 : 0);
    const best = (faces: FontFaceInfo[]) => {
      const candidates = faces.filter((f) => f.families.includes(family));
      if (!candidates.length) return undefined;
      return candidates.reduce((b, f) => (score(f) > score(b) ? f : b));
    };
    // The system fonts come after `folders` and only replace a face that
    // does not match exactly: reading them (hundreds of files on Windows)
    // is needed only then
    const own = best(this.ownFaces());
    if (own && score(own) === 3) return own;
    return best(this.faces());
  }

  /** The face to measure a typeface with; never undefined (falls back). */
  face(typeface: string, bold = false, italic = false): FontFace {
    const family = typeface.trim().toLowerCase();
    const key = `${family}|${bold}|${italic}`;
    const cached = this.resolved.get(key);
    if (cached) return cached;

    let info = this.find(family, bold, italic);
    for (const alias of METRIC_ALIASES[family] ?? []) {
      if (info) break;
      info = this.find(alias, bold, italic);
    }
    if (!info) {
      for (const fallback of FALLBACK) {
        info = this.find(fallback, bold, italic);
        if (info) break;
      }
      info ??= this.ownFaces()[0] ?? this.faces()[0];
      if (!info) {
        throw new Error(
          `flow(): no font to measure "${typeface}" with. Pass the font ` +
            'files or their folder in the `fonts` option.',
        );
      }
      if (!this.warned.has(family)) {
        this.warned.add(family);
        log.warn(
          `flow(): font "${typeface}" not found, measuring with ` +
            `"${info.families[0]}" instead. Pass its file in the \`fonts\` ` +
            'option for exact results.',
        );
      }
    }
    const face = faceOf(info);
    this.resolved.set(key, face);
    return face;
  }
}
