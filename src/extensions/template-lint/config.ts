import fs from 'fs';
import path from 'path';
import { nameKey, WatchTarget } from './watch';

/** Default config file name, looked up in the current folder. */
export const CONFIG_FILE = 'pptx-templates.json';

/**
 * `pptx-templates.json`: which decks to check, each with its data.
 *
 * ```json
 * { "templates": [
 *   { "files": "{folder}/first_page_{lang}.pptx", "data": "data/first_page_{lang}.json" }
 * ] }
 * ```
 */
export type TemplatesConfig = {
  templates: TemplatesConfigEntry[];
  openingTag?: string;
  closingTag?: string;
  fixSmartQuotes?: boolean;
};

export type TemplatesConfigEntry = {
  /**
   * Pattern(s) relative to the config file. `*` matches any part of a
   * name, `**` any number of folders, `{name}` like `*` and remembers the
   * part for `data` and `schema`.
   */
  files: string | string[];
  /** Example data; `{name}` placeholders from `files` are filled in. */
  data?: string;
  /** JSON Schema of the data; placeholders as in `data`. */
  schema?: string;
};

/** One deck from the config, with its data and schema file. */
export type ConfigJob = {
  file: string;
  dataFile?: string;
  schemaFile?: string;
};

export type ResolvedConfig = {
  jobs: ConfigJob[];
  /** Patterns that match no file (a typo, or decks not created yet). */
  unmatched: string[];
  /** Folders that may get new matching decks, for `--watch`. */
  targets: WatchTarget[];
};

const fail = (file: string, message: string): never => {
  throw new Error(`${file}: ${message}`);
};

const ENTRY_KEYS = ['files', 'data', 'schema'];
const TOP_KEYS = ['templates', 'openingTag', 'closingTag', 'fixSmartQuotes'];
const PLACEHOLDER = /\{([^{}]*)\}/g;
const NAME = /^[A-Za-z_][A-Za-z0-9_]*$/;

/** Reads and checks a config file; errors name the file and the entry. */
export const readConfig = (file: string): TemplatesConfig => {
  let config: unknown;
  try {
    config = JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (error) {
    fail(file, (error as Error).message);
  }
  if (!config || typeof config !== 'object' || Array.isArray(config)) {
    fail(file, 'expected a JSON object with "templates"');
  }
  const top = config as Record<string, unknown>;
  for (const key of Object.keys(top)) {
    if (!TOP_KEYS.includes(key)) {
      fail(file, `unknown key "${key}" (allowed: ${TOP_KEYS.join(', ')})`);
    }
  }
  if (!Array.isArray(top.templates) || !top.templates.length) {
    fail(file, '"templates" must be a non-empty list');
  }
  for (const key of ['openingTag', 'closingTag'] as const) {
    if (top[key] !== undefined && typeof top[key] !== 'string') {
      fail(file, `"${key}" must be a string`);
    }
  }
  if (
    top.fixSmartQuotes !== undefined &&
    typeof top.fixSmartQuotes !== 'boolean'
  ) {
    fail(file, '"fixSmartQuotes" must be true or false');
  }

  (top.templates as unknown[]).forEach((entry, i) => {
    const at = `templates[${i}]`;
    if (!entry || typeof entry !== 'object' || Array.isArray(entry)) {
      fail(file, `${at} must be an object`);
    }
    const e = entry as Record<string, unknown>;
    for (const key of Object.keys(e)) {
      if (!ENTRY_KEYS.includes(key)) {
        fail(
          file,
          `${at}: unknown key "${key}" (allowed: ${ENTRY_KEYS.join(', ')})`,
        );
      }
    }
    const patterns = Array.isArray(e.files) ? e.files : [e.files];
    if (
      !patterns.length ||
      patterns.some((p) => typeof p !== 'string' || !p.trim())
    ) {
      fail(file, `${at}: "files" must be a pattern or a list of patterns`);
    }
    const captured = new Set<string>();
    for (const pattern of patterns as string[]) {
      if (path.isAbsolute(pattern)) {
        fail(file, `${at}: "${pattern}" must be relative to the config file`);
      }
      const lastSegment = pattern.split(/[\\/]/).pop();
      if (!lastSegment || lastSegment === '**') {
        fail(file, `${at}: "${pattern}" must end with a file name pattern`);
      }
      for (const [, name] of pattern.matchAll(PLACEHOLDER)) {
        if (!NAME.test(name)) {
          fail(file, `${at}: "{${name}}" is not a valid placeholder name`);
        }
        captured.add(name);
      }
    }
    for (const key of ['data', 'schema'] as const) {
      if (e[key] === undefined) continue;
      if (typeof e[key] !== 'string')
        fail(file, `${at}: "${key}" must be a path`);
      for (const [, name] of (e[key] as string).matchAll(PLACEHOLDER)) {
        // Every pattern must provide it, else the path cannot be built
        const inAll = (patterns as string[]).every((p) =>
          p.includes(`{${name}}`),
        );
        if (!captured.has(name) || !inAll) {
          fail(
            file,
            `${at}: "${key}" uses {${name}}, but "files" does not define it` +
              (captured.has(name) ? ' in every pattern' : ''),
          );
        }
      }
    }
  });
  return top as TemplatesConfig;
};

/** Every character with a meaning in a regex, as literal text. */
const escapeRegExp = (text: string) =>
  // Not `-` or `/`: literal outside a class, and `\-` is invalid with /u
  text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** One path segment of a pattern → a regex over a name, with captures. */
const segmentRegExp = (segment: string): RegExp => {
  const seen = new Set<string>();
  let source = '';
  for (const part of segment.split(/(\*|\{[^{}]*\})/)) {
    if (part === '*') source += '.*?';
    else if (/^\{.*\}$/.test(part)) {
      const name = part.slice(1, -1);
      // The same name twice in one segment must match the same text
      source += seen.has(name) ? `\\k<${name}>` : `(?<${name}>.+?)`;
      seen.add(name);
    } else source += escapeRegExp(part);
  }
  const caseless =
    process.platform === 'win32' || process.platform === 'darwin';
  return new RegExp(`^${source}$`, caseless ? 'iu' : 'u');
};

const hasWildcard = (segment: string) => /\*|\{[^{}]*\}/.test(segment);

/** Folders never searched: dependencies, version control, hidden folders. */
const skipDir = (name: string) =>
  name === 'node_modules' || name.startsWith('.');

const listDir = (dir: string): fs.Dirent[] => {
  try {
    return fs.readdirSync(dir, { withFileTypes: true });
  } catch {
    return []; // not there (yet): nothing matches
  }
};

type Match = { file: string; captures: Record<string, string> };

/**
 * Files under `base` that match `pattern`, plus the folders a new match
 * could appear in. Captures of the same name must agree across segments.
 */
const matchPattern = (
  base: string,
  pattern: string,
): { matches: Match[]; targets: WatchTarget[] } => {
  const segments = pattern.split(/[\\/]+/).filter((s) => s && s !== '.');
  const matches: Match[] = [];
  const targets: WatchTarget[] = [];

  // Leading segments without wildcards are plain folders (`..` too)
  let start = base;
  let i = 0;
  while (i < segments.length - 1 && !hasWildcard(segments[i])) {
    start = path.join(start, segments[i++]);
  }

  const walk = (
    dir: string,
    index: number,
    captures: Record<string, string>,
  ): void => {
    const segment = segments[index];
    const last = index === segments.length - 1;
    if (segment === '**') {
      targets.push({ dir, accepts: () => true });
      walk(dir, index + 1, captures); // zero folders
      for (const entry of listDir(dir)) {
        if (entry.isDirectory() && !skipDir(entry.name)) {
          walk(path.join(dir, entry.name), index, captures);
        }
      }
      return;
    }
    const regexp = segmentRegExp(segment);
    targets.push({ dir, accepts: (name) => regexp.test(name) });
    for (const entry of listDir(dir)) {
      const m = regexp.exec(entry.name);
      if (!m) continue;
      const found = { ...captures };
      let agrees = true;
      for (const [name, value] of Object.entries(m.groups ?? {})) {
        if (name in found && nameKey(found[name]) !== nameKey(value)) {
          agrees = false;
        }
        found[name] = value;
      }
      if (!agrees) continue;
      const full = path.join(dir, entry.name);
      if (last) {
        if (entry.isFile() && !entry.name.startsWith('~$')) {
          matches.push({ file: full, captures: found });
        }
      } else if (entry.isDirectory() && !skipDir(entry.name)) {
        walk(full, index + 1, found);
      }
    }
  };
  walk(start, i, {});
  return { matches, targets };
};

const fill = (template: string, captures: Record<string, string>) =>
  template.replace(PLACEHOLDER, (_, name: string) => captures[name]);

/**
 * The decks a config names, each with its data and schema file, sorted by
 * path. Paths are relative to the config file's folder; a deck two entries
 * match is an error.
 */
export const resolveConfig = (
  config: TemplatesConfig,
  configFile: string,
): ResolvedConfig => {
  const base = path.dirname(configFile);
  const jobs = new Map<string, ConfigJob & { entry: number }>();
  const unmatched: string[] = [];
  const targets: WatchTarget[] = [];

  config.templates.forEach((entry, index) => {
    const patterns = Array.isArray(entry.files) ? entry.files : [entry.files];
    for (const pattern of patterns) {
      const found = matchPattern(base, pattern);
      targets.push(...found.targets);
      if (!found.matches.length) unmatched.push(pattern);
      for (const { file, captures } of found.matches) {
        const key = nameKey(path.resolve(file));
        const earlier = jobs.get(key);
        if (earlier && earlier.entry !== index) {
          fail(
            configFile,
            `${file} matches templates[${earlier.entry}] and templates[${index}]`,
          );
        }
        jobs.set(key, {
          entry: index,
          file,
          ...(entry.data && {
            dataFile: path.join(base, fill(entry.data, captures)),
          }),
          ...(entry.schema && {
            schemaFile: path.join(base, fill(entry.schema, captures)),
          }),
        });
      }
    }
  });

  const sorted = [...jobs.values()]
    .sort((a, b) => a.file.localeCompare(b.file, 'en', { numeric: true }))
    .map(({ file, dataFile, schemaFile }) => ({
      file,
      ...(dataFile && { dataFile }),
      ...(schemaFile && { schemaFile }),
    }));
  return { jobs: sorted, unmatched, targets };
};
