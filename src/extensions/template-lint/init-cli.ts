import fs from 'fs';
import path from 'path';
import { CliIO, printable, style } from './cli';
import { CONFIG_FILE, readConfig, resolveConfig } from './config';
import { templateDataSkeleton } from './lint';
import { mergeDataSkeleton, TemplateDataSkeleton } from './skeleton';
import { LintTemplateOptions } from './types';

export const INIT_USAGE = `Usage: pptx-templates init <file.pptx>... [-o <data.json>] [options]
       pptx-templates init [--config <file.json>] [options]

Writes example data for PowerPoint templates: every field the {{…}} tags
read, with a placeholder value. Several decks give one combined file.

An existing data file is completed, not replaced: its values stay, only
missing fields are added. Run it again after you add tags.

Options:
  -o, --out <data.json>     Write to this file (default: print the data)
  -c, --config <file.json>  Write the data files a config file names, each
                            from the decks mapped to it; without files,
                            ${CONFIG_FILE} in this folder is used
  --opening-tag <text>      Tag start (default: {{)
  --closing-tag <text>      Tag end (default: }})
  --fix-smart-quotes        Accept typographic quotes, like renderTemplate's
                            fixSmartQuotes option
  --color / --no-color      Force colors on or off
  -h, --help                Show this help

Placeholders: the path as text ("kunde.name"), true for values that are only
tested ({{IF kunde.vip}}), 1 for numbers (preis.toFixed(2), preis * 2), one
entry for lists ({{FOR p IN produkte}}).
Exit code: 0 done, 2 invalid call or unreadable file.`;

type InitArgs = {
  files: string[];
  out?: string;
  configFile?: string;
  options: LintTemplateOptions;
  color?: boolean;
  help: boolean;
};

const parseInitArgs = (args: string[]): InitArgs => {
  const parsed: InitArgs = { files: [], options: {}, help: false };
  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    const value = () => {
      const next = args[++i];
      if (next === undefined) throw new Error(`${arg} needs a value`);
      return next;
    };
    switch (arg) {
      case '-h':
      case '--help':
        parsed.help = true;
        break;
      case '-o':
      case '--out':
        parsed.out = value();
        break;
      case '-c':
      case '--config':
        parsed.configFile = value();
        break;
      case '--opening-tag':
        parsed.options.openingTag = value();
        break;
      case '--closing-tag':
        parsed.options.closingTag = value();
        break;
      case '--fix-smart-quotes':
        parsed.options.fixSmartQuotes = true;
        break;
      case '--color':
        parsed.color = true;
        break;
      case '--no-color':
        parsed.color = false;
        break;
      default:
        if (arg.startsWith('-')) throw new Error(`Unknown option ${arg}`);
        parsed.files.push(arg);
    }
  }
  if (parsed.configFile && (parsed.files.length || parsed.out)) {
    throw new Error(
      '--config names the files and their data: do not pass files or --out with it',
    );
  }
  return parsed;
};

/** One data file and the decks whose tags fill it. */
type Target = { out?: string; decks: string[]; options: LintTemplateOptions };

const toJson = (data: unknown) => `${JSON.stringify(data, null, 2)}\n`;

/** Reads an existing data file; undefined if there is none. */
const readExisting = (file: string): Record<string, unknown> | undefined => {
  if (!fs.existsSync(file)) return undefined;
  let value: unknown;
  try {
    value = JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (error) {
    throw new Error(
      `Cannot read ${file}: ${(error as Error).message}; it was not changed`,
      { cause: error },
    );
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`${file} is not a JSON object; it was not changed`);
  }
  return value as Record<string, unknown>;
};

/** Replaces `file` in one step, so a failed write leaves the old one. */
const writeFile = (file: string, text: string): void => {
  const temp = `${file}.${process.pid}.tmp`;
  try {
    fs.writeFileSync(temp, text);
    fs.renameSync(temp, file);
  } catch (error) {
    fs.rmSync(temp, { force: true });
    throw error;
  }
};

/** Writes or completes `out`; returns the line to report. */
const writeData = (
  out: string,
  skeleton: TemplateDataSkeleton,
  decks: number,
  paint: (fn: (t: string) => string, text: string) => string,
): string => {
  const existing = readExisting(out);
  const from = `from ${decks} file(s)`;
  if (!existing) {
    fs.mkdirSync(path.dirname(path.resolve(out)), { recursive: true });
    writeFile(out, toJson(skeleton.data));
    const fields = Object.keys(skeleton.data).length;
    return `${paint(style.bold, 'created')}  ${printable(out)}  ${paint(style.dim, `${fields} top-level field(s) ${from}`)}\n`;
  }
  const { data, added } = mergeDataSkeleton(existing, skeleton.data);
  if (!added.length) {
    return `${paint(style.bold, 'up to date')}  ${printable(out)}  ${paint(style.dim, from)}\n`;
  }
  writeFile(out, toJson(data));
  return (
    `${paint(style.bold, 'completed')}  ${printable(out)}  ${paint(style.dim, `${added.length} field(s) added ${from}`)}\n` +
    added.map((p) => `  + ${paint(style.cyan, printable(p))}\n`).join('')
  );
};

const configTargets = (
  configFile: string,
  cli: LintTemplateOptions,
): { targets: Target[]; notes: string[] } => {
  const config = readConfig(configFile);
  const resolved = resolveConfig(config, configFile);
  if (!resolved.jobs.length) {
    throw new Error(`${configFile}: no file matches its patterns`);
  }
  const options: LintTemplateOptions = {
    ...(config.openingTag !== undefined && { openingTag: config.openingTag }),
    ...(config.closingTag !== undefined && { closingTag: config.closingTag }),
    ...(config.fixSmartQuotes !== undefined && {
      fixSmartQuotes: config.fixSmartQuotes,
    }),
    ...cli,
  };
  const byData = new Map<string, Target>();
  const notes: string[] = [];
  for (const job of resolved.jobs) {
    if (!job.dataFile) {
      notes.push(`${job.file} has no "data" in ${configFile}; skipped.`);
      continue;
    }
    const key = path.resolve(job.dataFile);
    if (!byData.has(key)) {
      byData.set(key, { out: job.dataFile, decks: [], options });
    }
    byData.get(key).decks.push(job.file);
  }
  return { targets: [...byData.values()], notes };
};

/** `pptx-templates init …`; resolves to the exit code. */
export const runInitCommand = async (
  args: string[],
  io: CliIO,
): Promise<number> => {
  let parsed: InitArgs;
  try {
    parsed = parseInitArgs(args);
  } catch (error) {
    io.stderr(`${(error as Error).message}\n\n${INIT_USAGE}\n`);
    return 2;
  }
  if (parsed.help) {
    io.stdout(`${INIT_USAGE}\n`);
    return 0;
  }
  if (!parsed.files.length && !parsed.configFile) {
    if (fs.existsSync(CONFIG_FILE)) {
      parsed.configFile = CONFIG_FILE;
    } else {
      io.stderr(
        `No .pptx file given, and no ${CONFIG_FILE} in this folder.\n\n${INIT_USAGE}\n`,
      );
      return 2;
    }
  }
  const color = parsed.color ?? io.color ?? false;
  const paint = (fn: (t: string) => string, text: string) =>
    color ? fn(text) : text;

  let targets: Target[];
  try {
    if (parsed.configFile) {
      const planned = configTargets(parsed.configFile, parsed.options);
      planned.notes.forEach((note) => io.stderr(`${printable(note)}\n`));
      targets = planned.targets;
    } else {
      targets = [
        { out: parsed.out, decks: parsed.files, options: parsed.options },
      ];
    }
  } catch (error) {
    io.stderr(`${printable((error as Error).message)}\n`);
    return 2;
  }

  const functions = new Set<string>();
  let errors = 0;
  for (const target of targets) {
    let skeleton: TemplateDataSkeleton;
    try {
      skeleton = await templateDataSkeleton(target.decks, target.options);
    } catch (error) {
      io.stderr(
        `${printable(`Cannot read ${target.decks.join(', ')}: ${(error as Error).message}`)}\n`,
      );
      return 2;
    }
    skeleton.functions.forEach((name) => functions.add(name));
    errors += skeleton.errorCount;
    if (!target.out) {
      io.stdout(toJson(skeleton.data));
      continue;
    }
    try {
      io.stdout(writeData(target.out, skeleton, target.decks.length, paint));
    } catch (error) {
      io.stderr(`${printable((error as Error).message)}\n`);
      return 2;
    }
  }

  if (functions.size) {
    io.stderr(
      paint(
        style.dim,
        `Called as functions, so not in the data: ${printable([...functions].join(', '))}. Pass them to renderTemplate as additionalJsContext.\n`,
      ),
    );
  }
  if (errors) {
    io.stderr(
      paint(
        style.yellow,
        `${errors} tag error(s) skipped; run pptx-templates lint to see them.\n`,
      ),
    );
  }
  return 0;
};
