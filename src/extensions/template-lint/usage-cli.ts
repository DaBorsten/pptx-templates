import {
  CliIO,
  CommandRun,
  issueLocation,
  Job,
  ParsedArgs,
  printable,
  runFileCommand,
  style,
  unreadableRun,
} from './cli';
import { CONFIG_FILE } from './config';
import { collectDataUsage } from './lint';
import { DataUsageEntry, DataUsageLocation, TemplateDataUsage } from './usage';

/** Decks that share their data: "unused" across them makes sense. */
type DataGroup = {
  data?: string;
  schema?: string;
  files: string[];
  unused: string[];
};

export const USAGE_USAGE = `Usage: pptx-templates usage <file.pptx>... [options]
       pptx-templates usage [--config <file.json>] [options]

Shows which fields of the data the {{…}} tags of PowerPoint templates read,
which fields no tag reads, and which read paths are not in the data.

Options:
  --data <file.json>        Example data to compare with
  --schema <file.json>      JSON Schema of the data (alternative to --data)
  -c, --config <file.json>  Decks and their data from a config file; without
                            files, ${CONFIG_FILE} in this folder is used
  --opening-tag <text>      Tag start (default: {{)
  --closing-tag <text>      Tag end (default: }})
  --fix-smart-quotes        Accept typographic quotes, like renderTemplate's
                            fixSmartQuotes option
  --json                    Print the result as JSON
  -w, --watch               Run again whenever a file (or the --data/--schema
                            file) is saved. Stop with Ctrl+C
  --color / --no-color      Force colors on or off (default: on in a terminal;
                            NO_COLOR and FORCE_COLOR are respected)
  -h, --help                Show this help

Without --data or --schema, only the read paths are listed.
Exit code: 0 done, 2 invalid call or unreadable file.`;

/** "slide 1 › "Titel"", as the linter shows locations. */
const where = (location: DataUsageLocation): string =>
  issueLocation({
    part: location.part,
    slide: location.slide,
    shape: location.shape,
  } as Parameters<typeof issueLocation>[0]);

const section = (
  title: string,
  entries: (DataUsageEntry | string)[],
  indent: string,
  paint: (fn: (t: string) => string, text: string) => string,
  tone: (t: string) => string = (t) => t,
): string => {
  const rows = entries.map((entry) =>
    typeof entry === 'string'
      ? { path: entry, at: '' }
      : { path: entry.path, at: entry.locations.map(where).join(', ') },
  );
  const width = Math.max(...rows.map((r) => r.path.length));
  const lines = rows.map(({ path, at }) => {
    const shown = paint(tone, printable(path));
    return at
      ? `${indent}  ${shown}${' '.repeat(width - path.length)}  ${paint(style.dim, printable(at))}`
      : `${indent}  ${shown}`;
  });
  return `${indent}${paint(style.bold, `${title} (${entries.length})`)}\n${lines.join('\n')}\n`;
};

export const formatUsage = (
  file: string,
  usage: TemplateDataUsage,
  hasData: boolean,
  color = false,
): string => {
  const paint = (fn: (t: string) => string, text: string) =>
    color ? fn(text) : text;
  let text = `${paint(style.underline, printable(file))}\n`;
  text += usage.used.length
    ? section('used', usage.used, '  ', paint)
    : '  used: none\n';
  if (hasData) {
    text += usage.unused.length
      ? section('unused', usage.unused, '  ', paint, style.yellow)
      : '  unused: none\n';
  }
  if (usage.missing.length) {
    text += section('missing in data', usage.missing, '  ', paint, style.red);
  }
  return text;
};

const hasData = (job: Job) => !!(job.options.data || job.options.schema);

const usageFiles = async (
  jobs: Job[],
  parsed: ParsedArgs,
  color: boolean,
): Promise<CommandRun> => {
  const paint = (fn: (t: string) => string, text: string) =>
    color ? fn(text) : text;
  const files: Record<string, TemplateDataUsage> = {};
  let stdout = '';
  for (const job of jobs) {
    try {
      files[job.file] = await collectDataUsage([job.file], job.options);
    } catch (error) {
      return unreadableRun(job.file, error, stdout);
    }
    if (!parsed.json) {
      stdout +=
        (stdout ? '\n' : '') +
        formatUsage(job.file, files[job.file], hasData(job), color);
    }
  }

  // Decks with the same data and schema: what none of them reads
  const byData = new Map<string, Job[]>();
  for (const job of jobs.filter(hasData)) {
    const key = JSON.stringify([job.dataFile, job.schemaFile]);
    if (!byData.has(key)) byData.set(key, []);
    byData.get(key).push(job);
  }
  const groups: DataGroup[] = [];
  for (const members of byData.values()) {
    const decks = members.map((job) => job.file);
    let unused: string[];
    try {
      // The decks were readable a moment ago; read them once more together
      unused = (await collectDataUsage(decks, members[0].options)).unused;
    } catch (error) {
      return unreadableRun(decks.join(', '), error, stdout);
    }
    groups.push({
      ...(members[0].dataFile && { data: members[0].dataFile }),
      ...(members[0].schemaFile && { schema: members[0].schemaFile }),
      files: decks,
      unused,
    });
  }

  if (parsed.json) {
    stdout = JSON.stringify({ files, unusedInAll: groups }, null, 2) + '\n';
  } else if (!groups.length) {
    stdout += `\n${paint(style.dim, 'Pass --data or --schema to see unused fields.')}\n`;
  } else {
    for (const group of groups.filter((g) => g.files.length > 1)) {
      // Name the data only where decks use different data
      const source =
        byData.size > 1
          ? ` using ${[group.data, group.schema].filter(Boolean).join(' + ')}`
          : '';
      const n = group.files.length;
      stdout +=
        '\n' +
        (group.unused.length
          ? section(
              `Unused in all ${n} file(s)${source}`,
              group.unused,
              '',
              paint,
              style.yellow,
            )
          : `Every field is used in at least one of the ${n} file(s)${source}\n`);
    }
  }
  return { code: 0, stdout, stderr: '', unreadable: false };
};

/** `pptx-templates usage …`; resolves to the exit code. */
export const runUsageCommand = (args: string[], io: CliIO): Promise<number> =>
  runFileCommand(args, io, { usage: USAGE_USAGE, run: usageFiles });
