import {
  CliIO,
  CommandRun,
  issueLocation,
  ParsedArgs,
  printable,
  runFileCommand,
  style,
  unreadableRun,
} from './cli';
import { collectDataUsage } from './lint';
import { LintTemplateOptions } from './types';
import { DataUsageEntry, DataUsageLocation, TemplateDataUsage } from './usage';

export const USAGE_USAGE = `Usage: pptx-templates usage <file.pptx>... [options]

Shows which fields of the data the {{…}} tags of PowerPoint templates read,
which fields no tag reads, and which read paths are not in the data.

Options:
  --data <file.json>        Example data to compare with
  --schema <file.json>      JSON Schema of the data (alternative to --data)
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

const usageFiles = async (
  parsed: ParsedArgs,
  options: LintTemplateOptions,
  color: boolean,
): Promise<CommandRun> => {
  const hasData = !!(options.data || options.schema);
  const paint = (fn: (t: string) => string, text: string) =>
    color ? fn(text) : text;
  const files: Record<string, TemplateDataUsage> = {};
  let stdout = '';
  for (const file of parsed.files) {
    try {
      files[file] = await collectDataUsage([file], options);
    } catch (error) {
      return unreadableRun(file, error, stdout);
    }
    if (!parsed.json) {
      stdout +=
        (stdout ? '\n' : '') + formatUsage(file, files[file], hasData, color);
    }
  }
  // The files were readable a moment ago; read them once more together
  let unusedInAll: string[] | undefined;
  if (hasData) {
    try {
      unusedInAll = (await collectDataUsage(parsed.files, options)).unused;
    } catch (error) {
      return unreadableRun(parsed.files.join(', '), error, stdout);
    }
  }

  if (parsed.json) {
    stdout =
      JSON.stringify({ files, ...(hasData && { unusedInAll }) }, null, 2) +
      '\n';
  } else if (!hasData) {
    stdout += `\n${paint(style.dim, 'Pass --data or --schema to see unused fields.')}\n`;
  } else if (parsed.files.length > 1) {
    stdout +=
      '\n' +
      (unusedInAll.length
        ? section(
            `Unused in all ${parsed.files.length} file(s)`,
            unusedInAll,
            '',
            paint,
            style.yellow,
          )
        : `Every field is used in at least one of the ${parsed.files.length} file(s)\n`);
  }
  return { code: 0, stdout, stderr: '', unreadable: false };
};

/** `pptx-templates usage …`; resolves to the exit code. */
export const runUsageCommand = (args: string[], io: CliIO): Promise<number> =>
  runFileCommand(args, io, { usage: USAGE_USAGE, run: usageFiles });
