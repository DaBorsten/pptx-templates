import fs from 'fs';
import { lintTemplate } from './lint';
import { LintTemplateOptions, TemplateLintIssue } from './types';

export type CliIO = {
  stdout: (text: string) => void;
  stderr: (text: string) => void;
  /** Colored output; default off (the bin sets it from the terminal). */
  color?: boolean;
};

export const LINT_USAGE = `Usage: pptx-templates lint <file.pptx>... [options]

Checks the {{…}} tags of PowerPoint templates before rendering.

Options:
  --data <file.json>        Example data: report names and properties that
                            are not in it
  --schema <file.json>      JSON Schema of the data (alternative to --data)
  --opening-tag <text>      Tag start (default: {{)
  --closing-tag <text>      Tag end (default: }})
  --fix-smart-quotes        Accept typographic quotes, like renderTemplate's
                            fixSmartQuotes option
  --max-warnings <n>        Fail when there are more than n warnings
  --json                    Print the issues as JSON
  --color / --no-color      Force colors on or off (default: on in a terminal;
                            NO_COLOR and FORCE_COLOR are respected)
  -h, --help                Show this help

Exit code: 0 no errors, 1 errors (or too many warnings), 2 invalid call.`;

const readJson = (file: string, what: string): Record<string, unknown> => {
  try {
    const value = JSON.parse(fs.readFileSync(file, 'utf8'));
    if (!value || typeof value !== 'object' || Array.isArray(value)) {
      throw new Error('expected a JSON object');
    }
    return value;
  } catch (error) {
    throw new Error(
      `Cannot read ${what} ${file}: ${(error as Error).message}`,
      {
        cause: error,
      },
    );
  }
};

/** Where in the deck an issue is, for humans: "slide 2 › Title". */
export const issueLocation = (issue: TemplateLintIssue): string => {
  const parts: string[] = [];
  let m: RegExpExecArray;
  if ((m = /slides\/slide(\d+)\.xml$/.exec(issue.part)))
    parts.push(`slide ${m[1]}`);
  else if ((m = /slideLayout(\d+)\.xml$/.exec(issue.part)))
    parts.push(`layout ${m[1]}`);
  else if ((m = /slideMaster(\d+)\.xml$/.exec(issue.part)))
    parts.push(`master ${m[1]}`);
  else if (/notesSlide\d+\.xml$/.test(issue.part)) {
    parts.push(issue.slide ? `notes of slide ${issue.slide}` : 'notes');
  } else if ((m = /chart(\d+)\.xml$/.exec(issue.part))) {
    parts.push(
      `chart ${m[1]}` + (issue.slide ? ` on slide ${issue.slide}` : ''),
    );
  } else if ((m = /diagrams\/data(\d+)\.xml$/.exec(issue.part))) {
    parts.push(
      `SmartArt ${m[1]}` + (issue.slide ? ` on slide ${issue.slide}` : ''),
    );
  } else parts.push(issue.part);
  if (issue.shape) parts.push(`"${issue.shape}"`);
  return parts.join(' › ');
};

const ansi = (open: number, close: number) => (text: string) =>
  `\u001b[${open}m${text}\u001b[${close}m`;
const style = {
  bold: ansi(1, 22),
  dim: ansi(2, 22),
  underline: ansi(4, 24),
  red: ansi(31, 39),
  yellow: ansi(33, 39),
  cyan: ansi(36, 39),
};

/** Color on for a real terminal, off for pipes, `NO_COLOR` and `--no-color`. */
export const supportsColor = (
  env: NodeJS.ProcessEnv = process.env,
  isTTY: boolean = !!process.stdout.isTTY,
): boolean => {
  if (env.NO_COLOR !== undefined && env.NO_COLOR !== '') return false;
  if (env.FORCE_COLOR !== undefined && env.FORCE_COLOR !== '0') return true;
  return isTTY && env.TERM !== 'dumb';
};

/**
 * Text from the .pptx (tags, shape names) goes to a terminal: control
 * characters could hide or forge output there. Each becomes one space, so
 * the marker line under a tag still fits (a tab would shift it, too).
 */
const printable = (text: string): string =>
  // eslint-disable-next-line no-control-regex -- stripping control chars is the point here
  text.replace(/[\u0000-\u001f\u007f-\u009f]/g, ' ');

export const formatIssues = (
  file: string,
  issues: TemplateLintIssue[],
  color = false,
): string => {
  if (!issues.length) return '';
  const paint = (fn: (t: string) => string, text: string) =>
    color ? fn(text) : text;
  const lines = [paint(style.underline, printable(file))];
  issues.forEach((issue) => {
    const tone = issue.severity === 'error' ? style.red : style.yellow;
    lines.push(`  ${paint(style.cyan, printable(issueLocation(issue)))}`);
    lines.push(
      `    ${paint((t) => style.bold(tone(t)), issue.severity)}  ${printable(issue.message)}  ${paint(style.dim, `(${issue.rule})`)}`,
    );
    if (issue.tag) {
      lines.push(`    ${paint(style.dim, 'tag:')} ${printable(issue.tag)}`);
      if (issue.range) {
        const { start, end } = issue.range;
        lines.push(
          `         ${' '.repeat(start)}${paint(tone, '~'.repeat(Math.max(1, end - start)))}`,
        );
      }
    }
    if (issue.hint) {
      lines.push(`    ${paint(style.dim, 'hint:')} ${printable(issue.hint)}`);
    }
  });
  return lines.join('\n') + '\n';
};

type ParsedArgs = {
  files: string[];
  options: LintTemplateOptions;
  json: boolean;
  color?: boolean;
  maxWarnings?: number;
  help: boolean;
};

const parseArgs = (args: string[]): ParsedArgs => {
  const parsed: ParsedArgs = {
    files: [],
    options: {},
    json: false,
    help: false,
  };
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
      case '--json':
        parsed.json = true;
        break;
      case '--color':
        parsed.color = true;
        break;
      case '--no-color':
        parsed.color = false;
        break;
      case '--fix-smart-quotes':
        parsed.options.fixSmartQuotes = true;
        break;
      case '--data':
        parsed.options.data = readJson(value(), 'data');
        break;
      case '--schema':
        parsed.options.schema = readJson(value(), 'schema');
        break;
      case '--opening-tag':
        parsed.options.openingTag = value();
        break;
      case '--closing-tag':
        parsed.options.closingTag = value();
        break;
      case '--max-warnings': {
        const n = Number(value());
        if (!Number.isInteger(n) || n < 0) {
          throw new Error('--max-warnings needs a whole number');
        }
        parsed.maxWarnings = n;
        break;
      }
      default:
        if (arg.startsWith('-')) throw new Error(`Unknown option ${arg}`);
        parsed.files.push(arg);
    }
  }
  return parsed;
};

/** `pptx-templates lint …`; resolves to the exit code. */
export const runLintCommand = async (
  args: string[],
  io: CliIO,
): Promise<number> => {
  let parsed: ParsedArgs;
  try {
    parsed = parseArgs(args);
  } catch (error) {
    io.stderr(`${(error as Error).message}\n\n${LINT_USAGE}\n`);
    return 2;
  }
  if (parsed.help) {
    io.stdout(`${LINT_USAGE}\n`);
    return 0;
  }
  if (!parsed.files.length) {
    io.stderr(`No .pptx file given.\n\n${LINT_USAGE}\n`);
    return 2;
  }

  const color = parsed.color ?? io.color ?? false;
  let errors = 0;
  let warnings = 0;
  const report: Record<string, TemplateLintIssue[]> = {};
  for (const file of parsed.files) {
    let issues: TemplateLintIssue[];
    try {
      ({ issues } = await lintTemplate(file, parsed.options));
    } catch (error) {
      io.stderr(`Cannot read ${file}: ${(error as Error).message}\n`);
      return 2;
    }
    report[file] = issues;
    errors += issues.filter((i) => i.severity === 'error').length;
    warnings += issues.filter((i) => i.severity === 'warning').length;
    if (!parsed.json) io.stdout(formatIssues(file, issues, color));
  }

  if (parsed.json) {
    io.stdout(JSON.stringify(report, null, 2) + '\n');
  } else {
    const total = errors + warnings;
    const paint = (fn: (t: string) => string, text: string) =>
      color ? fn(text) : text;
    io.stdout(
      total
        ? `\n${paint(
            (t) => style.bold(errors ? style.red(t) : style.yellow(t)),
            `${total} problem(s): ${errors} error(s), ${warnings} warning(s)`,
          )}\n`
        : `${paint((t) => style.bold(ansi(32, 39)(t)), `No problems found in ${parsed.files.length} file(s)`)}\n`,
    );
  }

  const tooManyWarnings =
    parsed.maxWarnings !== undefined && warnings > parsed.maxWarnings;
  return errors || tooManyWarnings ? 1 : 0;
};
