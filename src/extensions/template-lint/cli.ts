import fs from 'fs';
import { lintTemplate } from './lint';
import { LintTemplateOptions, TemplateLintIssue } from './types';
import { FileWatcher, watchFiles } from './watch';

export type CliIO = {
  stdout: (text: string) => void;
  stderr: (text: string) => void;
  /** Colored output; default off (the bin sets it from the terminal). */
  color?: boolean;
  /** `--watch` clears the screen before each run (the bin sets it for a terminal). */
  clearScreen?: boolean;
  /** Ends `--watch`; without it the watch runs until the process ends. */
  signal?: AbortSignal;
  /** `--watch` timings; the defaults suit PowerPoint, tests shorten them. */
  watchTiming?: Partial<WatchTiming>;
};

type WatchTiming = {
  /** Quiet time after the last file event before linting. */
  debounceMs: number;
  /** Extra attempts when a file cannot be read (mid-save, locked). */
  retries: number;
  retryDelayMs: number;
};

const DEFAULT_WATCH_TIMING: WatchTiming = {
  debounceMs: 300,
  retries: 5,
  retryDelayMs: 200,
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
  -w, --watch               Lint again whenever a file (or the --data/--schema
                            file) is saved; the deck may stay open in
                            PowerPoint. Stop with Ctrl+C
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
    // JSON.parse quotes the file content in its message
    throw new Error(
      printable(`Cannot read ${what} ${file}: ${(error as Error).message}`),
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
  /** Read on every run, so `--watch` picks up changes. */
  dataFile?: string;
  schemaFile?: string;
  json: boolean;
  color?: boolean;
  maxWarnings?: number;
  watch: boolean;
  help: boolean;
};

const parseArgs = (args: string[]): ParsedArgs => {
  const parsed: ParsedArgs = {
    files: [],
    options: {},
    json: false,
    watch: false,
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
      case '-w':
      case '--watch':
        parsed.watch = true;
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
        parsed.dataFile = value();
        break;
      case '--schema':
        parsed.schemaFile = value();
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

const loadOptions = (parsed: ParsedArgs): LintTemplateOptions => {
  const options = { ...parsed.options };
  if (parsed.dataFile) options.data = readJson(parsed.dataFile, 'data');
  if (parsed.schemaFile) {
    options.schema = readJson(parsed.schemaFile, 'schema');
  }
  return options;
};

type LintRun = {
  code: number;
  stdout: string;
  stderr: string;
  /** A file could not be read; `--watch` tries again before reporting. */
  unreadable: boolean;
};

/** Lints all files; output is collected, not printed. */
const lintFiles = async (
  parsed: ParsedArgs,
  options: LintTemplateOptions,
  color: boolean,
): Promise<LintRun> => {
  let errors = 0;
  let warnings = 0;
  let stdout = '';
  const report: Record<string, TemplateLintIssue[]> = {};
  for (const file of parsed.files) {
    let issues: TemplateLintIssue[];
    try {
      ({ issues } = await lintTemplate(file, options));
    } catch (error) {
      return {
        code: 2,
        stdout,
        stderr: `${printable(`Cannot read ${file}: ${(error as Error).message}`)}\n`,
        unreadable: true,
      };
    }
    report[file] = issues;
    errors += issues.filter((i) => i.severity === 'error').length;
    warnings += issues.filter((i) => i.severity === 'warning').length;
    if (!parsed.json) stdout += formatIssues(file, issues, color);
  }

  if (parsed.json) {
    stdout += JSON.stringify(report, null, 2) + '\n';
  } else {
    const total = errors + warnings;
    const paint = (fn: (t: string) => string, text: string) =>
      color ? fn(text) : text;
    stdout += total
      ? `\n${paint(
          (t) => style.bold(errors ? style.red(t) : style.yellow(t)),
          `${total} problem(s): ${errors} error(s), ${warnings} warning(s)`,
        )}\n`
      : `${paint((t) => style.bold(ansi(32, 39)(t)), `No problems found in ${parsed.files.length} file(s)`)}\n`;
  }

  const tooManyWarnings =
    parsed.maxWarnings !== undefined && warnings > parsed.maxWarnings;
  return {
    code: errors || tooManyWarnings ? 1 : 0,
    stdout,
    stderr: '',
    unreadable: false,
  };
};

const delay = (ms: number) =>
  new Promise<void>((resolve) => setTimeout(resolve, ms));

const clock = () => new Date().toTimeString().slice(0, 8);

/**
 * `--watch`: lints, then again after every save of a deck or of the
 * `--data`/`--schema` file. A file read mid-save (half-written ZIP, briefly
 * missing during PowerPoint's rename, locked) is retried before the error is
 * shown, and a lint error never ends the watch. Resolves with 0 once
 * `io.signal` aborts, and with 2 when watching itself fails (the folder was
 * deleted); without a signal, it runs until the process ends.
 */
const watchAndLint = (
  parsed: ParsedArgs,
  io: CliIO,
  color: boolean,
): Promise<number> => {
  const timing = { ...DEFAULT_WATCH_TIMING, ...io.watchTiming };
  const { signal } = io;
  const watched = [parsed.dataFile, parsed.schemaFile, ...parsed.files].filter(
    Boolean,
  );
  const dim = (text: string) => (color ? style.dim(text) : text);
  let stopped = false;

  const attempt = async (): Promise<LintRun> => {
    try {
      return await lintFiles(parsed, loadOptions(parsed), color);
    } catch (error) {
      // --data/--schema unreadable or invalid JSON
      const stderr = `${(error as Error).message}\n`;
      return { code: 2, stdout: '', stderr, unreadable: true };
    }
  };

  const lintAndReport = async () => {
    let result = await attempt();
    for (let i = 0; result.unreadable && i < timing.retries; i++) {
      if (stopped) return;
      await delay(timing.retryDelayMs);
      result = await attempt();
    }
    if (stopped) return;
    if (io.clearScreen) io.stdout('\u001b[2J\u001b[3J\u001b[H');
    if (result.stdout) io.stdout(result.stdout);
    if (result.stderr) io.stderr(result.stderr);
    io.stdout(
      dim(
        `\n[${clock()}] Watching ${watched.length} file(s) for changes. Press Ctrl+C to stop.\n`,
      ),
    );
  };

  return new Promise<number>((resolve) => {
    let timer: ReturnType<typeof setTimeout>;
    let running: Promise<void> | undefined;
    let again = false;

    const run = () => {
      if (running) {
        // Saved while linting: lint once more afterwards
        again = true;
        return;
      }
      running = lintAndReport()
        .catch((error) => io.stderr(`${error?.stack ?? error}\n`))
        .finally(() => {
          running = undefined;
          if (again && !stopped) {
            again = false;
            run();
          }
        });
    };
    const schedule = () => {
      clearTimeout(timer);
      timer = setTimeout(run, timing.debounceMs);
    };

    let watcher: FileWatcher;
    const stop = (code: number) => {
      if (stopped) return;
      stopped = true;
      clearTimeout(timer);
      watcher?.close();
      signal?.removeEventListener('abort', onAbort);
      Promise.resolve(running).then(() => resolve(code));
    };
    const onAbort = () => stop(0);

    try {
      watcher = watchFiles(watched, schedule, (error) => {
        // Without a working watcher the command would wait forever
        io.stderr(`${printable(`Watching failed: ${error.message}`)}\n`);
        stop(2);
      });
    } catch (error) {
      io.stderr(`${printable(`Cannot watch: ${(error as Error).message}`)}\n`);
      resolve(2);
      return;
    }

    if (signal?.aborted) return onAbort();
    signal?.addEventListener('abort', onAbort, { once: true });
    run();
  });
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
  if (parsed.watch && parsed.json) {
    io.stderr(`--watch cannot be combined with --json.\n\n${LINT_USAGE}\n`);
    return 2;
  }

  const color = parsed.color ?? io.color ?? false;
  if (parsed.watch) return watchAndLint(parsed, io, color);

  let options: LintTemplateOptions;
  try {
    options = loadOptions(parsed);
  } catch (error) {
    io.stderr(`${(error as Error).message}\n\n${LINT_USAGE}\n`);
    return 2;
  }
  const result = await lintFiles(parsed, options, color);
  if (result.stdout) io.stdout(result.stdout);
  if (result.stderr) io.stderr(result.stderr);
  return result.code;
};
