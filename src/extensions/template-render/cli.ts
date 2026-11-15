import fs from 'fs';
import path from 'path';
import { CliIO, printable, style } from '../template-lint/cli';
import { renderDeck, RenderDeckOptions } from './index';

export const RENDER_USAGE = `Usage: pptx-templates render <file.pptx> --data <data.json> -o <out.pptx> [options]

Fills the {{…}} tags of every slide of a template with the data and writes
the result as a new .pptx.

Options:
  --data <file.json>        The data the tags read (required)
  -o, --out <file.pptx>     Where to write the result (required)
  --helpers <file.js>       A module whose exports the tags can call, such
                            as {{euro(price)}}; passed as additionalJsContext.
                            It runs as code: only use files you trust
  --image-root <dir>        Folder that picture paths in the data are read
                            from, as in {"logo": "images/logo.png"} with
                            {{IMAGE logo}} (default: the folder of --data)
  --slides <list>           Slides to keep, in this order, e.g. 1,3,2
                            (default: all)
  --opening-tag <text>      Tag start (default: {{)
  --closing-tag <text>      Tag end (default: }})
  --fix-smart-quotes        Accept typographic quotes from AutoCorrect
  --color / --no-color      Force colors on or off
  -h, --help                Show this help

Exit code: 0 written, 1 a tag failed to render, 2 invalid call or
unreadable file.`;

type RenderArgs = {
  file?: string;
  dataFile?: string;
  out?: string;
  helpers?: string;
  options: RenderDeckOptions;
  color?: boolean;
  help: boolean;
};

const parseRenderArgs = (args: string[]): RenderArgs => {
  const parsed: RenderArgs = { options: {}, help: false };
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
      case '--data':
        parsed.dataFile = value();
        break;
      case '-o':
      case '--out':
        parsed.out = value();
        break;
      case '--helpers':
        parsed.helpers = value();
        break;
      case '--image-root':
        parsed.options.imageRoot = value();
        break;
      case '--slides': {
        const list = value()
          .split(',')
          .map((s) => Number(s.trim()));
        if (!list.length || list.some((n) => !Number.isInteger(n) || n < 1)) {
          throw new Error('--slides needs slide numbers like 1,3,2');
        }
        parsed.options.slides = list;
        break;
      }
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
        if (parsed.file) throw new Error('render takes one .pptx file');
        parsed.file = arg;
    }
  }
  return parsed;
};

const readData = (file: string): Record<string, unknown> => {
  let value: unknown;
  try {
    value = JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch (error) {
    throw new Error(`Cannot read data ${file}: ${(error as Error).message}`, {
      cause: error,
    });
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`Cannot read data ${file}: expected a JSON object`);
  }
  return value as Record<string, unknown>;
};

/** The exports of a helpers module: `module.exports`, or its default export. */
export const loadHelpers = (file: string): Record<string, unknown> => {
  let mod: unknown;
  try {
    // eslint-disable-next-line @typescript-eslint/no-require-imports -- the CLI is CommonJS; require() also loads ESM on current Node
    mod = require(path.resolve(file));
  } catch (error) {
    throw new Error(
      `Cannot load helpers ${file}: ${(error as Error).message}`,
      { cause: error },
    );
  }
  const exports = mod as Record<string, unknown>;
  const value =
    exports && typeof exports.default === 'object' && exports.default
      ? exports.default
      : exports;
  if (!value || typeof value !== 'object') {
    throw new Error(`Cannot load helpers ${file}: it exports no object`);
  }
  return value as Record<string, unknown>;
};

/** `pptx-templates render …`; resolves to the exit code. */
export const runRenderCommand = async (
  args: string[],
  io: CliIO,
): Promise<number> => {
  let parsed: RenderArgs;
  try {
    parsed = parseRenderArgs(args);
  } catch (error) {
    io.stderr(`${(error as Error).message}\n\n${RENDER_USAGE}\n`);
    return 2;
  }
  if (parsed.help) {
    io.stdout(`${RENDER_USAGE}\n`);
    return 0;
  }
  const missing = [
    !parsed.file && 'a .pptx file',
    !parsed.dataFile && '--data',
    !parsed.out && '--out',
  ].filter(Boolean);
  if (missing.length) {
    io.stderr(`Missing ${missing.join(', ')}.\n\n${RENDER_USAGE}\n`);
    return 2;
  }
  const color = parsed.color ?? io.color ?? false;
  const paint = (fn: (t: string) => string, text: string) =>
    color ? fn(text) : text;

  let data: Record<string, unknown>;
  let content: Buffer;
  try {
    data = readData(parsed.dataFile);
    parsed.options.imageRoot ??= path.dirname(path.resolve(parsed.dataFile));
    if (parsed.helpers) {
      parsed.options.additionalJsContext = loadHelpers(parsed.helpers);
    }
    content = fs.readFileSync(parsed.file);
  } catch (error) {
    io.stderr(`${printable((error as Error).message)}\n`);
    return 2;
  }

  let result: Buffer;
  try {
    result = await renderDeck(content, data, parsed.options);
  } catch (error) {
    const err = error as Error;
    const isTagError = err?.name === 'TemplateRenderError';
    io.stderr(
      `${paint(style.underline, printable(parsed.file))}\n  ${paint((t) => style.bold(style.red(t)), 'error')}  ${printable(err?.message ?? String(error))}\n`,
    );
    return isTagError ? 1 : 2;
  }

  try {
    fs.mkdirSync(path.dirname(path.resolve(parsed.out)), { recursive: true });
    fs.writeFileSync(parsed.out, result);
  } catch (error) {
    io.stderr(
      `${printable(`Cannot write ${parsed.out}: ${(error as Error).message}`)}\n`,
    );
    return 2;
  }
  io.stdout(
    `${paint(style.bold, 'rendered')}  ${printable(parsed.out)}  ${paint(style.dim, `from ${printable(parsed.file)}`)}\n`,
  );
  return 0;
};
