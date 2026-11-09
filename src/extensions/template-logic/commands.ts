import escape from 'regexp.escape';
import { RenderTemplateOptions, TemplateCommandType } from './types';

export type ParsedCommand = {
  type: TemplateCommandType;
  /** The command without keyword. */
  code: string;
  /** FOR: loop variable; ALIAS: alias name. */
  name?: string;
};

const IDENT = '[\\p{L}_$][\\p{L}\\p{N}_$]*';

/** An operator that needs a value before it: no command's code starts so. */
const CONTINUES_EXPRESSION = /^(?:[+\-*/%<>=&|?:.,)\]]|!=)/;

/**
 * `{{each:row}}`: a marker for `modify.setTable`'s `expand` (nested
 * tables), not a template command. It is never a valid expression, so
 * renderTemplate leaves it as text for setTable, and the linter skips it.
 */
export const isTableMarker = (inner: string): boolean =>
  /^each:[^\s{}]+$/.test(inner.trim());

/** Resolved delimiters plus a global regex matching one tag. */
export const tagSyntax = (
  options?: Pick<
    RenderTemplateOptions,
    'openingTag' | 'closingTag' | 'cmdDelimiter'
  >,
) => {
  const delimiter = options?.cmdDelimiter;
  const [open, close] = Array.isArray(delimiter)
    ? delimiter
    : [delimiter, delimiter];
  const opening = options?.openingTag ?? open ?? '{{';
  const closing = options?.closingTag ?? close ?? '}}';
  return {
    opening,
    closing,
    pattern: new RegExp(
      `${escape(opening)}([\\s\\S]*?)${escape(closing)}`,
      'g',
    ),
  };
};

/**
 * Classifies the inside of a tag (`IF a > 1`, `= name`, `item.price`, …),
 * following the docx-templates command set. `*alias` references must be
 * resolved by the caller before.
 */
export const parseCommand = (inner: string, tag: string): ParsedCommand => {
  const content = inner.replace(/\u00A0/g, ' ').trim();
  if (!content) throw new Error(`Empty template tag ${tag}`);
  let m: RegExpExecArray;

  if ((m = /^IF\s+([\s\S]+)$/i.exec(content))) {
    return { type: 'IF', code: m[1].trim() };
  }
  if ((m = /^ELSE[-\s]?IF\s+([\s\S]+)$/i.exec(content))) {
    return { type: 'ELSE-IF', code: m[1].trim() };
  }
  if (/^ELSE$/i.test(content)) return { type: 'ELSE', code: '' };
  if (/^END-?IF\b/i.test(content)) return { type: 'END-IF', code: '' };
  if (
    (m = new RegExp(`^FOR\\s+\\$?(${IDENT})\\s+IN\\s+([\\s\\S]+)$`, 'iu').exec(
      content,
    ))
  ) {
    return { type: 'FOR', name: m[1], code: m[2].trim() };
  }
  if (/^END-?FOR\b/i.test(content)) return { type: 'END-FOR', code: '' };
  if ((m = /^(?:EXEC\s+|!)([\s\S]+)$/i.exec(content))) {
    return { type: 'EXEC', code: m[1].trim() };
  }
  if (
    (m = new RegExp(`^ALIAS\\s+(${IDENT})\\s+([\\s\\S]+)$`, 'iu').exec(content))
  ) {
    return { type: 'ALIAS', name: m[1], code: m[2].trim() };
  }
  // `{{link + 1}}`, `{{image ? 'a' : 'b'}}`: data called like a command. What
  // follows continues an expression, so the tag is an insertion.
  if (
    (m = /^(HTML|LINK|IMAGE)\s+([\s\S]+)$/i.exec(content)) &&
    !CONTINUES_EXPRESSION.test(m[2])
  ) {
    return {
      type: m[1].toUpperCase() as 'HTML' | 'LINK' | 'IMAGE',
      code: m[2].trim(),
    };
  }
  if ((m = /^(?:=|INS\s)([\s\S]+)$/i.exec(content))) {
    return { type: 'INS', code: m[1].trim() };
  }
  return { type: 'INS', code: content };
};

/**
 * docx-templates' `fixSmartQuotes`: typographic quotes from AutoCorrect
 * become straight ones, in the whole command.
 */
export const fixSmartQuotes = (command: string): string =>
  command.replace(/[“”„]/g, '"').replace(/[‘’‚]/g, "'");

/** `*name` → the aliased command; anything else unchanged. */
export const resolveAlias = (
  inner: string,
  aliases: Map<string, string>,
  tag: string,
): string => {
  const m = new RegExp(`^\\s*\\*(${IDENT})\\s*$`, 'u').exec(inner);
  if (!m) return inner;
  const command = aliases.get(m[1]);
  if (command === undefined) {
    throw new Error(`${tag}: unknown alias "${m[1]}"`);
  }
  return command;
};
