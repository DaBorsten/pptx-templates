/** Characters PowerPoint's AutoCorrect puts in for what was typed. */
const AUTOCORRECT: [string, string][] = [
  ['…', '...'],
  ['–', '-'],
  ['—', '--'],
  ['→', '->'],
  ['⇒', '=>'],
  ['←', '<-'],
  ['⇐', '<='],
  ['©', '(c)'],
  ['®', '(r)'],
  ['™', '(tm)'],
  ['½', '1/2'],
  ['¼', '1/4'],
  ['¾', '3/4'],
];

const TYPOGRAPHIC_QUOTES = /[‘’‚‛‹›“”„‟«»]/;

/** A misspelled command is parsed as an insertion and fails there. */
const COMMAND_FORMS: [RegExp, string][] = [
  [/^FOR\s/i, '{{FOR item IN list}}'],
  [/^END[\s_]/i, '{{END-IF}} or {{END-FOR}}, with a hyphen'],
  [/^ELSE[\s_]/i, '{{ELSE}} or {{ELSE-IF condition}}'],
  [/^IF\s/i, '{{IF condition}}'],
  [/^ALIAS\s/i, '{{ALIAS name command}}'],
];

const AUTOCORRECT_SETTINGS = 'File › Options › Proofing › AutoCorrect Options';

/**
 * Explains the usual cause of a syntax error in a tag written in
 * PowerPoint: AutoCorrect replaced quotes or character sequences.
 */
export const autocorrectHint = (
  code: string,
  message: string,
  fixSmartQuotes = false,
): string | undefined => {
  const hints: string[] = [];

  const replaced = AUTOCORRECT.filter(([char]) => code.includes(char));
  if (replaced.length) {
    hints.push(
      "PowerPoint's AutoCorrect replaced " +
        replaced
          .map(([char, typed]) => `"${typed}" with "${char}"`)
          .join(', ') +
        `. Type it again and press Ctrl+Z right after the replacement, or ` +
        `remove the entry in ${AUTOCORRECT_SETTINGS} › AutoCorrect.`,
    );
  }
  if (TYPOGRAPHIC_QUOTES.test(code)) {
    hints.push(
      "Strings need straight quotes ('Ada'), but PowerPoint's AutoCorrect " +
        'turned them into typographic ones (’ ‘ “ ” „). Switch off ' +
        `${AUTOCORRECT_SETTINGS} › AutoFormat As You Type › "Straight quotes" ` +
        'with "smart quotes", then type the quotes again' +
        (fixSmartQuotes ? '.' : ', or render with the fixSmartQuotes option.'),
    );
  }
  if (/Unexpected character "="/.test(message)) {
    hints.push('Use == or === to compare; = only assigns inside {{EXEC …}}.');
  }
  if (/Unexpected "(AND|OR|NOT)"/i.test(message)) {
    hints.push('Use && for AND, || for OR and ! for NOT.');
  }
  const form = COMMAND_FORMS.find(([keyword]) => keyword.test(code.trim()));
  if (form) {
    hints.push(
      'This looks like a command, but it does not match its form, so it is ' +
        `read as an expression. Write ${form[1]}.`,
    );
  }
  return hints.length ? hints.join(' ') : undefined;
};
