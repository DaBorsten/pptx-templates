import { FontFace } from './font-file';

/**
 * Line breaking and height of paragraphs, the way PowerPoint sets text:
 * greedy breaks at spaces and after hyphens, a word longer than the line
 * broken by characters, line height 1.2 × the largest font size of the line
 * (× the line spacing percentage, or a fixed point value).
 */

/** Spacing as points or as a share of the line (1 = 100 %). */
export type Spacing = { pts?: number; pct?: number };

type RunFormat = {
  face: FontFace;
  /** Font size in pt. */
  size: number;
  /** Extra space after each character in pt. */
  letterSpacing: number;
};

export type LayoutRun = RunFormat & ({ text: string } | { lineBreak: true });

export type LayoutParagraph = {
  runs: LayoutRun[];
  /** Size and font of an empty paragraph (endParaRPr). */
  endSize: number;
  endFace: FontFace;
  /** Margins and first line indent in pt. */
  marL: number;
  marR: number;
  indent: number;
  tabSize: number;
  bullet?: { text: string; face: FontFace; size: number };
  lineSpacing: Spacing;
  before?: Spacing;
  after?: Spacing;
};

/** Vertical space of one paragraph in pt. */
export type ParagraphMetrics = {
  /** Space before; not applied to the first paragraph of a text box. */
  before: number;
  /** Lines together. */
  body: number;
  /** Space after; does not count below the last paragraph. */
  after: number;
  lines: number;
};

/** Line height factor of "single" spacing. */
const LINE_FACTOR = 1.2;

type Glyph = {
  char: string;
  width: number;
  size: number;
  space: boolean;
  tab: boolean;
};

const widthOf = (text: string, format: RunFormat): number => {
  let width = 0;
  for (const char of text) {
    const cp = char.codePointAt(0);
    const em = format.face.has(cp) ? format.face.advance(cp) : 0.5;
    width += em * format.size + format.letterSpacing;
  }
  return width;
};

const lineHeight = (size: number, spacing: Spacing): number =>
  spacing.pts ?? LINE_FACTOR * size * (spacing.pct ?? 1);

const spacingOf = (spacing: Spacing | undefined, size: number): number =>
  !spacing ? 0 : (spacing.pts ?? LINE_FACTOR * size * (spacing.pct ?? 0));

/** Breaks a paragraph into lines; returns the font size of each line. */
export const breakLines = (p: LayoutParagraph, width: number): number[] => {
  // forced lines (<a:br/>) as glyph lists
  const segments: { glyphs: Glyph[]; size: number }[] = [
    { glyphs: [], size: 0 },
  ];
  for (const run of p.runs) {
    if ('lineBreak' in run) {
      const current = segments[segments.length - 1];
      current.size = Math.max(current.size, run.size);
      segments.push({ glyphs: [], size: 0 });
      continue;
    }
    for (const char of run.text) {
      segments[segments.length - 1].glyphs.push({
        char,
        width: widthOf(char, run),
        size: run.size,
        space: char === ' ', // a no-break space keeps words together
        tab: char === '\t',
      });
    }
  }

  const bulletWidth = p.bullet
    ? widthOf(p.bullet.text, { ...p.bullet, letterSpacing: 0 })
    : 0;
  const firstStart = p.bullet
    ? Math.max(p.marL, p.marL + p.indent + bulletWidth)
    : p.marL + p.indent;
  const right = width - p.marR;

  const sizes: number[] = [];
  let first = true;

  for (const segment of segments) {
    const glyphs = segment.glyphs;
    if (!glyphs.length) {
      sizes.push(segment.size || p.endSize);
      first = false;
      continue;
    }

    // words: non-space runs (split after hyphens) plus their trailing spaces
    type Word = { glyphs: Glyph[]; trailing: Glyph[] };
    const words: Word[] = [];
    let word: Word = { glyphs: [], trailing: [] };
    for (const g of glyphs) {
      if (g.space || g.tab) {
        word.trailing.push(g);
        continue;
      }
      if (word.trailing.length) {
        words.push(word);
        word = { glyphs: [], trailing: [] };
      }
      word.glyphs.push(g);
      if (g.char === '-' || g.char === '–') {
        words.push(word);
        word = { glyphs: [], trailing: [] };
      }
    }
    if (word.glyphs.length || word.trailing.length) words.push(word);

    let start = first ? firstStart : p.marL;
    let x = start;
    let lineSize = 0;
    let empty = true;
    const newLine = () => {
      sizes.push(lineSize || segment.size || p.endSize);
      first = false;
      start = p.marL;
      x = start;
      lineSize = 0;
      empty = true;
    };
    const advance = (g: Glyph) =>
      g.tab
        ? Math.max(Math.ceil((x + 0.01) / p.tabSize) * p.tabSize - x, 0)
        : g.width;

    for (const w of words) {
      const wordWidth = w.glyphs.reduce((sum, g) => sum + g.width, 0);
      if (!empty && x + wordWidth > right + 0.01) newLine();
      if (x + wordWidth > right + 0.01) {
        // longer than a whole line: break by characters
        for (const g of w.glyphs) {
          if (!empty && x + g.width > right + 0.01) newLine();
          x += g.width;
          lineSize = Math.max(lineSize, g.size);
          empty = false;
        }
      } else {
        x += wordWidth;
        w.glyphs.forEach((g) => (lineSize = Math.max(lineSize, g.size)));
        if (w.glyphs.length) empty = false;
      }
      // trailing spaces never wrap a line
      for (const g of w.trailing) {
        x += advance(g);
        lineSize = Math.max(lineSize, g.size);
        empty = false;
      }
    }
    newLine();
  }
  return sizes;
};

export const measureParagraph = (
  p: LayoutParagraph,
  width: number,
): ParagraphMetrics => {
  const sizes = breakLines(p, width);
  const firstSize = sizes[0] ?? p.endSize;
  return {
    before: spacingOf(p.before, firstSize),
    body: sizes.reduce((sum, size) => sum + lineHeight(size, p.lineSpacing), 0),
    after: spacingOf(p.after, sizes[sizes.length - 1] ?? firstSize),
    lines: sizes.length,
  };
};

/** Height of paragraphs stacked in one text box. */
export const stackHeight = (paragraphs: ParagraphMetrics[]): number =>
  paragraphs.reduce(
    (sum, p, i) =>
      sum +
      (i > 0 ? p.before : 0) +
      p.body +
      (i < paragraphs.length - 1 ? p.after : 0),
    0,
  );
