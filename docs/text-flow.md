---
title: Text Flow
description: Let the entries of a list flow through several text boxes, column by column, measured with the fonts' real metrics.
---

A list that is too long for one text box often has to continue in a second
and a third one, like the columns of a newspaper. Write the same loop into
every text box and wrap the list in `flow()`:

```text
{{FOR item IN flow(items)}}
{{item.title}}
{{item.text}}
{{END-FOR}}
```

```ts
import { modify } from 'pptx-templates';

const items = [{ title: 'Apples', text: 'Grown in the valley …' }];

pres.addSlide('Columns.pptx', 1, (slide) => {
  slide.modify(modify.renderTemplate({ items }));
});
```

`flow()` fills the text boxes in reading order: the leftmost box first, and
of boxes that sit on top of each other the upper one first. An entry is never
split: it goes into the first box that still has room for all of its
paragraphs. Boxes that are not needed stay empty.

Nothing has to be configured. How much fits into a box comes from the slide
itself: the size of the box and its margins, the font, size and weight of
each run, line spacing, space before and after paragraphs, indents and
bullets, including what the shape inherits from its layout, master and
theme. The boxes may be formatted differently and have different widths.

## How it measures

Before the boxes are filled, every entry is rendered once into a copy of each
box. Its paragraphs are broken into lines with the advance widths from the
font file, the way PowerPoint sets text: at spaces and after hyphens, a word
longer than the line by characters, and a line height of 1.2 × the font size
(× the line spacing). Text around the loop, such as a heading above it,
counts as well. The copies are thrown away; `EXEC` assignments of these trial
renders are undone, so a counter in the loop counts every entry once.

Measuring a slide with 24 entries in three boxes takes about 0.2 seconds.

## Fonts

The fonts are read from the system font folders (Windows, macOS, Linux).
Where a font is not installed, a font with the same metrics is used: Liberation
Sans or Arimo for Arial, Liberation Serif or Tinos for Times New Roman,
Liberation Mono or Cousine for Courier New, Carlito for Calibri, Caladea for
Cambria. On a Linux server without the Microsoft fonts, install these (Debian
and Ubuntu: `fonts-liberation fonts-crosextra-carlito fonts-crosextra-caladea`).

Fonts that are in neither place can be given as files or folders:

```ts
import { modify } from 'pptx-templates';

const items = [{ title: 'Apples', text: 'Grown in the valley …' }];

pres.addSlide('Columns.pptx', 1, (slide) => {
  slide.modify(
    modify.renderTemplate({ items }, { fonts: ['assets/fonts'] }),
  );
});
```

If a font cannot be found at all, `flow()` logs a warning and measures with
Liberation Sans, Arial or another sans-serif font it finds instead.

## Lists longer than one slide

When the boxes are full, the list continues on a copy of the slide: the same
boxes, filled with the next entries. `renderDeck` (and
`pptx-templates render`) adds these copies by itself, right after the slide:

```ts
import fs from 'fs';
import { renderDeck } from 'pptx-templates';

const items = [{ title: 'Apples', text: 'Grown in the valley …' }];

fs.writeFileSync('output/Columns.pptx', await renderDeck('templates/Columns.pptx', { items }));
```

Everything else on the slide repeats as well. `flowPage` (1, 2, …) and
`flowPages` tell the copies apart:

```text
Products {{IF flowPage > 1}}(continued){{END-IF}}
Page {{flowPage}} of {{flowPages}}
```

With `addSlide`, the deck is planned before anything is rendered, so ask
first how many slides the list needs and add them yourself. `flowPage` (0-based)
picks the part of the list each one shows:

```ts
import { countFlowPages, modify } from 'pptx-templates';

const items = [{ title: 'Apples', text: 'Grown in the valley …' }];
const data = { items };

const pages = await countFlowPages('templates/Columns.pptx', 1, data);
for (let page = 0; page < pages; page++) {
  pres.addSlide('Columns.pptx', 1, (slide) => {
    slide.modify(modify.renderTemplate(data, { flowPage: page }));
  });
}
```

`flowRest(list)` returns the entries after the current slide, for example
for a note at the bottom:

```text
{{IF flowRest(items).length}}{{flowRest(items).length}} more on the next slide{{END-IF}}
```

Rendered without `flowPage` (a single `addSlide`), the entries that need
another slide are left out and a warning is logged. An entry too large for
every box is left out as well, also by `renderDeck`. `flowOverflow: 'error'`
turns these warnings into errors:

```ts
import { modify } from 'pptx-templates';

const items = [{ title: 'Apples', text: 'Grown in the valley …' }];

pres.addSlide('Columns.pptx', 1, (slide) => {
  slide.modify(
    modify.renderTemplate({ items }, { flowOverflow: 'error' }),
  );
});
```

## Rules for the template

- `{{FOR … IN flow(…)}}` and `{{END-FOR}}` each need a paragraph of their own,
  and the loop must repeat whole paragraphs. A loop inside one paragraph
  (`{{FOR i IN flow(items)}}{{i}}, {{END-FOR}}`) is rejected.
- All boxes of one list write the same list: `flow(items)` in every box. A
  slide can have several lists, each with its own boxes.
- The boxes need a fixed size. "Shrink text on overflow" is measured with the
  scale PowerPoint stored in the file; "Resize shape to fit text" is ignored,
  the box keeps its size.
- Not supported: vertical text, text that does not wrap, columns inside one
  text box, and `flow()` in tables.
- Use `slide.modify(modify.renderTemplate(…))`. With `modifyElement()`, the
  layout and master of the slide are not available, and inherited formatting
  falls back to defaults.

## Good to know

- The measurement follows PowerPoint's rules, and the results were checked
  against LibreOffice. PowerPoint and LibreOffice can break a line one word
  apart now and then; where a line is a fraction of a point too wide,
  `flow()` breaks it, so a box rather ends up a little emptier than
  overflowing. Leave a little room at the bottom of the boxes if a line more
  matters.
- Hyphenation, kerning and right-to-left or East Asian line breaking rules
  are not taken into account.
- `flow`, `flowRest`, `flowPage` and `flowPages` are built in. The same name
  in your data or in `additionalJsContext` takes precedence.
- To count the slides, `renderDeck` and `countFlowPages` render the slide
  once more. Functions in your data are called again then.
