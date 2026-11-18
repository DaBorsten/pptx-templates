---
title: Hide Slides
description: Leave a slide out of the output when a condition is true, with a tag in its speaker notes.
---

Some slides only make sense with the right data: a slide for further
projects when there are more than two, a discount slide for VIP customers.
Write `{{HIDE-SLIDE condition}}` into the **speaker notes** of the slide,
and `renderDeck` leaves the slide out when the condition is true:

```text
{{HIDE-SLIDE projects.length <= 2}}
```

```ts
import fs from 'fs';
import { renderDeck } from 'pptx-templates';

const projects = [{ name: 'Portal' }, { name: 'App' }];

// The slide with the tag in its notes is not in the output
fs.writeFileSync(
  'output/Projects.pptx',
  await renderDeck('templates/Projects.pptx', { projects }),
);
```

The condition is an [expression](./template-logic.md#expressions) like the
one of `{{IF …}}`, with the data, `additionalJsContext` and the options of
the render (`strict`, `fixSmartQuotes`, other delimiters, `errorHandler`).
`pptx-templates render` on the command line does the same.

- `{{HIDE-SLIDE}}` without a condition always hides the slide, for example
  a slide that is still being worked on.
- With several tags in the notes, the slide is left out as soon as one of
  them is true.
- The tags are removed from the notes of the output; the rest of the notes
  stays. A paragraph that held nothing but the tag goes as well.
- A hidden slide leaves nothing behind in the output file: neither its
  notes nor its embedded objects.
- When every slide is hidden, the output has no slides.

## Why the speaker notes

The tag decides about the whole slide, not about a piece of text, so it
belongs to the slide once. The notes exist once per slide and are not part
of the design. On the slide itself, in a text box or a table,
`{{HIDE-SLIDE …}}` is removed with a warning and does nothing; the
[linter](./template-lint.md) reports it there as `unsupported-location`.
Other tags in the notes are not rendered and stay as they are.

## Copies of a flow() slide

A slide whose [`flow()`](./text-flow.md) list needs more than one slide is
repeated by `renderDeck`. The condition is evaluated for each copy, with
`flowPage` (1, 2, …) and `flowPages`:

```text
{{HIDE-SLIDE flowPage > 3}}
```

keeps the first three copies of the slide and leaves out the rest.
Each copy has its own speaker notes and its own copy of the slide's
embedded objects (OLE objects, embedded Excel or Word documents, charts
with their workbooks).

## With addSlide

`renderTemplate` renders one slide and cannot take it out of the deck.
With your own `addSlide` calls, ask `isSlideHidden` first and skip the
slide:

```ts
import { isSlideHidden, modify } from 'pptx-templates';

const data = { projects: [{ name: 'Portal' }] };

if (!(await isSlideHidden('templates/Projects.pptx', 2, data))) {
  pres.addSlide('Projects.pptx', 2, (slide) => {
    slide.modify(modify.renderTemplate(data));
  });
}
```

It takes the options of `renderTemplate`, plus `flowPage` (0-based) and
`flowPages` for a condition that depends on them. `addSlide` copies the
notes as they are, tag included.

## Good to know

- An error in the condition rejects `renderDeck` with a
  `TemplateRenderError` that starts with the slide number
  (`slide 2: {{HIDE-SLIDE check()}}: …`), unless `errorHandler` returns a
  value for it.
- The tag must be in one paragraph of the notes. PowerPoint may split it
  into several runs (spell check); that is fine.
- `flow()` and `flowRest()` in the condition do not distribute anything:
  `flow(list)` is the whole list there.
