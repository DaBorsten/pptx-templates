# pptx-templates

[![npm version](https://img.shields.io/npm/v/pptx-templates.svg)](https://www.npmjs.com/package/pptx-templates)
[![CI](https://github.com/DaBorsten/pptx-templates/actions/workflows/ci.yml/badge.svg)](https://github.com/DaBorsten/pptx-templates/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/DaBorsten/pptx-templates/blob/main/LICENSE)

**Fill PowerPoint templates with data.** Design the slide in PowerPoint, type
`{{tags}}` into it, and render it from JavaScript or TypeScript.

```text
Offer for {{customer.name}}                  Offer for Ada
{{FOR item IN items}}                    →   Apples: 1.50 €
{{item.name}}: {{item.price.toFixed(2)}} €   Pears: 2.00 €
{{END-FOR}}
```

The library edits your .pptx instead of rebuilding it. Fonts, colors,
positions, masters and layouts stay exactly as the designer left
them, and the designer can change the template later without touching your
code.

- **Logic in the slide**: values, `IF`/`ELSE`, `FOR` loops over paragraphs and
  table rows, variables, HTML, links and images, all written as text in
  PowerPoint
- **One slide per record**: render the same template slide for every customer,
  project or month into one deck
- **A linter for templates**: `npx pptx-templates lint` finds broken tags,
  unclosed blocks and misspelled field names before you render, also in CI and
  in watch mode while the deck is open in PowerPoint
- **Safe expressions**: tags run in a small interpreter, not `eval`, and can't
  reach Node.js
- **Charts, tables and images from code** when a tag is not enough
- TypeScript types included, runs on Node.js 20+ and Bun, MIT licensed

## Installation

With npm:

```bash
npm install pptx-templates
```

With Bun:

```bash
bun add pptx-templates
```

## Quick start

**1. Write the template.** Put this into a text box on slide 1 of
`templates/Offer.pptx` and format it however you like:

```text
Offer for {{customer.name}}
{{IF customer.vip}}
Thank you for being a loyal customer.
{{ELSE}}
Have a look at our new offers.
{{END-IF}}
{{FOR item IN items}}
{{$idx + 1}}. {{item.name}}: {{item.price.toFixed(2)}} €
{{END-FOR}}
```

**2. Render it.**

```ts
import Automizer, { modify } from 'pptx-templates';

const pres = new Automizer({
  templateDir: 'templates',
  outputDir: 'output',
  removeExistingSlides: true,
})
  .loadRoot('Offer.pptx')
  .load('Offer.pptx', 'offer');

const data = {
  customer: { name: 'Ada', vip: true },
  items: [
    { name: 'Apples', price: 1.5 },
    { name: 'Pears', price: 2 },
  ],
};

pres.addSlide('offer', 1, (slide) => {
  // renders every text box and table on the slide
  slide.modify(modify.renderTemplate(data));
});

await pres.write('Offer-Ada.pptx');
```

**3. Open `output/Offer-Ada.pptx`.** The text box now reads:

```text
Offer for Ada
Thank you for being a loyal customer.
1. Apples: 1.50 €
2. Pears: 2.00 €
```

Lines that held only control tags (`IF`, `END-IF`, `FOR`, …) are gone, and
every inserted value has the formatting of the tag it replaced.

## One slide per record

Add the same template slide once per entry. Each copy gets its own data:

```ts
const customers = [
  { name: 'Ada', vip: true },
  { name: 'Grace', vip: false },
];

for (const customer of customers) {
  pres.addSlide('offer', 1, (slide) => {
    slide.modify(modify.renderTemplate({ customer, items: [] }));
  });
}

await pres.write('All-Offers.pptx');
```

Load several template files and mix their slides to build a full deck: a
title slide from one file, a chart slide from another, a closing slide from a
third.

## Tags

| Tag | Effect |
|---|---|
| `{{expr}}` | Insert a value. `\n` becomes a line break, arrays are joined with `, ` |
| `{{IF expr}}` … `{{ELSE-IF expr}}` … `{{ELSE}}` … `{{END-IF}}` | Keep the first branch whose condition is true |
| `{{FOR item IN list}}` … `{{END-FOR}}` | Repeat paragraphs or table rows. `$idx` is the index |
| `{{EXEC total = 0}}`, `{{! total = total + x}}` | Set variables for the rest of the slide |
| `{{ALIAS name command}}`, `{{*name}}` | Shorthand for a command you use often |
| `{{HTML expr}}` | Insert HTML as formatted text: bold, italic, lists, links |
| `{{LINK expr}}` | Insert a hyperlink, to a URL or to a slide of the output |
| `{{IMAGE expr}}` | Place a picture where the text box is, optionally cropped to a shape with a border |
| `{{HIDE-SLIDE expr}}` | In the speaker notes: leave the slide out of the deck when `expr` is true |

Expressions are JavaScript without the dangerous parts: property paths,
comparisons, `&&`, `||`, `??`, `? :`, arrow functions
(`items.filter(i => i.active)`), methods of the values (`name.toUpperCase()`)
and functions you pass in the data.

**Table rows** repeat when `{{FOR}}` opens in the first cell and
`{{END-FOR}}` closes in the last cell of the same row:

| `{{FOR p IN products}}{{p.name}}` | `{{p.price}} €{{END-FOR}}` |
|---|---|

**Watch out for smart quotes.** `{{IF name == 'Ada'}}` needs straight
quotes, and PowerPoint turns them into `‘Ada’` as you type. Either turn
AutoCorrect off (*File › Options › Proofing › AutoCorrect Options... ›
AutoFormat As You Type > Replace as you type > "Straight quotes" with ‟smart quotes“*) or render with
`modify.renderTemplate(data, { fixSmartQuotes: true })`, which straightens
them before parsing. The linter finds tags where it already happened.

Full reference: [Template Logic](https://daborsten.github.io/pptx-templates/template-logic).

## Check templates before you render

Templates get edited in PowerPoint, often by people who never run the code.
The linter reports what they broke, slide by slide, with the wrong part
underlined:

```bash
npx pptx-templates lint templates/*.pptx --data example.json
```

```text
Offer.pptx
  slide 1 › "Price"
    error  Unterminated string in template expression "item.price.toFixed(2) + ' €"  (syntax)
    tag: {{item.price.toFixed(2) + ' €}}
                                   ~~~

1 problem(s): 1 error(s), 0 warning(s)
```

It finds unclosed `IF`/`FOR` blocks, tags split by a line break and quotes
replaced by AutoCorrect. With example data or a JSON Schema, it also finds
misspelled field names (`user.nmae` → *did you mean "name"?*).

- `--watch` lints again every time you press Ctrl+S in PowerPoint.
- The exit code fails a CI job when a template is broken.
- `npx pptx-templates usage` lists which data fields the tags read, and which
  ones no tag reads.
- `npx pptx-templates init` writes example data with every field the tags
  read, combined over several decks; run again, it only adds what is missing.
- `npx pptx-templates render deck.pptx --data data.json -o out.pptx` fills
  every slide with the data, to see the result without writing code
  (`renderDeck()` from code).
- A `pptx-templates.json` maps many decks to their data, for example one data
  file per language for every customer folder:

```json
{
  "templates": [
    { "files": "{customer}/cover_{lang}.pptx", "data": "data/cover_{lang}.json" }
  ]
}
```

Both commands are also available from code:

```ts
import { lintTemplate, listTemplateCommands } from 'pptx-templates';

const { issues, errorCount } = await lintTemplate('templates/Offer.pptx', {
  data: { customer: { name: 'Ada', vip: true }, items: [] },
});
const commands = await listTemplateCommands('templates/Offer.pptx');
```

More in [Template Linter](https://daborsten.github.io/pptx-templates/template-lint).

## When tags are not enough

Charts, images and shape geometry are easier to set from code. Every slide
callback can change single shapes by the name they have in PowerPoint's
Selection Pane:

```ts
pres.addSlide('offer', 1, (slide) => {
  slide.modifyElement('Title', modify.setText('Offer 2026'));
  slide.modifyElement('Logo', modify.setCornerRadius('10%'));
});
```

The [`modify.*` helpers](https://daborsten.github.io/pptx-templates/helpers)
fill [charts](https://daborsten.github.io/pptx-templates/charts) and
[tables](https://daborsten.github.io/pptx-templates/tables), swap
[images](https://daborsten.github.io/pptx-templates/images), crop pictures to
[any preset shape](https://daborsten.github.io/pptx-templates/shape-geometry)
and move or resize elements. For shapes that don't exist in any template,
[PptxGenJS is built in](https://daborsten.github.io/pptx-templates/generation).

## How it compares

| | pptx-templates | [PptxGenJS](https://github.com/gitbrent/PptxGenJS) | [pptx-automizer](https://github.com/singerla/pptx-automizer) |
|---|---|---|---|
| Slide design lives in | the .pptx, edited in PowerPoint | your code | the .pptx |
| Text logic | tags in the slide: `IF`, `FOR`, … | your code | a callback per shape, simple `{{tag}}` replacement |
| Repeat table rows from a list | `{{FOR}}` in the row | your code | from code |
| Template linter | yes | no | no |

Pick PptxGenJS when there is no designed template and the code should draw
every slide. Pick `pptx-templates` when someone designs the slides in
PowerPoint and your code only brings the data.

## Documentation

Guides for every feature and the API reference are at
**[daborsten.github.io/pptx-templates](https://daborsten.github.io/pptx-templates/)**.

Using an AI assistant? Point it to
[AI-INSTRUCTOR.md](https://github.com/DaBorsten/pptx-templates/blob/main/AI-INSTRUCTOR.md),
a condensed guide to the whole library that ships with the npm package.

## Contributing

Bug reports, ideas and pull requests are welcome:
[open an issue](https://github.com/DaBorsten/pptx-templates/issues/new). For a
security problem, follow the
[security policy](https://github.com/DaBorsten/pptx-templates/blob/main/SECURITY.md)
instead.

## Credits

`pptx-templates` is a fork of [pptx-automizer](https://github.com/singerla/pptx-automizer),
which copies slides and edits the OOXML underneath. The tag syntax comes from
[docx-templates](https://github.com/guigrpa/docx-templates), so if you fill
Word templates that way, you already know it.

## License

MIT
