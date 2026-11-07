---
title: Template Linter
description: Find broken template tags, unclosed blocks, unknown variables and AutoCorrect damage before you render.
---

A mistake in a [template](./template-logic.md) usually shows up only when you
render it: an `{{IF}}` without `{{END-IF}}`, a typo in a variable name, or a
quote that PowerPoint's AutoCorrect turned into `’`. `lintTemplate` reads a
.pptx and reports these problems up front, slide by slide and shape by shape.

```ts
import { lintTemplate } from 'pptx-templates';

const { issues, errorCount } = await lintTemplate('templates/Offer.pptx', {
  data: { user: { name: 'Ada', vip: true }, items: [{ name: 'Apples', price: 1.5 }] },
});
// [{ severity: 'warning', rule: 'unknown-property', slide: 2, shape: 'Title',
//    tag: '{{user.nmae}}', range: { start: 7, end: 11 },
//    message: '"nmae" does not exist on "user" (object); did you mean "name"?',
//    part: 'ppt/slides/slide2.xml' }, …]
```

`range` marks the characters of `tag` the issue is about (0-based, `end`
exclusive): the unterminated string, the unexpected character, the unknown
property. Where there is nothing more precise, for an unclosed block, say,
it covers the whole tag.

## Command line

```bash
npx pptx-templates lint deck.pptx
npx pptx-templates lint templates/*.pptx --data example.json
npx pptx-templates lint deck.pptx --schema data.schema.json --max-warnings 0
npx pptx-templates lint deck.pptx --json
```

The output underlines the part of the tag that is wrong:

```text
deck.pptx
  slide 1 › "Textfeld 3"
    error  Unterminated string in template expression "euro(p.preis) + 'Test`"  (syntax)
    tag: {{euro(p.preis) + 'Test`}}
                           ~~~~~~
```

In a terminal, errors and the underline are red, warnings yellow, and the
rule name and `tag:`/`hint:` labels dimmed. Colors are off when the output is
piped, when `NO_COLOR` is set, or with `--no-color`; `--color` or
`FORCE_COLOR=1` turns them on, for CI logs that render ANSI.

The exit code is `0` without errors, `1` with errors (or more warnings than
`--max-warnings`) and `2` for an invalid call or an unreadable file, so the
command can run in CI. `--opening-tag` and `--closing-tag` set other
delimiters.

## What it checks

| Rule | Severity | Finds |
|---|---|---|
| `syntax` | error | Tags that cannot be parsed. The hint names the usual cause: quotes or characters replaced by AutoCorrect (`(c)` → `©`, `...` → `…`, `->` → `→`), `=` instead of `==`, `AND` instead of `&&`, a misspelled command (`{{END IF}}`, `{{FOR x items}}`). Also blocked keys like `constructor` or `__proto__`, which `renderTemplate` would only reject while rendering |
| `unterminated-tag` | error | `{{` without `}}` or the other way round. Often a tag split by a line break (Shift+Enter) or by Enter. Such text stays in the output as it is |
| `unclosed-block` | error | `IF`/`FOR` without `END-IF`/`END-FOR` in the same shape or table |
| `unmatched-tag` | error | `END-IF`, `END-FOR`, `ELSE`, `ELSE-IF` without an opening tag, or an `ELSE` after `ELSE` |
| `unknown-alias` | error | `{{*name}}` without `{{ALIAS name …}}` on the slide |
| `unsupported-location` | error | `IMAGE` in a table cell, `IMAGE` in a placeholder that inherits its position from the layout, `ELSE` that spans table rows |
| `unsupported-location` | warning | Tags in slide layouts, masters, speaker notes, charts and SmartArt. `renderTemplate` does not render them, so they stay as they are |
| `unknown-variable` | warning | A name that is not in `data`/`schema`/`additionalJsContext` and not defined by `FOR`, `EXEC` or an arrow function |
| `unknown-property` | warning | A property path that does not exist, including methods: `user.name.lenght` on a string |
| `not-a-list` | warning | `FOR` over a value that is not a list in the data |

The last three run only with `data` or `schema`. Without them, the linter
cannot know which names your data has. Names inside `FOR` loops are checked
against the first entry of the list.

Markers like `{{each:row}}`, which `modify.setTable` uses to expand nested
tables (`expand: [{ tag: '{{each:row}}', … }]`), are not template tags. The
linter skips them, and `renderTemplate` leaves them in the text for
`setTable`.

## Example data or JSON Schema

`data` is an example of what you pass to `renderTemplate`. A property that is
missing in the example is reported, so the example should contain every
optional property, too. A property that is `null` in the example is not
checked further.

`schema` is a JSON Schema of the data. `$ref` to local definitions, `anyOf`,
`oneOf`, `allOf`, `items`, `prefixItems`, `patternProperties` and
`additionalProperties` are understood. An object with `properties` counts as
closed unless it sets `additionalProperties`: a name that is not listed is
reported, which is what catches typos. An object without `properties`
accepts any name.

```ts
import { lintTemplate } from 'pptx-templates';

const result = await lintTemplate('templates/Offer.pptx', {
  schema: {
    type: 'object',
    properties: {
      user: { type: 'object', properties: { name: { type: 'string' } } },
      items: { type: 'array', items: { $ref: '#/$defs/item' } },
    },
    $defs: {
      item: { type: 'object', properties: { name: { type: 'string' }, price: { type: 'number' } } },
    },
  },
  additionalJsContext: { money: (n: number) => `${n.toFixed(2)} €` },
});
```

| Option | Effect |
|---|---|
| `data` | Example data; turns on the name checks |
| `schema` | JSON Schema of the data; can be combined with `data` |
| `additionalJsContext` | Same as in `renderTemplate`: helper names that are not in the data |
| `openingTag`, `closingTag`, `cmdDelimiter` | Other tag delimiters, as in `renderTemplate` |
| `fixSmartQuotes` | Accept typographic quotes, as `renderTemplate` does with the same option. CLI: `--fix-smart-quotes` |

## Quotes from AutoCorrect

Strings in tags need straight quotes (`{{IF name == 'Ada'}}`). PowerPoint
replaces typed quotes with typographic ones (`{{IF name == ’Ada’}}`), which
`renderTemplate` rejects. The linter reports such a tag as a `syntax` error
before you render. Switch the replacement off once in *File › Options ›
Proofing › AutoCorrect Options › AutoFormat As You Type › "Straight quotes"
with "smart quotes"*, then type the quotes in the tag again.

If you render with `fixSmartQuotes: true`, lint with the same option
(`--fix-smart-quotes` on the command line), so the linter accepts the same
tags as the renderer.
