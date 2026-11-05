---
title: Template Logic
description: Conditions, loops, variables, HTML and links written directly into slide text, in the style of docx-templates.
---

`modify.renderTemplate` goes further than `replaceText`. You write the logic
straight into the text of your slide, in the style of
[docx-templates](https://github.com/guigrpa/docx-templates), and pass the data
when you build the deck:

```text
Hello {{user.name}}!
{{IF user.orders > 10}}
Thank you for being a loyal customer.
{{ELSE}}
Have a look at our new offers.
{{END-IF}}
{{FOR item IN items}}
{{$idx + 1}}. {{item.name}}: {{item.price.toFixed(2)}} €
{{END-FOR}}
```

```ts
import { modify } from 'pptx-templates';

const data = {
  user: { name: 'Ada', orders: 12 },
  items: [
    { name: 'Apples', price: 1.5 },
    { name: 'Pears', price: 2 },
  ],
  upper: (text: string) => text.toUpperCase(),
};

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  // A single shape (or table) ...
  slide.modifyElement('setText', modify.renderTemplate(data));
});

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  // ... or every shape and table on the slide
  slide.modify(modify.renderTemplate(data));
});
```

## Commands

| Tag | Effect |
|---|---|
| `{{expr}}`, `{{= expr}}`, `{{INS expr}}` | Insert the value. `null`/`undefined` insert nothing, arrays are joined with `, `, and `\n` becomes a line break |
| `{{IF expr}}` … `{{ELSE-IF expr}}` … `{{ELSE}}` … `{{END-IF}}` | Keep the first branch whose condition is truthy |
| `{{FOR item IN expr}}` … `{{END-FOR}}` | Repeat the content for each entry. Inside the loop, `item` and `$item` are the entry and `$idx` is the 0-based index |
| `{{EXEC name = expr; other = expr}}`, `{{! …}}` | Set variables for the rest of the slide. Nothing is inserted |
| `{{ALIAS name command}}` … `{{*name}}` | Define a shorthand for a command, e.g. `{{ALIAS fullName INS user.first + ' ' + user.last}}` |
| `{{HTML expr}}` | Insert HTML as formatted paragraphs (bold, italic, lists, links, …) through the [HTML converter](./text.md) |
| `{{LINK expr}}` | Insert a hyperlink. `expr` is a URL or `({ url, label })`. A number as `url` links to that slide of the output |
| `{{IMAGE expr}}` | Place a picture where the shape is (see [Images](#images)). Needs `slide.modify` |

Keywords are case-insensitive. A block can sit inside one paragraph
(`Status: {{IF ok}}fine{{ELSE}}broken{{END-IF}}`) or span paragraphs. A
paragraph that holds nothing but control tags (`IF`, `ELSE`, `END-FOR`,
`EXEC`, `ALIAS`, `HTML`, …) is removed from the output, so you can put each
tag on its own line, as in the example above.

Variables set with `EXEC` are visible in all later tags of the same render
call, also in later shapes, and inside loops they update the variable of the
enclosing scope:

```text
{{EXEC total = 0}}
{{FOR item IN items}}{{! total = total + item.price}}{{END-FOR}}
Total: {{total.toFixed(2)}}
```

Your `data` is never changed: an `EXEC` that assigns a name from `data`
defines a new variable that hides it.

## Tables

If a block opens in a cell and is not closed in that same cell, it spans
table rows. `{{FOR row IN rows}}` in the first cell and `{{END-FOR}}` in the
last cell of a row repeats that row once per entry. An `IF` placed the same
way keeps or removes the row. A row that holds nothing but control tags is
dropped. `ELSE` cannot span rows. Merged cells (`rowSpan`) are not adjusted,
and the table frame's height is set to the sum of its row heights.

## Images

`{{IMAGE …}}` takes an object like in docx-templates:

```ts
import fs from 'fs';
import { modify } from 'pptx-templates';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  // IMAGE writes media files, so it only works through slide.modify()
  slide.modify(
    modify.renderTemplate({
      logo: fs.readFileSync('images/logo.png'),
      photos: ['data:image/jpeg;base64,/9j/4AAQ…'],
    }),
  );
});
```

```text
{{IMAGE ({ data: logo, extension: '.png', width: 4, alt: 'Logo' })}}
{{FOR p IN photos}}{{IMAGE ({ data: p })}}{{END-FOR}}
{{IMAGE ({ data: logo, extension: '.png', shape: 'ellipse', border: { color: '#1f4e79', width: 2 } })}}
```

| Property | Effect |
|---|---|
| `data` | `Buffer`, `ArrayBuffer`, base64 string or data URL (`data:image/png;base64,…`) |
| `extension` | `.png`, `.jpg`, `.jpeg` or `.gif`. Optional for data URLs |
| `width`, `height` | Size in cm. With only one of them, the other keeps the aspect ratio. Without both, the picture fits into the shape, keeping its aspect ratio |
| `alt` | Alternative text |
| `rotation` | Rotation in degrees, clockwise |
| `shape` | Shape the picture is cropped to: any DrawingML preset, e.g. `'ellipse'`, `'heart'`, `'star5'` (see [Shape Geometry](./shape-geometry.md)). Default `'rect'`, or `'roundRect'` with `cornerRadius` |
| `cornerRadius` | Corner radius on the shorter side: a number in cm, or `'10%'`, `'12px'`, `'4pt'`. At most `'50%'`. Only for `roundRect`, `round1Rect`, `round2SameRect`, `round2DiagRect`, `snipRoundRect` |
| `shapeAdjustments` | Raw adjust values of the shape, e.g. `{ adj: 25000 }`. Win over `cornerRadius` |
| `border` | `{ color, width }`: hex color (`'#1f4e79'`) and width in pt |

PPTX text cannot contain pictures, so the picture is placed **at the shape
that holds the tag**, centered in its frame. Several pictures from one shape,
for example from a `FOR` loop, are placed side by side, each centered in an
equal share of the frame. If no text is left in the shape after rendering,
the shape is removed and only the pictures remain.

- The shape needs its own position and size. Placeholders that inherit
  them from the layout are not supported.
- `IMAGE` in table cells is not supported.
- Identical images are stored only once in the output.
- `IMAGE` writes a media file, so it needs `slide.modify(modify.renderTemplate(…))`.
  A `modifyElement()` callback cannot add files and rejects with an error.

## Expressions

Expressions support property paths (`a.b`, `a[0]`, `a['key']`, `a?.b`),
string/number/boolean literals, object and array literals
(`({ url: x, label: 'Home' })`, `[1, 2]`), arrow functions
(`items.filter(i => i.active)`, `(a, b) => a - b`), `! - + * / %`,
comparisons, `== != === !==`, `&& || ??` and `cond ? a : b`. You can call
your own functions (passed in `data`) and methods of the values
(`name.toUpperCase()`).

They are evaluated by a small built-in interpreter, not with `eval`, so a
template cannot run arbitrary code: there are no assignments outside `EXEC`,
no `new`, and keys like `constructor`, `prototype` or `__proto__` are
rejected. Typographic quotes inserted by PowerPoint's AutoCorrect (`“Ada”`)
are valid string delimiters.

Wrap an object literal in parentheses: `{{LINK ({ url: x })}}`. Without them,
the `}}` of the object closes the tag.

## Async data

`data` can be a Promise, and functions in `data` can return Promises. The
result of every tag is awaited, so `{{FOR p IN loadProjects()}}` works. To use
properties of an async result, store it first:
`{{EXEC user = loadUser()}}{{user.name}}`.

## Options

```ts
import { modify } from 'pptx-templates';

pres.addSlide('TextReplace.pptx', 1, (slide) => {
  slide.modify(
    modify.renderTemplate(
      { price: 2 },
      {
        cmdDelimiter: ['[[', ']]'], // or openingTag / closingTag
        additionalJsContext: {
          money: (n: number) => `${n.toFixed(2)} €`,
        },
        rejectNullish: true,
        errorHandler: (error, command) => `#ERR ${command}`,
      },
    ),
  );
});
```

| Option | Default | Effect |
|---|---|---|
| `openingTag`, `closingTag` | `{{`, `}}` | Tag delimiters |
| `cmdDelimiter` | | docx-templates spelling: one string for both sides (`'+++'`) or `[open, close]` |
| `additionalJsContext` | | Extra helpers and constants. `data` takes precedence |
| `strict` | `false` | Throw on unknown top-level names instead of treating them as `undefined` |
| `rejectNullish` | `false` | Throw when an insertion is `null` or `undefined` |
| `errorHandler(error, command)` | | Called when a command fails. Its return value is used instead (inserted text, condition, list, …) |
| `failFast` | `true` | `false`: keep going and throw one error listing all failed commands at the end |
| `processLineBreaks` | `true` | Turn `\n` in inserted values into line breaks |

## Listing the commands of a template

`listTemplateCommands` reads a .pptx and returns every tag, slide by slide,
the way `listCommands` does in docx-templates. Use it to check a template, or
to find out which data it expects:

```ts
import { listTemplateCommands } from 'pptx-templates';

const commands = await listTemplateCommands('templates/Offer.pptx');
// [{ slide: 1, shape: 'Title', raw: '{{IF user.vip}}', type: 'IF', code: 'user.vip' }, …]
```

## Good to know

- PowerPoint often splits text into several runs, even inside a tag (spell
  check, partial edits). Tags are joined across runs, and the inserted value
  takes the formatting of the run where the tag starts.
- An unknown name evaluates to `undefined`, so `{{IF user.vip}}` is simply
  false when `vip` is missing.
- Errors (an unclosed `{{IF}}`, an `END-FOR` without `FOR`, invalid syntax)
  reject `write()` with an error message that names the tag.
- Text inserted from your data is never rendered again, so a value that
  contains `{{` stays literal.
