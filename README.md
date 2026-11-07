# pptx-templates: A Powerful .pptx Modifier for Node.js

[![npm version](https://img.shields.io/npm/v/pptx-templates.svg)](https://www.npmjs.com/package/pptx-templates)
[![CI](https://github.com/DaBorsten/pptx-templates/actions/workflows/ci.yml/badge.svg)](https://github.com/DaBorsten/pptx-templates/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://github.com/DaBorsten/pptx-templates/blob/main/LICENSE)

`pptx-templates` is a Node.js-based PowerPoint (.pptx) generator that automates the manipulation of existing .pptx files. With `pptx-templates`, you can import your library of .pptx templates, merge templates, and customize slide content. `pptx-templates` will edit and merge existing pptx files. You can style template slides within PowerPoint, and these templates will be seamlessly integrated into the output presentation. Most of the content can be modified by using callbacks with [xmldom](https://github.com/xmldom/xmldom).

If you require to create elements from scratch, `pptx-templates` wraps around [PptxGenJS](https://github.com/gitbrent/PptxGenJS). Use the powerful syntax of `PptxGenJS` to add dynamic content to your existing .pptx template files.

`pptx-templates` is particularly well-suited for users who aim to manage their own library of .pptx template files, making it an ideal choice for those who work with intricate, well-designed customized layouts. With this tool, any existing slide or even a single element can serve as a data-driven template for generating output .pptx files.

## Documentation

**The full documentation lives at [daborsten.github.io/pptx-templates](https://daborsten.github.io/pptx-templates/)** — guides for every feature area plus the generated API reference:

- [Getting started](https://daborsten.github.io/pptx-templates/getting-started) · [Concepts](https://daborsten.github.io/pptx-templates/concepts) · [Selecting slides and shapes](https://daborsten.github.io/pptx-templates/selectors)
- Modify [Text](https://daborsten.github.io/pptx-templates/text) · [Tables](https://daborsten.github.io/pptx-templates/tables) · [Charts](https://daborsten.github.io/pptx-templates/charts) · [Images](https://daborsten.github.io/pptx-templates/images) · [Masters & Layouts](https://daborsten.github.io/pptx-templates/masters-layouts)
- [Generate shapes with PptxGenJS](https://daborsten.github.io/pptx-templates/generation) · [Hyperlinks](https://daborsten.github.io/pptx-templates/hyperlinks) · [Slide management](https://daborsten.github.io/pptx-templates/slide-management)
- [Helper utilities](https://daborsten.github.io/pptx-templates/helpers) · [Output](https://daborsten.github.io/pptx-templates/output) · [Requirements and limitations](https://daborsten.github.io/pptx-templates/limitations) · [Troubleshooting](https://daborsten.github.io/pptx-templates/troubleshooting)
- [API reference](https://daborsten.github.io/pptx-templates/api)

AI assistants consuming this library should read [AI-INSTRUCTOR.md](https://github.com/DaBorsten/pptx-templates/blob/main/AI-INSTRUCTOR.md), a condensed, self-contained guide that is also shipped in the npm package.

## Installation

```
$ yarn add pptx-templates
```

or

```
$ npm install pptx-templates
```

If `npm audit` reports `image-size`, see [Getting started](https://daborsten.github.io/pptx-templates/getting-started): the advisory is not reachable and an `overrides` entry silences it.

To explore the library from a clone instead, see [Getting started](https://daborsten.github.io/pptx-templates/getting-started).

## Example

```ts
import Automizer, { modify } from 'pptx-templates';

const automizer = new Automizer({
  // where your template pptx files are coming from:
  templateDir: 'my/pptx/templates',
  // where to write the final pptx output files:
  outputDir: 'my/pptx/output',
  // truncate the root presentation and start with zero slides:
  removeExistingSlides: true,
});

const pres = automizer
  .loadRoot('RootTemplate.pptx')
  .load('SlideWithGraph.pptx', 'graph')
  .load('SlideWithShapes.pptx', 'shapes');

pres
  .addSlide('graph', 1)
  .addSlide('shapes', 1, (slide) => {
    slide.modifyElement('Drum', [modify.setText('Hello world!')]);
  });

// Nothing has happened yet: all calls above only queue work.
// write() executes the whole queue and produces the output file.
const summary = await pres.write('myPresentation.pptx');
console.log(summary);
```

The single most important concept: calls like `addSlide()` and `modifyElement()` only **queue** work — including every modification callback you pass. Nothing is applied until `await pres.write(...)` (or `.stream()` / `.getJSZip()`) executes the whole queue in order. Read more in [Concepts](https://daborsten.github.io/pptx-templates/concepts).

Take a look into the [`__tests__`-directory](https://github.com/DaBorsten/pptx-templates/blob/main/__tests__) to see a lot of examples for several use cases. You will also find the example .pptx-files there.

## Support

Thanks to all contributors! You are always welcome to share code, tipps and ideas. We appreciate all levels of expertise and encourage everyone to get involved. Whether you're a seasoned pro or just starting out, your contributions are invaluable. [Get started](https://github.com/DaBorsten/pptx-templates/issues/new)

Found a security problem? Please follow the [security policy](https://github.com/DaBorsten/pptx-templates/blob/main/SECURITY.md) instead of opening a public issue.

# Special Thanks

This project was inspired by:

- [pptx-automizer](https://github.com/singerla/pptx-automizer)
- [docx-templates](https://github.com/guigrpa/docx-templates)
