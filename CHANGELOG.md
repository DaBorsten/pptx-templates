# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
semver, with the pre-1.0 convention that **breaking changes bump the minor**.

## [Unreleased]

## [0.3.4] - 2026-11-15

### Added

- `renderDeck` and `pptx-templates render`: fill the tags of every slide of
  a template with one set of data and get the new .pptx, without writing the
  `addSlide` loop. `slides` (`--slides 3,1`) keeps only some slides, in that
  order; `--helpers` loads a module whose exports the tags can call. A
  failing tag names its slide (`slide 2: {{euro(p.preis)}}: …`); the command
  exits with `1` then.
- `{{IMAGE logo}}`: the value can be the picture itself (a Buffer, a data
  URL or a file path) instead of `{ data, extension }`. A path such as
  `"images/logo.png"` is read from the new `imageRoot` option and cannot
  leave that folder; `pptx-templates render` uses the folder of the data
  file, `--image-root` another one. `extension` is optional everywhere: the
  type comes from the first bytes of the data.
- `templateDataSkeleton` lists the paths `{{IMAGE}}` places in `images`
  (also `({ data: x })` and both sides of `a ? b : c`) and gives them
  `null` instead of their path as text.

## [0.3.3] - 2026-11-14

### Added

- `templateDataSkeleton` and `pptx-templates init`: example data built from
  the tags of one or more templates, every read path with a placeholder
  (the path as text, `true` for tested values, `1` for numbers, a
  one-entry list for loops). Several decks give one combined file. An
  existing data file is completed, not replaced: its values stay, missing
  fields are added. Without file arguments, `init` writes every data file
  of `pptx-templates.json` from the decks mapped to it. Called names
  (`{{money(x)}}`) are reported as helpers instead of data.

### Changed

- Docs site: new design. One neutral grey scale with an orange accent used
  sparingly, Geist and Geist Mono (self-hosted), framed code panels with a
  quiet syntax theme, restyled sidebar, table of contents, tables and
  footer, and a matching dark mode. The sidebar is grouped into sections
  (Get started, Modify content, Templates, Build decks, Help, Reference);
  the navbar shows the version and npm/GitHub as icons. URLs, llms.txt and
  the Markdown twins are unchanged.
- npm package page links the docs site (`homepage` in `package.json`)
  instead of the GitHub README.
- README rewritten around templating in the .pptx file: tags, tables,
  linter, then the code API. New npm `description` and `keywords`.
- Removed the `npm audit` / `image-size` note from README and Getting
  started: since 0.2.1 the package no longer installs `image-size`.
- Docs: installing the package shows npm and Bun, working on a clone of
  the repository (install, dev, tests, validation) uses Bun instead of
  Yarn.

## [0.3.2] - 2026-11-13

### Added

- `templateDataUsage` and `pptx-templates usage`: which fields of the data
  the tags of a template read, which fields no tag reads, and which read
  paths are not in the data. Loop variables and EXEC names are traced back
  to the data (`produkte[].preis`); with several files, the command also
  lists the fields none of them reads. Supports `--json` and `--watch`.
- Config file `pptx-templates.json` for `lint` and `usage`: file patterns,
  each with its own data or schema. `{name}` placeholders carry parts of
  the deck path into the data path (`{customer}/cover_{lang}.pptx` →
  `data/cover_{lang}.json`). Without file arguments, the commands read
  `pptx-templates.json` from the current folder; `--watch` picks up new
  decks and folders that match.

## [0.3.1] - 2026-11-13

### Added

- `pptx-templates lint --watch` (`-w`): lints again whenever a deck or the
  `--data`/`--schema` file is saved. The deck can stay open in PowerPoint;
  saves via a temporary file and half-written files are handled.

### Security

- `pptx-templates lint`: error messages for an unreadable deck or
  `--data`/`--schema` file no longer pass control characters to the
  terminal (`JSON.parse` quotes the file content in its message).

## [0.3.0] - 2026-11-09

### Security

- `renderTemplate`: a template tag could run arbitrary code. Names inherited
  from `Object.prototype` (`{{constructor}}`, `{{__proto__}}`) were read as
  values, so a tag could reach `Object` and from there the `Function`
  constructor. Names now resolve only against the data. In addition, calling
  the `Function` constructor, `call`, `apply` and `bind` is rejected.
- `renderTemplate`: `match`, `matchAll` and `search` reject a string pattern
  (it was compiled to a regular expression; a crafted one blocked the process
  for minutes), and `repeat`, `padStart`, `padEnd` reject results over
  1,000,000 characters. **Behavior change:** `{{name.match('a')}}` now fails;
  pass a `RegExp` in the data, use `includes`, or set the new
  `allowStringPatterns` option for templates you trust. `compile` on a
  `RegExp` from the data is rejected for the same reason (it also changed the
  caller's `RegExp`).
- `renderTemplate`: `call`, `apply` and `bind` are rejected when read, so they
  cannot be passed on as a callback either (`list.map(fn.call, fn)`).
- `pptx-templates lint`: control characters in tags and shape names of the
  linted file are no longer printed to the terminal.

### Fixed

- `renderTemplate`: a paragraph holding only `{{IF …}}`, `{{END-IF}}` or
  `{{HTML …}}` left an empty paragraph in table cells.
- `renderTemplate`: `{{END-FOR}}` and a new `{{FOR …}}` in the same table row
  closed the new loop instead of the earlier one.
- `renderTemplate`: data fields named `link`, `image` or `html` were read as
  commands in `{{link + 1}}` or `{{image ? 'a' : 'b'}}`.
- `renderTemplate`: empty tags and unknown aliases now go to `errorHandler`
  and `failFast: false` like other errors, and reject with a
  `TemplateRenderError`.
- `renderTemplate`: only `__proto__` and the legacy `__define*__`/`__lookup*__`
  accessors are blocked, no longer every key starting with `__`
  (`{{user.__typename}}` works).
- `renderTemplate`: `\n`, `\t`, `\r` in string literals; numbers like `1e3`
  and `.5`.
- `listTemplateCommands`: an alias used before its definition no longer
  throws.
- `lintTemplate`: blocked keys are also reported on values that are not a
  name (`{{'x'.constructor}}`, `{{list.map(x => x.constructor)}}`).
- `lintTemplate`: no crash on a relationship without `Target` or a schema
  `$ref` with a malformed `%` escape. The CLI rejects `--data`/`--schema`
  files that are not a JSON object.

## [0.2.1] - 2026-11-07

### Security

- Bundle PptxGenJS 4.0.1 (`src/vendor/pptxgenjs/`, unmodified, MIT) instead of
  depending on the `pptxgenjs` package. Its manifest declares `image-size
  ^1.2.1`, which is affected by CVE-2025-71329 and CVE-2025-71330 (infinite
  loops in the JXL/HEIF and ICNS parsers). The PptxGenJS bundle never loads
  `image-size`, but the declared dependency could not be overridden by
  consumers of this package. `pptx-templates` no longer installs it.

### Changed

- `pptxgenjs` is no longer a dependency; `slide.generate()` works as before.
  To update the bundled copy, replace `index.js` and `index.d.ts` in
  `src/vendor/pptxgenjs/`.

## [0.2.0] - 2026-11-07

### Added

- `lintTemplate(pptx, options?)` checks the `{{…}}` tags of a template before
  rendering: syntax errors (with hints for quotes and characters replaced by
  PowerPoint's AutoCorrect and for misspelled commands like `{{END IF}}`),
  blocked keys (`constructor`, `__proto__`), unclosed or unmatched `IF`/`FOR`/`ELSE`, tags
  split by a line break, tags where `renderTemplate` cannot render them
  (layouts, masters, notes, charts, SmartArt, `IMAGE` in tables or in
  placeholders without their own position) and, given example `data` or a
  JSON `schema`, names and property paths that are not in the data
  (`src/extensions/template-lint/`). Each issue carries the `range` of the
  wrong characters in the tag; the CLI underlines it with `~~~`.
- CLI: `npx pptx-templates lint deck.pptx [--data data.json] [--schema schema.json]`
  (`src/extensions/cli/`).
- `renderTemplate` and `listTemplateCommands` leave `{{each:…}}` markers of
  `modify.setTable`'s `expand` (nested tables) as text instead of failing on
  them; `lintTemplate` skips them too.

### Documentation

- Getting started explains the `image-size` advisories `npm audit` reports via
  `pptxgenjs` (not reachable) and the `overrides` entry that silences them.

### Changed

- **Breaking:** strings in template expressions need straight quotes
  (`'Ada'`, `"Ada"`). Typographic quotes from PowerPoint's AutoCorrect
  (`“Ada”`, `‘Ada’`, `„Ada“`) were accepted before and are now a syntax
  error, as in docx-templates. Switch off AutoCorrect's "smart quotes" and
  retype them, or render with the new `fixSmartQuotes: true` option
  (docx-templates compatible). `lintTemplate` finds the affected tags.

## [0.1.3] - 2026-11-05

### Security

- Fixed CodeQL findings: a relationship filename like `a000…0` no longer
  triggers exponential regex backtracking, and the corner-radius parser no
  longer runs in polynomial time on long whitespace input. Relative targets
  (`../media/…`) are resolved with an anchored pattern.

## [0.1.2] - 2026-11-05

### Added

- `modify.setGeometry(shape, adjustments?)` crops a picture or shape to any
  DrawingML preset geometry (*Crop to Shape*), e.g. `'ellipse'`, `'heart'`.
- `modify.setCornerRadius(radius, shape?)` rounds corners; the radius is a
  number in cm or a string in `%`, `px`, `pt`, `cm` or `mm`.
- `{{IMAGE}}` takes `shape`, `cornerRadius`, `shapeAdjustments` and `border`,
  as in docx-templates (`src/extensions/shape-geometry/`).

## [0.1.1] - 2026-11-04

### Fixed

- `modify.replaceText` now finds `{{tags}}` that PowerPoint split over several
  runs (spell check `err="1"`, edit history). Before the search, a split tag
  is moved into the run it starts in; text before and after the tag keeps its
  own run and formatting (`src/extensions/run-merge/`).

## [0.1.0]

First release as `pptx-templates`, a fork of
[pptx-automizer](https://github.com/singerla/pptx-automizer). It contains
everything in pptx-automizer up to upstream commit `2ef0cf7` (2026-10-14,
after 0.9.4). See the
[upstream changelog](https://github.com/singerla/pptx-automizer/blob/main/CHANGELOG.md)
for that history.

### Breaking

- pptxgenjs 3.12 → 4.0.1. pptxgenjs 4 loads `node:fs` and `node:https` with
  a dynamic `import()` on every export. Projects that test `slide.generate()`
  with **Jest** must run Jest with `node --experimental-vm-modules`;
  otherwise the write fails and Jest reports an unhandled rejection. Plain
  Node is not affected.

### Added

- `modify.renderTemplate(data, options?)` renders docx-templates-style logic
  written into slide text: `{{expr}}`/`{{= …}}`/`{{INS …}}`,
  `IF`/`ELSE-IF`/`ELSE`/`END-IF`, `FOR … IN …`/`END-FOR` (also across table
  rows), `EXEC`/`!`, `ALIAS`/`*alias`, `HTML`, `LINK` and `IMAGE`. `IMAGE`
  places pictures at the shape that holds the tag (via `slide.modify`) and
  stores identical media only once. Expressions are evaluated by a built-in
  interpreter (no `eval`). They support arrow functions, object and array
  literals, optional chaining and async data. Options: `cmdDelimiter`,
  `openingTag`/`closingTag`, `additionalJsContext`, `strict`,
  `rejectNullish`, `errorHandler`, `failFast`, `processLineBreaks`.
- `listTemplateCommands(pptx, options?)` lists the template commands of a
  .pptx per slide and shape, like `listCommands` in docx-templates.

### Fixed

- pptxgenjs 4 writes the SVG bytes into the PNG fallback of an SVG image
  added via `slide.generate()`, and the media deduplication then merged both
  into one `.png` part. The fallback now gets the same placeholder PNG that
  pptxgenjs 3.12 wrote.

[Unreleased]: https://github.com/DaBorsten/pptx-templates/compare/v0.3.4...HEAD
[0.3.4]: https://github.com/DaBorsten/pptx-templates/compare/v0.3.3...v0.3.4
[0.3.3]: https://github.com/DaBorsten/pptx-templates/compare/v0.3.2...v0.3.3
[0.3.2]: https://github.com/DaBorsten/pptx-templates/compare/v0.3.1...v0.3.2
[0.3.1]: https://github.com/DaBorsten/pptx-templates/compare/v0.3.0...v0.3.1
[0.3.0]: https://github.com/DaBorsten/pptx-templates/compare/v0.2.1...v0.3.0
[0.2.1]: https://github.com/DaBorsten/pptx-templates/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/DaBorsten/pptx-templates/compare/v0.1.3...v0.2.0
[0.1.3]: https://github.com/DaBorsten/pptx-templates/compare/v0.1.2...v0.1.3
[0.1.2]: https://github.com/DaBorsten/pptx-templates/compare/v0.1.1...v0.1.2
[0.1.1]: https://github.com/DaBorsten/pptx-templates/compare/v0.1.0...v0.1.1
[0.1.0]: https://github.com/DaBorsten/pptx-templates/releases/tag/v0.1.0
