# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versions follow
semver, with the pre-1.0 convention that **breaking changes bump the minor**.

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

[Unreleased]: https://github.com/DaBorsten/pptx-templates/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/DaBorsten/pptx-templates/releases/tag/v0.1.0
