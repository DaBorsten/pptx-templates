# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

Node.js developers who generate PowerPoint decks from existing, designer-made
.pptx templates (reports, data-driven slides). They read the docs while
coding: looking up `modify.*` helpers, selectors, and the deferred
`write()` model. A large share of the reading happens through their AI
assistants (llms.txt, the Markdown twin of every page, `AI-INSTRUCTOR.md`), so
the human pages and the machine-readable twins must both stay clean.

## Product Purpose

`pptx-templates` modifies existing .pptx files instead of building slides from
scratch: load templates, copy slides and shapes between them, change text,
tables, charts and images through callbacks, and write a new deck. Success for
the docs: a developer finds the right helper and a working example quickly,
and an AI assistant can consume the same content without noise.

## Positioning

Template-first: the slide design lives in PowerPoint, made by whoever owns the
corporate template; code only fills and rearranges it. PptxGenJS is wrapped
for the rare from-scratch shape.

## Operating Context

- Docs site: Docusaurus 3 in `website/`, content in `docs/` at the repo root,
  typedoc-generated API reference in `docs/api/`, local search, deployed to
  GitHub Pages at `/pptx-templates/`.
- The repo is a fork of `singerla/pptx-automizer`; `website/` exists upstream
  too, so fork changes there are additive (new files, single added lines) and
  recorded in `FORK.md`.
- Every page has a Markdown twin and is listed in llms.txt (custom plugin).

## Capabilities and Constraints

- Docs-only site, served at the root URL; no landing page, no URL changes.
- Code examples are TypeScript throughout and typechecked by the test suite.
- No logo or brand assets exist yet.

## Brand Commitments

None binding. The current orange accent was a placeholder nod to PowerPoint.

## Evidence on Hand

README, `docs/*.md` guides, generated API reference, CHANGELOG. No
testimonials, user counts or benchmarks exist; none may be invented.

## Product Principles

1. The reference is the product: finding a helper and a working example beats any decoration.
2. One source for humans and machines: nothing on a page that the Markdown twin would lose meaning without.
3. Additive to upstream: fork changes live in new files.
