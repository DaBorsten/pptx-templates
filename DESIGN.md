---
name: pptx-templates docs
description: Template-first .pptx generation for Node.js, documented as claim then proof.
colors:
  accent: "#c2410c"
  accent-dark: "#fb923c"
  accent-soft: "rgba(194, 65, 12, 0.14)"
  accent-line: "rgba(194, 65, 12, 0.5)"
  syntax-string: "#B23A0A"
  syntax-string-dark: "#FDBA74"
  panel: "#ffffff"
  paper: "#fafafa"
  surface-2: "#f2f2f2"
  hairline: "#e6e6e6"
  hairline-strong: "#cfcfcf"
  marker: "#a6a6a6"
  ink-2: "#6b6b6b"
  ink-3: "#3d3d3d"
  ink-body: "#262626"
  ink: "#111111"
  panel-dark: "#111111"
  paper-dark: "#0a0a0a"
  surface-2-dark: "#171717"
  hairline-dark: "#262626"
  hairline-strong-dark: "#333333"
  marker-dark: "#525252"
  ink-2-dark: "#a3a3a3"
  ink-body-dark: "#d9d9d9"
  ink-dark: "#f5f5f5"
typography:
  display:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, -apple-system, Segoe UI, sans-serif"
    fontSize: "clamp(2rem, 1.5rem + 2vw, 2.75rem)"
    fontWeight: 650
    lineHeight: 1.08
    letterSpacing: "-0.035em"
    fontFeature: "'ss01' on"
  headline:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, sans-serif"
    fontSize: "1.5rem"
    fontWeight: 620
    lineHeight: 1.25
    letterSpacing: "-0.025em"
  title:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, sans-serif"
    fontSize: "1.1875rem"
    fontWeight: 620
    lineHeight: 1.25
    letterSpacing: "-0.015em"
  lead:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, sans-serif"
    fontSize: "1.1875rem"
    fontWeight: 400
    lineHeight: 1.6
    letterSpacing: "-0.01em"
  body:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, sans-serif"
    fontSize: "1rem"
    fontWeight: 400
    lineHeight: 1.7
  body-small:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, sans-serif"
    fontSize: "0.875rem"
    fontWeight: 450
    lineHeight: 1.4
  label:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, sans-serif"
    fontSize: "0.8125rem"
    fontWeight: 600
    lineHeight: 1.45
  group-label:
    fontFamily: "Geist Variable, ui-sans-serif, system-ui, sans-serif"
    fontSize: "0.75rem"
    fontWeight: 600
    letterSpacing: "0.01em"
  code:
    fontFamily: "Geist Mono Variable, ui-monospace, Cascadia Code, SF Mono, Consolas, monospace"
    fontSize: "0.8125rem"
    fontWeight: 400
    lineHeight: 1.65
  code-inline:
    fontFamily: "Geist Mono Variable, ui-monospace, Cascadia Code, SF Mono, Consolas, monospace"
    fontSize: "0.875em"
    fontWeight: 450
  code-badge:
    fontFamily: "Geist Mono Variable, ui-monospace, Cascadia Code, SF Mono, Consolas, monospace"
    fontSize: "0.6875rem"
    fontWeight: 500
    lineHeight: 1
rounded:
  xs: "4px"
  sm: "6px"
  md: "8px"
  lg: "10px"
spacing:
  xs: "0.375rem"
  sm: "0.75rem"
  md: "1.125rem"
  lg: "2rem"
  xl: "3.25rem"
components:
  code-panel:
    backgroundColor: "{colors.panel}"
    textColor: "{colors.ink}"
    typography: "{typography.code}"
    rounded: "{rounded.lg}"
    padding: "1rem 1.125rem"
  code-panel-title:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink-2}"
    padding: "0.5rem 1.125rem"
  code-inline:
    backgroundColor: "{colors.surface-2}"
    textColor: "{colors.ink}"
    typography: "{typography.code-inline}"
    rounded: "{rounded.sm}"
    padding: "0.1em 0.35em"
  link-prose:
    textColor: "{colors.ink}"
  link-prose-hover:
    textColor: "{colors.accent}"
  menu-link:
    textColor: "{colors.ink-2}"
    typography: "{typography.body-small}"
    rounded: "{rounded.md}"
    padding: "0.375rem 0.75rem"
  menu-link-active:
    backgroundColor: "{colors.surface-2}"
    textColor: "{colors.ink}"
    rounded: "{rounded.md}"
  sidebar-group-label:
    textColor: "{colors.ink}"
    typography: "{typography.group-label}"
    padding: "0.25rem 0.75rem"
  navbar-version:
    backgroundColor: "{colors.panel}"
    textColor: "{colors.ink-2}"
    typography: "{typography.code-badge}"
    rounded: "{rounded.sm}"
    padding: "0.3125rem 0.5rem"
  navbar-version-hover:
    textColor: "{colors.ink}"
  navbar-icon-link:
    textColor: "{colors.ink-2}"
    rounded: "{rounded.md}"
    padding: "0.5rem"
    size: "18px"
  navbar-icon-link-hover:
    backgroundColor: "{colors.surface-2}"
    textColor: "{colors.ink}"
  search-input:
    backgroundColor: "{colors.panel}"
    textColor: "{colors.ink}"
    rounded: "{rounded.md}"
    height: "2.25rem"
    width: "16rem"
  admonition:
    backgroundColor: "{colors.panel}"
    textColor: "{colors.ink-body}"
    rounded: "{rounded.lg}"
    padding: "1rem 1.125rem"
  pagination-link:
    backgroundColor: "{colors.panel}"
    textColor: "{colors.ink}"
    rounded: "{rounded.lg}"
    padding: "0.875rem 1.125rem"
  card:
    backgroundColor: "{colors.panel}"
    textColor: "{colors.ink}"
    rounded: "{rounded.lg}"
---

# Design System: pptx-templates docs

## Overview

**Creative North Star: "Claim, Then Proof"**

The docs are a reading surface for developers and their AI assistants: a short claim in prose, then the TypeScript that proves it, directly beneath. The whole page is one neutral scale so the code examples are the only thing that carries colour. The world is a monochrome product canon: Geist set tight in the headings, hairline structure, generous reading measure, and a single deep orange that marks links, focus and strings in code and nothing else.

Density is calm, not sparse: a fixed sidebar grouped into labelled sections in secondary ink, a centred reading column, an "On this page" rail. Depth is used once, deliberately: the code panel is the only surface lifted off the paper at rest. Everything else is drawn with 1px lines and tonal steps. Motion is a single moment, the article settling in on navigation.

The system rejects stock Infima: tinted pills for links, thick coloured admonition bars, a dark footer slab, and high-saturation syntax themes. There is no logo: the brand is the "pptx-templates" wordmark alone (user decision).

**Key Characteristics:**
- One neutral ramp (nine steps, light and dark), one accent, used sparingly.
- Geist Variable for text, Geist Mono Variable for code, both self-hosted.
- Hairline borders, 10px radii on every framed block.
- Code panels framed and lifted; prose, cards and admonitions flat.
- Active states shown with ink and weight, never with colour.
- No logo; the navbar carries the wordmark only (user decision).

## Colors

A neutral grey ramp carries the entire interface; a deep burnt orange is the only hue, and it is mostly drawn as lines.

### Primary
- **Burnt Orange** (`accent`): link underlines at half strength (`accent-line`), link and hash-link hover text, the focus ring, the text caret, selection tint (`accent-soft`), warning and danger admonition icons. In dark mode it lifts to **Ember Orange** (`accent-dark`).
- **Syntax String Orange** (`syntax-string`, `syntax-string-dark`): strings, numbers and booleans in code. Deliberately one step deeper than the accent on white and one step lighter on black, for legibility at 13px mono.

### Neutral
- **Panel White** (`panel`): code panels, admonitions, cards, pagination links, search input. Dark: `panel-dark`, a lifted near-black.
- **Paper** (`paper`): page, navbar (82% translucent with blur), sidebar, footer, code panel title strip. Dark: `paper-dark`.
- **Surface 2** (`surface-2`): inline code fill, active and hovered menu item, kbd keys, highlighted code line, search focus halo.
- **Hairline** (`hairline`): every structural line: navbar and sidebar edges, h2 rules, table rows, panel borders, TOC rail.
- **Hairline Strong** (`hairline-strong`): the hover step for bordered elements, table header underline, blockquote rule, scrollbar thumb.
- **Marker Grey** (`marker`): list markers and resting hash links.
- **Secondary Ink** (`ink-2`): sidebar and navbar links, breadcrumbs, table headers, TOC links, meta labels, code comments and punctuation.
- **Body Ink** (`ink-body`): running prose.
- **Ink** (`ink`): headings, the lead paragraph, strong text, link text, all active states, code keywords and function names (weight-differentiated).

### Named Rules
**The One Accent Rule.** The accent marks; it never fills. It appears on link underlines, hover text, focus, caret, warning/danger icons and code strings. No accent buttons, no accent backgrounds; the only tinted fill is text selection.

**The Ink Link Rule.** Links in prose are ink text, weight 500, with a 1px half-strength accent underline offset 0.22em. Hover turns text and underline full accent. Chrome links (sidebar, navbar, TOC, footer) go from secondary ink to ink and carry no accent.

**The Two-Tone Syntax Rule.** Code is ink plus one orange. Keywords and function names separate by weight (600 / 500), not by hue; comments step back to grey; only diff deleted/inserted lines use red and green.

## Typography

**Display Font:** Geist Variable (with ui-sans-serif, system-ui, Segoe UI)
**Body Font:** Geist Variable
**Label/Mono Font:** Geist Mono Variable (with ui-monospace, Cascadia Code, SF Mono, Consolas)

**Character:** One grotesque in fine variable weights (450, 560, 620, 650) does all the hierarchy work; tight negative tracking on headings, stylistic set ss01 on globally. Mono is reserved for code, code titles, kbd keys and the version badge.

### Hierarchy
- **Display** (650, fluid up to 2.75rem, 1.08, -0.035em, balanced): page title only.
- **Lead** (400, 1.1875rem, 1.6, -0.01em, ink): the first paragraph after the title; the page's claim, one step above body.
- **Headline** (620, 1.5rem, 1.25, -0.025em): h2, always preceded by a hairline rule and 2rem of air.
- **Title** (620, 1.1875rem, -0.015em): h3.
- **Body** (400, 1rem, 1.7, pretty wrapping): prose and lists, in body ink.
- **Body Small** (450, 0.875rem): sidebar, navbar links, tables, footer items.
- **Group Label** (600, 0.75rem, +0.01em, ink, sentence case): sidebar section headings; not links.
- **Label** (600, 0.8125rem): TOC heading, admonition headings (sentence case), footer titles, table headers (560, secondary ink).
- **Code** (Geist Mono, 0.8125rem, 1.65, ligatures off): code panels. Inline code at 0.875em, weight 450, with a hairline border.

### Named Rules
**The Weight Not Colour Rule.** Hierarchy and active state are expressed by ink and weight steps, not by hue or size jumps: active sidebar item 560, active TOC entry 500, strong 600.

**The No Uppercase Rule.** No uppercase labels anywhere; sidebar group labels, admonition headings and table headers stay in sentence case.

## Layout

Three columns on desktop (from 997px): sticky sidebar with a hairline right edge (grouped into fixed sections 1.25rem apart), a centred reading column capped at 44rem, and the right-hand "On this page" rail. The article starts 2.5rem below the navbar (3.75rem tall); breadcrumbs sit 1.25rem above the title. Vertical rhythm is set by headings: h2 gets 3.25rem above plus a rule with 2rem padding, h3 2.25rem, h4 1.75rem. Framed blocks keep a 1.125rem inline inset; menu items use 0.375rem by 0.75rem.

At 996px and below the sidebar moves into the navbar drawer, the search field collapses to a 2.25rem square that expands to 12rem on focus, the TOC becomes a bordered panel at the top of the article, and h2 spacing tightens to 2.5rem / 1.5rem. Tables keep their first column flush left so they align with the prose edge.

## Elevation & Depth

Hybrid, almost flat. Depth is tonal: paper below, panel white above, separated by hairlines. One shadow exists, a soft two-layer lift (see sidecar), and at rest it belongs only to code panels; pagination links borrow it on hover. In dark mode the shadow nearly disappears and the panel's lighter near-black carries the lift. The navbar floats via translucency and a 12px backdrop blur, not a shadow.

### Named Rules
**The Lifted Proof Rule.** Only the proof is raised. Code panels carry the panel shadow at rest; admonitions, cards, tables and the TOC stay flat. Interactive panels may gain the shadow on hover.

## Shapes

Gently rounded rectangles in four steps: 4px for focus rings and kbd keys, 6px for inline code and the copy button, 8px for menu items, search and the colour-mode toggle, 10px for every framed block (code panels, admonitions, cards, pagination, mobile TOC). Every frame is a 1px hairline; there are no thick borders and no coloured side bars. The only side rules are 1px: the blockquote's hairline-strong left edge, the guide line under a nested sidebar category (never inside a top-level group), and the TOC rail where the active entry turns ink.

## Components

### Code Panel (signature)
The proof. Panel white, 1px hairline, 10px radius, panel shadow, clipped. Code at 0.8125rem Geist Mono with 1rem by 1.125rem inset and the two-tone syntax theme. Optional title strip in mono 0.75rem secondary ink on paper, separated by a hairline. The copy button is a 6px hairline square on panel white, secondary ink, going to ink and hairline-strong on hover.

### Inline Code
Surface 2 fill, 1px hairline, 6px radius, ink, weight 450, breaks long words. Inside a link it inherits the link colour.

### Links
Prose links follow the Ink Link Rule with 160ms colour transitions. Heading hash links are marker grey and turn accent on hover.

### Navigation
- **Navbar:** paper at 82% (dark 80%) with saturate-and-blur backdrop and a hairline bottom. Brand: no logo, wordmark only (user decision): "pptx-templates" at 0.9375rem / 650, -0.02em, in ink, followed by the version badge. Text links 0.875rem / 500 in secondary ink, ink when hovered or active.
- **Version badge:** the package version (e.g. v0.3.2) in Geist Mono 0.6875rem / 500, secondary ink on panel white, 1px hairline, 6px radius; hover inks the text and steps the border to hairline-strong. Links to the CHANGELOG. Besides the search field it is the only framed object in the navbar.
- **Icon links:** npm and GitHub as 18px Simple Icons SVG marks (fill currentColor, labelled by aria-label), icon-only on desktop. 0.5rem padding, 8px radius, surface-2 fill on hover. Brand marks of external services only; no other icons are added to the chrome.
- **Drawer state:** while the mobile drawer is open the navbar drops its backdrop blur and turns solid paper (see the Contained Drawer Rule).
- **Search:** 2.25rem tall, 16rem wide on desktop, panel white, hairline border, 8px radius; hover steps to hairline-strong, focus to secondary ink with a 3px surface-2 halo (no accent). Kbd hints in mono 0.6875rem on surface 2.
- **Sidebar:** grouped into non-collapsible sections in a fixed order: Get started, Modify content, Templates, Build decks, Help, Reference (unknown upstream pages land in "More", so none is hidden). Each group opens with a group label (0.75rem / 600 ink, sentence case, not clickable); groups sit 1.25rem apart. Items are 0.875rem, weight 450, secondary ink, flush under their label with no indent and no guide line. Hover and active use the surface-2 pill (8px); active text is ink at 560. Only a linked category inside a group (the API reference) nests, hanging from a 1px hairline guide.
- **TOC:** 0.8125rem, an "On this page" label in ink, links on a 1px hairline rail; the active entry turns ink, weight 500, and inks its segment of the rail.
- **Breadcrumbs:** 0.8125rem secondary ink, no pills; current page ink at 500. The trail follows the sidebar groups: group, then page (e.g. "Templates > Template Logic").

**The Contained Drawer Rule.** Never leave `backdrop-filter` on an element that contains fixed-position UI while that UI is open: the blur turns the navbar into the containing block of the fixed mobile drawer. With the drawer open the navbar drops the blur and uses solid paper.

**The Labels Are Not Links Rule.** Sidebar group labels only organise; they never navigate, collapse or take a hover state.

### Tables
Open tables: no cell borders, no stripes, no header fill. 0.875rem at 1.55, a hairline under each row and hairline-strong under the header. Headers are 0.8125rem / 560 secondary ink, no wrap. First column flush left; code in it does not wrap; tabular figures throughout.

### Admonitions
Flat panel-white box, 1px hairline all round, 10px radius, 1rem by 1.125rem inset. Heading 0.8125rem / 600 ink in sentence case. Icons secondary ink; warning and danger icons take the accent. Types are told apart by heading and icon, not by colour fields.

### Cards / Pagination
Category and API index cards: flat panel white, hairline, 10px; hover strengthens the border to marker grey. Pagination links: same frame with 0.875rem by 1.125rem inset, a 0.75rem secondary-ink sublabel and a 0.9375rem / 600 ink label; hover gains hairline-strong and the panel shadow.

### Footer
Paper, not a dark slab: hairline top, secondary-ink links that turn ink, 0.8125rem / 600 titles.

### Motion
One moment: the article settles in on navigation (420ms, opacity from 0.35 and 6px rise, expo-out curve), only when the reader allows motion. All state changes are 160ms colour/border transitions on the same curve.

## Do's and Don'ts

### Do:
- **Do** keep all docs styling in the fork's own files (`website/src/css/fork-theme.css`, `website/src/fork/` for the prism themes, sidebar grouping and navbar items); upstream files such as `website/src/css/custom.css` stay untouched, and any single added line in an upstream file is recorded in `FORK.md`.
- **Do** put a code panel directly beneath the claim it proves.
- **Do** use the neutral ramp tokens for every new surface and line; reach for `hairline` first and `hairline-strong` only for hover or emphasis.
- **Do** frame new blocks with a 1px hairline and the 10px radius.
- **Do** show active and selected states with ink and a weight step, and a surface-2 pill where the item is a navigation target.
- **Do** keep the accent's focus ring (2px, 2px offset) on every interactive element.
- **Do** define both light and dark values for every new token.
- **Do** file every new docs page under one of the sidebar groups in `website/src/fork/sidebars.ts`; "More" is a safety net for upstream additions, not a destination.

### Don't:
- **Don't** move the design into upstream-owned files under `website/`; the fork boundary is the rule (see `FORK.md`).
- **Don't** fill components with the accent: no orange buttons, badges, pills or panel backgrounds.
- **Don't** use thick or coloured side bars on admonitions or quotes; frames and rules stay 1px.
- **Don't** put shadows on flat blocks (admonitions, cards at rest, tables); only code panels are lifted at rest.
- **Don't** add hues to the syntax theme beyond the single string orange and the diff red/green.
- **Don't** use uppercase labels, tinted link pills, striped tables or a dark footer slab.
- **Don't** keep `backdrop-filter` on the navbar while the mobile drawer is open.
- **Don't** add icons to the navbar beyond the external brand marks (npm, GitHub); internal navigation stays text.
- **Don't** add decoration the Markdown twin would lose meaning without; every page must read the same for an AI assistant.
