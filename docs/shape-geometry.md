---
title: Shape Geometry
description: Crop pictures and shapes to any preset shape, round corners in cm, px, pt or percent, and add borders.
---

Every shape and picture in PowerPoint has a geometry: a rectangle, an
ellipse, a heart, one of about 190 presets. With a picture, the geometry is
the mask the picture is cropped to (PowerPoint: *Crop to Shape*). Two
`modify.*` helpers change it. They work on `<p:sp>` (shapes) and `<p:pic>`
(pictures).

```ts
import { modify } from 'pptx-templates';

pres.addSlide('SlideWithImages.pptx', 2, (slide) => {
  // Crop a picture to an ellipse
  slide.modifyElement('imagePNG', modify.setGeometry('ellipse'));

  // Any preset, with the raw adjust values (the yellow handles)
  slide.modifyElement('imageJPG', modify.setGeometry('star5', { adj: 30000 }));

  // Rounded corners: in percent, px, pt, cm or mm
  slide.modifyElement('imagePNGduotone', [
    modify.setCornerRadius('12px'),
    // Border: setOutline, width in EMU (1pt = 12700)
    modify.setOutline({
      weight: 25400,
      color: { type: 'srgbClr', value: '1F4E79' },
    }),
  ]);
});
```

## `modify.setGeometry(shape, adjustments?)`

Replaces the preset or custom geometry of the element with `shape`.
`shape` is any name of the DrawingML spec (`ShapePresets` holds the full
list), e.g. `'rect'`, `'roundRect'`, `'ellipse'`, `'triangle'`, `'hexagon'`,
`'star5'`, `'heart'`, `'cloud'`, `'flowChartDocument'`.

`adjustments` sets the adjust values of the shape, keyed by guide name, e.g.
`{ adj: 25000 }`. Names and value ranges depend on the shape. Most are in
1/100000 of the shorter side, so `25000` is 25 %. Without `adjustments`,
the shape uses PowerPoint's defaults. Guides you leave out get their
default value, because PowerPoint asks to repair a file that lists only some
of them (`star5` with only `adj` is written as `adj`, `hf`, `vf`). A guide
name the shape does not have is an error.

## `modify.setCornerRadius(radius, shape?)`

Rounds the corners. The radius is measured on the shorter side of the
element and capped at half of it.

| `radius` | Meaning |
|---|---|
| `0.5` (number) | Centimeters, as in docx-templates |
| `'10%'` | Percent of the shorter side. `'50%'` makes the short sides fully round |
| `'12px'` | CSS pixels at 96 dpi (1px = 9525 EMU) |
| `'4pt'`, `'0.5cm'`, `'5mm'` | Points, centimeters, millimeters |

PPTX stores the radius as a fraction of the shorter side, not as a
length. `px`, `pt`, `cm` and `mm` are converted with the size the element
has when the callback runs. If you resize the element afterwards, the
corners scale with it. Absolute units need an element with its own size; a
placeholder that inherits its size from the layout only takes a percentage.

`shape` picks the rounded preset:

| `shape` | Corners |
|---|---|
| `'roundRect'` (default) | all four |
| `'round1Rect'` | top right |
| `'round2SameRect'` | top two |
| `'round2DiagRect'` | top left and bottom right |
| `'snipRoundRect'` | top left round, top right cut |

Without `shape`, an element that already has one of these presets keeps it
and only its radius changes; any other element becomes a `roundRect`.

## Borders

Borders are the outline of the shape, set with `modify.setOutline` (see
[Helpers](./helpers.md)). On a cropped picture, the border follows the
geometry, so an ellipse gets a round border.

## In templates

The `{{IMAGE}}` tag of [Template Logic](./template-logic.md#images) takes the
same options as docx-templates: `shape`, `cornerRadius`, `shapeAdjustments`
and `border`.

```text
{{IMAGE ({ data: logo, extension: '.png', width: 4, cornerRadius: '10%', border: { color: '#1f4e79', width: 2 } })}}
```
