import JSZip from 'jszip';
import { lintTemplate, LintTemplateOptions } from '../../src/index';

/**
 * Minimal in-memory .pptx files for the template linter: just the parts it
 * reads (slide XML, rels, layouts, notes, …), no full package.
 */

const escapeXml = (text: string) =>
  text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Every inner array is one paragraph, every string one run. */
export const paragraphsXml = (paragraphs: string[][]) =>
  paragraphs
    .map(
      (runs) =>
        '<a:p>' +
        runs.map((t) => `<a:r><a:rPr lang="de-DE"/><a:t>${escapeXml(t)}</a:t></a:r>`).join('') +
        '</a:p>',
    )
    .join('');

/** A shape; `ownFrame: false` is a placeholder that inherits its position. */
export const shape = (name: string, paragraphs: string[][], ownFrame = true) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr>${ownFrame ? '<a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' : ''}</p:spPr>` +
  `<p:txBody><a:bodyPr/>${paragraphsXml(paragraphs)}</p:txBody></p:sp>`;

/** A table; `rows[r][c]` is the text of one cell. */
export const table = (name: string, rows: string[][]) =>
  `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="${name}"/></p:nvGraphicFramePr>` +
  '<a:graphic><a:graphicData><a:tbl>' +
  rows
    .map(
      (cells) =>
        '<a:tr>' +
        cells
          .map((text) => `<a:tc><a:txBody><a:bodyPr/>${paragraphsXml([[text]])}</a:txBody></a:tc>`)
          .join('') +
        '</a:tr>',
    )
    .join('') +
  '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';

/** A slide/layout/master/notes part with the given shapes. */
export const partXml = (root: string, content: string) =>
  `<?xml version="1.0" encoding="UTF-8"?><p:${root} ` +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  `<p:cSld><p:spTree>${content}</p:spTree></p:cSld></p:${root}>`;

/** A chart part whose title is `text`. */
export const chartXml = (text: string) =>
  '<?xml version="1.0" encoding="UTF-8"?><c:chartSpace ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
  `<c:chart><c:title><c:tx><c:rich>${paragraphsXml([[text]])}</c:rich></c:tx></c:title></c:chart>` +
  '</c:chartSpace>';

/** A SmartArt data part with one text point. */
export const smartArtXml = (text: string) =>
  '<?xml version="1.0" encoding="UTF-8"?><dgm:dataModel ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
  `<dgm:ptLst><dgm:pt modelId="1"><dgm:t><a:bodyPr/>${paragraphsXml([[text]])}</dgm:t></dgm:pt></dgm:ptLst>` +
  '</dgm:dataModel>';

/** Relationships of slide 1 to the given targets (relative to ppt/slides/). */
export const slideRels = (...targets: string[]) =>
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  targets.map((t, i) => `<Relationship Id="rId${i + 1}" Type="x" Target="${t}"/>`).join('') +
  '</Relationships>';

/** A minimal .pptx for the linter: slide XML plus optional other parts. */
export const deck = async (slides: string[], parts: Record<string, string> = {}) => {
  const zip = new JSZip();
  slides.forEach((content, i) =>
    zip.file(`ppt/slides/slide${i + 1}.xml`, partXml('sld', content)),
  );
  Object.entries(parts).forEach(([name, xml]) => zip.file(name, xml));
  return zip.generateAsync({ type: 'nodebuffer' });
};

export const lint = async (
  slides: string[],
  options?: LintTemplateOptions,
  parts?: Record<string, string>,
) => (await lintTemplate(await deck(slides, parts), options)).issues;
