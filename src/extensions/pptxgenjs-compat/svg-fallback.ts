import * as path from 'path';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { AutomizerFile } from '../../types/types';
import { PptPaths } from '../../helper/ppt-paths';
import { IMG_BROKEN_PNG } from './img-broken-png';

const PNG_SIGNATURE = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
]);
const SLIDE_PART = /^ppt\/slides\/slide\d+\.xml$/;

/**
 * pptxgenjs 4 writes the SVG bytes into the PNG fallback of an SVG image when
 * it runs in Node. In 3.x the encoding step ran synchronously, so a
 * placeholder PNG was put there instead. In 4.x it runs after
 * `await import('node:fs')`, and the placeholder step no longer sees the
 * data. The result is a `.png` part with SVG content. Our MediaDeduplicator
 * then merges it with the byte-identical `.svg` part, so the `svgBlip` ends up
 * pointing to a part declared as `image/png`.
 *
 * This puts the 3.x placeholder back into every fallback that is not a real
 * PNG. `file` is returned unchanged when nothing needs repairing.
 */
export async function repairSvgFallbacks(
  file: AutomizerFile,
): Promise<AutomizerFile> {
  if (typeof file === 'string') {
    return file;
  }

  const zip = await JSZip.loadAsync(file);
  const fallbackParts = new Set<string>();

  const slideParts = Object.keys(zip.files).filter((name) =>
    SLIDE_PART.test(name),
  );
  for (const slidePart of slideParts) {
    for (const target of await svgFallbackTargets(zip, slidePart)) {
      fallbackParts.add(target);
    }
  }

  let repaired = false;
  for (const part of fallbackParts) {
    const entry = zip.file(part);
    if (!entry) continue;
    const content = await entry.async('nodebuffer');
    if (!content.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
      zip.file(part, Buffer.from(IMG_BROKEN_PNG, 'base64'));
      repaired = true;
    }
  }

  if (!repaired) {
    return file;
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'STORE' });
}

/**
 * Package paths of the PNG fallbacks of all SVG images on one slide: the
 * `r:embed` of every `a:blip` that carries an `asvg:svgBlip` extension.
 */
async function svgFallbackTargets(
  zip: JSZip,
  slidePart: string,
): Promise<string[]> {
  const slideXml = await zip.file(slidePart)!.async('string');
  if (!slideXml.includes('svgBlip')) {
    return [];
  }

  const relsEntry = zip.file(PptPaths.relsOf(slidePart));
  if (!relsEntry) {
    return [];
  }

  const parser = new DOMParser();
  const relsDoc = parser.parseFromString(
    await relsEntry.async('string'),
    'application/xml',
  );
  const targetsById = new Map<string, string>();
  const rels = relsDoc.getElementsByTagName('Relationship');
  for (let i = 0; i < rels.length; i++) {
    targetsById.set(
      rels[i].getAttribute('Id') ?? '',
      rels[i].getAttribute('Target') ?? '',
    );
  }

  const slideDoc = parser.parseFromString(slideXml, 'application/xml');
  const blips = slideDoc.getElementsByTagName('a:blip');
  const targets: string[] = [];
  for (let i = 0; i < blips.length; i++) {
    const blip = blips[i];
    if (blip.getElementsByTagName('asvg:svgBlip').length === 0) continue;
    const target = targetsById.get(blip.getAttribute('r:embed') ?? '');
    if (!target) continue;
    targets.push(
      target.startsWith('/')
        ? target.substring(1)
        : path.posix.normalize(
            path.posix.join(path.posix.dirname(slidePart), target),
          ),
    );
  }
  return targets;
}
