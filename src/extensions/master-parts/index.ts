import path from 'path';
import IArchive from '../../interfaces/iarchive';

/** Parts whose relationships the search starts from. */
const MASTER_DIRS = ['ppt/slideMasters', 'ppt/slideLayouts'];

/**
 * Folders `ModifyPresentationHelper.removeUnusedFiles` empties of the files
 * it did not copy with a slide, and whose parts masters and layouts use.
 */
const KEPT_DIRS = ['ppt/charts', 'ppt/embeddings'];

/**
 * Files in `ppt/charts` and `ppt/embeddings` that a slide master or layout
 * of the archive uses, by folder: OLE objects and embedded documents (e.g.
 * think-cell, a logo object on the master), charts, and what these point
 * at in turn (a chart's style, colors and workbook).
 *
 * `cleanup` keeps masters and layouts as they are, but removed every chart
 * and embedding that was not copied with a slide: the relationships of the
 * masters and layouts then pointed at missing parts, and PowerPoint offered
 * to repair the file. `removeUnusedFiles` keeps these files.
 */
export async function partsOfMastersAndLayouts(
  archive: IArchive,
): Promise<Record<string, string[]>> {
  const queue: string[] = [];
  for (const dir of MASTER_DIRS) {
    for (const file of await archive.folder(`${dir}/_rels`)) {
      if (file.name.endsWith('.rels')) queue.push(file.name);
    }
  }

  const keep: Record<string, string[]> = {};
  const seen = new Set<string>();
  while (queue.length) {
    for (const target of await internalTargets(archive, queue.shift())) {
      const dir = path.posix.dirname(target);
      if (!KEPT_DIRS.includes(dir) || seen.has(target)) continue;
      seen.add(target);
      keep[dir] = [...(keep[dir] || []), path.posix.basename(target)];
      const rels = relsOf(target);
      if (archive.fileExists(rels)) queue.push(rels);
    }
  }
  return keep;
}

/** Archive paths of the parts a .rels file points at, external ones left out. */
async function internalTargets(
  archive: IArchive,
  relsFile: string,
): Promise<string[]> {
  // Targets are relative to the part's folder, not to its _rels folder.
  const partDir = path.posix.dirname(path.posix.dirname(relsFile));
  const relationships = (await archive.readXml(relsFile)).getElementsByTagName(
    'Relationship',
  );
  const targets: string[] = [];
  for (let i = 0; i < relationships.length; i++) {
    const rel = relationships.item(i);
    if (rel.getAttribute('TargetMode') === 'External') continue;
    const target = rel.getAttribute('Target') || '';
    targets.push(
      target.startsWith('/')
        ? path.posix.normalize(target.slice(1))
        : path.posix.normalize(path.posix.join(partDir, target)),
    );
  }
  return targets;
}

/** `ppt/charts/chart1.xml` → `ppt/charts/_rels/chart1.xml.rels`. */
function relsOf(part: string): string {
  return `${path.posix.dirname(part)}/_rels/${path.posix.basename(part)}.rels`;
}
