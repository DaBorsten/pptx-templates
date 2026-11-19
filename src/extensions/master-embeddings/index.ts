import path from 'path';
import IArchive from '../../interfaces/iarchive';

const RELS_DIRS = ['ppt/slideMasters/_rels', 'ppt/slideLayouts/_rels'];
const EMBEDDINGS_DIR = 'ppt/embeddings';

/**
 * File names in `ppt/embeddings` that a slide master or layout of the
 * archive points at (OLE objects, embedded documents, e.g. think-cell or a
 * logo object on the master).
 *
 * `cleanup` keeps masters and layouts as they are, but removed every
 * embedding that was not copied with a slide: the relationships of the
 * masters and layouts then pointed at missing parts, and PowerPoint offered
 * to repair the file. `ModifyPresentationHelper.removeUnusedFiles` keeps
 * these files.
 */
export async function embeddingsOfMastersAndLayouts(
  archive: IArchive,
): Promise<string[]> {
  const keep: string[] = [];
  for (const dir of RELS_DIRS) {
    for (const file of await archive.folder(dir)) {
      if (!file.name.endsWith('.rels')) continue;
      const rels = await archive.readXml(file.name);
      const relationships = rels.getElementsByTagName('Relationship');
      for (let i = 0; i < relationships.length; i++) {
        const rel = relationships.item(i);
        if (rel.getAttribute('TargetMode') === 'External') continue;
        const target = resolveTarget(dir, rel.getAttribute('Target') || '');
        if (path.posix.dirname(target) === EMBEDDINGS_DIR) {
          keep.push(path.posix.basename(target));
        }
      }
    }
  }
  return keep;
}

/** `../embeddings/x.bin` from `ppt/slideLayouts/_rels` → `ppt/embeddings/x.bin`. */
function resolveTarget(relsDir: string, target: string): string {
  if (target.startsWith('/')) return path.posix.normalize(target.slice(1));
  // Targets are relative to the part's folder, not to its _rels folder.
  return path.posix.normalize(
    path.posix.join(path.posix.dirname(relsDir), target),
  );
}
