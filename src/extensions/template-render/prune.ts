import { posix } from 'path';
import JSZip from 'jszip';
import { readText } from './zip-text';

/**
 * The cleanup of renderDeck: removes every part of the package that no
 * relationship reaches any more, starting from `_rels/.rels`. These are the
 * template's own slides (removeExistingSlides only drops them from the slide
 * list) with their notes, charts, embeddings and pictures.
 *
 * It replaces Automizer's `cleanup` option there, which parses every slide
 * once more to find the pictures it uses; this reads the `.rels` parts as
 * text only. Like `cleanup`, a picture counts as used only where its
 * relationship id appears in the part (a replaced picture leaves its
 * relationship behind).
 */
export const removeUnreachableParts = async (zip: JSZip): Promise<void> => {
  const files = new Set(
    Object.keys(zip.files).filter((name) => !zip.files[name].dir),
  );
  const reachable = new Set<string>(['[Content_Types].xml']);
  const queue: string[] = [''];

  while (queue.length) {
    const part = queue.shift();
    const relsPath = relsOf(part);
    if (!files.has(relsPath)) continue;
    reachable.add(relsPath);
    const rels = await readText(zip, relsPath);
    const relations = parseRelationships(rels);
    // the part's own text, read only if it relates to pictures
    let content: string | undefined;
    for (const relation of relations) {
      if (relation.external) continue;
      const target = resolveTarget(part, relation.target);
      if (!target || reachable.has(target) || !files.has(target)) continue;
      if (relation.image && part) {
        content ??= await readText(zip, part);
        const id = relation.id;
        if (!content.includes(`"${id}"`) && !content.includes(`'${id}'`)) {
          continue;
        }
      }
      reachable.add(target);
      queue.push(target);
    }
  }

  [...files]
    .filter((file) => !reachable.has(file))
    .forEach((file) => {
      zip.remove(file);
      files.delete(file);
    });
  await removeContentTypeOverrides(zip, files);
};

type Relationship = {
  id: string;
  target: string;
  external: boolean;
  image: boolean;
};

const attribute = (tag: string, name: string): string | undefined => {
  const match = new RegExp(`\\s${name}\\s*=\\s*("([^"]*)"|'([^']*)')`).exec(
    tag,
  );
  return match ? decodeEntities(match[2] ?? match[3]) : undefined;
};

const decodeEntities = (value: string): string =>
  value.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-f]+|#\d+);/gi, (m, e) => {
    switch (e.toLowerCase()) {
      case 'amp':
        return '&';
      case 'lt':
        return '<';
      case 'gt':
        return '>';
      case 'quot':
        return '"';
      case 'apos':
        return "'";
      default:
        return String.fromCodePoint(
          e[1].toLowerCase() === 'x'
            ? parseInt(e.slice(2), 16)
            : parseInt(e.slice(1), 10),
        );
    }
  });

const parseRelationships = (xml: string): Relationship[] =>
  [...xml.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)].map(([tag]) => ({
    id: attribute(tag, 'Id') ?? '',
    target: attribute(tag, 'Target') ?? '',
    external: attribute(tag, 'TargetMode') === 'External',
    image: /\/image$/.test(attribute(tag, 'Type') ?? ''),
  }));

/** `ppt/slides/slide1.xml` → `ppt/slides/_rels/slide1.xml.rels`. */
const relsOf = (part: string): string =>
  part
    ? posix.join(posix.dirname(part), '_rels', `${posix.basename(part)}.rels`)
    : '_rels/.rels';

/** The part a relationship of `source` points at, as a zip path. */
const resolveTarget = (source: string, target: string): string | undefined => {
  if (!target) return undefined;
  let path = target.split('#')[0];
  try {
    path = decodeURIComponent(path);
  } catch {
    // not percent-encoded
  }
  const resolved = path.startsWith('/')
    ? posix.normalize(path.slice(1))
    : posix.normalize(posix.join(posix.dirname(source) || '.', path));
  return resolved.startsWith('../') ? undefined : resolved;
};

/** Drops the content types of parts that are not in the package. */
const removeContentTypeOverrides = async (
  zip: JSZip,
  files: Set<string>,
): Promise<void> => {
  const xml = await readText(zip, '[Content_Types].xml');
  if (xml === undefined) return;
  let changed = false;
  const result = xml.replace(/<(?:\w+:)?Override\b[^>]*>/g, (tag) => {
    const name = attribute(tag, 'PartName');
    if (!name) return tag;
    let part = name.replace(/^\//, '');
    try {
      part = decodeURIComponent(part);
    } catch {
      // not percent-encoded
    }
    if (files.has(part)) return tag;
    changed = true;
    return '';
  });
  if (changed) zip.file('[Content_Types].xml', result);
};
