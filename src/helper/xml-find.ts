import { XmlDocument, XmlElement } from '../types/xml-types';

/**
 * `root.getElementsByTagName(tagName)` as an array, in the same order
 * (depth-first, document order, `root` itself not included), about three
 * times faster: xmldom builds a live list with a callback per node. A
 * snapshot, not live. Documents of type 'html' (case-insensitive names) go
 * to xmldom.
 */
type Walked = {
  nodeType: number;
  tagName?: string;
  firstChild: Walked | null;
  nextSibling: Walked | null;
  parentNode: Walked | null;
  ownerDocument?: { type?: string } | null;
  type?: string;
};

const isHtml = (root: Walked): boolean =>
  (root.nodeType === 9 ? root.type : root.ownerDocument?.type) === 'html';

const walk = (
  root: Walked,
  tagName: string,
  visit: (el: Walked) => boolean | void,
): void => {
  const any = tagName === '*';
  let n = root.firstChild;
  while (n) {
    if (n.nodeType === 1 && (any || n.tagName === tagName) && visit(n)) {
      return;
    }
    if (n.firstChild) {
      n = n.firstChild;
      continue;
    }
    while (n && n !== root && !n.nextSibling) n = n.parentNode;
    n = n && n !== root ? n.nextSibling : null;
  }
};

export const elementsByTagName = (
  root: XmlDocument | XmlElement,
  tagName: string,
): XmlElement[] => {
  const r = root as unknown as Walked;
  if (isHtml(r)) {
    return Array.from(
      (root as XmlElement).getElementsByTagName(tagName),
    ) as XmlElement[];
  }
  const found: XmlElement[] = [];
  walk(r, tagName, (el) => {
    found.push(el as unknown as XmlElement);
  });
  return found;
};

/** `root.getElementsByTagName(tagName)[0]`, stopping at the first. */
export const firstElementByTagName = (
  root: XmlDocument | XmlElement,
  tagName: string,
): XmlElement | undefined => {
  const r = root as unknown as Walked;
  if (isHtml(r)) {
    return ((root as XmlElement).getElementsByTagName(tagName)[0] ??
      undefined) as XmlElement | undefined;
  }
  let found: XmlElement | undefined;
  walk(r, tagName, (el) => {
    found = el as unknown as XmlElement;
    return true;
  });
  return found;
};
