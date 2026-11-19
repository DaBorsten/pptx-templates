/**
 * A faster `node.cloneNode(deep)` for xmldom nodes, with the same result.
 *
 * xmldom 0.9 copies the scalar properties of each node with a `for … in`
 * loop over the whole prototype chain (about 80 keys per element), which
 * made cloning most of the time `flow()` takes. This copies the same own
 * scalar properties with `Object.keys` onto a node made by the document's
 * own factory, attaches cloned attributes with `setAttributeNode` and
 * children with `appendChild`, as xmldom does.
 *
 * Elements, attributes and text nodes of XML documents take the fast path.
 * Anything else (comments, CDATA, processing instructions, HTML documents)
 * and any error on the way fall back to xmldom's own `cloneNode`.
 */

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

/** The parts of an xmldom node this file touches. */
type DomNode = {
  nodeType: number;
  ownerDocument: DomDocument | null;
  firstChild: DomNode | null;
  nextSibling: DomNode | null;
  tagName?: string;
  data?: string;
  name?: string;
  attributes?: { length: number; item(i: number): DomNode | null };
  cloneNode(deep?: boolean): DomNode;
  appendChild(child: DomNode): DomNode;
  setAttributeNode?(attr: DomNode): DomNode | null;
};

type DomDocument = {
  type?: string;
  contentType?: string;
  createElement(tagName: string): DomNode;
  createAttribute(name: string): DomNode;
  createTextNode(data: string): DomNode;
};

/** Documents whose factories behave as xmldom's cloneNode (no HTML rules). */
const isPlainXml = (doc: DomDocument | null): doc is DomDocument =>
  !!doc &&
  doc.type !== 'html' &&
  doc.contentType !== 'text/html' &&
  doc.contentType !== 'application/xhtml+xml';

/** What xmldom's cloneNode copies: own properties that are not objects. */
const copyScalars = (source: DomNode, target: DomNode): void => {
  const from = source as unknown as Record<string, unknown>;
  const to = target as unknown as Record<string, unknown>;
  for (const key of Object.keys(from)) {
    const value = from[key];
    if (typeof value !== 'object' && value != to[key]) to[key] = value;
  }
};

const cloneAttribute = (doc: DomDocument, attr: DomNode): DomNode => {
  const copy = doc.createAttribute(attr.name);
  copyScalars(attr, copy);
  return copy;
};

/** One node without its children, or undefined: use xmldom's cloneNode. */
const cloneShallow = (doc: DomDocument, node: DomNode): DomNode | undefined => {
  if (node.nodeType === TEXT_NODE) {
    const copy = doc.createTextNode(node.data);
    copyScalars(node, copy);
    return copy;
  }
  if (node.nodeType !== ELEMENT_NODE) return undefined;
  const copy = doc.createElement(node.tagName);
  copyScalars(node, copy);
  const attrs = node.attributes;
  for (let i = 0; i < attrs.length; i++) {
    copy.setAttributeNode(cloneAttribute(doc, attrs.item(i)));
  }
  return copy;
};

const fastClone = (doc: DomDocument, root: DomNode, deep: boolean): DomNode => {
  const top = cloneShallow(doc, root) ?? root.cloneNode(deep);
  if (!deep || root.nodeType !== ELEMENT_NODE) return top;
  // depth-first without recursion: [source parent, its copy]
  const stack: [DomNode, DomNode][] = [[root, top]];
  while (stack.length) {
    const [source, target] = stack.pop();
    for (let child = source.firstChild; child; child = child.nextSibling) {
      const shallow = cloneShallow(doc, child);
      if (shallow) {
        target.appendChild(shallow);
        if (child.firstChild) stack.push([child, shallow]);
      } else {
        target.appendChild(child.cloneNode(true));
      }
    }
  }
  return top;
};

/** `node.cloneNode(deep)`, faster for elements and text of XML documents. */
export const cloneXml = <T>(node: T, deep = false): T => {
  const dom = node as unknown as DomNode;
  const doc = dom.ownerDocument;
  if (!isPlainXml(doc)) return dom.cloneNode(deep) as unknown as T;
  try {
    return fastClone(doc, dom, deep) as unknown as T;
  } catch {
    return dom.cloneNode(deep) as unknown as T;
  }
};
