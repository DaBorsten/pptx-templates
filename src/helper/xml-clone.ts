/**
 * A faster `node.cloneNode(deep)` for xmldom nodes, with the same result.
 *
 * xmldom 0.9 copies the scalar properties of each node with a `for … in`
 * loop over the whole prototype chain (about 80 keys per element), which
 * made cloning most of the time `flow()` takes. This copies the same own
 * scalar properties onto a node made by the document's own factory (by
 * name, see copiesByName), attaches cloned attributes with
 * `setAttributeNode` and children with `appendChild`, as xmldom does.
 *
 * Elements, attributes and text nodes of XML documents take the fast path.
 * Anything else (comments, CDATA, processing instructions, HTML documents)
 * and any error on the way fall back to xmldom's own `cloneNode`.
 */

import { DOMParser } from '@xmldom/xmldom';

const ELEMENT_NODE = 1;
const ATTRIBUTE_NODE = 2;
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

/*
 * The own scalar properties xmldom gives each kind of node, copied by name.
 * Reading them by name is several times faster than Object.keys with
 * computed keys (most of the time a clone took). Same rule as cloneNode:
 * copied where different (`!=`). Properties the factory sets already
 * (tagName, name, data) are compared like the rest.
 */
type Scalars = Record<string, unknown>;

const copyElementScalars = (from: Scalars, to: Scalars): void => {
  if (from.nodeName != to.nodeName) to.nodeName = from.nodeName;
  if (from.tagName != to.tagName) to.tagName = from.tagName;
  if (from.namespaceURI != to.namespaceURI) to.namespaceURI = from.namespaceURI;
  if (from.prefix != to.prefix) to.prefix = from.prefix;
  if (from.localName != to.localName) to.localName = from.localName;
  if (from.lineNumber != to.lineNumber) to.lineNumber = from.lineNumber;
  if (from.columnNumber != to.columnNumber) to.columnNumber = from.columnNumber;
};

const copyAttributeScalars = (from: Scalars, to: Scalars): void => {
  if (from.namespaceURI != to.namespaceURI) to.namespaceURI = from.namespaceURI;
  if (from.prefix != to.prefix) to.prefix = from.prefix;
  if (from.nodeName != to.nodeName) to.nodeName = from.nodeName;
  if (from.name != to.name) to.name = from.name;
  if (from.specified != to.specified) to.specified = from.specified;
  if (from.localName != to.localName) to.localName = from.localName;
  if (from.lineNumber != to.lineNumber) to.lineNumber = from.lineNumber;
  if (from.columnNumber != to.columnNumber) to.columnNumber = from.columnNumber;
  if (from.nodeValue != to.nodeValue) to.nodeValue = from.nodeValue;
  if (from.value != to.value) to.value = from.value;
};

const copyTextScalars = (from: Scalars, to: Scalars): void => {
  if (from._data != to._data) to._data = from._data;
  if (from.length != to.length) to.length = from.length;
  if (from.lineNumber != to.lineNumber) to.lineNumber = from.lineNumber;
  if (from.columnNumber != to.columnNumber) to.columnNumber = from.columnNumber;
};

/** The own scalar properties each copy function above handles. */
const HANDLED: Record<number, Set<string>> = {
  [ELEMENT_NODE]: new Set([
    'nodeName',
    'tagName',
    'namespaceURI',
    'prefix',
    'localName',
    'lineNumber',
    'columnNumber',
  ]),
  [ATTRIBUTE_NODE]: new Set([
    'namespaceURI',
    'prefix',
    'nodeName',
    'name',
    'specified',
    'localName',
    'lineNumber',
    'columnNumber',
    'nodeValue',
    'value',
  ]),
  [TEXT_NODE]: new Set(['_data', 'length', 'lineNumber', 'columnNumber']),
};

/**
 * Whether the copy by name covers every own scalar property xmldom gives
 * its nodes: checked once per process on a parsed sample (namespaces,
 * prefixed attributes, text), its clone and nodes made by the factories.
 * A property it does not know (another xmldom version): the generic copy.
 * A process-wide fact about the installed xmldom, not per-render state.
 */
let byName: boolean | undefined;

export const copiesByName = (): boolean => {
  if (byName !== undefined) return byName;
  byName = false;
  try {
    const sample = new DOMParser().parseFromString(
      '<a:r xmlns:a="urn:a" xmlns:r="urn:r" r:id="1" b="x"><a:t>t</a:t></a:r>',
      'application/xml',
    ) as unknown as { documentElement: DomNode } & DomDocument;
    const root = sample.documentElement;
    const nodes: DomNode[] = [];
    const collect = (node: DomNode) => {
      nodes.push(node);
      const attrs = node.attributes;
      for (let i = 0; attrs && i < attrs.length; i++) nodes.push(attrs.item(i));
      for (let c = node.firstChild; c; c = c.nextSibling) collect(c);
    };
    collect(root);
    collect(root.cloneNode(true));
    nodes.push(
      sample.createElement('a:p'),
      sample.createAttribute('x'),
      sample.createTextNode('t'),
    );
    byName = nodes.every((node) => {
      const handled = HANDLED[node.nodeType];
      const own = node as unknown as Scalars;
      return (
        !!handled &&
        Object.keys(own).every(
          (key) => typeof own[key] === 'object' || handled.has(key),
        )
      );
    });
  } catch {
    byName = false;
  }
  return byName;
};

const cloneAttribute = (
  doc: DomDocument,
  attr: DomNode,
  fast: boolean,
): DomNode => {
  const copy = doc.createAttribute(attr.name);
  if (fast) {
    copyAttributeScalars(
      attr as unknown as Scalars,
      copy as unknown as Scalars,
    );
  } else copyScalars(attr, copy);
  return copy;
};

/** One node without its children, or undefined: use xmldom's cloneNode. */
const cloneShallow = (
  doc: DomDocument,
  node: DomNode,
  fast: boolean,
): DomNode | undefined => {
  if (node.nodeType === TEXT_NODE) {
    const copy = doc.createTextNode(node.data);
    if (fast) {
      copyTextScalars(node as unknown as Scalars, copy as unknown as Scalars);
    } else copyScalars(node, copy);
    return copy;
  }
  if (node.nodeType !== ELEMENT_NODE) return undefined;
  const copy = doc.createElement(node.tagName);
  if (fast) {
    copyElementScalars(node as unknown as Scalars, copy as unknown as Scalars);
  } else copyScalars(node, copy);
  const attrs = node.attributes;
  for (let i = 0; i < attrs.length; i++) {
    copy.setAttributeNode(cloneAttribute(doc, attrs.item(i), fast));
  }
  return copy;
};

const fastClone = (
  doc: DomDocument,
  root: DomNode,
  deep: boolean,
  fast: boolean,
): DomNode => {
  const top = cloneShallow(doc, root, fast) ?? root.cloneNode(deep);
  if (!deep || root.nodeType !== ELEMENT_NODE) return top;
  // depth-first without recursion: [source parent, its copy]
  const stack: [DomNode, DomNode][] = [[root, top]];
  while (stack.length) {
    const [source, target] = stack.pop();
    for (let child = source.firstChild; child; child = child.nextSibling) {
      const shallow = cloneShallow(doc, child, fast);
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
    return fastClone(doc, dom, deep, copiesByName()) as unknown as T;
  } catch {
    return dom.cloneNode(deep) as unknown as T;
  }
};

/**
 * A copy of a parsed XML document, as parsing its source again would give
 * (same nodes, attributes, namespaces and serialization), or undefined:
 * parse it. Only for `application/xml` documents whose children are the
 * XML declaration, whitespace and the root element, as in OOXML parts.
 */
export const cloneDocument = <T>(source: T): T | undefined => {
  const doc = source as unknown as DomDocument & {
    nodeType: number;
    firstChild: DomNode | null;
    doctype?: unknown;
  };
  if (
    doc?.nodeType !== 9 ||
    doc.contentType !== 'application/xml' ||
    doc.type !== 'xml' ||
    doc.doctype ||
    !copiesByName()
  ) {
    return undefined;
  }
  try {
    // the parser's own document: implementation, doctype, contentType
    const copy = new DOMParser().parseFromString(
      '<x/>',
      'application/xml',
    ) as unknown as DomDocument & {
      documentElement: DomNode;
      removeChild(child: DomNode): DomNode;
      appendChild(child: DomNode): DomNode;
      createProcessingInstruction(target: string, data: string): DomNode;
    };
    copy.removeChild(copy.documentElement);
    for (let child = doc.firstChild; child; child = child.nextSibling) {
      if (child.nodeType === ELEMENT_NODE || child.nodeType === TEXT_NODE) {
        copy.appendChild(fastClone(copy, child, true, true));
      } else if (child.nodeType === 7) {
        const pi = child as unknown as Scalars & DomNode;
        const node = copy.createProcessingInstruction(
          pi.target as string,
          pi.data as string,
        ) as unknown as Scalars & DomNode;
        // a declaration without data has null, the factory makes ''
        if (pi._data !== node._data) node._data = pi._data;
        if (pi.length !== node.length) node.length = pi.length;
        if (pi.lineNumber != node.lineNumber) node.lineNumber = pi.lineNumber;
        if (pi.columnNumber != node.columnNumber) {
          node.columnNumber = pi.columnNumber;
        }
        copy.appendChild(node);
      } else {
        return undefined;
      }
    }
    return copy as unknown as T;
  } catch {
    return undefined;
  }
};
