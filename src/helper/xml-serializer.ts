import { XMLSerializer } from '@xmldom/xmldom';

/**
 * `new XMLSerializer().serializeToString(node)` of @xmldom/xmldom 0.9, with
 * the same output, about a third faster.
 *
 * xmldom walks the tree with closures and copies the list of visible
 * namespaces for every element (`namespaces.slice()`), which made
 * serializing a large share of writing a slide. This is the same algorithm
 * (default options: no node filter, not `requireWellFormed`) on one
 * namespace stack that each element pushes to and truncates again.
 *
 * Elements, text, documents and fragments are written here; comments,
 * CDATA sections, processing instructions, doctypes and entity references
 * are handed to xmldom one by one (they do not depend on namespaces).
 * HTML documents and elements in the HTML namespace, and an attribute as
 * the node to serialize, go to xmldom as a whole.
 */

const ELEMENT_NODE = 1;
const ATTRIBUTE_NODE = 2;
const TEXT_NODE = 3;
const DOCUMENT_NODE = 9;
const DOCUMENT_FRAGMENT_NODE = 11;

const NS_XML = 'http://www.w3.org/XML/1998/namespace';
const NS_XMLNS = 'http://www.w3.org/2000/xmlns/';
const NS_HTML = 'http://www.w3.org/1999/xhtml';

/** The parts of an xmldom node this file reads. */
type SNode = {
  nodeType: number;
  nodeName: string;
  tagName?: string;
  prefix?: string | null;
  localName?: string | null;
  namespaceURI?: string | null;
  name?: string;
  value?: string;
  data?: string;
  firstChild: SNode | null;
  nextSibling: SNode | null;
  ownerDocument?: { type?: string } | null;
  documentElement?: SNode | null;
  type?: string;
  attributes?: { length: number; item(i: number): SNode };
  lookupPrefix?(uri: string): string | null;
};

type Namespace = { prefix: string | null; namespace: string };

/** Thrown inside: the node goes to xmldom as a whole. */
const FALLBACK = Symbol('xmldom');

const encode = (c: string): string =>
  (c == '<' && '&lt;') ||
  (c == '>' && '&gt;') ||
  (c == '&' && '&amp;') ||
  (c == '"' && '&quot;') ||
  '&#' + c.charCodeAt(0) + ';';

// replace() only where needed: most values have nothing to escape
const ATTRIBUTE_SPECIAL = /[<>&"\t\n\r]/;
const TEXT_SPECIAL = /[<&>]/;

const attribute = (name: string, value: string): string =>
  ' ' +
  name +
  '="' +
  (ATTRIBUTE_SPECIAL.test(value)
    ? value.replace(/[<>&"\t\n\r]/g, encode)
    : value) +
  '"';

/** xmldom's needNamespaceDefine, for XML documents. */
const needsDefine = (node: SNode, namespaces: Namespace[]): boolean => {
  const prefix = node.prefix || '';
  const uri = node.namespaceURI;
  if (!uri) return false;
  if ((prefix === 'xml' && uri === NS_XML) || uri === NS_XMLNS) return false;
  for (let i = namespaces.length - 1; i >= 0; i--) {
    const ns = namespaces[i];
    if (ns.prefix === prefix) return ns.namespace !== uri;
  }
  return true;
};

let fallbackSerializer: XMLSerializer | undefined;
const xmldom = (node: SNode): string => {
  fallbackSerializer ??= new XMLSerializer();
  return fallbackSerializer.serializeToString(node as never);
};

// Built with +=: measured faster than pushing to an array and joining,
// also counting the copy V8 makes when the result is first read flat
const write = (node: SNode, namespaces: Namespace[]): string => {
  switch (node.nodeType) {
    case ELEMENT_NODE:
      return element(node, namespaces);
    case TEXT_NODE:
      return TEXT_SPECIAL.test(node.data)
        ? node.data.replace(/[<&>]/g, encode)
        : node.data;
    case DOCUMENT_NODE:
    case DOCUMENT_FRAGMENT_NODE: {
      let out = '';
      for (let c = node.firstChild; c; c = c.nextSibling) {
        out += write(c, namespaces);
      }
      return out;
    }
    case ATTRIBUTE_NODE:
      return attribute(node.name, node.value);
    default:
      // comment, CDATA, processing instruction, doctype, entity reference
      return xmldom(node);
  }
};

const element = (n: SNode, namespaces: Namespace[]): string => {
  if (n.namespaceURI === NS_HTML) throw FALLBACK;
  const attrs = n.attributes;
  const len = attrs.length;
  const nodeName = n.tagName;

  let prefixedNodeName = nodeName;
  if (!n.prefix && n.namespaceURI) {
    let defaultNS: string | undefined;
    for (let ai = 0; ai < len; ai++) {
      if (attrs.item(ai).name === 'xmlns') {
        defaultNS = attrs.item(ai).value;
        break;
      }
    }
    if (!defaultNS) {
      for (let i = namespaces.length - 1; i >= 0; i--) {
        const ns = namespaces[i];
        if (ns.prefix === '' && ns.namespace === n.namespaceURI) {
          defaultNS = ns.namespace;
          break;
        }
      }
    }
    if (defaultNS !== n.namespaceURI) {
      for (let i = namespaces.length - 1; i >= 0; i--) {
        const ns = namespaces[i];
        if (ns.namespace === n.namespaceURI) {
          if (ns.prefix) prefixedNodeName = ns.prefix + ':' + nodeName;
          break;
        }
      }
    }
  }

  let out = '<' + prefixedNodeName;
  // xmldom copies the list per element; pushing onto the same list and
  // truncating it afterwards gives every lookup the same entries
  const base = namespaces.length;
  for (let i = 0; i < len; i++) {
    const attr = attrs.item(i);
    if (attr.prefix == 'xmlns') {
      namespaces.push({ prefix: attr.localName, namespace: attr.value });
    } else if (attr.nodeName == 'xmlns') {
      namespaces.push({ prefix: '', namespace: attr.value });
    }
  }
  for (let i = 0; i < len; i++) {
    const attr = attrs.item(i);
    if (needsDefine(attr, namespaces)) {
      const attrPrefix = attr.prefix || '';
      const uri = attr.namespaceURI;
      out += attribute(attrPrefix ? 'xmlns:' + attrPrefix : 'xmlns', uri);
      namespaces.push({ prefix: attrPrefix, namespace: uri });
    }
    out += attribute(attr.name, attr.value);
  }
  if (nodeName === prefixedNodeName && needsDefine(n, namespaces)) {
    const nodePrefix = n.prefix || '';
    const uri = n.namespaceURI;
    out += attribute(nodePrefix ? 'xmlns:' + nodePrefix : 'xmlns', uri);
    namespaces.push({ prefix: nodePrefix, namespace: uri });
  }

  if (!n.firstChild) {
    namespaces.length = base;
    return out + '/>';
  }
  out += '>';
  for (let c = n.firstChild; c; c = c.nextSibling) {
    out += write(c, namespaces);
  }
  namespaces.length = base;
  return out + '</' + prefixedNodeName + '>';
};

/** `new XMLSerializer().serializeToString(node)`, faster. */
export const serializeXml = (node: unknown): string => {
  const n = node as SNode;
  const doc = n.nodeType === DOCUMENT_NODE ? n : n.ownerDocument;
  if (!doc || doc.type === 'html' || n.nodeType === ATTRIBUTE_NODE) {
    return xmldom(n);
  }
  // nodeSerializeToString: the namespace of the root element without a
  // prefix of its own counts as declared
  const refNode = (n.nodeType === DOCUMENT_NODE && n.documentElement) || n;
  const namespaces: Namespace[] = [];
  const uri = refNode.namespaceURI;
  if (uri && refNode.prefix == null && refNode.lookupPrefix(uri) == null) {
    namespaces.push({ namespace: uri, prefix: null });
  }
  try {
    return write(n, namespaces);
  } catch (error) {
    // FALLBACK, or a tree deeper than the call stack
    if (error === FALLBACK || error instanceof RangeError) return xmldom(n);
    throw error;
  }
};
