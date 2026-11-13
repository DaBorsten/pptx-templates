import { JsonSchema } from './types';

/** Where a tag reads a path. */
export type DataUsageLocation = {
  /** Part inside the .pptx, e.g. `ppt/slides/slide2.xml`. */
  part: string;
  slide?: number;
  /** Shape name (Selection pane). */
  shape?: string;
};

export type DataUsageEntry = {
  /** `kunde.name`, `produkte[].preis` (`[]`: every entry of a list). */
  path: string;
  locations: DataUsageLocation[];
};

export type TemplateDataUsage = {
  /**
   * Paths the tags read, as deep as the data goes, in the order they first
   * appear. A value that is output or passed on as a whole (`{{kunde}}`,
   * `{{format(kunde)}}`, `{{produkte.map(…)}}`) covers everything below it.
   */
  used: DataUsageEntry[];
  /**
   * Fields of the data or schema no tag reads. A subtree no tag reads at
   * all is one entry (`intern`, not `intern.a`, `intern.b`). Empty without
   * `data` and `schema`.
   */
  unused: string[];
  /** Paths the tags read that are not in the data or schema. */
  missing: DataUsageEntry[];
};

type TreeNode = {
  kind: 'object' | 'array' | 'leaf';
  children: Map<string, TreeNode>;
  /** Entries of a list; undefined when nothing is known about them. */
  element?: TreeNode;
  /** Keys outside `children` may exist (null in the data, open schema). */
  open: boolean;
  /** Read by a tag. */
  read: boolean;
  /** Used as a whole: everything below counts as read. */
  whole: boolean;
};

const node = (kind: TreeNode['kind'], open = false): TreeNode => ({
  kind,
  children: new Map(),
  open,
  read: false,
  whole: false,
});

const isIndex = (key: string) => /^\d+$/.test(key);

const isPlainObject = (value: unknown): value is Record<string, unknown> => {
  if (!value || typeof value !== 'object') return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
};

/** Combines what two sources (two list entries, data and schema) say. */
const merge = (a: TreeNode | undefined, b: TreeNode | undefined) => {
  if (!a) return b;
  if (!b) return a;
  const rank = { leaf: 0, array: 1, object: 2 };
  const result = node(rank[a.kind] >= rank[b.kind] ? a.kind : b.kind);
  result.open = a.open || b.open;
  for (const source of [a, b]) {
    source.children.forEach((child, key) =>
      result.children.set(key, merge(result.children.get(key), child)),
    );
  }
  result.element = merge(a.element, b.element);
  return result;
};

const fromData = (value: unknown): TreeNode => {
  if (value === null || value === undefined) return node('leaf', true);
  if (Array.isArray(value)) {
    const list = node('array');
    // Entries may differ: their fields are combined
    value.forEach(
      (entry) => (list.element = merge(list.element, fromData(entry))),
    );
    return list;
  }
  if (isPlainObject(value)) {
    const object = node('object');
    Object.entries(value).forEach(([key, child]) =>
      object.children.set(key, fromData(child)),
    );
    return object;
  }
  // Strings, numbers, dates, …
  return node('leaf');
};

const fromSchema = (root: JsonSchema): TreeNode => {
  const resolve = (schema: JsonSchema): JsonSchema | undefined => {
    let current = schema;
    for (let i = 0; current?.$ref && i < 20; i++) {
      if (!current.$ref.startsWith('#')) return undefined;
      let target: unknown = root;
      for (const segment of current.$ref.slice(1).split('/').filter(Boolean)) {
        let key: string;
        try {
          key = decodeURIComponent(segment)
            .replace(/~1/g, '/')
            .replace(/~0/g, '~');
        } catch {
          return undefined;
        }
        target = (target as Record<string, unknown>)?.[key];
      }
      current = target as JsonSchema;
    }
    return current;
  };

  const build = (input: JsonSchema, depth: number): TreeNode => {
    const schema = resolve(input);
    // Unknown or too deep (recursive schemas): anything may be there
    if (!schema || typeof schema !== 'object' || depth > 20) {
      return node('leaf', true);
    }
    const types = Array.isArray(schema.type)
      ? schema.type
      : schema.type
        ? [schema.type]
        : [];
    let result: TreeNode | undefined;

    if (schema.properties || types.includes('object')) {
      const object = node('object');
      Object.entries(schema.properties ?? {}).forEach(([key, child]) =>
        object.children.set(key, build(child, depth + 1)),
      );
      // Same convention as the linter: listed properties close the object
      object.open =
        !!schema.patternProperties ||
        (schema.additionalProperties !== undefined &&
          schema.additionalProperties !== false) ||
        !schema.properties;
      if (
        schema.additionalProperties &&
        typeof schema.additionalProperties === 'object'
      ) {
        object.open = true;
      }
      result = object;
    }
    if (schema.items || schema.prefixItems || types.includes('array')) {
      const list = node('array');
      const items = [
        ...(schema.prefixItems ?? []),
        ...(Array.isArray(schema.items)
          ? schema.items
          : schema.items
            ? [schema.items]
            : []),
      ];
      items.forEach(
        (item) => (list.element = merge(list.element, build(item, depth + 1))),
      );
      result = merge(result, list);
    }
    for (const branch of [
      ...(schema.allOf ?? []),
      ...(schema.anyOf ?? []),
      ...(schema.oneOf ?? []),
    ]) {
      result = merge(result, build(branch, depth + 1));
    }
    if (result) return result;
    return types.length ? node('leaf') : node('leaf', true);
  };

  return build(root, 0);
};

const IDENTIFIER = /^[\p{L}_$][\p{L}\p{N}_$]*$/u;

const formatPath = (keys: string[]): string =>
  keys
    .map((key, i) => {
      if (key === '[]') return '[]';
      if (isIndex(key)) return `[${key}]`;
      if (IDENTIFIER.test(key)) return i === 0 ? key : `.${key}`;
      return `[${JSON.stringify(key)}]`;
    })
    .join('');

const PROTOTYPES = {
  object: Object.prototype,
  array: Array.prototype,
} as const;

/**
 * Collects which data paths the tags of a template read. Without data and
 * schema it only lists the paths; with them, it also knows which fields are
 * not read and which read paths do not exist.
 */
export class DataUsageRecorder {
  private readonly root?: TreeNode;
  private readonly used = new Map<string, DataUsageEntry>();
  private readonly missing = new Map<string, DataUsageEntry>();

  constructor(data?: Record<string, unknown>, schema?: JsonSchema) {
    if (data !== undefined) this.root = fromData(data);
    if (schema !== undefined) this.root = merge(this.root, fromSchema(schema));
  }

  /**
   * A tag reads `keys` (from the data root; list entries as `[]` or an
   * index). `whole`: the value is used as a whole, not only tested.
   */
  record(keys: string[], whole: boolean, location: DataUsageLocation): void {
    if (!this.root) {
      this.add(this.used, keys, location);
      return;
    }
    let current = this.root;
    const walked: string[] = [];
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      current.read = true;
      let next: TreeNode | undefined;
      if (current.kind === 'array' && (key === '[]' || isIndex(key))) {
        walked.push('[]');
        next = current.element;
        if (!next) return this.add(this.used, walked, location);
      } else if (current.kind === 'object' && current.children.has(key)) {
        walked.push(key);
        next = current.children.get(key);
      } else if (current.kind === 'leaf') {
        // A property of a string or number (`name.length`), or of a value
        // nothing is known about (null in the data): the value is read
        current.whole = true;
        return this.add(this.used, walked, location);
      } else if (current.open && !(key in PROTOTYPES[current.kind])) {
        // A key the data or schema allows without listing it
        return this.add(this.used, [...walked, ...keys.slice(i)], location);
      } else {
        const proto = PROTOTYPES[current.kind];
        if (key in proto) {
          // `produkte.map(p => …)` hands every entry to the function;
          // `produkte.length` reads none of them
          if (typeof (proto as Record<string, unknown>)[key] === 'function') {
            current.whole = true;
          }
          return this.add(this.used, walked, location);
        }
        return this.add(this.missing, [...walked, key], location);
      }
      current = next;
    }
    current.read = true;
    if (whole) current.whole = true;
    this.add(this.used, walked, location);
  }

  private add(
    into: Map<string, DataUsageEntry>,
    keys: string[],
    location: DataUsageLocation,
  ): void {
    // Reading the root itself (`{{FOR x IN [1, 2]}}`) is not a data path
    if (!keys.length) return;
    const path = formatPath(keys);
    let entry = into.get(path);
    if (!entry) {
      entry = { path, locations: [] };
      into.set(path, entry);
    }
    const known = entry.locations.some(
      (l) =>
        l.part === location.part &&
        l.slide === location.slide &&
        l.shape === location.shape,
    );
    if (!known) entry.locations.push(location);
  }

  result(): TemplateDataUsage {
    const unused: string[] = [];
    const anyRead = (n: TreeNode): boolean =>
      n.read ||
      n.whole ||
      [...n.children.values()].some(anyRead) ||
      (!!n.element && anyRead(n.element));
    const collect = (n: TreeNode, keys: string[]) => {
      if (n.whole) return;
      if (!anyRead(n)) {
        unused.push(formatPath(keys));
        return;
      }
      n.children.forEach((child, key) => collect(child, [...keys, key]));
      if (n.element) collect(n.element, [...keys, '[]']);
    };
    if (this.root && !this.root.whole) {
      this.root.children.forEach((child, key) => collect(child, [key]));
    }
    return {
      used: [...this.used.values()],
      unused,
      missing: [...this.missing.values()],
    };
  }
}
