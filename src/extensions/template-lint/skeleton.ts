import { isForbiddenKey } from '../template-logic/expression';
import type { PathUse } from './lint';
import { formatPath } from './usage';

/**
 * Example data built from the tags of one or more templates: every path a
 * tag reads, with a placeholder value.
 */
export type TemplateDataSkeleton = {
  /**
   * The data: objects for paths with properties, a list with one entry for
   * `FOR` and list methods, `true` for values that are only tested
   * (`{{IF kunde.vip}}`), `1` for numbers (`preis.toFixed(2)`) and the path
   * itself as text for everything that is output (`"kunde.name"`).
   */
  data: Record<string, unknown>;
  /**
   * Names called as functions (`{{euro(preis)}}`): helpers, not data. Pass
   * them as `additionalJsContext`.
   */
  functions: string[];
  /** Tags with errors; they are skipped. `lint` names them. */
  errorCount: number;
};

type SkeletonNode = {
  kind: 'unknown' | 'object' | 'list' | 'value';
  children: Map<string, SkeletonNode>;
  element?: SkeletonNode;
  /** Output or passed on, not only tested. */
  output: boolean;
  number: boolean;
};

const create = (): SkeletonNode => ({
  kind: 'unknown',
  children: new Map(),
  output: false,
  number: false,
});

const isIndex = (key: string) => /^\d+$/.test(key);

const has = (proto: object, key: string) =>
  Object.prototype.hasOwnProperty.call(proto, key);

/** Collects the paths the tags read (see `UsageSink` in lint.ts). */
export class DataSkeletonRecorder {
  private readonly root = create();
  private readonly functions = new Set<string>();

  constructor() {
    this.root.kind = 'object';
  }

  record(
    keys: string[],
    whole: boolean,
    _location?: unknown,
    { called = false, operand }: PathUse = {},
  ): void {
    // `kunde.__proto__`: a tag error, never a field of the data
    if (!keys.length || keys.some(isForbiddenKey)) return;
    if (called && keys.length === 1) {
      this.functions.add(keys[0]);
      return;
    }
    const walk = called ? keys.slice(0, -1) : keys;
    let node = this.root;
    for (const key of walk) {
      if (key === '[]' || isIndex(key)) {
        node = this.asList(node);
        continue;
      }
      // `name.length`: a property of the value. Any other name that is not
      // called is a field, also one a string or list has (`kunde.link`,
      // `config.values`)
      if (key === 'length' && node !== this.root && node.kind !== 'object') {
        this.builtin(node, key);
        return;
      }
      if (node.kind === 'list') return;
      node.kind = 'object';
      if (!node.children.has(key)) node.children.set(key, create());
      node = node.children.get(key);
    }
    if (called) {
      // `preis.toFixed(2)`, `produkte.map(…)`: a method of the value;
      // `kunde.format()`: a helper on an object
      const method = keys[keys.length - 1];
      if (node.kind === 'object' || !this.builtin(node, method)) {
        this.functions.add(formatPath(keys));
        if (node.kind === 'unknown') node.kind = 'object';
      }
      return;
    }
    if (operand === 'number') node.number = true;
    if (whole || operand === 'add') node.output = true;
  }

  recordList(keys: string[]): void {
    let node = this.root;
    for (const key of keys) {
      if (key === '[]' || isIndex(key)) {
        node = this.asList(node);
        continue;
      }
      if (node.kind !== 'object') return;
      node = node.children.get(key);
      if (!node) return;
    }
    this.asList(node);
  }

  /** Makes `node` a list (unless it has fields); returns its entry. */
  private asList(node: SkeletonNode): SkeletonNode {
    if (node.kind === 'object') return create();
    node.kind = 'list';
    node.element ??= create();
    return node.element;
  }

  /** Applies what a built-in key says about the value; false if none. */
  private builtin(node: SkeletonNode, key: string): boolean {
    const onArray = has(Array.prototype, key);
    const onString = has(String.prototype, key);
    const onNumber = has(Number.prototype, key);
    if (onArray && !onString) {
      // `namen.join(', ')`, `produkte.map(…)`: the entries are used
      if (node.kind !== 'object') this.asList(node).output = true;
      return true;
    }
    if (onNumber && !onString) {
      if (node.kind === 'unknown') node.kind = 'value';
      node.number = true;
      return true;
    }
    if (onString || has(Object.prototype, key)) {
      if (node.kind === 'unknown') node.kind = 'value';
      if (node.kind === 'value') node.output = true;
      return true;
    }
    return false;
  }

  result(errorCount = 0): TemplateDataSkeleton {
    const value = (node: SkeletonNode | undefined, keys: string[]): unknown => {
      if (node?.kind === 'object') {
        return Object.fromEntries(
          [...node.children].map(([key, child]) => [
            key,
            value(child, [...keys, key]),
          ]),
        );
      }
      if (node?.kind === 'list') {
        return [value(node.element, [...keys, '[]'])];
      }
      if (node?.number) return 1;
      if (node && !node.output) return true;
      return formatPath(keys);
    };
    return {
      data: value(this.root, []) as Record<string, unknown>,
      functions: [...this.functions],
      errorCount,
    };
  }
}

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value);

/**
 * Adds what `skeleton` has and `existing` lacks; existing values are kept.
 * List entries that are objects each get the missing fields of the
 * skeleton's entry. Returns the merged data and the added paths.
 */
export const mergeDataSkeleton = (
  existing: Record<string, unknown>,
  skeleton: Record<string, unknown>,
): { data: Record<string, unknown>; added: string[] } => {
  const added = new Set<string>();
  const merge = (
    current: unknown,
    wanted: unknown,
    keys: string[],
  ): unknown => {
    if (current === undefined) {
      added.add(formatPath(keys));
      return wanted;
    }
    if (isPlainObject(current) && isPlainObject(wanted)) {
      // Own fields only: `toString` is missing unless the data has it
      const fields = new Map(Object.entries(current));
      for (const [key, child] of Object.entries(wanted)) {
        fields.set(key, merge(fields.get(key), child, [...keys, key]));
      }
      return Object.fromEntries(fields);
    }
    if (Array.isArray(current) && Array.isArray(wanted)) {
      const entry = wanted[0];
      if (!isPlainObject(entry)) return current;
      return current.map((item) =>
        isPlainObject(item) ? merge(item, entry, [...keys, '[]']) : item,
      );
    }
    return current;
  };
  const data = merge(existing, skeleton, []) as Record<string, unknown>;
  return { data, added: [...added] };
};
