import { JsonSchema } from './types';

/** What the linter knows about a value: nothing, an example, or a schema. */
export type Known =
  | { kind: 'any' }
  | { kind: 'value'; value: unknown }
  | { kind: 'schema'; schema: JsonSchema };

/** A lookup that found nothing; `candidates` feed "did you mean". */
export type Missing = { kind: 'missing'; candidates: string[]; owner: string };

const ANY: Known = { kind: 'any' };

const PROTOTYPES: Record<string, object> = {
  string: String.prototype,
  number: Number.prototype,
  integer: Number.prototype,
  boolean: Boolean.prototype,
  array: Array.prototype,
  object: Object.prototype,
};

const isIndex = (key: string) => /^\d+$/.test(key);

const prototypeNames = (proto: object): string[] =>
  Object.getOwnPropertyNames(proto).filter((name) => name !== 'constructor');

const typesOf = (schema: JsonSchema): string[] => {
  if (Array.isArray(schema.type)) return schema.type;
  if (schema.type) return [schema.type];
  const types: string[] = [];
  if (schema.properties || schema.additionalProperties !== undefined) {
    types.push('object');
  }
  if (schema.items || schema.prefixItems) types.push('array');
  return types;
};

/**
 * Answers "does `user.name` exist?" against example data and/or a JSON
 * Schema. Whatever cannot be decided (null in the example, an open schema,
 * a computed value) counts as existing: the linter only reports what is
 * certainly missing.
 */
export class DataModel {
  constructor(
    private readonly data?: Record<string, unknown>,
    private readonly schema?: JsonSchema,
    private readonly extra: Record<string, unknown> = {},
  ) {}

  /** Whether names can be checked at all (data or schema given). */
  get enabled(): boolean {
    return this.data !== undefined || this.schema !== undefined;
  }

  root(name: string): Known | Missing {
    const candidates: string[] = [...Object.keys(this.extra)];
    if (this.data) {
      if (Object.prototype.hasOwnProperty.call(this.data, name)) {
        return { kind: 'value', value: this.data[name] };
      }
      candidates.push(...Object.keys(this.data));
    }
    if (this.schema) {
      const fromSchema = this.member(
        { kind: 'schema', schema: this.schema },
        name,
      );
      if (fromSchema.kind !== 'missing') return fromSchema;
      candidates.push(...fromSchema.candidates);
    }
    if (Object.prototype.hasOwnProperty.call(this.extra, name)) {
      return { kind: 'value', value: this.extra[name] };
    }
    return { kind: 'missing', candidates, owner: 'data' };
  }

  member(known: Known, key: string): Known | Missing {
    if (known.kind === 'value') return this.valueMember(known.value, key);
    if (known.kind === 'schema') return this.schemaMember(known.schema, key);
    return ANY;
  }

  /** The type of a FOR loop's entries; `undefined`: not a list. */
  element(known: Known): Known | undefined {
    if (known.kind === 'value') {
      const value = known.value;
      if (value === null || value === undefined) return ANY;
      if (Array.isArray(value)) {
        return value.length ? { kind: 'value', value: value[0] } : ANY;
      }
      return typeof (value as Iterable<unknown>)[Symbol.iterator] === 'function'
        ? ANY
        : undefined;
    }
    if (known.kind === 'schema') return this.schemaElement(known.schema, 0);
    return ANY;
  }

  private valueMember(value: unknown, key: string): Known | Missing {
    if (value === null || value === undefined || typeof value === 'function') {
      return ANY;
    }
    if (Array.isArray(value)) {
      if (isIndex(key)) {
        return value.length
          ? { kind: 'value', value: value[Number(key)] ?? value[0] }
          : ANY;
      }
      if (key in value) return ANY;
      return {
        kind: 'missing',
        candidates: prototypeNames(Array.prototype),
        owner: 'list',
      };
    }
    if (typeof value === 'object') {
      if (Object.prototype.hasOwnProperty.call(value, key)) {
        return {
          kind: 'value',
          value: (value as Record<string, unknown>)[key],
        };
      }
      if (key in value) return ANY;
      // Class instances (Date, Buffer, Map, …): only plain objects are closed
      const proto = Object.getPrototypeOf(value);
      if (proto !== Object.prototype && proto !== null) return ANY;
      return {
        kind: 'missing',
        candidates: Object.keys(value),
        owner: 'object',
      };
    }
    if (key in Object(value)) return ANY;
    return {
      kind: 'missing',
      candidates: prototypeNames(Object.getPrototypeOf(Object(value))),
      owner: typeof value,
    };
  }

  private schemaMember(
    input: JsonSchema,
    key: string,
    depth = 0,
  ): Known | Missing {
    const schema = this.resolve(input);
    if (!schema || depth > 20) return ANY;

    const branches = [
      ...(schema.allOf ?? []),
      ...(schema.anyOf ?? []),
      ...(schema.oneOf ?? []),
    ];
    const own = this.ownSchemaMember(schema, key);
    if (!branches.length || own.kind !== 'missing') return own;

    // Found in any branch counts as existing
    const candidates = [...own.candidates];
    for (const branch of branches) {
      const result = this.schemaMember(branch, key, depth + 1);
      if (result.kind !== 'missing') return result;
      candidates.push(...result.candidates);
    }
    return { ...own, candidates };
  }

  private ownSchemaMember(schema: JsonSchema, key: string): Known | Missing {
    const types = typesOf(schema);
    const hasCombinators = !!(schema.allOf || schema.anyOf || schema.oneOf);
    if (!types.length && !hasCombinators) return ANY;

    if (types.includes('array') && isIndex(key)) {
      const element = this.schemaElement(schema, Number(key));
      return element ?? ANY;
    }
    if (
      schema.properties &&
      Object.prototype.hasOwnProperty.call(schema.properties, key)
    ) {
      return { kind: 'schema', schema: schema.properties[key] };
    }
    // An invalid pattern cannot rule the key out, but the others still apply
    let invalidPattern = false;
    for (const [pattern, sub] of Object.entries(
      schema.patternProperties ?? {},
    )) {
      let matches: boolean;
      try {
        matches = new RegExp(pattern, 'u').test(key);
      } catch {
        invalidPattern = true;
        continue;
      }
      if (matches) return { kind: 'schema', schema: sub };
    }
    if (invalidPattern) return ANY;
    for (const type of types) {
      if (PROTOTYPES[type] && key in PROTOTYPES[type]) return ANY;
    }
    if (types.includes('object')) {
      const additional = schema.additionalProperties;
      if (additional && typeof additional === 'object') {
        return { kind: 'schema', schema: additional };
      }
      // Lint convention: listed properties close the object
      if (
        additional === true ||
        (!schema.properties && !schema.patternProperties)
      ) {
        return ANY;
      }
    }
    const candidates = [
      ...Object.keys(schema.properties ?? {}),
      ...types.flatMap((type) =>
        type === 'object' || !PROTOTYPES[type]
          ? []
          : prototypeNames(PROTOTYPES[type]),
      ),
    ];
    return {
      kind: 'missing',
      candidates,
      owner: types.length === 1 ? types[0] : 'value',
    };
  }

  private schemaElement(
    input: JsonSchema,
    index: number,
    depth = 0,
  ): Known | undefined {
    const schema = this.resolve(input);
    if (!schema || depth > 20) return ANY;
    const types = typesOf(schema);
    if (types.includes('array')) {
      const item =
        schema.prefixItems?.[index] ??
        (Array.isArray(schema.items) ? schema.items[index] : schema.items);
      return item ? { kind: 'schema', schema: item } : ANY;
    }
    // A string can be iterated (its characters)
    if (types.includes('string')) return ANY;
    const branches = [
      ...(schema.anyOf ?? []),
      ...(schema.oneOf ?? []),
      ...(schema.allOf ?? []),
    ];
    if (!types.length && !branches.length) return ANY;
    for (const branch of branches) {
      const element = this.schemaElement(branch, index, depth + 1);
      if (element) return element;
    }
    return types.includes('null') && types.length === 1 ? ANY : undefined;
  }

  /** Follows local `$ref`s (`#/$defs/…`, `#/definitions/…`, `#`). */
  private resolve(schema: JsonSchema): JsonSchema | undefined {
    let current = schema;
    for (let i = 0; current?.$ref && i < 20; i++) {
      const ref = current.$ref;
      if (!ref.startsWith('#')) return undefined;
      let target: unknown = this.schema;
      for (const segment of ref.slice(1).split('/').filter(Boolean)) {
        let decoded: string;
        try {
          decoded = decodeURIComponent(segment);
        } catch {
          return undefined; // malformed %-escape: undecidable, like a remote ref
        }
        const key = decoded.replace(/~1/g, '/').replace(/~0/g, '~');
        target = (target as Record<string, unknown>)?.[key];
      }
      current = target as JsonSchema;
    }
    return current;
  }
}

/** Closest candidate within a small edit distance, for "did you mean". */
export const suggest = (
  name: string,
  candidates: string[],
): string | undefined => {
  const maxDistance = name.length <= 2 ? 0 : name.length <= 5 ? 1 : 2;
  let best: string | undefined;
  let bestDistance = maxDistance + 1;
  for (const candidate of new Set(candidates)) {
    const distance = editDistance(name.toLowerCase(), candidate.toLowerCase());
    if (distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  }
  return best;
};

/** Levenshtein distance where swapping two neighbours counts as one edit. */
const editDistance = (a: string, b: string): number => {
  const d: number[][] = Array.from({ length: a.length + 1 }, (_, i) =>
    Array.from({ length: b.length + 1 }, (_, j) =>
      i === 0 ? j : j === 0 ? i : 0,
    ),
  );
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      d[i][j] = Math.min(
        d[i - 1][j] + 1,
        d[i][j - 1] + 1,
        d[i - 1][j - 1] + cost,
      );
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
      }
    }
  }
  return d[a.length][b.length];
};
