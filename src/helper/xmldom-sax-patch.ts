/**
 * Performance patch for @xmldom/xmldom's SAX parser (verified on 0.9.12).
 *
 * `lib/sax.js` compiles a fresh RegExp for every closing tag it parses
 * (`g.reg('^', g.QName_group, g.S_OPT, '$')`) and for every `getMatch`
 * probe - measured at ~19% of total runtime when appending large slides,
 * since every part read from an archive goes through this parser.
 *
 * `grammar.reg` builds its RegExp purely from its arguments and returns it
 * without the `g` flag (no `lastIndex` state), and the runtime call sites
 * combine a small fixed set of grammar fragments - so caching the compiled
 * RegExp by source is safe and cache growth is bounded by the grammar.
 *
 * Applied via the exports object, which is the `g` the SAX parser calls
 * through; `grammar.js`-internal calls at module load are unaffected.
 * Remove once fixed upstream in @xmldom/xmldom.
 */

let applied = false;

export function patchXmldomSaxRegExpCache(): void {
  if (applied) {
    return;
  }
  applied = true;

  try {
    // Deep import on purpose: the patch targets the module object that
    // sax.js holds. Guarded, in case a future xmldom version restricts
    // deep imports via an `exports` map - the patch is an optimization
    // only, parsing works identically without it.
    // eslint-disable-next-line @typescript-eslint/no-require-imports
    const grammar = require('@xmldom/xmldom/lib/grammar');
    if (typeof grammar?.reg !== 'function') {
      return;
    }

    const original = grammar.reg;
    const cache = new Map<string, RegExp>();

    // sax.js calls with the same few argument lists (the same strings and
    // RegExp objects) for every closing tag: those are found by comparing
    // the arguments, without building the key
    const calls: { args: (string | RegExp)[]; compiled: RegExp }[] = [];
    const sameArgs = (a: (string | RegExp)[], b: (string | RegExp)[]) => {
      if (a.length !== b.length) return false;
      for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
      return true;
    };

    grammar.reg = function (...args: (string | RegExp)[]): RegExp {
      for (const call of calls) {
        if (sameArgs(call.args, args)) return call.compiled;
      }
      // Key = the exact source string `reg` compiles (same mapping and
      // join), so two calls share a cache entry iff they build the same
      // RegExp.
      const key = args
        .map((part) => (typeof part === 'string' ? part : part.source))
        .join('');
      let compiled = cache.get(key);
      if (!compiled) {
        compiled = original.apply(this, args);
        cache.set(key, compiled);
      }
      // same strings and RegExp objects: the same source, the same RegExp
      if (calls.length < 16) calls.push({ args, compiled });
      return compiled;
    };

    cacheQNameTests(grammar.QName_exact);
  } catch {
    // Patch is best-effort; never break parsing over it.
  }
}

/**
 * xmldom checks every element and attribute name it creates (parsing,
 * createElementNS, createAttribute) against `QName_exact`, a long Unicode
 * pattern. A document has few distinct names, so the result is cached per
 * name on that RegExp object (an own `test`; the pattern has no `g` flag,
 * so `test` has no state). Bounded: cleared when it grows large.
 */
const cacheQNameTests = (pattern: unknown): void => {
  if (!(pattern instanceof RegExp) || pattern.global || pattern.sticky) {
    return;
  }
  const test = RegExp.prototype.test;
  const results = new Map<string, boolean>();
  Object.defineProperty(pattern, 'test', {
    configurable: true,
    writable: true,
    value: function (input: string): boolean {
      if (typeof input !== 'string') return test.call(pattern, input);
      let result = results.get(input);
      if (result === undefined) {
        result = test.call(pattern, input);
        if (results.size > 10_000) results.clear();
        results.set(input, result);
      }
      return result;
    },
  });
};
