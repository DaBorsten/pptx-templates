import {
  LintTemplateOptions,
  TemplateLintIssue,
  TemplateLintRule,
  TemplateLintSeverity,
} from '../src/index';
import {
  chartXml,
  lint,
  partXml,
  shape,
  slideRels,
  smartArtXml,
  table,
} from './helpers/lint-deck';

/**
 * One case per use case of the template linter: what it must report (rule,
 * severity, and the characters it marks) and what it must not report.
 * docs/template-lint.md lists the rules these cases stand for.
 */

type Case = {
  /** Paragraphs of one text shape on slide 1 (each string one paragraph). */
  text?: string[];
  /** Slide 1 contents instead of `text`, e.g. tables. */
  slide?: string;
  options?: LintTemplateOptions;
  parts?: Record<string, string>;
  /** [rule, severity, marked text]; marked = tag.slice(range), else tag. */
  issues: [TemplateLintRule, TemplateLintSeverity, string][];
  /** Expected hint of the first issue. */
  hint?: RegExp;
};

const NBSP = String.fromCharCode(0xa0);

const data = {
  user: { name: 'Ada', vip: true, orders: [{ total: 3 }] },
  items: [{ name: 'A', price: 1 }],
  empty: [] as unknown[],
  nothing: null as unknown,
  count: 3,
  created: new Date(0),
  money: (n: number) => `${n} €`,
};

const schema = {
  type: 'object',
  properties: {
    user: { $ref: '#/$defs/user' },
    tags: { type: 'array', items: { type: 'string' } },
    meta: { type: 'object' },
    extra: { type: 'object', properties: { a: { type: 'number' } }, additionalProperties: true },
    maybe: {
      anyOf: [{ type: 'null' }, { type: 'object', properties: { x: { type: 'number' } } }],
    },
  },
  $defs: { user: { type: 'object', properties: { name: { type: 'string' } } } },
};

/** An invalid pattern ("(") must not hide the patterns after it. */
const patternSchema = {
  type: 'object',
  properties: {},
  patternProperties: {
    '(': { type: 'object' },
    '^x_': { type: 'object', properties: { a: { type: 'string' } } },
  },
};

const marked = (issue: TemplateLintIssue) =>
  issue.range ? issue.tag.slice(issue.range.start, issue.range.end) : issue.tag;

const run = async (c: Case) => {
  const slide = c.slide ?? shape('S', c.text.map((p) => [p]));
  const issues = await lint([slide], c.options, c.parts);
  if (c.hint) expect(issues[0]?.hint).toMatch(c.hint);
  return issues.map((i) => [i.rule, i.severity, marked(i)]);
};

/** Cases the linter has to report. */
const reported: Record<string, Case> = {
  // syntax
  'empty tag': { text: ['{{}}'], issues: [['syntax', 'error', '{{}}']] },
  'blank tag': { text: ['{{ }}'], issues: [['syntax', 'error', '{{ }}']] },
  'unterminated string': {
    text: ["{{'abc}}"],
    issues: [['syntax', 'error', "'abc"]],
  },
  'string closed with another quote': {
    text: ["{{euro(p) + 'Test`}}"],
    issues: [['syntax', 'error', "'Test`"]],
  },
  'typographic quotes': {
    text: ['{{’Ada’}}'],
    issues: [['syntax', 'error', '’']],
    hint: /Strings need straight quotes.*fixSmartQuotes/,
  },
  'typographic double quotes': { text: ['{{“Ada”}}'], issues: [['syntax', 'error', '“']] },
  'straight open, typographic close': {
    text: ["{{'Ada’}}"],
    issues: [['syntax', 'error', "'Ada’"]],
  },
  '= instead of ==': {
    text: ['{{IF a = 1}}{{END-IF}}'],
    issues: [['syntax', 'error', '=']],
    hint: /Use == or ===/,
  },
  'AND instead of &&': {
    text: ['{{IF a AND b}}{{END-IF}}'],
    issues: [['syntax', 'error', 'AND']],
    hint: /Use && for AND/,
  },
  'AutoCorrect …': {
    text: ['{{a…}}'],
    issues: [['syntax', 'error', '…']],
    hint: /replaced "\.\.\." with "…"/,
  },
  'AutoCorrect –': { text: ['{{a – 1}}'], issues: [['syntax', 'error', '–']] },
  'AutoCorrect →': { text: ['{{xs.map(x → x)}}'], issues: [['syntax', 'error', '→']] },
  'AutoCorrect ©': {
    text: ['{{FOR c IN xs.filter(©=> c)}}{{END-FOR}}'],
    issues: [['syntax', 'error', '©']],
    hint: /replaced "\(c\)" with "©"/,
  },
  'non-breaking spaces': {
    text: [`{{IF${NBSP}a${NBSP}=${NBSP}1}}{{END-IF}}`],
    issues: [['syntax', 'error', '=']],
  },
  'two expressions without operator': { text: ['{{a b}}'], issues: [['syntax', 'error', 'b']] },
  'missing closing bracket': { text: ['{{f(a}}'], issues: [['syntax', 'error', '}']] },
  'invalid EXEC statement': {
    text: ['{{EXEC total = 0; x = )}}'],
    issues: [['syntax', 'error', ')']],
  },
  'FOR without IN': {
    text: ['{{FOR x items}}{{END-FOR}}'],
    issues: [
      ['syntax', 'error', 'x'],
      ['unmatched-tag', 'error', '{{END-FOR}}'],
    ],
    hint: /Write \{\{FOR item IN list\}\}/,
  },
  'END IF with a space': {
    text: ['{{IF a}}x{{END IF}}'],
    issues: [
      ['syntax', 'error', 'IF'],
      ['unclosed-block', 'error', '{{IF a}}'],
    ],
    hint: /\{\{END-IF\}\} or \{\{END-FOR\}\}, with a hyphen/,
  },
  'forbidden key': {
    text: ['{{a.constructor}}'],
    issues: [['syntax', 'error', 'constructor']],
    hint: /blocked for safety/,
  },
  'forbidden __proto__': {
    text: ["{{a['__proto__']}}"],
    issues: [['syntax', 'error', "'__proto__'"]],
  },

  // unknown-alias
  'unknown alias': { text: ['{{*who}}'], issues: [['unknown-alias', 'error', '{{*who}}']] },

  // unterminated-tag
  '{{ without }}': { text: ['Hello {{user'], issues: [['unterminated-tag', 'error', '{{']] },
  '}} without {{': { text: ['Hello user}}'], issues: [['unterminated-tag', 'error', '}}']] },
  'tag split over two paragraphs': {
    text: ['{{user.', 'name}}'],
    issues: [
      ['unterminated-tag', 'error', '{{'],
      ['unterminated-tag', 'error', '}}'],
    ],
  },
  'tag split by a line break': {
    slide:
      '<p:sp><p:nvSpPr><p:cNvPr id="4" name="S"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>' +
      '<a:p><a:r><a:t>{{user.</a:t></a:r><a:br/><a:r><a:t>name}}</a:t></a:r></a:p></p:txBody></p:sp>',
    issues: [
      ['unterminated-tag', 'error', '{{'],
      ['unterminated-tag', 'error', '}}'],
    ],
  },
  'unpaired delimiter (same opening and closing)': {
    text: ['+++name+++ and +++rest'],
    options: { cmdDelimiter: '+++' },
    issues: [['unterminated-tag', 'error', '+++']],
  },

  // unclosed-block
  'IF without END-IF': {
    text: ['{{IF a}}', 'x'],
    issues: [['unclosed-block', 'error', '{{IF a}}']],
  },
  'FOR without END-FOR': {
    text: ['{{FOR x IN xs}}'],
    issues: [['unclosed-block', 'error', '{{FOR x IN xs}}']],
  },
  'block closed in another shape': {
    slide: shape('A', [['{{IF a}}']]) + shape('B', [['{{END-IF}}']]),
    issues: [
      ['unclosed-block', 'error', '{{IF a}}'],
      ['unmatched-tag', 'error', '{{END-IF}}'],
    ],
  },
  'inner block not closed before outer END': {
    text: ['{{FOR x IN xs}}{{IF x}}{{END-FOR}}'],
    issues: [['unclosed-block', 'error', '{{END-FOR}}']],
  },
  'FOR not closed in table': {
    slide: table('T', [['{{FOR r IN rows}}{{r}}', 'x']]),
    issues: [['unclosed-block', 'error', '{{FOR r IN rows}}']],
  },

  // unmatched-tag
  'END-IF without IF': { text: ['{{END-IF}}'], issues: [['unmatched-tag', 'error', '{{END-IF}}']] },
  'END-FOR without FOR': {
    text: ['{{END-FOR}}'],
    issues: [['unmatched-tag', 'error', '{{END-FOR}}']],
  },
  'END-FOR inside IF': {
    text: ['{{IF a}}{{END-FOR}}{{END-IF}}'],
    issues: [['unmatched-tag', 'error', '{{END-FOR}}']],
  },
  'ELSE without IF': { text: ['{{ELSE}}'], issues: [['unmatched-tag', 'error', '{{ELSE}}']] },
  'ELSE-IF without IF': {
    text: ['{{ELSE-IF a}}'],
    issues: [['unmatched-tag', 'error', '{{ELSE-IF a}}']],
  },
  'ELSE inside FOR': {
    text: ['{{FOR x IN xs}}{{ELSE}}{{END-FOR}}'],
    issues: [['unmatched-tag', 'error', '{{ELSE}}']],
  },
  'ELSE after ELSE': {
    text: ['{{IF a}}{{ELSE}}{{ELSE}}{{END-IF}}'],
    issues: [['unmatched-tag', 'error', '{{ELSE}}']],
  },
  'ELSE-IF after ELSE': {
    text: ['{{IF a}}{{ELSE}}{{ELSE-IF b}}{{END-IF}}'],
    issues: [['unmatched-tag', 'error', '{{ELSE-IF b}}']],
  },

  // unknown-variable
  'unknown name': {
    text: ['{{usr.name}}'],
    options: { data },
    issues: [['unknown-variable', 'warning', 'usr']],
  },
  'unknown function': {
    text: ['{{euro(1)}}'],
    options: { data },
    issues: [['unknown-variable', 'warning', 'euro']],
  },
  'unknown name in schema': {
    text: ['{{missing}}'],
    options: { schema },
    issues: [['unknown-variable', 'warning', 'missing']],
  },
  'loop variable used after the loop': {
    text: ['{{FOR x IN items}}{{END-FOR}}{{x}}{{$idx}}'],
    options: { data },
    issues: [
      ['unknown-variable', 'warning', 'x'],
      ['unknown-variable', 'warning', '$idx'],
    ],
  },
  'unknown name in a condition': {
    text: ['{{IF usr.vip}}{{END-IF}}'],
    options: { data },
    issues: [['unknown-variable', 'warning', 'usr']],
  },
  'unknown name in an arrow function body': {
    text: ['{{items.filter(i => i.price > limit).length}}'],
    options: { data },
    issues: [['unknown-variable', 'warning', 'limit']],
  },
  'unknown name in EXEC': {
    text: ['{{EXEC total = sum}}'],
    options: { data },
    issues: [['unknown-variable', 'warning', 'sum']],
  },

  // unknown-property
  'typo in a property': {
    text: ['{{user.nmae}}'],
    options: { data },
    issues: [['unknown-property', 'warning', 'nmae']],
  },
  'typo in a string method': {
    text: ['{{user.name.toUppercase()}}'],
    options: { data },
    issues: [['unknown-property', 'warning', 'toUppercase']],
  },
  'typo in an array property': {
    text: ['{{items.lenght}}'],
    options: { data },
    issues: [['unknown-property', 'warning', 'lenght']],
  },
  'property of a list entry by index': {
    text: ['{{items[0].nmae}}'],
    options: { data },
    issues: [['unknown-property', 'warning', 'nmae']],
  },
  'property of a loop entry': {
    text: ['{{FOR o IN user.orders}}{{o.totl}}{{$o.totl}}{{END-FOR}}'],
    options: { data },
    issues: [
      ['unknown-property', 'warning', 'totl'],
      ['unknown-property', 'warning', 'totl'],
    ],
  },
  'nested loop entry': {
    text: ['{{FOR u IN [user]}}{{FOR o IN user.orders}}{{o.price}}{{END-FOR}}{{END-FOR}}'],
    options: { data },
    issues: [['unknown-property', 'warning', 'price']],
  },
  'property of a number': {
    text: ['{{count.toFixd(2)}}'],
    options: { data },
    issues: [['unknown-property', 'warning', 'toFixd']],
  },
  'property not in a closed schema object': {
    text: ['{{user.mail}}'],
    options: { schema },
    issues: [['unknown-property', 'warning', 'mail']],
  },
  'property in no anyOf branch': {
    text: ['{{maybe.y}}'],
    options: { schema },
    issues: [['unknown-property', 'warning', 'y']],
  },
  'tags next to {{each:…}} markers in a table': {
    slide: table('T', [
      ['{{each:sub}}', '{{user.name}}'],
      ['{{each:row}}', '{{user.nmae}} {{each:keep}}', '{{IF user.vip}}x'],
    ]),
    options: { data },
    issues: [
      ['unknown-property', 'warning', 'nmae'],
      ['unclosed-block', 'error', '{{IF user.vip}}'],
    ],
  },
  'property behind a patternProperties pattern after an invalid one': {
    text: ['{{x_1.nope}}'],
    options: { schema: patternSchema },
    issues: [['unknown-property', 'warning', 'nope']],
  },
  'property of a schema list entry': {
    text: ['{{FOR t IN tags}}{{t.foo}}{{END-FOR}}'],
    options: { schema },
    issues: [['unknown-property', 'warning', 'foo']],
  },
  'property through an alias': {
    text: ['{{ALIAS n INS user.mial}}{{*n}}'],
    options: { data },
    issues: [['unknown-property', 'warning', '{{*n}}']],
  },

  // not-a-list
  'FOR over an object': {
    text: ['{{FOR u IN user}}{{END-FOR}}'],
    options: { data },
    issues: [['not-a-list', 'warning', 'user']],
  },
  'FOR over a number': {
    text: ['{{FOR n IN count}}{{END-FOR}}'],
    options: { data },
    issues: [['not-a-list', 'warning', 'count']],
  },
  'FOR over a schema object': {
    text: ['{{FOR u IN user}}{{END-FOR}}'],
    options: { schema },
    issues: [['not-a-list', 'warning', 'user']],
  },

  // unsupported-location
  'IMAGE in a table cell': {
    slide: table('T', [['{{IMAGE logo}}']]),
    issues: [['unsupported-location', 'error', '{{IMAGE logo}}']],
  },
  'IMAGE in a placeholder without own position': {
    slide: shape('P', [['{{IMAGE logo}}']], false),
    issues: [['unsupported-location', 'error', '{{IMAGE logo}}']],
  },
  'ELSE spanning table rows': {
    slide: table('T', [['{{IF a}}x'], ['{{ELSE}}'], ['{{END-IF}}']]),
    issues: [['unsupported-location', 'error', '{{ELSE}}']],
  },
  'ELSE-IF spanning table rows': {
    slide: table('T', [['{{IF a}}x'], ['{{ELSE-IF b}}'], ['{{END-IF}}']]),
    issues: [['unsupported-location', 'error', '{{ELSE-IF b}}']],
  },
  'tag in a slide layout': {
    text: [''],
    parts: { 'ppt/slideLayouts/slideLayout1.xml': partXml('sldLayout', shape('L', [['{{a}}']])) },
    issues: [['unsupported-location', 'warning', '{{a}}']],
  },
  'tag in a slide master': {
    text: [''],
    parts: { 'ppt/slideMasters/slideMaster1.xml': partXml('sldMaster', shape('M', [['{{a}}']])) },
    issues: [['unsupported-location', 'warning', '{{a}}']],
  },
  'tag in speaker notes': {
    text: [''],
    parts: {
      'ppt/slides/_rels/slide1.xml.rels': slideRels('../notesSlides/notesSlide1.xml'),
      'ppt/notesSlides/notesSlide1.xml': partXml('notes', shape('N', [['{{a}}']])),
    },
    issues: [['unsupported-location', 'warning', '{{a}}']],
  },
  'tag in a chart': {
    text: [''],
    parts: { 'ppt/charts/chart1.xml': chartXml('{{title}}') },
    issues: [['unsupported-location', 'warning', '{{title}}']],
  },
  'tag in SmartArt': {
    text: [''],
    parts: { 'ppt/diagrams/data1.xml': smartArtXml('{{step}}') },
    issues: [['unsupported-location', 'warning', '{{step}}']],
  },
};

/** Valid templates the linter must not report. */
const accepted: Record<string, Case> = {
  'text with quotes outside of tags': { text: ["Sie sind „VIP“ – {{a}} 'Test`"], issues: [] },
  'straight quotes': { text: [`{{IF a == 'x' || b == "y"}}{{END-IF}}`], issues: [] },
  'typographic quotes inside a string': { text: ["{{'Ada’s „Haus“'}}"], issues: [] },
  'typographic quotes with fixSmartQuotes': {
    text: ['{{IF a == ’x’}}{{“y”}}{{END-IF}}'],
    options: { fixSmartQuotes: true },
    issues: [],
  },
  'non-breaking spaces': { text: [`{{IF${NBSP}a}}x{{END-IF}}`], issues: [] },
  'keywords in lower case': {
    text: ['{{if a}}{{else-if b}}{{else}}{{end-if}}{{for x in xs}}{{endfor}}'],
    issues: [],
  },
  'IF, ELSE-IF and ELSE in one paragraph': {
    text: ['{{IF a}}1{{ELSE-IF b}}2{{ELSE}}3{{END-IF}}'],
    issues: [],
  },
  'blocks over several paragraphs': {
    text: ['{{FOR x IN xs}}', '{{IF x}}', 'y', '{{END-IF}}', '{{END-FOR}}'],
    issues: [],
  },
  'FOR over table rows': {
    slide: table('T', [['{{FOR r IN rows}}{{r.a}}', '{{r.b}}{{END-FOR}}']]),
    issues: [],
  },
  'IF over table rows': {
    slide: table('T', [['{{IF a}}'], ['x'], ['{{END-IF}}']]),
    issues: [],
  },
  'ELSE within one table cell': {
    slide: table('T', [['{{IF a}}x{{ELSE}}y{{END-IF}}']]),
    issues: [],
  },
  'IMAGE in a shape with own position': { text: ['{{IMAGE logo}}'], issues: [] },
  'other delimiters': {
    text: ['[[IF a]]x[[END-IF]] {{not a tag'],
    options: { openingTag: '[[', closingTag: ']]' },
    issues: [],
  },
  'any names without data or schema': { text: ['{{whatever.you.like}}'], issues: [] },
  'names from data': {
    text: ['{{user.name.toUpperCase()}} {{items.length}} {{items[0].price}} {{money(count)}}'],
    options: { data },
    issues: [],
  },
  'loop variables, $name and $idx': {
    text: ['{{FOR o IN user.orders}}{{o.total}}{{$o.total}}{{$idx + 1}}{{END-FOR}}'],
    options: { data },
    issues: [],
  },
  'arrow function parameters': {
    text: ['{{items.filter(i => i.price > 0).map((i, n) => i.name + n)}}'],
    options: { data },
    issues: [],
  },
  'EXEC variables, also in later shapes': {
    slide:
      shape('A', [['{{EXEC total = 0; label = "Sum"}}']]) +
      shape('B', [['{{label}}: {{total.toFixed(2)}}']]),
    options: { data },
    issues: [],
  },
  'ALIAS defined in another shape': {
    slide: shape('A', [['{{*who}}']]) + shape('B', [['{{ALIAS who INS user.name}}']]),
    options: { data },
    issues: [],
  },
  'additionalJsContext': {
    text: ['{{fmt(count)}}'],
    options: { data, additionalJsContext: { fmt: String } },
    issues: [],
  },
  'null, empty lists and class instances in the example': {
    text: [
      '{{nothing.any.thing}}',
      '{{FOR e IN empty}}{{e.anything}}{{END-FOR}}',
      '{{FOR n IN nothing}}{{END-FOR}}',
      '{{created.getFullYear()}}',
    ],
    options: { data },
    issues: [],
  },
  'open schema objects': {
    text: ['{{meta.anything}} {{extra.a}} {{extra.other}} {{maybe.x}} {{user.name.trim()}}'],
    options: { schema },
    issues: [],
  },
  'FOR over a string': {
    text: ['{{FOR c IN user.name}}{{c}}{{END-FOR}}'],
    options: { data },
    issues: [],
  },
  "setTable's {{each:…}} markers": {
    text: ['{{each:row}}', '{{each:subSub2}} {{ each:sub }}'],
    options: { data },
    issues: [],
  },
  'names an invalid patternProperties pattern might match': {
    text: ['{{other.anything}}'],
    options: { schema: patternSchema },
    issues: [],
  },
};

describe('the linter reports', () => {
  test.each(Object.entries(reported))('%s', async (_name, c) => {
    expect(await run(c)).toEqual(c.issues);
  });
});

describe('the linter accepts', () => {
  test.each(Object.entries(accepted))('%s', async (_name, c) => {
    expect(await run(c)).toEqual(c.issues);
  });
});
