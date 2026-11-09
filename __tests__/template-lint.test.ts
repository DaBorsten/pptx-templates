import Automizer, {
  lintTemplate,
  modify,
  TemplateLintIssue,
  XmlElement,
} from '../src/index';
import { runCli } from '../src/extensions/cli';
import { TemplateExpression } from '../src/extensions/template-logic/expression';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { deck, lint, partXml, shape, table } from './helpers/lint-deck';

const rules = (issues: TemplateLintIssue[]) => issues.map((i) => i.rule);

test('a valid template has no issues', async () => {
  const issues = await lint(
    [
      shape('Text', [
        ['Hello {{user.name}}!'],
        ['{{IF user.vip}}', 'VIP', '{{ELSE}}', 'regular{{END-IF}}'],
        ['{{FOR item IN items}}{{$idx + 1}}. {{item.name.toUpperCase()}}{{END-FOR}}'],
        ['{{EXEC total = 0}}{{total}}'],
      ]) + table('Table', [['{{FOR row IN items}}{{row.name}}', '{{row.price}}{{END-FOR}}']]),
    ],
    { data: { user: { name: 'Ada', vip: true }, items: [{ name: 'A', price: 1 }] } },
  );
  expect(issues).toEqual([]);
});

test("a deck with setTable's {{each:…}} markers has no issues", async () => {
  const { issues } = await lintTemplate(
    path.join(__dirname, 'pptx-templates', 'NestedTables.pptx'),
  );
  expect(issues).toEqual([]);
});

test('reports unclosed and unmatched blocks', async () => {
  const issues = await lint([
    shape('Open', [['{{IF a}}'], ['text']]) +
      shape('Close', [['{{END-IF}}']]) +
      shape('Mixed', [['{{FOR x IN xs}}{{IF x}}{{END-FOR}}']]) +
      shape('Else', [['{{ELSE}}'], ['{{IF b}}{{ELSE}}{{ELSE}}{{END-IF}}']]),
  ]);

  expect(issues.map((i) => [i.shape, i.rule, i.tag])).toEqual([
    ['Open', 'unclosed-block', '{{IF a}}'],
    ['Close', 'unmatched-tag', '{{END-IF}}'],
    ['Mixed', 'unclosed-block', '{{END-FOR}}'],
    ['Else', 'unmatched-tag', '{{ELSE}}'],
    ['Else', 'unmatched-tag', '{{ELSE}}'],
  ]);
  expect(issues[0]).toMatchObject({
    severity: 'error',
    part: 'ppt/slides/slide1.xml',
    slide: 1,
    message: '{{IF a}} is not closed with {{END-IF}} in this shape',
  });
  expect(issues[2].message).toBe('{{IF x}} is not closed before {{END-FOR}}');
  expect(issues[3].message).toBe('{{ELSE}} without a matching {{IF …}}');
  expect(issues[4].message).toBe('{{ELSE}} after {{ELSE}} in {{IF b}}');
});

test('reports tags split by a line break or missing a delimiter', async () => {
  const issues = await lint([
    shape('A', [['Hello {{user.na'], ['me}} and {{ok}}']]) +
      `<p:sp><p:nvSpPr><p:cNvPr id="4" name="B"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>` +
      '<a:p><a:r><a:t>{{IF</a:t></a:r><a:br/><a:r><a:t>x}}</a:t></a:r></a:p></p:txBody></p:sp>',
  ]);
  expect(issues.map((i) => [i.shape, i.rule, i.message])).toEqual([
    ['A', 'unterminated-tag', '"{{" without "}}"; the text stays as it is'],
    ['A', 'unterminated-tag', '"}}" without "{{"; the text stays as it is'],
    ['B', 'unterminated-tag', '"{{" without "}}"; the text stays as it is'],
    ['B', 'unterminated-tag', '"}}" without "{{"; the text stays as it is'],
  ]);
  expect(issues[0].tag).toBe('Hello {{user.na');
});

test('explains syntax errors caused by AutoCorrect', async () => {
  const issues = await lint([
    shape('Quotes', [["{{IF name == 'Ada’}}x{{END-IF}}"]]) +
      shape('Smart', [['{{IF name == ’Ada’}}x{{END-IF}}']]) +
      shape('Copyright', [['{{FOR c IN items.filter(©=> c.ok)}}{{END-FOR}}']]) +
      shape('Compare', [['{{IF a = 1}}{{END-IF}}']]) +
      shape('And', [['{{IF a AND b}}{{END-IF}}']]),
  ]);
  expect(rules(issues)).toEqual(['syntax', 'syntax', 'syntax', 'syntax', 'syntax']);
  expect(issues[0].hint).toMatch(/Strings need straight quotes/);
  expect(issues[1].message).toMatch(/Unexpected character "’"/);
  expect(issues[1].hint).toMatch(/Strings need straight quotes/);
  expect(issues[2].hint).toMatch(/replaced "\(c\)" with "©"/);
  expect(issues[3].hint).toMatch(/Use == or ===/);
  expect(issues[4].hint).toMatch(/Use && for AND/);
});

test('fixSmartQuotes lints typographic quotes as the renderer parses them', async () => {
  const slide = shape('T', [['{{IF name == ’Ada’}}x{{END-IF}}{{“a” + „b“}}']]);
  const strict = await lint([slide]);
  expect(rules(strict)).toEqual(['syntax', 'syntax']);
  expect(strict[0].hint).toMatch(/or render with the fixSmartQuotes option/);
  expect(await lint([slide], { fixSmartQuotes: true })).toEqual([]);
});

test('only straight quotes delimit strings', () => {
  const scope = { lookup: () => ({ found: false }) };
  expect(new TemplateExpression("'Ada'").evaluate(scope)).toBe('Ada');
  expect(new TemplateExpression('"Ada"').evaluate(scope)).toBe('Ada');
  // Typographic quotes inside a string are just text
  expect(new TemplateExpression("'Ada’s „Haus“'").evaluate(scope)).toBe('Ada’s „Haus“');
  for (const source of ['’Ada’', '‘Ada’', '“Ada”', '„Ada“', '»Ada«']) {
    expect(() => new TemplateExpression(source)).toThrow(/Unexpected character/);
  }
});

test('reports unknown variables and properties against example data', async () => {
  const issues = await lint(
    [
      shape('Text', [
        ['{{usr.name}} {{user.nmae}} {{user.name.lenght}} {{upper(user.name)}}'],
        ['{{FOR o IN user.orders}}{{o.totl}}{{$o.total}}{{$idx}}{{END-FOR}}'],
        ['{{FOR o IN user.name}}{{END-FOR}}{{FOR o IN user}}{{END-FOR}}'],
        ['{{EXEC sum = user.orders.length}}{{sum}}{{ALIAS n = user.name}}{{*n}}'],
        ['{{items.filter(i => i.ok).length}}'],
      ]),
    ],
    {
      data: { user: { name: 'Ada', orders: [{ total: 3 }] }, items: [] },
      additionalJsContext: { upper: (s: string) => s.toUpperCase() },
    },
  );

  expect(issues.map((i) => [i.rule, i.message])).toEqual([
    ['unknown-variable', '"usr" is not in the data (did you mean "user"?)'],
    ['unknown-property', '"nmae" does not exist on "user" (object); did you mean "name"?'],
    ['unknown-property', '"lenght" does not exist on "user.name" (string); did you mean "length"?'],
    ['unknown-property', '"totl" does not exist on "o" (object); did you mean "total"?'],
    ['not-a-list', '"user" is not a list in the example data'],
  ]);
  expect(issues.every((i) => i.severity === 'warning')).toBe(true);
});

const NBSP = String.fromCharCode(0xa0);

test('marks the characters an issue is about', async () => {
  /** The part of the tag the range marks. */
  const marked = (issue: TemplateLintIssue) =>
    issue.tag.slice(issue.range.start, issue.range.end);

  const issues = await lint(
    [
      shape('A', [
        ["{{euro(p.preis) + 'Test`}}"],
        ['{{IF a = 1}}{{END-IF}}'],
        // PowerPoint often stores non-breaking spaces
        [`{{IF${NBSP}a${NBSP}=${NBSP}2}}{{END-IF}}`],
        ['{{IF name == ’Ada’}}{{END-IF}}'],
        ['{{EXEC total = 0; x = )}}'],
        ['{{ user.nmae }}{{usr}}{{FOR o IN user}}{{END-FOR}}'],
        ['{{ALIAS n INS user.mial}}{{*n}}'],
        ['Hello {{user.na'],
        ['{{IF a}}'],
      ]),
    ],
    { data: { user: { name: 'Ada' }, euro: null, p: { preis: 1 }, a: 1, name: '' } },
  );

  expect(issues.map((i) => [i.rule, marked(i)])).toEqual([
    ['syntax', "'Test`"],
    ['syntax', '='],
    ['syntax', '='],
    ['syntax', '’'],
    ['syntax', ')'],
    ['unknown-property', 'nmae'],
    ['unknown-variable', 'usr'],
    ['not-a-list', 'user'],
    // Through an alias the code is not in the tag: the whole tag is marked
    ['unknown-property', '{{*n}}'],
    ['unterminated-tag', '{{'],
    ['unclosed-block', '{{IF a}}'],
  ]);
});

test('reports unknown names against a JSON Schema', async () => {
  const schema = {
    type: 'object',
    properties: {
      user: { $ref: '#/$defs/user' },
      tags: { type: 'array', items: { type: 'string' } },
      meta: { type: 'object' },
      maybe: { anyOf: [{ type: 'null' }, { type: 'object', properties: { x: { type: 'number' } } }] },
    },
    $defs: {
      user: {
        type: 'object',
        properties: { name: { type: 'string' } },
      },
    },
  };
  const issues = await lint(
    [
      shape('Text', [
        ['{{user.name.trim()}} {{user.mail}} {{meta.anything}} {{maybe.x}} {{maybe.y}}'],
        ['{{FOR t IN tags}}{{t.toUpperCase()}}{{t.foo}}{{END-FOR}}{{FOR u IN user}}{{END-FOR}}'],
        ['{{missing}}'],
      ]),
    ],
    { schema },
  );
  expect(issues.map((i) => [i.rule, i.message])).toEqual([
    ['unknown-property', '"mail" does not exist on "user" (object)'],
    ['unknown-property', '"y" does not exist on "maybe" (value)'],
    ['unknown-property', '"foo" does not exist on "t" (string)'],
    ['not-a-list', '"user" is not a list in the example data'],
    ['unknown-variable', '"missing" is not in the data'],
  ]);
});

test('without data or schema, names are not checked', async () => {
  expect(await lint([shape('Text', [['{{anything.at.all}}']])])).toEqual([]);
});

test('reports tags in places renderTemplate does not render', async () => {
  const issues = await lint(
    [
      shape('Placeholder', [['{{IMAGE logo}}']], false) +
        shape('Box', [['{{IMAGE logo}}']]) +
        table('Table', [
          ['{{IMAGE logo}}', ''],
          ['{{IF a}}x', ''],
          ['{{ELSE}}', '{{END-IF}}'],
        ]),
    ],
    undefined,
    {
      'ppt/slides/_rels/slide1.xml.rels':
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId2" Type="notesSlide" Target="../notesSlides/notesSlide7.xml"/>' +
        '</Relationships>',
      'ppt/notesSlides/notesSlide7.xml': partXml('notes', shape('Notes', [['{{note}}']])),
      'ppt/slideLayouts/slideLayout2.xml': partXml(
        'sldLayout',
        shape('Footer', [['{{a}} {{b}} {{c}} {{d}}']]),
      ),
    },
  );

  expect(issues.map((i) => [i.part, i.shape, i.rule, i.severity])).toEqual([
    ['ppt/slides/slide1.xml', 'Placeholder', 'unsupported-location', 'error'],
    ['ppt/slides/slide1.xml', 'Table', 'unsupported-location', 'error'],
    ['ppt/slides/slide1.xml', 'Table', 'unsupported-location', 'error'],
    ['ppt/notesSlides/notesSlide7.xml', undefined, 'unsupported-location', 'warning'],
    ['ppt/slideLayouts/slideLayout2.xml', undefined, 'unsupported-location', 'warning'],
  ]);
  expect(issues[1].message).toBe('IMAGE is not supported inside table cells');
  expect(issues[2].message).toBe('{{ELSE}} cannot span table rows; use it inside a cell');
  expect(issues[3]).toMatchObject({ slide: 1, message: '1 tag(s) in speaker notes are not rendered: {{note}}' });
  expect(issues[4].message).toBe(
    '4 tag(s) in slide layout 2 are not rendered: {{a}}, {{b}}, {{c}}, …',
  );
});

const writeTemplate =
  (paragraphs: string[][]) =>
  (element: XmlElement): void =>
    modify.setMultiText(
      paragraphs.map((runs) => ({
        paragraph: {},
        textRuns: runs.map((text) => ({ text })),
      })),
    )(element);

const lintedDeck = `${__dirname}/pptx-output/template-lint.test.pptx`;

beforeAll(async () => {
  const automizer = new Automizer({
    templateDir: `${__dirname}/pptx-templates`,
    outputDir: `${__dirname}/pptx-output`,
  });
  await automizer
    .loadRoot('RootTemplate.pptx')
    .load('TextReplace.pptx')
    .addSlide('TextReplace.pptx', 1, (slide) => {
      slide.modifyElement(
        'setText',
        writeTemplate([['{{IF us', 'er.vip}}'], ['{{user.nmae}}']]),
      );
    })
    .write('template-lint.test.pptx');
});

test('lints a deck written by PowerPoint/automizer from a file path', async () => {
  const result = await lintTemplate(lintedDeck, { data: { user: { name: 'Ada', vip: true } } });
  // The template's own {{replace}}/{{by}}/{{replacement}} tags count, too
  expect(result.issues.map((i) => [i.slide, i.shape, i.rule, i.tag])).toEqual([
    [2, 'replaceText', 'unknown-variable', '{{replace}}'],
    [2, 'replaceText', 'unknown-variable', '{{by}}'],
    [2, 'replaceText', 'unknown-variable', '{{replacement}}'],
    [2, 'setText', 'unknown-property', '{{user.nmae}}'],
    [2, 'setText', 'unclosed-block', '{{IF user.vip}}'],
  ]);
  expect(result.issues[3].message).toBe(
    '"nmae" does not exist on "user" (object); did you mean "name"?',
  );
  expect(result).toMatchObject({ errorCount: 1, warningCount: 4 });
});

describe('CLI', () => {
  const run = async (...args: string[]) => {
    let stdout = '';
    let stderr = '';
    const code = await runCli(args, {
      stdout: (t) => (stdout += t),
      stderr: (t) => (stderr += t),
    });
    return { code, stdout, stderr };
  };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-lint-'));
  const dataFile = path.join(dir, 'data.json');
  fs.writeFileSync(dataFile, JSON.stringify({ user: { name: 'Ada', vip: true } }));

  test('colors only when asked: red error, red underline, yellow warning', async () => {
    const args = ['lint', lintedDeck, '--data', dataFile];
    expect((await run(...args)).stdout).not.toContain('\u001b[');
    const colored = (await run(...args, '--color')).stdout;
    expect(colored).toContain('\u001b[1m\u001b[31merror\u001b[39m\u001b[22m');
    expect(colored).toContain('\u001b[1m\u001b[33mwarning\u001b[39m\u001b[22m');
    expect(colored).toContain('\u001b[33m~~~~\u001b[39m'); // {{user.nmae}} warning
    expect(colored).toContain('\u001b[31m~');
    expect((await run(...args, '--color', '--no-color')).stdout).not.toContain('\u001b[');
  });

  test('prints issues and fails on errors', async () => {
    const { code, stdout } = await run('lint', lintedDeck, '--data', dataFile);
    expect(code).toBe(1);
    expect(stdout).toContain('slide 2 › "setText"');
    expect(stdout).toContain('error  {{IF user.vip}} is not closed with {{END-IF}} in this shape  (unclosed-block)');
    expect(stdout).toContain('5 problem(s): 1 error(s), 4 warning(s)');
    // The issue's range is underlined below the tag
    expect(stdout).toContain(
      '    tag: {{user.nmae}}\n' +
      '                ~~~~\n',
    );
  });

  test('prints JSON', async () => {
    const { code, stdout } = await run('lint', '--json', lintedDeck);
    expect(code).toBe(1);
    expect(JSON.parse(stdout)[lintedDeck].map((i: TemplateLintIssue) => i.rule)).toEqual([
      'unclosed-block',
    ]);
  });

  test('passes a clean deck, --max-warnings fails on warnings', async () => {
    const clean = path.join(dir, 'clean.pptx');
    fs.writeFileSync(clean, await deck([shape('T', [['{{user.mail}}']])]));
    expect(await run('lint', clean)).toMatchObject({
      code: 0,
      stdout: 'No problems found in 1 file(s)\n',
    });
    expect((await run('lint', clean, '--data', dataFile)).code).toBe(0);
    expect((await run('lint', clean, '--data', dataFile, '--max-warnings', '0')).code).toBe(1);

    const smart = path.join(dir, 'smart.pptx');
    fs.writeFileSync(smart, await deck([shape('T', [['{{’Ada’}}']])]));
    expect((await run('lint', smart)).code).toBe(1);
    expect((await run('lint', smart, '--fix-smart-quotes')).code).toBe(0);
  });

  test('rejects invalid calls with exit code 2', async () => {
    // Usage goes to stderr when the call is invalid, to stdout for --help
    expect(await run()).toMatchObject({ code: 2, stdout: '' });
    expect((await run()).stderr).toContain('No command given.');
    expect(await run('--help')).toMatchObject({ code: 0, stderr: '' });
    expect((await run('lint')).code).toBe(2);
    expect((await run('lint', '--bogus')).stderr).toContain('Unknown option --bogus');
    expect((await run('nope')).stderr).toContain('Unknown command "nope"');
    expect((await run('lint', path.join(dir, 'missing.pptx'))).code).toBe(2);
    expect((await run('lint', '--help')).code).toBe(0);
  });
});

test('broken relationships and schema refs do not stop the linter', async () => {
  const noTarget =
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="x"/></Relationships>';
  expect(
    await lint([shape('T', [['{{a.b}}']])], undefined, {
      'ppt/slides/_rels/slide1.xml.rels': noTarget,
    }),
  ).toEqual([]);

  // '%E0%A4%A' is not a valid %-escape; such a ref decides nothing
  expect(
    await lint([shape('T', [['{{a.b}}']])], {
      schema: { type: 'object', properties: { a: { $ref: '#/%E0%A4%A' } } },
    }),
  ).toEqual([]);
});

describe('CLI input from files', () => {
  const run = async (...args: string[]) => {
    let stdout = '';
    let stderr = '';
    const code = await runCli(args, {
      stdout: (t) => (stdout += t),
      stderr: (t) => (stderr += t),
    });
    return { code, stdout, stderr };
  };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-lint-input-'));

  test('control characters from the .pptx do not reach the terminal', async () => {
    const evil = path.join(dir, 'evil.pptx');
    // ESC [2J would clear the screen; U+009B is the one-character CSI
    fs.writeFileSync(
      evil,
      await deck([shape('Evil&#27;[2J', [['{{IF a\u009b2K}}']])]),
    );
    const { code, stdout } = await run('lint', evil);
    expect(code).toBe(1);
    expect(stdout).not.toContain('\u001b');
    expect(stdout).not.toContain('\u009b');
    expect(stdout).toContain('slide 1 › "Evil [2J"');
    expect(stdout).toContain('tag: {{IF a 2K}}');
  });

  test('--data and --schema need a JSON object', async () => {
    const file = path.join(dir, 'null.json');
    fs.writeFileSync(file, 'null');
    for (const option of ['--data', '--schema']) {
      const { code, stderr } = await run('lint', 'deck.pptx', option, file);
      expect(code).toBe(2);
      expect(stderr).toContain('expected a JSON object');
    }
  });
});
