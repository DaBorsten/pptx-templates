import { escapeRegExp } from '../src/extensions/regexp-escape';

const specials = '\\^$.*+?()[]{}|-/';

test('an escaped string matches itself and nothing else', () => {
  const values = [specials, '{{', '}}', '[[', ']]', '<%', '%>', '${', 'a.b'];
  for (const value of values) {
    const pattern = new RegExp(`^${escapeRegExp(value)}$`);
    expect(pattern.test(value)).toBe(true);
  }
  expect(new RegExp(escapeRegExp('a.b')).test('axb')).toBe(false);
  expect(new RegExp(escapeRegExp('a*')).test('xa*x')).toBe(true);
  expect(new RegExp(`^${escapeRegExp('a*')}$`).test('aaa')).toBe(false);
});

test('inside a character class every character stays literal', () => {
  // as in mergeRunsSplitByTags and TextReplaceHelper.getRegExp
  const excluded = new RegExp(`^[^${escapeRegExp(specials)}]+$`);
  for (const char of specials) {
    expect(excluded.test(char)).toBe(false);
  }
  // `-` must not form a range between its neighbours
  const dash = new RegExp(`^[${escapeRegExp('a-z')}]$`);
  expect(dash.test('-')).toBe(true);
  expect(dash.test('m')).toBe(false);
});

test('custom delimiters still find tags', () => {
  const open = escapeRegExp('[[');
  const close = escapeRegExp(']]');
  const pattern = new RegExp(`${open}[^${open}${close}]+${close}`, 'g');
  expect('Hallo [[name]], [[ort]]!'.match(pattern)).toEqual([
    '[[name]]',
    '[[ort]]',
  ]);
});
