import { describe, expect, it } from 'vitest';
import { compileRegex, RegexError } from '../src/engine/regex';
import { defs } from '../src/engine/functions/text';
import { Err, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx();
const call = (name: string, ...args: Value[]): Value => defs.find((d) => d.name === name)!.fn(args, ctx);
const m = (p: string, s: string, i = false): (string | undefined)[] | null => {
  const r = compileRegex(p, i).exec(s);
  return r ? [r.whole, ...r.groups] : null;
};
const bad = (p: string): void => { expect(() => compileRegex(p)).toThrow(RegexError); };

describe('regex engine: basics', () => {
  it('literals, classes, quantifiers', () => {
    expect(m('(\\d+)-(\\d+)', 'x 12-345 y')).toEqual(['12-345', '12', '345']);
    expect(m('colou?r', 'my color')).toEqual(['color']);
    expect(m('a{2,3}', 'aaaa')).toEqual(['aaa']);
    expect(m('a{2,3}?', 'aaaa')).toEqual(['aa']);
    expect(m('a{,2}', 'aaa')).toEqual(['aa']);
    expect(m('a{3', 'a{3')).toEqual(['a{3']);
    expect(m('x{}', 'x{}')).toEqual(['x{}']);
    expect(m('[a-c-e]+', 'ab-ec')).toEqual(['ab-ec']);
    expect(m('[]a]+', 'a]a')).toEqual(['a]a']);
    expect(m('[^]a]+', 'xy]')).toEqual(['xy']);
    expect(m('[\\d.]+', 'v1.25b')).toEqual(['1.25']);
    expect(m('.+', 'a\nb')).toEqual(['a']);
    expect(m('\\N+', 'ab\nc')).toEqual(['ab']);
    expect(m('\\x41\\x{42}\\o{103}\\103', 'ABCC')).toEqual(['ABCC']);
    expect(m('\\x{1F600}', 'hi 😀')).toEqual(['😀']);
    expect(m('^.$', '😀')).toEqual(['😀']);
    expect(m('\\cA', '\u0001')).toEqual(['\u0001']);
  });
  it('lazy, greedy and alternation order', () => {
    expect(m('<.+>', '<a><b>')).toEqual(['<a><b>']);
    expect(m('<.+?>', '<a><b>')).toEqual(['<a>']);
    expect(m('(?:a|ab)(?:c|bcd)(d*)', 'abcd')).toEqual(['abcd', '']);
    expect(m('(?U)a+', 'aaa')).toEqual(['a']);
    expect(m('(?U)a+?', 'aaa')).toEqual(['aaa']);
  });
  it('anchors', () => {
    expect(m('a$', 'a\n')).toEqual(['a']);
    expect(m('a\\z', 'a\n')).toBeNull();
    expect(m('a\\Z', 'a\n')).toEqual(['a']);
    expect(m('a\\Z', 'a\nb')).toBeNull();
    expect(m('\\Ab', 'ab')).toBeNull();
    expect(m('(?m)^b$', 'a\nb\nc')).toEqual(['b']);
    expect(m('^b', 'a\nb')).toBeNull();
    expect(m('\\bcat\\b', 'concat cat!')).toEqual(['cat']);
    expect(m('\\Bcat', 'concat')).toEqual(['cat']);
    expect(compileRegex('(?m)^').matchAll('a\nb\n').length).toBe(2);
  });
});

describe('regex engine: PCRE2 features', () => {
  it('atomic groups', () => {
    expect(m('a(?>bc|b)c', 'abcc')).toEqual(['abcc']);
    expect(m('a(?>bc|b)c', 'abc')).toBeNull();
    expect(m('(?>a+)b', 'aaab')).toEqual(['aaab']);
    expect(m('(?>(\\d+))\\d', '123')).toBeNull();
    expect(m('(?>x(y))z', 'xyz')).toEqual(['xyz', 'y']);
  });
  it('possessive quantifiers', () => {
    expect(m('a++b', 'aaab')).toEqual(['aaab']);
    expect(m('a++a', 'aaa')).toBeNull();
    expect(m('a*+a', 'aaa')).toBeNull();
    expect(m('a?+a', 'a')).toBeNull();
    expect(m('x{2,3}+x', 'xxxx')).toEqual(['xxxx']);
    expect(m('x{2,3}+x', 'xxx')).toBeNull();
    expect(m('"[^"]*+"', '"abc"')).toEqual(['"abc"']);
  });
  it('inline modifiers', () => {
    expect(m('(?i)hello', 'HeLLo')).toEqual(['HeLLo']);
    expect(m('a(?i)b', 'aB')).toEqual(['aB']);
    expect(m('a(?i)b', 'AB')).toBeNull();
    expect(m('a(?i:b)c', 'aBc')).toEqual(['aBc']);
    expect(m('a(?i:b)c', 'aBC')).toBeNull();
    expect(m('(?i)a(?-i)b', 'Ab')).toEqual(['Ab']);
    expect(m('(?i)a(?-i)b', 'AB')).toBeNull();
    expect(m('(a(?i)b|c)', 'C')).toEqual(['C', 'C']);
    expect(m('(?s).+', 'a\nb')).toEqual(['a\nb']);
    expect(m('(?s:.)\\n?', '\n')).toEqual(['\n']);
    expect(m('(?x) a b # comment\n c', 'abc')).toEqual(['abc']);
    expect(m('(?x)a\\ b', 'a b')).toEqual(['a b']);
    expect(m('(?x)[a b]+', 'b a')).toEqual(['b a']);
    expect(m('(?xx)[a b]+', 'b a')).toEqual(['b']);
    expect(m('(?x: a b )c', 'abc')).toEqual(['abc']);
    expect(m('(?n)(a)(?<x>b)', 'ab')).toEqual(['ab', 'b']);
    expect(m('(?i)(?^)a', 'A')).toBeNull();
  });
  it('named groups and back-references', () => {
    expect(m('(?<y>\\d{4})-(?P<m>\\d\\d)-(?\'d\'\\d\\d)', '2024-05-09')).toEqual(['2024-05-09', '2024', '05', '09']);
    expect(m('(?<n>a)\\k<n>\\k{n}\\k\'n\'(?P=n)\\g{n}', 'aaaaaa')).toEqual(['aaaaaa', 'a']);
    expect(m('(\\w)\\1', 'abccd')).toEqual(['cc', 'c']);
    expect(m('(\\w)\\g1\\g{1}\\g{-1}', 'xxxx')).toEqual(['xxxx', 'x']);
    expect(m('(?i)(a)\\1', 'aA')).toEqual(['aA', 'a']);
    // An unset group never matches (unlike JavaScript, where it matches empty).
    expect(m('(a)|b\\1', 'b')).toBeNull();
    expect(m('(?J)(?<n>a)|(?<n>b)', 'b')).toEqual(['b', undefined, 'b']);
    bad('(?<n>a)(?<n>b)');
    bad('\\k<missing>');
    bad('(a)\\2');
  });
  it('\\Q...\\E, POSIX classes, \\h \\v \\R and properties', () => {
    expect(m('\\Qa.b*\\E+', 'a.b**')).toEqual(['a.b**']);
    expect(m('\\Q(x)', 'f(x)')).toEqual(['(x)']);
    expect(m('[\\Q]^\\E]+', '^]^')).toEqual(['^]^']);
    expect(m('[[:alpha:]]+', '12abc3')).toEqual(['abc']);
    expect(m('[[:^digit:]]+', '12abc3')).toEqual(['abc']);
    expect(m('[[:upper:][:digit:]]+', 'xAB1y')).toEqual(['AB1']);
    expect(m('[[:punct:]]+', 'a!?b')).toEqual(['!?']);
    expect(m('[[:xdigit:]]+', 'zz0fAg')).toEqual(['0fA']);
    expect(m('[[:space:]]+', 'a \t\nb')).toEqual([' \t\n']);
    expect(m('\\h+', 'a \t b')).toEqual([' \t ']);
    expect(m('\\v+', 'a\n\rb')).toEqual(['\n\r']);
    expect(m('a\\Rb', 'a\r\nb')).toEqual(['a\r\nb']);
    expect(m('a\\R{2}b', 'a\r\nb')).toBeNull();
    expect(m('\\p{Lu}+', 'abCDÉé')).toEqual(['CDÉ']);
    expect(m('\\P{L}+', 'ab12c')).toEqual(['12']);
    expect(m('\\p{Greek}+', 'abc αβγ')).toEqual(['αβγ']);
    expect(m('\\p{Arabic}+', 'x مرحبا')).toEqual(['مرحبا']);
    expect(m('\\pN+', 'a42')).toEqual(['42']);
    expect(m('\\p{^L}', 'a1')).toEqual(['1']);
    bad('[[:nope:]]');
    bad('\\p{NotAProperty}');
  });
  it('\\d and \\w are ASCII unless (*UCP)', () => {
    expect(m('\\w+', 'café')).toEqual(['caf']);
    expect(m('(*UCP)\\w+', 'café')).toEqual(['café']);
    expect(m('\\d', '٣')).toBeNull();
    expect(m('(*UCP)\\d', '٣')).toEqual(['٣']);
  });
  it('case-insensitive matching folds Unicode', () => {
    expect(m('é', 'É', true)).toEqual(['É']);
    expect(m('[à-ÿ]+', 'ÀÉ', true)).toEqual(['ÀÉ']);
    expect(m('σ', 'Σ', true)).toEqual(['Σ']);
    expect(m('k', 'K', true)).toEqual(['K']);
    expect(m('[^a]', 'A', true)).toBeNull();
  });
  it('lookaround, \\K, conditionals, branch reset, recursion', () => {
    expect(m('(?<=\\$)\\d+', 'cost $45')).toEqual(['45']);
    expect(m('(?<!x)y', 'xy zy')).toEqual(['y']);
    expect(m('(?<=ab|c)d', 'cd')).toEqual(['d']);
    expect(m('\\w+(?=;)', 'ab cd;')).toEqual(['cd']);
    expect(m('\\d+(?!px)\\b', '12px 30em 7')).toEqual(['7']);
    expect(m('(?=(\\w+))a', 'abc')).toEqual(['a', 'abc']);
    bad('(?<=a+)b');
    expect(m('foo\\Kbar', 'foobar')).toEqual(['bar']);
    expect(m('(a)?(?(1)b|c)', 'c')).toEqual(['c', undefined]);
    expect(m('(a)?(?(1)b|c)', 'ab')).toEqual(['ab', 'a']);
    expect(m('(?<q>")?\\w+(?(<q>)")', '"hi"')).toEqual(['"hi"', '"']);
    expect(m('(?(?=\\d)\\d{3}|[a-z]{2})', 'ab')).toEqual(['ab']);
    expect(m('(?(DEFINE)(?<byte>25[0-5]|2[0-4]\\d|1?\\d?\\d))^(?&byte)(?:\\.(?&byte)){3}$', '192.168.1.255')).not.toBeNull();
    expect(m('(?(DEFINE)(?<byte>25[0-5]|2[0-4]\\d|1?\\d?\\d))^(?&byte)(?:\\.(?&byte)){3}$', '192.168.1.256')).toBeNull();
    expect(m('(?|(a)|(b))\\1', 'bb')).toEqual(['bb', 'b']);
    expect(m('\\((?:[^()]|(?R))*\\)', 'x(a(b)c)y')).toEqual(['(a(b)c)']);
    expect(m('^(\\w)(?:(?1)|\\w?)*?\\1$', 'abca')).not.toBeNull();
    expect(m('^((.)(?:(?1)|.?)\\2)$', 'racecar')).not.toBeNull();
    expect(m('^((.)(?:(?1)|.?)\\2)$', 'racecat')).toBeNull();
    expect(m('(?:a(*FAIL)|b)', 'ab')).toEqual(['b']);
    expect(m('a+(*COMMIT)b', 'aaac aab')).toBeNull();
    expect(m('a(*ACCEPT)b', 'ac')).toEqual(['a']);
  });
  it('comments and errors', () => {
    expect(m('a(?#note)b', 'ab')).toEqual(['ab']);
    bad('(');
    bad(')');
    bad('[');
    bad('*a');
    bad('a**');
    bad('a{3,1}');
    bad('\\');
    bad('\\i');
    bad('(?<=\\1)a');
  });
  it('empty matches and catastrophic patterns', () => {
    expect(compileRegex('\\d*').matchAll('a12b').map((x) => x.whole)).toEqual(['', '12', '', '']);
    expect(compileRegex('').matchAll('abc')).toHaveLength(4);
    expect(m('(a|)+b', 'aab')).toEqual(['aab', '']);
    expect(m('(a*)*b', 'a'.repeat(40) + 'c')).toBeNull();
    expect(() => compileRegex('(a*)*$x?(?=c)').exec('a'.repeat(30) + 'b')).not.toThrow(TypeError);
  });
});

describe('REGEXTEST / REGEXEXTRACT / REGEXREPLACE', () => {
  it('REGEXTEST', () => {
    expect(call('REGEXTEST', 'alpha-42', '\\d+')).toBe(true);
    expect(call('REGEXTEST', 'ABC', 'abc')).toBe(false);
    expect(call('REGEXTEST', 'ABC', 'abc', 1)).toBe(true);
    expect(call('REGEXTEST', 'ABC', 'abc', 2)).toEqual(Err.value);
    expect(call('REGEXTEST', 'x', '(')).toEqual(Err.value);
    expect(call('REGEXTEST', 123, '^\\d{3}$')).toBe(true);
    expect(call('REGEXTEST', 'aaab', 'a++b')).toBe(true);
  });
  it('REGEXEXTRACT modes', () => {
    expect(call('REGEXEXTRACT', 'Call 555-1234 or 555-9876', '\\d{3}-\\d{4}')).toBe('555-1234');
    expect(call('REGEXEXTRACT', 'Call 555-1234 or 555-9876', '\\d{3}-\\d{4}', 1)).toEqual([['555-1234'], ['555-9876']]);
    expect(call('REGEXEXTRACT', 'Jane Doe', '(\\w+) (\\w+)', 2)).toEqual([['Jane', 'Doe']]);
    expect(call('REGEXEXTRACT', 'ab', '(a)(x)?', 2)).toEqual([['a', '']]);
    expect(call('REGEXEXTRACT', 'none', '\\d')).toEqual(Err.na);
    expect(call('REGEXEXTRACT', 'none', '\\d', 1)).toEqual(Err.na);
    expect(call('REGEXEXTRACT', 'a', 'a', 3)).toEqual(Err.value);
    expect(call('REGEXEXTRACT', 'HELLO', 'hel+', 0, 1)).toBe('HELL');
    expect(call('REGEXEXTRACT', 'price: $12.50', '(?<=\\$)[\\d.]+')).toBe('12.50');
    expect(call('REGEXEXTRACT', 'a😀b😀', '😀', 1)).toEqual([['😀'], ['😀']]);
  });
  it('REGEXREPLACE occurrences and replacements', () => {
    expect(call('REGEXREPLACE', 'a1b2c3', '\\d', '#')).toBe('a#b#c#');
    expect(call('REGEXREPLACE', 'a1b2c3', '\\d', '#', 2)).toBe('a1b#c3');
    expect(call('REGEXREPLACE', 'a1b2c3', '\\d', '#', -1)).toBe('a1b2c#');
    expect(call('REGEXREPLACE', 'a1b2c3', '\\d', '#', 4)).toBe('a1b2c3');
    expect(call('REGEXREPLACE', 'John Smith', '(\\w+) (\\w+)', '$2, $1')).toBe('Smith, John');
    expect(call('REGEXREPLACE', 'John Smith', '(?<f>\\w+) (?<l>\\w+)', '${l} ${f}')).toBe('Smith John');
    expect(call('REGEXREPLACE', 'John Smith', '(?<f>\\w+) (?<l>\\w+)', '$l-$f')).toBe('Smith-John');
    expect(call('REGEXREPLACE', 'ab', '(a)', '${1}0')).toBe('a0b');
    expect(call('REGEXREPLACE', 'ab', '(a)', '$10')).toBe('a0b');
    expect(call('REGEXREPLACE', 'ab', 'a', '$$')).toBe('$b');
    expect(call('REGEXREPLACE', 'ab', 'a', '[$0]')).toBe('[a]b');
    expect(call('REGEXREPLACE', 'ab', '(a)', '$2')).toEqual(Err.value);
    expect(call('REGEXREPLACE', 'abc', '', '-')).toBe('-a-b-c-');
    expect(call('REGEXREPLACE', 'Hello World', 'o', '0', 0, 1)).toBe('Hell0 W0rld');
    expect(call('REGEXREPLACE', 'AbA', 'a', 'x', 0, 1)).toBe('xbx');
    expect(call('REGEXREPLACE', '  trim  me ', '^\\s+|\\s+$', '')).toBe('trim  me');
    expect(call('REGEXREPLACE', 'x😀y', '.', '.')).toBe('...');
  });
});

describe('regex engine: verbs and full case folding', () => {
  it('(*THEN) moves to the next alternative of the innermost alternation', () => {
    expect(m('(?:a(*THEN)b|ac)', 'ac')).toEqual(['ac']);
    expect(m('(?:a(*PRUNE)b|ac)', 'ac')).toBeNull();
    // A group without "|" is not an alternation of its own: (*THEN) reaches the outer one.
    expect(m('(?:(?:a(*THEN)b)|ac)', 'ac')).toEqual(['ac']);
    // In the last alternative the alternation fails as a whole, and matching backtracks before it.
    expect(m('x?(?:a|b(*THEN)c)', 'xbd')).toBeNull();
    expect(m('^.*?(?:a(*THEN)b|c)', 'aac')).toEqual(['aac']);
    // No enclosing alternation: like (*PRUNE), the next start position is tried.
    expect(m('a+(*THEN)b', 'aac aab')).toEqual(['aab']);
    expect(m('(*THEN)x|y', 'y')).toEqual(['y']);
  });
  it('named verbs', () => {
    expect(m('(*MARK:A)a(*SKIP:A)b|ac', 'ac')).toBeNull();
    expect(m('a(*:M)b(*SKIP:M)c|.', 'abd')).toEqual(['b']);
    expect(m('a(*SKIP:nope)x|ab', 'ab')).toEqual(['ab']);
    expect(m('a(*PRUNE:N)b|ac', 'ac')).toBeNull();
    expect(m('(?:a(*THEN:T)b|ac)', 'ac')).toEqual(['ac']);
    expect(m('aaa(*SKIP)x|a', 'aaab')).toBeNull();
    expect(m('aaa(*SKIP)x|b', 'aaab')).toEqual(['b']);
    bad('(*MARK)');
    bad('(*:)');
    bad('(*FAIL:x)');
  });
  it('verbs inside atomic groups and assertions', () => {
    // Backtracked into while the atomic group is matching: normal effect.
    expect(m('(?>a(*COMMIT)b)|ac', 'ac')).toBeNull();
    // Once the group has matched, its verbs are discarded.
    expect(m('(?>a(*COMMIT))b|ac', 'ac')).toEqual(['ac']);
    // Negative assertion: (*COMMIT) backtracked into makes it true without trying other branches.
    expect(m('(?!a(*COMMIT)b|ac)a.', 'ac')).toEqual(['ac']);
    // Conditional positive assertion: (*PRUNE) makes the condition false.
    expect(m('(?(?=a(*PRUNE)b)x|ac)', 'ac')).toEqual(['ac']);
    // (*THEN) does not escape an assertion.
    expect(m('(?:(?=a(*THEN)b)|a)c', 'ac')).toEqual(['ac']);
    // (*ACCEPT) inside an atomic group ends the whole match.
    expect(m('(?>a(*ACCEPT)x)y', 'aq')).toEqual(['a']);
    // Inside a subroutine call (*COMMIT) only fails the call.
    expect(m('(?(DEFINE)(?<x>a(*COMMIT)b))(?&x)|ac', 'ac')).toEqual(['ac', undefined]);
    expect(m('(?<x>a(*COMMIT)b){0}(?&x)|ac', 'ac')).toEqual(['ac', undefined]);
  });
  it('full Unicode case folding', () => {
    expect(m('straße', 'STRASSE', true)).toEqual(['STRASSE']);
    expect(m('strasse', 'Straße', true)).toEqual(['Straße']);
    expect(m('ß', 'ss', true)).toEqual(['ss']);
    expect(m('ẞ', 'ß', true)).toEqual(['ß']);
    expect(m('ﬁle', 'FILE', true)).toEqual(['FILE']);
    expect(m('file', 'ﬁle', true)).toEqual(['ﬁle']);
    expect(m('(?i)(ss)\\1', 'ssß')).toEqual(['ssß', 'ss']);
    expect(m('ß', 'ss')).toBeNull();
    expect(m('ı', 'i', true)).toBeNull();
    expect(call('REGEXTEST', 'GROSSE', 'große', 1)).toBe(true);
    expect(call('REGEXREPLACE', 'Die Straße', 'STRASSE', 'Weg', 0, 1)).toBe('Die Weg');
  });
});
