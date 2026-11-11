// Regular expressions with PCRE2-style syntax for REGEXTEST, REGEXEXTRACT and REGEXREPLACE.
//
// JavaScript's RegExp differs from the PCRE2 dialect people write in spreadsheets in ways that change
// results (atomic groups, possessive quantifiers, unset back-references failing instead of matching
// empty, \Z, inline modifiers anywhere, branch reset, recursion, \K...). So this file has its own
// parser and a backtracking virtual machine with an explicit stack, written from the PCRE2 pattern
// syntax documentation. Matching works on Unicode code points. Defaults follow PCRE2 in UTF mode:
// \d \w \s \b and POSIX classes are ASCII unless the pattern starts with (*UCP); "." and \N exclude
// only LF; $ also matches before a final LF. Backtracking verbs ((*THEN), (*SKIP:name) and the rest)
// follow PCRE2's rules, including inside atomic groups, assertions and subroutine calls.
// Case-insensitive matching uses Unicode case folding with ASCII and non-ASCII forms mixing freely
// (Kelvin sign matches k). Going one step past PCRE2, literal text and back-references also use full
// folding, so "ß" matches "SS" and "ﬁ" matches "fi"; single characters in classes fold one-to-one.

export class RegexError extends Error {}
class StepLimit extends Error {}
/** Carries a verb's effect out of a positive assertion. */
class VerbExit extends Error { constructor(readonly code: number) { super('verb'); } }

const MAX_STEPS = 20_000_000;
const MAX_PROGRAM = 400_000;
const MAX_CALL_DEPTH = 2_000;

// ---------- character helpers ----------

const foldCache = new Map<number, number>();
/** Simple case fold: maps every case variant of a letter to one code point. */
const fold = (c: number): number => {
  if (c < 0x41) return c;
  if (c < 0x80) return c >= 0x41 && c <= 0x5a ? c + 32 : c;
  let f = foldCache.get(c);
  if (f !== undefined) return f;
  const s = String.fromCodePoint(c);
  const up = s.toUpperCase();
  const lo = [...up].length === 1 ? up.toLowerCase() : s.toLowerCase();
  f = [...lo].length === 1 ? lo.codePointAt(0)! : c;
  foldCache.set(c, f);
  return f;
};
const fullCache = new Map<number, number[]>();
/**
 * Full case folding: what a character becomes when case is ignored, possibly several characters
 * ("ß" and "ẞ" become "ss", "ﬁ" becomes "fi", "İ" becomes "i̇"). Dotless ı keeps its own form.
 */
export const fullFold = (c: number): number[] => {
  if (c < 0x80) return [c >= 0x41 && c <= 0x5a ? c + 32 : c];
  let f = fullCache.get(c);
  if (f) return f;
  if (c === 0x131) f = [c];
  else {
    let t = String.fromCodePoint(c);
    for (let k = 0; k < 2; k++) t = t.toUpperCase().toLowerCase();
    f = [...t].map((x) => x.codePointAt(0)!);
  }
  fullCache.set(c, f);
  return f;
};

const variants = (c: number): number[] => {
  const s = String.fromCodePoint(c);
  const out = new Set<number>([c, fold(c)]);
  for (const t of [s.toLowerCase(), s.toUpperCase()]) if ([...t].length === 1) out.add(t.codePointAt(0)!);
  return [...out];
};

type Test = (c: number) => boolean;
const isDigitA: Test = (c) => c >= 48 && c <= 57;
const isWordA: Test = (c) => (c >= 48 && c <= 57) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 95;
const isSpaceA: Test = (c) => c === 32 || (c >= 9 && c <= 13);
const isH: Test = (c) => c === 9 || c === 32 || c === 0xa0 || c === 0x1680 || c === 0x180e || (c >= 0x2000 && c <= 0x200a) || c === 0x202f || c === 0x205f || c === 0x3000;
const isV: Test = (c) => (c >= 10 && c <= 13) || c === 0x85 || c === 0x2028 || c === 0x2029;
const uTest = (re: RegExp): Test => (c) => re.test(String.fromCodePoint(c));
const isDigitU = uTest(/^\p{Nd}$/u);
const isWordU = uTest(/^[\p{L}\p{N}\p{Mn}\p{Pc}]$/u);
const isSpaceU: Test = (c) => isH(c) || isV(c) || /^\p{Z}$/u.test(String.fromCodePoint(c));

const POSIX: Record<string, Test> = {
  alnum: (c) => isDigitA(c) || (c >= 65 && c <= 90) || (c >= 97 && c <= 122),
  alpha: (c) => (c >= 65 && c <= 90) || (c >= 97 && c <= 122),
  ascii: (c) => c < 128,
  blank: (c) => c === 32 || c === 9,
  cntrl: (c) => c < 32 || c === 127,
  digit: isDigitA,
  graph: (c) => c > 32 && c < 127,
  lower: (c) => c >= 97 && c <= 122,
  print: (c) => c >= 32 && c < 127,
  punct: (c) => (c >= 33 && c <= 47) || (c >= 58 && c <= 64) || (c >= 91 && c <= 96) || (c >= 123 && c <= 126),
  space: isSpaceA,
  upper: (c) => c >= 65 && c <= 90,
  word: isWordA,
  xdigit: (c) => isDigitA(c) || (c >= 65 && c <= 70) || (c >= 97 && c <= 102),
};
const POSIX_UCP: Record<string, Test> = {
  alnum: uTest(/^[\p{L}\p{N}]$/u), alpha: uTest(/^\p{L}$/u), digit: isDigitU, lower: uTest(/^\p{Ll}$/u),
  upper: uTest(/^\p{Lu}$/u), space: isSpaceU, word: isWordU, punct: uTest(/^\p{P}$/u), graph: uTest(/^[^\p{Z}\p{C}]$/u), print: uTest(/^[^\p{C}]$/u),
};

const propCache = new Map<string, Test>();
const property = (raw: string): Test => {
  const key = raw;
  const hit = propCache.get(key);
  if (hit) return hit;
  const name = raw.replace(/[\s-]/g, '');
  const lower = name.toLowerCase();
  let t: Test | null = null;
  const special: Record<string, Test> = {
    any: () => true,
    'l&': uTest(/^[\p{Lu}\p{Ll}\p{Lt}]$/u), lc: uTest(/^[\p{Lu}\p{Ll}\p{Lt}]$/u),
    xan: uTest(/^[\p{L}\p{N}]$/u), xwd: isWordU, xsp: isSpaceU, xps: isSpaceU,
    xuc: (c) => c === 0x24 || c === 0x40 || c === 0x60 || (c >= 0xa0 && c <= 0xd7ff) || (c >= 0xe000),
  };
  if (special[lower]) t = special[lower]!;
  const tryRe = (src: string): Test | null => { try { const re = new RegExp(`^${src}$`, 'u'); return (c) => re.test(String.fromCodePoint(c)); } catch { return null; } };
  if (!t) {
    const title = name.split('_').map((w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join('_');
    const cands = [name, name.length <= 2 ? name.charAt(0).toUpperCase() + name.slice(1).toLowerCase() : title];
    for (const c of cands) {
      t = tryRe(`\\p{${c}}`) ?? tryRe(`\\p{Script=${c}}`);
      if (t) break;
    }
    if (!t && lower.startsWith('sc:')) t = tryRe(`\\p{Script=${name.slice(3)}}`);
    if (!t && lower.startsWith('scx:')) t = tryRe(`\\p{Script_Extensions=${name.slice(4)}}`);
  }
  if (!t) throw new RegexError(`unknown property \\p{${raw}}`);
  propCache.set(key, t);
  return t;
};

// ---------- syntax tree ----------

interface Flags { i: boolean; m: boolean; s: boolean; x: boolean; xx: boolean; n: boolean; U: boolean; J: boolean }

type Cond =
  | { k: 'grp'; n: number }
  | { k: 'name'; name: string }
  | { k: 'rec'; n: number | null; name?: string }
  | { k: 'define' }
  | { k: 'look'; node: Node };

type Node =
  | { t: 'empty' }
  | { t: 'char'; c: number; i: boolean }
  | { t: 'set'; test: Test }
  | { t: 'seq'; items: Node[] }
  | { t: 'alt'; alts: Node[] }
  | { t: 'group'; n: number; body: Node }
  | { t: 'nc'; body: Node }
  | { t: 'atomic'; body: Node }
  | { t: 'look'; ahead: boolean; neg: boolean; body: Node }
  | { t: 'rep'; body: Node; min: number; max: number; mode: 'g' | 'l' | 'p' }
  | { t: 'assert'; k: 'bol' | 'eol' | 'mbol' | 'meol' | 'A' | 'z' | 'Z' | 'b' | 'B' | 'G' }
  | { t: 'backref'; n: number; i: boolean }
  | { t: 'namedref'; name: string; i: boolean }
  | { t: 'relref'; n: number; i: boolean }
  | { t: 'numref'; digits: string; i: boolean }
  | { t: 'keep' }
  | { t: 'call'; n: number | null; name?: string }
  | { t: 'cond'; cond: Cond; yes: Node; no: Node }
  | { t: 'verb'; v: 'FAIL' | 'ACCEPT' | 'COMMIT' | 'PRUNE' | 'SKIP' | 'THEN' | 'MARK'; name?: string };

const seq = (items: Node[]): Node => (items.length === 0 ? { t: 'empty' } : items.length === 1 ? items[0]! : { t: 'seq', items });

class Parser {
  private i = 0;
  groupCount = 0;
  names = new Map<string, number[]>();
  ucp = false;
  private p: number[];
  constructor(src: string) { this.p = [...src].map((ch) => ch.codePointAt(0)!); }

  private peek(k = 0): number { return this.p[this.i + k] ?? -1; }
  private ch(k = 0): string { const c = this.peek(k); return c < 0 ? '' : String.fromCodePoint(c); }
  private eat(s: string): boolean {
    const cps = [...s];
    for (let k = 0; k < cps.length; k++) if (this.ch(k) !== cps[k]) return false;
    this.i += cps.length;
    return true;
  }
  private fail(msg: string): never { throw new RegexError(msg); }

  parse(flags: Flags): Node {
    // Leading start-of-pattern options such as (*UCP) or (*UTF).
    for (;;) {
      const m = /^\(\*([A-Z_]+(?:=\d+)?)\)/.exec(this.rest(40));
      if (!m || ['FAIL', 'F', 'ACCEPT', 'COMMIT', 'PRUNE', 'SKIP', 'THEN', 'MARK'].includes(m[1]!.replace(/=.*/, ''))) break;
      if (m[1] === 'UCP') this.ucp = true;
      this.i += [...m[0]].length;
    }
    const node = this.parseAlts({ ...flags }, false);
    if (this.i < this.p.length) this.fail('unmatched )');
    return node;
  }
  private rest(n: number): string { return String.fromCodePoint(...this.p.slice(this.i, this.i + n)); }

  private skipX(f: Flags): void {
    if (!f.x) return;
    for (;;) {
      const c = this.peek();
      if (c === 32 || (c >= 9 && c <= 13) || c === 0x85 || c === 0x2028 || c === 0x2029) { this.i++; continue; }
      if (c === 35) { while (this.i < this.p.length && this.peek() !== 10) this.i++; continue; }
      break;
    }
  }

  private parseAlts(f: Flags, branchReset: boolean): Node {
    const alts: Node[] = [];
    const g0 = this.groupCount;
    let gmax = g0;
    for (;;) {
      if (branchReset) this.groupCount = g0;
      alts.push(this.parseSeq(f));
      gmax = Math.max(gmax, this.groupCount);
      if (this.peek() === 124) { this.i++; continue; }
      break;
    }
    if (branchReset) this.groupCount = gmax;
    return alts.length === 1 ? alts[0]! : { t: 'alt', alts };
  }

  private parseSeq(f: Flags): Node {
    const items: Node[] = [];
    for (;;) {
      this.skipX(f);
      const c = this.peek();
      if (c < 0 || c === 124 || c === 41) break;
      const atom = this.parseAtom(f, items);
      if (atom === null) continue;
      this.skipX(f);
      items.push(this.parseQuant(atom, f));
    }
    return seq(items);
  }

  private parseQuant(atom: Node, f: Flags): Node {
    let node = atom;
    for (;;) {
      const c = this.ch();
      let min: number, max: number;
      if (c === '*') { min = 0; max = Infinity; this.i++; }
      else if (c === '+') { min = 1; max = Infinity; this.i++; }
      else if (c === '?') { min = 0; max = 1; this.i++; }
      else if (c === '{') {
        const m = /^\{\s*(\d*)\s*(?:(,)\s*(\d*)\s*)?\}/.exec(this.rest(30));
        if (!m || (m[1] === '' && (m[2] === undefined || m[3] === ''))) return node;
        min = m[1] === '' ? 0 : +m[1]!;
        max = m[2] === undefined ? min : m[3] === '' ? Infinity : +m[3]!;
        if (min > 65535 || (max !== Infinity && max > 65535)) this.fail('number too big in {} quantifier');
        if (max < min) this.fail('numbers out of order in {} quantifier');
        this.i += [...m[0]].length;
      } else return node;
      if (node.t === 'assert' || node.t === 'keep' || node.t === 'empty') {
        // PCRE allows a quantifier on an assertion and treats {0,...} as optional; keep it simple.
        if (min === 0) node = { t: 'empty' };
      }
      let mode: 'g' | 'l' | 'p' = f.U ? 'l' : 'g';
      if (this.ch() === '?') { mode = f.U ? 'g' : 'l'; this.i++; }
      else if (this.ch() === '+') { mode = 'p'; this.i++; }
      node = { t: 'rep', body: node, min, max, mode };
      this.skipX(f);
      if (!'*+?{'.includes(this.ch()) || this.ch() === '') return node;
      if (this.ch() === '{' && !/^\{\s*\d*\s*(,\s*\d*\s*)?\}/.test(this.rest(30))) return node;
      this.fail('quantifier does not follow a repeatable item');
    }
  }

  private name(term: string): string {
    const m = /^[A-Za-z_][A-Za-z0-9_]*/.exec(this.rest(64));
    if (!m) this.fail('group name expected');
    this.i += m[0].length;
    if (!this.eat(term)) this.fail('syntax error in group name');
    return m[0];
  }

  private addName(name: string, n: number, f: Flags): void {
    const list = this.names.get(name);
    if (list) {
      if (!f.J && !list.includes(n)) this.fail(`two named groups have the same name: ${name}`);
      if (!list.includes(n)) list.push(n);
    } else this.names.set(name, [n]);
  }

  private capture(f: Flags, name?: string): Node {
    const n = ++this.groupCount;
    if (name) this.addName(name, n, f);
    const body = this.parseAlts({ ...f }, false);
    this.close();
    return { t: 'group', n, body };
  }
  private close(): void { if (!this.eat(')')) this.fail('missing )'); }

  /** Returns null when the item produced nothing (comments, option settings). */
  private parseAtom(f: Flags, items: Node[]): Node | null {
    const c = this.ch();
    this.i++;
    switch (c) {
      case '.': return f.s ? { t: 'set', test: () => true } : { t: 'set', test: (x) => x !== 10 };
      case '^': return { t: 'assert', k: f.m ? 'mbol' : 'bol' };
      case '$': return { t: 'assert', k: f.m ? 'meol' : 'eol' };
      case '[': return this.parseClass(f);
      case '\\': return this.parseEscape(f, items);
      case '(': return this.parseGroup(f);
      case '*': case '+': case '?': this.fail('quantifier does not follow a repeatable item');
      // falls through (fail never returns)
      case '{': {
        const m = /^\s*(\d*)\s*(?:(,)\s*(\d*)\s*)?\}/.exec(this.rest(30));
        if (m && !(m[1] === '' && (m[2] === undefined || m[3] === ''))) this.fail('quantifier does not follow a repeatable item');
        return { t: 'char', c: 123, i: f.i };
      }
      default: return { t: 'char', c: this.p[this.i - 1]!, i: f.i };
    }
  }

  private parseGroup(f: Flags): Node | null {
    if (this.eat('*')) {
      const m = /^([A-Z]*)(?::([^)]*))?\)/.exec(this.rest(300));
      if (!m) this.fail('(*VERB) not recognized');
      this.i += [...m[0]].length;
      const name = m[2];
      let v = m[1]!;
      if (v === 'F') v = 'FAIL';
      if (v === '' && name !== undefined) v = 'MARK';
      if (name !== undefined && !name.length) this.fail(`(*${v}:) needs a name`);
      if (v === 'MARK' && name === undefined) this.fail('(*MARK) needs a name');
      if ((v === 'FAIL' || v === 'ACCEPT') && name !== undefined) this.fail(`(*${v}) does not take a name`);
      if (v === 'FAIL' || v === 'ACCEPT' || v === 'COMMIT' || v === 'PRUNE' || v === 'SKIP' || v === 'THEN' || v === 'MARK') return name === undefined ? { t: 'verb', v } : { t: 'verb', v, name };
      this.fail(`(*${v}) is not supported`);
    }
    if (!this.eat('?')) {
      if (f.n) { const body = this.parseAlts({ ...f }, false); this.close(); return { t: 'nc', body }; }
      return this.capture(f);
    }
    const c = this.ch();
    if (c === '#') { while (this.i < this.p.length && this.peek() !== 41) this.i++; this.close(); return null; }
    if (this.eat(':')) { const body = this.parseAlts({ ...f }, false); this.close(); return { t: 'nc', body }; }
    if (this.eat('|')) { const body = this.parseAlts({ ...f }, true); this.close(); return { t: 'nc', body }; }
    if (this.eat('>')) { const body = this.parseAlts({ ...f }, false); this.close(); return { t: 'atomic', body }; }
    if (this.eat('=')) { const body = this.parseAlts({ ...f }, false); this.close(); return { t: 'look', ahead: true, neg: false, body }; }
    if (this.eat('!')) { const body = this.parseAlts({ ...f }, false); this.close(); return { t: 'look', ahead: true, neg: true, body }; }
    if (this.eat('<=')) { const body = this.parseAlts({ ...f }, false); this.close(); return { t: 'look', ahead: false, neg: false, body }; }
    if (this.eat('<!')) { const body = this.parseAlts({ ...f }, false); this.close(); return { t: 'look', ahead: false, neg: true, body }; }
    if (this.eat('P<') || this.eat('<')) return this.capture(f, this.name('>'));
    if (this.eat("'")) return this.capture(f, this.name("'"));
    if (this.eat('P=')) return { t: 'namedref', name: this.name(')'), i: f.i };
    if (this.eat('P>') || this.eat('&')) return { t: 'call', n: null, name: this.name(')') };
    if (this.eat('R)')) return { t: 'call', n: 0 };
    const num = /^([+-]?)(\d+)\)/.exec(this.rest(12));
    if (num) {
      this.i += num[0].length;
      const k = +num[2]!;
      const n = num[1] === '+' ? this.groupCount + k : num[1] === '-' ? this.groupCount - k + 1 : k;
      if (num[1] && k === 0) this.fail('invalid relative group');
      if (n < 0) this.fail('reference to non-existent subpattern');
      return { t: 'call', n };
    }
    if (this.eat('(')) return this.parseCond(f);
    // Option settings: (?imnsxJU-imnsx) or (?^) scoped to the rest of the group, or (?i:...).
    const m = /^(\^?)([imnsxJU]*)(?:-([imnsxJU]*))?([:)])/.exec(this.rest(20));
    if (!m) this.fail('unrecognized character after (?');
    this.i += m[0].length;
    const nf: Flags = m[4] === ')' ? f : { ...f };
    if (m[1]) { nf.i = nf.m = nf.n = nf.s = nf.x = nf.xx = false; }
    const set = (s: string, on: boolean): void => {
      for (const ch of new Set(s)) {
        if (ch === 'x') { nf.x = on; nf.xx = on && s.includes('xx'); }
        else (nf as unknown as Record<string, boolean>)[ch] = on;
      }
    };
    set(m[2]!, true);
    if (m[3] !== undefined) set(m[3], false);
    if (m[4] === ')') return null;
    const body = this.parseAlts(nf, false);
    this.close();
    return { t: 'nc', body };
  }

  private parseCond(f: Flags): Node {
    let cond: Cond;
    const rest = this.rest(64);
    let m: RegExpExecArray | null;
    if ((m = /^([+-]?)(\d+)\)/.exec(rest))) {
      const k = +m[2]!;
      cond = { k: 'grp', n: m[1] === '+' ? this.groupCount + k : m[1] === '-' ? this.groupCount - k + 1 : k };
      this.i += m[0].length;
    } else if ((m = /^<([A-Za-z_]\w*)>\)/.exec(rest)) || (m = /^'([A-Za-z_]\w*)'\)/.exec(rest))) {
      cond = { k: 'name', name: m[1]! }; this.i += m[0].length;
    } else if ((m = /^R(\d*)\)/.exec(rest))) {
      cond = { k: 'rec', n: m[1] ? +m[1] : null }; this.i += m[0].length;
    } else if ((m = /^R&([A-Za-z_]\w*)\)/.exec(rest))) {
      cond = { k: 'rec', n: null, name: m[1]! }; this.i += m[0].length;
    } else if ((m = /^DEFINE\)/.exec(rest))) {
      cond = { k: 'define' }; this.i += m[0].length;
    } else if (rest.startsWith('?=') || rest.startsWith('?!') || rest.startsWith('?<=') || rest.startsWith('?<!')) {
      const node = this.parseGroup(f);
      if (!node || node.t !== 'look') this.fail('assertion expected after (?(');
      cond = { k: 'look', node };
    } else if ((m = /^([A-Za-z_]\w*)\)/.exec(rest))) {
      cond = { k: 'name', name: m[1]! }; this.i += m[0].length;
    } else this.fail('malformed number or name after (?(');
    const body = this.parseAlts({ ...f }, false);
    this.close();
    const alts = body.t === 'alt' ? body.alts : [body];
    if (alts.length > 2) this.fail('conditional subpattern contains more than two branches');
    if (cond.k === 'define' && alts.length > 1) this.fail('DEFINE subpattern contains more than one branch');
    return { t: 'cond', cond, yes: alts[0]!, no: alts[1] ?? { t: 'empty' } };
  }

  private hex(max: number): number {
    let s = '';
    while (s.length < max && /[0-9a-fA-F]/.test(this.ch())) { s += this.ch(); this.i++; }
    return s ? parseInt(s, 16) : 0;
  }

  private braced(re: RegExp, what: string): string {
    if (!this.eat('{')) this.fail(`${what}: { expected`);
    let s = '';
    while (this.i < this.p.length && this.ch() !== '}') { s += this.ch(); this.i++; }
    if (!this.eat('}') || !re.test(s.trim())) this.fail(`malformed ${what}`);
    return s.trim();
  }

  /** A single-character escape shared by class and non-class contexts; null if not one of these. */
  private charEscape(c: string): number | null {
    switch (c) {
      case 'a': return 7;
      case 'e': return 27;
      case 'f': return 12;
      case 'n': return 10;
      case 'r': return 13;
      case 't': return 9;
      case 'x': {
        if (this.ch() === '{') { const v = parseInt(this.braced(/^[0-9a-fA-F]+$/, '\\x{...}'), 16); if (v > 0x10ffff) this.fail('character code point value in \\x{} is too large'); return v; }
        return this.hex(2);
      }
      case 'o': return parseInt(this.braced(/^[0-7]+$/, '\\o{...}'), 8);
      case 'u': {
        if (this.ch() === '{') return parseInt(this.braced(/^[0-9a-fA-F]+$/, '\\u{...}'), 16);
        const v = this.hex(4); return v;
      }
      case 'c': {
        const x = this.peek();
        if (x < 32 || x > 126) this.fail('\\c must be followed by a printable ASCII character');
        this.i++;
        return (String.fromCharCode(x).toUpperCase().charCodeAt(0)) ^ 0x40;
      }
      case '0': {
        let s = '';
        while (s.length < 2 && /[0-7]/.test(this.ch())) { s += this.ch(); this.i++; }
        return parseInt(s || '0', 8);
      }
      default: return null;
    }
  }

  private classEscape(c: string, f: Flags): Test | null {
    const u = this.ucp;
    switch (c) {
      case 'd': return u ? isDigitU : isDigitA;
      case 'D': { const t = u ? isDigitU : isDigitA; return (x) => !t(x); }
      case 'w': return u ? isWordU : isWordA;
      case 'W': { const t = u ? isWordU : isWordA; return (x) => !t(x); }
      case 's': return u ? isSpaceU : isSpaceA;
      case 'S': { const t = u ? isSpaceU : isSpaceA; return (x) => !t(x); }
      case 'h': return isH;
      case 'H': return (x) => !isH(x);
      case 'v': return isV;
      case 'V': return (x) => !isV(x);
      case 'p': case 'P': {
        let name: string;
        if (this.ch() === '{') name = this.braced(/^\^?[\w&:= -]+$/, '\\p{...}');
        else { name = this.ch(); this.i++; if (!/[A-Za-z]/.test(name)) this.fail('malformed \\p'); }
        let neg = c === 'P';
        if (name.startsWith('^')) { neg = !neg; name = name.slice(1); }
        const t = property(name);
        void f;
        return neg ? (x) => !t(x) : t;
      }
      default: return null;
    }
  }

  private parseEscape(f: Flags, items: Node[]): Node | null {
    const c = this.ch();
    if (c === '') this.fail('\\ at end of pattern');
    this.i++;
    const t = this.classEscape(c, f);
    if (t) return { t: 'set', test: t };
    switch (c) {
      case 'b': return { t: 'assert', k: 'b' };
      case 'B': return { t: 'assert', k: 'B' };
      case 'A': return { t: 'assert', k: 'A' };
      case 'z': return { t: 'assert', k: 'z' };
      case 'Z': return { t: 'assert', k: 'Z' };
      case 'G': return { t: 'assert', k: 'G' };
      case 'K': return { t: 'keep' };
      case 'N': return { t: 'set', test: (x) => x !== 10 };
      case 'R': {
        // (?>\r\n|\n|\x0b|\f|\r|\x85|\x{2028}|\x{2029})
        return { t: 'atomic', body: { t: 'alt', alts: [seq([{ t: 'char', c: 13, i: false }, { t: 'char', c: 10, i: false }]), { t: 'set', test: isV }] } };
      }
      case 'X': return { t: 'atomic', body: seq([{ t: 'set', test: uTest(/^\P{M}$/u) }, { t: 'rep', body: { t: 'set', test: uTest(/^\p{M}$/u) }, min: 0, max: Infinity, mode: 'g' }]) };
      case 'Q': {
        const before = items.length;
        while (this.i < this.p.length) {
          if (this.ch() === '\\' && this.ch(1) === 'E') { this.i += 2; break; }
          items.push({ t: 'char', c: this.peek(), i: f.i });
          this.i++;
        }
        // The last quoted character is returned so a following quantifier applies to it alone.
        return items.length > before ? items.pop()! : null;
      }
      case 'E': return null;
      case 'g': {
        let m: RegExpExecArray | null;
        const rest = this.rest(40);
        if ((m = /^\{(-?\d+)\}/.exec(rest)) || (m = /^(-?\d+)/.exec(rest))) {
          this.i += m[0].length;
          const k = +m[1]!;
          if (k === 0) this.fail('a numbered reference must not be zero');
          return k < 0 ? { t: 'relref', n: this.groupCount + k + 1, i: f.i } : { t: 'backref', n: k, i: f.i };
        }
        if ((m = /^\{([A-Za-z_]\w*)\}/.exec(rest))) { this.i += m[0].length; return { t: 'namedref', name: m[1]!, i: f.i }; }
        if ((m = /^<([+-]?\d+)>/.exec(rest)) || (m = /^'([+-]?\d+)'/.exec(rest))) {
          this.i += m[0].length;
          const s = m[1]!;
          const k = +s.replace('+', '');
          const n = s.startsWith('+') ? this.groupCount + k : s.startsWith('-') ? this.groupCount + k + 1 : k;
          return { t: 'call', n };
        }
        if ((m = /^<([A-Za-z_]\w*)>/.exec(rest)) || (m = /^'([A-Za-z_]\w*)'/.exec(rest))) { this.i += m[0].length; return { t: 'call', n: null, name: m[1]! }; }
        this.fail('\\g is not followed by a group name or number');
      }
      // falls through (fail never returns)
      case 'k': {
        const m = /^(?:<([A-Za-z_]\w*)>|'([A-Za-z_]\w*)'|\{([A-Za-z_]\w*)\})/.exec(this.rest(40));
        if (!m) this.fail('\\k is not followed by a group name');
        this.i += m[0].length;
        return { t: 'namedref', name: (m[1] ?? m[2] ?? m[3])!, i: f.i };
      }
      default: break;
    }
    if (/[1-9]/.test(c)) {
      let digits = c;
      while (/\d/.test(this.ch())) { digits += this.ch(); this.i++; }
      return { t: 'numref', digits, i: f.i };
    }
    const ce = this.charEscape(c);
    if (ce !== null) return { t: 'char', c: ce, i: f.i };
    if (/[A-Za-z0-9]/.test(c)) this.fail(`unrecognized escape \\${c}`);
    return { t: 'char', c: this.p[this.i - 1]!, i: f.i };
  }

  private parseClass(f: Flags): Node {
    let neg = false;
    if (this.ch() === '^') { neg = true; this.i++; }
    const ranges: [number, number][] = [];
    const tests: Test[] = [];
    let first = true;
    // Reads one class member: a code point, or a Test for escapes like \d and POSIX names.
    const member = (): number | Test | null => {
      if (f.xx) while (this.ch() === ' ' || this.ch() === '\t') this.i++;
      const c = this.ch();
      if (c === '') this.fail('missing terminating ] for character class');
      if (c === '[') {
        const m = /^\[:(\^?)([a-z]+):\]/.exec(this.rest(16));
        if (m) {
          const base = (this.ucp ? POSIX_UCP[m[2]!] : undefined) ?? POSIX[m[2]!];
          if (!base) this.fail(`unknown POSIX class name ${m[2]}`);
          this.i += m[0].length;
          return m[1] ? (x) => !base(x) : base;
        }
        if (/^\[[.=]/.test(this.rest(2))) this.fail('POSIX collating elements are not supported');
      }
      if (c === '\\') {
        this.i++;
        const e = this.ch();
        this.i++;
        if (e === 'Q') {
          // Quoted run inside a class: each character is a member.
          const start = this.i;
          while (this.i < this.p.length && !(this.ch() === '\\' && this.ch(1) === 'E')) this.i++;
          const cps = this.p.slice(start, this.i);
          if (this.i < this.p.length) this.i += 2;
          for (const cp of cps) ranges.push([cp, cp]);
          return null;
        }
        if (e === 'E') return null;
        if (e === 'b') return 8;
        if (e === 'N' || e === 'R' || e === 'X' || e === 'B') this.fail(`escape \\${e} is not allowed in a class`);
        const t = this.classEscape(e, f);
        if (t) return t;
        const ce = this.charEscape(e);
        if (ce !== null) return ce;
        if (/[1-7]/.test(e)) {
          let s = e;
          while (s.length < 3 && /[0-7]/.test(this.ch())) { s += this.ch(); this.i++; }
          return parseInt(s, 8);
        }
        if (/[A-Za-z0-9]/.test(e)) this.fail(`unrecognized escape \\${e} in class`);
        return this.p[this.i - 1]!;
      }
      this.i++;
      return this.p[this.i - 1]!;
    };
    for (;;) {
      if (f.xx) while (this.ch() === ' ' || this.ch() === '\t') this.i++;
      if (this.ch() === ']' && !first) { this.i++; break; }
      first = false;
      const a = member();
      if (a === null) continue;
      if (typeof a !== 'number') { tests.push(a); continue; }
      if (this.ch() === '-' && this.ch(1) !== ']' && this.ch(1) !== '') {
        const save = this.i;
        this.i++;
        const b = member();
        if (typeof b === 'number') {
          if (b < a) this.fail('range out of order in character class');
          ranges.push([a, b]);
          continue;
        }
        // "a-\d": the hyphen is literal.
        this.i = save + 1;
        ranges.push([a, a], [45, 45]);
        if (b) tests.push(b);
        continue;
      }
      ranges.push([a, a]);
    }
    const base: Test = (x) => {
      for (const [lo, hi] of ranges) if (x >= lo && x <= hi) return true;
      for (const t of tests) if (t(x)) return true;
      return false;
    };
    const ci = f.i;
    const inClass: Test = ci ? (x) => base(x) || variants(x).some((v) => v !== x && base(v)) : base;
    return { t: 'set', test: neg ? (x) => !inClass(x) : inClass };
  }
}

// ---------- program ----------

const enum Op {
  Char, CharI, CharsI, Set, Split, Jmp, Save, Assert, Backref, Atomic, Look, Mark, Progress, Keep, Call, GEnd, CondGrp, CondName, CondRec, CondLook, Fail, Accept, Verb, Match, EndAt,
  AltEnter, Name,
}

interface Inst {
  op: Op;
  a: number;
  b: number;
  c?: number;
  test?: Test;
  ns?: number[];
  i?: boolean;
  k?: string;
  neg?: boolean;
  ahead?: boolean;
  /** Lookbehind alternatives: program start and the fixed length range of each. */
  alts?: { pc: number; min: number; max: number }[];
  /** Folded text for CharsI. */
  str?: number[];
  name?: string;
}

/** Minimum and maximum number of code points a node can match (max Infinity when unbounded or unknown). */
const width = (n: Node): [number, number] => {
  switch (n.t) {
    case 'empty': case 'assert': case 'keep': case 'look': case 'verb': return [0, 0];
    case 'char': case 'set': return [1, 1];
    case 'seq': return n.items.reduce<[number, number]>((acc, x) => { const w = width(x); return [acc[0] + w[0], acc[1] + w[1]]; }, [0, 0]);
    case 'alt': { const ws = n.alts.map(width); return [Math.min(...ws.map((w) => w[0])), Math.max(...ws.map((w) => w[1]))]; }
    case 'group': case 'nc': case 'atomic': return width(n.body);
    case 'rep': { const w = width(n.body); return [w[0] * n.min, n.max === Infinity ? (w[1] === 0 ? 0 : Infinity) : w[1] * n.max]; }
    case 'cond': { const a = width(n.yes), b = width(n.no); return [Math.min(a[0], b[0]), Math.max(a[1], b[1])]; }
    default: return [0, Infinity];
  }
};

export interface RegexMatch {
  /** Code point offsets of the whole match. */
  start: number;
  end: number;
  /** Captured text for groups 1..n; undefined when the group did not take part. */
  groups: (string | undefined)[];
  whole: string;
}

interface Frame { n: number; ret: number; caps: number[]; pos: number; up: Frame | null; depth: number }

// Backtracking stack entries: 0 an alternative to resume (alt = alternation it belongs to), 1 a capture
// to restore, 2 a register to restore, 3 a backtracking verb, 5 entry into an alternation, 6 entry into a
// subroutine call, 7 a (*MARK:name).
type BT =
  | { k: 0; pc: number; pos: number; cs: Frame | null; alt?: number }
  | { k: 1; i: number; v: number }
  | { k: 2; i: number; v: number }
  | { k: 3; v: 'COMMIT' | 'PRUNE' | 'SKIP' | 'THEN'; pos: number; frame: Frame | null; name?: string; alt: number }
  | { k: 5; id: number }
  | { k: 6; frame: Frame }
  | { k: 7; name: string; pos: number };

/** run() results besides a match position: -1 no match, then verbs ending the attempt, then (*ACCEPT) inside an atomic group. */
const ABORT = -2, ABORT_ALL = -3, THEN_ESC = -4, ACCEPTED = -5;
type RunKind = 'top' | 'atomic' | 'assert';

export class Regex {
  readonly groupCount: number;
  readonly names: ReadonlyMap<string, number[]>;
  private prog: Inst[] = [];
  private groupStart: number[] = [];
  private regCount = 0;
  private altCount = 0;
  private altStack: number[] = [];
  /** Alternation a (*THEN) escaping a sub-run belongs to, and where (*ACCEPT) ended inside an atomic group. */
  private thenAlt = -1;
  private acceptPos = -1;
  private anchored = false;
  // Per-exec state
  private s: number[] = [];
  private caps: number[] = [];
  private regs: number[] = [];
  private steps = 0;
  private searchStart = 0;
  private skipTo = -1;
  private notEmptyAt = -1;

  constructor(pattern: string, opts: { caseInsensitive?: boolean } = {}) {
    const parser = new Parser(pattern);
    const flags: Flags = { i: !!opts.caseInsensitive, m: false, s: false, x: false, xx: false, n: false, U: false, J: false };
    let root = parser.parse(flags);
    this.groupCount = parser.groupCount;
    this.ucp = parser.ucp;
    this.names = parser.names;
    root = this.resolve(root);
    this.anchored = this.isAnchored(root);
    this.required = /\(\*ACCEPT/.test(pattern) ? [] : this.requiredChars(root);
    const p = this.prog;
    p.push({ op: Op.Save, a: 0, b: 0 });
    this.groupStart[0] = 1;
    this.emit(root);
    p.push({ op: Op.Save, a: 1, b: 0 }, { op: Op.GEnd, a: 0, b: 0 }, { op: Op.Match, a: 0, b: 0 });
    for (const inst of p) if (inst.op === Op.Call && this.groupStart[inst.a] === undefined) throw new RegexError('reference to non-existent subpattern');
  }

  /** Characters every match must contain; a subject without one of them cannot match (saves hopeless backtracking). */
  private requiredChars(n: Node): { c: number; i: boolean }[] {
    switch (n.t) {
      case 'char': return [{ c: n.i ? fullFold(n.c)[0]! : n.c, i: n.i }];
      case 'seq': return n.items.flatMap((x) => this.requiredChars(x));
      case 'group': case 'nc': case 'atomic': return this.requiredChars(n.body);
      case 'rep': return n.min >= 1 ? this.requiredChars(n.body) : [];
      default: return [];
    }
  }
  private required: { c: number; i: boolean }[] = [];

  private isAnchored(n: Node): boolean {
    if (n.t === 'assert') return n.k === 'A' || n.k === 'bol';
    if (n.t === 'seq') return n.items.length > 0 && this.isAnchored(n.items[0]!);
    if (n.t === 'alt') return n.alts.every((a) => this.isAnchored(a));
    if (n.t === 'group' || n.t === 'nc' || n.t === 'atomic') return this.isAnchored(n.body);
    return false;
  }

  private nameToNums(name: string): number[] {
    const ns = this.names.get(name);
    if (!ns) throw new RegexError(`reference to non-existent subpattern ${name}`);
    return ns;
  }

  /** Turns ambiguous numbered escapes into back-references or octal characters, and checks references. */
  private resolve(n: Node): Node {
    const r = (x: Node): Node => this.resolve(x);
    switch (n.t) {
      case 'seq': return { ...n, items: n.items.map(r) };
      case 'alt': return { ...n, alts: n.alts.map(r) };
      case 'group': case 'nc': case 'atomic': case 'look': return { ...n, body: r(n.body) } as Node;
      case 'rep': return { ...n, body: r(n.body) };
      case 'cond': {
        let cond = n.cond;
        if (cond.k === 'look') cond = { k: 'look', node: r(cond.node) };
        if (cond.k === 'grp' && (cond.n < 1 || cond.n > this.groupCount)) throw new RegexError('reference to non-existent subpattern');
        if (cond.k === 'name' && !this.names.has(cond.name)) {
          throw new RegexError(`reference to non-existent subpattern ${cond.name}`);
        }
        return { ...n, cond, yes: r(n.yes), no: r(n.no) };
      }
      case 'numref': {
        const k = parseInt(n.digits, 10);
        if (k < 10 || k <= this.groupCount) {
          if (k > this.groupCount) throw new RegexError('reference to non-existent subpattern');
          return { t: 'backref', n: k, i: n.i };
        }
        const m = /^([0-7]{1,3})(.*)$/.exec(n.digits);
        if (!m) throw new RegexError('reference to non-existent subpattern');
        return seq([{ t: 'char', c: parseInt(m[1]!, 8), i: n.i }, ...[...m[2]!].map((ch): Node => ({ t: 'char', c: ch.codePointAt(0)!, i: n.i }))]);
      }
      case 'backref': if (n.n > this.groupCount) throw new RegexError('reference to non-existent subpattern'); return n;
      case 'relref': if (n.n < 1) throw new RegexError('reference to non-existent subpattern'); return { t: 'backref', n: n.n, i: n.i };
      case 'namedref': this.nameToNums(n.name); return n;
      case 'call': {
        const k = n.n ?? this.nameToNums(n.name!)[0]!;
        if (k > this.groupCount) throw new RegexError('reference to non-existent subpattern');
        return { t: 'call', n: k };
      }
      default: return n;
    }
  }

  private push(inst: Inst): number {
    this.prog.push(inst);
    if (this.prog.length > MAX_PROGRAM) throw new RegexError('regular expression is too large');
    return this.prog.length - 1;
  }

  private nullable(n: Node): boolean { return width(n)[0] === 0; }

  private emitSub(body: Node, tail?: Inst): number {
    const jmp = this.push({ op: Op.Jmp, a: 0, b: 0 });
    const start = this.prog.length;
    this.emit(body);
    if (tail) this.push(tail);
    this.push({ op: Op.Match, a: 0, b: 0 });
    this.prog[jmp]!.a = this.prog.length;
    return start;
  }

  private emit(n: Node): void {
    const p = this.prog;
    switch (n.t) {
      case 'empty': return;
      case 'char': {
        if (!n.i) { this.push({ op: Op.Char, a: n.c, b: 0 }); return; }
        const f = fullFold(n.c);
        this.push(f.length === 1 ? { op: Op.CharI, a: fold(n.c), b: 0 } : { op: Op.CharsI, a: 0, b: 0, str: f });
        return;
      }
      case 'set': this.push({ op: Op.Set, a: 0, b: 0, test: n.test }); return;
      case 'seq': {
        // Runs of caseless letters compare as folded text, so "strasse" matches "STRAßE".
        for (let k = 0; k < n.items.length; k++) {
          const x = n.items[k]!;
          if (x.t === 'char' && x.i && n.items[k + 1]?.t === 'char' && (n.items[k + 1] as { i: boolean }).i) {
            const str: number[] = [];
            while (k < n.items.length && n.items[k]!.t === 'char' && (n.items[k] as { i: boolean }).i) { str.push(...fullFold((n.items[k] as { c: number }).c)); k++; }
            k--;
            this.push({ op: Op.CharsI, a: 0, b: 0, str });
          } else this.emit(x);
        }
        return;
      }
      case 'alt': {
        const jumps: number[] = [];
        const id = this.altCount++;
        this.altStack.push(id);
        this.push({ op: Op.AltEnter, a: id, b: 0 });
        n.alts.forEach((alt, idx) => {
          if (idx < n.alts.length - 1) {
            const split = this.push({ op: Op.Split, a: 0, b: 0, c: id });
            p[split]!.a = p.length;
            this.emit(alt);
            jumps.push(this.push({ op: Op.Jmp, a: 0, b: 0 }));
            p[split]!.b = p.length;
          } else this.emit(alt);
        });
        for (const j of jumps) p[j]!.a = p.length;
        this.altStack.pop();
        return;
      }
      case 'group': {
        this.push({ op: Op.Save, a: 2 * n.n, b: 0 });
        if (this.groupStart[n.n] === undefined) this.groupStart[n.n] = p.length;
        this.emit(n.body);
        this.push({ op: Op.Save, a: 2 * n.n + 1, b: 0 });
        this.push({ op: Op.GEnd, a: n.n, b: 0 });
        return;
      }
      case 'nc': this.emit(n.body); return;
      case 'atomic': { const start = this.emitSub(n.body); this.push({ op: Op.Atomic, a: start, b: 0 }); return; }
      case 'look': {
        if (n.ahead) { const start = this.emitSub(n.body); this.push({ op: Op.Look, a: 0, b: 0, c: start, ahead: true, neg: n.neg }); return; }
        const alts = (n.body.t === 'alt' ? n.body.alts : [n.body]).map((alt) => {
          const [min, max] = width(alt);
          if (max === Infinity) throw new RegexError('lookbehind assertion is not fixed length');
          if (max > 65535) throw new RegexError('lookbehind is too long');
          return { pc: this.emitSub(alt, { op: Op.EndAt, a: 0, b: 0 }), min, max };
        });
        this.push({ op: Op.Look, a: 0, b: 0, ahead: false, neg: n.neg, alts });
        return;
      }
      case 'rep': {
        if (n.max === 0) {
          // X{0} matches nothing, but its groups can still be called as subroutines: keep the code, jump over it.
          const j = this.push({ op: Op.Jmp, a: 0, b: 0 });
          this.emit(n.body);
          p[j]!.a = p.length;
          return;
        }
        if (n.mode === 'p') { this.emit({ t: 'atomic', body: { ...n, mode: 'g' } }); return; }
        const greedy = n.mode === 'g';
        for (let k = 0; k < n.min; k++) this.emit(n.body);
        if (n.max === Infinity) {
          const nul = this.nullable(n.body);
          const loop = this.push({ op: Op.Split, a: 0, b: 0 });
          const bodyStart = p.length;
          const reg = nul ? this.regCount++ : -1;
          if (nul) this.push({ op: Op.Mark, a: reg, b: 0 });
          this.emit(n.body);
          const prog = nul ? this.push({ op: Op.Progress, a: reg, b: 0 }) : -1;
          this.push({ op: Op.Jmp, a: loop, b: 0 });
          const exit = p.length;
          p[loop]!.a = greedy ? bodyStart : exit;
          p[loop]!.b = greedy ? exit : bodyStart;
          if (prog >= 0) p[prog]!.b = exit;
          return;
        }
        const splits: number[] = [];
        for (let k = n.min; k < n.max; k++) {
          const s = this.push({ op: Op.Split, a: 0, b: 0 });
          splits.push(s);
          p[s]!.a = p.length;
          this.emit(n.body);
        }
        const exit = p.length;
        for (const s of splits) {
          const body = p[s]!.a;
          p[s]!.a = greedy ? body : exit;
          p[s]!.b = greedy ? exit : body;
        }
        return;
      }
      case 'assert': this.push({ op: Op.Assert, a: 0, b: 0, k: n.k }); return;
      case 'backref': this.push({ op: Op.Backref, a: 0, b: 0, ns: [n.n], i: n.i }); return;
      case 'namedref': this.push({ op: Op.Backref, a: 0, b: 0, ns: this.nameToNums(n.name), i: n.i }); return;
      case 'keep': this.push({ op: Op.Keep, a: 0, b: 0 }); return;
      case 'call': this.push({ op: Op.Call, a: n.n!, b: 0 }); return;
      case 'verb':
        if (n.v === 'FAIL') this.push({ op: Op.Fail, a: 0, b: 0 });
        else if (n.v === 'ACCEPT') this.push({ op: Op.Accept, a: 0, b: 0 });
        else if (n.v === 'MARK') this.push({ op: Op.Name, a: 0, b: 0, name: n.name });
        else {
          // (*PRUNE:n), (*THEN:n) and (*COMMIT:n) also set a mark; (*SKIP:n) skips to one.
          if (n.name !== undefined && n.v !== 'SKIP') this.push({ op: Op.Name, a: 0, b: 0, name: n.name });
          const inst: Inst = { op: Op.Verb, a: 0, b: 0, k: n.v, c: this.altStack.length ? this.altStack[this.altStack.length - 1]! : -1 };
          if (n.v === 'SKIP' && n.name !== undefined) inst.name = n.name;
          this.push(inst);
        }
        return;
      case 'cond': {
        const c = n.cond;
        let at: number;
        if (c.k === 'define') { const j = this.push({ op: Op.Jmp, a: 0, b: 0 }); this.emit(n.yes); p[j]!.a = p.length; return; }
        if (c.k === 'look') {
          const look = c.node as Extract<Node, { t: 'look' }>;
          const before = p.length;
          this.emit(look);
          const li = p.length - 1;
          void before;
          p[li]!.op = Op.CondLook;
          at = li;
        } else if (c.k === 'grp') at = this.push({ op: Op.CondGrp, a: 0, b: 0, ns: [c.n] });
        else if (c.k === 'name') at = this.push({ op: Op.CondName, a: 0, b: 0, ns: this.nameToNums(c.name) });
        else at = this.push({ op: Op.CondRec, a: 0, b: 0, c: c.name ? this.nameToNums(c.name)[0]! : c.n ?? -1 });
        // Condition instructions jump to `a` (yes) or `b` (no).
        p[at]!.a = p.length;
        this.emit(n.yes);
        const j = this.push({ op: Op.Jmp, a: 0, b: 0 });
        p[at]!.b = p.length;
        this.emit(n.no);
        p[j]!.a = p.length;
        return;
      }
      default: throw new RegexError('internal: unresolved reference');
    }
  }

  // ---------- matching ----------

  private setCap(stack: BT[], i: number, v: number): void {
    stack.push({ k: 1, i, v: this.caps[i]! });
    this.caps[i] = v;
  }

  private isWordAt(pos: number): boolean {
    const c = this.s[pos];
    if (c === undefined) return false;
    return this.ucp ? isWordU(c) : isWordA(c);
  }
  private ucp = false;

  private assert(k: string, pos: number): boolean {
    const s = this.s, n = s.length;
    switch (k) {
      case 'bol': case 'A': return pos === 0;
      case 'mbol': return pos === 0 || (s[pos - 1] === 10 && pos !== n);
      case 'eol': case 'Z': return pos === n || (pos === n - 1 && s[pos] === 10);
      case 'meol': return pos === n || s[pos] === 10;
      case 'z': return pos === n;
      case 'G': return pos === this.searchStart;
      case 'b': return this.isWordAt(pos - 1) !== this.isWordAt(pos);
      case 'B': return this.isWordAt(pos - 1) === this.isWordAt(pos);
      default: return false;
    }
  }

  /** Matches folded text at `pos`, consuming whole subject characters; returns the new position or -1. */
  private matchFolded(target: number[], pos: number): number {
    const s = this.s;
    let j = 0;
    while (j < target.length) {
      if (pos >= s.length) return -1;
      const f = fullFold(s[pos]!);
      if (j + f.length > target.length) return -1;
      for (let k = 0; k < f.length; k++) if (target[j + k] !== f[k]) return -1;
      j += f.length;
      pos++;
    }
    return pos;
  }

  /**
   * Runs the program from `pc` at `pos`. Returns the end position, -1 for no match, ABORT / ABORT_ALL
   * when a backtracking verb ends the attempt, THEN_ESC when a (*THEN) belongs to an alternation outside
   * this run, or ACCEPTED when (*ACCEPT) ended the match inside an atomic group. On success the captures
   * stay changed (callers log them for their own backtracking); otherwise they are restored.
   */
  private run(startPc: number, startPos: number, cs0: Frame | null, endAt: number, kind: RunKind): number {
    const prog = this.prog, s = this.s, n = s.length;
    const top = kind === 'top';
    const stack: BT[] = [];
    let pc = startPc, pos = startPos, cs = cs0;
    const capsSnap = (): number[] => this.caps.slice();
    const logDiff = (snap: number[]): void => {
      for (let i = 0; i < snap.length; i++) if (snap[i] !== this.caps[i]) stack.push({ k: 1, i, v: snap[i]! });
    };
    // A (*THEN) to process before ordinary backtracking (it escaped an atomic group).
    let pendingThen: { alt: number; frame: Frame | null } | null = null;
    for (;;) {
      if (++this.steps > MAX_STEPS) throw new StepLimit();
      const ins = prog[pc]!;
      let ok = true;
      switch (ins.op) {
        case Op.Char: if (pos < n && s[pos] === ins.a) { pos++; pc++; } else ok = false; break;
        case Op.CharI: if (pos < n && fold(s[pos]!) === ins.a && fullFold(s[pos]!).length === 1) { pos++; pc++; } else ok = false; break;
        case Op.CharsI: { const r = this.matchFolded(ins.str!, pos); if (r >= 0) { pos = r; pc++; } else ok = false; break; }
        case Op.Set: if (pos < n && ins.test!(s[pos]!)) { pos++; pc++; } else ok = false; break;
        case Op.Split: stack.push(ins.c === undefined ? { k: 0, pc: ins.b, pos, cs } : { k: 0, pc: ins.b, pos, cs, alt: ins.c }); pc = ins.a; break;
        case Op.AltEnter: stack.push({ k: 5, id: ins.a }); pc++; break;
        case Op.Name: stack.push({ k: 7, name: ins.name!, pos }); pc++; break;
        case Op.Jmp: pc = ins.a; break;
        case Op.Save: this.setCap(stack, ins.a, pos); pc++; break;
        case Op.Keep: this.setCap(stack, 0, pos); pc++; break;
        case Op.Assert: if (this.assert(ins.k!, pos)) pc++; else ok = false; break;
        case Op.Backref: {
          let gi = -1;
          for (const g of ins.ns!) if (this.caps[2 * g + 1]! >= 0 && this.caps[2 * g]! >= 0) { gi = g; break; }
          if (gi < 0) { ok = false; break; }
          const a = this.caps[2 * gi]!, b = this.caps[2 * gi + 1]!;
          if (ins.i) {
            const target: number[] = [];
            for (let k = a; k < b; k++) target.push(...fullFold(s[k]!));
            const r = this.matchFolded(target, pos);
            if (r >= 0) { pos = r; pc++; } else ok = false;
            break;
          }
          const len = b - a;
          if (pos + len > n) { ok = false; break; }
          for (let k = 0; k < len; k++) if (s[a + k] !== s[pos + k]) { ok = false; break; }
          if (ok) { pos += len; pc++; }
          break;
        }
        case Op.Atomic: {
          const snap = capsSnap();
          const r = this.run(ins.a, pos, cs, -1, 'atomic');
          // Verbs backtracked into while the group is being matched have their normal effect.
          if (r === ABORT || r === ABORT_ALL) return this.unwind(stack, r);
          if (r === ACCEPTED) {
            if (top) return this.acceptPos;
            if (kind === 'atomic') return ACCEPTED;
            return this.acceptPos;
          }
          if (r === THEN_ESC) { ok = false; pendingThen = { alt: this.thenAlt, frame: null }; break; }
          if (r < 0) { ok = false; break; }
          logDiff(snap);
          pos = r;
          pc++;
          break;
        }
        case Op.Look: case Op.CondLook: {
          const snap = capsSnap();
          let matched = false;
          const cond = ins.op === Op.CondLook;
          // Returns true when the attempt should stop (a verb or (*THEN) settled the assertion).
          const attempt = (start: number, pc0: number, end: number): boolean => {
            const r = this.run(pc0, start, cs, end, 'assert');
            if (r >= 0 || r === ACCEPTED) { matched = true; return true; }
            if (r === ABORT || r === ABORT_ALL) {
              if (!ins.neg && !cond) throw new VerbExit(r);
              return true;
            }
            return r === THEN_ESC;
          };
          try {
            if (ins.ahead) attempt(pos, ins.c!, -1);
            else {
              outer: for (const alt of ins.alts!) {
                for (let len = alt.max; len >= alt.min; len--) {
                  if (pos - len < 0) continue;
                  if (attempt(pos - len, alt.pc, pos)) break outer;
                }
              }
            }
          } catch (e) {
            if (e instanceof VerbExit) return this.unwind(stack, e.code);
            throw e;
          }
          if (ins.neg) {
            if (matched) for (let i = 0; i < snap.length; i++) this.caps[i] = snap[i]!;
            matched = !matched;
          } else if (matched) logDiff(snap);
          if (cond) pc = matched ? ins.a : ins.b;
          else if (matched) pc++;
          else ok = false;
          break;
        }
        case Op.Mark: stack.push({ k: 2, i: ins.a, v: this.regs[ins.a]! }); this.regs[ins.a] = pos; pc++; break;
        case Op.Progress: pc = this.regs[ins.a] === pos ? ins.b : pc + 1; break;
        case Op.Call: {
          const g = ins.a;
          let depth = 0;
          for (let f = cs; f; f = f.up) { depth++; if (f.n === g && f.pos === pos) { ok = false; break; } }
          if (!ok) break;
          if (depth >= MAX_CALL_DEPTH) throw new StepLimit();
          cs = { n: g, ret: pc + 1, caps: this.caps.slice(), pos, up: cs, depth: depth + 1 };
          stack.push({ k: 6, frame: cs });
          pc = this.groupStart[g]!;
          break;
        }
        case Op.GEnd: {
          if (cs && cs.n === ins.a) {
            for (let i = 0; i < cs.caps.length; i++) if (this.caps[i] !== cs.caps[i]) this.setCap(stack, i, cs.caps[i]!);
            pc = cs.ret;
            cs = cs.up;
          } else pc++;
          break;
        }
        case Op.CondGrp: case Op.CondName: {
          const set = ins.ns!.some((g) => this.caps[2 * g + 1]! >= 0);
          pc = set ? ins.a : ins.b;
          break;
        }
        case Op.CondRec: {
          const want = ins.c!;
          const yes = cs !== null && (want < 0 || cs.n === want);
          pc = yes ? ins.a : ins.b;
          break;
        }
        case Op.Fail: ok = false; break;
        case Op.Accept: {
          if (cs && cs !== cs0) {
            // Return from the innermost subroutine call.
            for (let i = 0; i < cs.caps.length; i++) if (this.caps[i] !== cs.caps[i]) this.setCap(stack, i, cs.caps[i]!);
            pc = cs.ret;
            cs = cs.up;
            break;
          }
          if (kind === 'atomic') { this.acceptPos = pos; return ACCEPTED; }
          return pos;
        }
        case Op.Verb: stack.push({ k: 3, v: ins.k as 'COMMIT', pos, frame: cs, alt: ins.c ?? -1, ...(ins.name !== undefined ? { name: ins.name } : {}) }); pc++; break;
        case Op.EndAt: if (pos === endAt) pc++; else ok = false; break;
        case Op.Match:
          if (top && pos === this.notEmptyAt && this.caps[0] === pos) { ok = false; break; }
          return pos;
      }
      if (ok) continue;
      // Backtrack.
      const popTo = (want: (e: BT) => boolean): BT | null => {
        while (stack.length) {
          const e = stack.pop()!;
          if (e.k === 1) this.caps[e.i] = e.v;
          else if (e.k === 2) this.regs[e.i] = e.v;
          if (want(e)) return e;
        }
        return null;
      };
      let resumed = false;
      for (;;) {
        let then = pendingThen;
        pendingThen = null;
        let e: BT | null = null;
        if (!then) {
          e = popTo((x) => x.k === 0 || x.k === 3);
          if (!e) break;
          if (e.k === 0) { pc = e.pc; pos = e.pos; cs = e.cs; resumed = true; break; }
          if (e.k !== 3) continue;
          const sub = e.frame !== null && e.frame !== cs0;
          if (e.v === 'THEN') then = { alt: e.alt, frame: sub ? e.frame : null };
          else {
            if (sub) {
              // Inside a subroutine call these verbs make the call fail; matching goes on outside it.
              const f = e.frame;
              popTo((x) => x.k === 6 && x.frame === f);
              continue;
            }
            if (e.v === 'SKIP') {
              if (e.name !== undefined) {
                const nm = e.name;
                let at = -1;
                for (let k = stack.length - 1; k >= 0; k--) { const x = stack[k]!; if (x.k === 7 && x.name === nm) { at = x.pos; break; } }
                if (at < 0) continue; // no such mark on this path: the (*SKIP:n) is ignored
                this.skipTo = at;
              } else this.skipTo = e.pos;
            }
            return this.unwind(stack, e.v === 'COMMIT' ? ABORT_ALL : ABORT);
          }
        }
        // (*THEN): on to the next alternative of the innermost enclosing alternation.
        const { alt, frame } = then;
        const hit = popTo((x) => (x.k === 0 && x.alt === alt && alt >= 0) || (x.k === 5 && x.id === alt) || (frame !== null && x.k === 6 && x.frame === frame));
        if (hit && hit.k === 0) { pc = hit.pc; pos = hit.pos; cs = hit.cs; resumed = true; break; }
        if (hit) continue; // the alternation (or subroutine) has no more branches: it fails as a whole
        if (alt < 0) return ABORT; // no enclosing alternation: acts like (*PRUNE)
        this.thenAlt = alt;
        return THEN_ESC;
      }
      if (!resumed) return -1;
    }
  }

  private unwind(stack: BT[], code: number): number {
    while (stack.length) {
      const e = stack.pop()!;
      if (e.k === 1) this.caps[e.i] = e.v;
      else if (e.k === 2) this.regs[e.i] = e.v;
    }
    return code;
  }

  /** Finds the first match at or after code point offset `from`. Throws RegexError when the step limit is hit. */
  execAt(subject: number[], from: number, notEmptyAt = -1): { start: number; end: number; caps: number[] } | null {
    this.s = subject;
    this.notEmptyAt = notEmptyAt;
    this.searchStart = from;
    for (const r of this.required) {
      let found = false;
      for (let k = from; k < subject.length && !found; k++) found = r.i ? fullFold(subject[k]!).includes(r.c) : subject[k] === r.c;
      if (!found) return null;
    }
    for (let st = from; st <= subject.length; st++) {
      if (this.anchored && st > 0) return null;
      this.caps = new Array<number>(2 * (this.groupCount + 1)).fill(-1);
      this.regs = new Array<number>(this.regCount).fill(-1);
      this.skipTo = -1;
      let r: number;
      try { r = this.run(0, st, null, -1, 'top'); }
      catch (e) { if (e instanceof StepLimit) throw new RegexError('match limit exceeded'); throw e; }
      if (r >= 0) {
        const start = this.caps[0]!;
        const caps = this.caps.slice();
        caps[1] = r;
        return { start, end: r, caps };
      }
      if (r === ABORT_ALL) return null;
      if (r === ABORT && this.skipTo > st) st = this.skipTo - 1;
    }
    return null;
  }

  private toMatch(subject: number[], m: { start: number; end: number; caps: number[] }): RegexMatch {
    const slice = (a: number, b: number): string => String.fromCodePoint(...subject.slice(a, b));
    const groups: (string | undefined)[] = [];
    for (let g = 1; g <= this.groupCount; g++) {
      const a = m.caps[2 * g]!, b = m.caps[2 * g + 1]!;
      groups.push(a >= 0 && b >= 0 ? slice(a, b) : undefined);
    }
    return { start: m.start, end: Math.max(m.start, m.end), groups, whole: m.end >= m.start ? slice(m.start, m.end) : '' };
  }

  /** First match in `text`, or null. */
  exec(text: string): RegexMatch | null {
    this.steps = 0;
    const subject = toCodePoints(text);
    const m = this.execAt(subject, 0);
    return m ? this.toMatch(subject, m) : null;
  }

  test(text: string): boolean { return this.exec(text) !== null; }

  /** Every non-overlapping match, left to right, with the usual rule for empty matches. */
  matchAll(text: string): RegexMatch[] {
    this.steps = 0;
    const subject = toCodePoints(text);
    const out: RegexMatch[] = [];
    let from = 0, notEmpty = -1;
    while (from <= subject.length) {
      const m = this.execAt(subject, from, notEmpty);
      if (!m) break;
      out.push(this.toMatch(subject, m));
      const end = Math.max(m.end, m.start);
      if (end === m.start) notEmpty = end; else notEmpty = -1;
      from = end;
      if (out.length > 1_000_000) break;
    }
    return out;
  }
}

/** UTF-16 text as code points; lone surrogates are kept as their own values. */
export const toCodePoints = (s: string): number[] => {
  const out: number[] = [];
  for (const ch of s) out.push(ch.codePointAt(0)!);
  return out;
};

const cache = new Map<string, Regex | RegexError>();
/** Compiles (with a small cache). Throws RegexError for a bad pattern. */
export const compileRegex = (pattern: string, caseInsensitive = false): Regex => {
  const key = (caseInsensitive ? 'i:' : 's:') + pattern;
  let r = cache.get(key);
  if (!r) {
    try { r = new Regex(pattern, { caseInsensitive }); } catch (e) { if (e instanceof RegexError) r = e; else throw e; }
    if (cache.size > 500) cache.clear();
    cache.set(key, r);
  }
  if (r instanceof RegexError) throw r;
  return r;
};

/**
 * Expands a REGEXREPLACE replacement for one match: $$, $n, ${n}, $name, ${name}, $0 and $&.
 * Throws RegexError for a reference to a group that does not exist.
 */
export const expandReplacement = (repl: string, re: Regex, m: RegexMatch): string => {
  let out = '';
  for (let i = 0; i < repl.length; i++) {
    const c = repl[i]!;
    if (c !== '$') { out += c; continue; }
    const rest = repl.slice(i + 1);
    let mm: RegExpExecArray | null;
    const group = (g: number): string => {
      if (g === 0) return m.whole;
      if (g > re.groupCount) throw new RegexError('reference to non-existent group in replacement');
      return m.groups[g - 1] ?? '';
    };
    const named = (name: string): string => {
      const ns = re.names.get(name);
      if (!ns) throw new RegexError(`unknown group name ${name} in replacement`);
      for (const g of ns) if (m.groups[g - 1] !== undefined) return m.groups[g - 1]!;
      return '';
    };
    if (rest.startsWith('$')) { out += '$'; i++; }
    else if (rest.startsWith('&')) { out += m.whole; i++; }
    else if ((mm = /^\{(\d+)\}/.exec(rest))) { out += group(+mm[1]!); i += mm[0].length; }
    else if ((mm = /^\{([A-Za-z_]\w*)\}/.exec(rest))) { out += named(mm[1]!); i += mm[0].length; }
    else if ((mm = /^<([A-Za-z_]\w*)>/.exec(rest))) { out += named(mm[1]!); i += mm[0].length; }
    else if ((mm = /^\d+/.exec(rest))) {
      // Take the longest digit run that names an existing group (so $10 is group 1 then "0" when there are fewer than 10 groups).
      let d = mm[0];
      while (d.length > 1 && +d > re.groupCount) d = d.slice(0, -1);
      out += group(+d);
      i += d.length;
    } else if ((mm = /^[A-Za-z_]\w*/.exec(rest)) && re.names.has(mm[0])) { out += named(mm[0]); i += mm[0].length; }
    else out += '$';
  }
  return out;
};
