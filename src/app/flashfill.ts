// Flash Fill: learns a text transformation from a few example rows and applies it to the others.
//
// The approach is program synthesis by example, after Gulwani, "Automating String Processing in
// Spreadsheets Using Input-Output Examples" (POPL 2011): an output is a concatenation of atoms, and each
// atom is a constant, a substring of an input whose two ends are given by token-relative positions (words,
// digit runs, case runs, punctuation...), possibly with a case change, or one of a few richer atoms
// (number reformatting, date parts, initials). For every example we build the graph of all ways to
// produce its output; the graphs are intersected so only programs that explain every example survive,
// and the cheapest path wins. Costs prefer token-based positions over fixed character offsets, few atoms
// over many, and atoms that also work on the rows still to be filled (multi-example disambiguation).

export type CaseOp = 'id' | 'up' | 'lo' | 'pr' | 'cap';
const CASES: readonly CaseOp[] = ['id', 'up', 'lo', 'pr', 'cap'];
const CASE_COST: Record<CaseOp, number> = { id: 0, up: 0.3, lo: 0.3, pr: 0.35, cap: 0.4 };

const proper = (s: string): string => s.toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase());
export const applyCase = (s: string, c: CaseOp): string => {
  switch (c) {
    case 'up': return s.toUpperCase();
    case 'lo': return s.toLowerCase();
    case 'pr': return proper(s);
    case 'cap': return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
    default: return s;
  }
};

// ---------- tokens and positions ----------

const CLASSES: Record<string, { re: RegExp; cost: number }> = {
  w: { re: /\p{L}+/gu, cost: 0 },
  d: { re: /\p{Nd}+/gu, cost: 0 },
  a: { re: /[\p{L}\p{Nd}]+/gu, cost: 0.1 },
  n: { re: /\S+/gu, cost: 0.1 },
  c: { re: /[^\s,;]+/gu, cost: 0.15 },
  p: { re: /[^\p{L}\p{Nd}\s]+/gu, cost: 0.2 },
  U: { re: /\p{Lu}[\p{Ll}\p{Nd}]*/gu, cost: 0.3 },
  u: { re: /\p{Lu}+/gu, cost: 0.35 },
  l: { re: /\p{Ll}+/gu, cost: 0.35 },
  s: { re: /\s+/gu, cost: 0.3 },
};
const CLASS_KEYS = Object.keys(CLASSES);

const tokenCache = new Map<string, Map<string, [number, number][]>>();
const tokens = (s: string, cls: string): [number, number][] => {
  let byCls = tokenCache.get(s);
  if (!byCls) {
    if (tokenCache.size > 5000) tokenCache.clear();
    byCls = new Map();
    tokenCache.set(s, byCls);
  }
  let t = byCls.get(cls);
  if (!t) {
    t = [];
    const re = new RegExp(CLASSES[cls]!.re.source, 'gu');
    for (let m = re.exec(s); m; m = re.exec(s)) t.push([m.index, m.index + m[0].length]);
    byCls.set(cls, t);
  }
  return t;
};

/** Every way to name position `a` of `s`, with its cost. */
const posExprs = (s: string, a: number): Map<string, number> => {
  const out = new Map<string, number>();
  out.set(`^${a}`, a === 0 ? 0 : 2 + 0.05 * a);
  out.set(`$${s.length - a}`, a === s.length ? 0 : 2.2 + 0.05 * (s.length - a));
  for (const cls of CLASS_KEYS) {
    const ts = tokens(s, cls);
    const base = CLASSES[cls]!.cost;
    ts.forEach(([st, en], i) => {
      if (a < st - 0 || a > en) return;
      const n = ts.length;
      for (const [side, off] of [['S', a - st], ['E', a - en]] as const) {
        if (Math.abs(off) > 40) continue;
        const c = base + (off === 0 ? 0 : 0.7 + 0.04 * Math.abs(off));
        out.set(`${cls}:${i + 1}:${side}:${off}`, c);
        out.set(`${cls}:${i - n}:${side}:${off}`, c + 0.05);
      }
    });
  }
  return out;
};

const evalPos = (key: string, s: string): number | null => {
  if (key[0] === '^') { const n = +key.slice(1); return n <= s.length ? n : null; }
  if (key[0] === '$') { const n = +key.slice(1); return n <= s.length ? s.length - n : null; }
  const [cls, idxS, side, offS] = key.split(':') as [string, string, string, string];
  const ts = tokens(s, cls);
  const idx = +idxS;
  const t = idx > 0 ? ts[idx - 1] : ts[ts.length + idx];
  if (!t) return null;
  const p = (side === 'S' ? t[0] : t[1]) + +offS;
  return p >= 0 && p <= s.length ? p : null;
};

// ---------- dates and numbers ----------

const EN_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const EN_DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

export interface FlashOptions {
  /** Order of day, month and year in dates written with numbers only. */
  dateOrder?: 'MDY' | 'DMY' | 'YMD';
  months?: readonly string[];
  monthsShort?: readonly string[];
  days?: readonly string[];
  daysShort?: readonly string[];
}

interface DateParts { y: number; m: number; d: number; wd: number }
const monthIndex = (name: string, opt: FlashOptions): number => {
  const n = name.toLowerCase().replace(/\.$/, '');
  const lists = [opt.months, opt.monthsShort, EN_MONTHS, EN_MONTHS.map((x) => x.slice(0, 3))];
  for (const l of lists) {
    if (!l) continue;
    const i = l.findIndex((x) => x.toLowerCase() === n || (n.length >= 3 && x.toLowerCase().startsWith(n) && l === EN_MONTHS));
    if (i >= 0) return i + 1;
  }
  return 0;
};
const mkDate = (y: number, m: number, d: number): DateParts | null => {
  if (y < 100) y += y < 30 ? 2000 : 1900;
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCMonth() !== m - 1) return null;
  return { y, m, d, wd: dt.getUTCDay() };
};
const dateCache = new Map<string, DateParts | null>();
export const parseLooseDate = (s: string, opt: FlashOptions = {}): DateParts | null => {
  const key = `${opt.dateOrder ?? 'MDY'}|${s}`;
  if (dateCache.has(key)) return dateCache.get(key)!;
  if (dateCache.size > 5000) dateCache.clear();
  const t = s.trim();
  let r: DateParts | null = null;
  let m: RegExpExecArray | null;
  if ((m = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:[T ].*)?$/.exec(t))) r = mkDate(+m[1]!, +m[2]!, +m[3]!);
  else if ((m = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})(?: .*)?$/.exec(t))) {
    const order = opt.dateOrder ?? 'MDY';
    r = order === 'DMY' ? mkDate(+m[3]!, +m[2]!, +m[1]!) : mkDate(+m[3]!, +m[1]!, +m[2]!);
  } else if ((m = /^(?:\p{L}+,?\s+)?(\p{L}+)\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})$/u.exec(t))) {
    const mi = monthIndex(m[1]!, opt);
    if (mi) r = mkDate(+m[3]!, mi, +m[2]!);
  } else if ((m = /^(?:\p{L}+,?\s+)?(\d{1,2})(?:st|nd|rd|th)?\.?\s+(\p{L}+)\.?,?\s+(\d{4})$/u.exec(t))) {
    const mi = monthIndex(m[2]!, opt);
    if (mi) r = mkDate(+m[3]!, mi, +m[1]!);
  } else if ((m = /^(\d{1,2})[-\s](\p{L}{3,})[-\s](\d{2,4})$/u.exec(t))) {
    const mi = monthIndex(m[2]!, opt);
    if (mi) r = mkDate(+m[3]!, mi, +m[1]!);
  }
  dateCache.set(key, r);
  return r;
};

const DATE_PARTS = ['yyyy', 'yy', 'm', 'mm', 'mmm', 'mmmm', 'd', 'dd', 'ddd', 'dddd', 'q'] as const;
const datePart = (p: DateParts, part: string, opt: FlashOptions): string => {
  const months = opt.months ?? EN_MONTHS, days = opt.days ?? EN_DAYS;
  switch (part) {
    case 'yyyy': return String(p.y);
    case 'yy': return String(p.y % 100).padStart(2, '0');
    case 'm': return String(p.m);
    case 'mm': return String(p.m).padStart(2, '0');
    case 'mmm': return (opt.monthsShort ?? months.map((x) => x.slice(0, 3)))[p.m - 1]!;
    case 'mmmm': return months[p.m - 1]!;
    case 'd': return String(p.d);
    case 'dd': return String(p.d).padStart(2, '0');
    case 'q': return String(Math.ceil(p.m / 3));
    case 'ddd': return (opt.daysShort ?? days.map((x) => x.slice(0, 3)))[p.wd]!;
    default: return days[p.wd]!;
  }
};

const NUM_RE = /-?\d[\d,]*(?:\.\d+)?|-?\.\d+/g;
const numbersIn = (s: string): { v: number; text: string }[] => {
  const out: { v: number; text: string }[] = [];
  for (const m of s.matchAll(NUM_RE)) {
    const txt = m[0];
    if (txt.includes(',') && !/^-?\d{1,3}(,\d{3})+(\.\d+)?$/.test(txt)) { for (const part of txt.split(',')) if (part) out.push({ v: Number(part), text: part }); continue; }
    out.push({ v: Number(txt.replace(/,/g, '')), text: txt });
  }
  return out.filter((x) => Number.isFinite(x.v));
};
const groupThousands = (s: string): string => s.replace(/^(-?\d+)/, (w) => w.replace(/\B(?=(\d{3})+(?!\d))/g, ','));
const NUM_SPECS = (o: string): string[] => {
  const specs: string[] = [];
  if (/^-?\d+$/.test(o)) specs.push(`pad${o.replace('-', '').length}`, 'int', 'round');
  if (/^-?\d+(\.\d+)?$/.test(o)) specs.push('gen');
  const dec = /^-?[\d,]*\.(\d+)$/.exec(o);
  if (dec) specs.push(`fix${dec[1]!.length}`, `grp${dec[1]!.length}`);
  if (/^-?\d{1,3}(,\d{3})+$/.test(o)) specs.push('grp0');
  return specs;
};
const fmtNum = (v: number, spec: string): string => {
  if (spec === 'int') return String(Math.trunc(v));
  if (spec === 'gen') return String(Number(v.toPrecision(15)));
  if (spec === 'round') return String(Math.round(v));
  if (spec.startsWith('pad')) { if (!Number.isInteger(v)) return ''; const w = +spec.slice(3); return (v < 0 ? '-' : '') + String(Math.abs(v)).padStart(w, '0'); }
  if (spec.startsWith('fix')) return v.toFixed(+spec.slice(3));
  if (spec.startsWith('grp')) return groupThousands(v.toFixed(+spec.slice(3)));
  return '';
};

// ---------- atoms ----------

interface SubAtom { t: 's'; k: number; c: CaseOp; pa: Map<string, number>; pb: Map<string, number> }
interface KeyAtom { t: 'k'; key: string; cost: number }
type Atom = SubAtom | KeyAtom;

/** A chosen atom of the final program. */
type Step =
  | { t: 'c'; s: string }
  | { t: 's'; k: number; c: CaseOp; a: string; b: string }
  | { t: 'n'; k: number; idx: number; spec: string }
  | { t: 'd'; k: number; part: string; c: CaseOp }
  | { t: 'i'; k: number; cls: string; sep: string; trail: boolean; c: CaseOp };

const isAlnum = (s: string): boolean => /[\p{L}\p{Nd}]/u.test(s);
const constCost = (s: string): number => (isAlnum(s) ? 3.2 + 0.6 * s.length : 0.6 + 0.15 * s.length);

const keyStep = (key: string): Step => {
  const parts = key.split('\u0001');
  switch (parts[0]) {
    case 'c': return { t: 'c', s: parts[1]! };
    case 'n': return { t: 'n', k: +parts[1]!, idx: +parts[2]!, spec: parts[3]! };
    case 'd': return { t: 'd', k: +parts[1]!, part: parts[2]!, c: parts[3] as CaseOp };
    default: return { t: 'i', k: +parts[1]!, cls: parts[2]!, sep: parts[3]!, trail: parts[4] === '1', c: parts[5] as CaseOp };
  }
};
const K = (...p: (string | number)[]): string => p.join('\u0001');

const initialsOf = (s: string, cls: string, sep: string, trail: boolean): string | null => {
  const ts = tokens(s, cls);
  if (ts.length < 1) return null;
  const letters = ts.map(([a]) => s.charAt(a));
  const tail = trail ? (sep.trim() || sep) : '';
  return letters.join(sep) + tail;
};

export const runStep = (st: Step, inputs: readonly string[], opt: FlashOptions = {}): string | null => {
  if (st.t === 'c') return st.s;
  const s = inputs[st.k];
  if (s === undefined) return null;
  switch (st.t) {
    case 's': {
      const a = evalPos(st.a, s), b = evalPos(st.b, s);
      if (a === null || b === null || b < a) return null;
      return applyCase(s.slice(a, b), st.c);
    }
    case 'n': {
      const nums = numbersIn(s);
      const x = st.idx > 0 ? nums[st.idx - 1] : nums[nums.length + st.idx];
      if (!x) return null;
      const r = fmtNum(x.v, st.spec);
      return r === '' ? null : r;
    }
    case 'd': {
      const p = parseLooseDate(s, opt);
      return p ? applyCase(datePart(p, st.part, opt), st.c) : null;
    }
    case 'i': {
      const r = initialsOf(s, st.cls, st.sep, st.trail);
      return r === null ? null : applyCase(r, st.c);
    }
  }
};

/** All atoms that produce out[i:j] for one example, indexed by start position and signature. */
const buildGraph = (inputs: readonly string[], out: string, opt: FlashOptions): Map<string, { j: number; atom: Atom }[]>[] => {
  const L = out.length;
  const idx: Map<string, { j: number; atom: Atom }[]>[] = Array.from({ length: L + 1 }, () => new Map());
  const add = (i: number, j: number, sig: string, atom: Atom): void => {
    const list = idx[i]!.get(sig);
    if (list) list.push({ j, atom }); else idx[i]!.set(sig, [{ j, atom }]);
  };
  const posMemo = new Map<string, Map<string, number>>();
  const pos = (k: number, a: number): Map<string, number> => {
    const key = `${k}:${a}`;
    let m = posMemo.get(key);
    if (!m) { m = posExprs(inputs[k]!, a); posMemo.set(key, m); }
    return m;
  };
  const lowers = inputs.map((s) => s.toLowerCase());
  const dates = inputs.map((s) => parseLooseDate(s, opt));
  const nums = inputs.map(numbersIn);
  for (let i = 0; i < L; i++) {
    for (let j = i + 1; j <= L; j++) {
      const o = out.slice(i, j);
      const ckey = K('c', o);
      add(i, j, ckey, { t: 'k', key: ckey, cost: constCost(o) });
      const lo = o.toLowerCase();
      if (lo.length !== o.length) continue;
      inputs.forEach((s, k) => {
        const low = lowers[k]!;
        for (let a = low.indexOf(lo); a >= 0; a = low.indexOf(lo, a + 1)) {
          const sub = s.slice(a, a + o.length);
          for (const c of CASES) {
            if (applyCase(sub, c) !== o) continue;
            if (c === 'cap' && applyCase(sub, 'pr') === o) continue;
            add(i, j, `s:${k}:${c}`, { t: 's', k, c, pa: pos(k, a), pb: pos(k, a + o.length) });
          }
        }
        // Richer atoms.
        const specs = NUM_SPECS(o);
        if (specs.length) nums[k]!.forEach((x, n) => {
          for (const spec of specs) {
            if (fmtNum(x.v, spec) !== o) continue;
            if (spec.startsWith('pad') && x.text === o) continue;
            for (const id of [n + 1, n - nums[k]!.length]) { const key = K('n', k, id, spec); add(i, j, key, { t: 'k', key, cost: 1.5 + (id < 0 ? 0.05 : 0) }); }
          }
        });
        const d = dates[k];
        if (d) for (const part of DATE_PARTS) {
          const v = datePart(d, part, opt);
          for (const c of ['id', 'up', 'lo'] as CaseOp[]) {
            if (applyCase(v, c) !== o || (c !== 'id' && v === o)) continue;
            const key = K('d', k, part, c);
            add(i, j, key, { t: 'k', key, cost: (part.length > 2 && /[md]{3}/.test(part) ? 0.9 : 1.1) + CASE_COST[c] });
          }
        }
        if (o.length >= 2) for (const cls of ['w', 'n']) {
          if (tokens(s, cls).length < 2) continue;
          for (const sep of ['', '.', ' ', '. ', '-']) for (const trail of [false, true]) {
            if (trail && sep === '') continue;
            const r = initialsOf(s, cls, sep, trail);
            if (r === null) continue;
            for (const c of ['id', 'up', 'lo'] as CaseOp[]) {
              if (applyCase(r, c) !== o) continue;
              const key = K('i', k, cls, sep, trail ? 1 : 0, c);
              add(i, j, key, { t: 'k', key, cost: 1.7 + CASE_COST[c] + (cls === 'n' ? 0.1 : 0) });
            }
          }
        }
      });
    }
  }
  return idx;
};

const intersectMaps = (a: Map<string, number>, b: Map<string, number>): Map<string, number> => {
  const out = new Map<string, number>();
  const [small, big] = a.size <= b.size ? [a, b] : [b, a];
  for (const [k, v] of small) { const w = big.get(k); if (w !== undefined) out.set(k, Math.max(v, w)); }
  return out;
};

const intersectAtom = (x: Atom, y: Atom): Atom | null => {
  if (x.t === 'k' || y.t === 'k') return x.t === 'k' && y.t === 'k' && x.key === y.key ? x : null;
  if (x.k !== y.k || x.c !== y.c) return null;
  const pa = intersectMaps(x.pa, y.pa);
  if (!pa.size) return null;
  const pb = intersectMaps(x.pb, y.pb);
  if (!pb.size) return null;
  return { t: 's', k: x.k, c: x.c, pa, pb };
};

const sigOf = (a: Atom): string => (a.t === 'k' ? a.key : `s:${a.k}:${a.c}`);

/** Best concrete step for an atom, considering how well it works on the unlabelled rows. */
const concretize = (atom: Atom, unlabeled: readonly (readonly string[])[], opt: FlashOptions): { step: Step; cost: number } | null => {
  const failRate = (st: Step): number => {
    if (!unlabeled.length) return 0;
    let bad = 0;
    for (const u of unlabeled) { const r = runStep(st, u, opt); if (r === null || (st.t === 's' && r === '')) bad++; }
    return bad / unlabeled.length;
  };
  if (atom.t === 'k') {
    const step = keyStep(atom.key);
    return { step, cost: atom.cost + (step.t === 'c' ? 0 : 4 * failRate(step)) };
  }
  const rank = (m: Map<string, number>): [string, number][] => {
    const scored = [...m].map(([key, c]): [string, number] => {
      if (!unlabeled.length) return [key, c];
      let bad = 0;
      for (const u of unlabeled) if (u[atom.k] === undefined || evalPos(key, u[atom.k]!) === null) bad++;
      return [key, c + 4 * (bad / unlabeled.length)];
    });
    return scored.sort((p, q) => p[1] - q[1] || p[0].localeCompare(q[0])).slice(0, 4);
  };
  const as = rank(atom.pa), bs = rank(atom.pb);
  let best: { step: Step; cost: number } | null = null;
  for (const [a, ca] of as) for (const [b, cb] of bs) {
    const step: Step = { t: 's', k: atom.k, c: atom.c, a, b };
    const fr = failRate(step);
    const cost = ca + cb + CASE_COST[atom.c] + 0.5 + 4 * fr;
    if (!best || cost < best.cost) best = { step, cost };
  }
  return best;
};

export interface FlashExample { inputs: readonly string[]; output: string }

export interface FlashProgram {
  steps: Step[];
  cost: number;
  run(inputs: readonly string[]): string | null;
}

const MAX_EXPANSIONS = 60_000;

const search = (examples: readonly FlashExample[], unlabeled: readonly (readonly string[])[], opt: FlashOptions): FlashProgram | null => {
  const graphs = examples.map((e) => buildGraph(e.inputs, e.output, opt));
  const ends = examples.map((e) => e.output.length);
  const E = examples.length;
  const start = new Array<number>(E).fill(0);
  const keyOf = (st: number[]): string => st.join(',');
  const dist = new Map<string, number>([[keyOf(start), 0]]);
  const prev = new Map<string, { from: string; step: Step }>();
  const states = new Map<string, number[]>([[keyOf(start), start]]);
  const done = new Set<string>();
  const heap: [number, string][] = [[0, keyOf(start)]];
  const concreteMemo = new Map<string, { step: Step; cost: number } | null>();
  let expansions = 0;
  const goal = keyOf(ends);
  while (heap.length) {
    let bi = 0;
    for (let i = 1; i < heap.length; i++) if (heap[i]![0] < heap[bi]![0]) bi = i;
    const [d, key] = heap[bi]!;
    heap[bi] = heap[heap.length - 1]!;
    heap.pop();
    if (done.has(key)) continue;
    done.add(key);
    if (key === goal) break;
    const st = states.get(key)!;
    // Combined edges from this state, found example by example through the signature index.
    const extend = (e: number, atom: Atom, next: number[]): void => {
      if (++expansions > MAX_EXPANSIONS) return;
      if (e === E) {
        const ck = `${sigOf(atom)}|${atom.t === 's' ? [...atom.pa.keys()].sort().join(';') + '#' + [...atom.pb.keys()].sort().join(';') : ''}`;
        let conc = concreteMemo.get(ck);
        if (conc === undefined) { conc = concretize(atom, unlabeled, opt); concreteMemo.set(ck, conc); }
        if (!conc) return;
        const nk = keyOf(next);
        const nd = d + conc.cost + 0.25;
        if (nd < (dist.get(nk) ?? Infinity)) {
          dist.set(nk, nd);
          states.set(nk, next);
          prev.set(nk, { from: key, step: conc.step });
          heap.push([nd, nk]);
        }
        return;
      }
      const cands = graphs[e]![st[e]!]!.get(sigOf(atom));
      if (!cands) return;
      for (const c of cands) {
        const x = intersectAtom(atom, c.atom);
        if (x) extend(e + 1, x, [...next, c.j]);
      }
    };
    for (const list of graphs[0]![st[0]!]!.values()) for (const c of list) extend(1, c.atom, [c.j]);
    if (expansions > MAX_EXPANSIONS) return null;
  }
  if (!prev.has(goal) && goal !== keyOf(start)) return null;
  const steps: Step[] = [];
  for (let k = goal; k !== keyOf(start);) { const p = prev.get(k)!; steps.unshift(p.step); k = p.from; }
  // Merge neighbouring constants.
  const merged: Step[] = [];
  for (const s of steps) {
    const last = merged[merged.length - 1];
    if (s.t === 'c' && last?.t === 'c') merged[merged.length - 1] = { t: 'c', s: last.s + s.s };
    else merged.push(s);
  }
  const prog: FlashProgram = {
    steps: merged,
    cost: dist.get(goal) ?? 0,
    run: (inputs) => {
      let r = '';
      for (const s of merged) { const v = runStep(s, inputs, opt); if (v === null) return null; r += v; }
      return r;
    },
  };
  return prog;
};

/**
 * Learns a program that turns each example's inputs into its output. `unlabeled` are the inputs of the
 * rows to be filled, used to prefer programs that work on them. Returns null when nothing fits.
 */
export const synthesize = (examples: readonly FlashExample[], unlabeled: readonly (readonly string[])[] = [], opt: FlashOptions = {}): FlashProgram | null => {
  const ex = examples.filter((e) => e.output !== '');
  if (!ex.length) return null;
  const sample = unlabeled.length > 60 ? unlabeled.filter((_, i) => i % Math.ceil(unlabeled.length / 60) === 0) : unlabeled;
  const verify = (p: FlashProgram | null): FlashProgram | null => (p && ex.every((e) => p.run(e.inputs) === e.output) ? p : null);
  // Shortest outputs first keep the product graph small; very many examples are learnt from a subset and checked.
  const ordered = [...ex].sort((a, b) => a.output.length - b.output.length);
  if (ordered.length <= 4) return verify(search(ordered, sample, opt));
  const p = verify(search(ordered.slice(0, 3), sample, opt));
  if (p) return p;
  return verify(search(ordered.slice(0, 6), sample, opt));
};
