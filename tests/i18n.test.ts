import { describe, expect, it } from 'vitest';
import { en } from '../src/ui/i18n/en';
import { formatMessage, LANGS, matchLang, pluralCases, tr, N, useCatalogForTest, t, type Catalog } from '../src/ui/i18n';
import { CATALOG } from '../src/engine/catalog';
import { layoutOf } from '../src/app/layout';
import { newSheet } from '../src/model/workbook';
import { anchorBox, firstStrong, generalAlign, logicalDx, mirrorX, paneAt, panesL, panesOf, rectIn, screenToSheet, sheetToScreen, type Viewport } from '../src/ui/render';
import { arrowDelta } from '../src/ui/keys';

const langs = import.meta.glob<{ default: Catalog }>('../src/ui/i18n/lang/*.ts', { eager: true });
const fns = import.meta.glob<{ default: Record<string, string> }>('../src/ui/i18n/fn/*.ts', { eager: true });
const keys = Object.keys(en) as (keyof typeof en)[];

/** Top-level {names} of a message (plural argument names included, plural bodies not). */
const params = (msg: string): string[] => {
  const out = new Set<string>();
  let depth = 0;
  for (let i = 0; i < msg.length; i++) {
    if (msg[i] === '{') { if (depth === 0) { const m = /^\{\s*(\w+)\s*[,}]/.exec(msg.slice(i)); if (m) out.add(m[1]!); } depth++; }
    else if (msg[i] === '}') depth = Math.max(0, depth - 1);
  }
  return [...out].sort();
};

/** Every plural block in a message, with its body. */
const plurals = (msg: string): string[] => {
  const out: string[] = [];
  const re = /\{\s*\w+\s*,\s*plural\s*,/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(msg))) {
    let depth = 1, i = m.index + m[0].length;
    const from = i;
    while (i < msg.length && depth) { if (msg[i] === '{') depth++; else if (msg[i] === '}') depth--; i++; }
    out.push(msg.slice(from, i - 1));
  }
  return out;
};

describe('interface languages', () => {
  it('lists 19 languages with a file for each one but English', () => {
    expect(LANGS.map((l) => l.code).sort()).toEqual(['ar', 'de', 'en', 'es', 'fr', 'he', 'hi', 'id', 'it', 'ja', 'ko', 'nl', 'pl', 'pt', 'ru', 'sv', 'tr', 'uk', 'zh'].sort());
    for (const l of LANGS) if (l.code !== 'en') expect(langs[`../src/ui/i18n/lang/${l.code}.ts`], l.code).toBeDefined();
    expect(LANGS.filter((l) => l.rtl).map((l) => l.code).sort()).toEqual(['ar', 'he']);
  });

  for (const l of LANGS.filter((x) => x.code !== 'en')) {
    describe(l.code, () => {
      const cat = langs[`../src/ui/i18n/lang/${l.code}.ts`]!.default;
      const allowed = new Set(new Intl.PluralRules(l.code).resolvedOptions().pluralCategories as string[]);
      it('has every key and no others', () => {
        expect(Object.keys(cat).sort()).toEqual([...keys].sort());
      });
      it('has no empty text and keeps the same {params}', () => {
        for (const k of keys) {
          expect(cat[k]?.trim(), `${l.code}: ${k}`).toBeTruthy();
          expect(params(cat[k]), `${l.code}: ${k}`).toEqual(params(k));
        }
      });
      it('uses valid plural forms', () => {
        for (const k of keys) {
          const src = plurals(k).length, ours = plurals(cat[k]);
          expect(ours.length, `${l.code}: ${k}`).toBe(src);
          for (const body of ours) {
            const cases = pluralCases(body);
            expect(cases, `${l.code}: ${k}`).not.toBeNull();
            for (const c of Object.keys(cases!)) if (!c.startsWith('=')) expect(allowed.has(c), `${l.code}: "${c}" in ${k}`).toBe(true);
          }
        }
      });
      it('is mostly translated', () => {
        const same = keys.filter((k) => cat[k] === k && /[a-z]{3}/.test(k)).length;
        expect(same / keys.length).toBeLessThan(0.15);
      });
    });
  }

  it('English plural messages are well formed', () => {
    for (const k of keys) for (const body of plurals(en[k])) expect(pluralCases(body), k).not.toBeNull();
  });

  it('function descriptions are keyed by catalog text', () => {
    const texts = new Set<string>();
    for (const f of CATALOG) { texts.add(f.about); if (f.returns) texts.add(f.returns); for (const a of f.args) texts.add(a.about); }
    for (const [file, mod] of Object.entries(fns)) for (const [k, v] of Object.entries(mod.default)) {
      expect(texts.has(k), `${file}: ${k}`).toBe(true);
      expect(v.trim(), `${file}: ${k}`).toBeTruthy();
    }
  });
});

describe('messages', () => {
  it('fills values and picks plural forms per language', () => {
    expect(formatMessage('{n, plural, one {# row} other {# rows}} in {name}', { n: 1, name: 'A' }, 'en')).toBe('1 row in A');
    expect(formatMessage('{n, plural, one {# row} other {# rows}}', { n: 5 }, 'en')).toBe('5 rows');
    const ar = '{n, plural, zero {لا} one {واحد} two {اثنان} few {# قليل} many {# كثير} other {# آخر}}';
    expect([0, 1, 2, 3, 11, 100].map((n) => formatMessage(ar, { n }, 'ar'))).toEqual(['لا', 'واحد', 'اثنان', '3 قليل', '11 كثير', '100 آخر']);
    const ru = '{n, plural, one {# строка} few {# строки} many {# строк} other {# строки}}';
    expect([1, 2, 5, 21].map((n) => formatMessage(ru, { n }, 'ru'))).toEqual(['1 строка', '2 строки', '5 строк', '21 строка']);
    expect(formatMessage('{n, plural, =0 {none} other {#}}', { n: 0 }, 'en')).toBe('none');
    expect(formatMessage('keep {unknown}', {}, 'en')).toBe('keep {unknown}');
  });

  it('translates keys, passes other text through and switches catalogues', () => {
    const key = N('Insert sheet');
    expect(tr(key)).toBe('Insert sheet');
    expect(tr('Not a key 123')).toBe('Not a key 123');
    const de = langs['../src/ui/i18n/lang/de.ts']!.default;
    useCatalogForTest('de', de);
    expect(t('Insert sheet')).toBe(de['Insert sheet']);
    expect(tr('Insert sheet')).toBe(de['Insert sheet']);
    useCatalogForTest('en', null);
    expect(t('Insert sheet')).toBe('Insert sheet');
  });

  it('matches browser languages', () => {
    expect(matchLang('pt-BR')).toBe('pt');
    expect(matchLang('zh-CN')).toBe('zh');
    expect(matchLang('iw')).toBe('he');
    expect(matchLang('in-ID')).toBe('id');
    expect(matchLang('xx')).toBe('en');
    expect(matchLang(undefined)).toBe('en');
  });
});

describe('right-to-left sheets', () => {
  const vpOf = (rtl: boolean, extra: Partial<Viewport> = {}): Viewport => {
    const s = newSheet('S');
    return { layout: layoutOf(s, 1), zoom: 1, headerW: 40, headerH: 20, width: 800, height: 400, scrollX: 0, scrollY: 0, frozenRows: 0, frozenCols: 0, rtl, ...extra };
  };

  it('puts column A at the right edge, next to the row headers', () => {
    const vp = vpOf(true);
    const p = panesOf(vp)[0]!;
    expect(p.m).toBe(true);
    const a1 = rectIn(p, vp, { r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(a1.x + a1.w).toBe(800 - 40);
    const b1 = rectIn(p, vp, { r1: 0, c1: 1, r2: 0, c2: 1 });
    expect(b1.x + b1.w).toBe(a1.x);
    const ltr = vpOf(false);
    const l = rectIn(panesOf(ltr)[0]!, ltr, { r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(a1.x).toBe(mirrorX(vp, l.x, l.w));
  });

  it('mirrors frozen panes and keeps them next to the headers', () => {
    const vp = vpOf(true, { frozenRows: 1, frozenCols: 2, scrollX: 300 });
    const screen = panesOf(vp), logical = panesL(vp);
    expect(screen.length).toBe(4);
    for (let i = 0; i < screen.length; i++) expect(screen[i]!.x).toBe(800 - logical[i]!.x - logical[i]!.w);
    const frozen = screen.find((p) => p.c2 < 2 && p.r2 >= 1)!;
    expect(frozen.x + frozen.w).toBe(760);
  });

  it('maps screen points to cells and back', () => {
    for (const rtl of [false, true]) {
      const vp = vpOf(rtl, { scrollX: 130, scrollY: 47, frozenRows: 2, frozenCols: 1 });
      for (const [sx, sy] of [[10, 5], [500, 300], [200, 30]] as const) {
        const [x, y] = sheetToScreen(vp, sx, sy);
        const [bx, by] = screenToSheet(vp, x, y);
        expect(bx).toBeCloseTo(sx, 6);
        expect(by).toBeCloseTo(sy, 6);
      }
    }
    const vp = vpOf(true);
    const [sx] = screenToSheet(vp, 800 - 40 - 1, 30);
    expect(vp.layout.cols.at(sx)).toBe(0);
    const [sx2] = screenToSheet(vp, 800 - 40 - 64 - 2, 30);
    expect(vp.layout.cols.at(sx2)).toBe(1);
  });

  it('finds the pane under a screen point', () => {
    const vp = vpOf(true, { frozenCols: 1 });
    const panes = panesOf(vp);
    const p = paneAt(panes, vp, 790 - 40, 100);
    expect(p.c2).toBe(0);
    expect(paneAt(panes, vp, 100, 100).c1).toBeGreaterThan(0);
  });

  it('places objects mirrored and turns screen drags into sheet moves', () => {
    const a = { row: 0, col: 0, dx: 10, dy: 5, width: 100, height: 50 };
    const ltr = anchorBox(vpOf(false), a), rtl = anchorBox(vpOf(true), a);
    expect(rtl.left).toBe(800 - ltr.left - ltr.width);
    expect(rtl.top).toBe(ltr.top);
    expect(logicalDx({ rtl: false }, 20, 0)).toBe(20);
    expect(logicalDx({ rtl: true }, 20, 0)).toBe(-20);
    expect(logicalDx({ rtl: true }, 0, 30)).toBe(-30);
  });

  it('swaps the left and right arrows', () => {
    expect(arrowDelta('ArrowRight', false)).toEqual([0, 1]);
    expect(arrowDelta('ArrowLeft', false)).toEqual([0, -1]);
    expect(arrowDelta('ArrowLeft', true)).toEqual([0, 1]);
    expect(arrowDelta('ArrowRight', true)).toEqual([0, -1]);
    expect(arrowDelta('ArrowDown', true)).toEqual([1, 0]);
    expect(arrowDelta('a', true)).toBeUndefined();
  });

  it('aligns General text by its own direction and numbers by the sheet', () => {
    expect(firstStrong('مرحبا 123')).toBe('rtl');
    expect(firstStrong('123 שלום')).toBe('rtl');
    expect(firstStrong('Total 5')).toBe('ltr');
    expect(firstStrong('123 ...')).toBeNull();
    expect(generalAlign('مرحبا', 'مرحبا', false, false)).toBe('right');
    expect(generalAlign('Hello', 'Hello', true, false)).toBe('left');
    expect(generalAlign(5, '5', false, false)).toBe('right');
    expect(generalAlign(5, '5', true, false)).toBe('left');
    expect(generalAlign('123', '123', true, false)).toBe('right');
    expect(generalAlign('Hello', 'Hello', false, true)).toBe('right');
  });
});
