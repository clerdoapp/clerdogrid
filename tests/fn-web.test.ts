import { afterEach, describe, expect, it } from 'vitest';
import { defs } from '../src/engine/functions/web';
import { cellImage, clearOnlineCache, setOnlineAllowed, settled } from '../src/engine/online';
import { toSerial } from '../src/engine/date';
import { CellError, type Value } from '../src/engine/types';
import { makeCtx } from './fnctx';

const ctx = makeCtx();
const call = (name: string, ...args: Value[]): Value => defs.find((d) => d.name === name)!.fn(args, ctx);
const code = (v: Value): string | null => (v instanceof CellError ? v.code : null);
const g = globalThis as unknown as { fetch?: unknown; Translator?: unknown; LanguageDetector?: unknown };
const realFetch = g.fetch;

afterEach(() => { setOnlineAllowed(false); clearOnlineCache(); g.fetch = realFetch; delete g.Translator; delete g.LanguageDetector; });

const mockFetch = (body: string, status = 200): string[] => {
  const seen: string[] = [];
  g.fetch = async (url: string) => { seen.push(url); return new Response(body, { status }); };
  return seen;
};

const XML = `<?xml version="1.0"?>
<library xmlns:x="urn:x">
  <book id="b1" year="1999"><title>Alpha</title><price>10.5</price><tag>a</tag><tag>b</tag></book>
  <book id="b2" year="2005"><title>Beta</title><price>20</price></book>
  <book id="b3" year="2012"><x:title>Gamma</x:title><price>7.25</price><tag>c</tag></book>
  <note>Hello &amp; welcome</note>
</library>`;

describe('FILTERXML', () => {
  it('paths, attributes and predicates', () => {
    expect(call('FILTERXML', XML, '//title')).toEqual([['Alpha'], ['Beta']]);
    expect(call('FILTERXML', XML, '//*[local-name()="title"]')).toEqual([['Alpha'], ['Beta'], ['Gamma']]);
    expect(call('FILTERXML', XML, '//x:*')).toBe('Gamma');
    expect(call('FILTERXML', XML, '/library/book[2]/title')).toBe('Beta');
    expect(call('FILTERXML', XML, '//book/@id')).toEqual([['b1'], ['b2'], ['b3']]);
    expect(call('FILTERXML', XML, '//book[@year>2000]/@id')).toEqual([['b2'], ['b3']]);
    expect(call('FILTERXML', XML, '//book[price<15]/@id')).toEqual([['b1'], ['b3']]);
    expect(call('FILTERXML', XML, '//book[last()]/@id')).toBe('b3');
    expect(call('FILTERXML', XML, '//book[position()=1 or @id="b3"]/@year')).toEqual([[1999], [2012]]);
    expect(call('FILTERXML', XML, '//book[tag]/@id')).toEqual([['b1'], ['b3']]);
    expect(call('FILTERXML', XML, '//book[not(tag)]/@id')).toBe('b2');
    expect(call('FILTERXML', XML, '//price')).toEqual([[10.5], [20], [7.25]]);
    expect(call('FILTERXML', XML, '//note')).toBe('Hello & welcome');
    expect(call('FILTERXML', XML, '//note/text()')).toBe('Hello & welcome');
    expect(call('FILTERXML', XML, '//x:title')).toBe('Gamma');
  });
  it('axes, unions and functions', () => {
    expect(call('FILTERXML', XML, '//title[.="Beta"]/../@id')).toBe('b2');
    expect(call('FILTERXML', XML, '//book[@id="b2"]/following-sibling::book/@id')).toBe('b3');
    expect(call('FILTERXML', XML, '//book[@id="b2"]/preceding-sibling::*[1]/@id')).toBe('b1');
    expect(call('FILTERXML', XML, '//tag[.="b"]/ancestor::book/@id')).toBe('b1');
    expect(call('FILTERXML', XML, '//book[1]/descendant::tag')).toEqual([['a'], ['b']]);
    expect(call('FILTERXML', XML, '//note | //book[1]/title')).toEqual([['Alpha'], ['Hello & welcome']]);
    expect(call('FILTERXML', XML, 'count(//book)')).toBe(3);
    expect(call('FILTERXML', XML, 'sum(//price)')).toBe(37.75);
    expect(call('FILTERXML', XML, 'concat(//book[1]/title, "-", //book[2]/title)')).toBe('Alpha-Beta');
    expect(call('FILTERXML', XML, '//book[starts-with(title, "Al")]/@id')).toBe('b1');
    expect(call('FILTERXML', XML, '//book[contains(@id, "3")]/price')).toBe(7.25);
    expect(call('FILTERXML', XML, 'substring-after(//note, "& ")')).toBe('welcome');
    expect(call('FILTERXML', XML, 'translate("abc", "abc", "ABC")')).toBe('ABC');
    expect(call('FILTERXML', XML, 'string-length(normalize-space("  a  b "))')).toBe(3);
    expect(call('FILTERXML', XML, '//book[floor(price) = 10]/@id')).toBe('b1');
    expect(call('FILTERXML', XML, '7 div 2 + 10 mod 4')).toBe(5.5);
    expect(call('FILTERXML', XML, 'name(/*)')).toBe('library');
    expect(call('FILTERXML', XML, 'boolean(//missing)')).toBe(false);
  });
  it('namespaces', () => {
    const NS = `<feed xmlns="urn:atom" xmlns:m="urn:media" xmlns:alias="urn:media">
      <entry><title>One</title><m:thumb m:w="64" w="9">t1</m:thumb></entry>
      <entry><title>Two</title><alias:thumb>t2</alias:thumb><x:extra xmlns:x="urn:x">e</x:extra></entry>
    </feed>`;
    // Both elements are in urn:media, whichever prefix the document used.
    expect(call('FILTERXML', NS, '//m:thumb')).toEqual([['t1'], ['t2']]);
    expect(call('FILTERXML', NS, '//alias:thumb')).toEqual([['t1'], ['t2']]);
    expect(call('FILTERXML', NS, '//m:thumb/@m:w')).toBe(64);
    expect(call('FILTERXML', NS, '//m:thumb/@w')).toBe(9);
    // A name without a prefix means no namespace, so the default namespace needs local-name().
    expect(code(call('FILTERXML', NS, '//title'))).toBe('#VALUE!');
    expect(call('FILTERXML', NS, '//*[local-name()="title"]')).toEqual([['One'], ['Two']]);
    expect(call('FILTERXML', NS, 'namespace-uri(/*)')).toBe('urn:atom');
    expect(call('FILTERXML', NS, 'namespace-uri(//m:thumb)')).toBe('urn:media');
    expect(call('FILTERXML', NS, 'local-name((//alias:thumb)[2])')).toBe('thumb');
    expect(call('FILTERXML', NS, 'name((//*[local-name()="thumb"])[2])')).toBe('alias:thumb');
    expect(call('FILTERXML', NS, '//x:extra')).toBe('e');
    expect(call('FILTERXML', NS, 'count(//m:*)')).toBe(2);
    expect(code(call('FILTERXML', NS, '//zz:thumb'))).toBe('#VALUE!');
    expect(code(call('FILTERXML', '<p:a/>', '/*'))).toBe('#VALUE!');
  });
  it('errors', () => {
    expect(code(call('FILTERXML', XML, '//missing'))).toBe('#VALUE!');
    expect(code(call('FILTERXML', XML, '//book['))).toBe('#VALUE!');
    expect(code(call('FILTERXML', 'not xml', '//a'))).toBe('#VALUE!');
    expect(code(call('FILTERXML', '<a><b></a>', '//a'))).toBe('#VALUE!');
    expect(call('FILTERXML', '<a>x<b>y</b>z<![CDATA[<w>]]></a>', '/a/text()')).toEqual([['x'], ['z<w>']]);
    expect(call('FILTERXML', '<a>x<b>y</b>z</a>', 'string(/a)')).toBe('xyz');
    expect(code(call('FILTERXML', XML, 'nosuch()'))).toBe('#VALUE!');
  });
});

describe('WEBSERVICE', () => {
  it('is offline unless allowed', () => {
    const seen = mockFetch('hi');
    expect(code(call('WEBSERVICE', 'https://example.com/a'))).toBe('#CONNECT!');
    expect(seen).toEqual([]);
  });
  it('fetches https text in the background', async () => {
    setOnlineAllowed(true);
    const seen = mockFetch('<rate>1.25</rate>');
    expect(code(call('WEBSERVICE', 'https://example.com/rate'))).toBe('#BUSY!');
    await settled();
    expect(call('WEBSERVICE', 'https://example.com/rate')).toBe('<rate>1.25</rate>');
    expect(seen).toEqual(['https://example.com/rate']);
    expect(call('FILTERXML', call('WEBSERVICE', 'https://example.com/rate'), '//rate')).toBe(1.25);
  });
  it('refuses other schemes, failures and huge answers', async () => {
    setOnlineAllowed(true);
    expect(code(call('WEBSERVICE', 'http://example.com'))).toBe('#VALUE!');
    expect(code(call('WEBSERVICE', 'file:///etc/passwd'))).toBe('#VALUE!');
    expect(code(call('WEBSERVICE', 'https://user:pw@example.com'))).toBe('#VALUE!');
    mockFetch('nope', 404);
    call('WEBSERVICE', 'https://example.com/missing');
    await settled();
    expect(code(call('WEBSERVICE', 'https://example.com/missing'))).toBe('#VALUE!');
    mockFetch('x'.repeat(40000));
    call('WEBSERVICE', 'https://example.com/big');
    await settled();
    expect(code(call('WEBSERVICE', 'https://example.com/big'))).toBe('#VALUE!');
  });
});

describe('IMAGE', () => {
  it('needs the internet setting and an https address', () => {
    expect(code(call('IMAGE', 'https://example.com/a.png'))).toBe('#CONNECT!');
    setOnlineAllowed(true);
    expect(code(call('IMAGE', 'http://example.com/a.png'))).toBe('#VALUE!');
    expect(code(call('IMAGE', 'https://example.com/a.png', 'x', 3))).toBe('#VALUE!');
    expect(code(call('IMAGE', 'https://example.com/a.png', 'x', 0, 10, 10))).toBe('#VALUE!');
    expect(code(call('IMAGE', 'https://example.com/a.png', 'x', 4))).toBe('#VALUE!');
  });
  it('registers the picture for the renderer, including spilled arrays', () => {
    setOnlineAllowed(true);
    expect(call('IMAGE', 'https://example.com/logo.png', 'Logo', 3, 40, 80)).toBe('Logo');
    expect(cellImage(0, 0, 0)).toMatchObject({ url: 'https://example.com/logo.png', alt: 'Logo', sizing: 3, height: 40, width: 80 });
    expect(call('IMAGE', [['https://example.com/1.png'], ['https://example.com/2.png']])).toEqual([[''], ['']]);
    expect(cellImage(0, 1, 0)?.url).toBe('https://example.com/2.png');
    expect(cellImage(0, 1, 0)?.anchor).toEqual({ sheet: 0, row: 0, col: 0 });
  });
});

describe('STOCKHISTORY', () => {
  const CSV = 'Date,Open,High,Low,Close,Volume\n2024-01-03,10,12,9,11,1000\n2024-01-02,9,10,8,9.5,800\n2024-01-04,11,13,10,12.5,1200\n';
  it('offline by default', () => {
    expect(code(call('STOCKHISTORY', 'MSFT', toSerial(2024, 1, 2)))).toBe('#CONNECT!');
  });
  it('reads daily prices with headers and chosen properties', async () => {
    setOnlineAllowed(true);
    const seen = mockFetch(CSV);
    const args: Value[] = ['XNAS:MSFT', toSerial(2024, 1, 2), toSerial(2024, 1, 4), 0, 1, 0, 1, 5];
    expect(code(call('STOCKHISTORY', ...args))).toBe('#BUSY!');
    await settled();
    expect(call('STOCKHISTORY', ...args)).toEqual([
      ['Date', 'Close', 'Volume'],
      [toSerial(2024, 1, 2), 9.5, 800],
      [toSerial(2024, 1, 3), 11, 1000],
      [toSerial(2024, 1, 4), 12.5, 1200],
    ]);
    expect(seen[0]).toBe('https://stooq.com/q/d/l/?s=msft.us&d1=20240102&d2=20240104&i=d');
  });
  it('argument checks and missing data', async () => {
    setOnlineAllowed(true);
    expect(code(call('STOCKHISTORY', 'MSFT', toSerial(2024, 1, 5), toSerial(2024, 1, 2)))).toBe('#VALUE!');
    expect(code(call('STOCKHISTORY', 'MSFT', toSerial(2024, 1, 2), toSerial(2024, 1, 4), 3))).toBe('#VALUE!');
    expect(code(call('STOCKHISTORY', 'MSFT', toSerial(2024, 1, 2), toSerial(2024, 1, 4), 0, 1, 6))).toBe('#VALUE!');
    mockFetch('No data');
    call('STOCKHISTORY', 'ZZZZ', toSerial(2024, 1, 2));
    await settled();
    expect(code(call('STOCKHISTORY', 'ZZZZ', toSerial(2024, 1, 2)))).toBe('#N/A');
    g.fetch = async () => { throw new TypeError('blocked'); };
    call('STOCKHISTORY', 'AAPL', toSerial(2024, 1, 2));
    await settled();
    expect(code(call('STOCKHISTORY', 'AAPL', toSerial(2024, 1, 2)))).toBe('#CONNECT!');
  });
});

describe('TRANSLATE and DETECTLANGUAGE', () => {
  it('#CONNECT! without a built-in translator', () => {
    expect(code(call('TRANSLATE', 'Hello', 'en', 'de'))).toBe('#CONNECT!');
    expect(code(call('DETECTLANGUAGE', 'Hello'))).toBe('#CONNECT!');
  });
  it('uses the browser translator and detector when present', async () => {
    g.Translator = {
      availability: async () => 'available',
      create: async ({ targetLanguage }: { targetLanguage: string }) => ({ translate: async (t: string) => `[${targetLanguage}] ${t}` }),
    };
    g.LanguageDetector = { create: async () => ({ detect: async () => [{ detectedLanguage: 'fr', confidence: 0.9 }] }) };
    expect(code(call('TRANSLATE', 'Bonjour', 'auto', 'en'))).toBe('#BUSY!');
    expect(code(call('DETECTLANGUAGE', 'Bonjour'))).toBe('#BUSY!');
    await settled();
    expect(call('TRANSLATE', 'Bonjour', 'auto', 'en')).toBe('[en] Bonjour');
    expect(call('DETECTLANGUAGE', 'Bonjour')).toBe('fr');
    expect(code(call('TRANSLATE', 'x', 'en', 'not a language'))).toBe('#VALUE!');
  });
});

describe('background answers recalculate the workbook', () => {
  it('a WEBSERVICE cell shows #BUSY! and then the answer', async () => {
    const { Engine } = await import('../src/engine/engine');
    const { newWorkbook } = await import('../src/model/workbook');
    const { registerAllFunctions } = await import('../src/engine/functions');
    const { onOnlineSettled } = await import('../src/engine/online');
    registerAllFunctions();
    const e = new Engine(newWorkbook());
    setOnlineAllowed(true);
    mockFetch('42');
    let settledCalls = 0;
    const stop = onOnlineSettled(() => { settledCalls++; e.recalc(); });
    try {
      e.setFormula(0, 0, 0, 'WEBSERVICE("https://example.com/n")');
      e.setFormula(0, 0, 1, 'VALUE(A1)*2');
      e.recalc();
      expect(code(e.getValue(0, 0, 0) as Value)).toBe('#BUSY!');
      expect(code(e.getValue(0, 0, 1) as Value)).toBe('#BUSY!');
      await settled();
      await new Promise((r) => setTimeout(r, 5));
      expect(settledCalls).toBe(1);
      expect(e.getValue(0, 0, 0)).toBe('42');
      expect(e.getValue(0, 0, 1)).toBe(84);
    } finally { stop(); }
  });
});
