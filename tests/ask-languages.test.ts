import { describe, expect, it } from 'vitest';
import { parseAsk, type ColumnSummary } from '../src/features/automate/insights';
import type { Catalog } from '../src/ui/i18n';

const langs = import.meta.glob<{ default: Catalog }>('../src/ui/i18n/lang/*.ts', { eager: true });

const colsOf = (headers: string[]): ColumnSummary[] =>
  headers.map((name, index) => ({ index, name, type: index === 0 ? 'date' : index === 1 ? 'text' : 'number', count: 10, blanks: 0, distinct: 10 }));

// Columns: date, category, two numbers. Questions: sum by, average by, top 5, count by, over time.
const CASES: Record<string, { headers: [string, string, string, string]; avg: string; top: string; bottom?: string; count: string; trend: string }> = {
  ar: { headers: ['التاريخ', 'المنطقة', 'المبيعات', 'الوحدات'], avg: 'متوسط الوحدات حسب المنطقة', top: 'أعلى 5 مناطق حسب المبيعات', bottom: 'أدنى 3 مناطق حسب المبيعات', count: 'عدد الصفوف حسب المنطقة', trend: 'المبيعات بمرور الوقت' },
  de: { headers: ['Datum', 'Region', 'Umsatz', 'Stück'], avg: 'Durchschnitt Stück pro Region', top: 'die besten 5 Regionen nach Umsatz', bottom: 'die schlechtesten 3 Regionen nach Umsatz', count: 'Anzahl nach Region', trend: 'Umsatz im Zeitverlauf' },
  es: { headers: ['Fecha', 'Región', 'Ventas', 'Unidades'], avg: 'promedio de unidades por región', top: 'las 5 mejores regiones por ventas', count: 'recuento por región', trend: 'ventas a lo largo del tiempo' },
  fr: { headers: ['Date', 'Région', 'Ventes', 'Unités'], avg: 'moyenne des unités par région', top: 'les 5 meilleures régions par ventes', count: 'nombre de lignes par région', trend: 'ventes au fil du temps' },
  it: { headers: ['Data', 'Regione', 'Vendite', 'Unità'], avg: 'media unità per regione', top: 'le 5 migliori regioni per vendite', count: 'conteggio per regione', trend: 'vendite nel tempo' },
  pt: { headers: ['Data', 'Região', 'Vendas', 'Unidades'], avg: 'média de unidades por região', top: 'as 5 melhores regiões por vendas', count: 'contagem por região', trend: 'vendas ao longo do tempo' },
  nl: { headers: ['Datum', 'Regio', 'Omzet', 'Aantal stuks'], avg: 'gemiddelde aantal stuks per regio', top: 'top 5 regio’s naar omzet', count: 'aantal per regio', trend: 'omzet in de tijd' },
  sv: { headers: ['Datum', 'Region', 'Försäljning', 'Enheter'], avg: 'genomsnitt enheter per region', top: 'bästa 5 regioner efter försäljning', count: 'antal per region', trend: 'försäljning över tid' },
  pl: { headers: ['Data', 'Region', 'Sprzedaż', 'Sztuki'], avg: 'średnia sztuk według regionu', top: 'najlepsze 5 regionów według sprzedaży', count: 'liczba według regionu', trend: 'sprzedaż w czasie' },
  tr: { headers: ['Tarih', 'Bölge', 'Satışlar', 'Adet'], avg: 'bölgeye göre ortalama adet', top: 'satışlara göre en iyi 5 bölge', count: 'bölge başına sayı', trend: 'zaman içinde satışlar' },
  ru: { headers: ['Дата', 'Регион', 'Продажи', 'Штуки'], avg: 'среднее штук по регионам', top: 'лучшие 5 регионов по продажам', count: 'количество по регионам', trend: 'динамика продаж' },
  uk: { headers: ['Дата', 'Область', 'Продажі', 'Штуки'], avg: 'середнє штук за областями', top: 'найкращі 5 областей за продажами', count: 'кількість за областями', trend: 'динаміка продажів' },
  zh: { headers: ['日期', '地区', '销售额', '数量'], avg: '按地区的数量平均', top: '按销售额前5的地区', bottom: '按销售额后3的地区', count: '按地区计数', trend: '销售额趋势' },
  ja: { headers: ['日付', '地域', '売上', '個数'], avg: '地域別の個数の平均', top: '売上で上位5の地域', bottom: '売上で下位3の地域', count: '地域別の件数', trend: '売上の推移' },
  ko: { headers: ['날짜', '지역', '매출', '수량'], avg: '지역별 수량 평균', top: '매출 기준 상위 5개 지역', bottom: '매출 기준 하위 3개 지역', count: '지역별 개수', trend: '매출 추이' },
  hi: { headers: ['दिनांक', 'क्षेत्र', 'बिक्री', 'इकाइयाँ'], avg: 'क्षेत्र के अनुसार औसत इकाइयाँ', top: 'बिक्री के अनुसार शीर्ष 5 क्षेत्र', count: 'क्षेत्र के अनुसार गिनती', trend: 'समय के साथ बिक्री' },
  id: { headers: ['Tanggal', 'Wilayah', 'Penjualan', 'Unit'], avg: 'rata-rata unit per wilayah', top: '5 wilayah teratas menurut penjualan', count: 'banyaknya per wilayah', trend: 'penjualan dari waktu ke waktu' },
  he: { headers: ['תאריך', 'אזור', 'מכירות', 'יחידות'], avg: 'ממוצע היחידות לפי אזור', top: '5 האזורים המובילים לפי מכירות', bottom: '3 האזורים הנמוכים לפי מכירות', count: 'ספירה לפי אזור', trend: 'המכירות לאורך זמן' },
};

describe('asking about data in every interface language', () => {
  it('covers every language', () => {
    expect(Object.keys(CASES).sort()).toEqual(Object.keys(langs).map((f) => /(\w+)\.ts$/.exec(f)![1]).sort());
  });

  for (const [code, c] of Object.entries(CASES)) {
    describe(code, () => {
      const cat = langs[`../src/ui/i18n/lang/${code}.ts`]!.default;
      const cols = colsOf(c.headers);
      it('reads the example shown in the box', () => {
        const h = [c.headers[0], cat['Region'], cat['Sales'], c.headers[3]];
        expect(parseAsk(cat['total sales by region'], h, colsOf(h))).toEqual({ kind: 'group', agg: 'sum', val: 2, by: 1 });
      });
      it('reads averages, top items, counts and trends', () => {
        expect(parseAsk(c.avg, c.headers, cols), c.avg).toEqual({ kind: 'group', agg: 'average', val: 3, by: 1 });
        expect(parseAsk(c.top, c.headers, cols), c.top).toEqual({ kind: 'top', n: 5, by: 1, val: 2, bottom: false });
        if (c.bottom) expect(parseAsk(c.bottom, c.headers, cols), c.bottom).toEqual({ kind: 'top', n: 3, by: 1, val: 2, bottom: true });
        expect(parseAsk(c.count, c.headers, cols), c.count).toMatchObject({ kind: 'group', agg: 'count', by: 1 });
        expect(parseAsk(c.trend, c.headers, cols), c.trend).toEqual({ kind: 'trend', val: 2, x: 0 });
      });
      it('says nothing when no column is named', () => {
        expect(parseAsk(cat['Ask about your data'], c.headers, cols)).toBeNull();
      });
    });
  }

  it('does not read a longer column name as a shorter one', () => {
    const h = ['Sales', 'Salesperson', 'Units'];
    expect(parseAsk('Summe Sales nach Salesperson', h, colsOf(['x', 'y', 'z']).map((x, i) => ({ ...x, type: i === 1 ? 'text' : 'number' })))).toEqual({ kind: 'group', agg: 'sum', val: 0, by: 1 });
  });
});
