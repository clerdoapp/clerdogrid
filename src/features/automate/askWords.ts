// Words for "Ask about your data" in the 19 interface languages. Questions are read by finding the
// column names they mention and these words around them, so word order can differ by language:
// "total Sales by Region", "Summe Umsatz nach Region", "مجموع المبيعات حسب المنطقة", "地域別の売上の合計".

export type AggWord = 'sum' | 'average' | 'count' | 'min' | 'max';

/** Aggregation words, matched as whole words (or as substrings in languages written without spaces). */
export const AGG: Record<AggWord, string[]> = {
  sum: ['total', 'sum', 'sum of', 'summe', 'gesamt', 'insgesamt', 'suma', 'somme', 'totale', 'soma', 'totaal', 'summa', 'totalt', 'łącznie', 'razem',
    'toplam', 'сумма', 'итог', 'всего', 'сума', 'підсумок', 'всього', 'jumlah', 'total', 'कुल', 'योग', 'סך הכול', 'סך', 'סכום', 'סה"כ', 'مجموع', 'إجمالي', 'اجمالي',
    '总计', '合计', '总和', '求和', '合計', '総計', '総額', '합계', '총합', '총계'],
  average: ['average', 'avg', 'mean', 'durchschnitt', 'mittelwert', 'promedio', 'media', 'moyenne', 'média', 'gemiddelde', 'gemiddeld', 'genomsnitt', 'medel',
    'średnia', 'ortalama', 'среднее', 'средний', 'середнє', 'середній', 'rata rata', 'rerata', 'औसत', 'ממוצע', 'متوسط', 'معدل', '平均', '평균'],
  count: ['count', 'count of', 'number of', 'how many', 'anzahl', 'wie viele', 'recuento', 'cuenta', 'número de', 'cuántos', 'nombre de', 'combien', 'compte',
    'conteggio', 'numero di', 'quanti', 'contagem', 'quantos', 'aantal', 'hoeveel', 'antal', 'hur många', 'liczba', 'ile', 'sayı', 'sayısı', 'kaç', 'adet',
    'количество', 'сколько', 'число', 'кількість', 'скільки', 'banyaknya', 'berapa', 'गिनती', 'कितने', 'ספירה', 'כמה', 'מספר', 'عدد', 'كم',
    '计数', '数量', '个数', '多少', '件数', '個数', '数', '개수', '건수', '몇'],
  min: ['min', 'minimum', 'lowest', 'smallest', 'kleinste', 'niedrigste', 'mínimo', 'menor', 'plus petit', 'plus bas', 'minimo', 'più basso', 'laagste',
    'kleinste', 'lägsta', 'minsta', 'najmniejsza', 'najniższa', 'minimum', 'en düşük', 'en az', 'минимум', 'наименьшее', 'мінімум', 'найменше', 'terendah',
    'terkecil', 'न्यूनतम', 'सबसे कम', 'מינימום', 'הנמוך', 'أدنى', 'أقل', '最小', '最低', '최소', '최저'],
  max: ['max', 'maximum', 'highest', 'largest', 'biggest', 'größte', 'höchste', 'máximo', 'mayor', 'plus grand', 'plus haut', 'massimo', 'più alto', 'hoogste',
    'grootste', 'högsta', 'största', 'największa', 'najwyższa', 'maksimum', 'en yüksek', 'en çok', 'максимум', 'наибольшее', 'максимум', 'найбільше', 'tertinggi',
    'terbesar', 'अधिकतम', 'सबसे ज़्यादा', 'מקסימום', 'הגבוה', 'أعلى', 'أكبر', '最大', '最高', '최대', '최고'],
};

/** "By" words that come before the grouping column ("by Region", "nach Region", "按地区", "حسب المنطقة"). */
export const BY_BEFORE = ['by', 'per', 'for each', 'grouped by', 'across', 'nach', 'pro', 'je', 'für jede', 'für jeden', 'por', 'por cada', 'par', 'pour chaque',
  'per', 'per ogni', 'para cada', 'voor elke', 'naar', 'op basis van', 'per', 'efter', 'för varje', 'per', 'według', 'wg', 'na', 'dla każdego', 'по', 'для каждого', 'за', 'для кожного',
  'menurut', 'berdasarkan', 'per', 'tiap', 'untuk setiap', 'לפי', 'עבור כל', 'בכל', 'حسب', 'بحسب', 'لكل', 'प्रति', '按', '按照', '每个', '各'];

/** "By" words that come after the grouping column ("Region'a göre", "地域別", "지역별", "क्षेत्र के अनुसार"). */
export const BY_AFTER = ['göre', 'başına', 'bazında', '別', 'ごと', '毎', '별', '기준', '마다', 'के अनुसार', 'अनुसार', 'वार'];

export const TOP = ['top', 'best', 'beste', 'besten', 'mejores', 'meilleurs', 'meilleures', 'migliori', 'melhores', 'bästa', 'najlepsze', 'en iyi', 'лучшие', 'топ',
  'найкращі', 'terbaik', 'teratas', 'शीर्ष', 'टॉप', 'הגבוהים', 'המובילים', 'أعلى', 'أفضل', '前', '上位', '상위'];
export const BOTTOM = ['bottom', 'worst', 'letzte', 'schlechteste', 'schlechtesten', 'peores', 'derniers', 'pires', 'peggiori', 'piores', 'slechtste', 'sämsta',
  'najgorsze', 'en kötü', 'худшие', 'найгірші', 'terbawah', 'terburuk', 'निचले', 'सबसे खराब', 'הנמוכים', 'أدنى', 'أسوأ', '后', '最后', '下位', '하위'];

export const TREND = ['over time', 'trend', 'by date', 'im zeitverlauf', 'verlauf', 'entwicklung', 'a lo largo del tiempo', 'tendencia', 'evolución', 'au fil du temps',
  'évolution', 'tendance', 'nel tempo', 'andamento', 'tendenza', 'ao longo do tempo', 'tendência', 'evolução', 'in de tijd', 'verloop', 'över tid', 'utveckling',
  'w czasie', 'zaman içinde', 'eğilim', 'trendi', 'во времени', 'динамика', 'тренд', 'з часом', 'динаміка', 'dari waktu ke waktu', 'tren', 'समय के साथ', 'रुझान',
  'לאורך זמן', 'מגמה', 'بمرور الوقت', 'عبر الزمن', 'اتجاه', '趋势', '随时间', '推移', '傾向', '추이', '추세'];

