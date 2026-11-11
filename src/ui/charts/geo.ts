// Region shapes for filled maps (Natural Earth, public domain), loaded on first use.

export interface Region { name: string; a2?: string; a3?: string; alt: string[]; rings: Float64Array[]; bbox: [number, number, number, number] }
export interface GeoData { world: Region[]; us: Region[] }

interface RawRegion { n: string; a2?: string; a3?: string; alt?: string[]; p: number[][] }

let cache: GeoData | null = null;
let loading: Promise<GeoData> | null = null;

const decode = (raw: { q: number; world: RawRegion[]; us: RawRegion[] }): GeoData => {
  const q = raw.q || 20;
  const conv = (r: RawRegion): Region => {
    let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
    const rings = r.p.map((flat) => {
      const out = new Float64Array(flat.length);
      let x = 0, y = 0;
      for (let i = 0; i < flat.length; i += 2) {
        x = i ? x + flat[i]! : flat[i]!; y = i ? y + flat[i + 1]! : flat[i + 1]!;
        out[i] = x / q; out[i + 1] = y / q;
        x0 = Math.min(x0, out[i]!); x1 = Math.max(x1, out[i]!); y0 = Math.min(y0, out[i + 1]!); y1 = Math.max(y1, out[i + 1]!);
      }
      return out;
    });
    return { name: r.n, a2: r.a2, a3: r.a3, alt: r.alt ?? [], rings, bbox: [x0, y0, x1, y1] };
  };
  return { world: raw.world.map(conv), us: raw.us.map(conv) };
};

export const geoData = (): GeoData | null => cache;
export const loadGeo = (): Promise<GeoData> => {
  loading ??= import('./geo/regions.json').then((m) => (cache = decode((m as { default: { q: number; world: RawRegion[]; us: RawRegion[] } }).default)));
  return loading;
};

export const normName = (s: string): string => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/&/g, ' and ').replace(/[^a-z0-9]+/g, ' ').replace(/^the /, '').trim();

/** Everyday names that differ from the data set's own. */
const ALIASES: Record<string, string> = {
  usa: 'USA', us: 'USA', 'united states': 'USA', america: 'USA', 'united states of america': 'USA', 'u s a': 'USA', 'u s': 'USA',
  uk: 'GBR', 'great britain': 'GBR', britain: 'GBR', england: 'GBR', 'united kingdom': 'GBR',
  russia: 'RUS', 'south korea': 'KOR', korea: 'KOR', 'republic of korea': 'KOR', 'north korea': 'PRK', 'czech republic': 'CZE', czechia: 'CZE',
  holland: 'NLD', burma: 'MMR', swaziland: 'SWZ', 'ivory coast': 'CIV', 'cote d ivoire': 'CIV', 'cape verde': 'CPV', 'east timor': 'TLS',
  vatican: 'VAT', 'vatican city': 'VAT', 'drc': 'COD', 'dr congo': 'COD', 'democratic republic of congo': 'COD', 'congo kinshasa': 'COD', 'congo brazzaville': 'COG',
  'republic of the congo': 'COG', 'macedonia': 'MKD', 'north macedonia': 'MKD', 'uae': 'ARE', 'emirates': 'ARE', 'turkiye': 'TUR', 'iran': 'IRN', 'syria': 'SYR',
  'laos': 'LAO', 'vietnam': 'VNM', 'viet nam': 'VNM', 'bolivia': 'BOL', 'venezuela': 'VEN', 'tanzania': 'TZA', 'moldova': 'MDA', 'palestine': 'PSE',
  'taiwan': 'TWN', 'brunei': 'BRN', 'bahamas': 'BHS', 'gambia': 'GMB', 'micronesia': 'FSM', 'st lucia': 'LCA', 'saint lucia': 'LCA',
};

export interface Match { area: 'world' | 'us'; regions: Region[]; index: (number)[] }

const indexOf = (list: Region[], us: boolean) => {
  const m = new Map<string, number>();
  list.forEach((r, i) => {
    for (const k of [r.name, ...r.alt]) { const n = normName(k); if (n && !m.has(n)) m.set(n, i); }
    if (r.a2) m.set(normName(r.a2), i);
    if (r.a3) m.set(normName(r.a3), i);
    if (us && r.a3) m.set(normName(r.a3.replace(/^US-/, '')), i);
  });
  return m;
};
const idxCache = new WeakMap<Region[], Map<string, number>>();

/** Finds the region for each name. `area` 'auto' picks US states when most names are states. */
export const matchRegions = (geo: GeoData, names: string[], area: 'auto' | 'world' | 'us' = 'auto'): Match => {
  const look = (list: Region[], us: boolean) => {
    let ix = idxCache.get(list);
    if (!ix) { ix = indexOf(list, us); idxCache.set(list, ix); }
    return names.map((nm) => {
      const n = normName(nm.replace(/,\s*(usa|us|united states)$/i, ''));
      if (!n) return -1;
      const hit = ix!.get(n);
      if (hit !== undefined) return hit;
      const alias = ALIASES[n];
      if (alias && !us) { const j = list.findIndex((r) => r.a3 === alias); if (j >= 0) return j; }
      return -1;
    });
  };
  const usIdx = area === 'world' ? [] : look(geo.us, true);
  const wIdx = area === 'us' ? [] : look(geo.world, false);
  const usHits = usIdx.filter((i) => i >= 0).length, wHits = wIdx.filter((i) => i >= 0).length;
  // Two-letter codes are shared by states and countries; full state names decide.
  const stateNames = names.filter((nm) => nm.trim().length > 2 && usIdx[names.indexOf(nm)]! >= 0).length;
  const useUs = area === 'us' || (area === 'auto' && usHits > 0 && (stateNames > wHits / 2 || usHits > wHits));
  return useUs ? { area: 'us', regions: geo.us, index: usIdx } : { area: 'world', regions: geo.world, index: wIdx };
};

// ---------- Projections ----------

const RAD = Math.PI / 180;
/** Equal Earth projection (Šavrič, Patterson and Jenny, 2018). */
export const equalEarth = (lon: number, lat: number): [number, number] => {
  const A1 = 1.340264, A2 = -0.081106, A3 = 0.000893, A4 = 0.003796, M = Math.sqrt(3) / 2;
  const th = Math.asin(M * Math.sin(lat * RAD)), t2 = th * th, t6 = t2 * t2 * t2;
  const x = (lon * RAD * Math.cos(th)) / (M * (A1 + 3 * A2 * t2 + t6 * (7 * A3 + 9 * A4 * t2)));
  const y = th * (A1 + A2 * t2 + t6 * (A3 + A4 * t2));
  return [x, -y];
};

/** Albers equal-area conic. */
export const albers = (p1: number, p2: number, lat0: number, lon0: number) => {
  const n = (Math.sin(p1 * RAD) + Math.sin(p2 * RAD)) / 2;
  const C = Math.cos(p1 * RAD) ** 2 + 2 * n * Math.sin(p1 * RAD);
  const r0 = Math.sqrt(C - 2 * n * Math.sin(lat0 * RAD)) / n;
  return (lon: number, lat: number): [number, number] => {
    const r = Math.sqrt(Math.max(0, C - 2 * n * Math.sin(lat * RAD))) / n;
    const th = n * (lon - lon0) * RAD;
    return [r * Math.sin(th), -(r0 - r * Math.cos(th))];
  };
};

const lower48 = albers(29.5, 45.5, 37.5, -96);
const alaska = albers(55, 65, 50, -154);
const hawaii = albers(8, 18, 3, -157);

/** US states with Alaska and Hawaii moved below the lower 48, in the lower-48 projection's units. */
export const usProject = (region: Region | null, lon: number, lat: number): [number, number] => {
  const code = region?.a2;
  if (code === 'AK' || (!region && lat > 50 && lon < -130)) { const [x, y] = alaska(lon, lat); return [x * 0.4 - 0.305, y * 0.4 + 0.249]; }
  if (code === 'HI' || (!region && lat < 25 && lon < -150)) { const [x, y] = hawaii(lon, lat); return [x - 0.144, y + 0.486]; }
  return lower48(lon, lat);
};
