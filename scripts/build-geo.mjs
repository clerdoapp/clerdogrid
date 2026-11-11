// Builds src/ui/charts/geo/regions.json, the region shapes for filled map charts, from
// Natural Earth (public domain, https://www.naturalearthdata.com/):
//   ne_50m_admin_0_countries.geojson and ne_110m_admin_1_states_provinces_lakes.geojson
// Usage: node scripts/build-geo.mjs <folder with the two .geojson files>
// Shapes are simplified (Douglas–Peucker), quantised to 1/Q degree and delta-encoded.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

const dir = process.argv[2];
if (!dir) { console.error('Pass the folder holding the Natural Earth .geojson files.'); process.exit(1); }
const Q = 20;

const sqSegDist = (p, a, b) => {
  let [x, y] = a; let dx = b[0] - x, dy = b[1] - y;
  if (dx || dy) {
    const t = ((p[0] - x) * dx + (p[1] - y) * dy) / (dx * dx + dy * dy);
    if (t > 1) { x = b[0]; y = b[1]; } else if (t > 0) { x += dx * t; y += dy * t; }
  }
  dx = p[0] - x; dy = p[1] - y;
  return dx * dx + dy * dy;
};
const simplify = (pts, tol) => {
  if (pts.length < 4) return pts;
  const keep = new Uint8Array(pts.length); keep[0] = keep[pts.length - 1] = 1;
  const stack = [[0, pts.length - 1]];
  while (stack.length) {
    const [a, b] = stack.pop();
    let max = 0, idx = -1;
    for (let i = a + 1; i < b; i++) { const d = sqSegDist(pts[i], pts[a], pts[b]); if (d > max) { max = d; idx = i; } }
    if (max > tol * tol) { keep[idx] = 1; stack.push([a, idx], [idx, b]); }
  }
  return pts.filter((_, i) => keep[i]);
};
const area = (r) => { let s = 0; for (let i = 0; i < r.length; i++) { const a = r[i], b = r[(i + 1) % r.length]; s += a[0] * b[1] - b[0] * a[1]; } return Math.abs(s / 2); };

const encode = (geom, tol, minArea) => {
  const polys = geom.type === 'Polygon' ? [geom.coordinates] : geom.type === 'MultiPolygon' ? geom.coordinates : [];
  const rings = polys.flatMap((p) => p.map((r, k) => ({ r, hole: k > 0 })));
  const largest = rings.reduce((m, x) => (area(x.r) > area(m.r) ? x : m), rings[0]);
  const out = [];
  for (const { r } of rings) {
    const big = r === largest.r;
    if (!big && area(r) < minArea) continue;
    let s = simplify(r, tol);
    if (s.length < 4) { if (!big) continue; s = simplify(r, tol / 6); }
    const q = s.map(([x, y]) => [Math.round(x * Q), Math.round(y * Q)]).filter((p, i, a) => i === 0 || p[0] !== a[i - 1][0] || p[1] !== a[i - 1][1]);
    if (q.length < 4 && !big) continue;
    const flat = [];
    q.forEach((p, i) => { if (i === 0) flat.push(p[0], p[1]); else flat.push(p[0] - q[i - 1][0], p[1] - q[i - 1][1]); });
    out.push(flat);
  }
  return out;
};

const read = (f) => JSON.parse(readFileSync(join(dir, f), 'utf8'));
const clean = (v) => (typeof v === 'string' && v && v !== '-99' ? v : undefined);

const countries = read('ne_50m_admin_0_countries.geojson').features.map((f) => {
  const p = f.properties;
  const name = p.NAME;
  const alt = [...new Set([p.NAME_LONG, p.FORMAL_EN, p.NAME_SORT, p.NAME_ALT, p.NAME_EN, p.ADMIN, p.BRK_NAME, p.GEOUNIT, p.ABBREV].filter((x) => clean(x) && x !== name))];
  return {
    n: name, a2: clean(p.ISO_A2) ?? clean(p.ISO_A2_EH), a3: clean(p.ISO_A3) ?? clean(p.ISO_A3_EH) ?? clean(p.ADM0_A3),
    ...(alt.length ? { alt } : {}), p: encode(f.geometry, 0.07, 0.01),
  };
}).filter((c) => c.p.length);

const states = read('ne_110m_admin_1_states_provinces_lakes.geojson').features
  .filter((f) => f.properties.adm0_a3 === 'USA')
  .map((f) => {
    const p = f.properties;
    return { n: p.name, a2: p.postal, a3: p.iso_3166_2, p: encode(f.geometry, 0.02, 0.002) };
  }).filter((s) => s.p.length);

const outDir = new URL('../src/ui/charts/geo/', import.meta.url);
mkdirSync(outDir, { recursive: true });
const json = JSON.stringify({ v: 1, q: Q, source: 'Natural Earth (public domain)', world: countries, us: states });
writeFileSync(new URL('regions.json', outDir), json);
console.log(`countries ${countries.length}, states ${states.length}, ${(json.length / 1024).toFixed(0)} KB`);
