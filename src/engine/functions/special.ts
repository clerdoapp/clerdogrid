// Numeric kernels for the statistical functions: gamma/beta families, normal, t, chi-square and F
// distributions and a bracketing root finder. Accuracy target is ~1e-14 relative in normal ranges.

const EPS = 3e-16;
const FPMIN = 1e-300;
const LN_SQRT_2PI = 0.9189385332046728;
const SQRT_2PI = 2.5066282746310002;

const LANCZOS = [
  0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
  12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
];

/** sin(πx) without the precision loss of Math.sin(Math.PI * x) for large x. */
export const sinPi = (x: number): number => {
  let r = x % 2;
  if (r < 0) r += 2;
  if (r === 0 || r === 1) return 0;
  if (r === 0.5) return 1;
  if (r === 1.5) return -1;
  return Math.sin(Math.PI * r);
};

/** ln|Γ(x)|. Lanczos (g=7) below 20, Stirling series above; reflection for x < 0.5. */
export const gammaln = (x: number): number => {
  if (Number.isNaN(x)) return NaN;
  if (x <= 0 && Number.isInteger(x)) return Infinity;
  if (x < 0.5) return Math.log(Math.PI / Math.abs(sinPi(x))) - gammaln(1 - x);
  if (x >= 20) {
    const z = 1 / (x * x);
    return (x - 0.5) * Math.log(x) - x + LN_SQRT_2PI + (1 / 12 - z * (1 / 360 - z * (1 / 1260 - z * (1 / 1680 - z / 1188)))) / x;
  }
  const y = x - 1;
  let a = LANCZOS[0]!;
  for (let i = 1; i < 9; i++) a += LANCZOS[i]! / (y + i);
  const t = y + 7.5;
  return LN_SQRT_2PI + (y + 0.5) * Math.log(t) - t + Math.log(a);
};

/** Γ(x), exact for small positive integers; NaN at poles. */
export const gamma = (x: number): number => {
  if (x <= 0 && Number.isInteger(x)) return NaN;
  if (Number.isInteger(x) && x <= 171) {
    let f = 1;
    for (let i = 2; i < x; i++) f *= i;
    return f;
  }
  if (x < 0.5) return Math.PI / (sinPi(x) * gamma(1 - x));
  return Math.exp(gammaln(x));
};

export const lnBeta = (a: number, b: number): number => gammaln(a) + gammaln(b) - gammaln(a + b);

const gser = (a: number, x: number): number => {
  let ap = a, del = 1 / a, sum = del;
  for (let n = 0; n < 1e6; n++) {
    ap += 1;
    del *= x / ap;
    sum += del;
    if (Math.abs(del) < Math.abs(sum) * 1e-17) break;
  }
  return sum * Math.exp(-x + a * Math.log(x) - gammaln(a));
};

const gcf = (a: number, x: number): number => {
  let b = x + 1 - a, c = 1 / FPMIN, d = 1 / b, h = d;
  for (let i = 1; i < 1e6; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    c = b + an / c;
    if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return Math.exp(-x + a * Math.log(x) - gammaln(a)) * h;
};

/** Regularized lower incomplete gamma P(a, x). */
export const gammaP = (a: number, x: number): number => {
  if (x <= 0) return 0;
  if (x === Infinity) return 1;
  return x < a + 1 ? gser(a, x) : 1 - gcf(a, x);
};

/** Regularized upper incomplete gamma Q(a, x) = 1 − P(a, x), accurate in the tail. */
export const gammaQ = (a: number, x: number): number => {
  if (x <= 0) return 1;
  if (x === Infinity) return 0;
  return x < a + 1 ? 1 - gser(a, x) : gcf(a, x);
};

// Continued fraction for the incomplete beta, modified Lentz.
const betacf = (x: number, a: number, b: number): number => {
  const qab = a + b, qap = a + 1, qam = a - 1;
  let c = 1, d = 1 - (qab * x) / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d;
  let h = d;
  for (let m = 1; m < 100000; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return h;
};

/**
 * Regularized incomplete beta as the pair [I_x(a,b), 1 − I_x(a,b)], each accurate even when tiny.
 * `y` must be 1 − x, passed separately so callers can supply it without cancellation.
 */
export const betaIncPair = (x: number, y: number, a: number, b: number): [number, number] => {
  if (x <= 0) return [0, 1];
  if (y <= 0) return [1, 0];
  const front = Math.exp(a * Math.log(x) + b * Math.log(y) - lnBeta(a, b));
  if (x < (a + 1) / (a + b + 2)) {
    const v = (front * betacf(x, a, b)) / a;
    return [v, 1 - v];
  }
  const w = (front * betacf(y, b, a)) / b;
  return [1 - w, w];
};

export const betaInc = (x: number, a: number, b: number): number => betaIncPair(x, 1 - x, a, b)[0];

/** Brent's method on a bracket [a, b] where f changes sign. */
export const zeroin = (f: (x: number) => number, x1: number, x2: number, tol = 1e-300): number => {
  let a = x1, b = x2, c = x2, fa = f(a), fb = f(b), fc = fb, d = b - a, e = d;
  if (fa === 0) return a;
  if (fb === 0) return b;
  for (let it = 0; it < 2000; it++) {
    if ((fb > 0 && fc > 0) || (fb < 0 && fc < 0)) { c = a; fc = fa; e = d = b - a; }
    if (Math.abs(fc) < Math.abs(fb)) { a = b; b = c; c = a; fa = fb; fb = fc; fc = fa; }
    const tol1 = 2 * Number.EPSILON * Math.abs(b) + 0.5 * tol;
    const xm = 0.5 * (c - b);
    if (Math.abs(xm) <= tol1 || fb === 0) return b;
    if (Math.abs(e) >= tol1 && Math.abs(fa) > Math.abs(fb)) {
      const s = fb / fa;
      let p: number, q: number;
      if (a === c) { p = 2 * xm * s; q = 1 - s; }
      else {
        const qq = fa / fc, r = fb / fc;
        p = s * (2 * xm * qq * (qq - r) - (b - a) * (r - 1));
        q = (qq - 1) * (r - 1) * (s - 1);
      }
      if (p > 0) q = -q;
      p = Math.abs(p);
      if (2 * p < Math.min(3 * xm * q - Math.abs(tol1 * q), Math.abs(e * q))) { e = d; d = p / q; }
      else { d = xm; e = d; }
    } else { d = xm; e = d; }
    a = b; fa = fb;
    b += Math.abs(d) > tol1 ? d : xm > 0 ? tol1 : -tol1;
    fb = f(b);
  }
  return b;
};

/** Root of a monotone f on [0, ∞), expanding the bracket upward. */
export const solvePositive = (f: (x: number) => number, guess = 1): number => {
  const f0 = f(0);
  if (f0 === 0) return 0;
  let lo = 0, hi = Math.max(guess, 1e-8);
  while ((f(hi) > 0) === (f0 > 0) && hi < 1e300) { lo = hi; hi *= 2; }
  return hi >= 1e300 ? Infinity : zeroin(f, lo, hi);
};

/** Root of a monotone increasing f on the real line. */
export const solveReal = (f: (x: number) => number): number => {
  let lo = -1, hi = 1;
  while (f(lo) > 0 && lo > -1e300) lo *= 2;
  while (f(hi) < 0 && hi < 1e300) hi *= 2;
  return zeroin(f, lo, hi);
};

// ---- Normal ----

export const normPdf = (z: number): number => Math.exp(-0.5 * z * z) / SQRT_2PI;

export const normCdf = (z: number): number => {
  if (Number.isNaN(z)) return NaN;
  if (!Number.isFinite(z)) return z > 0 ? 1 : 0;
  const q = 0.5 * gammaQ(0.5, 0.5 * z * z);
  return z < 0 ? q : 1 - q;
};

/** Φ(z) − 0.5 without cancellation near 0. */
export const normHalf = (z: number): number => (z < 0 ? -0.5 : 0.5) * gammaP(0.5, 0.5 * z * z);

const A = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.38357751867269e2, -3.066479806614716e1, 2.506628277459239];
const B = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
const C = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
const D = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];

/** Inverse standard normal: Acklam's rational approximation polished with Halley steps to full precision. */
export const normInv = (p: number): number => {
  if (!(p > 0 && p < 1)) return p === 0 ? -Infinity : p === 1 ? Infinity : NaN;
  if (p > 0.5) return -normInv(1 - p);
  let x: number;
  if (p < 0.02425) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((C[0]! * q + C[1]!) * q + C[2]!) * q + C[3]!) * q + C[4]!) * q + C[5]!) / ((((D[0]! * q + D[1]!) * q + D[2]!) * q + D[3]!) * q + 1);
  } else {
    const q = p - 0.5, r = q * q;
    x = ((((((A[0]! * r + A[1]!) * r + A[2]!) * r + A[3]!) * r + A[4]!) * r + A[5]!) * q) /
      (((((B[0]! * r + B[1]!) * r + B[2]!) * r + B[3]!) * r + B[4]!) * r + 1);
  }
  for (let i = 0; i < 3; i++) {
    const e = normCdf(x) - p;
    const u = e * SQRT_2PI * Math.exp(0.5 * x * x);
    const nx = x - u / (1 + 0.5 * x * u);
    if (!Number.isFinite(nx)) break;
    x = nx;
  }
  return x;
};

// ---- Student t ----

export const tPdf = (t: number, v: number): number =>
  Math.exp(gammaln((v + 1) / 2) - gammaln(v / 2) - 0.5 * Math.log(v * Math.PI) - ((v + 1) / 2) * Math.log1p((t * t) / v));

export const tCdf = (t: number, v: number): number => {
  if (Number.isNaN(t)) return NaN;
  if (!Number.isFinite(t)) return t > 0 ? 1 : 0;
  const t2 = t * t;
  // [P(|T| > |t|), P(|T| < |t|)]
  const [two, inner] = betaIncPair(v / (v + t2), t2 / (v + t2), v / 2, 0.5);
  return t > 0 ? 0.5 + 0.5 * inner : 0.5 * two;
};

export const tInv = (p: number, v: number): number => {
  if (p === 0.5) return 0;
  return solveReal((t) => tCdf(t, v) - p);
};

// ---- Gamma / chi-square / F / beta ----

export const gammaPdf = (x: number, a: number, scale: number): number => {
  if (x < 0) return 0;
  if (x === 0) return a < 1 ? Infinity : a === 1 ? 1 / scale : 0;
  return Math.exp((a - 1) * Math.log(x) - x / scale - gammaln(a) - a * Math.log(scale));
};

export const gammaInv = (p: number, a: number): number => solvePositive((x) => gammaP(a, x) - p, a);
export const gammaInvUpper = (q: number, a: number): number => solvePositive((x) => gammaQ(a, x) - q, a);

export const chisqCdf = (x: number, k: number): number => gammaP(k / 2, x / 2);
export const chisqSf = (x: number, k: number): number => gammaQ(k / 2, x / 2);

/** [CDF, survival] of the F distribution. */
export const fCdfPair = (x: number, d1: number, d2: number): [number, number] => {
  if (x <= 0) return [0, 1];
  const z = d1 * x;
  return betaIncPair(z / (z + d2), d2 / (z + d2), d1 / 2, d2 / 2);
};
export const fCdf = (x: number, d1: number, d2: number): number => fCdfPair(x, d1, d2)[0];

export const fPdf = (x: number, d1: number, d2: number): number => {
  if (x < 0) return 0;
  if (x === 0) return d1 < 2 ? Infinity : d1 === 2 ? 1 : 0;
  const z = d1 * x;
  return Math.exp(0.5 * (d1 * Math.log(z) + d2 * Math.log(d2) - (d1 + d2) * Math.log(z + d2)) - Math.log(x) - lnBeta(d1 / 2, d2 / 2));
};

export const betaPdf = (x: number, a: number, b: number): number => {
  if (x < 0 || x > 1) return 0;
  const la = a === 1 ? 0 : (a - 1) * Math.log(x);
  const lb = b === 1 ? 0 : (b - 1) * Math.log1p(-x);
  return Math.exp(la + lb - lnBeta(a, b));
};

export const betaIncInv = (p: number, a: number, b: number): number => {
  if (p <= 0) return 0;
  if (p >= 1) return 1;
  return zeroin((x) => betaInc(x, a, b) - p, 0, 1);
};

/** ln C(n, k) for real n ≥ k ≥ 0. */
export const lnChoose = (n: number, k: number): number => gammaln(n + 1) - gammaln(k + 1) - gammaln(n - k + 1);

/** x·ln(y) with 0·ln(0) = 0. */
export const xlogy = (x: number, y: number): number => (x === 0 ? 0 : x * Math.log(y));
