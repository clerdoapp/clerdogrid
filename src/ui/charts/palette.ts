// Chart colours, the one source for the grid (as CSS variables set at start) and for printing and PDF.
// Keys are CSS variable names without the leading "--".

export const CHART_LIGHT: Record<string, string> = {
  'chart-1': '#12a565',
  'chart-2': '#2f6fdf',
  'chart-3': '#f08c00',
  'chart-4': '#d6336c',
  'chart-5': '#7048e8',
  'chart-6': '#0b8fb3',
  'chart-7': '#e03131',
  'chart-8': '#74a10d',
  'chart-9': '#5c677d',
  'chart-10': '#c2620a',
  'chart-bg': 'var(--sheet)',
  'chart-ink': '#1f2a24',
  'chart-muted': '#63716a',
  'chart-grid': '#e6ebe8',
  'chart-axis': '#a6b2ab',
  'chart-hover': 'rgba(18, 165, 101, 0.08)',
  'chart-on': '#ffffff',
  'chart-up': '#12a565',
  'chart-down': '#e03131',
  'chart-total': '#4a6382',
  'chart-wall': 'rgba(127, 140, 133, 0.08)',
  'chart-volume': '#8a97a8',
  'chart-s1-1': '#0c7a4b',
  'chart-s1-2': '#12a565',
  'chart-s1-3': '#4cc38a',
  'chart-s1-4': '#8fd9b2',
  'chart-s1-5': '#0a5c39',
  'chart-s1-6': '#2e9c6e',
  'chart-s1-7': '#b9e8cf',
  'chart-s1-8': '#06402a',
  'chart-s2-1': '#12a565',
  'chart-s2-2': '#e8a317',
  'chart-s2-3': '#0c7a4b',
  'chart-s2-4': '#f2c24f',
  'chart-s2-5': '#4cc38a',
  'chart-s2-6': '#c27a0e',
  'chart-s2-7': '#0a5c39',
  'chart-s2-8': '#f5d98c',
  'chart-s3-1': '#2bb380',
  'chart-s3-2': '#5c6f82',
  'chart-s3-3': '#7fd1b0',
  'chart-s3-4': '#34495e',
  'chart-s3-5': '#a8e2ca',
  'chart-s3-6': '#8a9bab',
  'chart-s3-7': '#16805a',
  'chart-s3-8': '#c4ced8',
  'chart-s4-1': '#6cc49a',
  'chart-s4-2': '#9ab8d6',
  'chart-s4-3': '#e7b97a',
  'chart-s4-4': '#c9a3d4',
  'chart-s4-5': '#e8a0a8',
  'chart-s4-6': '#8fcfcf',
  'chart-s4-7': '#b8c77a',
  'chart-s4-8': '#a7adbb',
};

/** Dark theme: only what differs from light. */
export const CHART_DARK: Record<string, string> = {
  'chart-1': '#34c27f',
  'chart-2': '#6b9bff',
  'chart-3': '#ffb340',
  'chart-4': '#f06595',
  'chart-5': '#9f86ff',
  'chart-6': '#3bc9db',
  'chart-7': '#ff6b6b',
  'chart-8': '#a9e34b',
  'chart-9': '#a0a9bd',
  'chart-10': '#ff9f4a',
  'chart-ink': '#e5efe9',
  'chart-muted': '#91a398',
  'chart-grid': '#26332c',
  'chart-axis': '#4f6358',
  'chart-hover': 'rgba(52, 194, 127, 0.12)',
  'chart-on': '#ffffff',
  'chart-up': '#34c27f',
  'chart-down': '#ff7b72',
  'chart-total': '#8fa9c9',
  'chart-wall': 'rgba(160, 175, 166, 0.08)',
  'chart-volume': '#6f7d8f',
  'chart-s1-1': '#34c27f',
  'chart-s1-2': '#6fdba6',
  'chart-s1-3': '#1f9a62',
  'chart-s1-4': '#a6ecc8',
  'chart-s1-5': '#17784b',
  'chart-s1-6': '#4fd193',
  'chart-s1-7': '#c9f3dd',
  'chart-s1-8': '#0f5a37',
  'chart-s2-1': '#34c27f',
  'chart-s2-2': '#ffc14d',
  'chart-s2-3': '#1f9a62',
  'chart-s2-4': '#ffd986',
  'chart-s2-5': '#6fdba6',
  'chart-s2-6': '#e0a030',
  'chart-s2-7': '#17784b',
  'chart-s2-8': '#ffe7b0',
  'chart-s3-1': '#3dd6a4',
  'chart-s3-2': '#8da2b8',
  'chart-s3-3': '#7fe3c2',
  'chart-s3-4': '#5f7690',
  'chart-s3-5': '#b2f0dc',
  'chart-s3-6': '#aebccc',
  'chart-s3-7': '#22a77b',
  'chart-s3-8': '#d0d9e3',
  'chart-s4-1': '#7ad6aa',
  'chart-s4-2': '#a9c6e6',
  'chart-s4-3': '#f0c78e',
  'chart-s4-4': '#d6b3e0',
  'chart-s4-5': '#f2b0b8',
  'chart-s4-6': '#9fdede',
  'chart-s4-7': '#c9d88c',
  'chart-s4-8': '#b7bdcb',
};

const decls = (vars: Record<string, string>): string => Object.entries(vars).map(([k, v]) => `--${k}: ${v};`).join(' ');

/** The stylesheet text that defines the chart variables for both themes. */
export const chartPaletteCss = (): string =>
  `:root { ${decls(CHART_LIGHT)} }\n`
  + `@media (prefers-color-scheme: dark) { :root:not([data-theme='light']) { ${decls(CHART_DARK)} } }\n`
  + `:root[data-theme='dark'] { ${decls(CHART_DARK)} }\n`;

/** Light values keyed by full variable name ("--chart-1"), for drawing outside the page's styles. */
export const chartLightVars = (): Record<string, string> => Object.fromEntries(Object.entries(CHART_LIGHT).map(([k, v]) => [`--${k}`, v]));

let installed = false;
/** Adds the variables to the document once (browser only). */
export const installChartPalette = (): void => {
  if (installed || typeof document === 'undefined') return;
  installed = true;
  const el = document.createElement('style');
  el.dataset.chartPalette = '';
  el.textContent = chartPaletteCss();
  document.head.prepend(el);
};
installChartPalette();
