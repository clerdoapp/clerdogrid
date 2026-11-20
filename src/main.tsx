import { createRoot } from 'react-dom/client';
import '@fontsource/figtree/400.css';
import '@fontsource/figtree/500.css';
import '@fontsource/figtree/600.css';
import '@fontsource/outfit/600.css';
import '@clerdo/frame/style.css';
import './ui/fonts-bundled.css';
import './ui/app.css';
import { App } from './ui/App';
import { initLang } from './ui/i18n';

// The interface language loads before the first render so nothing flashes in English.
void initLang().finally(() => createRoot(document.getElementById('root')!).render(<App />));

if ('serviceWorker' in navigator && (location.protocol === 'https:' || location.hostname === 'localhost')) {
  void navigator.serviceWorker.register('./sw.js').catch(() => undefined);
}
