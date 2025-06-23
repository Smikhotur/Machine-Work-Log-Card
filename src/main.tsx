import { createRoot } from 'react-dom/client';
import { BrowserRouter } from 'react-router-dom';

import './shared/config/i18n/i18n';
import './app/styles/styles.scss';
import App from './app/App';

console.log(10, process.env.API_URL);

const rootElement = document.getElementById('root');
if (!rootElement) {
  throw new Error('Root element not found');
}
const root = createRoot(rootElement);

const futureConfig = {
  v7_fetcherPersist: true,
  v7_normalizeFormMethod: true,
  v7_partialHydration: false,
  v7_prependBasename: true,
  v7_relativeSplatPath: true,
  v7_startTransition: true,
  unstable_skipActionErrorRevalidation: false,
};

root.render(
  <BrowserRouter future={futureConfig}>
    <App />
  </BrowserRouter>
);
