import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';

import { LANGUAGE } from '@/shared/const/localStorage';

import en from './locales/en/translation.json';
import ua from './locales/ua/translation.json';

i18n.use(initReactI18next).init({
  resources: {
    en: { translation: en },
    ua: { translation: ua },
  },
  lng: localStorage.getItem(LANGUAGE) || 'en', // мова за замовчуванням
  fallbackLng: 'en',
  interpolation: {
    escapeValue: false, // React вже екранує
  },
  debug: true, // true, якщо хочеш бачити лог у консолі
});

export default i18n;
