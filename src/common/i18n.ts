import { en } from './i18n.en';

/**
 * Übersetzung der Oberfläche. Die deutschen Texte sind zugleich die Schlüssel;
 * i18n.en.ts enthält die englischen Entsprechungen. Platzhalter: {0}, {1}, …
 */

export type Language = 'de' | 'en';

let language: Language = 'de';
let culture: string = 'de-CH';

/** Setzt Sprache und Kultur anhand der Seiten-Kultur, z. B. «de-CH» oder «en-US». */
export function setCulture(cultureName: string | undefined): void {
  culture = cultureName || 'de-CH';
  language = /^de\b/i.test(culture) ? 'de' : 'en';
}

export function getLanguage(): Language {
  return language;
}

export function getCulture(): string {
  return culture;
}

export function t(key: string, ...args: Array<string | number>): string {
  const text = language === 'de' ? key : en[key] || key;
  return args.length ? text.replace(/\{(\d+)\}/g, (m, i: string) => (args[Number(i)] !== undefined ? String(args[Number(i)]) : m)) : text;
}
