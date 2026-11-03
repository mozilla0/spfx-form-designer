import { t } from './i18n';

export type FieldKind =
  | 'text'
  | 'note'
  | 'number'
  | 'choice'
  | 'multichoice'
  | 'boolean'
  | 'date'
  | 'datetime'
  | 'user'
  | 'usermulti'
  | 'lookup'
  | 'lookupmulti'
  | 'taxonomy'
  | 'taxonomymulti'
  | 'url'
  | 'calculated'
  | 'unsupported';

export interface ISpField {
  id: string;
  internalName: string;
  title: string;
  typeAsString: string;
  kind: FieldKind;
  required: boolean;
  readOnly: boolean;
  choices: string[];
  lookupListId?: string;
  lookupField?: string;
  /** Verwaltete Metadaten: Ausdruckssatz und optionaler Ankerbegriff. */
  termSetId?: string;
  anchorId?: string;
  /** Link-Spalte: «Hyperlink» oder «Image». */
  urlFormat?: 'hyperlink' | 'image';
  description: string;
}

export interface IUserValue {
  id: number;
  title: string;
  email?: string;
  loginName?: string;
}

export interface ILookupValue {
  id: number;
  title: string;
}

export interface ITermValue {
  termGuid: string;
  label: string;
}

export interface IUrlValue {
  url: string;
  description: string;
}

/**
 * Interne Werte im Formular:
 * text/note/choice/calculated: string · number: string (Eingabe) · multichoice: string[] · boolean: boolean
 * date: 'YYYY-MM-DD' · datetime: 'YYYY-MM-DDTHH:mm' (lokal) · user: IUserValue · usermulti: IUserValue[]
 * lookup: ILookupValue · lookupmulti: ILookupValue[] · taxonomy: ITermValue · taxonomymulti: ITermValue[] · url: IUrlValue
 */
export type FieldValue =
  | string
  | boolean
  | string[]
  | IUserValue
  | IUserValue[]
  | ILookupValue
  | ILookupValue[]
  | ITermValue
  | ITermValue[]
  | IUrlValue
  | null
  | undefined;

export type FieldValues = Record<string, FieldValue>;

export interface ICurrentUser {
  id: number;
  title: string;
  email: string;
  loginName: string;
}

export function fieldKindLabel(kind: FieldKind): string {
  switch (kind) {
    case 'text':
      return t('Einzelne Textzeile');
    case 'note':
      return t('Mehrere Textzeilen');
    case 'number':
      return t('Zahl');
    case 'choice':
      return t('Auswahl');
    case 'multichoice':
      return t('Mehrfachauswahl');
    case 'boolean':
      return t('Ja/Nein');
    case 'date':
      return t('Datum');
    case 'datetime':
      return t('Datum und Uhrzeit');
    case 'user':
      return t('Person');
    case 'usermulti':
      return t('Personen (mehrere)');
    case 'lookup':
      return t('Nachschlagen');
    case 'lookupmulti':
      return t('Nachschlagen (mehrere)');
    case 'taxonomy':
      return t('Verwaltete Metadaten');
    case 'taxonomymulti':
      return t('Verwaltete Metadaten (mehrere)');
    case 'url':
      return t('Link oder Bild');
    case 'calculated':
      return t('Berechnet (nur Anzeige)');
    default:
      return t('Nicht unterstützt');
  }
}

function itemText(v: unknown): string {
  if (!v || typeof v !== 'object') {
    return v === undefined || v === null ? '' : String(v);
  }
  const o = v as Record<string, unknown>;
  if (typeof o.title === 'string') {
    return o.title;
  }
  if (typeof o.label === 'string') {
    return o.label;
  }
  if (typeof o.url === 'string') {
    return (o.description as string) || o.url;
  }
  return '';
}

export function isEmptyValue(v: FieldValue): boolean {
  if (v === undefined || v === null || v === '' || v === false) {
    return true;
  }
  if (Array.isArray(v)) {
    return v.length === 0;
  }
  if (typeof v === 'object' && 'url' in v) {
    return !(v as IUrlValue).url;
  }
  return false;
}

/** Text-Darstellung eines Werts, z. B. für Titel-Platzhalter und Sichtbarkeitsregeln. */
export function valueToText(v: FieldValue): string {
  if (v === undefined || v === null) {
    return '';
  }
  if (typeof v === 'boolean') {
    return v ? t('Ja') : t('Nein');
  }
  if (Array.isArray(v)) {
    return (v as unknown[]).map(itemText).filter(Boolean).join(', ');
  }
  if (typeof v === 'object') {
    return itemText(v);
  }
  return v;
}
