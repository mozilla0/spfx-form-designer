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
  | 'lookup'
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

/**
 * Interne Werte im Formular:
 * text/note/choice: string · number: string (Eingabe) · multichoice: string[] · boolean: boolean
 * date: 'YYYY-MM-DD' · datetime: 'YYYY-MM-DDTHH:mm' (lokal) · user: IUserValue · lookup: ILookupValue
 */
export type FieldValue = string | boolean | string[] | IUserValue | ILookupValue | null | undefined;

export type FieldValues = Record<string, FieldValue>;

export interface ICurrentUser {
  id: number;
  title: string;
  email: string;
  loginName: string;
}

export const FIELD_KIND_LABELS: Record<FieldKind, string> = {
  text: 'Einzelne Textzeile',
  note: 'Mehrere Textzeilen',
  number: 'Zahl',
  choice: 'Auswahl',
  multichoice: 'Mehrfachauswahl',
  boolean: 'Ja/Nein',
  date: 'Datum',
  datetime: 'Datum und Uhrzeit',
  user: 'Person',
  lookup: 'Nachschlagen',
  unsupported: 'Nicht unterstützt'
};

export function isEmptyValue(v: FieldValue): boolean {
  return (
    v === undefined ||
    v === null ||
    v === '' ||
    v === false ||
    (Array.isArray(v) && v.length === 0)
  );
}

/** Text-Darstellung eines Werts, z. B. für Titel-Platzhalter und Sichtbarkeitsregeln. */
export function valueToText(v: FieldValue): string {
  if (v === undefined || v === null) {
    return '';
  }
  if (typeof v === 'boolean') {
    return v ? 'Ja' : 'Nein';
  }
  if (Array.isArray(v)) {
    return v.join(', ');
  }
  if (typeof v === 'object') {
    return v.title;
  }
  return v;
}
