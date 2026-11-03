import { DefaultMode, IFieldSettings } from './model';
import { FieldKind, FieldValue, FieldValues, ICurrentUser, ISpField } from './spTypes';
import { addDays, addWorkdays, endOfMonth, formatDate, today, toIsoDate } from './dateUtils';
import { t } from './i18n';

export interface IDefaultResult {
  has: boolean;
  value?: FieldValue;
  error?: string;
}

export interface IDefaultOption {
  key: DefaultMode;
  text: string;
}

export function defaultOptionsFor(kind: FieldKind): IDefaultOption[] {
  const none: IDefaultOption = { key: 'none', text: t('Keine') };
  const formula: IDefaultOption = { key: 'formula', text: t('Formel (JavaScript)') };
  switch (kind) {
    case 'date':
    case 'datetime':
      return [
        none,
        { key: 'today', text: t('Heute') },
        { key: 'todayPlus', text: t('Heute + n Tage') },
        { key: 'nextWorkday', text: t('Nächster Werktag') },
        { key: 'workdaysPlus', text: t('Heute + n Werktage') },
        { key: 'endOfMonth', text: t('Ende des Monats') },
        formula
      ];
    case 'user':
    case 'usermulti':
      return [none, { key: 'currentUser', text: t('Aktueller Benutzer') }];
    case 'boolean':
      return [none, { key: 'fixedTrue', text: t('Ja') }, formula];
    case 'text':
    case 'note':
    case 'number':
    case 'choice':
      return [none, { key: 'fixed', text: t('Fester Wert') }, formula];
    default:
      return [none];
  }
}

function normalizeResult(kind: FieldKind, v: unknown): FieldValue {
  if (v instanceof Date) {
    if (kind === 'date') {
      return toIsoDate(v);
    }
    if (kind === 'datetime') {
      return `${toIsoDate(v)}T00:00`;
    }
    return formatDate(v);
  }
  if (kind === 'boolean') {
    return !!v;
  }
  if (v === undefined || v === null) {
    return '';
  }
  if (kind === 'datetime' && typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) {
    return `${v}T00:00`;
  }
  return typeof v === 'string' ? v : String(v);
}

/** Berechnet den Standardwert eines Felds für das Formular «Neu». */
export function computeDefault(field: ISpField, settings: IFieldSettings, values: FieldValues, user: ICurrentUser | undefined): IDefaultResult {
  const rule = settings.defaultValue;
  if (!rule || rule.mode === 'none') {
    return { has: false };
  }
  try {
    let v: unknown;
    switch (rule.mode) {
      case 'today':
        v = today();
        break;
      case 'todayPlus':
        v = addDays(today(), rule.days || 0);
        break;
      case 'workdaysPlus':
        v = addWorkdays(today(), rule.days || 0);
        break;
      case 'nextWorkday':
        v = addWorkdays(today(), 1);
        break;
      case 'endOfMonth':
        v = endOfMonth(today());
        break;
      case 'currentUser':
        if (!user) {
          return { has: false };
        }
        return {
          has: true,
          value: field.kind === 'usermulti' ? [{ id: user.id, title: user.title, email: user.email }] : { id: user.id, title: user.title, email: user.email }
        };
      case 'fixed':
        v = rule.value || '';
        break;
      case 'fixedTrue':
        v = true;
        break;
      case 'formula': {
        const formula = (rule.formula || '').trim();
        if (!formula) {
          return { has: true, error: t('Die Formel ist leer.') };
        }
        // eslint-disable-next-line no-new-func
        const fn = new Function('today', 'addDays', 'addWorkdays', 'endOfMonth', 'field', 'user', `"use strict"; return (${formula}\n);`);
        v = fn(today, addDays, addWorkdays, endOfMonth, (name: string) => values[name], user ? { displayName: user.title, email: user.email } : {});
        break;
      }
    }
    return { has: true, value: normalizeResult(field.kind, v) };
  } catch (e) {
    return { has: true, error: t('Formelfehler: {0}', (e as Error).message) };
  }
}
