import { IFieldSettings, PatternKind } from './model';
import { FieldValue, ISpField, isEmptyValue } from './spTypes';
import { toIsoDate } from './dateUtils';
import { t } from './i18n';

export const PATTERNS: Record<Exclude<PatternKind, 'none' | 'custom'>, { re: RegExp; label: () => string; message: () => string }> = {
  email: {
    re: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    label: () => t('E-Mail-Adresse'),
    message: () => t('Bitte eine gültige E-Mail-Adresse eingeben.')
  },
  phoneCH: {
    re: /^(\+41|0041|0)\s?[1-9]\d(\s?\d){7}$/,
    label: () => t('Telefonnummer (CH)'),
    message: () => t('Bitte eine Schweizer Telefonnummer eingeben, z. B. 041 123 45 67.')
  },
  plzCH: {
    re: /^[1-9]\d{3}$/,
    label: () => t('Postleitzahl (CH)'),
    message: () => t('Bitte eine vierstellige Postleitzahl eingeben.')
  }
};

/**
 * Prüft einen Feldwert gegen die im Designer konfigurierten Regeln.
 * Liefert eine Fehlermeldung oder einen leeren String.
 */
export function validateFieldValue(field: ISpField, settings: IFieldSettings, value: FieldValue, required: boolean): string {
  const rule = settings.validation;
  const custom = (fallback: string): string => rule.message || fallback;

  if (required && isEmptyValue(value)) {
    return t('Dieses Feld ist erforderlich.');
  }
  if (isEmptyValue(value)) {
    return '';
  }

  if ((field.kind === 'text' || field.kind === 'note') && typeof value === 'string') {
    if (rule.minLength !== undefined && rule.minLength !== null && value.length < rule.minLength) {
      return custom(t('Mindestens {0} Zeichen eingeben.', rule.minLength));
    }
    if (rule.maxLength !== undefined && rule.maxLength !== null && value.length > rule.maxLength) {
      return custom(t('Höchstens {0} Zeichen erlaubt (aktuell {1}).', rule.maxLength, value.length));
    }
    if (rule.pattern === 'custom' && rule.regex) {
      let re: RegExp;
      try {
        // Der Ausdruck stammt bewusst aus der Feldkonfiguration des Designers.
        // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
        re = new RegExp(rule.regex);
      } catch {
        return t('Ungültiger regulärer Ausdruck in der Feldkonfiguration.');
      }
      if (!re.test(value)) {
        return custom(t('Ungültiges Format.'));
      }
    } else if (rule.pattern !== 'none' && rule.pattern !== 'custom') {
      const p = PATTERNS[rule.pattern];
      if (p && !p.re.test(value)) {
        return custom(p.message());
      }
    }
  }

  if (field.kind === 'number' && typeof value === 'string') {
    const n = Number(value);
    if (isNaN(n)) {
      return custom(t('Bitte eine Zahl eingeben.'));
    }
    if (rule.min !== undefined && rule.min !== null && n < rule.min) {
      return custom(t('Der Wert muss mindestens {0} sein.', rule.min));
    }
    if (rule.max !== undefined && rule.max !== null && n > rule.max) {
      return custom(t('Der Wert darf höchstens {0} sein.', rule.max));
    }
  }

  if ((field.kind === 'date' || field.kind === 'datetime') && typeof value === 'string' && rule.dateRule !== 'none') {
    const day = value.slice(0, 10);
    const now = toIsoDate(new Date());
    if (rule.dateRule === 'notPast' && day < now) {
      return custom(t('Das Datum darf nicht in der Vergangenheit liegen.'));
    }
    if (rule.dateRule === 'future' && day <= now) {
      return custom(t('Das Datum muss in der Zukunft liegen.'));
    }
    if (rule.dateRule === 'notFuture' && day > now) {
      return custom(t('Das Datum darf nicht in der Zukunft liegen.'));
    }
  }

  return '';
}

export function hasValidationRules(settings: IFieldSettings): boolean {
  const r = settings.validation;
  return (
    (r.minLength !== undefined && r.minLength !== null) ||
    (r.maxLength !== undefined && r.maxLength !== null) ||
    r.pattern !== 'none' ||
    (r.min !== undefined && r.min !== null) ||
    (r.max !== undefined && r.max !== null) ||
    r.dateRule !== 'none'
  );
}
