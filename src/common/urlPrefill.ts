import { FieldValue, FieldValues, ILookupValue, ISpField, ITermValue, IUserValue } from './spTypes';
import { SpService } from './spService';
import { addDays, toIsoDate } from './dateUtils';
import { t } from './i18n';

/**
 * Vorbelegung von Formularfeldern über URL-Parameter.
 * Format: fd_<InternerName>=<Wert>, z. B. ?fd_Title=Heizung&fd_Status=Offen
 * Mehrere Werte (Mehrfachauswahl, Personen, Nachschlagen, Metadaten) mit «;» trennen.
 */
export const URL_PREFIX: string = 'fd_';

/** Alle Query-Parameter einer URL (bereits dekodiert). */
export function readUrlParams(url: string | undefined): Record<string, string> {
  const out: Record<string, string> = {};
  if (!url) {
    return out;
  }
  try {
    new URL(url, 'https://localhost/').searchParams.forEach((value, key) => {
      out[key] = value;
    });
  } catch {
    // ignorieren
  }
  return out;
}

/** Nur die fd_-Parameter, Schlüssel = interner Feldname in Kleinbuchstaben. */
export function prefillParams(params: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  Object.keys(params).forEach((key) => {
    if (key.toLowerCase().indexOf(URL_PREFIX) === 0 && key.length > URL_PREFIX.length) {
      out[key.slice(URL_PREFIX.length).toLowerCase()] = params[key];
    }
  });
  return out;
}

function splitList(raw: string): string[] {
  return raw
    .split(';')
    .map((x) => x.trim())
    .filter(Boolean);
}

/** «today», «today+7», «today-1» oder JJJJ-MM-TT. */
export function parseDateParam(raw: string): string | undefined {
  const v = raw.trim().toLowerCase();
  const rel = /^(today|heute)\s*([+-]\s*\d+)?$/.exec(v);
  if (rel) {
    return toIsoDate(addDays(new Date(), rel[2] ? Number(rel[2].replace(/\s/g, '')) : 0));
  }
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
  return iso ? v : undefined;
}

/** Wandelt einen Parameterwert für Spaltentypen ohne Serverabfrage um. */
export function parseSimpleValue(field: ISpField, raw: string): { value?: FieldValue; error?: string } {
  switch (field.kind) {
    case 'text':
    case 'note':
      return { value: raw };
    case 'number':
      return isNaN(Number(raw)) ? { error: t('«{0}» ist keine Zahl.', raw) } : { value: raw };
    case 'choice':
      return field.choices.length && field.choices.indexOf(raw) < 0 ? { error: t('«{0}» ist keine gültige Auswahl.', raw) } : { value: raw };
    case 'multichoice': {
      const items = splitList(raw);
      const invalid = items.filter((x) => field.choices.length > 0 && field.choices.indexOf(x) < 0);
      return invalid.length ? { error: t('«{0}» ist keine gültige Auswahl.', invalid.join(', ')) } : { value: items };
    }
    case 'boolean': {
      const v = raw.trim().toLowerCase();
      if (['1', 'true', 'ja', 'yes'].indexOf(v) >= 0) {
        return { value: true };
      }
      if (['0', 'false', 'nein', 'no'].indexOf(v) >= 0) {
        return { value: false };
      }
      return { error: t('«{0}» ist kein Ja/Nein-Wert.', raw) };
    }
    case 'date': {
      const d = parseDateParam(raw);
      return d ? { value: d } : { error: t('«{0}» ist kein Datum (JJJJ-MM-TT oder today+n).', raw) };
    }
    case 'datetime': {
      const m = /^(\d{4}-\d{2}-\d{2})(?:[T ](\d{2}:\d{2}))?$/.exec(raw.trim());
      if (m) {
        return { value: `${m[1]}T${m[2] || '00:00'}` };
      }
      const d = parseDateParam(raw);
      return d ? { value: `${d}T00:00` } : { error: t('«{0}» ist kein Datum (JJJJ-MM-TT oder today+n).', raw) };
    }
    case 'url': {
      const [url, ...rest] = raw.split('|');
      return { value: { url: url.trim(), description: rest.join('|').trim() } };
    }
    default:
      return {};
  }
}

const ASYNC_KINDS: string[] = ['user', 'usermulti', 'lookup', 'lookupmulti', 'taxonomy', 'taxonomymulti'];

async function resolveAsync(field: ISpField, raw: string, sp: SpService): Promise<FieldValue> {
  const items = splitList(raw);
  const multi = field.kind === 'usermulti' || field.kind === 'lookupmulti' || field.kind === 'taxonomymulti';
  if (field.kind === 'user' || field.kind === 'usermulti') {
    const users: IUserValue[] = [];
    for (const login of items) {
      users.push(await sp.ensureUser(login));
    }
    return multi ? users : users[0] || null;
  }
  if (field.kind === 'lookup' || field.kind === 'lookupmulti') {
    const options = await sp.getLookupItems(field);
    const found: ILookupValue[] = items.map((x) => {
      const hit = /^\d+$/.test(x) ? options.filter((o) => o.id === Number(x))[0] : options.filter((o) => o.title.toLowerCase() === x.toLowerCase())[0];
      if (!hit) {
        throw new Error(t('«{0}» wurde in der Nachschlageliste nicht gefunden.', x));
      }
      return hit;
    });
    return multi ? found : found[0] || null;
  }
  const terms = await sp.getTerms(field);
  const found: ITermValue[] = items.map((x) => {
    const key = x.toLowerCase();
    const hit = terms.filter((o) => o.termGuid.toLowerCase() === key)[0] || terms.filter((o) => o.label.toLowerCase() === key)[0];
    if (!hit) {
      throw new Error(t('Der Begriff «{0}» wurde nicht gefunden.', x));
    }
    return { termGuid: hit.termGuid, label: hit.label };
  });
  return multi ? found : found[0] || null;
}

/** Ermittelt die Werte aller fd_-Parameter. Fehler werden gesammelt, nicht geworfen. */
export async function resolvePrefill(params: Record<string, string>, fields: ISpField[], sp: SpService | undefined): Promise<{ values: FieldValues; errors: string[] }> {
  const wanted = prefillParams(params);
  const values: FieldValues = {};
  const errors: string[] = [];
  for (const field of fields) {
    const raw = wanted[field.internalName.toLowerCase()];
    if (raw === undefined || field.readOnly || field.kind === 'calculated' || field.kind === 'unsupported') {
      continue;
    }
    try {
      if (ASYNC_KINDS.indexOf(field.kind) >= 0) {
        if (!sp) {
          continue;
        }
        values[field.internalName] = await resolveAsync(field, raw, sp);
      } else {
        const r = parseSimpleValue(field, raw);
        if (r.error) {
          errors.push(`${field.title}: ${r.error}`);
        } else if (r.value !== undefined) {
          values[field.internalName] = r.value;
        }
      }
    } catch (e) {
      errors.push(`${field.title}: ${(e as Error).message}`);
    }
  }
  return { values, errors };
}

/** Baut einen Link auf das Formular «Neu» mit Vorbelegung. */
export function buildFormLink(options: {
  webUrl: string;
  listId: string;
  listUrl: string;
  contentTypeId: string;
  values: Record<string, string>;
  viaListPanel: boolean;
}): string {
  const extra = Object.keys(options.values)
    .filter((k) => options.values[k] !== '')
    .map((k) => `${URL_PREFIX}${encodeURIComponent(k)}=${encodeURIComponent(options.values[k])}`)
    .join('&');
  const origin = options.webUrl.replace(/^(https?:\/\/[^/]+).*$/, '$1');
  if (options.viaListPanel) {
    return `${origin}${encodeURI(options.listUrl)}?fdOpen=new${extra ? `&${extra}` : ''}`;
  }
  const base = `${options.webUrl.replace(/\/$/, '')}/_layouts/15/SPListForm.aspx?PageType=8&List=${encodeURIComponent(`{${options.listId}}`)}&ContentTypeId=${encodeURIComponent(options.contentTypeId)}&RootFolder=${encodeURIComponent(options.listUrl)}`;
  return extra ? `${base}&${extra}` : base;
}
