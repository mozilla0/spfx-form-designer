import { getCulture } from './i18n';

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

/** Wandelt Date, 'YYYY-MM-DD' oder 'YYYY-MM-DDTHH:mm' in ein lokales Datum (00:00 Uhr) um. */
export function toLocalDate(x: unknown): Date {
  if (x instanceof Date) {
    return new Date(x.getFullYear(), x.getMonth(), x.getDate());
  }
  if (typeof x === 'string') {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(x);
    if (m) {
      return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    }
  }
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

export function today(): Date {
  return toLocalDate(new Date());
}

export function addDays(d: unknown, n: number): Date {
  const x = toLocalDate(d);
  x.setDate(x.getDate() + Number(n || 0));
  return x;
}

/** Addiert Werktage (Mo–Fr). Feiertage werden nicht berücksichtigt. */
export function addWorkdays(d: unknown, n: number): Date {
  const x = toLocalDate(d);
  let k = Number(n || 0);
  const step = k < 0 ? -1 : 1;
  k = Math.abs(k);
  while (k > 0) {
    x.setDate(x.getDate() + step);
    const w = x.getDay();
    if (w !== 0 && w !== 6) {
      k--;
    }
  }
  return x;
}

export function endOfMonth(d: unknown): Date {
  const x = toLocalDate(d);
  return new Date(x.getFullYear(), x.getMonth() + 1, 0);
}

export function toIsoDate(d: unknown): string {
  const x = toLocalDate(d);
  return `${x.getFullYear()}-${pad2(x.getMonth() + 1)}-${pad2(x.getDate())}`;
}

type DatePart = 'day' | 'month' | 'year';

/** Reihenfolge von Tag, Monat und Jahr in der aktuellen Kultur. */
function partOrder(): DatePart[] {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const fmt = new Intl.DateTimeFormat(getCulture(), { day: '2-digit', month: '2-digit', year: 'numeric' }) as any;
    const parts: Array<{ type: string }> = fmt.formatToParts(new Date(2000, 10, 22));
    const order = parts.filter((p) => p.type === 'day' || p.type === 'month' || p.type === 'year').map((p) => p.type as DatePart);
    return order.length === 3 ? order : ['day', 'month', 'year'];
  } catch {
    return ['day', 'month', 'year'];
  }
}

/** Datum im Format der Kultur, z. B. 03.11.2026 (de-CH) oder 11/03/2026 (en-US). */
export function formatDate(d: unknown): string {
  const x = toLocalDate(d);
  try {
    return new Intl.DateTimeFormat(getCulture(), { day: '2-digit', month: '2-digit', year: 'numeric' }).format(x);
  } catch {
    return `${pad2(x.getDate())}.${pad2(x.getMonth() + 1)}.${x.getFullYear()}`;
  }
}

export function formatDateWithWeekday(d: unknown): string {
  const x = toLocalDate(d);
  let weekday = '';
  try {
    weekday = new Intl.DateTimeFormat(getCulture(), { weekday: 'short' }).format(x);
  } catch {
    weekday = '';
  }
  return weekday ? `${weekday}, ${formatDate(x)}` : formatDate(x);
}

/** Platzhalter für die Datumseingabe, z. B. «TT.MM.JJJJ» oder «MM/DD/YYYY». */
export function datePlaceholder(): string {
  const de = /^de\b/i.test(getCulture());
  const names: Record<DatePart, string> = de ? { day: 'TT', month: 'MM', year: 'JJJJ' } : { day: 'DD', month: 'MM', year: 'YYYY' };
  const sep = formatDate(new Date(2000, 10, 22)).replace(/[0-9]/g, '').charAt(0) || '.';
  return partOrder().map((p) => names[p]).join(sep);
}

/** Liest ein Datum in der Reihenfolge der Kultur (Trennzeichen . / -) oder als JJJJ-MM-TT. */
export function parseDate(s: string): Date | null {
  const text = (s || '').trim();
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
  let day: number;
  let month: number;
  let year: number;
  if (iso) {
    year = Number(iso[1]);
    month = Number(iso[2]);
    day = Number(iso[3]);
  } else {
    const m = /^(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})$/.exec(text);
    if (!m) {
      return null;
    }
    const values = [Number(m[1]), Number(m[2]), Number(m[3])];
    const order = partOrder();
    day = values[order.indexOf('day')];
    month = values[order.indexOf('month')];
    year = values[order.indexOf('year')];
  }
  if (year < 100) {
    year += 2000;
  }
  const d = new Date(year, month - 1, day);
  return isNaN(d.getTime()) || d.getMonth() !== month - 1 ? null : d;
}

/** Abwärtskompatibel: TT.MM.JJJJ. */
export function parseSwissDate(s: string): Date | null {
  const m = /^\s*(\d{1,2})\.(\d{1,2})\.(\d{2}|\d{4})\s*$/.exec(s || '');
  if (!m) {
    return null;
  }
  let year = Number(m[3]);
  if (year < 100) {
    year += 2000;
  }
  const d = new Date(year, Number(m[2]) - 1, Number(m[1]));
  return isNaN(d.getTime()) ? null : d;
}

/** Datumshilfen, die im Formular-Skript als form.date verfügbar sind. */
export const scriptDateHelpers = {
  today,
  addDays,
  addWorkdays,
  endOfMonth,
  format: formatDate,
  iso: toIsoDate
};

export function pad(n: number): string {
  return pad2(n);
}
